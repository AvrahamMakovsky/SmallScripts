"""
Remote Present PnP Device Search
================================

Small Windows utility for checking whether specific Plug and Play devices are
currently detected on multiple remote hosts.

The script asks for a FriendlyName identifier, opens Notepad so the operator can
paste a list of hosts, then queries each host with PowerShell remoting and
Get-PnpDevice -PresentOnly. Results are normalized into a compact console table
and can be exported to Excel or CSV.

Typical use cases:
- Find which lab hosts currently detect a specific USB/network/debug device.
- Verify whether a device is present before starting validation or support work.
- Collect quick evidence across several machines without opening Device Manager
  manually on each one.

Requirements:
- Windows host running the script.
- PowerShell remoting / WinRM access to the target hosts.
- Permissions to run Get-PnpDevice on the target hosts.
- Optional: pandas + openpyxl for Excel export. CSV export works without them.

Notes:
- The search is based on a case-insensitive literal substring match against
  Device Manager FriendlyName.
- Ping pre-check is optional and disabled by default because ICMP is often
  blocked even when WinRM works.

Author: Avraham Makovsky
Contact: Avraham Makovsky
"""

import csv
import json
import os
import subprocess
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Tuple

try:
    import pandas as pd

    PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    PANDAS_AVAILABLE = False

try:
    import tkinter as tk
    from tkinter import filedialog

    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False


__version__ = "2.1.0"

POWERSHELL_EXE = "powershell.exe"
DEFAULT_TIMEOUT_SECONDS = 120
DEFAULT_MAX_WORKERS = 16
PING_PRECHECK_DEFAULT = False

# Keep exported columns stable so repeated runs are easy to compare or merge.
RESULT_COLUMNS = [
    "Host",
    "Identifier_Found",
    "Match_Count",
    "Matched_Device_Name",
    "Device_Class",
    "Present",
    "Device_Status",
    "InstanceId",
    "Problem",
    "Query_Error",
]


# Embedded PowerShell keeps the Python file self-contained while still using
# native Windows inventory data. Arguments are passed through -File parameters
# instead of string interpolation to avoid fragile quoting and accidental command
# injection when host names or identifiers contain special characters.
REMOTE_PNP_SEARCH_PS = r'''
param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [Parameter(Mandatory = $true)]
    [string]$Identifier
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$scriptBlock = {
    param([string]$Identifier)

    $ErrorActionPreference = "Stop"

    $devices = @(
        Get-PnpDevice -PresentOnly | Where-Object {
            $_.FriendlyName -and ($_.FriendlyName.IndexOf($Identifier, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
        } | Select-Object Status, Class, FriendlyName, InstanceId, Problem, Present
    )

    ConvertTo-Json -InputObject $devices -Compress -Depth 4
}

Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $Identifier
'''


def create_temp_powershell_script() -> str:
    """Write the embedded PowerShell helper to a temporary .ps1 file."""
    fd, path = tempfile.mkstemp(prefix="pnp_device_search_", suffix=".ps1", text=True)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(REMOTE_PNP_SEARCH_PS)
    return path


def get_hosts_from_notepad() -> List[str]:
    """Collect host names through Notepad, which is convenient for ad-hoc lab lists."""
    temp_path = None
    try:
        fd, temp_path = tempfile.mkstemp(prefix="hosts_", suffix=".txt", text=True)
        os.close(fd)

        initial_text = (
            "# Put one host per line\n"
            "# Lines starting with # or ; are ignored\n"
            "# Duplicate hosts are removed automatically\n"
            "# Example:\n"
            "PC123\n"
            "PC456\n"
        )

        with open(temp_path, "w", encoding="utf-8") as f:
            f.write(initial_text)

        subprocess.run(["notepad.exe", temp_path], check=False)

        with open(temp_path, "r", encoding="utf-8") as f:
            raw_lines = f.readlines()

        hosts = []
        seen = set()

        for line in raw_lines:
            host = line.strip()
            if not host or host.startswith("#") or host.startswith(";"):
                continue

            key = host.lower()
            if key not in seen:
                seen.add(key)
                hosts.append(host)

        return hosts

    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def ask_int(prompt: str, default: int, minimum: int = 1, maximum: int = 64) -> int:
    raw = input(f"{prompt} [{default}]: ").strip()
    if not raw:
        return default

    try:
        value = int(raw)
        return max(minimum, min(maximum, value))
    except ValueError:
        print(f"Invalid value. Using default: {default}")
        return default


def ask_yes_no(prompt: str, default: bool = False) -> bool:
    default_text = "Y" if default else "N"
    raw = input(f"{prompt} [Y/N, default {default_text}]: ").strip().lower()
    if not raw:
        return default
    return raw in ("y", "yes", "1", "true")


def choose_save_path(extension: str, file_label: str) -> Optional[str]:
    """Use a save dialog when Tkinter is available, otherwise save in cwd."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_name = f"PnP_Device_Search_{timestamp}.{extension}"

    if TK_AVAILABLE:
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)

            path = filedialog.asksaveasfilename(
                title=f"Save results as {file_label}",
                defaultextension=f".{extension}",
                initialfile=default_name,
                filetypes=[(f"{file_label} files", f"*.{extension}"), ("All files", "*.*")],
            )

            root.destroy()
            return path
        except Exception:
            pass

    return os.path.join(os.getcwd(), default_name)


def unique_preserve_order(items: Iterable[object]) -> List[str]:
    seen = set()
    out = []
    for item in items:
        value = str(item).strip()
        if not value:
            continue
        key = value.lower()
        if key not in seen:
            seen.add(key)
            out.append(value)
    return out


def join_values(items: Iterable[object]) -> str:
    values = unique_preserve_order(items)
    return " | ".join(values) if values else "NA"


def empty_result(host: str, status: str, error: str = "") -> Dict[str, object]:
    """Return one normalized row so success, no-match, and failure cases align."""
    return {
        "Host": host,
        "Identifier_Found": "No",
        "Match_Count": 0,
        "Matched_Device_Name": "NA",
        "Device_Class": "NA",
        "Present": "No",
        "Device_Status": status,
        "InstanceId": "NA",
        "Problem": "",
        "Query_Error": error,
    }


def classify_error(stderr: str) -> Tuple[str, str]:
    """Convert common remoting failures into statuses an operator can act on."""
    text = (stderr or "").strip()
    lowered = text.lower()

    if not text:
        return "Failed", "PowerShell returned a non-zero exit code without stderr"

    if "access is denied" in lowered or "unauthorizedaccess" in lowered:
        return "Access Denied", text

    if "winrm" in lowered or "wsman" in lowered or "psremotingtransportexception" in lowered:
        return "WinRM/Remoting Failed", text

    if "cannot find the computer" in lowered or "could not resolve" in lowered or "name resolution" in lowered:
        return "Host Not Resolved", text

    if "network path was not found" in lowered or "the client cannot connect" in lowered:
        return "Host Unreachable", text

    return "Failed", text


def ping_host(host: str, timeout_ms: int = 1000) -> bool:
    """Best-effort ICMP check. False here should not be treated as proof of failure."""
    if os.name != "nt":
        return True

    try:
        completed = subprocess.run(
            ["ping", "-n", "1", "-w", str(timeout_ms), host],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        return completed.returncode == 0
    except Exception:
        # If ping itself fails locally, don't block the real remoting query.
        return True


def run_powershell_query(ps_script_path: str, host: str, identifier: str, timeout_seconds: int) -> Tuple[str, str, int]:
    proc = subprocess.Popen(
        [
            POWERSHELL_EXE,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            ps_script_path,
            "-ComputerName",
            host,
            "-Identifier",
            identifier,
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    stdout, stderr = proc.communicate(timeout=timeout_seconds)
    return stdout, stderr, proc.returncode


def normalize_json_to_list(data: object) -> List[dict]:
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if isinstance(data, dict):
        return [data]
    return []


def parse_host_result(host: str, stdout: str, stderr: str, returncode: int) -> Dict[str, object]:
    """Parse PowerShell JSON and reduce any number of matches to one report row."""
    if returncode != 0:
        status, error = classify_error(stderr)
        return empty_result(host, status, error)

    output = (stdout or "").strip()
    if not output:
        return empty_result(host, "Not Found", "No output returned")

    try:
        data = json.loads(output)
        devices = normalize_json_to_list(data)

        real_devices = []
        for dev in devices:
            friendly_name = str(dev.get("FriendlyName", "NA") or "NA")
            present = bool(dev.get("Present", False))
            if friendly_name != "NA" and present:
                real_devices.append(dev)

        if not real_devices:
            return empty_result(host, "Not Found")

        matched_names = []
        device_classes = []
        statuses = []
        instance_ids = []
        problems = []

        for dev in real_devices:
            matched_names.append(dev.get("FriendlyName", ""))
            device_classes.append(dev.get("Class", ""))
            statuses.append(dev.get("Status", ""))
            instance_ids.append(dev.get("InstanceId", ""))
            problems.append(dev.get("Problem", ""))

        return {
            "Host": host,
            "Identifier_Found": "Yes",
            "Match_Count": len(real_devices),
            "Matched_Device_Name": join_values(matched_names),
            "Device_Class": join_values(device_classes),
            "Present": "Yes",
            "Device_Status": join_values(statuses),
            "InstanceId": join_values(instance_ids),
            "Problem": join_values(problems),
            "Query_Error": "",
        }

    except Exception as e:
        return empty_result(host, "Failed", f"JSON parse failed: {e}; Raw output: {output[:500]}")


def query_host(ps_script_path: str, host: str, identifier: str, timeout_seconds: int, ping_precheck: bool) -> Dict[str, object]:
    if ping_precheck and not ping_host(host):
        return empty_result(host, "Host Unreachable", "Ping pre-check failed. ICMP may be blocked, so retry without ping pre-check if needed.")

    try:
        stdout, stderr, returncode = run_powershell_query(ps_script_path, host, identifier, timeout_seconds)
        return parse_host_result(host, stdout, stderr, returncode)
    except subprocess.TimeoutExpired:
        return empty_result(host, "Timed Out", f"Timed out after {timeout_seconds} seconds")
    except FileNotFoundError:
        return empty_result(host, "Local Error", f"{POWERSHELL_EXE} was not found on this computer")
    except Exception as e:
        return empty_result(host, "Local Error", str(e))


def print_result_line(index: int, total: int, result: Dict[str, object]) -> None:
    print(
        f"[{index}/{total}] {result['Host']}: "
        f"Found={result['Identifier_Found']}, "
        f"Count={result['Match_Count']}, "
        f"Class={result['Device_Class']}, "
        f"Status={result['Device_Status']}"
    )


def print_table(results: List[Dict[str, object]]) -> None:
    print("\n" + "=" * 120)
    print(f"{'Host':<22} {'Found':<8} {'Count':<8} {'Device Class':<25} {'Status':<22} {'Matched Device Name'}")
    print("-" * 160)

    for row in results:
        print(
            f"{str(row['Host']):<22} "
            f"{str(row['Identifier_Found']):<8} "
            f"{str(row['Match_Count']):<8} "
            f"{str(row['Device_Class'])[:24]:<25} "
            f"{str(row['Device_Status'])[:21]:<22} "
            f"{str(row['Matched_Device_Name'])}"
        )


def export_csv(results: List[Dict[str, object]], path: str) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=RESULT_COLUMNS)
        writer.writeheader()
        writer.writerows(results)


def export_excel(results: List[Dict[str, object]], path: str) -> None:
    if not PANDAS_AVAILABLE:
        raise RuntimeError("pandas/openpyxl is not installed")

    df = pd.DataFrame(results, columns=RESULT_COLUMNS)
    df.to_excel(path, index=False)


def export_results(results: List[Dict[str, object]]) -> None:
    print("\n" + "=" * 70)

    if PANDAS_AVAILABLE:
        choice = input("Export results? [X=Excel / C=CSV / N=No, default X]: ").strip().lower()
        if not choice:
            choice = "x"
    else:
        print("Excel export requires pandas and openpyxl. CSV export is available without extra modules.")
        choice = input("Export results? [C=CSV / N=No, default C]: ").strip().lower()
        if not choice:
            choice = "c"

    if choice in ("n", "no", "0"):
        print("\nResults not exported.")
        return

    if choice in ("x", "excel", "xlsx") and PANDAS_AVAILABLE:
        path = choose_save_path("xlsx", "Excel")
        if not path:
            print("\nExport cancelled.")
            return
        export_excel(results, path)
        print(f"\nExcel file saved: {path}")
        return

    if choice in ("c", "csv") or not PANDAS_AVAILABLE:
        path = choose_save_path("csv", "CSV")
        if not path:
            print("\nExport cancelled.")
            return
        export_csv(results, path)
        print(f"\nCSV file saved: {path}")
        return

    print("\nInvalid export choice. Results not exported.")


def main() -> None:
    print(f"Remote Present PnP Device Search v{__version__}")
    print("=" * 78)
    print("What it does:")
    print("  Searches multiple remote Windows hosts for currently present Plug and Play")
    print("  devices whose Device Manager FriendlyName contains a given identifier.")
    print()
    print("How it works:")
    print("  1. You enter a device identifier, for example: ASIX, J5, USB Serial, Realtek.")
    print("  2. Notepad opens so you can paste one host name per line.")
    print("  3. The script runs Get-PnpDevice -PresentOnly remotely on each host.")
    print("  4. Results are shown in the console and can be exported to Excel or CSV.")
    print()
    print("Requirements:")
    print("  Windows, PowerShell remoting/WinRM access, and permission to query the hosts.")
    print("  Excel export needs pandas + openpyxl; CSV export has no extra dependency.")
    print("=" * 78)

    if os.name != "nt":
        print("Warning: this tool is intended to run on Windows with powershell.exe and notepad.exe.")

    identifier = input("\nEnter the friendly-name identifier to search for: ").strip()
    if not identifier:
        print("No identifier provided.")
        input("\nPress Enter to exit...")
        return

    hosts = get_hosts_from_notepad()
    if not hosts:
        print("No hosts were provided.")
        input("\nPress Enter to exit...")
        return

    max_workers = ask_int("Max parallel host checks", DEFAULT_MAX_WORKERS, minimum=1, maximum=64)
    ping_precheck = ask_yes_no("Run ping pre-check before remoting", PING_PRECHECK_DEFAULT)

    print(f"\nSearching for identifier: {identifier}")
    print(f"Checking {len(hosts)} host(s), max parallel checks: {max_workers}")
    if ping_precheck:
        print("Ping pre-check is enabled. If many hosts show unreachable, retry with ping pre-check disabled.")
    print("-" * 70)

    ps_script_path = create_temp_powershell_script()
    results = []
    completed_count = 0

    try:
        # ThreadPoolExecutor limits local PowerShell processes while keeping the
        # tool fast enough for typical lab-sized batches.
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_host = {
                executor.submit(
                    query_host,
                    ps_script_path,
                    host,
                    identifier,
                    DEFAULT_TIMEOUT_SECONDS,
                    ping_precheck,
                ): host
                for host in hosts
            }

            for future in as_completed(future_to_host):
                completed_count += 1
                host = future_to_host[future]
                try:
                    result = future.result()
                except Exception as e:
                    result = empty_result(host, "Local Error", str(e))

                results.append(result)
                print_result_line(completed_count, len(hosts), result)

    finally:
        try:
            os.remove(ps_script_path)
        except Exception:
            pass

    results.sort(key=lambda row: str(row["Host"]).lower())

    found_count = sum(1 for row in results if row["Identifier_Found"] == "Yes")
    not_found_count = sum(1 for row in results if row["Device_Status"] == "Not Found")
    failed_count = len(results) - found_count - not_found_count
    multi_count = sum(1 for row in results if int(row["Match_Count"]) > 1)

    print("\n" + "=" * 100)
    print(
        f"SUMMARY: {found_count} found, "
        f"{not_found_count} not found, "
        f"{failed_count} failed/unreachable, "
        f"{multi_count} with multiple matches"
    )
    print("=" * 100)

    print_table(results)
    export_results(results)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
