# SmallScripts

Small collection of practical utility scripts that I use for everyday tasks, experiments, and technical work.

## About this repository

These scripts started as personal side projects that I wrote on my own time and hardware.

They are general-purpose tools that can be useful at home or in professional environments, for example validation labs, support work, Windows administration, workflow automation, or general system maintenance.

They are **not official tools of any employer**, past or present.

The scripts are intentionally kept simple and direct. Some of them are monolithic by design because they are easier to copy, inspect, adapt, and run in restricted Windows environments.

Over time I may add scripts in different languages, such as:

- Python
- PowerShell
- Windows batch
- Other small tools and helpers

## Safety and scope

Use these scripts only on systems, files, and environments where you are authorized to work.

Before publishing, sharing, or adapting any script for another environment, review it for:

- Internal hostnames
- Private domains
- Corporate names
- Real ticket/work-item references
- Workbook links
- User names
- Exported operational data
- Local logs or SQLite databases

Runtime files such as logs, databases, exports, and temporary files should not be committed to the repository.

Recommended `.gitignore` additions:

```gitignore
*.log
*.db
*.sqlite
*.sqlite3
__pycache__/
*.pyc
```

## Related repositories

- [KVM Auto Login Helper](https://github.com/AvrahamMakovsky/kvm-auto-login-helper) - Chrome MV3 helper for auto-filling repetitive local KVM login pages.

---

## Scripts

### spreadsheet_work_item_flow.py

Windows desktop utility for converting spreadsheet rows into reviewed work-item proposals using the local Microsoft Excel desktop application.

#### What it does

- Opens or reconnects to a workbook through the installed Excel desktop app.
- Reads worksheet data through Excel COM automation.
- Auto-detects likely columns such as status, owner, station/host, description, completion date, priority, and work-item reference.
- Allows manual correction of the detected column mapping.
- Converts spreadsheet rows into reviewable work-item proposals.
- Lets the user edit, approve, skip, or bulk-select proposals.
- Supports dry-run mode before any real creation flow.
- Uses local SQLite duplicate protection to reduce repeated processing.
- Can write created work-item references back into the source workbook.
- Uses a mocked work-item creation service by default, so it is safe to adapt before connecting to a real system.

#### Typical use case

A team receives operational tasks in spreadsheet form, but direct access to the workbook through APIs is unavailable, unreliable, or requires permissions the user does not have.

Instead of manually converting rows into tickets or work items, this script uses the already-opened Excel desktop session as a bridge, then keeps the final decision human-reviewed.

#### Requirements

- Windows
- Python 3
- Microsoft Excel desktop app
- pywin32

Install dependency:

```bash
pip install pywin32
```

#### How to use

1. Run the script:

```bash
python spreadsheet_work_item_flow.py
```

2. Choose a source:
   - SharePoint / Excel URL
   - Local Excel file

3. Open or reconnect the workbook in Excel.

4. Select the worksheet.

5. Load or refresh worksheet data.

6. Review the detected column mapping.

7. Review/edit the generated proposals.

8. Use `Dry Run` first.

9. Replace the mocked work-item service with an approved integration if needed.

10. Create work items and optionally write the created reference back to Excel.

#### Notes

The public version intentionally uses a mocked work-item creation service.

Do not commit runtime logs, local SQLite databases, real workbook links, internal field names, or exported spreadsheet data.

---

### remote_pnp_device_search.py

Remote Plug and Play device search utility for Windows hosts using Python and PowerShell Remoting.

#### What it does

- Searches remote Windows hosts for currently present Plug and Play devices.
- Uses `Get-PnpDevice -PresentOnly` through PowerShell Remoting / WinRM.
- Matches devices by a literal, case-insensitive substring in the device `FriendlyName`.
- Opens Notepad so a host list can be pasted quickly.
- Deduplicates hostnames while preserving order.
- Supports multiple matching devices per host.
- Shows progress and a compact summary in the terminal.
- Exports results to Excel when `pandas` and `openpyxl` are installed.
- Falls back to CSV export when Excel dependencies are unavailable.
- Classifies common query failures such as access denied, remoting failure, unresolved host, timeout, or local execution error.

#### Typical use case

You need to quickly check whether a specific USB adapter, serial adapter, network device, lab dongle, or other Device Manager item is currently detected on a list of remote Windows machines.

Example questions this script helps answer:

- Which hosts currently detect this device?
- Is the device visible in Device Manager on the remote machine?
- What Windows device class is reported for the match?
- Are there multiple matching devices on the same host?
- Did the query fail because of access, WinRM, DNS, timeout, or no match?

#### Requirements

- Windows
- Python 3
- PowerShell
- PowerShell Remoting / WinRM enabled and reachable on target hosts
- Permissions to query the remote machines

Optional, for Excel export:

```bash
pip install pandas openpyxl
```

Without these packages, CSV export can still be used.

#### How to use

1. Run the script:

```bash
python remote_pnp_device_search.py
```

2. Enter a device friendly-name identifier, for example:

```text
ASIX
```

3. Paste one hostname per line into the Notepad window that opens:

```text
PC123
PC456
LAB-HOST-07
```

4. Save and close Notepad.
5. Wait for the remote checks to finish.
6. Export the results if needed.

#### Output columns

| Column | Meaning |
|---|---|
| `Host` | Target machine name |
| `Identifier_Found` | Whether a matching present device was found |
| `Match_Count` | Number of matched present devices |
| `Matched_Device_Name` | Friendly name reported by Windows |
| `Device_Class` | Windows PnP device class |
| `Present` | Whether the device is currently present |
| `Device_Status` | Device status reported by Windows |
| `InstanceId` | PnP instance ID |
| `Problem` | Device problem code, if reported |
| `Query_Error` | Remoting, query, parsing, or timeout details if something failed |

#### Notes

This is a focused support/lab utility, not a replacement for full inventory platforms such as SCCM, Intune, Lansweeper, or PDQ Inventory.

For large managed environments, a central inventory platform is usually more appropriate. For quick checks across a limited host list, this script is intended to stay simple and fast.

---

### SSDHealthMonitor.py

Command-line SSD health monitor.

#### What it does

- Reads SMART data from local drives using `smartctl` from smartmontools.
- Shows key information such as:
  - Drive model and serial
  - Capacity
  - Temperature
  - Power-on hours
  - Basic health indicators
- Prints a compact summary in the terminal so you can quickly see the state of your drives.

#### Requirements

- Python 3
- smartmontools / `smartctl`

#### How to use

1. Install Python 3.
2. Install `smartctl` and ensure it is in your system PATH.
3. Run the script:

```bash
python SSDHealthMonitor.py
```

---

### extractor.py

Hostname extractor for ticket titles, logs, or exported text.

#### What it does

- Scans a text file with ticket titles, logs, or copied text.
- Extracts hostnames that match the configured hostname pattern.
- Outputs a clean list of unique hostnames for use in other tools.

#### Typical use case

You have exported many tickets or logs into a `.txt` file and want to identify all mentioned hostnames automatically.

#### How to use

1. Save your ticket titles or text lines into a `.txt` file.
2. Place the file next to the script.
3. Run:

```bash
python extractor.py
```

---

### BulkPingLauncher.bat

Paste a list of hosts and open a continuous `ping -t` for each one.

#### What it does

- Opens a Notepad window for entering hosts.
- Accepts hosts separated by space, comma, or newline.
- If Windows Terminal (`wt.exe`) is found, opens each host in a new tab.
- If Windows Terminal is not found, falls back to separate CMD windows.

#### Typical use case

- Monitor connectivity of several servers, lab machines, kiosks, or workstations.
- Run quick bulk diagnostics.

---

### BulkVNCLauncher.bat

Paste a list of hosts and launch RealVNC Viewer sessions for all of them.

#### What it does

- Opens Notepad for list entry.
- Accepts hosts separated by space, comma, or newline.
- Starts a RealVNC session for each host using the installed `vncviewer.exe`.
- Requires RealVNC Viewer to be installed.
- The viewer path can be adjusted inside the script.

#### Typical use case

Bulk-connect to lab stations, remote machines, kiosks, or support targets.

---

### ipconfig_release.bat

Remote network repair helper for Windows.

#### What it does

- Schedules a reboot first, so the machine restarts even if the remote session drops.
- Runs:
  - `ipconfig /release`
  - `ipconfig /renew`
  - `netsh winsock reset`
- Writes a small log file under `%TEMP%` for troubleshooting.

#### Why it exists

When you run `ipconfig /release` on a remote host, the connection can drop before you can manually reboot the machine.

This script schedules the reboot first, then issues the network commands as best-effort.

#### How to use

1. Copy the `.bat` file to the remote host.
2. Run it as Administrator.

Note: `ipconfig /release` can temporarily drop the remote connection. That is expected.

---

### offlinehostnameupdate.ps1

Offline hostname editor for an offline Windows SSD workflow.

#### What it does

This script was created to reproduce, offline, hostname and network identity changes that normally happen while a machine is online.

It can update selected files and offline Windows registry locations on a connected SSD.

#### Behavior summary

- Edits configured hostname files on the offline installation.
- Mounts and updates the EFI/System partition when needed.
- Loads offline Windows registry hives and updates hostname-related entries.
- Updates relevant ControlSets.
- Uses safe unload logic for registry hives.
- Shows a summary before applying changes.
- Attempts to unmount what it mounted.

#### Typical use case

You have an offline Windows installation connected to a working machine and need to adjust its hostname-related configuration before booting it again.

#### How to use

1. Connect the target SSD to your working machine.
2. Run PowerShell as Administrator.
3. Run the script and follow the on-screen flow.
4. Double-check that you are targeting the offline SSD, not the live OS.

#### Notes

This is a beta tool. Use carefully.

Before publishing or adapting this script, sanitize any environment-specific hostname patterns, DNS suffixes, paths, or domain values.

---

### userrebuildreminder.ps1

Fullscreen user notification for a final rebuild or maintenance reminder.

#### What it does

- Displays a fullscreen, topmost notification window.
- Intended as a final on-screen reminder before scheduled disruptive maintenance.
- Requires explicit user confirmation before closing.
- Assumes the user was already notified through normal communication channels.

#### Behavior summary

- Borderless fullscreen window covering the entire display.
- Large, readable text suitable for shared workstations, lab stations, kiosks, or similar systems.
- Window is always on top.
- Close attempts trigger a confirmation dialog.
- Window closes only after the user explicitly acknowledges the message.

#### Typical use case

A machine is scheduled for rebuild, reimage, shutdown, or maintenance, and the user needs one last clear reminder to save work and close applications.

#### How to use

1. Run PowerShell as Administrator if needed.
2. Execute the script:

```powershell
.\userrebuildreminder.ps1
```

3. Optionally customize the message:

```powershell
.\userrebuildreminder.ps1 -Title "Final Reminder - Maintenance" -Message "This machine is scheduled for maintenance. Please save your work now."
```

---

## Getting started

1. Clone the repository:

```bash
git clone https://github.com/AvrahamMakovsky/SmallScripts.git
cd SmallScripts
```

2. Run the relevant script using Python, PowerShell, or Windows batch as appropriate.

3. Check each script's requirements before running it.

---

## Notes

These scripts are small and focused. I use them as building blocks and utilities in my day-to-day work and personal projects, and they may evolve over time.

Author: Avraham Makovsky  
Jerusalem, Israelhe reminder

#### Typical use case

- Validation lab or support environment
- Station is scheduled for a rebuild, reimage, or destructive maintenance
- User has already been informed verbally or by ticket/email
- Script is used as a final reminder to save work and close applications

#### How to use

1. Run PowerShell (Administrator recommended).
2. Execute the script:

```powershell
.\userrebuildreminder.ps1
```

3. Optionally customize the message:

```powershell
.\userrebuildreminder.ps1 -Title "Final Reminder - Station Rebuild" -Message "This station is scheduled for rebuild. Please save your work now."
```

---

## Getting started

1. Clone the repository:

```bash
git clone https://github.com/AvrahamMakovsky/SmallScripts.git
cd SmallScripts
```

2. Run the relevant scripts using Python, PowerShell, or Windows batch as appropriate.

---

## Notes

These scripts are small and focused. I use them as building blocks and utilities in my day-to-day work and personal projects, and they may evolve over time.

Author: Avraham Makovsky  
Jerusalem, Israel
## Scripts

### SSDHealthMonitor.py

Command-line SSD health monitor (Python).

#### What it does

- Reads SMART data from local drives (via `smartctl` from smartmontools).
- Shows key information such as:
  - Drive model and serial
  - Capacity
  - Temperature
  - Power-on hours
  - Basic health indicators (e.g. reallocated or pending sectors)
- Prints a compact summary in the terminal so you can quickly see the state of your drives.

#### How to use

1. Install Python 3.
2. Install `smartctl` (from smartmontools) and ensure it's in your system PATH.
3. Run the script:

```bash
python SSDHealthMonitor.py
```

---

### extractor.py

Hostname extractor for ticket titles or log text (Python).

#### What it does

- Scans a text file with ticket titles or logs.
- Extracts hostnames that match lab station patterns like `ST01****`.
- Outputs a clean list of unique hostnames for use in other tools.

#### Typical use case

- You have exported many tickets or logs into a `.txt` file.
- You want to identify all mentioned station hostnames automatically.
- You run the script and get a plain text list of hosts.

#### How to use

1. Save your ticket titles or text lines into a `.txt` file.
2. Place the file next to the script.
3. Run:

```bash
python extractor.py
```

---

### BulkPingLauncher.bat

Paste a list of hosts and open a continuous `ping -t` for each one - either in multiple tabs (Windows Terminal) or separate CMD windows.

#### What it does

- Opens a Notepad window for you to paste a list of hostnames (delimited by space, comma, or newline).
- If Windows Terminal (`wt.exe`) is found, each host is opened in a new tab with its own `ping -t`.
- If not found, falls back to opening a classic CMD window for each host.

#### Use case

- Monitor connectivity of several servers, lab machines, or kiosks
- Fast bulk diagnostics

---

### BulkVNCLauncher.bat

Paste a list of hosts and launch RealVNC Viewer sessions for all of them.

#### What it does

- Opens Notepad for list entry (hosts separated by space, comma, or newline).
- Starts a RealVNC session for each host using the installed `vncviewer.exe`.
- Requires RealVNC Viewer to be installed (path adjustable in script).

#### Use case

- Bulk-connect to lab stations or remote machines for access and support

---

### ipconfig_release.bat

Remote network repair helper (Windows batch) - release/renew DHCP, reset Winsock, and reboot immediately.

#### What it does

- Schedules a reboot immediately (default is 5 seconds) so the machine will restart even if your remote session drops.
- Runs:
  - `ipconfig /release`
  - `ipconfig /renew`
  - `netsh winsock reset`
- Writes a small log file under `%TEMP%` for troubleshooting.

#### Why it exists

When you run `ipconfig /release` on a remote host (RDP/SMB), the connection can drop before you can reboot the machine manually.  
This script schedules the reboot first, then issues the network commands as best-effort.

#### How to use

1. Copy the `.bat` file to the remote host.
2. Run it **as Administrator**.

Note: `ipconfig /release` can temporarily drop the remote connection. That's expected.

---

### offlinehostnameupdate.ps1

Offline hostname editor for an offline SSD workflow (Imaging + EFI + offline registry).

#### What it does

This script was created to simulate and reproduce, offline, the hostname and DNS configuration changes that normally happen online in a specific lab environment.

In that lab, target machines use hostnames with a suffix such as `_TA` and are paired to host systems.

After analyzing the original online scripts, I identified which files and registry locations are responsible for those changes and built this offline variant:

- Edits `Imaging\host.txt` and `Imaging\hostfqdn.txt` on the offline SSD.
- Mounts and updates `EFI\host.txt` on the ESP partition.
- Loads the offline Windows `SYSTEM` and `SOFTWARE` hives and updates the relevant hostname and DNS-related registry entries.

#### Behavior summary

- Registry only: ALL-CAPS hostname + optional ALL-CAPS suffix (`_TA`, etc.)
- `Imaging\host.txt` - bare HOSTNAME only (no suffix, no DNS)
- `Imaging\hostfqdn.txt` - HOSTNAME.FQDN-SUFFIX (true FQDN, suffix from config)
- `EFI\host.txt` - HOSTNAME.FQDN-SUFFIX (same as `Imaging\hostfqdn.txt`)

Other implementation details:

- EFI access: prefers already-mounted ESP on the same disk; otherwise uses `mountvol` to map (prefers `M:`)
- Robust volume GUID matching, readiness wait, provider priming, and safe read/write
- Updates all ControlSets (Tcpip, Tcpip6, ComputerName)
- Mirrors SOFTWARE ComputerName (with safe unload)
- Shows `Imaging\pair.json` in the summary and then removes it
- Always unmounts what it mounted

#### Configuration

The DNS suffix for the FQDN files is configured inside the script:

```powershell
$HostFqdnSuffixConfig = "iil.company.com"
```

#### How to use

1. Connect the target SSD (offline Windows installation) to your working machine.
2. Run PowerShell as Administrator.
3. Run the script and follow the on-screen flow to ensure the correct offline drive is selected.

Note: This is a beta tool. Use carefully and double-check that you are targeting the offline SSD, not the live OS.

---

### userrebuildreminder.ps1

Fullscreen user notification for final rebuild reminder (PowerShell / WinForms).

#### What it does

- Displays a fullscreen, topmost notification window on the user's station.
- Intended as a final on-screen reminder before a scheduled station rebuild.
- Assumes the user was already notified verbally or via other channels.
- Requires explicit user confirmation before closing.

This script was written for controlled lab or enterprise environments where users may remain logged in or overlook standard notifications, and a clear, unavoidable reminder is required before disruptive maintenance.

#### Behavior summary

- Borderless fullscreen window covering the entire display
- Large, readable text suitable for lab stations, kiosks, or shared machines
- Window is always on top
- Any close attempt triggers a confirmation dialog
- Window closes only after the user explicitly acknowledges the reminder

#### Typical use case

- Validation lab or support environment
- Station is scheduled for a rebuild, reimage, or destructive maintenance
- User has already been informed verbally or by ticket/email
- Script is used as a final reminder to save work and close applications

#### How to use

1. Run PowerShell (Administrator recommended).
2. Execute the script:

```powershell
.\userrebuildreminder.ps1
```

3. Optionally customize the message:

```powershell
.\userrebuildreminder.ps1 -Title "Final Reminder - Station Rebuild" -Message "This station is scheduled for rebuild. Please save your work now."
```

---

## Getting started

1. Clone the repository:

```bash
git clone https://github.com/AvrahamMakovsky/SmallScripts.git
cd SmallScripts
```

2. Run the relevant scripts using Python or PowerShell as appropriate.

---

## Notes

These scripts are small and focused. I use them as building blocks and utilities in my day-to-day work and personal projects, and they may evolve over time.

Author: Avraham Makovsky  
Jerusalem, Israel
