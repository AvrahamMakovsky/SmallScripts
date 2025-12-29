# SmallScripts

Small collection of utility scripts that I use for everyday tasks, experiments, and technical work.

## About this repository

These scripts started as personal side projects that I wrote on my own time and hardware.  
They are general-purpose tools that can be useful at home or in professional environments  
(for example, validation labs, support work, or general system maintenance).

They are **not official tools of any employer**, past or present.

Over time I may add scripts in different languages, such as:

- Python
- PowerShell
- Windows batch
- Other small tools and helpers

---

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

### ipconfig release.bat

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
3. If the filename contains a space, run it with quotes:

```bat
".\ipconfig release.bat"
```

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
