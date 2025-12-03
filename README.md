# SmallScripts

Small collection of utility scripts that I use for everyday tasks, experiments, and technical work.

## About this repository

These scripts started as personal side projects that I wrote on my own time and hardware.  
They are general-purpose tools that can be useful at home or in professional environments  
(for example, validation labs, support work, or general system maintenance).

They are not official tools of any employer, past or present.

Over time I may add scripts in different languages, such as:

- Python  
- PowerShell  
- Windows batch  
- Other small tools and helpers  

---

## Scripts

### `SSDHealthMonitor.py`  
Command-line SSD health monitor (Python).

#### What it does

- Reads SMART data from local drives (via `smartctl` from smartmontools).
- Shows key information such as:
  - Drive model and serial  
  - Capacity  
  - Temperature  
  - Power-on hours  
  - Basic health indicators (e.g. reallocated or pending sectors).
- Prints a compact summary in the terminal so you can quickly see the state of your drives.

#### How to use

1. Install Python 3.
2. Install `smartctl` (from smartmontools) and ensure it's in your system PATH.
3. Run the script:

   ```bash
   python SSDHealthMonitor.py
   ```

---

### `extractor.py`  
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

### `BulkPingLauncher.bat`  
Paste a list of hosts and open a continuous `ping -t` for each oneвЂ”either in **multiple tabs** (Windows Terminal) or separate CMD windows.

#### What it does

- Opens a Notepad window for you to paste a list of hostnames (delimited by space, comma, or newline).
- If Windows Terminal (`wt.exe`) is found, each host is opened in a new tab with its own `ping -t`.
- If not found, falls back to opening a classic CMD window for each host.

#### Use case

- Monitor connectivity of several servers, lab machines, or kiosks.
- Fast bulk diagnostics.

---

### `BulkVNCLauncher.bat`  
Paste a list of hosts and launch RealVNC Viewer sessions for all of them.

#### What it does

- Opens Notepad for list entry (hosts separated by space, comma, or newline).
- Starts a RealVNC session for each host using the installed `vncviewer.exe`.
- Requires RealVNC Viewer to be installed (path adjustable in script).

#### Use case

- Bulk-connect to lab stations or remote machines for access and support.

---

## Getting started

1. Clone the repository:

   ```bash
   git clone https://github.com/AvrahamMakovsky/SmallScripts.git
   cd SmallScripts
   ```

2. Run the relevant scripts using Python or double-click batch files as needed.

---

## Notes

These scripts are small and focused. I use them as building blocks and utilities in my day-to-day work and personal projects, and they may evolve over time.

Author: **Avraham Makovsky**  
Jerusalem, Israel
