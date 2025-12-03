# SmallScripts

Small collection of utility scripts that I use for everyday tasks, experiments and technical work.

## About this repository

These scripts started as personal side projects that I wrote on my own time and hardware.  
They are general purpose tools that can be useful at home or in professional environments  
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

**What it does**

- Reads SMART data from local drives (via `smartctl` from smartmontools).
- Shows key information such as:
  - Drive model and serial  
  - Capacity  
  - Temperature  
  - Power-on hours  
  - Basic health indicators (for example, reallocated or pending sectors, when available).
- Prints a compact summary in the terminal so you can quickly see the state of your drives.

**How to use**

1. Install Python 3.
2. Install `smartctl` (from smartmontools) and make sure it is available in your system PATH.
3. Run the script from a terminal in the repository folder:

   ```bash
   python SSDHealthMonitor.py

   extractor.py

### Hostname extractor for ticket titles / text (Python).

What it does

Scans a text file that contains ticket titles or other lines of text.

Searches for hostnames that match patterns like ST01****
(for example: ST01WVAW0123 or similar station / lab hostnames).

Outputs a clean list of all hostnames it finds, which can then be used in other tools or scripts.


Typical use case

You have many tickets exported to a text file.

You want to quickly collect all hostnames mentioned in those tickets.

You run the script once and get a plain list of hosts.


How to use

1. Put your ticket titles or text into a simple .txt file (one title or line per row).


2. Place that file in the same folder as the script (or use the file name expected by the script).


3. Run the script from a terminal in the repository folder:

python extractor.py




---

Getting started

1. Clone the repository:

git clone https://github.com/AvrahamMakovsky/SmallScripts.git
cd SmallScripts


2. Run the scripts with Python as shown above.



PowerShell or batch scripts that may be added in the future will include short usage notes either here or in their headers.


---

Notes

These scripts are small and focused. I use them as building blocks and utilities in my day-to-day work and personal projects, and they may evolve over time.

Author: Avraham Makovsky
