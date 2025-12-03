@echo off
setlocal enabledelayedexpansion

:: BulkVNCLauncher.bat
:: Author: Avraham Makovsky
:: Description:
::     Prompts user to paste a list of hostnames/IPs, then opens a separate RealVNC session for each.
::     Each VNC window launches independently (no need for wt.exe)
:: Supported delimiters: comma, space, newline
:: License: MIT / Public Domain

:: Path to VNC Viewer (update if needed)
set "VNC_PATH=C:\Program Files\RealVNC\VNC Viewer\vncviewer.exe"

if not exist "%VNC_PATH%" (
    echo VNC Viewer not found! Please update the script path.
    pause
    exit /b
)

:: Prompt user for hosts
echo Paste your VNC hosts (separated by new lines, spaces, or commas), then SAVE and CLOSE Notepad. > hosts.tmp
start /wait notepad hosts.tmp

:: Normalize input
set "rawHosts="
for /f "delims=" %%A in (hosts.tmp) do (
    set "rawHosts=!rawHosts! %%A "
)
set "rawHosts=%rawHosts:,= %"

:: Launch VNC sessions
for %%A in (%rawHosts%) do (
    echo Opening VNC session for: %%A
    start "" "%VNC_PATH%" %%A
)

del hosts.tmp
exit
