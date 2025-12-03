@echo off
setlocal enabledelayedexpansion

:: BulkPingLauncher.bat
:: Author: Avraham Makovsky
:: Description:
::     Prompts user to paste a list of hostnames/IPs, then opens a separate ping window for each.
::     Uses Windows Terminal tabs if available, else falls back to classic CMD windows.
:: Supported delimiters: comma, space, newline
:: License: MIT / Public Domain

:: Check if Windows Terminal (wt.exe) is available
where wt >nul 2>nul
if %errorlevel% NEQ 0 goto FallbackCMD

:: Windows Terminal mode
echo Windows Terminal detected. Using multiple tabs.

:: Prompt user for input
echo Paste your hosts (separated by new lines, spaces, or commas), then SAVE and CLOSE Notepad. > hosts.tmp
start /wait notepad hosts.tmp

:: Gather and normalize input
set "rawHosts="
for /f "delims=" %%A in (hosts.tmp) do (
    set "rawHosts=!rawHosts! %%A "
)
set "rawHosts=%rawHosts:,= %"

:: Parse and store each host
for %%A in (%rawHosts%) do (
    set /a count+=1
    set "host!count!=%%A"
)

:: Exit if nothing was provided
if %count% EQU 0 (
    echo No hosts provided. Exiting...
    del hosts.tmp
    exit /b
)

:: Build wt command
set "wtCommand=wt -w 0 new-tab --title Host_1 cmd /k ping -t !host1!"
for /L %%i in (2,1,%count%) do (
    set "currentHost=!host%%i!"
    set "wtCommand=!wtCommand! ; new-tab --title Host_%%i cmd /k ping -t !currentHost!"
)

:: Execute
start "" %wtCommand%
del hosts.tmp
exit

:: Fallback: Classic CMD for older Windows
:FallbackCMD
echo Windows Terminal not found. Using separate CMD windows.
echo Paste your hosts (separated by new lines, spaces, or commas), then SAVE and CLOSE Notepad. > hosts.tmp
start /wait notepad hosts.tmp

set "rawHosts="
for /f "delims=" %%A in (hosts.tmp) do (
    set "rawHosts=!rawHosts! %%A "
)
set "rawHosts=%rawHosts:,= %"

for %%A in (%rawHosts%) do (
    start cmd /k ping -t %%A
)

del hosts.tmp
exit
