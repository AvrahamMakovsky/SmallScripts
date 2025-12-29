@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ============================================================
rem ipconfig_release.bat
rem - Schedules reboot ASAP (5 seconds) so it will reboot even if
rem   remote connection drops during /release.
rem - Releases + renews DHCP
rem - Resets Winsock
rem
rem Author: Avraham Makovsky
rem ============================================================

rem ---- Require Admin ----
net session >nul 2>&1
if not "%errorlevel%"=="0" (
  echo [ERROR] Run as Administrator.
  exit /b 5
)

rem ---- Log file ----
set "TS="
for /f "usebackq delims=" %%I in (`powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss" 2^>nul`) do set "TS=%%I"
if not defined TS set "TS=%RANDOM%_%RANDOM%"
set "LOG=%TEMP%\ipconfig_release_%COMPUTERNAME%_%TS%.log"

call :log "============================================================"
call :log "ipconfig release - started (immediate reboot scheduled)"
call :log "Computer: %COMPUTERNAME%"
call :log "User: %USERNAME%"
call :log "Log: %LOG%"
call :log "============================================================"

rem ---- Schedule reboot NOW (ASAP) ----
shutdown /r /t 5 /c "Immediate reboot scheduled: ipconfig release/renew + winsock reset" >> "%LOG%" 2>&1
call :log "[INFO] Reboot scheduled in 5 seconds."

rem ---- Run commands quickly (best-effort) ----
call :run "ipconfig /release"   "Release DHCP lease"
call :run "ipconfig /renew"     "Renew DHCP lease"
call :run "netsh winsock reset" "Reset Winsock catalog"

call :log "[INFO] Commands issued. Host will reboot momentarily."
exit /b 0

rem ================== Helpers ==================
:run
set "CMD=%~1"
set "DESC=%~2"
call :log "---- %DESC% ----"
call :log "CMD: %CMD%"
cmd /c %CMD% >> "%LOG%" 2>&1
set "RC=%errorlevel%"
if not "%RC%"=="0" (
  call :log "[WARN] Exit code: %RC% (see log)"
) else (
  call :log "[OK] Completed"
)
exit /b 0

:log
echo %~1
>> "%LOG%" echo %~1
exit /b 0
