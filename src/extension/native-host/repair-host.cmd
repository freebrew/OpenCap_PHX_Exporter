@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-host.ps1" -Action Repair
echo.
pause
