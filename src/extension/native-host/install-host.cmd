@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-host.ps1" -Action Install
echo.
echo If Chrome was already open, click Recheck in OpenCap Slidesheet after this window closes.
pause
