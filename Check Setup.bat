@echo off
rem Checks Windows, PowerPoint and Python, and installs anything missing. Exports nothing.
rem To rebuild the Python environment from scratch, run:  "Check Setup.bat" -Repair
setlocal
title Paper Figure Exporter - setup check
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0engine\launcher.ps1" -CheckOnly %*
exit /b %errorlevel%
