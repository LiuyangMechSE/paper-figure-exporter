@echo off
rem Double-click to choose a PowerPoint file, or drag one or more .pptx files onto this icon.
rem The first run checks the computer and installs anything missing (Python + packages).
setlocal
title Paper Figure Exporter
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0engine\launcher.ps1" %*
exit /b %errorlevel%
