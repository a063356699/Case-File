@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "work\start-mobile-scan.ps1"
if errorlevel 1 pause
