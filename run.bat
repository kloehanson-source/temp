@echo off
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "flatten_merge.ps1"
echo.
pause
