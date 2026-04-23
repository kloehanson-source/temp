@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0flatten_merge.ps1"
if %errorlevel% neq 0 (
    echo.
    echo Script exited with an error. See red text above.
    pause
)
