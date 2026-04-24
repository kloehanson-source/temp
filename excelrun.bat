@echo off
cd /d "%~dp0"

:: Use the Windows Python Launcher (py.exe) which is registered system-wide.
:: Fall back to python if py is not available.
where py >nul 2>&1 && set PYTHON=py || set PYTHON=python

:: If no folder was passed (e.g. double-clicked), prompt for one.
if "%~1"=="" (
    set /p FOLDER="Enter folder path: "
) else (
    set FOLDER=%*
)

%PYTHON% excelmerge.py "%FOLDER%"
echo.
pause
