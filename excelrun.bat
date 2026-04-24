@echo off
cd /d "%~dp0"

:: -----------------------------------------------------------------------
:: Find Python, skipping Windows Store app-execution-alias stubs.
:: Those stubs live under %LOCALAPPDATA%\Microsoft\WindowsApps and only
:: open the Store instead of running Python.
:: -----------------------------------------------------------------------
set PYTHON=

:: 1. Check PATH entries (py launcher or python), ignoring WindowsApps stubs
for /f "usebackq delims=" %%P in (
    `powershell -NoProfile -Command "Get-Command python,py -ErrorAction SilentlyContinue | Where-Object {$_.Source -notlike '*WindowsApps*'} | Select-Object -First 1 -ExpandProperty Source"`
) do set PYTHON=%%P

:: 2. Fall back: search the default per-user install location
if not defined PYTHON (
    for /f "usebackq delims=" %%P in (
        `powershell -NoProfile -Command "Get-ChildItem $env:LOCALAPPDATA\Programs\Python -Filter python.exe -Recurse -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName"`
    ) do set PYTHON=%%P
)

:: 3. Fall back: search Program Files
if not defined PYTHON (
    for /f "usebackq delims=" %%P in (
        `powershell -NoProfile -Command "Get-ChildItem 'C:\Program Files\Python*','C:\Python*' -Filter python.exe -Recurse -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName"`
    ) do set PYTHON=%%P
)

if not defined PYTHON (
    echo.
    echo ERROR: Python was not found on this machine.
    echo.
    echo  1. Download and install Python from  https://www.python.org/downloads/
    echo  2. During setup tick "Add Python to PATH"
    echo  3. Re-run this script
    echo.
    echo If Python IS installed but you still see this, also go to:
    echo  Settings ^> Apps ^> Advanced app settings ^> App execution aliases
    echo  and turn OFF "App Installer - python.exe"
    echo.
    pause
    exit /b 1
)

echo Using: %PYTHON%
echo.

:: -----------------------------------------------------------------------
:: Prompt for folder if the bat was double-clicked with no argument
:: -----------------------------------------------------------------------
if "%~1"=="" (
    set /p FOLDER="Enter folder path: "
    "%PYTHON%" excelmerge.py "%FOLDER%"
) else (
    "%PYTHON%" excelmerge.py %*
)

echo.
pause
