@echo off
cd /d "%~dp0"

set PYTHON=

:: --- 1. Check PATH, skip Windows Store stubs (they live under WindowsApps) ---
for /f "delims=" %%P in ('where py 2^>nul') do (
    echo %%P | findstr /i "WindowsApps" >nul || ( set PYTHON=%%P & goto :found )
)
for /f "delims=" %%P in ('where python 2^>nul') do (
    echo %%P | findstr /i "WindowsApps" >nul || ( set PYTHON=%%P & goto :found )
)

:: --- 2. Check common per-user install locations (no PowerShell needed) ---
for %%D in (
    "%LOCALAPPDATA%\Programs\Python\Python314"
    "%LOCALAPPDATA%\Programs\Python\Python313"
    "%LOCALAPPDATA%\Programs\Python\Python312"
    "%LOCALAPPDATA%\Programs\Python\Python311"
    "%LOCALAPPDATA%\Programs\Python\Python310"
    "%LOCALAPPDATA%\Programs\Python\Python39"
    "%LOCALAPPDATA%\Programs\Python\Python38"
) do if exist "%%~D\python.exe" ( set PYTHON=%%~D\python.exe & goto :found )

:: --- 3. Check common system-wide install locations ---
for %%D in (
    "C:\Python314"
    "C:\Python313"
    "C:\Python312"
    "C:\Python311"
    "C:\Python310"
    "C:\Python39"
    "C:\Python38"
    "C:\Program Files\Python314"
    "C:\Program Files\Python313"
    "C:\Program Files\Python312"
    "C:\Program Files\Python311"
    "C:\Program Files\Python310"
) do if exist "%%~D\python.exe" ( set PYTHON=%%~D\python.exe & goto :found )

echo.
echo  ERROR: Python was not found on this machine.
echo.
echo  Install from:  https://www.python.org/downloads/
echo  During setup, tick "Add Python to PATH"
echo.
echo  If Python is installed but you still see this:
echo    Settings ^> Apps ^> Advanced app settings ^> App execution aliases
echo    Turn OFF "App Installer - python.exe"
echo.
pause
exit /b 1

:found
echo Using: %PYTHON%
echo.

if "%~1"=="" (
    set /p FOLDER="Enter folder path: "
    "%PYTHON%" excelmerge.py "%FOLDER%"
) else (
    "%PYTHON%" excelmerge.py %*
)

echo.
pause
