@echo off
cd /d "%~dp0"
python excelmerge.py %*
echo.
pause
