@echo off
cd /d "%~dp0"

:: Office is almost always installed as 32-bit (even on 64-bit Windows).
:: 64-bit PowerShell cannot drive a 32-bit Excel COM object -- it fails with
:: TYPE_E_ELEMENTNOTFOUND on every property access.
:: Solution: use the 32-bit PowerShell in SysWOW64 when it exists.

set PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
set PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe

if exist "%PS32%" (
    "%PS32%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0excelmerge.ps1" %*
) else (
    "%PS64%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0excelmerge.ps1" %*
)

echo.
pause
