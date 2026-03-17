@echo off
cd /d "%~dp0"

echo ----------------------------------------------------------------------
echo QSBar Registration Script
echo ----------------------------------------------------------------------
echo.

powershell.exe -ExecutionPolicy Bypass -File "Register-QSBar.ps1" -DllPath "%~dp0QSBar.dll" -RestartApps

echo.
if %errorlevel% equ 0 (
    echo Registration Completed Successfully!
) else (
    echo Registration Completed with Errors (See above).
)
echo.
pause
