@echo off
cd /d "%~dp0"
echo Running Quick Setup with ExecutionPolicy Bypass...
powershell.exe -ExecutionPolicy Bypass -File "quick_setup.ps1"
if %errorlevel% neq 0 (
    echo.
    echo Script failed with error code %errorlevel%.
    pause
) else (
    echo.
    echo Script finished successfully.
    pause
)
