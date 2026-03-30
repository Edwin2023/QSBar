@echo off
cd /d "%~dp0"
echo ----------------------------------------------------------------------
echo Requesting Administrator Privileges for Deep Clean...
echo ----------------------------------------------------------------------

:: Check for admin rights
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Administrator rights confirmed. Proceeding with deep clean...
    powershell.exe -ExecutionPolicy Bypass -File "DeepClean-QSBar.ps1"
) else (
    echo Elevating to Administrator...
    powershell.exe -Command "Start-Process cmd -ArgumentList '/c %~dp0DeepClean-QSBar.bat' -Verb RunAs"
)
