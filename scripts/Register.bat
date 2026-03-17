@echo off
cd /d "%~dp0"

:: Check for Administrator privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with Administrator privileges...
    goto :RunScript
) else (
    echo Requesting Administrator privileges...
    goto :Elevate
)

:Elevate
:: Create a temporary VBScript to launch the batch file as Administrator
echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
"%temp%\getadmin.vbs"
del "%temp%\getadmin.vbs"
exit /B

:RunScript
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
