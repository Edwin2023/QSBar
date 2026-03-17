@echo off
cd /d "%~dp0"
echo ----------------------------------------------------------------------
echo QSBar Registration Script
echo ----------------------------------------------------------------------
echo.

:: Detect if running as Administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [INFO] Running with Administrator privileges.
    set "ADMIN_SWITCH="
) else (
    echo [INFO] Running as Standard User.
    echo [INFO] Admin privileges are recommended for full system registration,
    echo [INFO] but basic user registration will proceed.
    set "ADMIN_SWITCH=-UserConfigOnly"
)

echo.
echo Registering QSBar...
echo.

powershell.exe -ExecutionPolicy Bypass -File "Register-QSBar.ps1" -DllPath "%~dp0QSBar.dll" -RestartApps %ADMIN_SWITCH%
if %errorlevel% neq 0 (
    echo [ERROR] Registration failed with exit code %errorlevel%
    :: Only pause if not running in silent/installer mode (detect by checking if %1 is -Silent)
    if /I "%~1" NEQ "-Silent" pause
    exit /b %errorlevel%
)

echo.
echo Registration Completed Successfully!
echo.
if /I "%~1" NEQ "-Silent" pause
