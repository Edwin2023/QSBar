@echo off
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "Register-QSBar.ps1" -Unregister
if %errorlevel% neq 0 pause
