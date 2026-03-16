@echo off
cd /d "%~dp0"
echo Requesting Administrator privileges to clean conflicting HKLM registry keys...
powershell -Command "Start-Process powershell -Verb RunAs -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0Register-QSBar.ps1\" -CleanHKLM'"
echo Cleanup command sent. Please accept the UAC prompt if it appears.
pause
