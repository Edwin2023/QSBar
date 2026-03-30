@echo off
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "Build-Test-1.0.1.0.ps1"
