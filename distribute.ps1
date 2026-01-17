# QSBar Distribution Script
# This script builds the project and prepares a 'dist' folder for the installer.

$ErrorActionPreference = "Stop"
$baseDir = $PSScriptRoot
$msbuildPath = "D:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$distDir = Join-Path $baseDir "dist"

Write-Host "--- 1. Cleaning and Creating dist folder ---" -ForegroundColor Cyan
if (Test-Path $distDir) { Remove-Item $distDir -Recurse -Force }
New-Item -ItemType Directory -Path $distDir | Out-Null

Write-Host "`n--- 2. Building QSBar (Add-in) ---" -ForegroundColor Cyan
& $msbuildPath (Join-Path $baseDir "QSBar\QSBar.csproj") /p:Configuration=Release /p:Platform="AnyCPU" /p:RegisterForComInterop=false

Write-Host "`n--- 3. Building QSBar.Setup (Installer) ---" -ForegroundColor Cyan
& $msbuildPath (Join-Path $baseDir "QSBar.Setup\QSBar.Setup.csproj") /p:Configuration=Release /p:Platform="AnyCPU"

Write-Host "`n--- 4. Preparing dist folder ---" -ForegroundColor Cyan
# Copy Installer
Copy-Item (Join-Path $baseDir "QSBar.Setup\bin\Release\QSBar.Setup.exe") (Join-Path $distDir "Install_QSBar.exe")

# Copy Add-in and dependencies
$payloadDir = New-Item -ItemType Directory -Path (Join-Path $distDir "payload")
Copy-Item (Join-Path $baseDir "QSBar\bin\Release\QSBar.dll") $payloadDir
Copy-Item (Join-Path $baseDir "QSBar\bin\Release\*.dll") $payloadDir -Exclude "QSBar.dll"

Write-Host "`nDistribution package created in '$distDir'!" -ForegroundColor Green
Write-Host "To install: Run dist\Install_QSBar.exe"
Write-Host "To uninstall: Run dist\Install_QSBar.exe --uninstall"
