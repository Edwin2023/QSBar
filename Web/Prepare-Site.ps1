# Prepare-Site.ps1
# Automate the preparation of website files

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$OutputDir = Join-Path $ScriptDir "output"
$InstallerDir = Join-Path $ScriptDir "..\Installer"

# 1. Create Output Directory
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# 2. Copy Web Files
Copy-Item (Join-Path $ScriptDir "index.html") -Destination (Join-Path $OutputDir "index.html") -Force

# 3. Find and Copy Latest Installer
$pattern = Join-Path $InstallerDir "QSBar_Setup_v*.exe"
$latestInstaller = Get-ChildItem $pattern | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($latestInstaller) {
    Copy-Item $latestInstaller.FullName -Destination (Join-Path $OutputDir "QSBar_Setup.exe") -Force
    Write-Host "Success: Copied installer $($latestInstaller.Name) -> QSBar_Setup.exe" -ForegroundColor Green
} else {
    Write-Host "Warning: Installer not found (QSBar_Setup_v*.exe)" -ForegroundColor Yellow
    Write-Host "Search Path: $InstallerDir" -ForegroundColor Gray
}

Write-Host "`nWebsite files are ready!" -ForegroundColor Cyan
Write-Host "Location: $OutputDir"
Write-Host "Please upload all files in this directory to your Alibaba Cloud server."
