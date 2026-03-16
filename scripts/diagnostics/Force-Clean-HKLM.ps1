
# Force-Clean-HKLM.ps1
# This script forces the removal of all HKLM registry keys for QSBar to resolve conflicts.
# Must be run as Administrator.

$ErrorActionPreference = "Stop"

function Test-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit
}

Write-Host "--- Starting HKLM Cleanup for QSBar ---" -ForegroundColor Cyan

$guids = @("{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}")
$progId = "QSBar.WpsAddIn"

$paths = @(
    "HKLM:\Software\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}",
    "HKLM:\Software\Classes\$progId",
    "HKLM:\Software\Microsoft\Office\Excel\Addins\$progId",
    "HKLM:\Software\Kingsoft\Office\ET\Addins\$progId",
    "HKLM:\Software\Kingsoft\Office\WPS\Addins\$progId",
    "HKLM:\Software\Wow6432Node\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}",
    "HKLM:\Software\Wow6432Node\Classes\$progId",
    "HKLM:\Software\Wow6432Node\Microsoft\Office\Excel\Addins\$progId",
    "HKLM:\Software\Wow6432Node\Kingsoft\Office\ET\Addins\$progId",
    "HKLM:\Software\Wow6432Node\Kingsoft\Office\WPS\Addins\$progId"
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Write-Host "Removing: $path" -ForegroundColor Yellow
        try {
            Remove-Item -Path $path -Recurse -Force
            Write-Host "  [OK] Removed." -ForegroundColor Green
        } catch {
            Write-Host "  [ERROR] Failed to remove: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Skipping (Not Found): $path" -ForegroundColor Gray
    }
}

Write-Host "`n--- Cleanup Complete ---" -ForegroundColor Cyan
Write-Host "Please verify that the HKLM keys are gone."
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
