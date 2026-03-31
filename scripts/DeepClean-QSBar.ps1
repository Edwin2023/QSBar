param(
    [switch]$Interactive = $true
)

$ErrorActionPreference = "Continue"

$CLSID = "{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"
$ProgID = "QSBar.WpsAddIn"

# Helper Function: Check Administrator
function Test-IsAdmin {
    return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Remove-RegistryKeySafely {
    param([string]$Path)
    if (Test-Path $Path) {
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            Write-Host "[OK] Removed: $Path" -ForegroundColor Green
        } catch {
            Write-Host "[FAILED] Could not remove: $Path" -ForegroundColor Red
            Write-Host "         $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "[-] Not found (Skipped): $Path" -ForegroundColor DarkGray
    }
}

function Remove-RegistryValueSafely {
    param([string]$Path, [string]$Name)
    if (Test-Path $Path) {
        $val = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        if ($null -ne $val) {
            try {
                Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction Stop
                Write-Host "[OK] Removed Value: $Path\$Name" -ForegroundColor Green
            } catch {
                Write-Host "[FAILED] Could not remove value: $Path\$Name" -ForegroundColor Red
            }
        } else {
            Write-Host "[-] Value Not found (Skipped): $Path\$Name" -ForegroundColor DarkGray
        }
    }
}

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "      QSBar Deep Uninstaller & Registry Cleaner" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

$IsAdmin = Test-IsAdmin
if (-not $IsAdmin) {
    Write-Host "WARNING: You are NOT running as Administrator!" -ForegroundColor Yellow
    Write-Host "HKLM (System-wide) registry keys cannot be cleaned." -ForegroundColor Yellow
    Write-Host "Please right-click and 'Run as Administrator' for a complete clean." -ForegroundColor Yellow
    Write-Host ""
} else {
    Write-Host "Running as Administrator. Full system clean enabled." -ForegroundColor Green
    Write-Host ""
}

Write-Host "`n--- Step 1: Closing Office Applications ---" -ForegroundColor Cyan
$procs = @("excel", "wps", "et")
foreach ($proc in $procs) {
    $running = Get-Process -Name $proc -ErrorAction SilentlyContinue
    if ($running) {
        Write-Host "Stopping $proc..."
        Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
    }
}
Start-Sleep -Seconds 1

Write-Host "`n--- Step 2: Cleaning Current User (HKCU) Registry ---" -ForegroundColor Cyan

# CLSID and ProgID
Remove-RegistryKeySafely "HKCU:\Software\Classes\CLSID\$CLSID"
Remove-RegistryKeySafely "HKCU:\Software\Classes\$ProgID"
Remove-RegistryKeySafely "HKCU:\Software\Classes\Wow6432Node\CLSID\$CLSID"
Remove-RegistryKeySafely "HKCU:\Software\Classes\Wow6432Node\$ProgID"

# Addins (Excel & WPS)
Remove-RegistryKeySafely "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgID"
Remove-RegistryKeySafely "HKCU:\Software\Kingsoft\Office\ET\Addins\$ProgID"
Remove-RegistryKeySafely "HKCU:\Software\Kingsoft\Office\ET\AddinsData\$ProgID"
Remove-RegistryKeySafely "HKCU:\Software\Kingsoft\Office\WPS\Addins\$ProgID"

# WPS AddinsWL Values
Remove-RegistryValueSafely "HKCU:\Software\Kingsoft\Office\ET\AddinsWL" $ProgID
Remove-RegistryValueSafely "HKCU:\Software\Kingsoft\Office\WPS\AddinsWL" $ProgID
Remove-RegistryValueSafely "HKCU:\Software\Kingsoft\Office\Common\AddinsWL" $ProgID
Remove-RegistryValueSafely "HKCU:\Software\Kingsoft\Office\6.0\AddinsWL" $ProgID

# Clean Legacy/Alternate ProgIDs
$legacyProgIDs = @("BMToolkits.WpsAddIn")
foreach ($p in $legacyProgIDs) {
    Remove-RegistryKeySafely "HKCU:\Software\Microsoft\Office\Excel\Addins\$p"
    Remove-RegistryKeySafely "HKCU:\Software\Kingsoft\Office\ET\Addins\$p"
    Remove-RegistryKeySafely "HKCU:\Software\Kingsoft\Office\ET\AddinsData\$p"
}


Write-Host "`n--- Step 3: Cleaning Local Machine (HKLM) Registry ---" -ForegroundColor Cyan
if ($IsAdmin) {
    # 64-bit / Native Registry
    Remove-RegistryKeySafely "HKLM:\Software\Classes\CLSID\$CLSID"
    Remove-RegistryKeySafely "HKLM:\Software\Classes\$ProgID"
    Remove-RegistryKeySafely "HKLM:\Software\Microsoft\Office\Excel\Addins\$ProgID"
    
    # 32-bit (WOW6432Node) Registry
    Remove-RegistryKeySafely "HKLM:\Software\WOW6432Node\Classes\CLSID\$CLSID"
    Remove-RegistryKeySafely "HKLM:\Software\WOW6432Node\Classes\$ProgID"
    Remove-RegistryKeySafely "HKLM:\Software\WOW6432Node\Microsoft\Office\Excel\Addins\$ProgID"
} else {
    Write-Host "[SKIP] Skipping HKLM cleanup due to lack of Administrator privileges." -ForegroundColor Yellow
}

# Inno Setup Uninstall Registry Keys
$InnoAppId = "{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}_is1"
Remove-RegistryKeySafely "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$InnoAppId"
Remove-RegistryKeySafely "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$InnoAppId"
Remove-RegistryKeySafely "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$InnoAppId"

Write-Host "`n--- Step 4: Attempting to Unregister using RegAsm ---" -ForegroundColor Cyan
$dllPaths = @(
    "$PSScriptRoot\..\QSBar\bin\Release\QSBar.dll",
    "$PSScriptRoot\..\Release\QSBar.dll",
    "$env:LOCALAPPDATA\QSBar\QSBar.dll"
)

$dllFound = $false
foreach ($path in $dllPaths) {
    if (Test-Path $path) {
        $dllFound = $true
        Write-Host "Found DLL at: $path. Attempting RegAsm /unregister..." -ForegroundColor Cyan
        
        $regasm32 = "$env:windir\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
        $regasm64 = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

        if (Test-Path $regasm32) {
            Start-Process -FilePath $regasm32 -ArgumentList "/unregister /codebase `"$path`"" -Wait -NoNewWindow -ErrorAction SilentlyContinue
        }
        if (Test-Path $regasm64) {
            Start-Process -FilePath $regasm64 -ArgumentList "/unregister /codebase `"$path`"" -Wait -NoNewWindow -ErrorAction SilentlyContinue
        }
        break # Only unregister the first found
    }
}

if (-not $dllFound) {
    Write-Host "No QSBar.dll found to run RegAsm against. Relying entirely on manual registry cleanup." -ForegroundColor Yellow
}

Write-Host "`n--- Step 5: Removing Physical Files and Folders ---" -ForegroundColor Cyan
$installDir = "$env:LOCALAPPDATA\QSBar"
if (Test-Path $installDir) {
    try {
        Remove-Item -Path $installDir -Recurse -Force -ErrorAction Stop
        Write-Host "[OK] Removed installation directory: $installDir" -ForegroundColor Green
    } catch {
        Write-Host "[FAILED] Could not remove directory: $installDir" -ForegroundColor Red
        Write-Host "         $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "[-] Installation directory not found (Skipped): $installDir" -ForegroundColor DarkGray
}

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "          Cleanup Completed Successfully!                 " -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Cyan
if ($Interactive) {
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
