# QSBar Quick Setup Script for Excel & WPS
# This script restores dependencies, builds the project, and registers it for WPS.

$ErrorActionPreference = "Stop"
$baseDir = $PSScriptRoot
$projectPath = Join-Path $baseDir "QSBar\QSBar.csproj"
$slnPath = Join-Path $baseDir "QSBar.sln"
$nugetExe = Join-Path $baseDir "nuget.exe"
$msbuildPath = "D:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"

Write-Host "--- 1. Cleaning Old Registrations & Resiliency ---" -ForegroundColor Cyan

# 1.1 Clear Excel/WPS Disabled Items (To prevent "Hard Disable")
Write-Host "Clearing Disabled Items list..." -ForegroundColor Gray
$disabledPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Kingsoft\Office\ET\Resiliency\DisabledItems"
)
foreach ($dp in $disabledPaths) {
    if (Test-Path $dp) {
        Remove-Item -Path "$dp\*" -Force -ErrorAction SilentlyContinue
    }
}

# 1.7 Remove old COM keys
$progIDs = @("BMToolkits.WpsAddIn", "QSBar.WpsAddIn")
foreach ($p in $progIDs) {
    $oldPaths = @(
        "HKCU:\Software\Microsoft\Office\Excel\Addins\$p",
        "HKCU:\Software\Kingsoft\Office\ET\AddinsData\$p"
    )
    foreach ($oldPath in $oldPaths) {
        if (Test-Path $oldPath) {
            Remove-Item -Path $oldPath -Force -Recurse
            Write-Host "Removed COM key: $oldPath" -ForegroundColor Yellow
        }
    }
}

Write-Host "`n--- 2. Restoring NuGet Packages ---" -ForegroundColor Cyan
if (Test-Path $nugetExe) {
    & $nugetExe restore $slnPath -Source "https://api.nuget.org/v3/index.json"
} else {
    Write-Warning "nuget.exe not found. Skipping restore."
}

Write-Host "`n--- 2. Building QSBar ---" -ForegroundColor Cyan
if (Test-Path $msbuildPath) {
    # Kill any locking processes
    Stop-Process -Name "et" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "wps" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "excel" -Force -ErrorAction SilentlyContinue

    & $msbuildPath $projectPath /p:Configuration=Debug /p:Platform="AnyCPU" /p:RegisterForComInterop=false
    Write-Host "Build successful!" -ForegroundColor Green
} else {
    Write-Error "MSBuild not found at $msbuildPath"
    exit
}

# 3. Registering for WPS & Excel (COM)
Write-Host "`n--- 3. Registering for WPS & Excel (COM) ---" -ForegroundColor Cyan
$dllPath = Join-Path $baseDir "QSBar\bin\Debug\QSBar.dll"
$regasm32 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
$regasm64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (Test-Path $dllPath) {
    # 3.1 COM Registration via RegAsm
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    # Set environment variable to show Excel/VSTO errors
    [Environment]::SetEnvironmentVariable("VSTO_LOGALERTS", "1", "User")
    Write-Host "Set VSTO_LOGALERTS=1 to enable error reporting in Excel." -ForegroundColor Gray

    if ($isAdmin) {
        if (Test-Path $regasm32) {
            Write-Host "Registering 32-bit COM..."
            & $regasm32 /codebase "$dllPath" /tlb | Out-Null
        }
        if (Test-Path $regasm64) {
            Write-Host "Registering 64-bit COM..."
            & $regasm64 /codebase "$dllPath" /tlb | Out-Null
        }
    } else {
        Write-Warning "Not running as Administrator. Attempting to elevate RegAsm..."
        if (Test-Path $regasm32) {
            Start-Process $regasm32 -ArgumentList "/codebase `"$dllPath`" /tlb" -Verb RunAs -Wait
        }
        if (Test-Path $regasm64) {
            Start-Process $regasm64 -ArgumentList "/codebase `"$dllPath`" /tlb" -Verb RunAs -Wait
        }
    }

    $ProgID = "QSBar.WpsAddIn"
    $CLSID = "{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"
    $FriendlyName = "QSBar (COM)"
    $Description = "QSBar COM Add-in for Excel and WPS"

    # 3.2 Explicitly write CLSID to HKCU to ensure Excel/WPS visibility
    Write-Host "Forcing CLSID registration in HKCU..." -ForegroundColor Gray
    $clsidRoot = "HKCU:\Software\Classes\CLSID\$CLSID"
    if (-not (Test-Path $clsidRoot)) { New-Item -Path $clsidRoot -Force | Out-Null }
    Set-ItemProperty -Path $clsidRoot -Name "(Default)" -Value "QSBar.WpsAddIn"
    
    $inproc = New-Item -Path "$clsidRoot\InprocServer32" -Force
    Set-ItemProperty -Path $inproc.PSPath -Name "(Default)" -Value "C:\Windows\System32\mscoree.dll"
    Set-ItemProperty -Path $inproc.PSPath -Name "ThreadingModel" -Value "Both"
    Set-ItemProperty -Path $inproc.PSPath -Name "Class" -Value "QSBar.WpsExcelAddIn"
    Set-ItemProperty -Path $inproc.PSPath -Name "Assembly" -Value "QSBar, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    Set-ItemProperty -Path $inproc.PSPath -Name "RuntimeVersion" -Value "v4.0.30319"
    Set-ItemProperty -Path $inproc.PSPath -Name "CodeBase" -Value "file:///$($dllPath.Replace('\', '/'))"

    $progIdKey = "HKCU:\Software\Classes\$ProgID"
    if (-not (Test-Path $progIdKey)) { New-Item -Path $progIdKey -Force | Out-Null }
    $curVer = New-Item -Path "$progIdKey\CLSID" -Force
    Set-ItemProperty -Path $curVer.PSPath -Name "(Default)" -Value $CLSID

    # 3.3 Standard COM Add-in paths for WPS and Excel
    $comRegPaths = @(
        "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgID",
        "HKCU:\Software\Kingsoft\Office\ET\Addins\$ProgID",
        "HKCU:\Software\Kingsoft\Office\ET\AddinsData\$ProgID",
        "HKCU:\Software\Kingsoft\Office\WPS\Addins\$ProgID"
    )

    foreach ($path in $comRegPaths) {
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force | Out-Null
        }
        Set-ItemProperty -Path $path -Name "Description" -Value $Description -Force
        Set-ItemProperty -Path $path -Name "FriendlyName" -Value $FriendlyName -Force
        Set-ItemProperty -Path $path -Name "LoadBehavior" -Value 3 -Type DWord -Force
        Set-ItemProperty -Path $path -Name "CommandLineSafe" -Value 1 -Type DWord -Force
        # Remove Manifest key if it exists (from old VSTO attempts) to ensure it's treated as COM
        Remove-ItemProperty -Path $path -Name "Manifest" -ErrorAction SilentlyContinue
    }

    # 3.3 WPS Whitelist (Crucial for WPS)
    $wlProducts = @("ET", "WPS", "Common", "6.0")
    foreach ($prod in $wlProducts) {
        $wlPath = "HKCU:\Software\Kingsoft\Office\$prod\AddinsWL"
        if (-not (Test-Path $wlPath)) {
            New-Item -Path $wlPath -Force | Out-Null
        }
        Set-ItemProperty -Path $wlPath -Name $ProgID -Value "" -Force
        Write-Host "Added $ProgID to $prod whitelist" -ForegroundColor Gray
    }

    Write-Host "WPS & Excel Registration complete!" -ForegroundColor Green

    # 4. Diagnostic Test
    Write-Host "`n--- 4. Diagnostic Test ---" -ForegroundColor Cyan
    try {
        $testObj = New-Object -ComObject $ProgID -ErrorAction Stop
        Write-Host "SUCCESS: COM object '$ProgID' created successfully in PowerShell!" -ForegroundColor Green
        $testObj = $null
    } catch {
        Write-Warning "FAILED: Could not create COM object '$ProgID'. Excel will likely fail too."
        Write-Warning "Error: $($_.Exception.Message)"
    }
} else {
    Write-Error "DLL not found after build! Expected at $dllPath"
}

Write-Host "`n--- Setup Finished! ---" -ForegroundColor Green
Write-Host "1. Both Excel and WPS are now configured to use COM (not VSTO)."
Write-Host "2. For WPS: Restart WPS Spreadsheets."
Write-Host "3. For Excel: Restart Excel."
Write-Host "If the tab doesn't appear, check 'COM Add-ins' and ensure 'QSBar (COM)' is checked."
