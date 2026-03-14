# QSBar Quick Setup Script for Excel and WPS
$ErrorActionPreference = "Stop"
$baseDir = $PSScriptRoot
$rootDir = (Get-Item $baseDir).Parent.FullName
$projectPath = Join-Path $rootDir "QSBar\QSBar.csproj"
$slnPath = Join-Path $rootDir "QSBar.sln"
$nugetExe = Join-Path $rootDir "nuget.exe"
$scriptsDir = $baseDir

# Find MSBuild
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuildPath = $null

if (Test-Path $vswhere) {
    $vsPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
    if ($vsPath) {
        $tryPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
        if (Test-Path $tryPath) { $msbuildPath = $tryPath }
        else {
            $tryPath = Join-Path $vsPath "MSBuild\15.0\Bin\MSBuild.exe"
            if (Test-Path $tryPath) { $msbuildPath = $tryPath }
        }
    }
}

if ([string]::IsNullOrEmpty($msbuildPath)) {
    $commonPaths = @(
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "D:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe",
        "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
    )
    foreach ($p in $commonPaths) {
        if (Test-Path $p) { $msbuildPath = $p; break }
    }
}

if ([string]::IsNullOrEmpty($msbuildPath)) {
    Write-Error "Cannot find MSBuild.exe."
    exit
}

Write-Host "Using MSBuild: $msbuildPath" -ForegroundColor Gray
Write-Host "--- 1. Cleaning Old Registrations ---" -ForegroundColor Cyan

# Clear Disabled Items
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

# Remove old COM keys
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

Write-Host "`n--- 2. Building QSBar ---" -ForegroundColor Cyan
if (Test-Path $msbuildPath) {
    Stop-Process -Name "et" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "wps" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "excel" -Force -ErrorAction SilentlyContinue

    & $msbuildPath $projectPath /p:Configuration=Debug /p:Platform="AnyCPU"
    Write-Host "Build successful!" -ForegroundColor Green
} else {
    Write-Error "MSBuild not found at $msbuildPath"
    exit
}

# 3. Registering
Write-Host "`n--- 3. Registering for WPS and Excel (COM) ---" -ForegroundColor Cyan
$dllPath = Join-Path $rootDir "QSBar\bin\Debug\QSBar.dll"
$registerScript = Join-Path $scriptsDir "Register-QSBar.ps1"
$regasm32 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
$regasm64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (Test-Path $dllPath) {
    # Delegate to Register-QSBar.ps1
    if (Test-Path $registerScript) {
        Write-Host "Calling Register-QSBar.ps1..." -ForegroundColor Cyan
        & powershell.exe -ExecutionPolicy Bypass -File $registerScript -DllPath $dllPath
    } else {
        Write-Error "CRITICAL ERROR: Register-QSBar.ps1 not found!"
        Write-Error "Please restore e:\Code\QSBar\scripts\Register-QSBar.ps1"
        exit 1
    }
} else {
    Write-Error "DLL not found after build!"
}

Write-Host "`n--- Setup Finished! ---" -ForegroundColor Green
Write-Host "1. Both Excel and WPS are now configured to use COM (not VSTO)."
Write-Host "2. Restart WPS/Excel to see the changes."
Write-Host "If the tab does not appear, check COM Add-ins settings."
