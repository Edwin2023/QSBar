# QSBar Quick Setup Script for Excel and WPS
$ErrorActionPreference = "Stop"
<<<<<<< HEAD
$baseDir = (Split-Path -Parent $PSScriptRoot)
$projectPath = Join-Path $baseDir "QSBar\QSBar.csproj"
$slnPath = Join-Path $baseDir "QSBar.sln"
$nugetExe = Join-Path $baseDir "nuget.exe"
=======
$baseDir = $PSScriptRoot
$rootDir = (Get-Item $baseDir).Parent.FullName
$projectPath = Join-Path $rootDir "QSBar\QSBar.csproj"
$slnPath = Join-Path $rootDir "QSBar.sln"
$nugetExe = Join-Path $rootDir "nuget.exe"
$scriptsDir = $baseDir
>>>>>>> fe804e7ec53608d2c57b7c870bc1bdf7383dccb5

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
    # Use existing Register-QSBar.ps1 script
    if (Test-Path $registerScript) {
        Write-Host "Calling Register-QSBar.ps1..." -ForegroundColor Cyan
        & powershell.exe -ExecutionPolicy Bypass -File $registerScript -DllPath $dllPath
        exit
    }
    
    # Fallback to manual registration if script missing (legacy code below)
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    [Environment]::SetEnvironmentVariable("VSTO_LOGALERTS", "1", "User")
    
    if ($isAdmin) {
        if (Test-Path $regasm32) {
            & $regasm32 /codebase "$dllPath" /tlb | Out-Null
        }
        if (Test-Path $regasm64) {
            & $regasm64 /codebase "$dllPath" /tlb | Out-Null
        }
    } else {
        Write-Warning "Not running as Administrator. Attempting to elevate RegAsm..."
        if (Test-Path $regasm32) {
            Start-Process $regasm32 -ArgumentList "/codebase ""$dllPath"" /tlb" -Verb RunAs -Wait
        }
        if (Test-Path $regasm64) {
            Start-Process $regasm64 -ArgumentList "/codebase ""$dllPath"" /tlb" -Verb RunAs -Wait
        }
    }

    $ProgID = "QSBar.WpsAddIn"
    $CLSID = "{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"
    $FriendlyName = "QSBar (COM)"
    $Description = "QSBar COM Add-in for Excel and WPS"

    $clsidRoot = "HKCU:\Software\Classes\CLSID\$CLSID"
    if (-not (Test-Path $clsidRoot)) { New-Item -Path $clsidRoot -Force | Out-Null }
    Set-ItemProperty -Path $clsidRoot -Name "(Default)" -Value "QSBar.WpsAddIn"
    
    $inproc = New-Item -Path "$clsidRoot\InprocServer32" -Force
    Set-ItemProperty -Path $inproc.PSPath -Name "(Default)" -Value "C:\Windows\System32\mscoree.dll"
    Set-ItemProperty -Path $inproc.PSPath -Name "ThreadingModel" -Value "Both"
    Set-ItemProperty -Path $inproc.PSPath -Name "Class" -Value "QSBar.WpsExcelAddIn"
    Set-ItemProperty -Path $inproc.PSPath -Name "Assembly" -Value "QSBar, Version=1.0.0.1, Culture=neutral, PublicKeyToken=null"
    Set-ItemProperty -Path $inproc.PSPath -Name "RuntimeVersion" -Value "v4.0.30319"
    Set-ItemProperty -Path $inproc.PSPath -Name "CodeBase" -Value "file:///$($dllPath.Replace('\', '/'))"

    $progIdKey = "HKCU:\Software\Classes\$ProgID"
    if (-not (Test-Path $progIdKey)) { New-Item -Path $progIdKey -Force | Out-Null }
    $curVer = New-Item -Path "$progIdKey\CLSID" -Force
    Set-ItemProperty -Path $curVer.PSPath -Name "(Default)" -Value $CLSID

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
        Remove-ItemProperty -Path $path -Name "Manifest" -ErrorAction SilentlyContinue
    }

    $wlProducts = @("ET", "WPS", "Common", "6.0")
    foreach ($prod in $wlProducts) {
        $wlPath = "HKCU:\Software\Kingsoft\Office\$prod\AddinsWL"
        if (-not (Test-Path $wlPath)) {
            New-Item -Path $wlPath -Force | Out-Null
        }
        Set-ItemProperty -Path $wlPath -Name $ProgID -Value "" -Force
    }

    Write-Host "Registration complete!" -ForegroundColor Green

    try {
        $testObj = New-Object -ComObject $ProgID -ErrorAction Stop
        Write-Host "SUCCESS: COM object created successfully!" -ForegroundColor Green
        $testObj = $null
    } catch {
        Write-Warning "FAILED: Could not create COM object."
        Write-Warning "Error: $($_.Exception.Message)"
    }
} else {
    Write-Error "DLL not found after build!"
}

Write-Host "`n--- Setup Finished! ---" -ForegroundColor Green
Write-Host "1. Both Excel and WPS are now configured to use COM (not VSTO)."
Write-Host "2. Restart WPS/Excel to see the changes."
Write-Host "If the tab does not appear, check COM Add-ins settings."
