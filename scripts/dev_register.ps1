param(
    [string]$dllPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $dllPath)) {
    Write-Error "DLL not found at $dllPath"
    exit 1
}

$ProgID = "QSBar.WpsAddIn"
$CLSID = "{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"
$FriendlyName = "QSBar (COM)"
$Description = "QSBar COM Add-in for Excel and WPS"

Write-Host "--- Auto Register COM (Excel & WPS) ---" -ForegroundColor Cyan

# 1. Run RegAsm
$regasm32 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
$regasm64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (Test-Path $regasm32) {
    Write-Host "Registering 32-bit COM..."
    & $regasm32 /codebase "$dllPath" /tlb | Out-Null
}
if (Test-Path $regasm64) {
    Write-Host "Registering 64-bit COM..."
    & $regasm64 /codebase "$dllPath" /tlb | Out-Null
}

# 2. Force CLSID into HKCU
$clsidRoot = "HKCU:\Software\Classes\CLSID\$CLSID"
if (-not (Test-Path $clsidRoot)) { New-Item -Path $clsidRoot -Force | Out-Null }
Set-ItemProperty -Path $clsidRoot -Name "(Default)" -Value $ProgID

$inproc = New-Item -Path "$clsidRoot\InprocServer32" -Force
Set-ItemProperty -Path $inproc.PSPath -Name "(Default)" -Value "C:\Windows\System32\mscoree.dll"
Set-ItemProperty -Path $inproc.PSPath -Name "ThreadingModel" -Value "Both"
Set-ItemProperty -Path $inproc.PSPath -Name "Class" -Value "QSBar.WpsExcelAddIn"
Set-ItemProperty -Path $inproc.PSPath -Name "Assembly" -Value "QSBar, Version=1.0.0.1, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty -Path $inproc.PSPath -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty -Path $inproc.PSPath -Name "CodeBase" -Value "file:///$($dllPath.Replace('\', '/'))"

$progIdKey = "HKCU:\Software\Classes\$ProgID"
if (-not (Test-Path $progIdKey)) { New-Item -Path $progIdKey -Force | Out-Null }
$clsidKey = New-Item -Path "$progIdKey\CLSID" -Force
Set-ItemProperty -Path $clsidKey.PSPath -Name "(Default)" -Value $CLSID

# 3. Register Addin paths
$comRegPaths = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgID",
    "HKCU:\Software\Kingsoft\Office\ET\Addins\$ProgID",
    "HKCU:\Software\Kingsoft\Office\ET\AddinsData\$ProgID",
    "HKCU:\Software\Kingsoft\Office\WPS\Addins\$ProgID"
)

foreach ($path in $comRegPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "Description" -Value $Description -Force
    Set-ItemProperty -Path $path -Name "FriendlyName" -Value $FriendlyName -Force
    Set-ItemProperty -Path $path -Name "LoadBehavior" -Value 3 -Type DWord -Force
    Set-ItemProperty -Path $path -Name "CommandLineSafe" -Value 1 -Type DWord -Force
}

# 4. WPS Whitelist
$wlProducts = @("ET", "WPS", "Common", "6.0")
foreach ($prod in $wlProducts) {
    $wlPath = "HKCU:\Software\Kingsoft\Office\$prod\AddinsWL"
    if (-not (Test-Path $wlPath)) { New-Item -Path $wlPath -Force | Out-Null }
    Set-ItemProperty -Path $wlPath -Name $ProgID -Value "" -Force
}

Write-Host "Registration Complete! Please restart Excel or WPS." -ForegroundColor Green
