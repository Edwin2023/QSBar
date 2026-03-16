
# Debug-QSBar-State.ps1
# Checks the state of QSBar registration for Excel and WPS

$ErrorActionPreference = "SilentlyContinue"

Write-Host "--- QSBar Diagnostic Report ---" -ForegroundColor Cyan

# 1. Check File Existence
$dllDebug = "E:\Code\qsbar\QSBar\bin\Debug\QSBar.dll"
$dllRelease = "E:\Code\qsbar\QSBar\bin\Release\QSBar.dll"
$dllInstall = "$env:LOCALAPPDATA\QSBar\QSBar.dll"

Write-Host "Checking DLL Files:"
if (Test-Path $dllDebug) { Write-Host "  [Debug]   Found: $dllDebug" -ForegroundColor Green } else { Write-Host "  [Debug]   Missing" -ForegroundColor Red }
if (Test-Path $dllRelease) { Write-Host "  [Release] Found: $dllRelease" -ForegroundColor Green } else { Write-Host "  [Release] Missing" -ForegroundColor Gray }
if (Test-Path $dllInstall) { Write-Host "  [Install] Found: $dllInstall" -ForegroundColor Green } else { Write-Host "  [Install] Missing" -ForegroundColor Gray }

# 2. Check Registry (HKCU - User)
$clsid = "{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"
$progId = "QSBar.WpsAddIn"

Write-Host "`nChecking HKCU (User) Registry:"
$hkcuClsid = Get-ItemProperty "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
if ($hkcuClsid) {
    Write-Host "  [CLSID] Found." -ForegroundColor Green
    Write-Host "    CodeBase: $($hkcuClsid.CodeBase)"
    Write-Host "    Assembly: $($hkcuClsid.Assembly)"
} else {
    Write-Host "  [CLSID] MISSING in HKCU!" -ForegroundColor Red
}

$wpsAddin = Get-ItemProperty "HKCU:\Software\Kingsoft\Office\WPS\Addins\$progId"
if ($wpsAddin) {
    Write-Host "  [WPS Addin] Found." -ForegroundColor Green
    Write-Host "    LoadBehavior: $($wpsAddin.LoadBehavior)"
} else {
    Write-Host "  [WPS Addin] MISSING in HKCU!" -ForegroundColor Red
}

$etAddin = Get-ItemProperty "HKCU:\Software\Kingsoft\Office\ET\Addins\$progId"
if ($etAddin) {
    Write-Host "  [ET Addin]  Found." -ForegroundColor Green
    Write-Host "    LoadBehavior: $($etAddin.LoadBehavior)"
} else {
    Write-Host "  [ET Addin]  MISSING in HKCU!" -ForegroundColor Red
}

# 3. Check HKLM Registry (Machine - 64-bit)
Write-Host "`nChecking HKLM (System) Registry:"
$hklmClsid = Get-ItemProperty "HKLM:\Software\Classes\CLSID\$clsid\InprocServer32"
if ($hklmClsid) {
    Write-Host "  [CLSID] Found (64-bit)." -ForegroundColor Yellow
    Write-Host "    CodeBase: $($hklmClsid.CodeBase)"
} else {
    Write-Host "  [CLSID] Not found (64-bit)." -ForegroundColor Gray
}

$hklmWps = Get-ItemProperty "HKLM:\Software\Kingsoft\Office\WPS\Addins\$progId"
if ($hklmWps) {
    Write-Host "  [WPS Addin] Found (64-bit). LoadBehavior: $($hklmWps.LoadBehavior)" -ForegroundColor Yellow
} else {
    Write-Host "  [WPS Addin] Not found (64-bit)." -ForegroundColor Gray
}

# 4. Check HKLM Registry (Machine - 32-bit/WOW64)
Write-Host "`nChecking HKLM (WOW6432Node) Registry:"
$wowClsid = Get-ItemProperty "HKLM:\Software\Classes\WOW6432Node\CLSID\$clsid\InprocServer32"
if ($wowClsid) {
    Write-Host "  [CLSID] Found (32-bit)." -ForegroundColor Yellow
    Write-Host "    CodeBase: $($wowClsid.CodeBase)"
} else {
    Write-Host "  [CLSID] Not found (32-bit)." -ForegroundColor Gray
}

$wowWps = Get-ItemProperty "HKLM:\Software\Wow6432Node\Kingsoft\Office\WPS\Addins\$progId"
if ($wowWps) {
    Write-Host "  [WPS Addin] Found (32-bit). LoadBehavior: $($wowWps.LoadBehavior)" -ForegroundColor Yellow
} else {
    Write-Host "  [WPS Addin] Not found (32-bit)." -ForegroundColor Gray
}

# 5. Check Disabled Items
Write-Host "`nChecking Disabled Items:"
$disabled = Get-ChildItem "HKCU:\Software\Kingsoft\Office\ET\Resiliency\DisabledItems"
if ($disabled) {
    Write-Host "  [WPS/ET] Found Disabled Items!" -ForegroundColor Red
    $disabled | Format-Table Name
} else {
    Write-Host "  [WPS/ET] Clean." -ForegroundColor Green
}

Write-Host "`n--- End Report ---"
