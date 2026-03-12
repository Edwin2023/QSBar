param(
    [string]$DllPath,
    [switch]$CleanHKLM,
    [switch]$RestartApps,
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"
$CLSID = "{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"
$ProgID = "QSBar.WpsAddIn"
$FriendlyName = "QSBar (COM)"
$Description = "QSBar COM Add-in for Excel and WPS"

# --- Helper Function: Check Administrator ---
function Test-IsAdmin {
    return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# --- Task 0: Clean Disabled Items (Fix silent load failures) ---
if (-not $Unregister) {
    Write-Host "--- Checking for Disabled Items ---" -ForegroundColor Cyan
    $disabledPaths = @(
        "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
        "HKCU:\Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems",
        "HKCU:\Software\Kingsoft\Office\ET\Resiliency\DisabledItems"
    )
    foreach ($dp in $disabledPaths) {
        if (Test-Path $dp) {
            Remove-Item -Path "$dp\*" -Force -ErrorAction SilentlyContinue
            Write-Host "Cleaned DisabledItems in: $dp" -ForegroundColor Gray
        }
    }
}

# --- Task 1: Clean HKLM (Optional) ---
if ($CleanHKLM) {
    Write-Host "--- Cleaning HKLM Registration ---" -ForegroundColor Cyan
    if (-not (Test-IsAdmin)) {
        Write-Warning "Cleaning HKLM requires Administrator privileges."
        Write-Warning "Please run this script as Administrator to use -CleanHKLM."
    } else {
        $hklmPath = "HKLM:\Software\Classes\CLSID\$CLSID"
        if (Test-Path $hklmPath) {
            Write-Host "Removing HKLM CLSID key..." -ForegroundColor Yellow
            Remove-Item -Path $hklmPath -Recurse -Force
            Write-Host "HKLM Cleanup complete." -ForegroundColor Green
        } else {
            Write-Host "No HKLM registration found." -ForegroundColor Gray
        }
    }
    if ($Unregister -and -not $DllPath) { exit }
}

# --- Task 2: Restart Apps (Optional) ---
if ($RestartApps) {
    Write-Host "--- Restarting Office Applications ---" -ForegroundColor Cyan
    $apps = @("EXCEL", "wps", "et")
    foreach ($app in $apps) {
        Get-Process -Name $app -ErrorAction SilentlyContinue | Stop-Process -Force
        Write-Host "Stopped $app" -ForegroundColor DarkGray
    }
}

# --- Resolve DLL Path ---
if (-not $DllPath) {
    # 1. Look in script directory (Deployment scenario)
    $DllPath = Join-Path $PSScriptRoot "QSBar.dll"
    
    # 2. If not found, look in project output (Dev scenario)
    if (-not (Test-Path $DllPath)) {
        $devPath = Join-Path $PSScriptRoot "..\QSBar\bin\Debug\QSBar.dll"
        if (Test-Path $devPath) { 
            $DllPath = $devPath 
            Write-Host "Found DLL in Debug folder: $DllPath" -ForegroundColor Gray
        }
    }
}

if (-not (Test-Path $DllPath)) {
    if (-not $Unregister) {
        Write-Error "DLL not found at: $DllPath"
        exit 1
    }
} else {
    $DllPath = (Get-Item $DllPath).FullName
}

# --- Task 3: Unregister (if requested) ---
if ($Unregister) {
    Write-Host "--- Unregistering ---" -ForegroundColor Cyan
    
    # Remove HKCU Keys
    $hkcuClsid = "HKCU:\Software\Classes\CLSID\$CLSID"
    if (Test-Path $hkcuClsid) { Remove-Item -Path $hkcuClsid -Recurse -Force; Write-Host "Removed HKCU CLSID" }
    
    $hkcuProg = "HKCU:\Software\Classes\$ProgID"
    if (Test-Path $hkcuProg) { Remove-Item -Path $hkcuProg -Recurse -Force; Write-Host "Removed HKCU ProgID" }
    
    # Remove Addin Keys
    $comRegPaths = @(
        "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgID",
        "HKCU:\Software\Kingsoft\Office\ET\Addins\$ProgID",
        "HKCU:\Software\Kingsoft\Office\ET\AddinsData\$ProgID",
        "HKCU:\Software\Kingsoft\Office\WPS\Addins\$ProgID"
    )
    foreach ($path in $comRegPaths) {
        if (Test-Path $path) { Remove-Item -Path $path -Recurse -Force; Write-Host "Removed Addin Key: $path" }
    }

    # Clean legacy/alternate ProgIDs
    $legacyProgIDs = @("BMToolkits.WpsAddIn")
    foreach ($p in $legacyProgIDs) {
        $legacyPaths = @(
            "HKCU:\Software\Microsoft\Office\Excel\Addins\$p",
            "HKCU:\Software\Kingsoft\Office\ET\Addins\$p",
            "HKCU:\Software\Kingsoft\Office\ET\AddinsData\$p"
        )
        foreach ($lp in $legacyPaths) {
            if (Test-Path $lp) { Remove-Item -Path $lp -Recurse -Force; Write-Host "Removed Legacy Key: $lp" }
        }
    }

    Write-Host "Unregistration Complete." -ForegroundColor Green
    exit
}

# --- Task 4: Registration (Default) ---
Write-Host "--- Registering QSBar ($DllPath) ---" -ForegroundColor Cyan

# 1. Run RegAsm (Updates HKCR, might fail if no Admin, but we rely on manual HKCU mostly)
$regasm32 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
$regasm64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (Test-Path $regasm32) {
    Write-Host "Running RegAsm (32-bit)..."
    $output = & $regasm32 /codebase "$DllPath" /tlb 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "RegAsm (32-bit) FAILED (Likely due to no Admin rights)."
        Write-Warning "Proceeding with HKCU registration (Sufficient for User/Dev mode)..."
    }
}
if (Test-Path $regasm64) {
    Write-Host "Running RegAsm (64-bit)..."
    $output = & $regasm64 /codebase "$DllPath" /tlb 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "RegAsm (64-bit) FAILED (Likely due to no Admin rights)."
        Write-Warning "Proceeding with HKCU registration (Sufficient for User/Dev mode)..."
    }
}

# 2. Check for HKLM Conflict
$hklmInproc = "HKLM:\Software\Classes\CLSID\$CLSID\InprocServer32"
if (Test-Path $hklmInproc) {
    $hklmCodeBase = Get-ItemProperty -Path $hklmInproc -Name "CodeBase" -ErrorAction SilentlyContinue
    if ($hklmCodeBase -and $hklmCodeBase.CodeBase -ne "file:///$($DllPath.Replace('\', '/'))") {
        Write-Warning "CONFLICT DETECTED: HKLM registration points to a different DLL!"
        Write-Warning "Current HKLM: $($hklmCodeBase.CodeBase)"
        Write-Warning "Target DLL:   file:///$($DllPath.Replace('\', '/'))"
        Write-Warning "This will likely prevent Excel from loading your changes."
        Write-Warning "Run this script with '-CleanHKLM' as Administrator to fix this."
    }
}

# 3. Force CLSID into HKCU (User Registration)
try {
    $asmName = [System.Reflection.AssemblyName]::GetAssemblyName($DllPath)
    $asmVersion = $asmName.Version.ToString()
    $pktBytes = $asmName.GetPublicKeyToken()
    if ($pktBytes) {
        $pkt = [BitConverter]::ToString($pktBytes).Replace("-", "").ToLower()
    } else {
        $pkt = "null"
    }
    $fullAsmName = "$($asmName.Name), Version=$asmVersion, Culture=$($asmName.CultureInfo.Name), PublicKeyToken=$pkt"
    if ($asmName.CultureInfo.Name -eq "") { $fullAsmName = $fullAsmName.Replace("Culture=", "Culture=neutral") }
} catch {
    Write-Warning "Failed to read assembly metadata. Using default fallback."
    $fullAsmName = "QSBar, Version=1.0.0.1, Culture=neutral, PublicKeyToken=null"
}

$clsidRoot = "HKCU:\Software\Classes\CLSID\$CLSID"
if (-not (Test-Path $clsidRoot)) { New-Item -Path $clsidRoot -Force | Out-Null }
Set-ItemProperty -Path $clsidRoot -Name "(Default)" -Value $ProgID

$inproc = New-Item -Path "$clsidRoot\InprocServer32" -Force
Set-ItemProperty -Path $inproc.PSPath -Name "(Default)" -Value "C:\Windows\System32\mscoree.dll"
Set-ItemProperty -Path $inproc.PSPath -Name "ThreadingModel" -Value "Both"
Set-ItemProperty -Path $inproc.PSPath -Name "Class" -Value "QSBar.WpsExcelAddIn"
Set-ItemProperty -Path $inproc.PSPath -Name "Assembly" -Value $fullAsmName
Set-ItemProperty -Path $inproc.PSPath -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty -Path $inproc.PSPath -Name "CodeBase" -Value "file:///$($DllPath.Replace('\', '/'))"

$progIdKey = "HKCU:\Software\Classes\$ProgID"
if (-not (Test-Path $progIdKey)) { New-Item -Path $progIdKey -Force | Out-Null }
$clsidKey = New-Item -Path "$progIdKey\CLSID" -Force
Set-ItemProperty -Path $clsidKey.PSPath -Name "(Default)" -Value $CLSID

# 4. Register Addin paths (Excel & WPS)
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
    # Remove Manifest to prevent VSTO conflict
    Remove-ItemProperty -Path $path -Name "Manifest" -ErrorAction SilentlyContinue
}

# 5. WPS Whitelist
$wlProducts = @("ET", "WPS", "Common", "6.0")
foreach ($prod in $wlProducts) {
    $wlPath = "HKCU:\Software\Kingsoft\Office\$prod\AddinsWL"
    if (-not (Test-Path $wlPath)) { New-Item -Path $wlPath -Force | Out-Null }
    Set-ItemProperty -Path $wlPath -Name $ProgID -Value "" -Force
}

Write-Host "Registration Complete!" -ForegroundColor Green

# 6. Verify COM Object (Self-Test)
try {
            $testObj = New-Object -ComObject $ProgID -ErrorAction Stop
            Write-Host "SUCCESS: COM object created successfully!" -ForegroundColor Green
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($testObj)) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testObj) | Out-Null
            }
        } catch {
    Write-Warning "FAILED: Could not create COM object."
    Write-Warning "Error: $($_.Exception.Message)"
    Write-Warning "This might mean the DLL is not loadable (missing dependencies?) or registration failed."
}
