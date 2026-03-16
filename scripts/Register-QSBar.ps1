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

# --- Determine Registry Root ---
$IsAdmin = Test-IsAdmin
if ($IsAdmin) {
    $RegRoot = "HKLM:"
    Write-Host "Running as Administrator (Using HKLM - All Users)" -ForegroundColor Yellow
} else {
    $RegRoot = "HKCU:"
    Write-Host "Running as Standard User (Using HKCU - Current User)" -ForegroundColor Cyan
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

# --- Task 1: Clean Opposite Registry (Prevent Conflicts) ---
# If Admin (HKLM), try to clean HKCU CLSID to avoid shadowing.
# We DO NOT clean HKCU Addins anymore, because we will explicitly write to them.
if ($IsAdmin) {
    Write-Host "--- Cleaning HKCU CLSID (to enforce HKLM) ---" -ForegroundColor Cyan
    $hkcuPaths = @(
        "HKCU:\Software\Classes\CLSID\$CLSID",
        "HKCU:\Software\Classes\$ProgID"
    )
    foreach ($p in $hkcuPaths) {
        if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force; Write-Host "Removed HKCU shadow: $p" -ForegroundColor Gray }
    }
} elseif ($CleanHKLM) {
    # Existing logic for cleaning HKLM if requested (and failed IsAdmin check earlier?)
    # Actually Test-IsAdmin check is at top now.
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

# 0. Unblock DLL (Fix "Mark of the Web" issues)
try { Unblock-File -Path $DllPath -ErrorAction SilentlyContinue } catch {}

# 1. Run RegAsm (Updates HKCR/HKLM if Admin, HKCU if User)
$regasm32 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
$regasm64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (Test-Path $regasm32) {
    Write-Host "Running RegAsm (32-bit)..."
    $output = & $regasm32 /codebase "$DllPath" /tlb 2>&1
}
if (Test-Path $regasm64) {
    Write-Host "Running RegAsm (64-bit)..."
    $output = & $regasm64 /codebase "$DllPath" /tlb 2>&1
}

# 2. Force CLSID into Registry (Backup for RegAsm failure/quirks)
# We always write to HKCU CLSID as a fallback because WPS heavily relies on HKCU.
# Even if Admin (where RegAsm writes to HKLM), writing to HKCU ensures the current user can load it.
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

# Always write CLSID to HKCU to ensure current user can load it regardless of UAC
$targetRoots = @("HKCU:")
if ($IsAdmin) { $targetRoots += "HKLM:" }

foreach ($root in $targetRoots) {
    $clsidRoot = "$root\Software\Classes\CLSID\$CLSID"
    if (-not (Test-Path $clsidRoot)) { New-Item -Path $clsidRoot -Force | Out-Null }
    Set-ItemProperty -Path $clsidRoot -Name "(Default)" -Value $ProgID

    $inproc = New-Item -Path "$clsidRoot\InprocServer32" -Force
    Set-ItemProperty -Path $inproc.PSPath -Name "(Default)" -Value "C:\Windows\System32\mscoree.dll"
    Set-ItemProperty -Path $inproc.PSPath -Name "ThreadingModel" -Value "Both"
    Set-ItemProperty -Path $inproc.PSPath -Name "Class" -Value "QSBar.WpsExcelAddIn"
    Set-ItemProperty -Path $inproc.PSPath -Name "Assembly" -Value $fullAsmName
    Set-ItemProperty -Path $inproc.PSPath -Name "RuntimeVersion" -Value "v4.0.30319"
    Set-ItemProperty -Path $inproc.PSPath -Name "CodeBase" -Value "file:///$($DllPath.Replace('\', '/'))"

    $progIdKey = "$root\Software\Classes\$ProgID"
    if (-not (Test-Path $progIdKey)) { New-Item -Path $progIdKey -Force | Out-Null }
    $clsidKey = New-Item -Path "$progIdKey\CLSID" -Force
    Set-ItemProperty -Path $clsidKey.PSPath -Name "(Default)" -Value $CLSID
}

# 3. Register Addin paths (Excel & WPS) - Use $RegRoot
$comRegPaths = @(
    "$RegRoot\Software\Microsoft\Office\Excel\Addins\$ProgID",
    "$RegRoot\Software\Kingsoft\Office\ET\Addins\$ProgID",
    "$RegRoot\Software\Kingsoft\Office\ET\AddinsData\$ProgID",
    "$RegRoot\Software\Kingsoft\Office\WPS\Addins\$ProgID",
    "$RegRoot\Software\Kingsoft\Office\WPS\AddinsData\$ProgID"
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

# 4. WPS Whitelist (AddinsWL)
# Kingsoft might check HKLM whitelist too. Safe to add.
$wlPaths = @(
    "$RegRoot\Software\Kingsoft\Office\ET\AddinsWL",
    "$RegRoot\Software\Kingsoft\Office\WPS\AddinsWL",
    "$RegRoot\Software\Kingsoft\Office\6.0\Common\AddinsWL"
)
foreach ($path in $wlPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name $ProgID -Value "1" -Type String -Force
}

# 5. WOW6432Node Support (If Admin/HKLM, ensure 32-bit apps see it too)
if ($IsAdmin) {
    Write-Host "Adding WOW6432Node keys for 32-bit WPS/Office..." -ForegroundColor Cyan
    
    # 5.1 Addin Registration (WOW6432Node)
    $wowPaths = @(
        "HKLM:\Software\Wow6432Node\Kingsoft\Office\ET\Addins\$ProgID",
        "HKLM:\Software\Wow6432Node\Kingsoft\Office\WPS\Addins\$ProgID"
    )
    foreach ($path in $wowPaths) {
         if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
         Set-ItemProperty -Path $path -Name "Description" -Value $Description -Force
         Set-ItemProperty -Path $path -Name "FriendlyName" -Value $FriendlyName -Force
         Set-ItemProperty -Path $path -Name "LoadBehavior" -Value 3 -Type DWord -Force
         Set-ItemProperty -Path $path -Name "CommandLineSafe" -Value 1 -Type DWord -Force
         Remove-ItemProperty -Path $path -Name "Manifest" -ErrorAction SilentlyContinue
    }

    # 5.2 Whitelist (WOW6432Node) - Crucial for 32-bit WPS on 64-bit OS
    $wowWlPaths = @(
        "HKLM:\Software\Wow6432Node\Kingsoft\Office\ET\AddinsWL",
        "HKLM:\Software\Wow6432Node\Kingsoft\Office\WPS\AddinsWL",
        "HKLM:\Software\Wow6432Node\Kingsoft\Office\6.0\Common\AddinsWL"
    )
    foreach ($path in $wowWlPaths) {
        if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
        Set-ItemProperty -Path $path -Name $ProgID -Value "1" -Type String -Force
    }
}

# 6. Explicitly write to HKCU (Crucial for WPS)
# Office and WPS prioritize HKCU for Addins and Whitelist.
# Even if running as Admin, we write to the current HKCU profile.
Write-Host "Ensuring HKCU Addins and Whitelist keys..." -ForegroundColor Cyan
$hkcuAddinPaths = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgID",
    "HKCU:\Software\Kingsoft\Office\ET\Addins\$ProgID",
    "HKCU:\Software\Kingsoft\Office\ET\AddinsData\$ProgID",
    "HKCU:\Software\Kingsoft\Office\WPS\Addins\$ProgID",
    "HKCU:\Software\Kingsoft\Office\WPS\AddinsData\$ProgID"
)
foreach ($path in $hkcuAddinPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "Description" -Value $Description -Force
    Set-ItemProperty -Path $path -Name "FriendlyName" -Value $FriendlyName -Force
    Set-ItemProperty -Path $path -Name "LoadBehavior" -Value 3 -Type DWord -Force
    Set-ItemProperty -Path $path -Name "CommandLineSafe" -Value 1 -Type DWord -Force
    Remove-ItemProperty -Path $path -Name "Manifest" -ErrorAction SilentlyContinue
}

$hkcuWlPaths = @(
    "HKCU:\Software\Kingsoft\Office\ET\AddinsWL",
    "HKCU:\Software\Kingsoft\Office\WPS\AddinsWL",
    "HKCU:\Software\Kingsoft\Office\6.0\Common\AddinsWL"
)
foreach ($path in $hkcuWlPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name $ProgID -Value "1" -Type String -Force
}

Write-Host "Registration Complete!" -ForegroundColor Green

# 7. Verify COM Object (Self-Test)
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
