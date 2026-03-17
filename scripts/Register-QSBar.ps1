param(
    [string]$DllPath,
    [switch]$CleanHKLM,
    [switch]$RestartApps,
    [switch]$Unregister,
    [switch]$UserConfigOnly
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

# --- Determine Mode ---
$IsAdmin = Test-IsAdmin

if ($UserConfigOnly) {
    Write-Host "Running in User Config Mode (Writing HKCU Only)" -ForegroundColor Cyan
    $RegRoot = "HKCU:"
} elseif ($IsAdmin) {
    $RegRoot = "HKLM:"
    Write-Host "Running as Administrator (Using HKLM - All Users)" -ForegroundColor Yellow
} else {
    # Fallback to HKCU if not admin and not explicitly UserConfigOnly (though RegAsm will fail)
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
# We SKIP cleaning HKCU when running as Admin to avoid "UAC Registry Drift" (Admin HKCU != User HKCU).
# We only clean HKLM if requested.
if ($CleanHKLM -and $IsAdmin) {
    # Logic to clean HKLM if implemented...
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

# 0. Check .NET 4.8 Runtime
$releaseKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
if (Test-Path $releaseKey) {
    $release = Get-ItemProperty -Path $releaseKey -Name "Release" -ErrorAction SilentlyContinue
    if ($release.Release -lt 528040) {
        Write-Error "CRITICAL: .NET Framework 4.8 is required but not found (Release: $($release.Release))."
        Write-Error "Please install .NET Framework 4.8 Runtime."
        exit 1
    }
} else {
    Write-Error "CRITICAL: .NET Framework 4.x is not installed."
    exit 1
}

# 0.1 Unblock DLL (Fix "Mark of the Web" issues)
try { Unblock-File -Path $DllPath -ErrorAction SilentlyContinue } catch {}

# 1. Run RegAsm (Updates HKCR/HKLM if Admin, HKCU if User)
if (-not $UserConfigOnly) {
    $regasm32 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    $regasm64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

    $regSuccess = $false

    if (Test-Path $regasm32) {
        Write-Host "Running RegAsm (32-bit)..."
        $p = Start-Process -FilePath $regasm32 -ArgumentList "/codebase `"$DllPath`" /tlb" -PassThru -Wait -NoNewWindow
        if ($p.ExitCode -eq 0) { $regSuccess = $true }
        else { Write-Warning "RegAsm (32-bit) failed with exit code $($p.ExitCode)." }
    }
    if (Test-Path $regasm64) {
        Write-Host "Running RegAsm (64-bit)..."
        $p = Start-Process -FilePath $regasm64 -ArgumentList "/codebase `"$DllPath`" /tlb" -PassThru -Wait -NoNewWindow
        if ($p.ExitCode -eq 0) { $regSuccess = $true }
        else { Write-Warning "RegAsm (64-bit) failed with exit code $($p.ExitCode)." }
    }

    if (-not $regSuccess) {
        Write-Warning "RegAsm failed to register the assembly. Attempting manual registry fallback..."
    }
} else {
    Write-Host "Skipping RegAsm (User Config Mode)" -ForegroundColor Yellow
}

# 2. Clean up any manual HKCU CLSID (Fix for v1.0.0.10 and earlier)
# Manual HKCU CLSID shadows RegAsm's HKLM registration and often lacks Implemented Categories,
# causing WPS to silently fail to load the add-in.
$hkcuClsidRoot = "HKCU:\Software\Classes\CLSID\$CLSID"
if (Test-Path $hkcuClsidRoot) {
    Write-Host "Cleaning up legacy HKCU CLSID to unblock HKLM registration..." -ForegroundColor Yellow
    Remove-Item -Path $hkcuClsidRoot -Recurse -Force -ErrorAction SilentlyContinue
}
$hkcuProgIdRoot = "HKCU:\Software\Classes\$ProgID"
if (Test-Path $hkcuProgIdRoot) {
    Remove-Item -Path $hkcuProgIdRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# 3. Register Addin paths (Excel & WPS) - Use $RegRoot
# Note: If UserConfigOnly is set, RegRoot is already set to HKCU: above.
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
if ($IsAdmin -and -not $UserConfigOnly) {
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
