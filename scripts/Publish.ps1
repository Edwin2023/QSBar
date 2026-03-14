# QSBar One-Click Publish Script (Gitee Edition)
# Usage: .\Publish.ps1 -Version "1.0.0.1" -Log "Fixed version detection and optimized update process"

param (
    [string]$Version,
    [string]$Log
)

$ErrorActionPreference = "Stop" # Stop on error
$baseDir = $PSScriptRoot
$rootDir = (Get-Item $baseDir).Parent.FullName
$DllSource = Join-Path $rootDir "QSBar\bin\Release\QSBar.dll"
$PublishDir = Join-Path $rootDir "Release"
$VersionJson = Join-Path $rootDir "version.json"
$InstallInfoDir = Join-Path $rootDir "Installinfo"
$InstallInfoVersionJson = Join-Path $InstallInfoDir "version.json"
$InstallerScript = Join-Path $InstallInfoDir "QSBar_Installer.iss"
$projectPath = Join-Path $rootDir "QSBar\QSBar.csproj"
$slnPath = Join-Path $rootDir "QSBar.sln"
$nugetExe = Join-Path $rootDir "nuget.exe"

# 1. Check parameters
if (-not $Version -or -not $Log) {
    Write-Host "Error: Please provide Version and ChangeLog." -ForegroundColor Red
    Write-Host "Usage Example: .\Publish.ps1 -Version '1.0.0.1' -Log 'Updated something...'" -ForegroundColor Yellow
    exit
}

Write-Host "--- Starting Publish Process v$Version ---" -ForegroundColor Cyan

# 2.5 Update AssemblyInfo.cs and .csproj with new version
Write-Host "--- Updating Version to $Version ---" -ForegroundColor Cyan
$AssemblyInfoPath = Join-Path $rootDir "QSBar\Properties\AssemblyInfo.cs"
if (Test-Path $AssemblyInfoPath) {
    # Use System.IO.File to handle encoding reliably (UTF-8)
    $content = [System.IO.File]::ReadAllText($AssemblyInfoPath, [System.Text.Encoding]::UTF8)
    $content = $content -replace 'AssemblyVersion\(".*?"\)', "AssemblyVersion(`"$Version`")"
    $content = $content -replace 'AssemblyFileVersion\(".*?"\)', "AssemblyFileVersion(`"$Version`")"
    [System.IO.File]::WriteAllText($AssemblyInfoPath, $content, [System.Text.Encoding]::UTF8)
    Write-Host "Updated AssemblyInfo.cs" -ForegroundColor Gray
}

if (Test-Path $projectPath) {
    $content = [System.IO.File]::ReadAllText($projectPath, [System.Text.Encoding]::UTF8)
    $content = $content -replace '<ApplicationVersion>.*?</ApplicationVersion>', "<ApplicationVersion>$Version</ApplicationVersion>"
    [System.IO.File]::WriteAllText($projectPath, $content, [System.Text.Encoding]::UTF8)
    Write-Host "Updated QSBar.csproj ApplicationVersion" -ForegroundColor Gray
}

# 3. Build Setup (Find MSBuild)
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
    Write-Error "Could not find MSBuild.exe. Please build the project manually in Release mode."
    exit
}

Write-Host "Using MSBuild: $msbuildPath" -ForegroundColor Gray

# 3. Restore NuGet Packages
Write-Host "`n--- Restoring NuGet Packages ---" -ForegroundColor Cyan
if (Test-Path $nugetExe) {
    try {
        & $nugetExe restore $slnPath -Source "https://api.nuget.org/v3/index.json"
    } catch {
        Write-Warning "NuGet restore failed or had warnings. Continuing..."
    }
} else {
    Write-Warning "nuget.exe not found. Skipping restore."
}

# 4. Build Project (Release)
Write-Host "`n--- Building Project (Release) ---" -ForegroundColor Cyan
& $msbuildPath $projectPath /p:Configuration=Release /p:Platform="AnyCPU" /t:Rebuild
if ($LASTEXITCODE -ne 0) {
    Write-Error "Build Failed!"
    exit
}

# 5. Check Build Result
if (-not (Test-Path $DllSource)) {
    Write-Host "Error: Cannot find compiled DLL at: $DllSource" -ForegroundColor Red
    exit
}

# 6. Update publish directory
if (-not (Test-Path $PublishDir)) { New-Item -ItemType Directory -Path $PublishDir | Out-Null }
Copy-Item $DllSource -Destination "$PublishDir\QSBar.dll" -Force
Write-Host "Successfully copied latest DLL to publish directory." -ForegroundColor Green

# 7. Update version.json
$jsonObj = New-Object PSObject
$jsonObj | Add-Member NoteProperty "version" $Version
$jsonObj | Add-Member NoteProperty "downloadUrl" "https://gitee.com/kevin137/qsbar/raw/master/publish/QSBar.dll"
$jsonObj | Add-Member NoteProperty "changeLog" $Log

$jsonString = $jsonObj | ConvertTo-Json
[System.IO.File]::WriteAllText($VersionJson, $jsonString, [System.Text.Encoding]::UTF8)
Write-Host "Successfully updated version.json." -ForegroundColor Green

# 7.5 Update Installinfo files
# Update Installinfo/version.json
[System.IO.File]::WriteAllText($InstallInfoVersionJson, $jsonString, [System.Text.Encoding]::UTF8)
Write-Host "Successfully updated Installinfo/version.json." -ForegroundColor Green

# Update QSBar_Installer.iss version
if (Test-Path $InstallerScript) {
    $issContent = [System.IO.File]::ReadAllText($InstallerScript, [System.Text.Encoding]::UTF8)
    # Replace #define MyAppVersion "x.x.x.x"
    $issContent = $issContent -replace '#define MyAppVersion ".*?"', "#define MyAppVersion `"$Version`""
    [System.IO.File]::WriteAllText($InstallerScript, $issContent, [System.Text.Encoding]::UTF8)
    Write-Host "Successfully updated QSBar_Installer.iss version." -ForegroundColor Green
}


# 8. Push to Gitee
Write-Host "Pushing to Gitee..." -ForegroundColor Cyan
try {
    Push-Location $rootDir
    git add .
    git commit -m "Release v$Version : $Log"
    git push origin master
    Pop-Location
    Write-Host "--- Publish Successful! ---" -ForegroundColor Green
    Write-Host "Users will now receive the update notification for v$Version." -ForegroundColor Cyan
} catch {
    Write-Warning "Git push failed. Please push manually."
    Write-Warning $_
}
