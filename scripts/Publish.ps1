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
$PublishDir = Join-Path $rootDir "publish"
$VersionJson = Join-Path $rootDir "version.json"
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

# 2. Build Setup (Find MSBuild)
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
