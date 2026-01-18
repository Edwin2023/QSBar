# QSBar One-Click Publish Script (Gitee Edition)
# Usage: .\Publish.ps1 -Version "1.0.0.1" -Log "Fixed version detection and optimized update process"

param (
    [string]$Version,
    [string]$Log
)

$RootDir = Get-Location
$DllSource = "$RootDir\QSBar\bin\Release\QSBar.dll"
$PublishDir = "$RootDir\publish"
$VersionJson = "$RootDir\version.json"

# 1. Check parameters
if (-not $Version -or -not $Log) {
    Write-Host "Error: Please provide Version and ChangeLog." -ForegroundColor Red
    Write-Host "Usage Example: .\Publish.ps1 -Version '1.0.0.1' -Log 'Updated something...'" -ForegroundColor Yellow
    exit
}

Write-Host "--- Starting Publish Process v$Version ---" -ForegroundColor Cyan

# 2. Check Build Result
if (-not (Test-Path $DllSource)) {
    Write-Host "Error: Cannot find compiled DLL at: $DllSource" -ForegroundColor Red
    Write-Host "Please make sure to build the project in 'Release' mode first." -ForegroundColor Yellow
    exit
}

# 3. Update publish directory
if (-not (Test-Path $PublishDir)) { New-Item -ItemType Directory -Path $PublishDir }
Copy-Item $DllSource -Destination "$PublishDir\QSBar.dll" -Force
Write-Host "Successfully copied latest DLL to publish directory." -ForegroundColor Green

# 4. Update version.json
$jsonObj = New-Object PSObject
$jsonObj | Add-Member NoteProperty "version" $Version
$jsonObj | Add-Member NoteProperty "downloadUrl" "https://gitee.com/kevin137/qsbar/raw/master/publish/QSBar.dll"
$jsonObj | Add-Member NoteProperty "changeLog" $Log

$jsonString = $jsonObj | ConvertTo-Json
[System.IO.File]::WriteAllText($VersionJson, $jsonString, [System.Text.Encoding]::UTF8)
Write-Host "Successfully updated version.json." -ForegroundColor Green

# 5. Push to Gitee
Write-Host "Pushing to Gitee..." -ForegroundColor Cyan
git add .
git commit -m "Release v$Version : $Log"
git push origin master

Write-Host "--- Publish Successful! ---" -ForegroundColor Green
Write-Host "Users will now receive the update notification for v$Version." -ForegroundColor Cyan
