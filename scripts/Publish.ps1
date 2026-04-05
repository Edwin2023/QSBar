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
$jsonObj | Add-Member NoteProperty "downloadUrl" "https://gitee.com/kevin137/qsbar-release/raw/master/QSBar_Setup_v$Version.exe"
$jsonObj | Add-Member NoteProperty "changeLog" $Log

$jsonString = $jsonObj | ConvertTo-Json
[System.IO.File]::WriteAllText($VersionJson, $jsonString, [System.Text.Encoding]::UTF8)
Write-Host "Successfully updated version.json." -ForegroundColor Green

# 7.5 Update Installinfo files
# Update Installinfo/version.json
[System.IO.File]::WriteAllText($InstallInfoVersionJson, $jsonString, [System.Text.Encoding]::UTF8)
Write-Host "Successfully updated Installinfo/version.json." -ForegroundColor Green

# Update UpdateNotes.html
$UpdateNotesPath = Join-Path $InstallInfoDir "UpdateNotes.html"
if (Test-Path $UpdateNotesPath) {
    $notesContent = [System.IO.File]::ReadAllText($UpdateNotesPath, [System.Text.Encoding]::UTF8)
    
    # Update Version
    $notesContent = $notesContent -replace '<!-- VERSION_START -->.*?<!-- VERSION_END -->', "<!-- VERSION_START -->$Version<!-- VERSION_END -->"
    
    # Update Log
    # Split the log by newline or semicolon if it's a single string, to create list items
    $logItems = $Log -split "[;\n]" | Where-Object { $_.Trim() -ne "" }
    $logHtml = ""
    foreach ($item in $logItems) {
        $logHtml += "            <li>$($item.Trim())</li>`n"
    }
    
    # Replace the block between LOG_START and LOG_END
    $notesContent = [regex]::Replace($notesContent, '(?s)(?<=<!-- LOG_START -->\s*).*?(?=\s*<!-- LOG_END -->)', "`n$logHtml")
    
    [System.IO.File]::WriteAllText($UpdateNotesPath, $notesContent, [System.Text.Encoding]::UTF8)
    Write-Host "Successfully updated UpdateNotes.html." -ForegroundColor Green
}

# Update QSBar_Installer.iss version
if (Test-Path $InstallerScript) {
    $issContent = [System.IO.File]::ReadAllText($InstallerScript, [System.Text.Encoding]::UTF8)
    # Update QSBar_Installer.iss version is no longer needed if using GetFileVersion, but we keep it for backward compatibility or we can remove it.
    # We will just rely on the DLL compilation to update the version.
    # $issContent = $issContent -replace '#define MyAppVersion ".*?"', "#define MyAppVersion `"$Version`""
    [System.IO.File]::WriteAllText($InstallerScript, $issContent, [System.Text.Encoding]::UTF8)
    Write-Host "Successfully updated QSBar_Installer.iss version." -ForegroundColor Green
}

# 7.8 Compile Inno Setup Installer
Write-Host "`n--- Compiling Inno Setup Installer ---" -ForegroundColor Cyan
$isccPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if (Test-Path $isccPath) {
    if (Test-Path $InstallerScript) {
        & $isccPath $InstallerScript
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Inno Setup compilation failed with exit code $LASTEXITCODE."
        } else {
            Write-Host "Successfully compiled installer to Release folder." -ForegroundColor Green
        }
    } else {
        Write-Warning "Installer script not found at $InstallerScript"
    }
} else {
    Write-Warning "Inno Setup compiler (ISCC.exe) not found at $isccPath. Skipping installer generation."
}


# 8. Push to Source Gitee (Private)
Write-Host "`n--- Pushing to Source Repo (Private) ---" -ForegroundColor Cyan
try {
    Push-Location $rootDir
    git add .
    git commit -m "Release v$Version : $Log"
    git push origin master
    Pop-Location
    Write-Host "Source push successful." -ForegroundColor Green
} catch {
    Write-Warning "Source git push failed. Please push manually."
}

# 9. Update Release Repo (Public) - The Release folder itself is the repo
Write-Host "`n--- Pushing to Release Repo (Public) at $PublishDir ---" -ForegroundColor Cyan

if (Test-Path $PublishDir) {
    try {
        # Initialize Git repo in Release folder if not already initialized
        Push-Location $PublishDir
        if (-not (Test-Path ".git")) {
            Write-Host "Initializing Git repository in Release folder..." -ForegroundColor Yellow
            git init
            git remote add origin https://gitee.com/kevin137/QSBar-release.git
        }
        Pop-Location

        # Copy necessary files from Source to Release Repo
        Copy-Item "$rootDir\README.md" -Destination "$PublishDir\README.md" -Force
        Copy-Item "$rootDir\version.json" -Destination "$PublishDir\version.json" -Force
        
        $relInstallinfoDir = Join-Path $PublishDir "Installinfo"
        if (-not (Test-Path $relInstallinfoDir)) { New-Item -ItemType Directory -Path $relInstallinfoDir | Out-Null }
        
        # Copy resources
        Copy-Item "$InstallInfoDir\*.bmp" -Destination "$relInstallinfoDir\" -Force
        Copy-Item "$InstallInfoDir\*.png" -Destination "$relInstallinfoDir\" -Force
        Copy-Item "$InstallInfoDir\*.ico" -Destination "$relInstallinfoDir\" -Force
        Copy-Item "$InstallInfoDir\version.json" -Destination "$relInstallinfoDir\" -Force
        Copy-Item "$InstallInfoDir\UpdateNotes.html" -Destination "$relInstallinfoDir\" -Force
        
        # Push Release Repo
        Push-Location $PublishDir
        git add .
        git commit -m "Release v$Version : $Log"
        # Push forcefully to ensure the remote matches the local Release folder
        git push origin master -f
        Pop-Location
        
        Write-Host "--- Publish to Public Release Repo Successful! ---" -ForegroundColor Green
        Write-Host "Users will now receive the update notification for v$Version from the public repo." -ForegroundColor Cyan
    } catch {
        Write-Warning "Release repo push failed. Please push manually."
        Write-Warning $_
    }
} else {
    Write-Warning "Release directory not found at $PublishDir. Skipping public push."
}
