$ErrorActionPreference = "Stop"

$baseDir = "$PSScriptRoot\.."
$assemblyInfo = "$baseDir\QSBar\Properties\AssemblyInfo.cs"
$csproj = "$baseDir\QSBar\QSBar.csproj"
$sln = "$baseDir\QSBar.sln"
$iss = "$baseDir\Installinfo\QSBar_Installer.iss"
$msbuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$iscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "      QSBar 1.0.1.0 Test Package Generator" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

Write-Host "`n[Step 1] Backing up original source files..." -ForegroundColor Yellow
Copy-Item $assemblyInfo "$assemblyInfo.bak" -Force
Copy-Item $csproj "$csproj.bak" -Force

try {
    Write-Host "[Step 2] Temporarily modifying DLL version to 1.0.1.0..." -ForegroundColor Yellow
    $asmContent = Get-Content $assemblyInfo -Raw -Encoding UTF8
    $asmContent = $asmContent -replace '(?m)^\[assembly:\s*AssemblyVersion\(".*?"\)\]', '[assembly: AssemblyVersion("1.0.1.0")]'
    $asmContent = $asmContent -replace '(?m)^\[assembly:\s*AssemblyFileVersion\(".*?"\)\]', '[assembly: AssemblyFileVersion("1.0.1.0")]'
    Set-Content -Path $assemblyInfo -Value $asmContent -Encoding UTF8

    $projContent = Get-Content $csproj -Raw -Encoding UTF8
    $projContent = $projContent -replace '<ApplicationVersion>.*?</ApplicationVersion>', '<ApplicationVersion>1.0.1.0</ApplicationVersion>'
    Set-Content -Path $csproj -Value $projContent -Encoding UTF8

    Write-Host "[Step 3] Compiling QSBar DLL (v1.0.1.0)..." -ForegroundColor Yellow
    $buildResult = & $msbuild $sln /p:Configuration=Release /p:Platform="Any CPU" /t:Rebuild /p:CscToolPath="C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\Roslyn"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Build failed! Output:" -ForegroundColor Red
        $buildResult | Out-Host
        throw "MSBuild failed to compile the DLL."
    }

    Write-Host "[Step 4] Packaging Installer using Inno Setup..." -ForegroundColor Yellow
    Copy-Item "$baseDir\QSBar\bin\Release\QSBar.dll" "$baseDir\Release\QSBar.dll" -Force
    & $iscc $iss | Out-Null

    Write-Host "`n[SUCCESS] Successfully created Release\QSBar_Setup_v1.0.1.0.exe!" -ForegroundColor Green
}
finally {
    Write-Host "`n[Step 5] Restoring original source files (v1.0.1.1+)..." -ForegroundColor Yellow
    Move-Item "$assemblyInfo.bak" $assemblyInfo -Force
    Move-Item "$csproj.bak" $csproj -Force

    Write-Host "[Step 6] Recompiling original DLL to restore environment..." -ForegroundColor Yellow
    & $msbuild $sln /p:Configuration=Release /p:Platform="Any CPU" /t:Rebuild /p:CscToolPath="C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\Roslyn" | Out-Null
    if (Test-Path "$baseDir\QSBar\bin\Release\QSBar.dll") {
        Copy-Item "$baseDir\QSBar\bin\Release\QSBar.dll" "$baseDir\Release\QSBar.dll" -Force
    }
    
    Write-Host "`n[DONE] Environment fully restored. You can continue developing." -ForegroundColor Cyan
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}