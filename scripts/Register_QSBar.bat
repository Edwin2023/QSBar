@echo off
setlocal EnableDelayedExpansion

set "REGASM32=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
set "REGASM64=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
set "LOG_FILE=%TEMP%\QSBar_Register.log"

echo Starting Registration at %DATE% %TIME% > "%LOG_FILE%"
cd /d "%~dp0"

echo ========================================
echo QSBar COM Registration Tool (Diagnostic Mode)
echo ========================================

:: 1. 32-bit Registration
echo [1/3] Registering for 32-bit (WPS / Office 32)...
if exist "%REGASM32%" (
    echo Running 32-bit RegAsm...
    "%REGASM32%" /codebase "QSBar.dll" /tlb "QSBar32.tlb" >> "%LOG_FILE%" 2>&1
    if !errorlevel! neq 0 (
        echo ERROR: 32-bit registration failed. See %LOG_FILE%
    ) else (
        echo Success: 32-bit registration completed.
    )
) else (
    echo SKIP: 32-bit .NET Framework not found.
)

:: 2. 64-bit Registration
echo.
echo [2/3] Registering for 64-bit (Office 64)...
if exist "%REGASM64%" (
    echo Running 64-bit RegAsm...
    "%REGASM64%" /codebase "QSBar.dll" /tlb "QSBar64.tlb" >> "%LOG_FILE%" 2>&1
    if !errorlevel! neq 0 (
        echo ERROR: 64-bit registration failed. See %LOG_FILE%
    ) else (
        echo Success: 64-bit registration completed.
    )
) else (
    echo SKIP: 64-bit .NET Framework not found.
)

:: 3. ProgID and CLSID Mapping (Force both HKCU and HKLM if possible)
echo.
echo [3/3] Finalizing Registry Mappings...
set "GUID={D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"
set "PROGID=QSBar.WpsAddIn"

:: Register ProgID
reg add "HKCU\Software\Classes\!PROGID!" /ve /t REG_SZ /d "!PROGID!" /f >> "%LOG_FILE%" 2>&1
reg add "HKCU\Software\Classes\!PROGID!\CLSID" /ve /t REG_SZ /d "!GUID!" /f >> "%LOG_FILE%" 2>&1

:: Force Excel to see the Add-in (sometimes HKCU is not enough for 64-bit Office)
echo [Extra] Ensuring Excel Add-in keys exist...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\!PROGID!" /v "LoadBehavior" /t REG_DWORD /d 3 /f >> "%LOG_FILE%" 2>&1
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\!PROGID!" /v "FriendlyName" /t REG_SZ /d "QSBar" /f >> "%LOG_FILE%" 2>&1
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\!PROGID!" /v "Description" /t REG_SZ /d "QSBar Productivity Tool" /f >> "%LOG_FILE%" 2>&1

:: Check if registered
echo.
echo Verification:
reg query "HKCR\CLSID\!GUID!" /ve >nul 2>&1
if !errorlevel! equ 0 (
    echo CLSID found in registry (HKCR).
) else (
    echo WARNING: CLSID NOT found in HKCR. Trying HKCU...
    reg query "HKCU\Software\Classes\CLSID\!GUID!" /ve >nul 2>&1
    if !errorlevel! equ 0 (
        echo CLSID found in HKCU.
    ) else (
        echo CRITICAL ERROR: CLSID not found anywhere!
    )
)

echo.
echo Registration finished. 
echo Log saved to: %LOG_FILE%
echo Please restart Excel/WPS.
timeout /t 5
