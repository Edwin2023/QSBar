$regasm = "$env:windir\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
Set-Location $PSScriptRoot
& $regasm /codebase "QSBar.dll" /tlb
Pause
