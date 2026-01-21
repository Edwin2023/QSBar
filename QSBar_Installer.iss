; QSBar One-Click Installer Script for Inno Setup
#define MyAppName "QSBar"
#define MyAppVersion "1.0.0.1"
#define MyAppPublisher "Bookmen"
#define MyAppExeName "QSBar.dll"
#define SourcePath "QSBar\bin\Release"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://gitee.com/kevin137/qsbar
AppSupportURL=https://gitee.com/kevin137/qsbar
AppUpdatesURL=https://gitee.com/kevin137/qsbar
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=Installer
OutputBaseFilename=QSBar_Setup_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
; Run in 64-bit mode on 64-bit systems to access both registry views easily
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern
; 检测 Excel 和 WPS 是否正在运行，防止文件占用
AppMutex=Excel,wps.exe,et.exe,wpp.exe
; 展示个性化说明
InfoBeforeFile=InstallInfo.rtf
; Simplify installation process
DisableProgramGroupPage=yes
DisableReadyPage=no
; SetupIconFile=QSBar.ico

; Digital Signing Configuration
; You need to configure a SignTool named 'Standard' in Inno Setup IDE:
; Tools -> Configure Sign Tools... -> Add
; Name: Standard
; Command: "signtool.exe" sign /f "C:\QSBar_Cert.pfx" /p 123456 /fd SHA256 /t http://timestamp.digicert.com $f
SignTool=Standard
SignedUninstaller=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
english.WelcomeLabel1=Welcome to QSBar Toolbox
english.WelcomeLabel2=This toolbox is designed to significantly enhance your productivity in Excel and WPS.%n%nBefore continuing, please ensure all Office applications are closed.
english.ReadyLabel1=Setup is ready to inject the "Productivity Soul" into your computer.
english.ReadyLabel2=Click "Install" to begin the deployment of QSBar.
; Custom labels for Info page
english.InfoBeforeLabel=Important Information
english.InfoBeforeClickLabel=Please read the following information about QSBar before continuing.
english.ClickNext=When you are ready to continue with Setup, click Next.

[InstallDelete]
Type: files; Name: "{app}\QSBar.dll"
Type: files; Name: "{app}\QSBar.tlb"
Type: files; Name: "{app}\QSBar.dll.config"

[Files]
Source: "{#SourcePath}\QSBar.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\QSBar.tlb"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.Interfaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.System.Drawing.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\System.ComponentModel.Annotations.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\QSBar.dll.config"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; --- 32-bit COM Registration (For 32-bit Office/WPS) ---
; Register CLSID in 32-bit view
Root: HKCU32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"; ValueType: string; ValueName: ""; ValueData: "QSBar.WpsAddIn"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "QSBar.WpsExcelAddIn"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "QSBar, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}/QSBar.dll"; Flags: uninsdeletekey

Root: HKLM32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"; ValueType: string; ValueName: ""; ValueData: "QSBar.WpsAddIn"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "QSBar.WpsExcelAddIn"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "QSBar, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}/QSBar.dll"; Flags: uninsdeletekey

; Register ProgID (32-bit)
Root: HKCU32; Subkey: "Software\Classes\QSBar.WpsAddIn"; ValueType: string; ValueName: ""; ValueData: "QSBar.WpsAddIn"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\QSBar.WpsAddIn\CLSID"; ValueType: string; ValueName: ""; ValueData: "{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"; Flags: uninsdeletekey

; Register Add-in for Excel & WPS (32-bit)
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar Excel Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

Root: HKCU32; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

Root: HKCU32; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; --- 64-bit COM Registration (For 64-bit Office/WPS) ---
; Register CLSID in 64-bit view
Root: HKCU64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"; ValueType: string; ValueName: ""; ValueData: "QSBar.WpsAddIn"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "QSBar.WpsExcelAddIn"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "QSBar, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}/QSBar.dll"; Flags: uninsdeletekey; Check: IsWin64

Root: HKLM64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"; ValueType: string; ValueName: ""; ValueData: "QSBar.WpsAddIn"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "QSBar.WpsExcelAddIn"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "QSBar, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}/QSBar.dll"; Flags: uninsdeletekey; Check: IsWin64

; Register Add-in for Excel & WPS (64-bit)
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar Excel Productivity Add-in"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64

Root: HKCU64; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64

Root: HKCU64; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64

; --- WPS Whitelist (Both views) ---
Root: HKCU32; Subkey: "Software\Kingsoft\Office\ET\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU32; Subkey: "Software\Kingsoft\Office\WPS\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU32; Subkey: "Software\Kingsoft\Office\Common\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU32; Subkey: "Software\Kingsoft\Office\6.0\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue

Root: HKCU64; Subkey: "Software\Kingsoft\Office\ET\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\WPS\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\Common\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue; Check: IsWin64
Root: HKCU64; Subkey: "Software\Kingsoft\Office\6.0\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue; Check: IsWin64

[Run]
; 安装前清理：尝试注销旧组件（即使文件不存在也会静默执行）
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; Flags: runhidden; StatusMsg: "Cleaning up old 32-bit registration..."; BeforeInstall: TaskKillOffice
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; Flags: runhidden; StatusMsg: "Cleaning up old 64-bit registration..."; Check: Is64BitInstallMode

; Register COM using RegAsm
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/codebase ""{app}\QSBar.dll"" /tlb:""{app}\QSBar.tlb"""; StatusMsg: "Registering components..."; Flags: runhidden

[Code]
procedure TaskKillOffice();
var
  ResultCode: Integer;
begin
  Exec('taskkill.exe', '/f /im excel.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im wps.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im et.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
