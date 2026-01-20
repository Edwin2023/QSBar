; QSBar One-Click Installer Script for Inno Setup
#define MyAppName "QS工具箱"
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
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern
; 检测 Excel 和 WPS 是否正在运行，防止文件占用
AppMutex=Excel,wps.exe,et.exe,wpp.exe
; 展示个性化说明
InfoBeforeFile=InstallInfo.rtf
; 简化安装过程
DisableProgramGroupPage=yes
DisableReadyPage=no
; 设置图标（如果有的话）
; SetupIconFile=QSBar.ico

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
Source: "{#SourcePath}\EPPlus.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.Interfaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.System.Drawing.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\System.ComponentModel.Annotations.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\QSBar.dll.config"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Register for Excel
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar Excel Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName} (COM)"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; Register for WPS (Standard Path)
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName} (WPS)"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; Register for WPS (Excel Compatible Path)
Root: HKCU; Subkey: "Software\Kingsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName} (WPS)"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; WPS Whitelist
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\Common\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue

[Run]
; 安装前清理：尝试注销旧组件（即使文件不存在也会静默执行）
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; Flags: runhidden; StatusMsg: "Cleaning up old 32-bit registration..."; BeforeInstall: TaskKillOffice
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; Flags: runhidden; StatusMsg: "Cleaning up old 64-bit registration..."; Check: Is64BitInstallMode

; Register COM using RegAsm
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/codebase ""{app}\QSBar.dll"" /tlb"; StatusMsg: "Registering components..."; Flags: runhidden
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/codebase ""{app}\QSBar.dll"" /tlb"; StatusMsg: "Registering 64-bit components..."; Flags: runhidden; Check: Is64BitInstallMode

[UninstallRun]
; Unregister COM
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; StatusMsg: "正在注销组件..."; Flags: runhidden
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; StatusMsg: "正在注销 64 位组件..."; Flags: runhidden; Check: Is64BitInstallMode

[Code]
procedure TaskKillOffice();
var
  ResultCode: Integer;
begin
  // 使用 taskkill 强制关闭可能占用的进程，确保文件不被锁定
  Exec('taskkill.exe', '/f /im wps.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im et.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im wpp.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im excel.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

