; QSBar One-Click Installer Script for Inno Setup
#define MyAppName "QSBar"
#define MyAppPublisher "Bookmen"
#define MyAppExeName "QSBar.dll"
#define SourcePath "..\QSBar\bin\Release"
#define ScriptPath "..\scripts"
#define MyAppVersion GetVersionNumbersString("..\QSBar\bin\Release\QSBar.dll")

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
OutputDir=..\Release
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
SetupIconFile=qsbar.ico
UninstallDisplayIcon={app}\qsbar.ico

; Digital Signing Configuration
; You need to configure a SignTool named 'Standard' in Inno Setup IDE:
; Tools -> Configure Sign Tools... -> Add
; Name: Standard
; Command: "signtool.exe" sign /f "C:\QSBar_Cert.pfx" /p 123456 /fd SHA256 /t http://timestamp.digicert.com $f
; SignTool=Standard
; SignedUninstaller=yes

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
Source: "{#SourcePath}\QSBar.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\..\..\..\lib\Office.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\..\..\..\lib\Microsoft.Office.Interop.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\..\..\..\lib\Interop.AddInDesignerObjects.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\QSBar.tlb"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.Interfaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.System.Drawing.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\System.ComponentModel.Annotations.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\QSBar.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#ScriptPath}\Register.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#ScriptPath}\Register-QSBar.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#ScriptPath}\DeepClean-QSBar.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "Ribbon_EXCEL.bmp"; DestDir: "{app}"; Flags: ignoreversion
Source: "Ribbon_WPS.bmp"; DestDir: "{app}"; Flags: ignoreversion
Source: "qsbar.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo_horizontal.png"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; --- Office/WPS 加载项注册 (部分由 Register-QSBar.ps1 处理) ---
; 注意：Installer 运行在 Admin 模式，HKCU 指向 Admin 用户的 HKCU，而不是登录用户的 HKCU。
; 因此，必须依靠 [Run] 部分的 runasoriginaluser 来注册当前用户的 HKCU。
; 下面的 Registry 部分仅作为 Admin 用户的备份，或者 HKLM 注册。

; 1. HKLM 注册 (所有用户) - 仅在 PrivilegesRequired=admin 时有效
; 仅删除，不写入，避免干扰 RegAsm
Root: HKLM; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"; Flags: uninsdeletekey dontcreatekey
Root: HKLM; Subkey: "Software\Classes\QSBar.WpsAddIn"; Flags: uninsdeletekey dontcreatekey

; 2. Admin 用户的 HKCU (可选，主要用于调试)
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: "1"; Flags: uninsdeletevalue

Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: "1"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\6.0\Common\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: "1"; Flags: uninsdeletevalue

; 强制在 Admin HKCU 也写入 Addins 注册表
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "QSBar (COM)"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar COM Add-in for Excel and WPS"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletevalue

Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "QSBar (COM)"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar COM Add-in for Excel and WPS"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletevalue

Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "QSBar (COM)"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar COM Add-in for Excel and WPS"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletevalue

[Run]
; 1. Admin Registration (System-wide HKLM) - Executed by Installer (Admin)
Filename: "{app}\Register.bat"; Parameters: "-Silent"; WorkingDir: "{app}"; Flags: waituntilterminated; StatusMsg: "Registering QSBar (System-Level)..."

; 2. User Registration (Current User HKCU) - Executed as Original User
; This is CRITICAL for WPS/Excel to see the add-in in the user's profile
Filename: "{app}\Register.bat"; Parameters: "-Silent"; WorkingDir: "{app}"; Flags: runasoriginaluser waituntilterminated; StatusMsg: "Registering QSBar (User-Level)..."

; 3. Auto-launch Excel after install (instead of showing notes)
Filename: "excel.exe"; Description: "Launch Excel now (立即启动 Excel)"; Flags: postinstall nowait shellexec skipifsilent


[UninstallRun]
; 1. Unregister for User (HKCU) - Executed as Original User
; Note: runasoriginaluser is NOT supported in [UninstallRun]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\Register-QSBar.ps1"" -Unregister"; WorkingDir: "{app}"; Flags: runhidden

; 2. Unregister for System (HKLM)
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\Register-QSBar.ps1"" -Unregister"; WorkingDir: "{app}"; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{app}"


[Code]
var
  ExcelImage, WpsImage: TBitmapImage;
  NeedPreInstallCleanup: Boolean;
  PreInstallCleanupDone: Boolean;

procedure TaskKillOffice();
var
  ResultCode: Integer;
begin
  Exec('taskkill.exe', '/f /im excel.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im wps.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im et.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(3000); // 确保进程已完全退出并释放文件锁
end;

// 获取已安装版本的卸载字符串
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

// 检查是否已安装
function IsInstalled(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

// 解析版本号字符串为数字，用于比较
procedure DecodeVersion(verstr: String; var v1, v2, v3, v4: Integer);
var
  i, p: Integer;
  s: String;
begin
  v1 := 0; v2 := 0; v3 := 0; v4 := 0;
  s := verstr;
  for i := 1 to 4 do
  begin
    p := Pos('.', s);
    if p > 0 then
    begin
      if i = 1 then v1 := StrToIntDef(Copy(s, 1, p - 1), 0)
      else if i = 2 then v2 := StrToIntDef(Copy(s, 1, p - 1), 0)
      else if i = 3 then v3 := StrToIntDef(Copy(s, 1, p - 1), 0);
      s := Copy(s, p + 1, Length(s));
    end
    else
    begin
      if i = 1 then v1 := StrToIntDef(s, 0)
      else if i = 2 then v2 := StrToIntDef(s, 0)
      else if i = 3 then v3 := StrToIntDef(s, 0)
      else if i = 4 then v4 := StrToIntDef(s, 0);
      break;
    end;
  end;
end;

// 比较版本号：如果 Ver1 > Ver2 返回 1，Ver1 < Ver2 返回 -1，相等返回 0
function CompareVersion(ver1, ver2: String): Integer;
var
  v1_1, v1_2, v1_3, v1_4: Integer;
  v2_1, v2_2, v2_3, v2_4: Integer;
begin
  DecodeVersion(ver1, v1_1, v1_2, v1_3, v1_4);
  DecodeVersion(ver2, v2_1, v2_2, v2_3, v2_4);
  
  if v1_1 > v2_1 then Result := 1 else if v1_1 < v2_1 then Result := -1
  else if v1_2 > v2_2 then Result := 1 else if v1_2 < v2_2 then Result := -1
  else if v1_3 > v2_3 then Result := 1 else if v1_3 < v2_3 then Result := -1
  else if v1_4 > v2_4 then Result := 1 else if v1_4 < v2_4 then Result := -1
  else Result := 0;
end;

// 获取已安装程序的版本号
function GetInstalledVersion(): String;
var
  sUnInstPath: String;
  sVersion: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}_is1');
  sVersion := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'DisplayVersion', sVersion) then
    RegQueryStringValue(HKCU, sUnInstPath, 'DisplayVersion', sVersion);
  Result := sVersion;
end;

procedure ForceCleanupFallback();
var
  Rc: Integer;
begin
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\Wow6432Node\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\Wow6432Node\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn');
  RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\ET\AddinsWL', 'QSBar.WpsAddIn');
  RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\WPS\AddinsWL', 'QSBar.WpsAddIn');
  RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\Common\AddinsWL', 'QSBar.WpsAddIn');
  RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\6.0\AddinsWL', 'QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}');
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\Classes\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\WOW6432Node\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}');
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\WOW6432Node\Classes\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\WOW6432Node\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn');
  RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}_is1');
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}_is1');
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}_is1');
  Exec('cmd.exe', '/c rmdir /s /q "' + ExpandConstant('{localappdata}\QSBar') + '"', '', SW_HIDE, ewWaitUntilTerminated, Rc);
end;

procedure ForceCleanupUserFallback();
var
  Rc: Integer;
begin
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Classes\QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Classes\Wow6432Node\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Classes\Wow6432Node\QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Kingsoft\Office\ET\AddinsWL" /v "QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Kingsoft\Office\WPS\AddinsWL" /v "QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Kingsoft\Office\Common\AddinsWL" /v "QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Kingsoft\Office\6.0\AddinsWL" /v "QSBar.WpsAddIn" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('reg.exe', 'delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}_is1" /f', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  ExecAsOriginalUser('cmd.exe', '/c rmdir /s /q "' + ExpandConstant('{localappdata}\QSBar') + '"', '', SW_HIDE, ewWaitUntilTerminated, Rc);
end;

procedure RunPreInstallCleanup();
var
  iResultCode: Integer;
  OkAdmin: Boolean;
  OkUser: Boolean;
  CmdLine: String;
begin
  if PreInstallCleanupDone then Exit;
  if WizardForm <> nil then
  begin
    WizardForm.StatusLabel.Caption := 'Preparing update environment...';
    WizardForm.Update();
  end;
  TaskKillOffice();
  if WizardForm <> nil then
  begin
    WizardForm.StatusLabel.Caption := 'Running deep cleanup...';
    WizardForm.Update();
  end;
  ExtractTemporaryFile('DeepClean-QSBar.ps1');
  CmdLine := '-ExecutionPolicy Bypass -WindowStyle Hidden -File "' + ExpandConstant('{tmp}\DeepClean-QSBar.ps1') + '" -Interactive:$false';
  OkAdmin := Exec('powershell.exe', CmdLine, '', SW_HIDE, ewWaitUntilTerminated, iResultCode) and (iResultCode = 0);
  OkUser := ExecAsOriginalUser('powershell.exe', CmdLine, '', SW_HIDE, ewWaitUntilTerminated, iResultCode) and (iResultCode = 0);
  if not (OkAdmin or OkUser) then
    ForceCleanupFallback();
  ForceCleanupUserFallback();
  if WizardForm <> nil then
  begin
    WizardForm.StatusLabel.Caption := 'Deep cleanup completed. Installing...';
    WizardForm.Update();
  end;
  PreInstallCleanupDone := True;
end;

function InitializeSetup(): Boolean;
var
  V: Integer;
  InstalledVer: String;
  CurrentVer: String;
  Release: Cardinal;
begin
  Result := True;
  NeedPreInstallCleanup := True;
  PreInstallCleanupDone := False;
  
  // Check for .NET Framework 4.8
  if RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full') then
  begin
    if RegQueryDwordValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) then
    begin
      if Release < 528040 then
      begin
        MsgBox('.NET Framework 4.8 is required.' + #13#10 + 
               'Please install it and run setup again.', mbCriticalError, MB_OK);
        Result := False;
        Exit;
      end;
    end;
  end
  else
  begin
    MsgBox('.NET Framework 4.8 is required.' + #13#10 + 
           'Please install it and run setup again.', mbCriticalError, MB_OK);
    Result := False;
    Exit;
  end;

  InstalledVer := '';
  CurrentVer := '{#MyAppVersion}';
  if IsInstalled() then
  begin
    InstalledVer := GetInstalledVersion();
    
    if (InstalledVer <> '') and (CompareVersion(InstalledVer, CurrentVer) > 0) then
    begin
      V := MsgBox('检测到系统中已安装了更高版本的 QSBar (v' + InstalledVer + ')。' + #13#10 + 
             '您当前正在尝试安装较低的版本 (v' + CurrentVer + ')。' + #13#10#13#10 + 
             '继续安装将会覆盖高版本，是否确定继续？', mbConfirmation, MB_YESNO);
      if V = IDNO then
      begin
        Result := False;
        Exit;
      end;
    end;
  end;

end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep = ssInstall) and NeedPreInstallCleanup then
    RunPreInstallCleanup();
end;

function IsDotNet40Installed: Boolean;
begin
  // 简化版检测：只要存在该注册表项，就认为已安装 .NET 4.0
  Result := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full');
end;

// 获取符合 URL 格式的 CodeBase 路径
function GetCodeBase(Param: String): String;
var
  AppPath: String;
begin
  AppPath := ExpandConstant('{app}');
  // 将反斜杠替换为正斜杠，以符合 file:/// 协议
  StringChangeEx(AppPath, '\', '/', True);
  Result := 'file:///' + AppPath + '/QSBar.dll';
end;

// 卸载时的额外清理
procedure CurUninstallStepChanged(UninstallStep: TUninstallStep);
begin
  if UninstallStep = usPostUninstall then
  begin
    // 强制清理可能残留的注册表项 (HKCU)
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}');
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\QSBar.WpsAddIn');
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\ET\AddinsWL', 'QSBar.WpsAddIn');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\WPS\AddinsWL', 'QSBar.WpsAddIn');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\Common\AddinsWL', 'QSBar.WpsAddIn');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\6.0\AddinsWL', 'QSBar.WpsAddIn');
    
    // 强制清理可能残留的注册表项 (HKLM)
    RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}');
    RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\Classes\QSBar.WpsAddIn');
    RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn');
    
    // 强制清理 WOW6432Node
    RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\WOW6432Node\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}');
    RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\WOW6432Node\Classes\QSBar.WpsAddIn');
    RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\WOW6432Node\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn');
  end;
end;

procedure InitializeWizard;
var
  ExcelLabel, WpsLabel: TLabel;
begin
  // --- 环境检测：检查是否安装了 .NET 4.0 ---
  if not IsDotNet40Installed then
  begin
    MsgBox('This application requires .NET Framework 4.0 or higher.' + #13#10 +
           'Please install .NET Framework and try again.', mbCriticalError, MB_OK);
    WizardForm.Close;
    exit;
  end;

  // 提取临时图片文件
  ExtractTemporaryFile('SHOW1_EXCEL_BAR.bmp');
  ExtractTemporaryFile('SHOW2_WPS_BAR.bmp');

  // 调整 ReadyMemo（准备安装界面的文本框）的位置 and 高度，为图片腾出空间
  // ReadyPage 是安装前的最后一个确认页面
  WizardForm.ReadyMemo.Top := WizardForm.ReadyMemo.Top + ScaleY(140);
  WizardForm.ReadyMemo.Height := WizardForm.ReadyMemo.Height - ScaleY(140);

  // Excel 展示图
  ExcelImage := TBitmapImage.Create(WizardForm);
  ExcelImage.Parent := WizardForm.ReadyPage;
  ExcelImage.Left := WizardForm.ReadyLabel.Left;
  ExcelImage.Top := WizardForm.ReadyLabel.Top + WizardForm.ReadyLabel.Height + ScaleY(30);
  ExcelImage.Width := ScaleX(400);
  ExcelImage.Height := ScaleY(50);
  ExcelImage.Stretch := True;
  ExcelImage.Bitmap.LoadFromFile(ExpandConstant('{tmp}\SHOW1_EXCEL_BAR.bmp'));
  ExcelImage.BringToFront;

  ExcelLabel := TLabel.Create(WizardForm);
  ExcelLabel.Parent := WizardForm.ReadyPage;
  ExcelLabel.Caption := 'Excel Interface Preview:';
  ExcelLabel.Left := ExcelImage.Left;
  ExcelLabel.Top := ExcelImage.Top - ScaleY(18);
  ExcelLabel.Font.Style := [fsBold];
  ExcelLabel.BringToFront;

  // WPS 展示图
  WpsImage := TBitmapImage.Create(WizardForm);
  WpsImage.Parent := WizardForm.ReadyPage;
  WpsImage.Left := ExcelImage.Left;
  WpsImage.Top := ExcelImage.Top + ExcelImage.Height + ScaleY(30);
  WpsImage.Width := ScaleX(400);
  WpsImage.Height := ScaleY(50);
  WpsImage.Stretch := True;
  WpsImage.Bitmap.LoadFromFile(ExpandConstant('{tmp}\SHOW2_WPS_BAR.bmp'));
  WpsImage.BringToFront;

  WpsLabel := TLabel.Create(WizardForm);
  WpsLabel.Parent := WizardForm.ReadyPage;
  WpsLabel.Caption := 'WPS Interface Preview:';
  WpsLabel.Left := WpsImage.Left;
  WpsLabel.Top := WpsImage.Top - ScaleY(18);
  WpsLabel.Font.Style := [fsBold];
  WpsLabel.BringToFront;
end;
