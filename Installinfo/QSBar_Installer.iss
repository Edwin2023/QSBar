; QSBar One-Click Installer Script for Inno Setup
#define MyAppName "QSBar"
#define MyAppVersion "1.0.0.5"
#define MyAppPublisher "Bookmen"
#define MyAppExeName "QSBar.dll"
#define SourcePath "..\QSBar\bin\Release"
#define ScriptPath "..\scripts"

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
DefaultDirName=D:\Program Files\{#MyAppName}
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
; SetupIconFile=QSBar.ico

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
Source: "{#SourcePath}\QSBar.tlb"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.Interfaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\EPPlus.System.Drawing.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\Microsoft.IO.RecyclableMemoryStream.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\System.ComponentModel.Annotations.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\QSBar.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#ScriptPath}\Register.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "UpdateNotes.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "SHOW1_EXCEL_BAR.bmp"; DestDir: "{app}"; Flags: ignoreversion
Source: "SHOW2_WPS_BAR.bmp"; DestDir: "{app}"; Flags: ignoreversion
Source: "SHOW1_EXCEL_BAR.bmp"; Flags: dontcopy
Source: "SHOW2_WPS_BAR.bmp"; Flags: dontcopy

[Registry]
; --- Office/WPS 加载项注册 (告诉 Office 去哪里找这个 COM 组件) ---

; Excel (32-bit & 64-bit)
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar Excel Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; --- 强制 COM 注册 (解决 64位 Excel 加载失败的兜底方案) ---
; 即使 RegAsm 失败，这里的注册表项也能确保 Excel 找到组件
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"; ValueType: string; ValueName: ""; ValueData: "QSBar.WpsAddIn"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "QSBar.WpsExcelAddIn"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "QSBar, Version={#MyAppVersion}, Culture=neutral, PublicKeyToken=null"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}/{#MyAppExeName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\QSBar.WpsAddIn"; ValueType: string; ValueName: ""; ValueData: "QSBar.WpsAddIn"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\QSBar.WpsAddIn\CLSID"; ValueType: string; ValueName: ""; ValueData: "{{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}"; Flags: uninsdeletekey

; WPS 表格 (ET)
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; WPS 表格 (ET) - AddinsData (关键：防止重启后需要重新勾选)
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsData\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; WPS 通用 - AddinsData (多处备份，确保自动加载)
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\AddinsData\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\AddinsData\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\AddinsData\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\AddinsData\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; WPS 通用
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "Description"; ValueData: "QSBar WPS Productivity Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; --- WPS 白名单 (防止被禁用) ---
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\WPS\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\Common\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Kingsoft\Office\6.0\AddinsWL"; ValueType: string; ValueName: "QSBar.WpsAddIn"; ValueData: ""; Flags: uninsdeletevalue

[Run]
; 安装前清理旧的注册信息
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; WorkingDir: "{app}"; Flags: runhidden; StatusMsg: "Cleaning up old 32-bit registration..."; BeforeInstall: TaskKillOffice
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; WorkingDir: "{app}"; Flags: runhidden; StatusMsg: "Cleaning up old 64-bit registration..."; Check: IsWin64

; 核心注册逻辑：调用 bat 脚本进行注册，确保环境与手动执行一致
Filename: "{app}\Register_QSBar.bat"; Parameters: ""; WorkingDir: "{app}"; StatusMsg: "Registering QSBar (Finalizing)..."

Filename: "{app}\UpdateNotes.html"; Description: "View update instructions (查看更新说明)"; Flags: postinstall shellexec skipifsilent

[UninstallRun]
; 卸载时彻底反注册 COM 组件
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; WorkingDir: "{app}"; Flags: runhidden; StatusMsg: "Unregistering 32-bit COM..."
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/u ""{app}\QSBar.dll"""; WorkingDir: "{app}"; Flags: runhidden; StatusMsg: "Unregistering 64-bit COM..."; Check: IsWin64

[UninstallDelete]
Type: filesandordirs; Name: "{app}"


[Code]
var
  ExcelImage, WpsImage: TBitmapImage;

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

// 安装程序初始化时执行
function InitializeSetup(): Boolean;
var
  V: Integer;
  iResultCode: Integer;
  sUnInstallString: String;
begin
  Result := True;
  if IsInstalled() then
  begin
    V := MsgBox('A version of QSBar is already installed. Do you want to uninstall it before continuing?' + #13#10 +
                '检测到已安装旧版本，是否在继续安装前先卸载？', mbConfirmation, MB_YESNO);
    if V = IDYES then
    begin
      sUnInstallString := RemoveQuotes(GetUninstallString());
      if Exec(sUnInstallString, '/SILENT /VERYSILENT /SUPPRESSMSGBOXES', '', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      begin
        // 卸载成功后继续
        Result := True;
      end
      else
      begin
        MsgBox('Uninstallation failed. Please uninstall manually and try again.' + #13#10 +
               '自动卸载失败，请手动卸载后再试。', mbError, MB_OK);
        Result := False;
      end;
    end;
  end;
end;

function IsDotNet40Installed: Boolean;
begin
  // 简化版检测：只要存在该注册表项，就认为已安装 .NET 4.0
  Result := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full');
end;

procedure TaskKillOffice();
var
  ResultCode: Integer;
begin
  Exec('taskkill.exe', '/f /im excel.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im wps.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/f /im et.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
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
    // 强制清理可能残留的注册表项
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\CLSID\{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}');
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Classes\QSBar.WpsAddIn');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\ET\AddinsWL', 'QSBar.WpsAddIn');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\WPS\AddinsWL', 'QSBar.WpsAddIn');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\Common\AddinsWL', 'QSBar.WpsAddIn');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Kingsoft\Office\6.0\AddinsWL', 'QSBar.WpsAddIn');
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
