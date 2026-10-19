; ═══════════════════════════════════════════════════════════
; 地址坐标转换工具 - Inno Setup 安装脚本
; 每次发新版，只需修改第 8 行的版本号即可
; ═══════════════════════════════════════════════════════════

#define MyAppName      "地址坐标转换工具"
#define MyAppVersion   "2.8.66"
#define MyAppURL       "https://addressmap.mypdftool1.top"
#define MyAppExeName   "addressmap_tool_v" + MyAppVersion + ".exe"
#define MyAppSetupName "setup_addressmap_tool_v" + MyAppVersion
#define MyAppId        "{{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppSupportURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename={#MyAppSetupName}
OutputDir=.\Output
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
CloseApplications=force
CloseApplicationsFilter=*.exe
RestartApplications=no
AllowNoIcons=yes
DisableDirPage=auto
DisableProgramGroupPage=auto
ShowLanguageDialog=no
MinVersion=6.1sp1

[Files]
Source: ".\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\{#MyAppExeName}"; Flags: nowait shellexec runhidden; Check: IsVerySilent

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/F /IM ""{#MyAppExeName}"""; Flags: runhidden waituntilterminated; RunOnceId: "KillApp"

[UninstallDelete]
Type: files; Name: "{tmp}\setup_addressmap_tool_*.exe"

[Code]
var
  VerySilentMode: Boolean;

function IsVerySilent(): Boolean;
begin
  Result := VerySilentMode;
end;

function InitializeSetup(): Boolean;
var
  i: Integer;
  s: String;
begin
  VerySilentMode := False;
  for i := 0 to ParamCount do
  begin
    s := AnsiUppercase(ParamStr(i));
    if (s = '/VERYSILENT') or (s = '/VERYSLIENT') then
    begin
      VerySilentMode := True;
      Break;
    end;
  end;
  
  if FileExists(ExpandConstant('{sys}\taskkill.exe')) then
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM ""{#MyAppExeName}""', '', SW_HIDE, ewWaitUntilTerminated, i)
  else
    ShellExec('', 'taskkill', '/F /IM ""{#MyAppExeName}""', '', SW_HIDE, ewWaitUntilTerminated, i);
  
  Sleep(500);
  Result := True;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  i: Integer;
begin
  if FileExists(ExpandConstant('{sys}\taskkill.exe')) then
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM ""{#MyAppExeName}""', '', SW_HIDE, ewWaitUntilTerminated, i)
  else
    ShellExec('', 'taskkill', '/F /IM ""{#MyAppExeName}""', '', SW_HIDE, ewWaitUntilTerminated, i);
  Sleep(300);
  Result := '';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  i: Integer;
begin
  if (CurStep = ssPostInstall) and WizardSilent then
    ShellExec('', ExpandConstant('{app}\{#MyAppExeName}'), '', '', SW_SHOWNORMAL, ewNoWait, i);
end;

function InitializeUninstall(): Boolean;
var
  i: Integer;
begin
  if FileExists(ExpandConstant('{sys}\taskkill.exe')) then
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM ""{#MyAppExeName}""', '', SW_HIDE, ewWaitUntilTerminated, i)
  else
    ShellExec('', 'taskkill', '/F /IM ""{#MyAppExeName}""', '', SW_HIDE, ewWaitUntilTerminated, i);
  Sleep(500);
  Result := True;
end;