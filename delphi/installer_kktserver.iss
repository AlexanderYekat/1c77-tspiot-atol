; Installation script for Delphi KKT server (GUI + HTTP)
#define MyAppName "KKT Server TSPiOT"
#define MyAppVersion "2026.07.23.03"
#define MyAppPublisher "CTO KSM"
#define MyAppExeName "kktserverindyProject.exe"
#define MyAppDirName "kktserver"
#define MyAppIniName "kktserverindy.ini"
#define DefaultHttpPort "2580"

[Setup]
AppId={{A3C71E8B-4F2D-4A9E-9C1B-6D8E5F0A2B47}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://cto-ksm.ru
AppSupportURL=https://cto-ksm.ru
AppUpdatesURL=https://cto-ksm.ru
; Uniform tree with other CTO KSM products (Latin path)
DefaultDirName={autopf}\CTO_KSM\{#MyAppDirName}
DefaultGroupName=CTO KSM\{#MyAppName}
OutputDir=output
OutputBaseFilename=kktserver_setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
; Остановку делаем сами через HTTP /exit (закрытие окна = трей, не выход)
CloseApplications=no
RestartApplications=no
MinVersion=6.1sp1
DisableProgramGroupPage=yes
InfoBeforeFile=installer_info_ru.txt

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
; Shared settings/logs — writable without admin
Name: "{commonappdata}\CTO_KSM\{#MyAppDirName}"; Permissions: users-modify

[Files]
; Берём свежую сборку Delphi (Win32 Debug)
Source: "Win32\Debug\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Settings live in %ProgramData%\CTO_KSM\kktserver\ (created on first run)

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
; Mandatory autostart so HTTP comes up after reboot (FormCreate starts the server)
Name: "{commonstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Run]
; Launch after install — HTTP starts automatically in FormCreate
Filename: "{app}\{#MyAppExeName}"; Flags: nowait postinstall skipifsilent; Description: "{cm:LaunchProgram,{#MyAppName}}"

[Code]
function GetServiceHttpPort: Integer;
var
  IniPath, PortStr: String;
begin
  IniPath := ExpandConstant('{commonappdata}\CTO_KSM\{#MyAppDirName}\{#MyAppIniName}');
  PortStr := GetIniString('HttpServer', 'Port', '{#DefaultHttpPort}', IniPath);
  Result := StrToIntDef(PortStr, {#DefaultHttpPort});
end;

procedure RequestHttpExit;
var
  ResultCode: Integer;
  Port: Integer;
  Url: String;
begin
  Port := GetServiceHttpPort;
  Url := 'http://127.0.0.1:' + IntToStr(Port) + '/exit';
  // Штатный выход новых сборок (hide-to-tray). curl есть в Windows 10+.
  Exec(
    ExpandConstant('{cmd}'),
    '/C curl.exe -s -m 5 -X POST "' + Url + '" >nul 2>&1',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

function IsProcessRunning(const ImageName: String): Boolean;
var
  ResultCode: Integer;
begin
  Exec(
    ExpandConstant('{cmd}'),
    '/C tasklist /NH /FI "IMAGENAME eq ' + ImageName + '" 2>nul | find /I "' + ImageName + '" >nul',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := ResultCode = 0;
end;

procedure WaitForProcessExit(const ImageName: String; TimeoutMs: Integer);
var
  Elapsed: Integer;
begin
  Elapsed := 0;
  while (Elapsed < TimeoutMs) and IsProcessRunning(ImageName) do
  begin
    Sleep(500);
    Elapsed := Elapsed + 500;
  end;
end;

procedure EnsureProcessStopped(const ImageName: String; TimeoutMs: Integer);
var
  ResultCode: Integer;
begin
  if not IsProcessRunning(ImageName) then
    Exit;
  RequestHttpExit;
  WaitForProcessExit(ImageName, TimeoutMs);
  if IsProcessRunning(ImageName) then
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM "' + ImageName + '" /T', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  NeedsRestart := False;
  EnsureProcessStopped('kktserverindyProject.exe', 15000);
  EnsureProcessStopped('kktserverindy.exe', 15000);
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  EnsureProcessStopped('kktserverindyProject.exe', 15000);
  EnsureProcessStopped('kktserverindy.exe', 15000);
end;
