unit uFormSettings;

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.Samples.Spin;

const
  DefaultHttpPort = 2580;
  DefaultConnectionType = 'usb';
  DefaultComNumber = 3;
  DefaultIpPort = 5555;
  DefaultCashierName = 'Кассир';
  DefaultUiMaxLines = 2000;
  DefaultMaxBodyLen = 4096;

type
  TFormSettings = class(TForm)
    ScrollBox: TScrollBox;
    gbHttpServer: TGroupBox;
    lblHttpPort: TLabel;
    seHttpPort: TSpinEdit;
    gbFRCash: TGroupBox;
    lblConnectionType: TLabel;
    cbConnectionType: TComboBox;
    lblComNumber: TLabel;
    seComNumber: TSpinEdit;
    lblIpAddress: TLabel;
    edIpAddress: TEdit;
    lblIpPort: TLabel;
    seIpPort: TSpinEdit;
    lblRemoteServerAddr: TLabel;
    edRemoteServerAddr: TEdit;
    lblCashierName: TLabel;
    edCashierName: TEdit;
    lblCashierInn: TLabel;
    edCashierInn: TEdit;
    gbKKT: TGroupBox;
    chkEmulation: TCheckBox;
    chkTestReceiptMode: TCheckBox;
    pnlEmulation: TPanel;
    lblEmulatedSerial: TLabel;
    edEmulatedSerial: TEdit;
    lblEmulatedInn: TLabel;
    edEmulatedInn: TEdit;
    lblEmulatedFnNumber: TLabel;
    edEmulatedFnNumber: TEdit;
    lblEmulatedCheckNumber: TLabel;
    seEmulatedCheckNumber: TSpinEdit;
    lblEmulatedSessionNumber: TLabel;
    seEmulatedSessionNumber: TSpinEdit;
    gbLog: TGroupBox;
    chkLogEnabled: TCheckBox;
    chkLogUiEnabled: TCheckBox;
    lblLogUiMaxLines: TLabel;
    seLogUiMaxLines: TSpinEdit;
    lblLogPath: TLabel;
    edLogPath: TEdit;
    btnBrowseLogPath: TButton;
    lblLogLevel: TLabel;
    cbLogLevel: TComboBox;
    lblLogMaxBodyLen: TLabel;
    seLogMaxBodyLen: TSpinEdit;
    lblRestartHint: TLabel;
    pnlButtons: TPanel;
    btnOk: TButton;
    btnCancel: TButton;
    dlgSaveLog: TSaveDialog;
    procedure FormCreate(Sender: TObject);
    procedure chkEmulationClick(Sender: TObject);
    procedure cbConnectionTypeChange(Sender: TObject);
    procedure btnBrowseLogPathClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
  private
    FSettings: TCustomIniFile;
    FServerRunning: Boolean;
    procedure LoadFromIni;
    procedure SaveToIni;
    procedure UpdateEmulationPanel;
    procedure UpdateConnectionTypePanel;
    procedure ApplyRunningRestrictions;
    function ValidateInput: Boolean;
    function ConnectionTypeToComboIndex(const ConnType: string): Integer;
    function ComboIndexToConnectionType(Index: Integer): string;
  public
    /// <summary>%ProgramData%\CTO_KSM\kktserver — без прав админа.</summary>
    class function GetDataDirectory: string;
    class function GetSettingsFileName: string;
    class function GetDefaultLogFileName: string;
    class function Execute(ASettings: TCustomIniFile; ServerRunning: Boolean): Boolean;
    /// <summary>Записывает полный набор настроек по умолчанию (первый запуск).</summary>
    class procedure WriteDefaultSettings(ASettings: TCustomIniFile);
  end;

implementation

{$R *.dfm}

procedure TFormSettings.FormCreate(Sender: TObject);
begin
  cbConnectionType.Items.Clear;
  cbConnectionType.Items.Add('USB');
  cbConnectionType.Items.Add('COM');
  cbConnectionType.Items.Add('TCP-IP');

  cbLogLevel.Items.Clear;
  cbLogLevel.Items.Add('off');
  cbLogLevel.Items.Add('error');
  cbLogLevel.Items.Add('info');
  cbLogLevel.Items.Add('debug');

  seHttpPort.MinValue := 1;
  seHttpPort.MaxValue := 65535;
  seComNumber.MinValue := 1;
  seComNumber.MaxValue := 256;
  seIpPort.MinValue := 1;
  seIpPort.MaxValue := 65535;
  seLogUiMaxLines.MinValue := 100;
  seLogUiMaxLines.MaxValue := 100000;
  seLogMaxBodyLen.MinValue := 256;
  seLogMaxBodyLen.MaxValue := 1048576;
  seEmulatedCheckNumber.MinValue := 0;
  seEmulatedCheckNumber.MaxValue := 999999999;
  seEmulatedSessionNumber.MinValue := 0;
  seEmulatedSessionNumber.MaxValue := 999999999;
end;

function TFormSettings.ConnectionTypeToComboIndex(const ConnType: string): Integer;
var
  T: string;
begin
  T := LowerCase(Trim(ConnType));
  if T = 'com' then
    Result := 1
  else if (T = 'ip') or (T = 'tcp') or (T = 'tcpip') or (T = 'tcp-ip') then
    Result := 2
  else
    Result := 0; // usb
end;

function TFormSettings.ComboIndexToConnectionType(Index: Integer): string;
begin
  case Index of
    1: Result := 'com';
    2: Result := 'ip';
  else
    Result := 'usb';
  end;
end;

procedure TFormSettings.LoadFromIni;
var
  LevelStr: string;
  LevelIndex: Integer;
  ConnType: string;
  LegacyPort: Integer;
begin
  seHttpPort.Value := FSettings.ReadInteger('HttpServer', 'Port', DefaultHttpPort);

  ConnType := FSettings.ReadString('FRCash', 'ConnectionType', '');
  if ConnType = '' then
    ConnType := DefaultConnectionType; // usb

  // Если ComNumber ещё не задан — подхватить старый Port из INI только как номер COM
  if FSettings.ValueExists('FRCash', 'ComNumber') then
    seComNumber.Value := FSettings.ReadInteger('FRCash', 'ComNumber', DefaultComNumber)
  else
  begin
    LegacyPort := FSettings.ReadInteger('FRCash', 'Port', 0);
    if LegacyPort > 0 then
      seComNumber.Value := LegacyPort
    else
      seComNumber.Value := DefaultComNumber;
  end;

  cbConnectionType.ItemIndex := ConnectionTypeToComboIndex(ConnType);
  edIpAddress.Text := FSettings.ReadString('FRCash', 'IpAddress', '');
  seIpPort.Value := FSettings.ReadInteger('FRCash', 'IpPort', DefaultIpPort);
  edRemoteServerAddr.Text := FSettings.ReadString('FRCash', 'RemoteServerAddr', '');
  edCashierName.Text := FSettings.ReadString('FRCash', 'CashierName', DefaultCashierName);
  edCashierInn.Text := FSettings.ReadString('FRCash', 'CashierInn', '');

  chkEmulation.Checked := FSettings.ReadBool('KKT', 'Emulation', False);
  chkTestReceiptMode.Checked := FSettings.ReadBool('KKT', 'TestReceiptMode', False);
  edEmulatedSerial.Text := FSettings.ReadString('KKT', 'EmulatedSerial', '0000000000000001');
  edEmulatedInn.Text := FSettings.ReadString('KKT', 'EmulatedInn', '1234567890');
  edEmulatedFnNumber.Text := FSettings.ReadString('KKT', 'EmulatedFnNumber', '9999078900000001');
  seEmulatedCheckNumber.Value := FSettings.ReadInteger('KKT', 'EmulatedCheckNumber', 0);
  seEmulatedSessionNumber.Value := FSettings.ReadInteger('KKT', 'EmulatedSessionNumber', 0);

  chkLogEnabled.Checked := FSettings.ReadBool('Log', 'Enabled', True);
  chkLogUiEnabled.Checked := FSettings.ReadBool('Log', 'UiEnabled', True);
  seLogUiMaxLines.Value := FSettings.ReadInteger('Log', 'UiMaxLines', DefaultUiMaxLines);
  edLogPath.Text := FSettings.ReadString('Log', 'Path', GetDefaultLogFileName);
  seLogMaxBodyLen.Value := FSettings.ReadInteger('Log', 'MaxBodyLen', DefaultMaxBodyLen);

  LevelStr := LowerCase(Trim(FSettings.ReadString('Log', 'Level', 'info')));
  LevelIndex := cbLogLevel.Items.IndexOf(LevelStr);
  if LevelIndex < 0 then
    LevelIndex := cbLogLevel.Items.IndexOf('info');
  cbLogLevel.ItemIndex := LevelIndex;

  UpdateEmulationPanel;
  UpdateConnectionTypePanel;
  ApplyRunningRestrictions;
end;

procedure TFormSettings.SaveToIni;
begin
  FSettings.WriteInteger('HttpServer', 'Port', seHttpPort.Value);

  FSettings.WriteString('FRCash', 'ConnectionType', ComboIndexToConnectionType(cbConnectionType.ItemIndex));
  FSettings.WriteInteger('FRCash', 'ComNumber', seComNumber.Value);
  FSettings.WriteString('FRCash', 'IpAddress', Trim(edIpAddress.Text));
  FSettings.WriteInteger('FRCash', 'IpPort', seIpPort.Value);
  FSettings.WriteString('FRCash', 'RemoteServerAddr', Trim(edRemoteServerAddr.Text));
  FSettings.WriteString('FRCash', 'CashierName', Trim(edCashierName.Text));
  FSettings.WriteString('FRCash', 'CashierInn', Trim(edCashierInn.Text));

  // Убрать устаревшие ключи Port/Baud/Password, если остались
  FSettings.DeleteKey('FRCash', 'Port');
  FSettings.DeleteKey('FRCash', 'Baud');
  FSettings.DeleteKey('FRCash', 'Password');

  FSettings.WriteBool('KKT', 'Emulation', chkEmulation.Checked);
  FSettings.WriteBool('KKT', 'TestReceiptMode', chkTestReceiptMode.Checked);
  FSettings.WriteString('KKT', 'EmulatedSerial', Trim(edEmulatedSerial.Text));
  FSettings.WriteString('KKT', 'EmulatedInn', Trim(edEmulatedInn.Text));
  FSettings.WriteString('KKT', 'EmulatedFnNumber', Trim(edEmulatedFnNumber.Text));
  FSettings.WriteInteger('KKT', 'EmulatedCheckNumber', seEmulatedCheckNumber.Value);
  FSettings.WriteInteger('KKT', 'EmulatedSessionNumber', seEmulatedSessionNumber.Value);

  FSettings.WriteBool('Log', 'Enabled', chkLogEnabled.Checked);
  FSettings.WriteBool('Log', 'UiEnabled', chkLogUiEnabled.Checked);
  FSettings.WriteInteger('Log', 'UiMaxLines', seLogUiMaxLines.Value);
  FSettings.WriteString('Log', 'Path', Trim(edLogPath.Text));
  if cbLogLevel.ItemIndex >= 0 then
    FSettings.WriteString('Log', 'Level', cbLogLevel.Items[cbLogLevel.ItemIndex])
  else
    FSettings.WriteString('Log', 'Level', 'info');
  FSettings.WriteInteger('Log', 'MaxBodyLen', seLogMaxBodyLen.Value);

  FSettings.UpdateFile;
end;

procedure TFormSettings.UpdateEmulationPanel;
begin
  pnlEmulation.Enabled := chkEmulation.Checked;
end;

procedure TFormSettings.UpdateConnectionTypePanel;
var
  ConnType: string;
  IsCom, IsIp: Boolean;
begin
  ConnType := ComboIndexToConnectionType(cbConnectionType.ItemIndex);
  IsCom := ConnType = 'com';
  IsIp := ConnType = 'ip';

  lblComNumber.Enabled := IsCom;
  seComNumber.Enabled := IsCom;
  lblIpAddress.Enabled := IsIp;
  edIpAddress.Enabled := IsIp;
  lblIpPort.Enabled := IsIp;
  seIpPort.Enabled := IsIp;
end;

procedure TFormSettings.ApplyRunningRestrictions;
begin
  gbHttpServer.Enabled := not FServerRunning;
  gbFRCash.Enabled := not FServerRunning;
  gbKKT.Enabled := not FServerRunning;

  if FServerRunning then
    lblRestartHint.Caption :=
      'Сервер запущен: HTTP, ККТ и эмуляция недоступны для изменения. ' +
      'Остановите сервер и запустите снова, чтобы применить эти параметры.'
  else
    lblRestartHint.Caption :=
      'Изменения HTTP-порта, параметров подключения ККТ и режима эмуляции ' +
      'вступят в силу после перезапуска сервера.';

  if not FServerRunning then
    UpdateConnectionTypePanel;
end;

function TFormSettings.ValidateInput: Boolean;
var
  ConnType: string;
begin
  Result := False;

  if Trim(edCashierName.Text) = '' then
  begin
    ShowMessage('Укажите имя кассира.');
    ActiveControl := edCashierName;
    Exit;
  end;

  ConnType := ComboIndexToConnectionType(cbConnectionType.ItemIndex);
  if ConnType = 'ip' then
  begin
    if Trim(edIpAddress.Text) = '' then
    begin
      ShowMessage('Укажите IP-адрес ККТ.');
      ActiveControl := edIpAddress;
      Exit;
    end;
  end;

  if Trim(edLogPath.Text) = '' then
  begin
    ShowMessage('Укажите путь к файлу лога.');
    ActiveControl := edLogPath;
    Exit;
  end;

  if cbLogLevel.ItemIndex < 0 then
  begin
    ShowMessage('Выберите уровень логирования.');
    ActiveControl := cbLogLevel;
    Exit;
  end;

  Result := True;
end;

procedure TFormSettings.chkEmulationClick(Sender: TObject);
begin
  UpdateEmulationPanel;
end;

procedure TFormSettings.cbConnectionTypeChange(Sender: TObject);
begin
  UpdateConnectionTypePanel;
end;

procedure TFormSettings.btnBrowseLogPathClick(Sender: TObject);
var
  Dir, FileName: string;
begin
  FileName := Trim(edLogPath.Text);
  if FileName = '' then
    FileName := GetDefaultLogFileName;

  dlgSaveLog.FileName := ExtractFileName(FileName);
  Dir := ExtractFilePath(FileName);
  if (Dir <> '') and DirectoryExists(Dir) then
    dlgSaveLog.InitialDir := Dir
  else
    dlgSaveLog.InitialDir := GetDataDirectory;

  if dlgSaveLog.Execute then
    edLogPath.Text := dlgSaveLog.FileName;
end;

procedure TFormSettings.btnOkClick(Sender: TObject);
begin
  if not ValidateInput then
    ModalResult := mrNone
  else
    ModalResult := mrOk;
end;

class function TFormSettings.GetDataDirectory: string;
var
  ProgramData: string;
begin
  ProgramData := GetEnvironmentVariable('PROGRAMDATA');
  if ProgramData = '' then
    ProgramData := 'C:\ProgramData';
  Result := IncludeTrailingPathDelimiter(ProgramData) + 'CTO_KSM\kktserver';
  ForceDirectories(Result);
end;

class function TFormSettings.GetSettingsFileName: string;
begin
  Result := IncludeTrailingPathDelimiter(GetDataDirectory) + 'kktserverindy.ini';
end;

class function TFormSettings.GetDefaultLogFileName: string;
begin
  Result := IncludeTrailingPathDelimiter(GetDataDirectory) + 'kktserver.log';
end;

class function TFormSettings.Execute(ASettings: TCustomIniFile; ServerRunning: Boolean): Boolean;
var
  Form: TFormSettings;
begin
  Form := TFormSettings.Create(nil);
  try
    Form.FSettings := ASettings;
    Form.FServerRunning := ServerRunning;
    Form.LoadFromIni;
    Result := Form.ShowModal = mrOk;
    if Result then
      Form.SaveToIni;
  finally
    Form.Free;
  end;
end;

class procedure TFormSettings.WriteDefaultSettings(ASettings: TCustomIniFile);
begin
  if ASettings = nil then
    Exit;

  ASettings.WriteInteger('HttpServer', 'Port', DefaultHttpPort);

  ASettings.WriteString('FRCash', 'ConnectionType', DefaultConnectionType);
  ASettings.WriteInteger('FRCash', 'ComNumber', DefaultComNumber);
  ASettings.WriteString('FRCash', 'IpAddress', '');
  ASettings.WriteInteger('FRCash', 'IpPort', DefaultIpPort);
  ASettings.WriteString('FRCash', 'RemoteServerAddr', '');
  ASettings.WriteString('FRCash', 'CashierName', DefaultCashierName);
  ASettings.WriteString('FRCash', 'CashierInn', '');

  ASettings.WriteBool('KKT', 'Emulation', False);
  ASettings.WriteBool('KKT', 'TestReceiptMode', False);
  ASettings.WriteString('KKT', 'EmulatedSerial', '0000000000000001');
  ASettings.WriteString('KKT', 'EmulatedInn', '1234567890');
  ASettings.WriteString('KKT', 'EmulatedRegNumber', '1234567111890');
  ASettings.WriteString('KKT', 'EmulatedFnNumber', '9999078900000001');
  ASettings.WriteInteger('KKT', 'EmulatedCheckNumber', 0);
  ASettings.WriteInteger('KKT', 'EmulatedSessionNumber', 0);

  ASettings.WriteBool('Log', 'Enabled', True);
  ASettings.WriteBool('Log', 'UiEnabled', True);
  ASettings.WriteInteger('Log', 'UiMaxLines', DefaultUiMaxLines);
  ASettings.WriteString('Log', 'Path', GetDefaultLogFileName);
  ASettings.WriteString('Log', 'Level', 'info');
  ASettings.WriteInteger('Log', 'MaxBodyLen', DefaultMaxBodyLen);

  ASettings.UpdateFile;
end;

end.
