unit uVariantPrint;

interface

uses
  System.SysUtils, System.Variants, System.JSON, System.IniFiles,
  System.SyncObjs, System.NetEncoding, System.Generics.Collections,
  System.Classes, System.DateUtils, Winapi.Windows, Winapi.ActiveX,
  ComObj, uKktLog, uFormSettings;

type
  TCheckMarkTaskEntry = class
    Status: string;
    ResultJson: string;
    CreatedAt: TDateTime;
  end;

  /// <summary>Кэш результата проверки одной марки (ККТ + РР) в рамках текущего чека.</summary>
  TMarkCacheEntry = record
    KktAccepted: Boolean;
    KktResultCode: Integer;
    KktDescription: string;
    KktWarning: string;
    /// <summary>JSON ответа getMarkingCodeValidationStatus (для imcParams.itemInfoCheckResult).</summary>
    KktValidationJson: string;
    RrChecked: Boolean;
    RrAccepted: Boolean;
    RrCode: Integer;
    RrDescription: string;
    Uuid: string;
    TimeStr: string;
    Inst: string;
    Version: string;
    CheckedAt: TDateTime;
  end;

  TVariantPrint = class
  private
    ovObject: OleVariant;      // COM-объект драйвера ККТ (AddIn.Fptr10) — держим один на всё время
    FSettings: TCustomIniFile;
    FLog: TKktLogger;
    FLock: TCriticalSection;   // защита от параллельных обращений к ККТ
    FMarkTasksLock: TCriticalSection;
    FMarkTasks: TObjectDictionary<string, TCheckMarkTaskEntry>;
    FMarkCache: TDictionary<string, TMarkCacheEntry>; // кэш ККТ+РР в рамках текущего чека
    FConnected: Boolean;       // текущее состояние соединения, как мы его понимаем
    FEmulation: Boolean;
    FTestReceiptMode: Boolean; // тестовый режим: cancelReceipt вместо фискализации

    FEmulatedSerial: string;
    FEmulatedInn: string;
    FEmulatedRegNumber: string;
    FEmulatedFnNumber: string;
    FEmulatedCheckNumber: Integer;
    FEmulatedSessionNumber: Integer;

    // STA-поток для AddIn.Fptr10: Indy и async-задачи вызывают COM не из main thread
    FFrThread: TThread;
    FFrThreadId: TThreadID;
    FFrWorkAvail: TEvent;
    FFrWorkDone: TEvent;
    FFrReady: TEvent;
    FFrRequest: TProc;
    FFrError: string;
    FFrStop: Boolean;

    procedure StartFrComThread;
    procedure StopFrComThread;
    procedure ExecOnFrThread(const Proc: TProc);
    function EnsureFRConnected: Boolean;
    procedure DoDisconnect;
    function CheckLinkAlive: Boolean;
    function ReadSerialNumber: string;
    function FRObjectReady: Boolean;

    procedure ApplyAtolConnectionSettings;
    function AtolErrorCode: Integer;
    function AtolErrorDescription: string;
    /// <summary>Единая точка processJson (как ВыполнитьJSONАТОЛ). True = успех.</summary>
    function ProcessJsonAtol(const JsonText, Stage: string;
      out ResponseJson: string; out ErrorCode: Integer;
      out ErrorDescription: string): Boolean;

    /// <summary>Клиентский JSON чека → JSON-задание АТОЛ sell/sellReturn (как СформироватьJSONЧека).</summary>
    function BuildAtolReceiptJson(Req: TJSONObject): string;
    function BuildAtolReceiptItem(Item: TJSONObject; ReturnCheck: Boolean;
      DefaultDepartment: Integer): TJSONObject;
    function BuildAtolImcParams(Item: TJSONObject; const MarkB64: string;
      ReturnCheck: Boolean): TJSONObject;
    function BuildAtolIndustryInfo(Item: TJSONObject; const MarkDecoded: string;
      ReturnCheck: Boolean): TJSONArray;
    function BuildAtolPayments(Req: TJSONObject): TJSONArray;
    /// <summary>Клиентский nonFiscal (слип) → JSON-задание АТОЛ nonFiscal.</summary>
    function BuildAtolNonFiscalJson(Req: TJSONObject): string;
    function BuildAtolNonFiscalItem(Item: TJSONObject): TJSONObject;
    function PrintNonFiscalFromRequest(Req: TJSONObject): string;
    function ExtractItemInfoCheckResult(const ValidationJson: string): TJSONObject;
    function DefaultItemInfoCheckResult(Accepted: Boolean): TJSONObject;
    function ExtractCheckNumberFromAtolResponse(const ResponseJson: string): Integer;

    procedure AfterCommand(Req: TJSONObject);

    function DecodeMarkBase64(const B64: string): string;
    function GetReceiptReturnCheck(Req: TJSONObject): Boolean;
    function GetCashierName(Req: TJSONObject): string;
    function GetCashierInn(Req: TJSONObject): string;
    procedure ClearStuckOpenDocument(const Context: string);
    function DoOpenSession(const CashierName: string;
      const CashierInn: string = ''): Integer;
    procedure EnsureSessionOpen(const CashierName: string;
      const CashierInn: string = '');
    procedure EnsureReadyForXReport;
    procedure EnsureReadyForZReport;
    function BuildAtolJsonGetShiftStatus: string;
    function BuildAtolJsonOpenShift(const CashierName, CashierInn: string): string;
    function BuildAtolJsonCloseShift(const CashierName, CashierInn: string): string;
    function BuildAtolJsonReportX(const CashierName, CashierInn: string): string;
    function BuildAtolJsonContinuePrint: string;
    function BuildAtolJsonBeginMarkValidation(const MarkB64, EstimatedStatus: string): string;
    function BuildAtolJsonGetMarkValidationStatus: string;
    function BuildAtolJsonAcceptMarkingCode: string;
    function BuildAtolJsonCancelMarkValidation: string;
    function BuildAtolJsonClearMarkValidationResult: string;
    function MockAtolJsonResponse(const JsonRequest: string): string;
    function ExtractShiftNumberFromAtolResponse(const ResponseJson: string): Integer;
    function QueryShiftStatus(out Opened, Expired: Boolean;
      out ShiftNumber: Integer): Boolean;
    function CancelReceiptAtol(const Stage: string;
      RaiseOnError: Boolean = True): Boolean;
    function GetEstimatedMarkStatus(ReturnCheck, DraftBeer,
      DraftBeerLeftovers: Boolean): string;
    function GetMarkValidationTimeoutSec: Integer;
    function IsTspiotRetryCode(Code: Integer): Boolean;
    procedure CancelMarkValidationQuiet;
    function WaitMarkValidationReady(CanRetryTspiot: Boolean;
      out ResponseJson: string; out ErrorCode: Integer;
      out ErrorDescription: string; out NeedRestart, TimedOut: Boolean): Boolean;
    procedure ClassifyMarkValidationResponse(const ResponseJson: string;
      out Accepted: Boolean; out Description, Warning: string;
      out ResultCode: Integer);
    function CheckSingleMark(const MarkDecoded, MarkB64: string;
      ReturnCheck: Boolean; DraftBeer: Boolean = False;
      DraftBeerLeftovers: Boolean = False): TJSONObject;
    function CheckMarkInternal(const Body: string): string;
    procedure TrimOldMarkTasks;
    procedure StartAsyncCheckMark(const Body, TaskId: string);
    function NewTaskId: string;
    procedure CollectMarksFromRequest(Req: TJSONObject; out MarksB64: TArray<string>);
    procedure CollectMarksFromReceiptItems(Req: TJSONObject; out MarksB64: TArray<string>);
    function EnsureMarksReadyForPrint(Req: TJSONObject): string;

    function MarkCacheKey(const MarkDecoded: string; ReturnCheck: Boolean): string;
    function TryGetMarkFromCache(const MarkDecoded: string; ReturnCheck: Boolean;
      out Entry: TMarkCacheEntry): Boolean;
    function IsMarkFullyOk(const MarkDecoded: string; ReturnCheck,
      CheckPermission: Boolean): Boolean;
    procedure UpsertMarkCacheKkt(const MarkDecoded: string; ReturnCheck: Boolean;
      KktAccepted: Boolean; KktResultCode: Integer; const KktDescription: string;
      const KktWarning: string = ''; const KktValidationJson: string = '');
    procedure UpsertMarkCacheRr(const MarkDecoded: string; ReturnCheck: Boolean;
      RrAccepted: Boolean; RrCode: Integer; const RrDescription, Uuid, TimeStr,
      Inst, Version: string);
    procedure UpdateMarkCacheFromMarkObj(const MarkDecoded: string;
      ReturnCheck: Boolean; MarkObj: TJSONObject);
    procedure ClearMarkCache;
    function MarkObjFromCacheEntry(const Entry: TMarkCacheEntry): TJSONObject;
    function PermissionObjFromCacheEntry(const Entry: TMarkCacheEntry): TJSONObject;

    procedure ReadInnAndFn(out Inn, FnNumber: string);
    function ResolvePermissionInn(Req: TJSONObject): string;
    function ResolvePermissionFn(Req: TJSONObject): string;
    function BuildDocumentCheckJson(const MarkB64, Inn, Fn: string): string;
    function HttpPostJson(const Url, Body: string): string;
    function ExtractFirstTruemarkCode(Truemark: TJSONObject): TJSONObject;
    procedure AddTruemarkFields(PermObj, Truemark: TJSONObject);
    function ParsePermissionCheckResponse(const ResponseBody: string): TJSONObject;
    function CheckSingleMarkPermission(const MarkB64, Inn, Fn: string): TJSONObject;
  public
    constructor Create(ASettings: TCustomIniFile);
    destructor Destroy; override;

    function KktInfoJson: string;
    function KktConnectionStatusJson: string;
    function PrintCheckJson(const Body: string): string;
    function CloseShiftJson(const Body: string): string;
    function XReportJson(const Body: string = ''): string;
    function OpenShiftJson(const Body: string = ''): string;
    function CancelCheckJson(const Body: string = ''): string;
    function DisconnectJson: string;
    function ConnectJson: string;
    function ClearBufferOfMarksJson(const Body: string = ''): string;
    function CheckMarkJson(const Body: string): string;
    function CheckMarkStatusJson(const TaskId: string): string;

    property Logger: TKktLogger read FLog;
    property Emulation: Boolean read FEmulation;
  end;

implementation

uses
  IdHTTP;

const
  PERMISSION_FMU_HOST = 'localhost';
  PERMISSION_FMU_PORT = 2578;
  PERMISSION_FMU_PATH = '/document';

  ATOL_PROG_ID = 'AddIn.Fptr10';
  // Fallback-константы Fptr10 (если свойства COM недоступны)
  ATOL_SETTING_PORT = 'Port';
  ATOL_SETTING_COM_FILE = 'ComFile';
  ATOL_SETTING_IPADDRESS = 'IPAddress';
  ATOL_SETTING_IPPORT = 'IPPort';
  ATOL_SETTING_REMOTE_SERVER_ADDR = 'RemoteServerAddr';
  ATOL_PORT_COM = 0;
  ATOL_PORT_USB = 1;
  ATOL_PORT_TCPIP = 2;
  ATOL_PARAM_JSON_DATA = 65645;

  // Проверка марок (как ТаймаутПроверкиМарки / ретраи ТС ПИоТ в BSL)
  ATOL_MARK_DEFAULT_TIMEOUT_SEC = 15;
  ATOL_MARK_POLL_INTERVAL_MS = 1000;
  ATOL_MARK_TSPIOT_WAIT_MS = 5000;
  ATOL_MARK_MAX_TSPIOT_RETRIES = 3;

{-------------------------------------------------------------------------------
 Вспомогательные функции уровня модуля
-------------------------------------------------------------------------------}

function JsonEscape(const S: string): string;
begin
  Result := S.Replace('\', '\\')
             .Replace('"', '\"')
             .Replace(#13, '\r')
             .Replace(#10, '\n')
             .Replace(#9, '\t')
             .Replace(#29, '\u001D');
end;

function UnescapeJsonUnicode(const S: string): string;
var
  I, Code: Integer;
  Hex: string;
begin
  Result := '';
  I := 1;
  while I <= Length(S) do
  begin
    if (I + 5 <= Length(S)) and (S[I] = '\') and (UpCase(S[I + 1]) = 'U') then
    begin
      Hex := Copy(S, I + 2, 4);
      if TryStrToInt('$' + Hex, Code) then
      begin
        Result := Result + WideChar(Code);
        Inc(I, 6);
        Continue;
      end;
    end;
    Result := Result + S[I];
    Inc(I);
  end;
end;

function JsonToPlainText(const V: TJSONAncestor): string;
begin
  // ToJSON в старых Delphi экранирует кириллицу как \uXXXX; раскодируем для 1С 7.7 и логов.
  Result := UnescapeJsonUnicode(V.ToJSON);
end;

function JsonBodyPreviewPlain(const Body: string; MaxLen: Integer = 300): string;
var
  Parsed: TJSONValue;
begin
  Result := Body;
  Parsed := TJSONObject.ParseJSONValue(Body);
  if Parsed = nil then
  begin
    if Length(Result) > MaxLen then
      Result := Copy(Result, 1, MaxLen) + '...';
    Exit;
  end;
  try
    if Parsed is TJSONAncestor then
      Result := JsonToPlainText(TJSONAncestor(Parsed))
    else
      Result := UnescapeJsonUnicode(Parsed.ToJSON);
  finally
    Parsed.Free;
  end;
  if Length(Result) > MaxLen then
    Result := Copy(Result, 1, MaxLen) + '...';
end;

function JsonOk(const Extra: string = ''): string;
begin
  Result := '{"result":1,"description":"OK"';
  if Extra <> '' then
    Result := Result + ',' + Extra;
  Result := Result + '}';
end;

function JsonError(const Msg: string): string;
begin
  Result := '{"result":0,"description":"' + JsonEscape(Msg) + '"}';
end;

function JsonBodyPreview(const Body: string; MaxLen: Integer = 300): string;
begin
  Result := Body;
  if Length(Result) > MaxLen then
    Result := Copy(Result, 1, MaxLen) + '...';
end;

function ParseJsonObjectOrRaise(const Body: string; const LogPrefix: string;
  ALog: TKktLogger): TJSONObject;
var
  V: TJSONValue;
begin
  if Trim(Body) = '' then
    raise Exception.Create('Empty request body');
  V := TJSONObject.ParseJSONValue(Body);
  if not (V is TJSONObject) then
  begin
    if Assigned(ALog) then
      ALog.LogError(Format('%s: Invalid JSON, bodyLen=%d, head=%s',
        [LogPrefix, Length(Body), JsonBodyPreview(Body)]));
    if V <> nil then
      V.Free;
    raise Exception.Create(Format('Invalid JSON, bodyLen=%d', [Length(Body)]));
  end;
  Result := TJSONObject(V);
end;

function JStr(O: TJSONObject; const Name: string; const Def: string = ''): string;
var
  V: TJSONValue;
begin
  Result := Def;
  if O = nil then Exit;
  V := O.GetValue(Name);
  if V <> nil then
    Result := V.Value;
end;

function JInt(O: TJSONObject; const Name: string; const Def: Integer = 0): Integer;
begin
  Result := StrToIntDef(JStr(O, Name, IntToStr(Def)), Def);
end;

function JBool(O: TJSONObject; const Name: string; const Def: Boolean = False): Boolean;
var
  V: TJSONValue;
begin
  Result := Def;
  if O = nil then Exit;
  V := O.GetValue(Name);
  if V is TJSONBool then
    Result := TJSONBool(V).AsBoolean
  else if V is TJSONTrue then
    Result := True
  else if V is TJSONFalse then
    Result := False;
end;

function JFloat(O: TJSONObject; const Name: string; const Def: Double = 0): Double;
var
  S: string;
  FS: TFormatSettings;
begin
  FS := TFormatSettings.Create;
  FS.DecimalSeparator := '.';
  S := StringReplace(JStr(O, Name, FloatToStr(Def, FS)), ',', '.', [rfReplaceAll]);
  Result := StrToFloatDef(S, Def, FS);
end;

/// <summary>СНО клиента → код taxationType АТОЛ.</summary>
function NormalizeAtolTaxationType(const S: string): string;
var
  T: string;
begin
  T := LowerCase(Trim(S));
  if T = 'osn' then
    Result := 'osn'
  else if T = 'usnincome' then
    Result := 'usnIncome'
  else if T = 'usnincomeoutcome' then
    Result := 'usnIncomeOutcome'
  else if T = 'esn' then
    Result := 'esn'
  else if T = 'patent' then
    Result := 'patent'
  else
    Result := Trim(S);
end;

/// <summary>Ставка НДС → tax.type АТОЛ (клиент уже шлёт коды АТОЛ).</summary>
function NormalizeAtolTaxType(const S: string): string;
var
  T: string;
begin
  T := LowerCase(Trim(S));
  if (T = '') or (T = 'без ндс') or (T = 'безндс') then
    Result := 'none'
  else if (T = '20%') or (T = '20') then
    Result := 'vat20'
  else if (T = '22%') or (T = '22') then
    Result := 'vat22'
  else if (T = '10%') or (T = '10') then
    Result := 'vat10'
  else if (T = '5%') or (T = '5') then
    Result := 'vat5'
  else if (T = '7%') or (T = '7') then
    Result := 'vat7'
  else if (T = '0%') or (T = '0') then
    Result := 'vat0'
  else
    Result := T;
end;

/// <summary>paymentObject клиента → строковый код АТОЛ (ФФД ≥ 1.2).</summary>
function MapAtolPaymentObject(const Obj: string; HasMark: Boolean): string;
var
  T: string;
begin
  T := LowerCase(Trim(Obj));
  if (T = 'commoditywithmarking') or (T = '33') then
    Exit('commodityWithMarking');
  if (T = 'commoditywithoutmarking') or (T = '32') then
    Exit('commodityWithoutMarking');
  if (T = 'excisewithmarking') or (T = '31') then
    Exit('exciseWithMarking');
  if (T = 'excisewithoutmarking') or (T = '30') then
    Exit('exciseWithoutMarking');
  if (T = 'commodity') or (T = '1') then
  begin
    if HasMark then
      Exit('commodityWithMarking');
    Exit('commodity');
  end;
  if (T = 'excise') or (T = '2') then
  begin
    if HasMark then
      Exit('exciseWithMarking');
    Exit('excise');
  end;
  if (T = 'job') or (T = '3') then
    Exit('job');
  if (T = 'service') or (T = '4') or (T = 'услуга') then
    Exit('service');
  if (T = 'payment') or (T = '10') or (T = 'платеж') then
    Exit('payment');
  if HasMark then
  begin
    if Pos('excise', T) > 0 then
      Exit('exciseWithMarking');
    Exit('commodityWithMarking');
  end;
  if Trim(Obj) <> '' then
    Result := Trim(Obj)
  else
    Result := 'commodity';
end;

function NormalizeAtolPaymentMethod(const S: string): string;
var
  T: string;
begin
  T := LowerCase(Trim(S));
  if (T = '') or (T = 'fullpayment') or (T = '4') then
    Result := 'fullPayment'
  else if (T = 'fullprepayment') or (T = '1') then
    Result := 'fullPrepayment'
  else if (T = 'prepayment') or (T = '2') then
    Result := 'prepayment'
  else if (T = 'advance') or (T = '3') then
    Result := 'advance'
  else if (T = 'partialpayment') or (T = '5') then
    Result := 'partialPayment'
  else if (T = 'credit') or (T = '6') then
    Result := 'credit'
  else if (T = 'creditpayment') or (T = '7') then
    Result := 'creditPayment'
  else
    Result := Trim(S);
end;

/// <summary>Единица измерения → код АТОЛ (строка «0» = шт, как в BSL).</summary>
function MapAtolMeasurementUnit(const U: string): string;
var
  T: string;
begin
  T := LowerCase(Trim(U));
  if (T = '') or (T = 'piece') or (T = 'шт') or (T = '0') then
    Result := '0'
  else if (T = 'liter') or (T = 'л') or (T = '41') then
    Result := '41'
  else
    Result := Trim(U);
end;

{-------------------------------------------------------------------------------
 TVariantPrint — конструктор/деструктор
-------------------------------------------------------------------------------}

constructor TVariantPrint.Create(ASettings: TCustomIniFile);
begin
  inherited Create;
  FSettings := ASettings;
  FLock := TCriticalSection.Create;
  FMarkTasksLock := TCriticalSection.Create;
  FMarkTasks := TObjectDictionary<string, TCheckMarkTaskEntry>.Create([doOwnsValues]);
  FMarkCache := TDictionary<string, TMarkCacheEntry>.Create;
  FLog := TKktLogger.Create(ASettings, TFormSettings.GetDefaultLogFileName);
  ovObject := Unassigned;
  FConnected := False;
  FFrStop := False;
  FFrThreadId := 0;
  FFrRequest := nil;
  FFrError := '';
  FFrWorkAvail := TEvent.Create(nil, False, False, '');
  FFrWorkDone := TEvent.Create(nil, False, False, '');
  FFrReady := TEvent.Create(nil, True, False, '');

  FEmulation := FSettings.ReadBool('KKT', 'Emulation', False);
  FTestReceiptMode := FSettings.ReadBool('KKT', 'TestReceiptMode', False);
  FEmulatedSerial := FSettings.ReadString('KKT', 'EmulatedSerial', '0000000000000001');
  FEmulatedInn := FSettings.ReadString('KKT', 'EmulatedInn', '1234567890');
  FEmulatedRegNumber := FSettings.ReadString('KKT', 'EmulatedRegNumber', '1234567111890');
  FEmulatedFnNumber := FSettings.ReadString('KKT', 'EmulatedFnNumber', '9999078900000001');
  FEmulatedCheckNumber := FSettings.ReadInteger('KKT', 'EmulatedCheckNumber', 0);
  FEmulatedSessionNumber := FSettings.ReadInteger('KKT', 'EmulatedSessionNumber', 0);

  StartFrComThread;

  if FEmulation then
    FLog.LogInfo('Режим эмуляции ККТ включён — ошибки драйвера логируются, ответ API = OK')
  else
    FLog.LogInfo('Боевой режим ККТ (АТОЛ Fptr10)');
  if FTestReceiptMode then
    FLog.LogInfo('Тестовый режим чека: вместо фискализации будет cancelReceipt');
end;

destructor TVariantPrint.Destroy;
begin
  FLock.Enter;
  try
    DoDisconnect;
  finally
    FLock.Leave;
  end;
  StopFrComThread;
  FFrWorkAvail.Free;
  FFrWorkDone.Free;
  FFrReady.Free;
  FMarkTasksLock.Enter;
  try
    FMarkTasks.Free;
  finally
    FMarkTasksLock.Leave;
    FMarkTasksLock.Free;
  end;
  FMarkCache.Free;
  FLog.Free;
  FLock.Free;
  inherited;
end;

procedure TVariantPrint.StartFrComThread;
begin
  FFrThread := TThread.CreateAnonymousThread(
    procedure
    begin
      CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
      try
        FFrThreadId := GetCurrentThreadId;
        FFrReady.SetEvent;
        while True do
        begin
          FFrWorkAvail.WaitFor(INFINITE);
          if FFrStop then
            Break;
          FFrError := '';
          try
            if Assigned(FFrRequest) then
              FFrRequest();
          except
            on E: Exception do
              FFrError := E.Message;
          end;
          FFrRequest := nil;
          FFrWorkDone.SetEvent;
        end;
        try
          if not (VarIsEmpty(ovObject) or VarIsNull(ovObject)) then
            ovObject := Unassigned;
        except
        end;
      finally
        CoUninitialize;
      end;
    end);
  FFrThread.FreeOnTerminate := False;
  FFrThread.Start;
  if FFrReady.WaitFor(15000) <> wrSignaled then
    raise Exception.Create('Не удалось запустить STA-поток драйвера ККТ');
end;

procedure TVariantPrint.StopFrComThread;
begin
  if FFrThread = nil then
    Exit;
  FFrStop := True;
  FFrWorkAvail.SetEvent;
  FFrThread.WaitFor;
  FreeAndNil(FFrThread);
  FFrThreadId := 0;
end;

procedure TVariantPrint.ExecOnFrThread(const Proc: TProc);
begin
  if not Assigned(Proc) then
    Exit;
  if GetCurrentThreadId = FFrThreadId then
  begin
    Proc();
    Exit;
  end;
  FFrError := '';
  FFrRequest := Proc;
  FFrWorkDone.ResetEvent;
  FFrWorkAvail.SetEvent;
  if FFrWorkDone.WaitFor(120000) <> wrSignaled then
    raise Exception.Create('Таймаут ожидания STA-потока драйвера ККТ');
  if FFrError <> '' then
    raise Exception.Create(FFrError);
end;

function TVariantPrint.ReadSerialNumber: string;
var
  Local, Resp, ErrDesc: string;
  ErrCode: Integer;
  V: TJSONValue;
  Root: TJSONObject;
begin
  Result := '';
  if FEmulation then
  begin
    Result := FEmulatedSerial;
    Exit;
  end;
  if VarIsEmpty(ovObject) or VarIsNull(ovObject) then
    Exit;

  Local := '';
  ExecOnFrThread(procedure
  begin
    try
      Local := VarToStr(ovObject.SerialNumber);
    except
      on E: Exception do
        FLog.LogDebug('[ATOL] ReadSerialNumber property: ' + E.Message);
    end;
  end);
  Result := Trim(Local);
  if Result <> '' then
    Exit;

  // АТОЛ Fptr10: серийный номер через getDeviceInfo
  if ProcessJsonAtol('{"type":"getDeviceInfo"}', 'getDeviceInfo', Resp, ErrCode, ErrDesc) then
  begin
    V := TJSONObject.ParseJSONValue(Resp);
    if V <> nil then
    try
      if V is TJSONObject then
      begin
        Root := TJSONObject(V);
        Result := Trim(JStr(Root, 'serial'));
        if Result = '' then
          Result := Trim(JStr(Root, 'serialNumber'));
      end;
    finally
      V.Free;
    end;
  end;
end;

function TVariantPrint.FRObjectReady: Boolean;
begin
  Result := not (VarIsEmpty(ovObject) or VarIsNull(ovObject));
end;

{-------------------------------------------------------------------------------
 Подключение / переиспользование соединения (АТОЛ Fptr10)
-------------------------------------------------------------------------------}

function TVariantPrint.AtolErrorCode: Integer;
begin
  try
    Result := Integer(ovObject.errorCode);
  except
    Result := -1;
  end;
end;

function TVariantPrint.AtolErrorDescription: string;
begin
  Result := '';
  try
    Result := VarToStr(ovObject.errorDescription);
  except
    Result := '';
  end;
end;

procedure TVariantPrint.ApplyAtolConnectionSettings;
var
  ConnType, IpAddr, RemoteAddr, SettingPort, SettingCom, SettingIp,
    SettingIpPort, SettingRemote, PortVal: string;
  ComNum, IpPortNum: Integer;
begin
  // Вызывать только из STA-потока при готовом ovObject
  ConnType := LowerCase(Trim(FSettings.ReadString('FRCash', 'ConnectionType',
    DefaultConnectionType)));
  ComNum := FSettings.ReadInteger('FRCash', 'ComNumber', DefaultComNumber);
  IpAddr := Trim(FSettings.ReadString('FRCash', 'IpAddress', ''));
  IpPortNum := FSettings.ReadInteger('FRCash', 'IpPort', DefaultIpPort);
  RemoteAddr := Trim(FSettings.ReadString('FRCash', 'RemoteServerAddr', ''));

  try SettingPort := VarToStr(ovObject.LIBFPTR_SETTING_PORT); except SettingPort := ATOL_SETTING_PORT; end;
  try SettingCom := VarToStr(ovObject.LIBFPTR_SETTING_COM_FILE); except SettingCom := ATOL_SETTING_COM_FILE; end;
  try SettingIp := VarToStr(ovObject.LIBFPTR_SETTING_IPADDRESS); except SettingIp := ATOL_SETTING_IPADDRESS; end;
  try SettingIpPort := VarToStr(ovObject.LIBFPTR_SETTING_IPPORT); except SettingIpPort := ATOL_SETTING_IPPORT; end;
  try SettingRemote := VarToStr(ovObject.LIBFPTR_SETTING_REMOTE_SERVER_ADDR); except SettingRemote := ATOL_SETTING_REMOTE_SERVER_ADDR; end;

  if ConnType = 'com' then
  begin
    try PortVal := VarToStr(ovObject.LIBFPTR_PORT_COM); except PortVal := IntToStr(ATOL_PORT_COM); end;
    ovObject.setSingleSetting(SettingPort, PortVal);
    ovObject.setSingleSetting(SettingCom, Format('COM%d', [ComNum]));
    FLog.LogInfo(Format('[ATOL] settings: COM%d', [ComNum]));
  end
  else if ConnType = 'ip' then
  begin
    try PortVal := VarToStr(ovObject.LIBFPTR_PORT_TCPIP); except PortVal := IntToStr(ATOL_PORT_TCPIP); end;
    ovObject.setSingleSetting(SettingPort, PortVal);
    ovObject.setSingleSetting(SettingIp, IpAddr);
    ovObject.setSingleSetting(SettingIpPort, IntToStr(IpPortNum));
    FLog.LogInfo(Format('[ATOL] settings: TCP %s:%d', [IpAddr, IpPortNum]));
  end
  else
  begin
    // usb по умолчанию
    try PortVal := VarToStr(ovObject.LIBFPTR_PORT_USB); except PortVal := IntToStr(ATOL_PORT_USB); end;
    ovObject.setSingleSetting(SettingPort, PortVal);
    FLog.LogInfo('[ATOL] settings: USB');
  end;

  if RemoteAddr <> '' then
  begin
    ovObject.setSingleSetting(SettingRemote, RemoteAddr);
    FLog.LogInfo('[ATOL] remote server: ' + RemoteAddr);
  end;

  ovObject.applySingleSettings;
end;

function TVariantPrint.ProcessJsonAtol(const JsonText, Stage: string;
  out ResponseJson: string; out ErrorCode: Integer;
  out ErrorDescription: string): Boolean;
var
  LocalResp: string;
  LocalCode, ExecCode: Integer;
  LocalDesc: string;
  ParamJson: Integer;
begin
  ResponseJson := '';
  ErrorCode := 0;
  ErrorDescription := '';

  EnsureFRConnected;

  FLog.LogDebug(Format('[ATOL] %s REQUEST: %s', [Stage, JsonText]));

  LocalResp := '';
  LocalCode := 0;
  LocalDesc := '';
  ExecCode := 0;

  ExecOnFrThread(procedure
  begin
    if not FRObjectReady then
    begin
      if FEmulation then
      begin
        LocalCode := 0;
        LocalDesc := '';
        LocalResp := '{}';
        Exit;
      end;
      raise Exception.Create(Stage + ': объект драйвера АТОЛ недоступен');
    end;

    try
      try
        ParamJson := Integer(ovObject.LIBFPTR_PARAM_JSON_DATA);
      except
        ParamJson := ATOL_PARAM_JSON_DATA;
      end;

      ovObject.setParam(ParamJson, JsonText);
      ExecCode := Integer(ovObject.processJson);
      LocalCode := AtolErrorCode;
      LocalDesc := AtolErrorDescription;
      try
        LocalResp := VarToStr(ovObject.getParamString(ParamJson));
      except
        LocalResp := '';
      end;
    except
      on E: Exception do
      begin
        if FEmulation then
        begin
          FLog.LogError(Format('[ATOL] %s ERROR (IGNORED): %s', [Stage, E.Message]));
          LocalCode := 0;
          LocalDesc := '';
          LocalResp := '{}';
          Exit;
        end;
        raise;
      end;
    end;
  end);

  ResponseJson := LocalResp;
  ErrorCode := LocalCode;
  ErrorDescription := LocalDesc;

  if FEmulation then
  begin
    if (ExecCode <> 0) or (LocalCode <> 0) then
      FLog.LogError(Format('[ATOL] %s ERROR (IGNORED): code=%d %s resp=%s',
        [Stage, LocalCode, LocalDesc, LocalResp]));
    // В эмуляции для бизнес-логики всегда отдаём мок ответа АТОЛ (как в BSL)
    ResponseJson := MockAtolJsonResponse(JsonText);
    ErrorCode := 0;
    ErrorDescription := '';
    Result := True;
    FLog.LogDebug(Format('[ATOL] %s RESPONSE (emu): %s', [Stage, ResponseJson]));
    Exit;
  end;

  if (ExecCode <> 0) or (LocalCode <> 0) then
  begin
    if ErrorDescription = '' then
      ErrorDescription := Format('processJson failed (exec=%d, errorCode=%d)',
        [ExecCode, LocalCode]);
    FLog.LogError(Format('[ATOL] %s: %s', [Stage, ErrorDescription]));
    Result := False;
    Exit;
  end;

  FLog.LogDebug(Format('[ATOL] %s RESPONSE: %s', [Stage, ResponseJson]));
  Result := True;
end;

function TVariantPrint.CheckLinkAlive: Boolean;
var
  Local: Boolean;
begin
  Result := False;
  if not FConnected then
    Exit;

  if FEmulation then
  begin
    Result := True;
    Exit;
  end;

  if VarIsEmpty(ovObject) or VarIsNull(ovObject) then
    Exit;

  Local := False;
  ExecOnFrThread(procedure
  begin
    try
      Local := Boolean(ovObject.isOpened);
    except
      on E: Exception do
      begin
        FLog.LogError('[ATOL] CheckLinkAlive: ' + E.Message);
        Local := False;
      end;
    end;
  end);
  Result := Local;
end;

function TVariantPrint.EnsureFRConnected: Boolean;
var
  OpenCode, ErrCode: Integer;
  ErrDesc, ConnType: string;
  OpenOk: Boolean;
begin
  if FConnected and CheckLinkAlive then
  begin
    Result := True;
    Exit;
  end;

  FConnected := False;
  OpenOk := False;
  OpenCode := -1;
  ErrCode := 0;
  ErrDesc := '';

  if VarIsEmpty(ovObject) or VarIsNull(ovObject) then
  begin
    try
      ExecOnFrThread(procedure
      begin
        ovObject := CreateOleObject(ATOL_PROG_ID);
      end);
      FLog.LogInfo('[ATOL] CreateOleObject ' + ATOL_PROG_ID);
    except
      on E: Exception do
      begin
        FLog.LogError('[ATOL] CreateOleObject: ' + E.Message);
        if FEmulation then
        begin
          FConnected := True;
          Result := True;
          Exit;
        end;
        raise Exception.Create('Драйвер АТОЛ v10 не установлен: ' + E.Message);
      end;
    end;
  end;

  ConnType := LowerCase(Trim(FSettings.ReadString('FRCash', 'ConnectionType',
    DefaultConnectionType)));
  FLog.LogInfo(Format('[ATOL] open ConnectionType=%s', [ConnType]));

  try
    ExecOnFrThread(procedure
    begin
      if not FRObjectReady then
        raise Exception.Create('объект драйвера АТОЛ недоступен');

      try
        if Boolean(ovObject.isOpened) then
        begin
          OpenOk := True;
          OpenCode := 0;
          Exit;
        end;
      except
        // isOpened недоступен — пробуем open
      end;

      ApplyAtolConnectionSettings;
      OpenCode := Integer(ovObject.open);
      ErrCode := AtolErrorCode;
      ErrDesc := AtolErrorDescription;
      OpenOk := (OpenCode = 0);
    end);
  except
    on E: Exception do
    begin
      FLog.LogError('[ATOL] open: ' + E.Message);
      if FEmulation then
      begin
        FConnected := True;
        Result := True;
        Exit;
      end;
      raise;
    end;
  end;

  if not OpenOk then
  begin
    if ErrDesc = '' then
      ErrDesc := Format('open failed (code=%d, errorCode=%d)', [OpenCode, ErrCode]);
    FLog.LogError('[ATOL] open: ' + ErrDesc);
    if FEmulation then
    begin
      FConnected := True;
      Result := True;
      Exit;
    end;
    raise Exception.Create(ErrDesc);
  end;

  FConnected := True;
  Result := True;
  FLog.LogInfo('[ATOL] open OK');
end;

procedure TVariantPrint.DoDisconnect;
begin
  if FRObjectReady then
  begin
    try
      ExecOnFrThread(procedure
      begin
        try
          if Boolean(ovObject.isOpened) then
            ovObject.close;
        except
          on E: Exception do
            FLog.LogError('[ATOL] close: ' + E.Message);
        end;
      end);
    except
      on E: Exception do
        FLog.LogError('[ATOL] DoDisconnect: ' + E.Message);
    end;
  end;
  FConnected := False;
  FLog.LogInfo('[ATOL] disconnected (COM-объект сохранён)');
end;

procedure TVariantPrint.AfterCommand(Req: TJSONObject);
begin
  if (Req <> nil) and JBool(Req, 'disconnect', False) then
    DoDisconnect;
end;

function TVariantPrint.DisconnectJson: string;
begin
  FLock.Enter;
  try
    try
      DoDisconnect;
      Result := JsonOk;
    except
      on E: Exception do
      begin
        FLog.LogError('DisconnectJson: ' + E.Message);
        if FEmulation then
          Result := JsonOk
        else
          Result := JsonError(E.Message);
      end;
    end;
  finally
    FLock.Leave;
  end;
end;

function TVariantPrint.ConnectJson: string;
begin
  FLock.Enter;
  try
    try
      EnsureFRConnected;
      Result := JsonOk('"connected":true');
      FLog.LogInfo('ConnectJson OK');
    except
      on E: Exception do
      begin
        FLog.LogError('ConnectJson: ' + E.Message);
        if FEmulation then
          Result := JsonOk('"connected":true')
        else
          Result := JsonError(E.Message);
      end;
    end;
  finally
    FLock.Leave;
  end;
end;

{-------------------------------------------------------------------------------
 /connection-status — статус связи и настройки FSettings (без принудительного open)
-------------------------------------------------------------------------------}

function TVariantPrint.KktConnectionStatusJson: string;
var
  ComNumber, IpPort, HttpPort: Integer;
  ConnType, IpAddress, RemoteAddr, Cashier: string;
  Connected: Boolean;
begin
  FLock.Enter;
  try
    ConnType := LowerCase(Trim(FSettings.ReadString('FRCash', 'ConnectionType',
      DefaultConnectionType)));
    ComNumber := FSettings.ReadInteger('FRCash', 'ComNumber', DefaultComNumber);
    IpAddress := FSettings.ReadString('FRCash', 'IpAddress', '');
    IpPort := FSettings.ReadInteger('FRCash', 'IpPort', DefaultIpPort);
    RemoteAddr := FSettings.ReadString('FRCash', 'RemoteServerAddr', '');
    HttpPort := FSettings.ReadInteger('HttpServer', 'Port', DefaultHttpPort);
    Cashier := FSettings.ReadString('FRCash', 'CashierName', DefaultCashierName);
    Connected := CheckLinkAlive;

    Result := JsonOk(
      '"connected":' + BoolToStr(Connected, True).ToLower + ',' +
      '"connectionType":"' + JsonEscape(ConnType) + '",' +
      '"comNumber":' + IntToStr(ComNumber) + ',' +
      '"ipAddress":"' + JsonEscape(IpAddress) + '",' +
      '"ipPort":' + IntToStr(IpPort) + ',' +
      '"remoteServerAddr":"' + JsonEscape(RemoteAddr) + '",' +
      '"cashierName":"' + JsonEscape(Cashier) + '",' +
      '"httpPort":' + IntToStr(HttpPort) + ',' +
      '"emulation":' + BoolToStr(FEmulation, True).ToLower + ',' +
      '"testReceiptMode":' + BoolToStr(FTestReceiptMode, True).ToLower);
  finally
    FLock.Leave;
  end;
end;

{-------------------------------------------------------------------------------
 /info — реквизиты ККТ через JSON АТОЛ (getDeviceInfo / getFnInfo / getRegistrationInfo)
-------------------------------------------------------------------------------}

function TVariantPrint.KktInfoJson: string;
var
  SerialNumber, KKTRegistrationNumber, Inn, FnNumber: string;
  Resp, ErrDesc: string;
  ErrCode: Integer;
  V: TJSONValue;
  Root, Org, Device: TJSONObject;

  function NestedStr(Parent: TJSONObject; const ChildName, FieldName: string): string;
  var
    Child: TJSONObject;
  begin
    Result := '';
    if Parent = nil then
      Exit;
    Result := Trim(JStr(Parent, FieldName));
    if Result <> '' then
      Exit;
    Child := Parent.GetValue(ChildName) as TJSONObject;
    if Child <> nil then
      Result := Trim(JStr(Child, FieldName));
  end;

begin
  FLock.Enter;
  try
    try
      EnsureFRConnected;

      SerialNumber := '';
      Inn := '';
      KKTRegistrationNumber := '';
      FnNumber := '';

      // Серийный номер ККТ
      SerialNumber := ReadSerialNumber;
      if (SerialNumber = '') and
        ProcessJsonAtol('{"type":"getDeviceInfo"}', 'getDeviceInfo', Resp, ErrCode, ErrDesc) then
      begin
        V := TJSONObject.ParseJSONValue(Resp);
        if V <> nil then
        try
          if V is TJSONObject then
          begin
            Root := TJSONObject(V);
            SerialNumber := Trim(JStr(Root, 'serial'));
            if SerialNumber = '' then
              SerialNumber := Trim(JStr(Root, 'serialNumber'));
            if SerialNumber = '' then
              SerialNumber := NestedStr(Root, 'deviceInfo', 'serial');
            if SerialNumber = '' then
              SerialNumber := NestedStr(Root, 'deviceInfo', 'serialNumber');
          end;
        finally
          V.Free;
        end;
      end;

      // ФН
      if ProcessJsonAtol('{"type":"getFnInfo"}', 'getFnInfo', Resp, ErrCode, ErrDesc) then
      begin
        V := TJSONObject.ParseJSONValue(Resp);
        if V <> nil then
        try
          if V is TJSONObject then
          begin
            Root := TJSONObject(V);
            FnNumber := Trim(JStr(Root, 'serial'));
            if FnNumber = '' then
              FnNumber := Trim(JStr(Root, 'fnNumber'));
            if FnNumber = '' then
              FnNumber := NestedStr(Root, 'fnInfo', 'serial');
            if FnNumber = '' then
              FnNumber := NestedStr(Root, 'fnInfo', 'fnNumber');
          end;
        finally
          V.Free;
        end;
      end;

      // ИНН / РН ККТ
      if ProcessJsonAtol('{"type":"getRegistrationInfo"}', 'getRegistrationInfo',
        Resp, ErrCode, ErrDesc) then
      begin
        V := TJSONObject.ParseJSONValue(Resp);
        if V <> nil then
        try
          if V is TJSONObject then
          begin
            Root := TJSONObject(V);
            Org := Root.GetValue('organization') as TJSONObject;
            Device := Root.GetValue('device') as TJSONObject;
            if Org <> nil then
            begin
              Inn := Trim(JStr(Org, 'vatin'));
              if Inn = '' then
                Inn := Trim(JStr(Org, 'inn'));
            end;
            if Inn = '' then
              Inn := Trim(JStr(Root, 'vatin'));
            if Inn = '' then
              Inn := Trim(JStr(Root, 'inn'));

            if Device <> nil then
            begin
              KKTRegistrationNumber := Trim(JStr(Device, 'registrationNumber'));
              if (FnNumber = '') then
                FnNumber := Trim(JStr(Device, 'fnNumber'));
            end;
            if KKTRegistrationNumber = '' then
              KKTRegistrationNumber := Trim(JStr(Root, 'registrationNumber'));
          end;
        finally
          V.Free;
        end;
      end;

      if FEmulation then
      begin
        if SerialNumber = '' then
          SerialNumber := FEmulatedSerial;
        if Inn = '' then
          Inn := FEmulatedInn;
        if KKTRegistrationNumber = '' then
          KKTRegistrationNumber := FEmulatedRegNumber;
        if FnNumber = '' then
          FnNumber := FEmulatedFnNumber;
      end;

      Result := JsonOk(
        '"serialNumber":"' + JsonEscape(SerialNumber) + '",' +
        '"inn":"'       + JsonEscape(Inn) + '",' +
        '"regNumber":"' + JsonEscape(KKTRegistrationNumber) + '",' +
        '"fnNumber":"'  + JsonEscape(FnNumber) + '",' +
        '"connected":'  + BoolToStr(FConnected, True).ToLower + ',' +
        '"emulation":'  + BoolToStr(FEmulation, True).ToLower);
    except
      on E: Exception do
      begin
        FLog.LogError('KktInfoJson: ' + E.Message);
        if FEmulation then
          Result := JsonOk(
            '"serialNumber":"' + JsonEscape(FEmulatedSerial) + '",' +
            '"inn":"' + JsonEscape(FEmulatedInn) + '",' +
            '"regNumber":"' + JsonEscape(FEmulatedRegNumber) + '",' +
            '"fnNumber":"' + JsonEscape(FEmulatedFnNumber) + '",' +
            '"connected":true,' +
            '"emulation":true')
        else
          Result := JsonError(E.Message);
      end;
    end;
  finally
    FLock.Leave;
  end;
end;

function TVariantPrint.MarkCacheKey(const MarkDecoded: string; ReturnCheck: Boolean): string;
begin
  if ReturnCheck then
    Result := 'R|' + MarkDecoded
  else
    Result := 'S|' + MarkDecoded;
end;

function TVariantPrint.TryGetMarkFromCache(const MarkDecoded: string;
  ReturnCheck: Boolean; out Entry: TMarkCacheEntry): Boolean;
begin
  Result := False;
  Entry := Default(TMarkCacheEntry);
  if MarkDecoded = '' then
    Exit;
  Result := FMarkCache.TryGetValue(MarkCacheKey(MarkDecoded, ReturnCheck), Entry);
end;

function TVariantPrint.IsMarkFullyOk(const MarkDecoded: string; ReturnCheck,
  CheckPermission: Boolean): Boolean;
var
  Entry: TMarkCacheEntry;
begin
  Result := False;
  if not TryGetMarkFromCache(MarkDecoded, ReturnCheck, Entry) then
    Exit;
  if not Entry.KktAccepted then
    Exit;
  if CheckPermission and not ReturnCheck then
  begin
    if not Entry.RrChecked then
      Exit;
    if not Entry.RrAccepted then
      Exit;
    if (Trim(Entry.Uuid) = '') or (Trim(Entry.TimeStr) = '') then
      Exit;
  end;
  Result := True;
end;

procedure TVariantPrint.UpsertMarkCacheKkt(const MarkDecoded: string;
  ReturnCheck: Boolean; KktAccepted: Boolean; KktResultCode: Integer;
  const KktDescription: string; const KktWarning: string;
  const KktValidationJson: string);
var
  Key: string;
  Entry: TMarkCacheEntry;
begin
  if MarkDecoded = '' then
    Exit;
  Key := MarkCacheKey(MarkDecoded, ReturnCheck);
  if not FMarkCache.TryGetValue(Key, Entry) then
  begin
    Entry := Default(TMarkCacheEntry);
    Entry.RrAccepted := True; // РР ещё не проверялся
  end;
  Entry.KktAccepted := KktAccepted;
  Entry.KktResultCode := KktResultCode;
  Entry.KktDescription := KktDescription;
  Entry.KktWarning := KktWarning;
  if KktValidationJson <> '' then
    Entry.KktValidationJson := KktValidationJson;
  Entry.CheckedAt := Now;
  FMarkCache.AddOrSetValue(Key, Entry);
  FLog.LogInfo(Format('Кэш марки ККТ: accepted=%s code=%d warn=%s len=%d return=%s',
    [BoolToStr(KktAccepted, True), KktResultCode, KktWarning, Length(MarkDecoded),
     BoolToStr(ReturnCheck, True)]));
end;

procedure TVariantPrint.UpsertMarkCacheRr(const MarkDecoded: string;
  ReturnCheck: Boolean; RrAccepted: Boolean; RrCode: Integer;
  const RrDescription, Uuid, TimeStr, Inst, Version: string);
var
  Key: string;
  Entry: TMarkCacheEntry;
begin
  if MarkDecoded = '' then
    Exit;
  Key := MarkCacheKey(MarkDecoded, ReturnCheck);
  if not FMarkCache.TryGetValue(Key, Entry) then
    Entry := Default(TMarkCacheEntry);
  Entry.RrChecked := True;
  Entry.RrAccepted := RrAccepted;
  Entry.RrCode := RrCode;
  Entry.RrDescription := RrDescription;
  Entry.Uuid := Uuid;
  Entry.TimeStr := TimeStr;
  Entry.Inst := Inst;
  Entry.Version := Version;
  Entry.CheckedAt := Now;
  FMarkCache.AddOrSetValue(Key, Entry);
  FLog.LogInfo(Format('Кэш марки РР: accepted=%s code=%d uuid=%s len=%d',
    [BoolToStr(RrAccepted, True), RrCode, Uuid, Length(MarkDecoded)]));
end;

procedure TVariantPrint.UpdateMarkCacheFromMarkObj(const MarkDecoded: string;
  ReturnCheck: Boolean; MarkObj: TJSONObject);
var
  PermObj: TJSONObject;
  CodeStr, Desc: string;
begin
  if (MarkDecoded = '') or (MarkObj = nil) then
    Exit;

  Desc := JStr(MarkObj, 'description', '');
  if (Length(Desc) >= 9) and (Copy(Desc, Length(Desc) - 8, 9) = ' (cached)') then
    Desc := Copy(Desc, 1, Length(Desc) - 9);

  UpsertMarkCacheKkt(MarkDecoded, ReturnCheck,
    JBool(MarkObj, 'accepted', False),
    JInt(MarkObj, 'resultCode', 0),
    Desc,
    JStr(MarkObj, 'warning', ''));

  PermObj := MarkObj.GetValue('permission') as TJSONObject;
  if PermObj = nil then
    Exit;

  CodeStr := JStr(PermObj, 'code', '0');
  UpsertMarkCacheRr(MarkDecoded, ReturnCheck,
    JBool(PermObj, 'accepted', False),
    StrToIntDef(CodeStr, 0),
    JStr(PermObj, 'error', ''),
    JStr(PermObj, 'uuid', ''),
    JStr(PermObj, 'time', ''),
    JStr(PermObj, 'inst', ''),
    JStr(PermObj, 'version', ''));
end;

procedure TVariantPrint.ClearMarkCache;
begin
  if FMarkCache.Count > 0 then
  begin
    FLog.LogInfo('Сброс кэша марок (' + IntToStr(FMarkCache.Count) + ')');
    FMarkCache.Clear;
  end;
end;

function TVariantPrint.MarkObjFromCacheEntry(const Entry: TMarkCacheEntry): TJSONObject;
var
  Desc: string;
begin
  Result := TJSONObject.Create;
  Result.AddPair('accepted', TJSONBool.Create(Entry.KktAccepted));
  Result.AddPair('resultCode', TJSONNumber.Create(Entry.KktResultCode));
  Result.AddPair('checkItemLocalResult', TJSONNumber.Create(15));
  Result.AddPair('checkItemLocalError', TJSONNumber.Create(0));
  Result.AddPair('markingType2', TJSONNumber.Create(0));
  Result.AddPair('kmServerErrorCode', TJSONNumber.Create(0));
  Result.AddPair('kmServerCheckingStatus', TJSONNumber.Create(15));
  Desc := Entry.KktDescription;
  if Desc = '' then
    Desc := 'OK';
  if Pos('(cached)', Desc) = 0 then
    Desc := Desc + ' (cached)';
  Result.AddPair('description', Desc);
  if Entry.KktWarning <> '' then
    Result.AddPair('warning', Entry.KktWarning);
end;

function TVariantPrint.PermissionObjFromCacheEntry(const Entry: TMarkCacheEntry): TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('accepted', TJSONBool.Create(Entry.RrAccepted));
  Result.AddPair('code', IntToStr(Entry.RrCode));
  Result.AddPair('error', Entry.RrDescription);
  Result.AddPair('uuid', Entry.Uuid);
  Result.AddPair('time', Entry.TimeStr);
  Result.AddPair('inst', Entry.Inst);
  Result.AddPair('version', Entry.Version);
end;

function TVariantPrint.GetReceiptReturnCheck(Req: TJSONObject): Boolean;
var
  T: string;
begin
  T := LowerCase(Trim(JStr(Req, 'type')));
  Result := T = 'sellreturn';
end;

function TVariantPrint.GetCashierName(Req: TJSONObject): string;
var
  Op: TJSONObject;
begin
  Result := '';
  if Req <> nil then
  begin
    Result := Trim(JStr(Req, 'cashierName'));
    if Result = '' then
    begin
      Op := Req.GetValue('operator') as TJSONObject;
      if Op <> nil then
        Result := Trim(JStr(Op, 'name'));
    end;
  end;
  if Result = '' then
    Result := FSettings.ReadString('FRCash', 'CashierName', DefaultCashierName);
end;

function TVariantPrint.GetCashierInn(Req: TJSONObject): string;
var
  Op: TJSONObject;
begin
  Result := '';
  if Req <> nil then
  begin
    Result := Trim(JStr(Req, 'cashierInn'));
    if Result = '' then
    begin
      Op := Req.GetValue('operator') as TJSONObject;
      if Op <> nil then
        Result := Trim(JStr(Op, 'vatin'));
    end;
  end;
  if Result = '' then
    Result := Trim(FSettings.ReadString('FRCash', 'CashierInn', ''));
end;

function TVariantPrint.BuildAtolJsonGetShiftStatus: string;
begin
  Result := '{"type":"getShiftStatus"}';
end;

function TVariantPrint.BuildAtolJsonContinuePrint: string;
begin
  Result := '{"type":"continuePrint"}';
end;

function TVariantPrint.BuildAtolJsonBeginMarkValidation(const MarkB64,
  EstimatedStatus: string): string;
var
  Job, Params: TJSONObject;
begin
  Job := TJSONObject.Create;
  try
    Job.AddPair('type', 'beginMarkingCodeValidation');
    Params := TJSONObject.Create;
    Params.AddPair('imcType', 'auto');
    Params.AddPair('imc', MarkB64);
    Params.AddPair('itemEstimatedStatus', EstimatedStatus);
    Params.AddPair('imcModeProcessing', TJSONNumber.Create(0));
    Job.AddPair('params', Params);
    Result := JsonToPlainText(Job);
  finally
    Job.Free;
  end;
end;

function TVariantPrint.BuildAtolJsonGetMarkValidationStatus: string;
begin
  Result := '{"type":"getMarkingCodeValidationStatus"}';
end;

function TVariantPrint.BuildAtolJsonAcceptMarkingCode: string;
begin
  Result := '{"type":"acceptMarkingCode"}';
end;

function TVariantPrint.BuildAtolJsonCancelMarkValidation: string;
begin
  Result := '{"type":"cancelMarkingCodeValidation"}';
end;

function TVariantPrint.BuildAtolJsonClearMarkValidationResult: string;
begin
  // Как СформироватьJSONОчисткиРезультатаПроверкиМарки в BSL
  Result := '{"type":"clearMarkingCodeValidationResult"}';
end;

function TVariantPrint.GetEstimatedMarkStatus(ReturnCheck, DraftBeer,
  DraftBeerLeftovers: Boolean): string;
begin
  // Как ПолучитьПланируемыйСтатусМарки в BSL
  if ReturnCheck then
  begin
    if DraftBeer then
    begin
      if DraftBeerLeftovers then
        Result := 'itemStatusUnchanged'
      else
        Result := 'itemDryReturn';
    end
    else
      Result := 'itemPieceReturn';
  end
  else if DraftBeer then
  begin
    if DraftBeerLeftovers then
      Result := 'itemStatusUnchanged'
    else
      Result := 'itemDryForSale';
  end
  else
    Result := 'itemPieceSold';
end;

function TVariantPrint.GetMarkValidationTimeoutSec: Integer;
begin
  Result := FSettings.ReadInteger('FRCash', 'MarkValidationTimeout',
    ATOL_MARK_DEFAULT_TIMEOUT_SEC);
  if Result < 1 then
    Result := ATOL_MARK_DEFAULT_TIMEOUT_SEC;
end;

function TVariantPrint.IsTspiotRetryCode(Code: Integer): Boolean;
begin
  Result := (Code = 403) or (Code = 407) or (Code = 410);
end;

procedure TVariantPrint.CancelMarkValidationQuiet;
var
  Resp, ErrDesc: string;
  ErrCode: Integer;
begin
  try
    ProcessJsonAtol(BuildAtolJsonCancelMarkValidation, 'cancelMarkingCodeValidation',
      Resp, ErrCode, ErrDesc);
  except
    on E: Exception do
      FLog.LogInfo('cancelMarkingCodeValidation: ' + E.Message);
  end;
end;

function TVariantPrint.BuildAtolJsonOpenShift(const CashierName,
  CashierInn: string): string;
var
  Job, Op: TJSONObject;
begin
  Job := TJSONObject.Create;
  try
    Job.AddPair('type', 'openShift');
    Op := TJSONObject.Create;
    Op.AddPair('name', CashierName);
    if CashierInn <> '' then
      Op.AddPair('vatin', CashierInn);
    Job.AddPair('operator', Op);
    Result := JsonToPlainText(Job);
  finally
    Job.Free;
  end;
end;

function TVariantPrint.BuildAtolJsonCloseShift(const CashierName,
  CashierInn: string): string;
var
  Job, Op: TJSONObject;
begin
  Job := TJSONObject.Create;
  try
    Job.AddPair('type', 'closeShift');
    Op := TJSONObject.Create;
    Op.AddPair('name', CashierName);
    if CashierInn <> '' then
      Op.AddPair('vatin', CashierInn);
    Job.AddPair('operator', Op);
    Result := JsonToPlainText(Job);
  finally
    Job.Free;
  end;
end;

function TVariantPrint.BuildAtolJsonReportX(const CashierName,
  CashierInn: string): string;
var
  Job, Op: TJSONObject;
begin
  Job := TJSONObject.Create;
  try
    Job.AddPair('type', 'reportX');
    if CashierName <> '' then
    begin
      Op := TJSONObject.Create;
      Op.AddPair('name', CashierName);
      if CashierInn <> '' then
        Op.AddPair('vatin', CashierInn);
      Job.AddPair('operator', Op);
    end;
    Result := JsonToPlainText(Job);
  finally
    Job.Free;
  end;
end;

function TVariantPrint.MockAtolJsonResponse(const JsonRequest: string): string;
var
  V: TJSONValue;
  O: TJSONObject;
  TaskType: string;
  ShiftNo: Integer;
begin
  TaskType := '';
  V := TJSONObject.ParseJSONValue(JsonRequest);
  try
    if V is TJSONObject then
    begin
      O := TJSONObject(V);
      TaskType := LowerCase(Trim(JStr(O, 'type')));
    end;
  finally
    V.Free;
  end;

  ShiftNo := FEmulatedSessionNumber;
  if ShiftNo <= 0 then
    ShiftNo := 1;

  if TaskType = 'getshiftstatus' then
    Result := Format(
      '{"shiftStatus":{"state":"closed","number":%d,"documentsCount":0}}',
      [ShiftNo])
  else if TaskType = 'getdeviceinfo' then
    Result := Format(
      '{"deviceInfo":{"serial":"%s","modelName":"ATOL Emulation","firmwareVersion":"0.0"},' +
      '"serial":"%s"}',
      [JsonEscape(FEmulatedSerial), JsonEscape(FEmulatedSerial)])
  else if TaskType = 'getfninfo' then
    Result := Format(
      '{"fnInfo":{"serial":"%s","version":"fn_emu"},"serial":"%s"}',
      [JsonEscape(FEmulatedFnNumber), JsonEscape(FEmulatedFnNumber)])
  else if TaskType = 'getregistrationinfo' then
    Result := Format(
      '{"organization":{"vatin":"%s","name":"ATOL Emulation"},' +
      '"device":{"registrationNumber":"%s","fnNumber":"%s"},' +
      '"registrationNumber":"%s"}',
      [JsonEscape(FEmulatedInn), JsonEscape(FEmulatedRegNumber),
       JsonEscape(FEmulatedFnNumber), JsonEscape(FEmulatedRegNumber)])
  else if (TaskType = 'openshift') or (TaskType = 'closeshift') then
    Result := Format(
      '{"fiscalParams":{"shiftNumber":%d,"fiscalDocumentNumber":1,' +
      '"fiscalDocumentDateTime":"2026-08-03T12:00:00+05:00",' +
      '"fnNumber":"%s","registrationNumber":"%s"}}',
      [ShiftNo, JsonEscape(FEmulatedFnNumber), JsonEscape(FEmulatedRegNumber)])
  else if TaskType = 'reportx' then
    Result := '{}'
  else if TaskType = 'continueprint' then
    Result := '{}'
  else if TaskType = 'beginmarkingcodevalidation' then
    Result :=
      '{"offlineValidation":{"fmCheck":false,"fmCheckResult":false,' +
      '"fmCheckErrorReason":"noKeys"}}'
  else if TaskType = 'getmarkingcodevalidationstatus' then
    Result :=
      '{"ready":true,"sentImcRequest":true,' +
      '"driverError":{"code":0,"error":"","description":""},' +
      '"onlineValidation":{' +
      '"itemInfoCheckResult":{"imcCheckFlag":true,"imcCheckResult":true,' +
      '"imcStatusInfo":true,"ecrStandAloneFlag":true},' +
      '"markOperatorItemStatus":"itemEstimatedStatusCorrect",' +
      '"markOperatorResponse":{"responseStatus":true,"itemStatusCheck":true},' +
      '"markOperatorResponseResult":"correct","imcType":"auto","imcModeProcessing":0}}'
  else if TaskType = 'acceptmarkingcode' then
    Result :=
      '{"itemInfoCheckResult":{"ecrStandAloneFlag":true,"imcCheckFlag":true,' +
      '"imcCheckResult":true,"imcEstimatedStatusCorrect":true,"imcStatusInfo":true}}'
  else if (TaskType = 'cancelmarkingcodevalidation') or
    (TaskType = 'declinemarkingcode') or
    (TaskType = 'clearmarkingcodevalidationresult') then
    Result := '{}'
  else if (TaskType = 'sell') or (TaskType = 'sellreturn') then
  begin
    Inc(FEmulatedCheckNumber);
    Result := Format(
      '{"fiscalParams":{"fiscalDocumentNumber":%d,"fiscalReceiptNumber":%d,' +
      '"shiftNumber":%d,"fiscalDocumentDateTime":"2026-08-03T12:00:00+05:00",' +
      '"fnNumber":"%s","registrationNumber":"%s","receiptsCount":%d}}',
      [FEmulatedCheckNumber, FEmulatedCheckNumber, ShiftNo,
       JsonEscape(FEmulatedFnNumber), JsonEscape(FEmulatedRegNumber),
       FEmulatedCheckNumber]);
  end
  else if TaskType = 'nonfiscal' then
    Result := '{}'
  else
    Result := '{}';
end;

function TVariantPrint.ExtractShiftNumberFromAtolResponse(
  const ResponseJson: string): Integer;
var
  V: TJSONValue;
  Root, FP, SS: TJSONObject;
begin
  Result := 0;
  if Trim(ResponseJson) = '' then
    Exit;
  V := TJSONObject.ParseJSONValue(ResponseJson);
  if V = nil then
    Exit;
  try
    if not (V is TJSONObject) then
      Exit;
    Root := TJSONObject(V);
    FP := Root.GetValue('fiscalParams') as TJSONObject;
    if FP <> nil then
    begin
      Result := JInt(FP, 'shiftNumber', 0);
      if Result <> 0 then
        Exit;
    end;
    SS := Root.GetValue('shiftStatus') as TJSONObject;
    if SS <> nil then
      Result := JInt(SS, 'number', 0);
  finally
    V.Free;
  end;
end;

function TVariantPrint.QueryShiftStatus(out Opened, Expired: Boolean;
  out ShiftNumber: Integer): Boolean;
var
  Resp, ErrDesc, State: string;
  ErrCode: Integer;
  V: TJSONValue;
  Root, SS: TJSONObject;
begin
  Opened := False;
  Expired := False;
  ShiftNumber := 0;

  if not ProcessJsonAtol(BuildAtolJsonGetShiftStatus, 'getShiftStatus', Resp,
    ErrCode, ErrDesc) then
    raise Exception.Create(ErrDesc);

  V := TJSONObject.ParseJSONValue(Resp);
  if V = nil then
    raise Exception.Create('Не удалось разобрать JSON ответа статуса смены');
  try
    if not (V is TJSONObject) then
      raise Exception.Create('Не удалось разобрать JSON ответа статуса смены');
    Root := TJSONObject(V);
    SS := Root.GetValue('shiftStatus') as TJSONObject;
    if SS = nil then
      raise Exception.Create('Не удалось разобрать JSON ответа статуса смены');

    State := LowerCase(Trim(JStr(SS, 'state')));
    ShiftNumber := JInt(SS, 'number', 0);
    if State = 'opened' then
      Opened := True
    else if State = 'expired' then
    begin
      Opened := True;
      Expired := True;
    end;
    // closed — Opened=False
    Result := True;
    FLog.LogInfo(Format('getShiftStatus: state=%s number=%d',
      [State, ShiftNumber]));
  finally
    V.Free;
  end;
end;

function TVariantPrint.CancelReceiptAtol(const Stage: string;
  RaiseOnError: Boolean): Boolean;
var
  LocalCode: Integer;
  LocalDesc: string;
begin
  EnsureFRConnected;

  LocalCode := 0;
  LocalDesc := '';

  ExecOnFrThread(procedure
  begin
    if not FRObjectReady then
    begin
      if FEmulation then
      begin
        LocalCode := 0;
        LocalDesc := '';
        Exit;
      end;
      raise Exception.Create(Stage + ': объект драйвера АТОЛ недоступен');
    end;

    try
      ovObject.cancelReceipt;
      LocalCode := AtolErrorCode;
      LocalDesc := AtolErrorDescription;
    except
      on E: Exception do
      begin
        if FEmulation then
        begin
          FLog.LogError(Format('[ATOL] %s ERROR (IGNORED): %s',
            [Stage, E.Message]));
          LocalCode := 0;
          LocalDesc := '';
          Exit;
        end;
        raise;
      end;
    end;
  end);

  if FEmulation then
  begin
    if LocalCode <> 0 then
      FLog.LogError(Format('[ATOL] %s ERROR (IGNORED): code=%d %s',
        [Stage, LocalCode, LocalDesc]));
    Result := True;
    Exit;
  end;

  if LocalCode <> 0 then
  begin
    if LocalDesc = '' then
      LocalDesc := Format('cancelReceipt failed (errorCode=%d)', [LocalCode]);
    FLog.LogError(Format('[ATOL] %s: %s', [Stage, LocalDesc]));
    if RaiseOnError then
      raise Exception.Create(LocalDesc);
    Result := False;
    Exit;
  end;

  FLog.LogInfo('[ATOL] ' + Stage + ': cancelReceipt OK');
  Result := True;
end;

{-------------------------------------------------------------------------------
 Сборка JSON чека АТОЛ (sell / sellReturn) — как СформироватьJSONЧека в BSL
-------------------------------------------------------------------------------}

function TVariantPrint.DefaultItemInfoCheckResult(Accepted: Boolean): TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('imcCheckFlag', TJSONBool.Create(Accepted));
  Result.AddPair('imcCheckResult', TJSONBool.Create(Accepted));
  Result.AddPair('imcStatusInfo', TJSONBool.Create(Accepted));
  Result.AddPair('imcEstimatedStatusCorrect', TJSONBool.Create(Accepted));
  Result.AddPair('ecrStandAloneFlag', TJSONBool.Create(Accepted));
end;

function TVariantPrint.ExtractItemInfoCheckResult(
  const ValidationJson: string): TJSONObject;
var
  V: TJSONValue;
  Root, Online, Info: TJSONObject;
  Clone: TJSONValue;
begin
  Result := nil;
  if Trim(ValidationJson) = '' then
    Exit;
  V := TJSONObject.ParseJSONValue(ValidationJson);
  if V = nil then
    Exit;
  try
    if not (V is TJSONObject) then
      Exit;
    Root := TJSONObject(V);
    Online := Root.GetValue('onlineValidation') as TJSONObject;
    if Online = nil then
      Exit;
    Info := Online.GetValue('itemInfoCheckResult') as TJSONObject;
    if Info = nil then
      Exit;
    Clone := TJSONObject.ParseJSONValue(Info.ToJSON);
    if Clone is TJSONObject then
      Result := TJSONObject(Clone)
    else
      Clone.Free;
  finally
    V.Free;
  end;
end;

function TVariantPrint.ExtractCheckNumberFromAtolResponse(
  const ResponseJson: string): Integer;
var
  V: TJSONValue;
  Root, FP: TJSONObject;
begin
  Result := 0;
  if Trim(ResponseJson) = '' then
    Exit;
  V := TJSONObject.ParseJSONValue(ResponseJson);
  if V = nil then
    Exit;
  try
    if not (V is TJSONObject) then
      Exit;
    Root := TJSONObject(V);
    FP := Root.GetValue('fiscalParams') as TJSONObject;
    if FP = nil then
      Exit;
    Result := JInt(FP, 'fiscalReceiptNumber', 0);
    if Result = 0 then
      Result := JInt(FP, 'fiscalDocumentNumber', 0);
  finally
    V.Free;
  end;
end;

function TVariantPrint.BuildAtolIndustryInfo(Item: TJSONObject;
  const MarkDecoded: string; ReturnCheck: Boolean): TJSONArray;
var
  PermObj, Elem: TJSONObject;
  UUID, TimeStr, Inst, Version, Tag1262, Tag1263, Tag1264, Attr: string;
  Cached: TMarkCacheEntry;
begin
  Result := nil;
  if Item = nil then
    Exit;

  UUID := '';
  TimeStr := '';
  Inst := '';
  Version := '';
  Tag1262 := '030';
  Tag1263 := '21.11.2023';
  Tag1264 := '1944';

  PermObj := Item.GetValue('permission') as TJSONObject;
  if PermObj <> nil then
  begin
    UUID := JStr(PermObj, 'uuid');
    TimeStr := JStr(PermObj, 'time');
    Inst := JStr(PermObj, 'inst');
    Version := JStr(PermObj, 'version');
    Tag1262 := JStr(PermObj, 'tag1262', Tag1262);
    Tag1263 := JStr(PermObj, 'tag1263', Tag1263);
    Tag1264 := JStr(PermObj, 'tag1264', Tag1264);
  end;

  if (not ReturnCheck) and ((UUID = '') or (TimeStr = '')) and (MarkDecoded <> '') then
  begin
    if TryGetMarkFromCache(MarkDecoded, ReturnCheck, Cached) and Cached.RrChecked and
      (Cached.Uuid <> '') and (Cached.TimeStr <> '') then
    begin
      UUID := Cached.Uuid;
      TimeStr := Cached.TimeStr;
      Inst := Cached.Inst;
      Version := Cached.Version;
      FLog.LogInfo('industryInfo: permission из кэша uuid=' + UUID);
    end;
  end;

  if (UUID = '') or (TimeStr = '') then
    Exit;

  Attr := 'UUID=' + UUID + '&Time=' + TimeStr;
  if Inst <> '' then
    Attr := Attr + '&Inst=' + Inst + '&Ver=' + Version;

  Elem := TJSONObject.Create;
  Elem.AddPair('fois', Tag1262);
  Elem.AddPair('date', Tag1263);
  Elem.AddPair('number', Tag1264);
  Elem.AddPair('industryAttribute', Attr);

  Result := TJSONArray.Create;
  Result.AddElement(Elem);
end;

function TVariantPrint.BuildAtolImcParams(Item: TJSONObject;
  const MarkB64: string; ReturnCheck: Boolean): TJSONObject;
var
  MarkDecoded, EstimatedStatus, ValidationJson: string;
  Cached: TMarkCacheEntry;
  Info: TJSONObject;
  DraftBeer, DraftBeerLeftovers: Boolean;
begin
  Result := TJSONObject.Create;
  Result.AddPair('imcType', 'auto');
  Result.AddPair('imc', MarkB64);
  Result.AddPair('imcModeProcessing', TJSONNumber.Create(0));

  MarkDecoded := DecodeMarkBase64(MarkB64);
  DraftBeer := JBool(Item, 'draftBeer', False) or JBool(Item, 'isDraftBeer', False);
  DraftBeerLeftovers := JBool(Item, 'draftBeerLeftovers', False);
  EstimatedStatus := GetEstimatedMarkStatus(ReturnCheck, DraftBeer, DraftBeerLeftovers);
  Result.AddPair('itemEstimatedStatus', EstimatedStatus);

  ValidationJson := '';
  if TryGetMarkFromCache(MarkDecoded, ReturnCheck, Cached) then
    ValidationJson := Cached.KktValidationJson;

  Info := ExtractItemInfoCheckResult(ValidationJson);
  if Info = nil then
  begin
    // После acceptMarkingCode флаги обычно true; иначе [M] / без проверки
    if TryGetMarkFromCache(MarkDecoded, ReturnCheck, Cached) and Cached.KktAccepted then
      Info := DefaultItemInfoCheckResult(True)
    else if JBool(Item, 'markAccepted', False) then
      Info := DefaultItemInfoCheckResult(True)
    else
      Info := DefaultItemInfoCheckResult(False);
  end;
  Result.AddPair('itemInfoCheckResult', Info);
end;

function TVariantPrint.BuildAtolReceiptItem(Item: TJSONObject;
  ReturnCheck: Boolean; DefaultDepartment: Integer): TJSONObject;
var
  ItemType, Name, MarkB64, MarkDecoded, TaxTypeStr: string;
  TaxObj, TaxOut: TJSONObject;
  Industry: TJSONArray;
  Qty, Price, Amount: Double;
  Department: Integer;
  HasMark: Boolean;
begin
  Result := nil;
  if Item = nil then
    Exit;

  ItemType := LowerCase(Trim(JStr(Item, 'type', 'position')));
  if ItemType = 'text' then
  begin
    Result := TJSONObject.Create;
    Result.AddPair('type', 'text');
    Result.AddPair('text', JStr(Item, 'text'));
    Exit;
  end;

  if (ItemType <> 'position') and (ItemType <> '') and (JStr(Item, 'name') = '') then
    Exit;

  Name := JStr(Item, 'name');
  Qty := JFloat(Item, 'quantity', 1);
  Price := JFloat(Item, 'price', 0);
  Amount := JFloat(Item, 'amount', 0);
  if Amount <= 0 then
    Amount := Price * Qty;
  Department := JInt(Item, 'department', DefaultDepartment);

  MarkB64 := Trim(JStr(Item, 'mark'));
  HasMark := MarkB64 <> '';
  if HasMark then
    MarkDecoded := DecodeMarkBase64(MarkB64)
  else
    MarkDecoded := '';

  Result := TJSONObject.Create;
  Result.AddPair('type', 'position');
  Result.AddPair('name', Name);
  Result.AddPair('price', TJSONNumber.Create(Price));
  Result.AddPair('quantity', TJSONNumber.Create(Qty));
  Result.AddPair('amount', TJSONNumber.Create(Amount));
  Result.AddPair('measurementUnit',
    MapAtolMeasurementUnit(JStr(Item, 'measurementUnit')));
  Result.AddPair('paymentMethod',
    NormalizeAtolPaymentMethod(JStr(Item, 'paymentMethod', 'fullPayment')));
  Result.AddPair('paymentObject',
    MapAtolPaymentObject(JStr(Item, 'paymentObject', 'commodity'), HasMark));
  if Department > 0 then
    Result.AddPair('department', TJSONNumber.Create(Department));

  TaxObj := Item.GetValue('tax') as TJSONObject;
  if TaxObj <> nil then
    TaxTypeStr := NormalizeAtolTaxType(JStr(TaxObj, 'type'))
  else
    TaxTypeStr := 'none';
  TaxOut := TJSONObject.Create;
  TaxOut.AddPair('type', TaxTypeStr);
  Result.AddPair('tax', TaxOut);

  if HasMark then
  begin
    if not ReturnCheck then
    begin
      Industry := BuildAtolIndustryInfo(Item, MarkDecoded, ReturnCheck);
      if Industry <> nil then
        Result.AddPair('industryInfo', Industry)
      else
        FLog.LogInfo('WARN: маркированная позиция без permission — industryInfo не добавлен');
    end;
    Result.AddPair('imcParams', BuildAtolImcParams(Item, MarkB64, ReturnCheck));
  end;
end;

function TVariantPrint.BuildAtolPayments(Req: TJSONObject): TJSONArray;
var
  Payments: TJSONArray;
  Pay, OutPay: TJSONObject;
  I, Payment: Integer;
  PayType: string;
  Sum, Total, Received, Summ1, Summ2, SummCredit: Double;
begin
  Result := TJSONArray.Create;
  if Req = nil then
    Exit;

  Payments := Req.GetValue('payments') as TJSONArray;
  if (Payments <> nil) and (Payments.Count > 0) then
  begin
    for I := 0 to Payments.Count - 1 do
    begin
      if not (Payments.Items[I] is TJSONObject) then
        Continue;
      Pay := TJSONObject(Payments.Items[I]);
      PayType := LowerCase(Trim(JStr(Pay, 'type')));
      Sum := JFloat(Pay, 'sum', 0);
      if Sum <= 0 then
        Continue;
      if PayType = 'cash' then
        PayType := 'cash'
      else if PayType = 'electronically' then
        PayType := 'electronically'
      else if PayType = 'credit' then
        PayType := 'credit'
      else
      begin
        FLog.LogInfo('WARN: неизвестный тип оплаты "' + PayType + '" — пропуск');
        Continue;
      end;
      OutPay := TJSONObject.Create;
      OutPay.AddPair('type', PayType);
      OutPay.AddPair('sum', TJSONNumber.Create(Sum));
      Result.AddElement(OutPay);
    end;
    Exit;
  end;

  // Legacy: payment + total / received
  Total := JFloat(Req, 'total', 0);
  Received := JFloat(Req, 'received', Total);
  Summ1 := 0;
  Summ2 := 0;
  SummCredit := 0;
  Payment := JInt(Req, 'payment', 0);
  case Payment of
    0: Summ1 := Total;
    1: Summ2 := Total;
    2: SummCredit := Received;
  else
    Summ1 := Total;
  end;
  if Summ1 > 0 then
  begin
    OutPay := TJSONObject.Create;
    OutPay.AddPair('type', 'cash');
    OutPay.AddPair('sum', TJSONNumber.Create(Summ1));
    Result.AddElement(OutPay);
  end;
  if Summ2 > 0 then
  begin
    OutPay := TJSONObject.Create;
    OutPay.AddPair('type', 'electronically');
    OutPay.AddPair('sum', TJSONNumber.Create(Summ2));
    Result.AddElement(OutPay);
  end;
  if SummCredit > 0 then
  begin
    OutPay := TJSONObject.Create;
    OutPay.AddPair('type', 'credit');
    OutPay.AddPair('sum', TJSONNumber.Create(SummCredit));
    Result.AddElement(OutPay);
  end;
end;

function TVariantPrint.BuildAtolReceiptJson(Req: TJSONObject): string;
var
  Job, Op, Item, AtolItem: TJSONObject;
  ItemsIn, ItemsOut, PaymentsOut: TJSONArray;
  I, DefaultDept: Integer;
  ItemType, CashierName, CashierInn, TaxType: string;
  Total: Double;
begin
  Job := TJSONObject.Create;
  try
    if GetReceiptReturnCheck(Req) then
      Job.AddPair('type', 'sellReturn')
    else
      Job.AddPair('type', 'sell');

    TaxType := NormalizeAtolTaxationType(JStr(Req, 'taxationType'));
    if TaxType <> '' then
      Job.AddPair('taxationType', TaxType);

    CashierName := GetCashierName(Req);
    CashierInn := GetCashierInn(Req);
    if CashierName <> '' then
    begin
      Op := TJSONObject.Create;
      Op.AddPair('name', CashierName);
      if CashierInn <> '' then
        Op.AddPair('vatin', CashierInn);
      Job.AddPair('operator', Op);
    end;

    DefaultDept := JInt(Req, 'department', 1);
    ItemsOut := TJSONArray.Create;
    ItemsIn := Req.GetValue('items') as TJSONArray;
    if ItemsIn <> nil then
    begin
      for I := 0 to ItemsIn.Count - 1 do
      begin
        if not (ItemsIn.Items[I] is TJSONObject) then
          Continue;
        Item := TJSONObject(ItemsIn.Items[I]);
        ItemType := LowerCase(Trim(JStr(Item, 'type', 'position')));
        if (ItemType <> 'text') and (ItemType <> 'position') and (ItemType <> '') and
          (JStr(Item, 'name') = '') then
        begin
          FLog.LogInfo('Пропуск элемента items type="' + ItemType + '"');
          Continue;
        end;
        AtolItem := BuildAtolReceiptItem(Item, GetReceiptReturnCheck(Req), DefaultDept);
        if AtolItem <> nil then
          ItemsOut.AddElement(AtolItem);
      end;
    end;
    Job.AddPair('items', ItemsOut);

    PaymentsOut := BuildAtolPayments(Req);
    Job.AddPair('payments', PaymentsOut);

    Total := JFloat(Req, 'total', 0);
    if Total > 0 then
      Job.AddPair('total', TJSONNumber.Create(Total));

    Result := JsonToPlainText(Job);
  finally
    Job.Free;
  end;
end;

function TVariantPrint.BuildAtolNonFiscalItem(Item: TJSONObject): TJSONObject;
var
  ItemType, Align, BarcodeType, TextVal, BarcodeVal: string;
  Scale: Integer;
begin
  Result := nil;
  if Item = nil then
    Exit;

  ItemType := LowerCase(Trim(JStr(Item, 'type', 'text')));
  if ItemType = 'text' then
  begin
    TextVal := JStr(Item, 'text');
    // Пустую строку пропускаем (как в сборке слипа 1С)
    if TextVal = '' then
      Exit;
    Result := TJSONObject.Create;
    Result.AddPair('type', 'text');
    Result.AddPair('text', TextVal);
    Align := LowerCase(Trim(JStr(Item, 'alignment')));
    if (Align = 'left') or (Align = 'center') or (Align = 'right') then
      Result.AddPair('alignment', Align);
    if Item.GetValue('font') <> nil then
      Result.AddPair('font', TJSONNumber.Create(JInt(Item, 'font', 0)));
    if Item.GetValue('doubleWidth') <> nil then
      Result.AddPair('doubleWidth', TJSONBool.Create(JBool(Item, 'doubleWidth', False)));
    if Item.GetValue('doubleHeight') <> nil then
      Result.AddPair('doubleHeight', TJSONBool.Create(JBool(Item, 'doubleHeight', False)));
  end
  else if ItemType = 'barcode' then
  begin
    BarcodeVal := JStr(Item, 'barcode');
    if BarcodeVal = '' then
      Exit;
    Result := TJSONObject.Create;
    Result.AddPair('type', 'barcode');
    Result.AddPair('barcode', BarcodeVal);
    BarcodeType := Trim(JStr(Item, 'barcodeType', 'QR'));
    if BarcodeType = '' then
      BarcodeType := 'QR';
    Result.AddPair('barcodeType', BarcodeType);
    Scale := JInt(Item, 'scale', 0);
    if Scale > 0 then
      Result.AddPair('scale', TJSONNumber.Create(Scale));
    Align := LowerCase(Trim(JStr(Item, 'alignment')));
    if (Align = 'left') or (Align = 'center') or (Align = 'right') then
      Result.AddPair('alignment', Align);
  end
  else
    FLog.LogInfo('nonFiscal: пропуск items type="' + ItemType + '"');
end;

function TVariantPrint.BuildAtolNonFiscalJson(Req: TJSONObject): string;
var
  Job, Item, AtolItem: TJSONObject;
  ItemsIn, ItemsOut: TJSONArray;
  I: Integer;
begin
  Job := TJSONObject.Create;
  try
    Job.AddPair('type', 'nonFiscal');

    ItemsOut := TJSONArray.Create;
    ItemsIn := Req.GetValue('items') as TJSONArray;
    if ItemsIn <> nil then
    begin
      for I := 0 to ItemsIn.Count - 1 do
      begin
        if not (ItemsIn.Items[I] is TJSONObject) then
          Continue;
        Item := TJSONObject(ItemsIn.Items[I]);
        AtolItem := BuildAtolNonFiscalItem(Item);
        if AtolItem <> nil then
          ItemsOut.AddElement(AtolItem);
      end;
    end;
    if ItemsOut.Count = 0 then
    begin
      ItemsOut.Free;
      raise Exception.Create('nonFiscal: items пуст — нечего печатать');
    end;
    Job.AddPair('items', ItemsOut);

    if Req.GetValue('printFooter') <> nil then
      Job.AddPair('printFooter', TJSONBool.Create(JBool(Req, 'printFooter', True)))
    else
      Job.AddPair('printFooter', TJSONBool.Create(True));

    Result := JsonToPlainText(Job);
  finally
    Job.Free;
  end;
end;

function TVariantPrint.PrintNonFiscalFromRequest(Req: TJSONObject): string;
var
  AtolJson, Resp, ErrDesc: string;
  ErrCode: Integer;
begin
  EnsureFRConnected;

  AtolJson := BuildAtolNonFiscalJson(Req);
  FLog.LogInfo('PrintNonFiscal ATOL JSON: ' + JsonBodyPreviewPlain(AtolJson, 800));

  if not ProcessJsonAtol(AtolJson, 'nonFiscal', Resp, ErrCode, ErrDesc) then
    raise Exception.Create(ErrDesc);

  Result := JsonOk('"nonFiscal":true');
  FLog.LogInfo('PrintNonFiscal OK');
  // Кэш марок не сбрасываем — слип не завершает фискальный чек
  AfterCommand(Req);
end;

function TVariantPrint.PrintCheckJson(const Body: string): string;
var
  Req: TJSONObject;
  Items: TJSONArray;
  CheckNumber: Integer;
  CashierName, PreflightErr, AtolJson, Resp, ErrDesc, SerialNumber, Desc: string;
  ErrCode: Integer;
  IsNonFiscal: Boolean;
begin
  FLock.Enter;
  try
    Result := JsonError('Unknown error');
    Req := nil;
    IsNonFiscal := False;
    try
      FLog.LogInfo('PrintCheckJson begin');
      EnsureFRConnected;

      Req := ParseJsonObjectOrRaise(Body, 'PrintCheckJson', FLog);
      IsNonFiscal := SameText(Trim(JStr(Req, 'type')), 'nonFiscal');

      // Слип / нефискальный документ АТОЛ (как ПечатьСлипаЧерезВнешнююПрограмму в 1С)
      if IsNonFiscal then
      begin
        Result := PrintNonFiscalFromRequest(Req);
      end
      else
      begin
      CashierName := GetCashierName(Req);
      EnsureSessionOpen(CashierName, GetCashierInn(Req));

      Items := Req.GetValue('items') as TJSONArray;
      if Items = nil then
        raise Exception.Create('items is empty');

      PreflightErr := EnsureMarksReadyForPrint(Req);
      if PreflightErr <> '' then
      begin
        Result := PreflightErr;
        FLog.LogError('PrintCheckJson: pre-flight — марки не прошли, чек не открываем');
        // кэш не сбрасываем — результаты проверки уже известны
      end
      else
      begin
        AtolJson := BuildAtolReceiptJson(Req);
        FLog.LogInfo('PrintCheckJson ATOL JSON: ' +
          JsonBodyPreviewPlain(AtolJson, 800));

        if FTestReceiptMode and (not FEmulation) then
        begin
          // Как «НеПробиватьЧек» в BSL: марки проверены, processJson(sell) не вызываем
          FLog.LogInfo('Тестовый режим: processJson(sell) пропущен, cancelReceipt');
          CancelReceiptAtol('testReceiptMode cancelReceipt', True);
          SerialNumber := ReadSerialNumber;
          Desc := 'Проверки пройдены. Чек не пробит (тестовый режим). Выполнена отмена чека.';
          Result :=
            '{"result":1,"description":"' + JsonEscape(Desc) + '",' +
            '"checkNumber":0,' +
            '"serialNumber":"' + JsonEscape(SerialNumber) + '",' +
            '"testReceipt":true}';
          FLog.LogInfo('PrintCheckJson testReceiptMode OK');
          ClearMarkCache;
          AfterCommand(Req);
        end
        else
        begin
          if not ProcessJsonAtol(AtolJson, 'sell/sellReturn', Resp, ErrCode, ErrDesc) then
            raise Exception.Create(ErrDesc);

          CheckNumber := ExtractCheckNumberFromAtolResponse(Resp);
          if (CheckNumber = 0) and FEmulation then
            CheckNumber := FEmulatedCheckNumber;

          SerialNumber := ReadSerialNumber;

          Result := JsonOk(
            '"checkNumber":' + IntToStr(CheckNumber) + ',' +
            '"serialNumber":"' + JsonEscape(SerialNumber) + '"');

          FLog.LogInfo(Format('PrintCheckJson OK checkNumber=%d serial=%s',
            [CheckNumber, SerialNumber]));
          ClearMarkCache;
          AfterCommand(Req);
        end;
      end;
      end; // not nonFiscal
    except
      on E: Exception do
      begin
        FLog.LogError('PrintCheckJson: ' + E.Message);
        if IsNonFiscal then
        begin
          if FEmulation then
          begin
            Result := JsonOk('"nonFiscal":true');
            AfterCommand(Req);
          end
          else
            Result := JsonError(E.Message);
        end
        else if FEmulation then
        begin
          try
            CancelReceiptAtol('CancelReceipt after error', False);
          except
          end;
          ClearMarkCache;
          Inc(FEmulatedCheckNumber);
          Result := JsonOk(
            '"checkNumber":' + IntToStr(FEmulatedCheckNumber) + ',' +
            '"serialNumber":"' + JsonEscape(FEmulatedSerial) + '"');
          AfterCommand(Req);
        end
        else
        begin
          try
            CancelReceiptAtol('CancelReceipt after error', False);
          except
          end;
          ClearMarkCache;
          Result := JsonError(E.Message);
        end;
      end;
    end;
    Req.Free;
  finally
    FLock.Leave;
  end;
end;

{-------------------------------------------------------------------------------
 Служебные команды
-------------------------------------------------------------------------------}

function TVariantPrint.CloseShiftJson(const Body: string): string;
var
  V: TJSONValue;
  Req: TJSONObject;
  SessionNumber: Integer;
  CashierName, CashierInn, Json, Resp, ErrDesc: string;
  ErrCode: Integer;
begin
  FLock.Enter;
  try
    V := nil;
    Req := nil;
    try
      FLog.LogInfo('CloseShiftJson begin');
      V := TJSONObject.ParseJSONValue(Body);
      if V is TJSONObject then
        Req := TJSONObject(V);

      EnsureFRConnected;
      EnsureReadyForZReport;

      CashierName := GetCashierName(Req);
      CashierInn := GetCashierInn(Req);
      Json := BuildAtolJsonCloseShift(CashierName, CashierInn);
      if not ProcessJsonAtol(Json, 'closeShift', Resp, ErrCode, ErrDesc) then
        raise Exception.Create(ErrDesc);

      SessionNumber := ExtractShiftNumberFromAtolResponse(Resp);
      if FEmulation then
      begin
        Inc(FEmulatedSessionNumber);
        SessionNumber := FEmulatedSessionNumber;
      end;

      Result := JsonOk('"documentNumber":' + IntToStr(SessionNumber));
      FLog.LogInfo('CloseShiftJson OK documentNumber=' + IntToStr(SessionNumber));
      ClearMarkCache;
      AfterCommand(Req);
    except
      on E: Exception do
      begin
        FLog.LogError('CloseShiftJson: ' + E.Message);
        if FEmulation then
        begin
          Inc(FEmulatedSessionNumber);
          Result := JsonOk('"documentNumber":' + IntToStr(FEmulatedSessionNumber));
          ClearMarkCache;
          AfterCommand(Req);
        end
        else
          Result := JsonError(E.Message);
      end;
    end;
    V.Free;
  finally
    FLock.Leave;
  end;
end;

function TVariantPrint.XReportJson(const Body: string = ''): string;
var
  V: TJSONValue;
  Req: TJSONObject;
  CashierName, CashierInn, Json, Resp, ErrDesc: string;
  ErrCode: Integer;
begin
  FLock.Enter;
  try
    V := nil;
    Req := nil;
    try
      FLog.LogInfo('XReportJson begin');
      V := TJSONObject.ParseJSONValue(Body);
      if V is TJSONObject then
        Req := TJSONObject(V);

      EnsureFRConnected;
      EnsureReadyForXReport;

      CashierName := GetCashierName(Req);
      CashierInn := GetCashierInn(Req);
      Json := BuildAtolJsonReportX(CashierName, CashierInn);
      if not ProcessJsonAtol(Json, 'reportX', Resp, ErrCode, ErrDesc) then
        raise Exception.Create(ErrDesc);

      Result := JsonOk;
      FLog.LogInfo('XReportJson OK');
      AfterCommand(Req);
    except
      on E: Exception do
      begin
        FLog.LogError('XReportJson: ' + E.Message);
        if FEmulation then
        begin
          Result := JsonOk;
          AfterCommand(Req);
        end
        else
          Result := JsonError(E.Message);
      end;
    end;
    V.Free;
  finally
    FLock.Leave;
  end;
end;

function TVariantPrint.OpenShiftJson(const Body: string = ''): string;
var
  V: TJSONValue;
  Req: TJSONObject;
  CashierName, CashierInn: string;
  SessionNumber: Integer;
  Opened, Expired: Boolean;
begin
  FLock.Enter;
  try
    V := nil;
    Req := nil;
    try
      FLog.LogInfo('OpenShiftJson begin');
      V := TJSONObject.ParseJSONValue(Body);
      if V is TJSONObject then
        Req := TJSONObject(V);

      EnsureFRConnected;

      ClearStuckOpenDocument('OpenShift');
      QueryShiftStatus(Opened, Expired, SessionNumber);

      if Expired then
        raise Exception.Create(
          'Смена открыта более 24 часов. Снимите Z-отчёт (закрытие смены) и повторите.');

      if Opened then
      begin
        FLog.LogInfo('OpenShiftJson: смена уже открыта');
        Result := JsonOk('"documentNumber":' + IntToStr(SessionNumber) +
          ',"alreadyOpen":true');
        AfterCommand(Req);
        Exit;
      end;

      CashierName := GetCashierName(Req);
      CashierInn := GetCashierInn(Req);
      SessionNumber := DoOpenSession(CashierName, CashierInn);

      Result := JsonOk('"documentNumber":' + IntToStr(SessionNumber));
      FLog.LogInfo('OpenShiftJson OK documentNumber=' + IntToStr(SessionNumber));
      AfterCommand(Req);
    except
      on E: Exception do
      begin
        FLog.LogError('OpenShiftJson: ' + E.Message);
        if FEmulation then
        begin
          Inc(FEmulatedSessionNumber);
          Result := JsonOk('"documentNumber":' + IntToStr(FEmulatedSessionNumber));
          AfterCommand(Req);
        end
        else
          Result := JsonError(E.Message);
      end;
    end;
    V.Free;
  finally
    FLock.Leave;
  end;
end;

function TVariantPrint.CancelCheckJson(const Body: string = ''): string;
var
  V: TJSONValue;
  Req: TJSONObject;
begin
  FLock.Enter;
  try
    V := nil;
    Req := nil;
    try
      FLog.LogInfo('CancelCheckJson begin');
      V := TJSONObject.ParseJSONValue(Body);
      if V is TJSONObject then
        Req := TJSONObject(V);

      EnsureFRConnected;
      // Нет открытого чека — не ошибка: cancel-check ещё и сбрасывает FMarkCache
      if not CancelReceiptAtol('cancelReceipt', False) then
        FLog.LogInfo('CancelCheckJson: cancelReceipt без документа/с ошибкой — кэш сбрасываем');
      ClearMarkCache;
      Result := JsonOk;
      FLog.LogInfo('CancelCheckJson OK');
      AfterCommand(Req);
    except
      on E: Exception do
      begin
        FLog.LogError('CancelCheckJson: ' + E.Message);
        // Кэш сбрасываем даже при сбое связи — иначе «залипшие» марки блокируют чек
        ClearMarkCache;
        if FEmulation then
        begin
          Result := JsonOk;
          AfterCommand(Req);
        end
        else
          Result := JsonError(E.Message);
      end;
    end;
    V.Free;
  finally
    FLock.Leave;
  end;
end;

function TVariantPrint.ClearBufferOfMarksJson(const Body: string = ''): string;
var
  V: TJSONValue;
  Req: TJSONObject;
  Resp, ErrDesc: string;
  ErrCode: Integer;
begin
  FLock.Enter;
  try
    V := nil;
    Req := nil;
    try
      FLog.LogInfo('ClearBufferOfMarksJson begin');
      V := TJSONObject.ParseJSONValue(Body);
      if V is TJSONObject then
        Req := TJSONObject(V);

      EnsureFRConnected;
      if not ProcessJsonAtol(BuildAtolJsonClearMarkValidationResult,
        'clearMarkingCodeValidationResult', Resp, ErrCode, ErrDesc) then
        raise Exception.Create(ErrDesc);

      ClearMarkCache;
      Result := JsonOk;
      FLog.LogInfo('ClearBufferOfMarksJson OK');
      AfterCommand(Req);
    except
      on E: Exception do
      begin
        FLog.LogError('ClearBufferOfMarksJson: ' + E.Message);
        if FEmulation then
        begin
          Result := JsonOk;
          ClearMarkCache;
          AfterCommand(Req);
        end
        else
          Result := JsonError(E.Message);
      end;
    end;
    V.Free;
  finally
    FLock.Leave;
  end;
end;

function TVariantPrint.WaitMarkValidationReady(CanRetryTspiot: Boolean;
  out ResponseJson: string; out ErrorCode: Integer;
  out ErrorDescription: string; out NeedRestart, TimedOut: Boolean): Boolean;
var
  I, MaxPolls: Integer;
  Resp, ErrDesc: string;
  ErrCode: Integer;
  StatusObj, DriverErr: TJSONObject;
  V: TJSONValue;
  Ready: Boolean;
begin
  Result := False;
  ResponseJson := '';
  ErrorCode := 0;
  ErrorDescription := '';
  NeedRestart := False;
  TimedOut := False;

  MaxPolls := GetMarkValidationTimeoutSec;
  if MaxPolls < 1 then
    MaxPolls := ATOL_MARK_DEFAULT_TIMEOUT_SEC;

  for I := 1 to MaxPolls do
  begin
    if not ProcessJsonAtol(BuildAtolJsonGetMarkValidationStatus,
      'getMarkingCodeValidationStatus', Resp, ErrCode, ErrDesc) then
    begin
      if IsTspiotRetryCode(ErrCode) and CanRetryTspiot then
      begin
        FLog.LogInfo(Format(
          'getMarkingCodeValidationStatus: ТС ПИоТ %d — повтор begin', [ErrCode]));
        NeedRestart := True;
        ErrorCode := ErrCode;
        ErrorDescription := ErrDesc;
        Exit;
      end;
      ErrorCode := ErrCode;
      ErrorDescription := ErrDesc;
      if ErrorDescription = '' then
        ErrorDescription := 'Ошибка опроса статуса проверки марки';
      Exit;
    end;

    Ready := False;
    V := TJSONObject.ParseJSONValue(Resp);
    try
      if V is TJSONObject then
      begin
        StatusObj := TJSONObject(V);
        Ready := JBool(StatusObj, 'ready', False);
        DriverErr := StatusObj.GetValue('driverError') as TJSONObject;
        if (not Ready) and (DriverErr <> nil) then
        begin
          ErrCode := JInt(DriverErr, 'code', 0);
          if IsTspiotRetryCode(ErrCode) and CanRetryTspiot then
          begin
            FLog.LogInfo(Format(
              'getMarkingCodeValidationStatus: driverError ТС ПИоТ %d — повтор begin',
              [ErrCode]));
            NeedRestart := True;
            ErrorCode := ErrCode;
            ErrorDescription := JStr(DriverErr, 'description', ErrDesc);
            Exit;
          end;
        end;
      end;
    finally
      V.Free;
    end;

    if Ready then
    begin
      ResponseJson := Resp;
      Result := True;
      Exit;
    end;

    FLog.LogDebug(Format('getMarkingCodeValidationStatus: не ready, опрос %d/%d',
      [I, MaxPolls]));
    if I < MaxPolls then
      Sleep(ATOL_MARK_POLL_INTERVAL_MS);
  end;

  TimedOut := True;
  ErrorDescription := 'Превышено максимальное число опросов статуса проверки марки';
end;

procedure TVariantPrint.ClassifyMarkValidationResponse(const ResponseJson: string;
  out Accepted: Boolean; out Description, Warning: string; out ResultCode: Integer);
var
  V: TJSONValue;
  Root, DriverErr, Online, Info: TJSONObject;
  DriverCode: Integer;
  DriverDesc, OpResult, ItemStatus: string;
  SentRequest, ImcCheckFlag, ImcCheckResult, ImcStatusInfo: Boolean;
begin
  // Как РазобратьОтветПроверкиМаркиНаККТ + Предупреждение в BSL
  Accepted := False;
  Description := '';
  Warning := '';
  ResultCode := 0;

  V := TJSONObject.ParseJSONValue(ResponseJson);
  if not (V is TJSONObject) then
  begin
    if V <> nil then
      V.Free;
    Description := 'Некорректный JSON ответа проверки марки';
    ResultCode := -1;
    Exit;
  end;

  Root := TJSONObject(V);
  try
    if not JBool(Root, 'ready', False) then
    begin
      Description := 'Результат проверки ещё не готов';
      Exit;
    end;

    DriverErr := Root.GetValue('driverError') as TJSONObject;
    DriverCode := 0;
    DriverDesc := '';
    if DriverErr <> nil then
    begin
      DriverCode := JInt(DriverErr, 'code', 0);
      DriverDesc := JStr(DriverErr, 'description', '');
    end;

    // 421 / 402 — принять с предупреждением
    if (DriverCode = 421) or (DriverCode = 402) then
    begin
      Accepted := True;
      ResultCode := 0;
      Description := 'Код маркировки успешно проверен';
      if DriverDesc <> '' then
        Warning := DriverDesc
      else
        Warning := 'Ошибка драйвера ' + IntToStr(DriverCode);
      Exit;
    end;

    if DriverCode <> 0 then
    begin
      Accepted := False;
      ResultCode := DriverCode;
      if DriverDesc <> '' then
        Description := DriverDesc
      else
        Description := 'Ошибка проверки кода маркировки (' + IntToStr(DriverCode) + ')';
      Exit;
    end;

    Online := Root.GetValue('onlineValidation') as TJSONObject;
    if Online = nil then
    begin
      Description := 'В ответе нет onlineValidation';
      ResultCode := -1;
      Exit;
    end;

    OpResult := LowerCase(Trim(JStr(Online, 'markOperatorResponseResult', '')));
    if OpResult <> 'correct' then
    begin
      if OpResult = 'incorrect' then
        Description := 'Запрос имеет некорректный формат'
      else
        Description :=
          'Указанный в запросе код маркировки имеет некорректный формат (не распознан)';
      Exit;
    end;

    Info := Online.GetValue('itemInfoCheckResult') as TJSONObject;
    ImcCheckFlag := False;
    ImcCheckResult := False;
    ImcStatusInfo := False;
    if Info <> nil then
    begin
      ImcCheckFlag := JBool(Info, 'imcCheckFlag', False);
      ImcCheckResult := JBool(Info, 'imcCheckResult', False);
      ImcStatusInfo := JBool(Info, 'imcStatusInfo', False);
    end;

    if ImcCheckFlag and (not ImcCheckResult) then
    begin
      Description := 'Код маркировки не прошёл проверку';
      Exit;
    end;

    SentRequest := JBool(Root, 'sentImcRequest', False);
    if not ImcStatusInfo then
    begin
      Accepted := True;
      ResultCode := 0;
      Description := 'Код маркировки успешно проверен';
      if SentRequest then
        Warning := 'Запрос отправлен, но ответ ОИСМ не получен'
      else
        Warning := 'Не удалось отправить запрос в ОИСМ';
      Exit;
    end;

    ItemStatus := Trim(JStr(Online, 'markOperatorItemStatus', ''));
    if SameText(ItemStatus, 'itemEstimatedStatusCorrect') then
    begin
      Accepted := True;
      ResultCode := 0;
      Description := 'Код маркировки успешно проверен';
      Exit;
    end;

    if SameText(ItemStatus, 'itemSaleStopped') or
      SameText(ItemStatus, 'itemStatusCheckSuspended') then
    begin
      Description := 'Оборот товара приостановлен';
      Exit;
    end;

    Description := 'Планируемый статус товара некорректен';
  finally
    Root.Free;
  end;
end;

{-------------------------------------------------------------------------------
 Проверка маркировки (АТОЛ: begin → poll → accept/cancel)
-------------------------------------------------------------------------------}

function TVariantPrint.DecodeMarkBase64(const B64: string): string;
var
  Bytes: TBytes;
  I, Len, RawLen: Integer;
begin
  if Trim(B64) = '' then
    raise Exception.Create('Пустая марка (base64)');

  try
    Bytes := TNetEncoding.Base64.DecodeStringToBytes(B64);
  except
    on E: Exception do
      raise Exception.Create('Некорректный base64 марки: ' + E.Message);
  end;

  RawLen := Length(Bytes);
  if RawLen = 0 then
    raise Exception.Create('Декодированная марка пуста');

  // 1С часто отдаёт КМ из фиксированного поля со пробелами/нулями справа —
  // обрезаем хвост для стабильного ключа кэша и логов.
  Len := RawLen;
  while (Len > 0) and ((Bytes[Len - 1] = Ord(' ')) or (Bytes[Len - 1] = 0)) do
    Dec(Len);
  if Len = 0 then
    raise Exception.Create('Декодированная марка пуста после обрезки хвоста');
  if Len <> RawLen then
    FLog.LogInfo(Format('DecodeMarkBase64: обрезан хвост %d→%d', [RawLen, Len]));

  SetLength(Result, Len);
  for I := 0 to Len - 1 do
    Result[I + 1] := Char(Bytes[I]);
end;

function TVariantPrint.NewTaskId: string;
var
  G: TGUID;
begin
  CreateGUID(G);
  Result := GUIDToString(G);
  Result := StringReplace(Result, '{', '', [rfReplaceAll]);
  Result := StringReplace(Result, '}', '', [rfReplaceAll]);
end;

procedure TVariantPrint.TrimOldMarkTasks;
var
  Keys: TArray<string>;
  Key: string;
  Entry: TCheckMarkTaskEntry;
  I, RemoveCount: Integer;
begin
  if FMarkTasks.Count <= 50 then
    Exit;

  Keys := FMarkTasks.Keys.ToArray;
  RemoveCount := FMarkTasks.Count - 50;
  for I := 0 to High(Keys) do
  begin
    if RemoveCount <= 0 then
      Break;
    Key := Keys[I];
    if FMarkTasks.TryGetValue(Key, Entry) then
    begin
      if (Entry.Status <> 'pending') or (MinutesBetween(Now, Entry.CreatedAt) > 60) then
      begin
        FMarkTasks.Remove(Key);
        Dec(RemoveCount);
      end;
    end;
  end;
end;

procedure TVariantPrint.CollectMarksFromRequest(Req: TJSONObject;
  out MarksB64: TArray<string>);
var
  Arr: TJSONArray;
  I: Integer;
  SingleMark: string;
begin
  SetLength(MarksB64, 0);
  if Req = nil then
    raise Exception.Create('Invalid request');

  Arr := Req.GetValue('marks') as TJSONArray;
  if Arr <> nil then
  begin
    if Arr.Count = 0 then
      raise Exception.Create('marks is empty');
    SetLength(MarksB64, Arr.Count);
    for I := 0 to Arr.Count - 1 do
      MarksB64[I] := Arr.Items[I].Value;
    Exit;
  end;

  SingleMark := JStr(Req, 'mark');
  if SingleMark = '' then
    raise Exception.Create('Укажите mark (base64) или marks (массив base64)');

  SetLength(MarksB64, 1);
  MarksB64[0] := SingleMark;
end;

procedure TVariantPrint.CollectMarksFromReceiptItems(Req: TJSONObject;
  out MarksB64: TArray<string>);
var
  Items: TJSONArray;
  Item: TJSONObject;
  I: Integer;
  ItemType, MarkB64: string;
begin
  SetLength(MarksB64, 0);
  if Req = nil then
    Exit;
  Items := Req.GetValue('items') as TJSONArray;
  if Items = nil then
    Exit;

  for I := 0 to Items.Count - 1 do
  begin
    if not (Items.Items[I] is TJSONObject) then
      Continue;
    Item := TJSONObject(Items.Items[I]);
    ItemType := LowerCase(Trim(JStr(Item, 'type', 'position')));
    if (ItemType <> 'position') and (ItemType <> '') then
      Continue;
    MarkB64 := Trim(JStr(Item, 'mark'));
    if MarkB64 = '' then
      Continue;
    SetLength(MarksB64, Length(MarksB64) + 1);
    MarksB64[High(MarksB64)] := MarkB64;
  end;
end;

function TVariantPrint.EnsureMarksReadyForPrint(Req: TJSONObject): string;
var
  MarksB64: TArray<string>;
  MarksOut: TJSONArray;
  Root, MarkObj, PermObj: TJSONObject;
  ReturnCheck, CheckPermission: Boolean;
  PermissionInn, PermissionFn: string;
  I: Integer;
  MarkDecoded: string;
  AllAccepted, AllPermissionAccepted, Accepted, PermAccepted: Boolean;
  Cached: TMarkCacheEntry;
  HasErrors: Boolean;
begin
  Result := '';
  if Req = nil then
    Exit;

  if JBool(Req, 'skipMarkPreflight', False) then
  begin
    FLog.LogInfo('print-check: skipMarkPreflight=true — pre-flight пропущен');
    Exit;
  end;

  CollectMarksFromReceiptItems(Req, MarksB64);
  if Length(MarksB64) = 0 then
  begin
    FLog.LogInfo('print-check: pre-flight — маркированных позиций нет');
    Exit;
  end;

  ReturnCheck := GetReceiptReturnCheck(Req);
  // РР только для sell; по умолчанию включаем (ленивый клиент без permission в JSON)
  CheckPermission := (not ReturnCheck) and JBool(Req, 'checkPermission', True);

  FLog.LogInfo(Format(
    'print-check: pre-flight begin, marks=%d return=%s checkPermission=%s',
    [Length(MarksB64), BoolToStr(ReturnCheck, True), BoolToStr(CheckPermission, True)]));

  if CheckPermission then
  begin
    PermissionInn := ResolvePermissionInn(Req);
    PermissionFn := ResolvePermissionFn(Req);
  end;

  MarksOut := TJSONArray.Create;
  try
    AllAccepted := True;
    AllPermissionAccepted := True;
    HasErrors := False;

    for I := 0 to High(MarksB64) do
    begin
      MarkDecoded := DecodeMarkBase64(MarksB64[I]);
      FLog.LogInfo(Format('pre-flight марка [%d/%d], len=%d',
        [I + 1, Length(MarksB64), Length(MarkDecoded)]));

      if IsMarkFullyOk(MarkDecoded, ReturnCheck, CheckPermission) then
      begin
        TryGetMarkFromCache(MarkDecoded, ReturnCheck, Cached);
        FLog.LogInfo(Format('pre-flight [%d]: из кэша — skip', [I + 1]));
        MarkObj := MarkObjFromCacheEntry(Cached);
        try
          MarkObj.AddPair('index', TJSONNumber.Create(I));
          MarkObj.AddPair('mark', MarksB64[I]);
          if CheckPermission and Cached.RrChecked then
            MarkObj.AddPair('permission', PermissionObjFromCacheEntry(Cached));
          MarksOut.AddElement(MarkObj);
          MarkObj := nil;
        finally
          MarkObj.Free;
        end;
        Continue;
      end;

      MarkObj := CheckSingleMark(MarkDecoded, MarksB64[I], ReturnCheck,
        JBool(Req, 'draftBeer', False) or JBool(Req, 'isDraftBeer', False),
        JBool(Req, 'draftBeerLeftovers', False));
      try
        MarkObj.AddPair('index', TJSONNumber.Create(I));
        MarkObj.AddPair('mark', MarksB64[I]);

        Accepted := JBool(MarkObj, 'accepted', False);
        if not Accepted then
        begin
          AllAccepted := False;
          HasErrors := True;
        end;

        if CheckPermission then
        begin
          if TryGetMarkFromCache(MarkDecoded, ReturnCheck, Cached) and Cached.RrChecked then
          begin
            FLog.LogInfo(Format('pre-flight РР [%d]: из кэша — skip POST /document', [I + 1]));
            PermObj := PermissionObjFromCacheEntry(Cached);
          end
          else
          begin
            FLog.LogInfo(Format('pre-flight РР [%d]: проверка', [I + 1]));
            PermObj := CheckSingleMarkPermission(MarksB64[I], PermissionInn, PermissionFn);
          end;
          try
            MarkObj.AddPair('permission', PermObj);
            PermObj := nil;
            PermAccepted := False;
            if MarkObj.GetValue('permission') is TJSONObject then
              PermAccepted := JBool(TJSONObject(MarkObj.GetValue('permission')), 'accepted', False);
            if not PermAccepted then
            begin
              AllPermissionAccepted := False;
              HasErrors := True;
            end;
          finally
            PermObj.Free;
          end;
        end;

        UpdateMarkCacheFromMarkObj(MarkDecoded, ReturnCheck, MarkObj);
        MarksOut.AddElement(MarkObj);
        MarkObj := nil;
      finally
        MarkObj.Free;
      end;
    end;

    if not HasErrors then
    begin
      FLog.LogInfo('print-check: pre-flight OK — все марки готовы');
      Exit;
    end;

    Root := TJSONObject.Create;
    try
      Root.AddPair('result', TJSONNumber.Create(0));
      Root.AddPair('description', 'Не все марки прошли проверку');
      Root.AddPair('allAccepted', TJSONBool.Create(AllAccepted));
      if CheckPermission then
        Root.AddPair('allPermissionAccepted', TJSONBool.Create(AllPermissionAccepted));
      Root.AddPair('marks', MarksOut);
      MarksOut := nil;
      Result := JsonToPlainText(Root);
    finally
      Root.Free;
    end;
    FLog.LogError('print-check: pre-flight FAIL allAccepted=' +
      BoolToStr(AllAccepted, True));
  finally
    MarksOut.Free;
  end;
end;

procedure TVariantPrint.ClearStuckOpenDocument(const Context: string);
var
  Resp, ErrDesc: string;
  ErrCode: Integer;
begin
  // АТОЛ: continuePrint → cancelReceipt при залипшем документе
  if FEmulation then
    Exit;

  FLog.LogInfo(Context + ': continuePrint (очистка залипшего документа)');
  try
    if not ProcessJsonAtol(BuildAtolJsonContinuePrint, Context + ' continuePrint',
      Resp, ErrCode, ErrDesc) then
      FLog.LogInfo(Context + ': continuePrint не помог — ' + ErrDesc);
  except
    on E: Exception do
      FLog.LogInfo(Context + ': continuePrint — ' + E.Message);
  end;

  FLog.LogInfo(Context + ': cancelReceipt (очистка залипшего документа)');
  try
    CancelReceiptAtol(Context + ' cancelReceipt', False);
  except
    on E: Exception do
      FLog.LogInfo(Context + ': cancelReceipt — ' + E.Message);
  end;
end;

function TVariantPrint.DoOpenSession(const CashierName: string;
  const CashierInn: string): Integer;
var
  Name, Inn, Json, Resp, ErrDesc: string;
  ErrCode: Integer;
begin
  Name := Trim(CashierName);
  if Name = '' then
    Name := FSettings.ReadString('FRCash', 'CashierName', DefaultCashierName);
  Inn := Trim(CashierInn);
  if Inn = '' then
    Inn := Trim(FSettings.ReadString('FRCash', 'CashierInn', ''));

  FLog.LogInfo('Открываем смену (кассир: ' + Name + ')');

  Json := BuildAtolJsonOpenShift(Name, Inn);
  if not ProcessJsonAtol(Json, 'openShift', Resp, ErrCode, ErrDesc) then
    raise Exception.Create(ErrDesc);

  Result := ExtractShiftNumberFromAtolResponse(Resp);
  if FEmulation then
  begin
    Inc(FEmulatedSessionNumber);
    Result := FEmulatedSessionNumber;
  end;
  FLog.LogInfo('After openShift: shiftNumber=' + IntToStr(Result));
end;

procedure TVariantPrint.EnsureSessionOpen(const CashierName: string;
  const CashierInn: string);
var
  Opened, Expired: Boolean;
  ShiftNumber: Integer;
begin
  // Для продажи / проверки марки: нужна открытая смена < 24 ч
  ClearStuckOpenDocument('EnsureSessionOpen');

  if FEmulation then
    Exit;

  QueryShiftStatus(Opened, Expired, ShiftNumber);

  if Expired then
    raise Exception.Create(
      'Смена открыта более 24 часов. Снимите Z-отчёт (закрытие смены) и повторите.');

  if Opened then
    Exit;

  DoOpenSession(CashierName, CashierInn);
end;

procedure TVariantPrint.EnsureReadyForXReport;
var
  Opened, Expired: Boolean;
  ShiftNumber: Integer;
begin
  // X-отчёт: смена должна быть открыта; expired допустим; смену не открываем
  ClearStuckOpenDocument('XReport');

  if FEmulation then
    Exit;

  QueryShiftStatus(Opened, Expired, ShiftNumber);

  if Opened or Expired then
    Exit;

  raise Exception.Create('Смена закрыта. Откройте смену и повторите X-отчёт.');
end;

procedure TVariantPrint.EnsureReadyForZReport;
var
  Opened, Expired: Boolean;
  ShiftNumber: Integer;
begin
  // Z-отчёт: opened и expired допустимы; смену не открываем
  ClearStuckOpenDocument('ZReport');

  if FEmulation then
    Exit;

  QueryShiftStatus(Opened, Expired, ShiftNumber);

  if Opened or Expired then
    Exit;

  raise Exception.Create('Смена уже закрыта.');
end;

function TVariantPrint.CheckSingleMark(const MarkDecoded, MarkB64: string;
  ReturnCheck: Boolean; DraftBeer, DraftBeerLeftovers: Boolean): TJSONObject;
var
  Cached: TMarkCacheEntry;
  EstimatedStatus: string;
  Attempt, MaxAttempts, ErrCode, ResultCode: Integer;
  HasMore, Accepted, NeedRestart, TimedOut: Boolean;
  Resp, ErrDesc, Description, Warning, StatusJson: string;
  LocalResult: Integer;

  procedure FillMarkResult(Obj: TJSONObject; AAccepted: Boolean; ACode: Integer;
    const ADesc, AWarning: string; ALocalResult: Integer;
    const AValidationJson: string);
  begin
    Obj.AddPair('accepted', TJSONBool.Create(AAccepted));
    Obj.AddPair('resultCode', TJSONNumber.Create(ACode));
    Obj.AddPair('checkItemLocalResult', TJSONNumber.Create(ALocalResult));
    Obj.AddPair('checkItemLocalError', TJSONNumber.Create(Ord(not AAccepted)));
    Obj.AddPair('markingType2', TJSONNumber.Create(0));
    Obj.AddPair('kmServerErrorCode', TJSONNumber.Create(0));
    Obj.AddPair('kmServerCheckingStatus', TJSONNumber.Create(ALocalResult));
    Obj.AddPair('description', ADesc);
    if AWarning <> '' then
      Obj.AddPair('warning', AWarning);
    UpsertMarkCacheKkt(MarkDecoded, ReturnCheck, AAccepted, ACode, ADesc, AWarning,
      AValidationJson);
  end;

begin
  if TryGetMarkFromCache(MarkDecoded, ReturnCheck, Cached) and Cached.KktAccepted then
  begin
    FLog.LogInfo('check-marks: марка из кэша — skip beginMarkingCodeValidation, len=' +
      IntToStr(Length(MarkDecoded)) + ', desc=' + Cached.KktDescription);
    Result := MarkObjFromCacheEntry(Cached);
    Exit;
  end;

  Result := TJSONObject.Create;

  // Sentinel для smoke / эмуляции
  if FEmulation then
  begin
    if Pos('FAIL_MARK', MarkDecoded) = 1 then
    begin
      FillMarkResult(Result, False, 1, 'Отказ ККТ (emulation FAIL_MARK)', '', 0, '');
      Exit;
    end;
    if Pos('WARN_MARK', MarkDecoded) = 1 then
    begin
      FillMarkResult(Result, True, 0, 'Код маркировки успешно проверен',
        'Запрос отправлен, но ответ ОИСМ не получен (emulation WARN_MARK)', 15, '');
      Exit;
    end;
  end;

  EstimatedStatus := GetEstimatedMarkStatus(ReturnCheck, DraftBeer, DraftBeerLeftovers);
  MaxAttempts := FSettings.ReadInteger('FRCash', 'MarkTspiotRetries',
    ATOL_MARK_MAX_TSPIOT_RETRIES);
  if MaxAttempts < 1 then
    MaxAttempts := ATOL_MARK_MAX_TSPIOT_RETRIES;

  FLog.LogInfo(Format(
    'beginMarkingCodeValidation: len=%d status=%s return=%s draft=%s',
    [Length(MarkDecoded), EstimatedStatus, BoolToStr(ReturnCheck, True),
     BoolToStr(DraftBeer, True)]));

  for Attempt := 1 to MaxAttempts do
  begin
    HasMore := Attempt < MaxAttempts;

    if not ProcessJsonAtol(
      BuildAtolJsonBeginMarkValidation(MarkB64, EstimatedStatus),
      'beginMarkingCodeValidation', Resp, ErrCode, ErrDesc) then
    begin
      // 401: процедура проверки КМ уже запущена — отменить и начать заново
      if (ErrCode = 401) and HasMore then
      begin
        FLog.LogInfo('beginMarkingCodeValidation: 401 — cancel и повтор');
        CancelMarkValidationQuiet;
        Continue;
      end;
      if IsTspiotRetryCode(ErrCode) and HasMore then
      begin
        FLog.LogInfo(Format('beginMarkingCodeValidation: ТС ПИоТ %d — ожидание',
          [ErrCode]));
        Sleep(ATOL_MARK_TSPIOT_WAIT_MS);
        Continue;
      end;
      if ErrDesc = '' then
        ErrDesc := 'Ошибка начала проверки марки';
      FLog.LogError('beginMarkingCodeValidation: ' + ErrDesc);
      FillMarkResult(Result, False, ErrCode, ErrDesc, '', 0, '');
      Exit;
    end;

    if not WaitMarkValidationReady(HasMore, StatusJson, ErrCode, ErrDesc,
      NeedRestart, TimedOut) then
    begin
      if NeedRestart then
      begin
        CancelMarkValidationQuiet;
        Sleep(ATOL_MARK_TSPIOT_WAIT_MS);
        Continue;
      end;
      CancelMarkValidationQuiet;
      if TimedOut then
        FillMarkResult(Result, False, 0, ErrDesc, '', 0, '')
      else
        FillMarkResult(Result, False, ErrCode, ErrDesc, '', 0, '');
      Exit;
    end;

    ClassifyMarkValidationResponse(StatusJson, Accepted, Description, Warning,
      ResultCode);

    if Accepted then
    begin
      if not ProcessJsonAtol(BuildAtolJsonAcceptMarkingCode, 'acceptMarkingCode',
        Resp, ErrCode, ErrDesc) then
      begin
        if IsTspiotRetryCode(ErrCode) and HasMore then
        begin
          FLog.LogInfo(Format('acceptMarkingCode: ТС ПИоТ %d — повтор', [ErrCode]));
          CancelMarkValidationQuiet;
          Sleep(ATOL_MARK_TSPIOT_WAIT_MS);
          Continue;
        end;
        CancelMarkValidationQuiet;
        if ErrDesc = '' then
          ErrDesc := 'Ошибка принятия марки';
        FLog.LogError('acceptMarkingCode: ' + ErrDesc);
        FillMarkResult(Result, False, ErrCode, ErrDesc, '', 0, '');
        Exit;
      end;

      LocalResult := 15;
      if Description = '' then
        Description := 'Код маркировки успешно проверен';
      FLog.LogInfo(Format('check-marks OK accepted=true warn=%s', [Warning]));
      FillMarkResult(Result, True, 0, Description, Warning, LocalResult, StatusJson);
      Exit;
    end;

    // Марка не прошла проверку
    CancelMarkValidationQuiet;
    if Description = '' then
      Description := 'Код маркировки не прошёл проверку';
    FLog.LogInfo('check-marks rejected: ' + Description);
    FillMarkResult(Result, False, ResultCode, Description, '', 0, '');
    Exit;
  end;

  FillMarkResult(Result, False, 0, 'Не удалось выполнить проверку марки', '', 0, '');
end;

procedure TVariantPrint.ReadInnAndFn(out Inn, FnNumber: string);
var
  Resp, ErrDesc: string;
  ErrCode: Integer;
  V: TJSONValue;
  Root, Org, Device: TJSONObject;

  function NestedStr(Parent: TJSONObject; const ChildName, FieldName: string): string;
  var
    Child: TJSONObject;
  begin
    Result := '';
    if Parent = nil then
      Exit;
    Result := Trim(JStr(Parent, FieldName));
    if Result <> '' then
      Exit;
    Child := Parent.GetValue(ChildName) as TJSONObject;
    if Child <> nil then
      Result := Trim(JStr(Child, FieldName));
  end;

begin
  Inn := '';
  FnNumber := '';
  if FEmulation then
  begin
    Inn := FEmulatedInn;
    FnNumber := FEmulatedFnNumber;
    Exit;
  end;
  if not FRObjectReady then
    Exit;

  // ИНН / ФН через JSON АТОЛ (как /info), для разрешительного режима
  if ProcessJsonAtol('{"type":"getRegistrationInfo"}', 'getRegistrationInfo permission',
    Resp, ErrCode, ErrDesc) then
  begin
    V := TJSONObject.ParseJSONValue(Resp);
    if V <> nil then
    try
      if V is TJSONObject then
      begin
        Root := TJSONObject(V);
        Org := Root.GetValue('organization') as TJSONObject;
        Device := Root.GetValue('device') as TJSONObject;
        if Org <> nil then
        begin
          Inn := Trim(JStr(Org, 'vatin'));
          if Inn = '' then
            Inn := Trim(JStr(Org, 'inn'));
        end;
        if Inn = '' then
          Inn := Trim(JStr(Root, 'vatin'));
        if Inn = '' then
          Inn := Trim(JStr(Root, 'inn'));
        if Device <> nil then
          FnNumber := Trim(JStr(Device, 'fnNumber'));
      end;
    finally
      V.Free;
    end;
  end;

  if FnNumber = '' then
  begin
    if ProcessJsonAtol('{"type":"getFnInfo"}', 'getFnInfo permission',
      Resp, ErrCode, ErrDesc) then
    begin
      V := TJSONObject.ParseJSONValue(Resp);
      if V <> nil then
      try
        if V is TJSONObject then
        begin
          Root := TJSONObject(V);
          FnNumber := Trim(JStr(Root, 'serial'));
          if FnNumber = '' then
            FnNumber := Trim(JStr(Root, 'fnNumber'));
          if FnNumber = '' then
            FnNumber := NestedStr(Root, 'fnInfo', 'serial');
          if FnNumber = '' then
            FnNumber := NestedStr(Root, 'fnInfo', 'fnNumber');
        end;
      finally
        V.Free;
      end;
    end;
  end;
end;

function TVariantPrint.ResolvePermissionInn(Req: TJSONObject): string;
var
  Fn: string;
begin
  Result := JStr(Req, 'inn', '');
  if Result = '' then
    Result := JStr(Req, 'ИНН', '');
  if Result = '' then
    ReadInnAndFn(Result, Fn);
end;

function TVariantPrint.ResolvePermissionFn(Req: TJSONObject): string;
var
  Inn: string;
begin
  Result := JStr(Req, 'fnNumber', '');
  if Result = '' then
    Result := JStr(Req, 'fn', '');
  if Result = '' then
    Result := JStr(Req, 'ФН', '');
  if Result = '' then
    ReadInnAndFn(Inn, Result);
end;

function TVariantPrint.BuildDocumentCheckJson(const MarkB64, Inn, Fn: string): string;
var
  Root, Pos, Org: TJSONObject;
  PositionsArr, CodesArr: TJSONArray;
begin
  Root := TJSONObject.Create;
  try
    Root.AddPair('action', 'check');
    Root.AddPair('type', 'receipt');
    if Trim(Fn) <> '' then
      Root.AddPair('shift', Fn);

    PositionsArr := TJSONArray.Create;
    Pos := TJSONObject.Create;
    Org := TJSONObject.Create;
    Org.AddPair('inn', Inn);
    Pos.AddPair('organization', Org);

    CodesArr := TJSONArray.Create;
    CodesArr.AddElement(TJSONString.Create(MarkB64));
    Pos.AddPair('marking_codes', CodesArr);

    PositionsArr.AddElement(Pos);
    Root.AddPair('positions', PositionsArr);
    Result := JsonToPlainText(Root);
  finally
    Root.Free;
  end;
end;

function TVariantPrint.HttpPostJson(const Url, Body: string): string;
var
  Http: TIdHTTP;
  ReqStream: TStringStream;
begin
  Http := TIdHTTP.Create(nil);
  ReqStream := TStringStream.Create(Body, TEncoding.UTF8);
  try
    Http.Request.ContentType := 'application/json';
    Http.ConnectTimeout := 30000;
    Http.ReadTimeout := 180000;
    Result := Http.Post(Url, ReqStream);
  finally
    ReqStream.Free;
    Http.Free;
  end;
end;

function TVariantPrint.ExtractFirstTruemarkCode(Truemark: TJSONObject): TJSONObject;
var
  CodesV: TJSONValue;
  Arr: TJSONArray;
begin
  Result := nil;
  if Truemark = nil then
    Exit;
  CodesV := Truemark.GetValue('codes');
  if CodesV is TJSONArray then
  begin
    Arr := TJSONArray(CodesV);
    if (Arr.Count > 0) and (Arr.Items[0] is TJSONObject) then
      Result := TJSONObject(Arr.Items[0]);
  end;
end;

procedure TVariantPrint.AddTruemarkFields(PermObj, Truemark: TJSONObject);
const
  FieldNames: array[0..8] of string = (
    'valid', 'verified', 'found', 'realizable', 'utilised',
    'isBlocked', 'sold', 'isOwner', 'grayZone');
var
  TruemarkOut, CodeObj: TJSONObject;
  I: Integer;
  Val: string;
begin
  if PermObj = nil then
    Exit;

  CodeObj := ExtractFirstTruemarkCode(Truemark);
  if (CodeObj = nil) and (Truemark = nil) then
    Exit;

  TruemarkOut := TJSONObject.Create;
  try
    for I := Low(FieldNames) to High(FieldNames) do
    begin
      Val := '';
      if CodeObj <> nil then
        Val := JStr(CodeObj, FieldNames[I], '');
      if (Val = '') and (Truemark <> nil) then
        Val := JStr(Truemark, FieldNames[I], '');
      if Val <> '' then
        TruemarkOut.AddPair(FieldNames[I], Val);
    end;

    if CodeObj <> nil then
    begin
      Val := JStr(CodeObj, 'message', '');
      if Val <> '' then
        TruemarkOut.AddPair('message', Val);
      Val := JStr(CodeObj, 'errorCode', '');
      if Val <> '' then
        TruemarkOut.AddPair('errorCode', Val);
      Val := JStr(CodeObj, 'cis', '');
      if Val <> '' then
        TruemarkOut.AddPair('cis', Val);
      Val := JStr(CodeObj, 'gtin', '');
      if Val <> '' then
        TruemarkOut.AddPair('gtin', Val);
    end;

    if Truemark <> nil then
    begin
      Val := JStr(Truemark, 'description', '');
      if Val <> '' then
        TruemarkOut.AddPair('description', Val);
      Val := JStr(Truemark, 'code', '');
      if Val <> '' then
        TruemarkOut.AddPair('truemarkCode', Val);
    end;

    if TruemarkOut.Count > 0 then
      PermObj.AddPair('truemark', TruemarkOut)
    else
      TruemarkOut.Free;
  except
    TruemarkOut.Free;
    raise;
  end;
end;

function TVariantPrint.ParsePermissionCheckResponse(const ResponseBody: string): TJSONObject;
var
  V, TruemarkV: TJSONValue;
  Root, Truemark: TJSONObject;
  CodeStr, Err, Uuid, TimeStr, Inst, Version: string;
  Accepted: Boolean;
begin
  Result := TJSONObject.Create;

  V := TJSONObject.ParseJSONValue(ResponseBody);
  if not (V is TJSONObject) then
  begin
    if V <> nil then
      V.Free;
    Result.AddPair('accepted', TJSONBool.Create(False));
    Result.AddPair('code', '-1');
    Result.AddPair('error', 'Некорректный JSON ответа FMU');
    Exit;
  end;

  Root := TJSONObject(V);
  try
    CodeStr := JStr(Root, 'Code', '');
    if CodeStr = '' then
      CodeStr := JStr(Root, 'code', '');
    Accepted := (CodeStr = '0') or SameText(CodeStr, '0.0');

    Err := JStr(Root, 'Error', '');
    if Err = '' then
      Err := JStr(Root, 'error', '');

    Uuid := '';
    TimeStr := '';
    Inst:= '';
    Version:='';
    TruemarkV := Root.GetValue('Truemark_response');
    if TruemarkV = nil then
      TruemarkV := Root.GetValue('truemark_response');
    if TruemarkV is TJSONObject then
    begin
      Truemark := TJSONObject(TruemarkV);
      Uuid := JStr(Truemark, 'reqId', '');
      TimeStr := JStr(Truemark, 'reqTimestamp', '');
      Inst := JStr(Truemark, 'inst', '');
      Version := JStr(Truemark, 'version', '');
      AddTruemarkFields(Result, Truemark);
    end
    else
    begin
      Uuid := JStr(Root, 'reqId', '');
      TimeStr := JStr(Root, 'reqTimestamp', '');
      Inst := JStr(Root, 'inst', '');
      Version := JStr(Root, 'version', '');
    end;

    Result.AddPair('accepted', TJSONBool.Create(Accepted));
    Result.AddPair('code', CodeStr);
    Result.AddPair('error', Err);
    Result.AddPair('uuid', Uuid);
    Result.AddPair('time', TimeStr);
    Result.AddPair('inst', Inst);
    Result.AddPair('version', Version);
  finally
    Root.Free;
  end;
end;

function TVariantPrint.CheckSingleMarkPermission(const MarkB64, Inn, Fn: string): TJSONObject;
var
  Url, ReqJson, RespBody: string;
begin
  // Emulation влияет только на ККТ/ФР; РР всегда идёт на реальный FMU /document
  if FEmulation then
    FLog.LogInfo('РР: Emulation=true — фейковый ответ отключён, POST /document как обычно');

  if Trim(Inn) = '' then
    raise Exception.Create('ИНН не задан для проверки разрешительного режима');

  ReqJson := BuildDocumentCheckJson(MarkB64, Inn, Fn);
  Url := Format('http://%s:%d%s', [PERMISSION_FMU_HOST, PERMISSION_FMU_PORT, PERMISSION_FMU_PATH]);
  FLog.LogInfo('POST /document (РР): ' + Url);

  try
    RespBody := HttpPostJson(Url, ReqJson);
  except
    on E: Exception do
    begin
      FLog.LogError('POST /document (РР): ' + E.Message);
      Result := TJSONObject.Create;
      Result.AddPair('accepted', TJSONBool.Create(False));
      Result.AddPair('code', '-1');
      Result.AddPair('error', E.Message);
      Result.AddPair('uuid', '');
      Result.AddPair('time', '');
      Result.AddPair('inst', '');
      Result.AddPair('version', '');
      Exit;
    end;
  end;

  FLog.LogInfo('POST /document (РР) ответ: ' + JsonBodyPreviewPlain(RespBody, 500));
  Result := ParsePermissionCheckResponse(RespBody);
end;

function TVariantPrint.CheckMarkInternal(const Body: string): string;
var
  V: TJSONValue;
  Req: TJSONObject;
  MarksB64: TArray<string>;
  MarksOut: TJSONArray;
  Root: TJSONObject;
  MarkObj, PermObj: TJSONObject;
  ReturnCheck: Boolean;
  CheckPermission: Boolean;
  DraftBeer, DraftBeerLeftovers: Boolean;
  CashierName, PermissionInn, PermissionFn: string;
  I: Integer;
  MarkDecoded: string;
  AllAccepted, AllPermissionAccepted, PermAccepted, Accepted: Boolean;
  HasWarnings: Boolean;
  Cached: TMarkCacheEntry;
begin
  V := TJSONObject.ParseJSONValue(Body);
  if not (V is TJSONObject) then
  begin
    if V <> nil then
      V.Free;
    FLog.LogError(Format('CheckMarkInternal: Invalid JSON, bodyLen=%d, head=%s',
      [Length(Body), JsonBodyPreview(Body)]));
    raise Exception.Create(Format('Invalid JSON, bodyLen=%d', [Length(Body)]));
  end;
  Req := TJSONObject(V);

  try
    FLog.LogInfo('CheckMarkInternal begin');
    EnsureFRConnected;

    CollectMarksFromRequest(Req, MarksB64);
    ReturnCheck := GetReceiptReturnCheck(Req);
    CheckPermission := JBool(Req, 'checkPermission', False);
    DraftBeer := JBool(Req, 'draftBeer', False) or JBool(Req, 'isDraftBeer', False);
    DraftBeerLeftovers := JBool(Req, 'draftBeerLeftovers', False);
    CashierName := GetCashierName(Req);

    if CheckPermission and not ReturnCheck then
    begin
      PermissionInn := ResolvePermissionInn(Req);
      PermissionFn := ResolvePermissionFn(Req);
    end;

    EnsureSessionOpen(CashierName, GetCashierInn(Req));

    MarksOut := TJSONArray.Create;
    Root := TJSONObject.Create;
    try
      AllAccepted := True;
      AllPermissionAccepted := True;
      HasWarnings := False;
      for I := 0 to High(MarksB64) do
      begin
        MarkDecoded := DecodeMarkBase64(MarksB64[I]);
        FLog.LogInfo(Format('Проверка марки [%d/%d], длина=%d',
          [I + 1, Length(MarksB64), Length(MarkDecoded)]));

        if IsMarkFullyOk(MarkDecoded, ReturnCheck, CheckPermission) then
        begin
          TryGetMarkFromCache(MarkDecoded, ReturnCheck, Cached);
          FLog.LogInfo(Format('Марка [%d/%d] полностью из кэша (ККТ+РР) — skip',
            [I + 1, Length(MarksB64)]));
          MarkObj := MarkObjFromCacheEntry(Cached);
          try
            if CheckPermission and not ReturnCheck and Cached.RrChecked then
              MarkObj.AddPair('permission', PermissionObjFromCacheEntry(Cached));
            MarkObj.AddPair('index', TJSONNumber.Create(I));
            if not Cached.KktAccepted then
              AllAccepted := False;
            if CheckPermission and not ReturnCheck and not Cached.RrAccepted then
              AllPermissionAccepted := False;
            if Cached.KktWarning <> '' then
              HasWarnings := True;
            MarksOut.AddElement(MarkObj);
            MarkObj := nil;
          finally
            MarkObj.Free;
          end;
          Continue;
        end;

        MarkObj := CheckSingleMark(MarkDecoded, MarksB64[I], ReturnCheck,
          DraftBeer, DraftBeerLeftovers);
        try
          if CheckPermission and not ReturnCheck then
          begin
            if TryGetMarkFromCache(MarkDecoded, ReturnCheck, Cached) and Cached.RrChecked then
            begin
              FLog.LogInfo(Format('РР [%d/%d] из кэша — skip POST /document',
                [I + 1, Length(MarksB64)]));
              PermObj := PermissionObjFromCacheEntry(Cached);
            end
            else
            begin
              FLog.LogInfo(Format('Проверка РР [%d/%d]', [I + 1, Length(MarksB64)]));
              PermObj := CheckSingleMarkPermission(MarksB64[I], PermissionInn, PermissionFn);
            end;
            try
              MarkObj.AddPair('permission', PermObj);
              PermObj := nil;
              PermAccepted := False;
              if MarkObj.GetValue('permission') is TJSONObject then
              begin
                if TJSONObject(MarkObj.GetValue('permission')).GetValue('accepted') is TJSONBool then
                  PermAccepted := TJSONBool(TJSONObject(MarkObj.GetValue('permission')).GetValue('accepted')).AsBoolean;
              end;
              if not PermAccepted then
                AllPermissionAccepted := False;
            finally
              PermObj.Free;
            end;
          end;

          UpdateMarkCacheFromMarkObj(MarkDecoded, ReturnCheck, MarkObj);

          MarkObj.AddPair('index', TJSONNumber.Create(I));
          Accepted := False;
          if MarkObj.GetValue('accepted') is TJSONBool then
            Accepted := TJSONBool(MarkObj.GetValue('accepted')).AsBoolean;
          if not Accepted then
            AllAccepted := False;
          if Trim(JStr(MarkObj, 'warning', '')) <> '' then
            HasWarnings := True;
          MarksOut.AddElement(MarkObj);
          MarkObj := nil;
        finally
          MarkObj.Free;
        end;
      end;

      Root.AddPair('result', TJSONNumber.Create(1));
      Root.AddPair('description', 'OK');
      Root.AddPair('marks', MarksOut);
      MarksOut := nil;
      Root.AddPair('allAccepted', TJSONBool.Create(AllAccepted));
      Root.AddPair('hasWarnings', TJSONBool.Create(HasWarnings));
      if CheckPermission and not ReturnCheck then
        Root.AddPair('allPermissionAccepted', TJSONBool.Create(AllPermissionAccepted));
      Result := JsonToPlainText(Root);
      if CheckPermission and not ReturnCheck then
        FLog.LogInfo('CheckMarkInternal OK allAccepted=' + BoolToStr(AllAccepted, True) +
          ', allPermissionAccepted=' + BoolToStr(AllPermissionAccepted, True) +
          ', hasWarnings=' + BoolToStr(HasWarnings, True))
      else
        FLog.LogInfo('CheckMarkInternal OK allAccepted=' + BoolToStr(AllAccepted, True) +
          ', hasWarnings=' + BoolToStr(HasWarnings, True));
      AfterCommand(Req);
    finally
      MarksOut.Free;
      Root.Free;
    end;
  except
    on E: Exception do
    begin
      FLog.LogError('CheckMarkInternal: ' + E.Message);
      if FEmulation then
      begin
        Result := JsonOk('"allAccepted":true,"hasWarnings":false,"marks":[{"index":0,"accepted":true,' +
          '"resultCode":0,"checkItemLocalResult":15,"checkItemLocalError":0,' +
          '"markingType2":0,"kmServerErrorCode":0,"kmServerCheckingStatus":15,' +
          '"description":"OK (emulation)"}]');
        AfterCommand(Req);
      end
      else
        Result := JsonError(E.Message);
    end;
  end;

  V.Free;
end;

procedure TVariantPrint.StartAsyncCheckMark(const Body, TaskId: string);
var
  Entry: TCheckMarkTaskEntry;
begin
  Entry := TCheckMarkTaskEntry.Create;
  Entry.Status := 'pending';
  Entry.CreatedAt := Now;

  FMarkTasksLock.Enter;
  try
    TrimOldMarkTasks;
    FMarkTasks.Add(TaskId, Entry);
  finally
    FMarkTasksLock.Leave;
  end;

  FLog.LogInfo('async check-marks: задача принята taskId=' + TaskId);

  TThread.CreateAnonymousThread(
    procedure
    var
      JsonResult: string;
      TaskEntry: TCheckMarkTaskEntry;
      Parsed: TJSONValue;
      Ok: Boolean;
    begin
      // FLock сериализует ККТ: async не пересекается с sync / print / сменой.
      // COM Fptr10 — только через ExecOnFrThread внутри CheckMarkInternal.
      FLock.Enter;
      try
        try
          // CheckMarkInternal пишет marks[] (+ permission) в FMarkCache
          JsonResult := CheckMarkInternal(Body);
        except
          on E: Exception do
            JsonResult := JsonError(E.Message);
        end;
      finally
        FLock.Leave;
      end;

      Ok := False;
      Parsed := TJSONObject.ParseJSONValue(JsonResult);
      try
        if Parsed is TJSONObject then
          Ok := JInt(TJSONObject(Parsed), 'result', 0) = 1;
      finally
        Parsed.Free;
      end;

      FMarkTasksLock.Enter;
      try
        if FMarkTasks.TryGetValue(TaskId, TaskEntry) then
        begin
          if Ok then
            TaskEntry.Status := 'completed'
          else
            TaskEntry.Status := 'error';
          TaskEntry.ResultJson := JsonResult;
        end;
      finally
        FMarkTasksLock.Leave;
      end;

      if Ok then
        FLog.LogInfo('async check-marks: completed taskId=' + TaskId +
          ' — FMarkCache обновлён (через CheckMarkInternal)')
      else
        FLog.LogError('async check-marks: error taskId=' + TaskId + ' ' +
          JsonBodyPreviewPlain(JsonResult, 200));
    end).Start;
end;

function TVariantPrint.CheckMarkJson(const Body: string): string;
var
  V: TJSONValue;
  Req: TJSONObject;
  TaskId: string;
begin
  V := TJSONObject.ParseJSONValue(Body);
  try
    if not (V is TJSONObject) then
    begin
      FLog.LogError(Format('CheckMarkJson: Invalid JSON, bodyLen=%d, head=%s',
        [Length(Body), JsonBodyPreview(Body)]));
      raise Exception.Create(Format('Invalid JSON, bodyLen=%d', [Length(Body)]));
    end;
    Req := TJSONObject(V);

    if JBool(Req, 'async', False) then
    begin
      TaskId := NewTaskId;
      StartAsyncCheckMark(Body, TaskId);
      Result := JsonOk('"async":true,"taskId":"' + JsonEscape(TaskId) +
        '","status":"pending"');
      Exit;
    end;

    FLock.Enter;
    try
      Result := CheckMarkInternal(Body);
    finally
      FLock.Leave;
    end;
  finally
    V.Free;
  end;
end;

function TVariantPrint.CheckMarkStatusJson(const TaskId: string): string;
var
  Entry: TCheckMarkTaskEntry;
  Root: TJSONObject;
  Parsed: TJSONValue;
begin
  if Trim(TaskId) = '' then
    Exit(JsonError('taskId is required'));

  FMarkTasksLock.Enter;
  try
    if not FMarkTasks.TryGetValue(TaskId, Entry) then
      Exit(JsonError('task not found'));

    if Entry.Status = 'pending' then
      Exit(JsonOk('"taskId":"' + JsonEscape(TaskId) + '","status":"pending"'));

    Parsed := TJSONObject.ParseJSONValue(Entry.ResultJson);
    if not (Parsed is TJSONObject) then
      Exit(JsonError('invalid task result'));

    Root := TJSONObject(Parsed);
    try
      Root.AddPair('taskId', TaskId);
      Root.AddPair('status', Entry.Status);
      Result := JsonToPlainText(Root);
    finally
      Root.Free;
    end;
  finally
    FMarkTasksLock.Leave;
  end;
end;

end.
