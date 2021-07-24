//
// OFX - Open Financial Exchange
// OFC - Open Financial Connectivity

// 2006 - Eduardo Bento da Rocha (YoungArts)
// 2016 - Leonardo Gregianin - github.com/leogregianin
// 2021 - Marivaldo Santos - github.com/onyxadm

unit ofxreader;

interface

{$DEFINE HAS_FORMATSETTINGS}

uses
  Classes,
  SysUtils,
  StrUtils;

type
  TOFXItem = class
  public
    MovType    : string;
    MovDate    : TDateTime;
    Value      : Currency;
    ID         : string;
    Document   : string;
    Description: string;
  end;

  TOFXReader = class(TComponent)
  private
    FOFXFile     : string;
    FListItems   : TList;
    FDateEnd     : TDateTime;
    FBranchID    : string;
    FBankID      : string;
    FAccountID   : string;
    FAccountType : string;
    FDateStart   : TDateTime;
    FFinalBalance: Currency;
    FDateBalance : TDateTime;

    procedure Clear;
    procedure Delete(iIndex: Integer);

    function Add: TOFXItem;
    function InfLine(sLine: string): string;
    function FindString(sSubString, sString: string): Boolean;
  public
    constructor Create(AOwner: TComponent); overload;
    constructor Create(AFile: string); overload;
    destructor Destroy; override;

    function Import(OFXFile: string = ''): Boolean;
    function Get(iIndex: Integer): TOFXItem;
    function Count: Integer;
  published
    property OFXFile: string read FOFXFile write FOFXFile;

    property BankID      : string read FBankID write FBankID;
    property BranchID    : string read FBranchID write FBranchID;
    property AccountID   : string read FAccountID write FAccountID;
    property AccountType : string read FAccountType write FAccountType;
    property DateStart   : TDateTime read FDateStart write FDateStart;
    property DateEnd     : TDateTime read FDateEnd write FDateEnd;
    property DateBalance : TDateTime read FDateBalance write FDateBalance;
    property FinalBalance: Currency read FFinalBalance write FFinalBalance;
  end;

procedure Register;

implementation

{$REGION ' Funções auxiliares '}

{ -----------------------------------------------------------------------------
  Converte uma <NumString> para Double, semelhante ao StrToFloat, mas
  verifica se a virgula é '.' ou ',' efetuando a conversão se necessário
  Se não for possivel converter, dispara Exception
  ---------------------------------------------------------------------------- }
function StringToFloat(NumString: String): Double;
var
  DS: Char;

  { -----------------------------------------------------------------------------
    Retorna quantas ocorrencias de <SubStr> existem em <AString>
    ---------------------------------------------------------------------------- }
  function CountStr(const AString, SubStr: String): Integer;
  Var
    ini: Integer;
  begin
    Result := 0;
    if SubStr = '' then
      exit;

    ini := Pos(SubStr, AString);
    while ini > 0 do
    begin
      Result := Result + 1;
      ini    := PosEx(SubStr, AString, ini + 1);
    end;
  end;

begin
  NumString := Trim(NumString);

  DS := {$IFDEF HAS_FORMATSETTINGS}FormatSettings.{$ENDIF}DecimalSeparator;

  if DS <> '.' then
    NumString := StringReplace(NumString, '.', DS, [rfReplaceAll]);

  if DS <> ',' then
    NumString := StringReplace(NumString, ',', DS, [rfReplaceAll]);

  while CountStr(NumString, DS) > 1 do
    NumString := StringReplace(NumString, DS, '', []);

  Result := StrToFloat(NumString);
end;

{ -----------------------------------------------------------------------------
  Converte uma <NumString> para Double, semelhante ao StrToFloatDef, mas
  verifica se a virgula é '.' ou ',' efetuando a conversão se necessário
  Se não for possivel converter, retorna <DefaultValue>
  ---------------------------------------------------------------------------- }
function StringToFloatDef(const NumString: String; const DefaultValue: Double): Double;
begin
  if Trim(NumString).isEmpty then
    Result := DefaultValue
  else
  begin
    try
      Result := StringToFloat(NumString);
    except
      Result := DefaultValue;
    end;
  end;
end;

{ -----------------------------------------------------------------------------
  Remove todos os espacos duplos do texto
  ---------------------------------------------------------------------------- }
function RemoverEspacosDuplos(const AString: String): String;
begin
  Result := Trim(AString);
  while Pos('  ', Result) > 0 do
    Result := StringReplace(Result, '  ', ' ', [rfReplaceAll]);
end;
{$ENDREGION}

constructor TOFXReader.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FListItems := TList.Create;
end;

constructor TOFXReader.Create(AFile: string);
begin
  Create(nil);

  if FileExists(AFile) then
  begin
    FOFXFile := AFile;
    Import(AFile);
  end;
end;

destructor TOFXReader.Destroy;
begin
  Clear;
  FListItems.Free;
  inherited Destroy;
end;

procedure TOFXReader.Delete(iIndex: Integer);
begin
  TOFXItem(FListItems.Items[iIndex]).Free;
  FListItems.Delete(iIndex);
end;

procedure TOFXReader.Clear;
begin
  while FListItems.Count > 0 do
    Delete(0);
  FListItems.Clear;
end;

function TOFXReader.Count: Integer;
begin
  Result := FListItems.Count;
end;

function TOFXReader.Get(iIndex: Integer): TOFXItem;
begin
  Result := TOFXItem(FListItems.Items[iIndex]);
end;

function TOFXReader.Import(OFXFile: string = ''): Boolean;
var
  oFile: TStringList;
  i    : Integer;
  bOFX : Boolean;
  oItem: TOFXItem;
  sLine: string;
begin
  Clear;
  DateStart := 0;
  DateEnd   := 0;
  bOFX      := False;

  if not FileExists(OFXFile) or Trim(OFXFile).isEmpty then
  begin
    if FileExists(FOFXFile) then
      OFXFile := FOFXFile;
  end;

  if not FileExists(OFXFile) then
    raise Exception.Create('File not found!');

  FOFXFile := OFXFile;

  oFile := TStringList.Create;
  try
    oFile.LoadFromFile(OFXFile);
    i := 0;

    while i < oFile.Count do
    begin
      sLine := oFile.Strings[i];
      if FindString('<OFX>', sLine) or FindString('<OFC>', sLine) then
        bOFX := True;

      if bOFX then
      begin
        // Bank
        if FindString('<BANKID>', sLine) then
          BankID := InfLine(sLine);

        // Agency
        if FindString('<BRANCHID>', sLine) then
          BranchID := InfLine(sLine);

        // Account
        if FindString('<ACCTID>', sLine) then
          AccountID := InfLine(sLine);

        // Account type
        if FindString('<ACCTTYPE>', sLine) then
          AccountType := InfLine(sLine);

        // Date Start and Date End
        if FindString('<DTSTART>', sLine) then
        begin
          if Trim(sLine) <> '' then
            DateStart := EncodeDate( //
              StrToIntDef(copy(InfLine(sLine), 1, 4), 0), //
              StrToIntDef(copy(InfLine(sLine), 5, 2), 0), //
              StrToIntDef(copy(InfLine(sLine), 7, 2), 0));
        end;
        if FindString('<DTEND>', sLine) then
        begin
          if Trim(sLine) <> '' then
            DateEnd := EncodeDate( //
              StrToIntDef(copy(InfLine(sLine), 1, 4), 0), //
              StrToIntDef(copy(InfLine(sLine), 5, 2), 0), //
              StrToIntDef(copy(InfLine(sLine), 7, 2), 0));
        end;

        // Final
        if FindString('<LEDGER>', sLine) or FindString('<BALAMT>', sLine) then
          FinalBalance := StringToFloatDef(InfLine(sLine), 0);

        // Date Balance
        if FindString('<DTASOF>', sLine) then
        begin
          if Trim(sLine) <> '' then
            DateBalance := EncodeDate( //
              StrToIntDef(copy(InfLine(sLine), 1, 4), 0), //
              StrToIntDef(copy(InfLine(sLine), 5, 2), 0), //
              StrToIntDef(copy(InfLine(sLine), 7, 2), 0));
        end;

        // Movement
        if FindString('<STMTTRN>', sLine) then
        begin
          oItem := Add;
          while not FindString('</STMTTRN>', sLine) do
          begin
            Inc(i);
            sLine := oFile.Strings[i];

            if FindString('<TRNTYPE>', sLine) then
            begin
              if (InfLine(sLine) = '0') or (InfLine(sLine) = 'CREDIT') OR (InfLine(sLine) = 'DEP') then
                oItem.MovType := 'C'
              else
                if (InfLine(sLine) = '1') or (InfLine(sLine) = 'DEBIT') OR (InfLine(sLine) = 'XFER') then
                  oItem.MovType := 'D'
                else
                  oItem.MovType := 'OTHER';
            end;

            if FindString('<DTPOSTED>', sLine) then
              oItem.MovDate := EncodeDate( //
                StrToIntDef(copy(InfLine(sLine), 1, 4), 0), //
                StrToIntDef(copy(InfLine(sLine), 5, 2), 0), //
                StrToIntDef(copy(InfLine(sLine), 7, 2), 0));

            if FindString('<FITID>', sLine) then
              oItem.ID := Trim(InfLine(sLine));

            if FindString('<CHKNUM>', sLine) or FindString('<CHECKNUM>', sLine) or FindString('<REFNUM>', sLine)
            then
              oItem.Document := InfLine(sLine);

            if Trim(oItem.Document).isEmpty then
              oItem.Document := oItem.ID;

            oItem.Document := RemoverEspacosDuplos(oItem.Document);

            if FindString('<MEMO>', sLine) then
              oItem.Description := RemoverEspacosDuplos(InfLine(sLine));

            if FindString('<TRNAMT>', sLine) then
              oItem.Value := StringToFloatDef(InfLine(sLine), 0);
          end;
        end;

      end;

      Inc(i);
    end;

    Result := bOFX;
  finally
    oFile.Free;
  end;
end;

function TOFXReader.InfLine(sLine: string): string;
var
  iTemp: Integer;
begin
  Result := '';
  sLine  := Trim(sLine);
  if FindString('>', sLine) then
  begin
    sLine := Trim(sLine);
    iTemp := Pos('>', sLine);

    if Pos('</', sLine) > 0 then
      Result := copy(sLine, iTemp + 1, Pos('</', sLine) - iTemp - 1)
    else
      // allows you to read the whole line when there is no completion of </ on the same line
      // made by weberdepaula@gmail.com
      Result := copy(sLine, iTemp + 1, length(sLine));
  end;
end;

function TOFXReader.Add: TOFXItem;
var
  oItem: TOFXItem;
begin
  oItem := TOFXItem.Create;
  FListItems.Add(oItem);
  Result := oItem;
end;

function TOFXReader.FindString(sSubString, sString: string): Boolean;
begin
  Result := Pos(UpperCase(sSubString), UpperCase(sString)) > 0;
end;

procedure Register;
begin
  RegisterComponents('OFXReader ONYX', [TOFXReader]);
end;

end.
