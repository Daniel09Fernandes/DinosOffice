{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOffice_calc.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }

{ Documentation:                                           }
{
  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Editing_Spreadsheet_Documents
  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Cells_and_Ranges
  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Templates
  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/StarDesktop
}
{ ******************************************************* }

unit uOpenOffice_calc;

interface

uses
  System.Classes, DB, ActiveX, uOpenOffice,
  dbWeb, ComObj, XMLDoc, XMLIntf, Vcl.Dialogs, System.Variants,
  Windows, uOpenOfficeEvents, Datasnap.DBClient;

type

  TTypeValue = (ftString, ftNumeric);

  TFieldsSheet = record
  private
  var
    arrFields: array of string;
    procedure setArrayFieldsSheet;

  public
    function getField(aIndex: integer): String;
    function getIndex(aNameField: String): integer;
  end;

  TOpenOffice_calc = class(TOpenOffice)
  private
    const
    DefaultNewSheetNamePT = 'Planilha1';
    DefaultNewSheetNameEn = 'Sheet1';
    procedure ValidateSheetName;

    var
    //--------events------//
    FOnBeforeStartFile: TBeforeStartFile;
    FOnAfterStartFile : TAfterStartFile;
    //--------------------//
    FFields: TFieldsSheet;
    FSheetName: string;

    procedure SetSheetName(const Value: string);
  public
  var
    Value: string;
    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
    procedure startSheet;
    procedure positionSheetByName(aSheetName: string);
    procedure addNewSheet(aSheetName: string; aPosition: integer);
    function setFormula(aCellNumber: integer; aCollName: string; aFormula: string): TOpenOffice_calc;
    function SetValue(aCellNumber: integer; aCollName: string; aValue: variant; TypeValue: TTypeValue = ftString; Wrapped: boolean = false): TOpenOffice_calc;
    function GetValue(aCellNumber: integer; aCollName: String): TOpenOffice_calc;
    procedure DataSetToSheet(const aCds : TClientDataSet);
  published
    property ServicesManager: OleVariant read objServiceManager;
    property Cell: OleVariant read objCell write objCell;
    property Table: OleVariant read objSCalc;
    property Fields: TFieldsSheet read FFields;
    property SheetName: string read FSheetName write SetSheetName;

    //---------events-----------//
    property OnBeforeStartFile: TBeforeStartFile read FOnBeforeStartFile write FOnBeforeStartFile;
    property OnAfterStartFile : TAfterStartFile  read FOnAfterStartFile  write FOnAfterStartFile;

  end;

procedure Register;

implementation

uses
  System.SysUtils, math,uOpenOfficeHelper, uOpenOfficeCollors;

procedure Register;
begin
  RegisterComponents('DinosOffice', [TOpenOffice_calc]);
end;

procedure TOpenOffice_calc.ValidateSheetName;
var
  LCID: LangID;
  Language: array [0 .. 100] of char;
begin

  LCID := GetSystemDefaultLangID;

  if SheetName.Trim.IsEmpty then
  begin

    VerLanguageName(LCID, Language, 100);

    if pos('Português', String(Language)) > 0 then
      SheetName := DefaultNewSheetNamePT
    else
      SheetName := DefaultNewSheetNameEn;
  end;
end;

function TOpenOffice_calc.SetValue(aCellNumber: integer; aCollName: string; aValue: variant; TypeValue: TTypeValue; Wrapped: boolean): TOpenOffice_calc;
var
  map: string;
begin
  if aCellNumber = 0 then
    aCellNumber := 1;

  map := aCollName + aCellNumber.ToString;
  objCell := objSCalc.getCellRangeByName(map);

  if  assigned(OnBeforeSetValue) then
    OnBeforeSetValue(self);

  if TypeValue = ftString then
  begin
    objCell.IsTextWrapped := false;

    if Wrapped then
      objCell.IsTextWrapped := True;

    objCell.setString(aValue);
  end
  else
    objCell.SetValue(aValue);

  Result := self;

  if  assigned(OnAfterSetValue) then
    OnAfterSetValue(self);
end;

constructor TOpenOffice_calc.Create(AOwner: TComponent);
begin
  inherited;
  Fields.setArrayFieldsSheet;
end;

procedure TOpenOffice_calc.SetSheetName(const Value: string);
begin
  FSheetName := Value;
end;

procedure TOpenOffice_calc.DataSetToSheet(const aCds: TClientDataSet);
var idx,idxFields : integer;
    lTypeVl : TTypeValue;
begin
  aCds.DisableControls;
  try
    //Create header
    for idx := 0 to pred(aCds.Fields.Count) do
      SetValue(0,FFields.arrFields[idx],aCds.Fields[idx].DisplayName)
      .setBold(true)
      .setBorder([bAll], opBlack)
      .changeFont('Liberation Sans',11)
      .setColor(opWhite,opSoftGray);

      aCds.First;
      while not aCds.Eof do
      begin
        for idxFields := 0 to pred(aCds.Fields.Count) do
        begin
          if (aCds.Fields[idx] is TCurrencyField) or
             (aCds.Fields[idx] is TIntegerField)  or
             (aCds.Fields[idx] is TFloatField)    or
             (aCds.Fields[idx] is TNumericField)  then
            lTypeVl := ftNumeric
           else
             lTypeVl := ftString;

          SetValue(aCds.RecNo +1, FFields.arrFields[idxFields],aCds.Fields[idxFields].Value, lTypeVl)
          .setBorder([bAll], opBlack);
        end;
        aCds.Next;
      end;
  finally
     aCds.EnableControls;
  end;
end;

destructor TOpenOffice_calc.Destroy;
begin
  inherited;
end;

function TOpenOffice_calc.GetValue(aCellNumber: integer; aCollName: String) : TOpenOffice_calc;
var
  map: string;
begin
  if assigned(onBeforeGetValue) then
    onBeforeGetValue(self);


  map := aCollName + aCellNumber.ToString;
  objCell := objSCalc.getCellRangeByName(map);
  Value := VarToStr(objCell.String);

  Result := self;

  if assigned(onAfterGetValue) then
    onAfterGetValue(self);
end;

procedure TOpenOffice_calc.positionSheetByName(aSheetName: string);
begin
  objSCalc := objDocument.Sheets.getByName(aSheetName);
end;

procedure TOpenOffice_calc.addNewSheet(aSheetName: string; aPosition: integer);
begin
  objDocument.Sheets.insertNewByName(aSheetName, aPosition);
  objSCalc := objDocument.Sheets.getByName(aSheetName);
end;

function TOpenOffice_calc.setFormula(aCellNumber: integer; aCollName: string;
  aFormula: string): TOpenOffice_calc;
var
  map: string;
begin
  map := aCollName + aCellNumber.ToString;
  objCell := objSCalc.getCellByPosition(Fields.getIndex(aCollName),
    aCellNumber);
  objCell.Formula := (aFormula);

  Result := self;
end;

procedure TOpenOffice_calc.startSheet;
begin

  if Assigned( FOnBeforeStartFile) then
    FOnBeforeStartFile(self);

  if URlFile.Trim.IsEmpty then
    URlFile := NewFile[integer(TpCalc)];

  ValidateSheetName;
  LoadDocument(SheetName);

  if objDocument.Sheets.hasByName(SheetName) then
    objSCalc := objDocument.Sheets.getByName(SheetName)
  else
  begin
    objSCalc := objDocument.createInstance('com.sun.star.sheet.Spreadsheet');
    objDocument.Sheets.insertByName(SheetName, objSCalc);
  end;

  if Assigned( FOnAfterStartFile) then
     FOnAfterStartFile(self);
end;



{ TFieldsSheet }

function TFieldsSheet.getField(aIndex: integer): string;
begin
  Result := arrFields[aIndex];
end;

function TFieldsSheet.getIndex(aNameField: String): integer;
var
  i: integer;
begin
  Result := 0;

  for i := 0 to High(arrFields) do
    if arrFields[i] = aNameField then
    begin
      Result := i;
      exit;
    end;
end;

procedure TFieldsSheet.setArrayFieldsSheet;
begin
  SetLength(arrFields, 24);

  arrFields[0] := 'A';
  arrFields[1] := 'B';
  arrFields[2] := 'C';
  arrFields[3] := 'D';
  arrFields[4] := 'E';
  arrFields[5] := 'F';
  arrFields[6] := 'G';
  arrFields[7] := 'H';
  arrFields[8] := 'I';
  arrFields[9] := 'J';
  arrFields[10] := 'K';
  arrFields[11] := 'L';
  arrFields[12] := 'M';
  arrFields[13] := 'N';
  arrFields[14] := 'O';
  arrFields[15] := 'P';
  arrFields[16] := 'Q';
  arrFields[17] := 'R';
  arrFields[18] := 'S';
  arrFields[19] := 'T';
  arrFields[20] := 'U';
  arrFields[21] := 'V';
  arrFields[22] := 'X';
  arrFields[23] := 'Y';
  arrFields[24] := 'Z';
end;

end.
