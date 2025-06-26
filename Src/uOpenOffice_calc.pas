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
  System.Classes, data.DB, ActiveX, uOpenOffice, System.Generics.Collections,
  dbWeb, ComObj, XMLDoc, XMLIntf, Vcl.Dialogs, System.Variants,
  Windows, uOpenOfficeEvents, Datasnap.DBClient, System.SysUtils,
  FireDAC.Comp.Client;

type

  TTypeValue = (ftString, ftNumeric);
  //0; //1234,2
  //1; //1234
  //3; //1.234
  //4; //1.234,20
  //10; //123420%
  //11; //123420,00%
  //20; //R$ 1.234
  //21; //R$ 1.234,20
  //24; // 1.234,20 BRL
  //30; // DD/MM/YY
  //31; // DDDD DD/MM/YY
  //32; // MM/YY
  //33; // DD/MMMM
  //34; // MMMM
  //35; // TRIMESTRE
  //36; // DD/MM/YYYY
  //39; // DD MMMM YY
  //40; // DD MMMM YYYY
  TNumberMask = (Default, NoDecimalSeparator, WithDecimalSeparator = 3, WithDecimalSeparatorAndComma, Percent = 10, PercentWithComma, CurrencySymbolWithoutDecimal = 20,
                   CurrencySymbolWithDecimal, CurrencySuffix = 24, Date_dd_mm_yy = 30, Date_dddd_dd_mm_yyy,  Date_mm_yy, Date_dd_mmmm, Date_mmmm,
                   Date_QUARTER, Date_Default, Date_dd_mmmm_yy = 39, Date_dd_mmmm_yyyy);


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

   var
    //--------events------//
    FOnBeforeStartFile: TBeforeStartFile;
    FOnAfterStartFile : TAfterStartFile;
    //--------------------//
    FFields: TFieldsSheet;
    FSheetName: string;
    FNumberMask: TNumberMask;
    FValue: string;

    procedure ValidateSheetName;
    procedure SetSheetName(const Value: string);
  public
    function StartSheet: TOpenOffice_calc;
    function AddNewSheet(const aSheetName: string; aPosition: integer): TOpenOffice_calc;
   	function DatasetToSheet(const aCds : TClientDataSet): TOpenOffice_calc; overload;
    function DatasetToSheet(const aCds : TFDMemTable): TOpenOffice_calc; overload;
    function CallConversorPDFTOSheet: TOpenOffice_calc;
	function ExeThread(pProc : Tproc): TOpenOffice_calc;
    function PositionSheetByIndex(const aSheetIndex: integer): TOpenOffice_calc;
    function PositionSheetByName(const aSheetName: string):TOpenOffice_calc;
    function SetFormula(aCellNumber: integer; const aCollName: string; const aFormula: string): TOpenOffice_calc;
    function SetValue(aCellNumber: integer; const aCollName: string; aValue: variant; TypeValue: TTypeValue = ftString; Wrapped: boolean = false): TOpenOffice_calc;
    function GetValue(aCellNumber: integer; const aCollName: String) : TOpenOffice_calc;    
    function SheetToDataSet(const TabSheetName: String; TabSheetIndex: Integer = 0; IndexOfHeaderToFieldCds: Integer = 1): TClientDataSet;
    function TabSheetExists(ATabSheetName: string):Boolean;
  	function RemoveSheet(const aSheetName: string):TOpenOffice_calc; overload;
    function RemoveSheet(aSheetIndex: Integer):TOpenOffice_calc; overload;
    function GetSheetList: TDictionary<integer, string>;

    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
  published
    property SheetName: string read FSheetName write SetSheetName;
    property NumberMask: TNumberMask read FNumberMask write FNumberMask;
    property Fields: TFieldsSheet read FFields;
    property Value: string read FValue write FValue;

    //---------events-----------//
    property OnBeforeStartFile: TBeforeStartFile read FOnBeforeStartFile write FOnBeforeStartFile;
    property OnAfterStartFile : TAfterStartFile  read FOnAfterStartFile  write FOnAfterStartFile;
  end;

procedure Register;

implementation

uses
  math,
  uOpenOfficeHelper,
  uOpenOfficeCollors,
  uConvertPDFToSheet;

procedure Register;
begin
  RegisterComponents('DinosOffice', [TOpenOffice_calc]);
end;

function TOpenOffice_calc.CallConversorPDFTOSheet: TOpenOffice_calc;
var PdfToSheet : TConvertPDFToSheet;
begin
  PdfToSheet := TConvertPDFToSheet.create;
  try
    PdfToSheet.callConversor;
  finally
    freeAndNil(PdfToSheet);
    Result := Self;
  end;
end;

function TOpenOffice_calc.GetSheetList: TDictionary<integer, string>;
var
  I: Integer;
  lSheets: OleVariant;
begin
  Result := TDictionary<integer, string>.Create;

  lSheets := FobjDocument.Sheets;

  for I := 0 to lSheets.getCount - 1 do
    Result.Add(I, lSheets.getByIndex(I).getName);
end;

function TOpenOffice_calc.RemoveSheet(const aSheetName: string):TOpenOffice_calc;
begin
  if FobjDocument.Sheets.hasByName(aSheetName) then
    FobjDocument.Sheets.removeByName(aSheetName)
  else
    raise Exception.CreateFmt('A aba "%s" não existe no documento', [aSheetName]);
end;

function TOpenOffice_calc.RemoveSheet(aSheetIndex: Integer):TOpenOffice_calc;
var
  lSheetName: string;
begin
  if (aSheetIndex >= 0) and (aSheetIndex < FobjDocument.Sheets.getCount) then
  begin
    lSheetName := FobjDocument.Sheets.getByIndex(aSheetIndex).getName;
    FobjDocument.Sheets.removeByName(lSheetName);
  end
  else
    raise Exception.CreateFmt('Índice de aba inválido: %d', [aSheetIndex]);
end;

procedure TOpenOffice_calc.ValidateSheetName;
var
  LCID: LangID;
  Languages: array [0 .. 100] of char;
  Language : string;
begin

  LCID := GetSystemDefaultLangID;

  if SheetName.Trim.IsEmpty then
  begin

    VerLanguageName(LCID, Languages, 100);
    Language := String(Languages);

    if pos('Português', Language) > 0 then
      SheetName := DefaultNewSheetNamePT
    else
      SheetName := DefaultNewSheetNameEn;
  end;
end;

function TOpenOffice_calc.SetValue(aCellNumber: integer; const aCollName: string; aValue: variant; TypeValue: TTypeValue; Wrapped: boolean): TOpenOffice_calc;
var
  map: string;
begin
  if aCellNumber = 0 then
    aCellNumber := 1;

  map := aCollName + aCellNumber.ToString;
  FobjCell := FobjSCalc.getCellRangeByName(map);

  if  assigned(OnBeforeSetValue) then
    OnBeforeSetValue(self);

  if TypeValue = ftString then
  begin
    FobjCell.IsTextWrapped := false;

    if Wrapped then
      FobjCell.IsTextWrapped := True;

    FobjCell.setString(aValue);
  end
  else
  begin
    FobjCell.NumberFormat := Integer(NumberMask);
    FobjCell.SetValue(aValue);
  end;

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

function TOpenOffice_calc.DatasetToSheet(const aCds: TFDMemTable): TOpenOffice_calc;
var idx,idxFields : integer;
    lTypeVl : TTypeValue;
begin
  aCds.DisableControls;
  try
    //Create header
    for idx := 0 to pred(aCds.Fields.Count) do
      SetValue(0,Fields.arrFields[idx],aCds.Fields[idx].DisplayName)
      .setBold(true)
      .setBorder([bAll], opBlack)
      .setColor(opBlack,opSoftGray);

      aCds.First;
      while not aCds.Eof do
      begin
        for idxFields := 0 to pred(aCds.Fields.Count) do
        begin
          if (aCds.Fields[idxFields] is TCurrencyField) or
             (aCds.Fields[idxFields] is TIntegerField)  or
             (aCds.Fields[idxFields] is TFloatField)    or
             (aCds.Fields[idxFields] is TNumericField)  then
            lTypeVl := ftNumeric
           else
             lTypeVl := ftString;

          SetValue(aCds.RecNo +1, Fields.arrFields[idxFields],aCds.Fields[idxFields].Value, lTypeVl)
          .setBorder([bAll], opBlack);
        end;
        aCds.Next;
      end;
  finally
     aCds.EnableControls;
     Result := Self
  end;
end;

function TOpenOffice_calc.DatasetToSheet(const aCds: TClientDataSet): TOpenOffice_calc;
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
      .setColor(opBlack,opSoftGray);

      aCds.First;
      while not aCds.Eof do
      begin
        for idxFields := 0 to pred(aCds.Fields.Count) do
        begin
          if (aCds.Fields[idxFields] is TCurrencyField) or
             (aCds.Fields[idxFields] is TIntegerField)  or
             (aCds.Fields[idxFields] is TFloatField)    or
             (aCds.Fields[idxFields] is TNumericField)  then
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
     Result := Self;
  end;
end;

destructor TOpenOffice_calc.Destroy;
begin
  inherited;
end;

function TOpenOffice_calc.ExeThread(pProc: Tproc): TOpenOffice_calc;
begin
  HungThread.ExecProc := pProc;
  HungThread.Start;
  Result := Self;
end;

function TOpenOffice_calc.GetValue(aCellNumber: integer; const aCollName: String) : TOpenOffice_calc;
var
  map: string;
begin
  if assigned(onBeforeGetValue) then
    onBeforeGetValue(self);

  map := aCollName + aCellNumber.ToString;
  FobjCell := FobjSCalc.getCellRangeByName(map);
  Value := VarToStr(FobjCell.String);

  Result := self;

  if assigned(onAfterGetValue) then
    onAfterGetValue(self);
end;

function TOpenOffice_calc.positionSheetByName(const aSheetName: string):TOpenOffice_calc;
begin
  FobjSCalc := FobjDocument.Sheets.getByName(aSheetName);
  Result := self;
end;

function TOpenOffice_calc.positionSheetByIndex(const aSheetIndex: integer) :TOpenOffice_calc;
begin
  FobjSCalc := FobjDocument.Sheets.getByIndex(aSheetIndex);
  Result := self;
end;

function TOpenOffice_calc.addNewSheet(const aSheetName: string; aPosition: integer): TOpenOffice_calc;
begin
  FobjDocument.Sheets.insertNewByName(aSheetName, aPosition);
  FobjSCalc := FobjDocument.Sheets.getByName(aSheetName);
  Result := Self;
end;

function TOpenOffice_calc.setFormula(aCellNumber: integer; const aCollName: string;
  const aFormula: string): TOpenOffice_calc;
var
  map: string;
begin
  map := aCollName + aCellNumber.ToString;
  FobjCell := FobjSCalc.getCellByPosition(Fields.getIndex(aCollName), aCellNumber);
  FobjCell.FormulaLocal  := aFormula;
  Result := self;
end;

function TOpenOffice_calc.startSheet: TOpenOffice_calc;
begin
  if Assigned( FOnBeforeStartFile) then
    FOnBeforeStartFile(self);

  if URlFile.Trim.IsEmpty then
    URlFile := FNewFile[integer(TpCalc)];

  ValidateSheetName;
  LoadDocument(SheetName);

  if FobjDocument.Sheets.hasByName(SheetName) then
    FobjSCalc := FobjDocument.Sheets.getByName(SheetName)
  else
  begin
    FobjSCalc := FobjDocument.createInstance('com.sun.star.sheet.Spreadsheet');
    FobjDocument.Sheets.insertByName(SheetName, FobjSCalc);
  end;

  if Assigned( FOnAfterStartFile) then
     FOnAfterStartFile(self);

  Result := Self;
end;

function TOpenOffice_calc.TabSheetExists(ATabSheetName: string):Boolean;
begin
  Result := FobjDocument.Sheets.hasByName(ATabSheetName); 
end;

function TOpenOffice_calc.SheetToDataSet(const TabSheetName: String; TabSheetIndex: Integer = 0; IndexOfHeaderToFieldCds: Integer = 1): TClientDataSet;
var
  I, IdxField, lCountRow,
  lCountCell : Integer;
  lFieldName: string;
begin
  Result := TClientDataSet.Create(nil);
  try
    if not TabSheetName.trim.IsEmpty then
      positionSheetByName(TabSheetName)
    else
      positionSheetByIndex(TabSheetindex);

	  lCountCell := CountCell;
    for I := 0 to lCountCell -1 do
    begin
      lFieldName := GetValue(IndexOfHeaderToFieldCds, Fields.getField(I)).Value;
      if lFieldName.Trim.IsEmpty then
        Continue;

      Result.FieldDefs.Add(lFieldName, TFieldType.ftString, 30);
    end;

    Result.CreateDataSet;
    Result.DisableControls;
    Result.LogChanges := false;
	  lCountRow := CountRow;

    for I := (IndexOfHeaderToFieldCds+1) to lCountRow do
    begin
      Result.Append;
      for IdxField := 0 to pred(Result.FieldCount) do
        Result.Fields[IdxField] .AsString := GetValue(I,Fields.getField(IdxField)).Value;

      Result.Post;
    end;
  finally
    Result.EnableControls;   
  end;
end;

{ TFieldsSheet }

function TFieldsSheet.getField(aIndex: integer): string;
var DifIdx: double;
    Letter: String;
    Idx: Int64;
begin

  if (aIndex > 25) then //Multiplica as letras EX AA, AB, AAA, AAB...
  begin
     DifIdx := aIndex / 26;
     DifIdx := Trunc(DifIdx);

     SetLength(arrFields, aIndex+1);
     arrFields[aIndex] := '';

     Letter := arrFields[trunc(DifIdx-1)];

     if DifIdx = 0 then
       DifIdx := 1;

     DifIdx := DifIdx * 26;

     Idx := (aIndex - Trunc(DifIdx - Trunc(aIndex / 26) )) - Trunc(aIndex / 26);  //Reseta o alfabeto para cada range A, AA, AAA
     if Idx < 0 then
      Idx := Idx *-1;

     Letter := Letter + arrFields[Idx];

    arrFields[aIndex] := Letter;
  end;

  Result := arrFields[aIndex];
end;

function TFieldsSheet.getIndex(aNameField: String): integer;
var
  i, idx: integer;
  rep,aux, firstIdx,
  secondIdx : integer;
begin
  Result := 0;
  rep := 26;
  aux := 0;
  secondIdx := 0;
  aNameField := aNameField.ToUpper;
  if aNameField.length <= 1  then
  begin
    for i := 0 to High(arrFields) do
      if arrFields[i] = aNameField then
      begin
        Result := i;
        exit;
      end;
  end else
  begin
    for idx := 1 to aNameField.Length - 2 do
      rep := (rep * 26) + 26;

      //26*26 + 26

    for idx := 2 to aNameField.Length do
    begin
      for i := 0 to High(arrFields) do
          if arrFields[i] = aNameField[idx] then
            break;
    end;
    firstIdx  :=  getIndex(aNameField[1]) + 1;

    if  aNameField.Length = 3 then
    begin
      secondIdx := getIndex(aNameField[2]) + 27;
      aux := 26;
    end;

    Result := ( (firstIdx * 26) + (secondIdx * 26) + i) - aux;
  end;
end;

procedure TFieldsSheet.setArrayFieldsSheet;
begin
  SetLength(arrFields, 26);

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
  arrFields[22] := 'W';
  arrFields[23] := 'X';
  arrFields[24] := 'Y';
  arrFields[25] := 'Z';
end;

end.
