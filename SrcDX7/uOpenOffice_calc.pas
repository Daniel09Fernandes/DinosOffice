{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOffice_calc.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MPL/GPL/LGPL three license - see LICENSE.md}

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
  Classes, DB, ActiveX, uOpenOffice, uOpenOfficeCollors,
  dbWeb, ComObj, XMLDoc, XMLIntf, Dialogs, Variants,
  Windows, uOpenOfficeEvents, DBClient;

type

  TTypeValue = (ftString, ftNumeric);
  TTypeChart = (ctDefault, ctVertical, ctPie, ctLine);
  TBorder = (bAll, bLeft, bRight, bBottom, bTop);

  TBoderSheet = set of TBorder;

  { STANDARD : é o alinhamento padrão tanto para números como para textos, sendo a esqueda para as strings e a direita para os números;
    LEFT : o conteúdo é alinhado no lado esquerdo da célula;
    CENTER : o conteúdo é alinhado no centro da célula;
    RIGHT : o conteúdo é alinhado no lado direito da célula;
    BLOCK : o conteúdo é alinhando em relação ao comprimento da célula;
    REPEAT : o conteúdo é repetido dentro da célula para preenchê-la. }
  THoriJustify = (fthSTANDARD, fthLEFT, fthCENTER, fthRIGHT, fthBLOCK,
    fthREPEAT);
  { STANDARD : é o valor usado como padrão;
    TOP : o conteúdo da célula é alinhado pelo topo;
    CENTER : o conteúdo da célula é alinhado pelo centro;
    BOTTOM : o conteúdo da célula é alinhado pela base. }
  TVertJustify = (ftvSTANDARD, ftvTOP, ftvCENTER, ftvBOTTOM);

  TFields = record
    arrFields: array of string;
  end;

  TOpenOffice_calc = class(TOpenOffice)
  private
    FFields : TFields;
    DefaultNewSheetNamePT : string;
    DefaultNewSheetNameEn : string;

    //--------events------//
    FOnBeforeStartFile: TBeforeStartFile;
    FOnAfterStartFile : TAfterStartFile;
    //--------------------//
    FSheetName: string;
    procedure ValidateSheetName;
    procedure SetSheetName(const Value: string);
    procedure setArrayFieldsSheet;
  public
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
    procedure addChart(typeChart: TTypeChart; StartRow, EndRow: Integer; StartColumn, EndColumn, ChartName: string; PositionSheet: Integer);
    function setBorder(borderPosition: TBoderSheet; opColor: TOpenColor; RemoveBorder: boolean = false) : TOpenOffice_calc;
    function changeFont(aNameFont: string; aHeight: Integer): TOpenOffice_calc;
    function changeJustify(aTypeHori: THoriJustify; aTypeVert: TVertJustify) : TOpenOffice_calc;
    function setColor(aFontColor, aBackgroud: TOpenColor): TOpenOffice_calc;
    function setBold(aBold: boolean): TOpenOffice_calc;
    function SetUnderline(aUnderline: boolean): TOpenOffice_calc;
    function CountRow: Integer;
    function CountCell: Integer;
    function HoryJustifyToInteger(pValue:THoriJustify):integer;
    function VertJustifyToInteger(pValue:TVertJustify):integer;
    function getField(aIndex: integer): string;
    function getIndex(aNameField: String): integer;
  published
    property ServicesManager: OleVariant read objServiceManager;
    property Cell: OleVariant read objCell write objCell;
    property Table: OleVariant read objSCalc;
    property Fields: TFields read FFields;
    property SheetName: string read FSheetName write SetSheetName;

    //---------events-----------//
    property OnBeforeStartFile: TBeforeStartFile read FOnBeforeStartFile write FOnBeforeStartFile;
    property OnAfterStartFile : TAfterStartFile  read FOnAfterStartFile  write FOnAfterStartFile;

  end;

procedure Register;

implementation

uses
  SysUtils, math,uOpenOfficeHelper;

procedure Register;
begin
  RegisterComponents('DinosOffice', [TOpenOffice_calc]);
end;

function TOpenOffice_calc.HoryJustifyToInteger(pValue:THoriJustify):Integer;
begin
case pValue of
    fthSTANDARD:
      result := 0;
    fthLEFT:
      result := 1;
    fthCENTER:
      result := 2;
    fthRIGHT:
      result := 3;
    fthBLOCK:
      result := 4;
    fthREPEAT:
      result := 5;
  end;
end;

function TOpenOffice_calc.VertJustifyToInteger(pValue:TVertJustify):Integer;
begin
  case pValue of
    ftvSTANDARD:
      result := 0;
    ftvTOP:
      result := 1;
    ftvCENTER:
      result := 2;
    ftvBOTTOM:
      result := 3;
  end;
end;

procedure TOpenOffice_calc.ValidateSheetName;
var
  LCID: LangID;
  Language: array [0 .. 100] of char;
begin

  LCID := GetSystemDefaultLangID;

  if Trim(SheetName) = '' then
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

  map := aCollName + IntToStr(aCellNumber);
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
   setArrayFieldsSheet;
   DefaultNewSheetNamePT := 'Planilha1';
   DefaultNewSheetNameEn := 'Sheet1';
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


  map := aCollName + IntToStr(aCellNumber);
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
  map := aCollName + intToStr(aCellNumber);
  objCell := objSCalc.getCellByPosition(getIndex(aCollName),
    aCellNumber);
  objCell.Formula := (aFormula);

  Result := self;
end;

procedure TOpenOffice_calc.startSheet;
begin

  if Assigned( FOnBeforeStartFile) then
    FOnBeforeStartFile(self);

  if trim(URlFile) = '' then
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

 
procedure TOpenOffice_calc.addChart(typeChart: TTypeChart;
  StartRow, EndRow: Integer; StartColumn, EndColumn, ChartName: string;
  PositionSheet: Integer);
var
  Chart, Rect, sheet, cursor: OleVariant;
  RangeAddress: Variant;
  countChart: Integer;
begin
  countChart := 1;

  if Trim(ChartName) = '' then
    ChartName := 'MyChar_' + (StartColumn + intTostr(StartRow)) + ':' +
      (EndColumn + intTostr(EndRow));

  sheet := objDocument.Sheets.getByIndex(PositionSheet);
  // getByName(aCollName);
  Charts := sheet.Charts;

  while Charts.hasByName(ChartName) do
  begin
    ChartName := ChartName + '_' + intTostr(countChart);
    inc(countChart);
  end;

  Rect := objServiceManager.Bridge_GetStruct('com.sun.star.awt.Rectangle');
  RangeAddress := sheet.Bridge_GetStruct('com.sun.star.table.CellRangeAddress');

  Rect.Width := 12000;
  Rect.Height := 12000;
  Rect.X := 8000 * countChart + 1;
  Rect.Y := 1000;

  RangeAddress.sheet := PositionSheet;
  RangeAddress.StartColumn := getIndex(StartColumn);
  RangeAddress.StartRow := StartRow;
  RangeAddress.EndColumn := getIndex(EndColumn);
  RangeAddress.EndRow := EndRow;

  Charts.addNewByName(ChartName, Rect, VarArrayOf(RangeAddress), true, true);

  if typeChart <> ctDefault then
  begin
    Chart := Charts.getByName(ChartName).embeddedObject;
    Chart.Title.String := ChartName;
    case typeChart of
      ctVertical:
        Chart.Diagram.Vertical := true;
      ctPie:
        begin
          Chart.Diagram := Chart.createInstance
            ('com.sun.star.chart.PieDiagram');
          Chart.HasMainTitle := true;
        end;
      ctLine:
        begin
          Chart.Diagram := Chart.createInstance
            ('com.sun.star.chart.LineDiagram');
        end;
    end;
  end;

end;

function TOpenOffice_calc.changeFont(aNameFont: string; aHeight: Integer)
  : TOpenOffice_calc;
begin
  // Cell := Table.getCellRangeByName(aCollName+aCellNumber.ToString);
  Cell.CharFontName := aNameFont;
  Cell.CharHeight := inttostr(aHeight);
  result := self;
end;

function TOpenOffice_calc.changeJustify(aTypeHori: THoriJustify;
  aTypeVert: TVertJustify): TOpenOffice_calc;
begin
  Cell.HoriJustify := HoryJustifyToInteger(aTypeHori);
  Cell.VertJustify := VertJustifyToInteger(aTypeVert);
  result := self;
end;

function TOpenOffice_calc.CountRow: Integer;
var
  FRow, FCountRow: Integer;
  FCountBlank: Integer;
  FBreak, allBlank: boolean;
  I: Integer;
begin
  FBreak := false;
  FRow := 1;
  FCountRow := 0;
  FCountBlank := 0;

  while not FBreak do
  begin
    for I := 0 to 21 do
    begin
      if GetValue(FRow, getField(I)).Value = '' then
      begin
        allBlank := true;
      end
      else
      begin
        if FCountBlank > 0 then // An empty column behind a valued column
          FCountRow := FCountRow + FCountBlank;

        allBlank := false;
        FCountBlank := 0;

        inc(FCountRow);
        break;
      end;
    end;
    inc(FRow);

    if FCountBlank = 50 then
      FBreak := true;

    if allBlank then
      inc(FCountBlank);

  end;
  result := FCountRow;
end;

function TOpenOffice_calc.CountCell: Integer;
var
  FCell, FCountCell, FCountBlank: Integer;
  I: Integer;
  allBlank: boolean;
begin
  FCell := 1;
  FCountCell := 0;
  FCountBlank := 0;

  for I := 0 to 21 do
  begin
    for FCell := 1 to 10 do
    begin
      if GetValue(FCell, getField(I)).Value <> '' then
      begin

        if FCountBlank > 0 then
          FCountCell := FCountCell + FCountBlank;

        allBlank := false;

        inc(FCountCell);
        break;
      end
      else
        allBlank := true;
    end;

    if FCountBlank = 10 then
    begin
      FCountBlank := 0;
      break;
    end;

    if allBlank then
      inc(FCountBlank);
  end;

  result := FCountCell;
end;

function TOpenOffice_calc.seTBorder(borderPosition: TBoderSheet; opColor: TOpenColor; RemoveBorder: boolean): TOpenOffice_calc;
var
  border: Variant;
  settings: Variant;
begin
  border := ServicesManager.createInstance('com.sun.star.reflection.CoreReflection');
  border.forName('com.sun.star.table.BorderLine2').createObject(settings);

 if not RemoveBorder then
  begin
    settings.Color := opColor;
    settings.InnerLineWidth := 20;
    settings.LineDistance := 60;
    settings.LineWidth := 2;
    settings.OuterLineWidth := 20;
  end else
  begin
    settings.Color := 0;
    settings.InnerLineWidth := 0;
    settings.LineDistance := 0;
    settings.LineWidth := 0;
    settings.OuterLineWidth := 0;
  end;

  if bAll in borderPosition then
  begin
    Cell.TopBorder := settings;
    Cell.LeftBorder := settings;
    Cell.RightBorder := settings;
    Cell.BottomBorder := settings;
  end;

  if bTop in borderPosition then
    Cell.TopBorder := settings;

  if bLeft in borderPosition then
    Cell.LeftBorder := settings;

  if bRight in borderPosition then
    Cell.RightBorder := settings;

  if bBottom in borderPosition then
    Cell.BottomBorder := settings;

  result := self;
end;

function TOpenOffice_calc.setColor(aFontColor, aBackgroud: TOpenColor)
  : TOpenOffice_calc;
begin
  Cell.CharColor := aFontColor;
  Cell.CellBackColor := aBackgroud;
  result := self;
end;

function TOpenOffice_calc.setBold(aBold: boolean): TOpenOffice_calc;
begin
  Cell.CharWeight := ifthen(aBold, 150, 0);
  result := self;
end;

function TOpenOffice_calc.SetUnderline(aUnderline: boolean): TOpenOffice_calc;
begin
  Cell.CharUnderline := ifthen(aUnderline, 1, 0);
  result := self;
end;


function TOpenOffice_calc.getField(aIndex: integer): string;
begin
  Result := FFields.arrFields[aIndex];
end;

function TOpenOffice_calc.getIndex(aNameField: String): integer;
var
  i: integer;
begin
  Result := 0;

  for i := 0 to High(FFields.arrFields) do
    if FFields.arrFields[i] = aNameField then
    begin
      Result := i;
      exit;
    end;
end;

procedure TOpenOffice_calc.setArrayFieldsSheet;
begin
  SetLength(FFields.arrFields, 25);
  with FFields do
  begin
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
  end
end;

end.
