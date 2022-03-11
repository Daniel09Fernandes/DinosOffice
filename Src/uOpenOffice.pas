{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOffice.pas }
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

unit uOpenOffice;

interface

uses
  System.Classes, DB, ActiveX,
  dbWeb, ComObj, XMLDoc, XMLIntf, Vcl.Dialogs, System.Variants, Windows, uOpenOfficeEvents, uOpenOfficeSetPrinter;

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


  TOpenOffice = class(TComponent)
  private const
    NewSheet = 'private:factory/scalc';
    DefaultNewSheetNamePT = 'Planilha1';
    DefaultNewSheetNameEn = 'Sheet1';
  var
    //--------events------//
    FOnBeforeStartSheet: TBeforeStartSheet;
    FOnAfterStartSheet : TAfterStartSheet;
    FOnBeforeCloseSheet: TBeforeCloseSheet;
    FOnAfterCloseSheet : TAfterCloseSheet;
    FOnBeforePrint     : TBeforePrint;

    FOnAfterGetValue   : TAfterGetValue;
    FOnBeforeGetValue  : TBeforeGetValue;
    FOnAfterSetValue   : TAfterSetValue;
    FOnBeforeSetValue  : TBeforeSetValue;

    //--------------------//
    FFields: TFieldsSheet;
    FURlFile: string;
    FSheetName: string;
    FSetPrinter : TSetPrinter;
    procedure SetURlFile(const Value: string);
    procedure SetSheetName(const Value: string);
    function convertFilePathToUrlFile(aFilePath: string): string;
    procedure ValidateSheetName;
    procedure inicialization;
  protected
    { Protected declarations }
    // LibreOffice
      objCoreReflection, objDesktop, objServiceManager, objDocument, objTable,
      objCell, Charts: OleVariant;
  public
  var
    Value: string;
    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
    procedure startSheet;
    procedure closeSheet;
    procedure positionSheetByName(aSheetName: string);
    procedure addNewSheet(aSheetName: string; aPosition: integer);
    procedure saveSheet(aFileName: String = '');
    procedure print;
    function setFormula(aCellNumber: integer; aCollName: string; aFormula: string): TOpenOffice;
    function SetValue(aCellNumber: integer; aCollName: string; aValue: variant; TypeValue: TTypeValue = ftString; Wrapped: boolean = false): TOpenOffice;
    function GetValue(aCellNumber: integer; aCollName: String): TOpenOffice;

  published
    property ServicesManager: OleVariant read objServiceManager;
    property Cell: OleVariant read objCell write objCell;
    property Table: OleVariant read objTable;
    property Fields: TFieldsSheet read FFields;

    property URlFile: string read FURlFile write SetURlFile;
    property SheetName: string read FSheetName write SetSheetName;

    //---------events-----------//
    property OnBeforeStartSheet: TBeforeStartSheet read FOnBeforeStartSheet write FOnBeforeStartSheet;
    property OnAfterStartSheet : TAfterStartSheet  read FOnAfterStartSheet  write FOnAfterStartSheet;
    property OnBeforeCloseSheet: TBeforeCloseSheet read FOnBeforeCloseSheet write FOnBeforeCloseSheet;
    property OnAfterCloseSheet : TAfterCloseSheet  read FOnAfterCloseSheet  write FOnAfterCloseSheet;
    property OnBeforePrint     : TBeforePrint      read FOnBeforePrint      write FOnBeforePrint;
    property OnBeforeGetValue  : TBeforeGetValue   read FOnBeforeGetValue   write FOnBeforeGetValue;
    property OnAfterGetValue   : TAfterGetValue    read FOnAfterGetValue    write FOnAfterGetValue;
    property OnBeforeSetValue  : TBeforeSetValue   read FOnBeforeSetValue   write FOnBeforeSetValue;
    property OnAfterSetValue   : TAfterSetValue    read FOnAfterSetValue    write FOnAfterSetValue;

  end;

procedure Register;

implementation

uses
  System.SysUtils, math;

procedure Register;
begin
  RegisterComponents('DinosOffice', [TOpenOffice]);
end;

function TOpenOffice.SetValue(aCellNumber: integer; aCollName: string;
  aValue: variant; TypeValue: TTypeValue; Wrapped: boolean): TOpenOffice;
var
  map: string;
begin
  map := aCollName + aCellNumber.ToString;
  objCell := objTable.getCellRangeByName(map);

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

procedure TOpenOffice.closeSheet;
begin
  if Assigned( FOnBeforeCloseSheet) then
    FOnBeforeCloseSheet(self);

  objDocument.close(True);

  if Assigned( FOnAfterCloseSheet) then
    FOnAfterCloseSheet(self);
end;

function TOpenOffice.convertFilePathToUrlFile(aFilePath: string): string;
begin
  if (pos('FILE', UpperCase(aFilePath)) <= 0) then
  begin
    aFilePath := StringReplace(aFilePath, '\', '/', [rfReplaceAll]);
    aFilePath := 'file:///' + aFilePath;
  end;
  Result := aFilePath;
end;

constructor TOpenOffice.Create(AOwner: TComponent);
begin
  inherited;
  inicialization;
  Fields.setArrayFieldsSheet;
  FSetPrinter := TSetPrinter.create(nil);
end;

procedure TOpenOffice.SetSheetName(const Value: string);
begin
  FSheetName := Value;
end;

destructor TOpenOffice.Destroy;
begin
  objCoreReflection := Unassigned;
  objDesktop := Unassigned;
  objServiceManager := Unassigned;
  objDocument := Unassigned;
  objTable := Unassigned;
  objCell := Unassigned;
  FSetPrinter.Free;
  inherited;
end;

function TOpenOffice.GetValue(aCellNumber: integer; aCollName: String)
  : TOpenOffice;
var
  map: string;
begin
  if assigned(onBeforeGetValue) then
    onBeforeGetValue(self);


  map := aCollName + aCellNumber.ToString;
  objCell := objTable.getCellRangeByName(map);
  Value := VarToStr(objCell.String);

  Result := self;

  if assigned(onAfterGetValue) then
    onAfterGetValue(self);
end;

procedure TOpenOffice.inicialization;
begin
  try
    // Libre office
    objServiceManager := CreateOleObject('com.sun.star.ServiceManager');
    objCoreReflection := objServiceManager.createInstance
      ('com.sun.star.reflection.CoreReflection');
    objDesktop := objServiceManager.createInstance
      ('com.sun.star.frame.Desktop');
  except
    messageDlg('Erro(pt-Br):  Instale o LibreOffice para usar o sistema' + #13 +
      #13 + 'Error(En)  :  install  the LibreOffice to use the system',
      TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], 0);
  end;
end;

procedure TOpenOffice.positionSheetByName(aSheetName: string);
begin
  objTable := objDocument.Sheets.getByName(aSheetName);
end;

procedure TOpenOffice.print;
var PaperSize: variant;
    printerProperties : array[0..3] of variant;
begin

  if Assigned(onBeforePrint) then
  begin
    onBeforePrint(self, FSetPrinter);

    PaperSize := objServiceManager.Bridge_GetStruct('com.sun.star.awt.Size');

    PaperSize.Width  := FSetPrinter.PaperSize_Width;
    PaperSize.Height := FSetPrinter.PaperSize_Height;

    PrinterProperties[0] := objServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    PrinterProperties[0].Name  :='Name';
    PrinterProperties[0].Value := FSetPrinter.PrinterName;


    PrinterProperties[1] := objServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    PrinterProperties[1].Name  :='PaperSize';
    PrinterProperties[1].Value := PaperSize;

    PrinterProperties[2] := objServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    PrinterProperties[2].Name  :='Pages';
    PrinterProperties[2].Value := FSetPrinter.Pages;

    objDocument.Printer :=  VarArrayOf(PrinterProperties);

    objDocument.print(VarArrayOf(PrinterProperties));
  end
  else
    objDocument.print(VarArrayOf([]));
end;

procedure TOpenOffice.addNewSheet(aSheetName: string; aPosition: integer);
begin
  objDocument.Sheets.insertNewByName(aSheetName, aPosition);
  objTable := objDocument.Sheets.getByName(aSheetName);
end;

procedure TOpenOffice.saveSheet(aFileName: String);
begin
  aFileName := convertFilePathToUrlFile(aFileName);

  if aFileName.Trim.IsEmpty then
    aFileName := FURlFile;

  objDocument.storeAsURL(aFileName, VarArrayOf([]));
end;

function TOpenOffice.setFormula(aCellNumber: integer; aCollName: string;
  aFormula: string): TOpenOffice;
var
  map: string;
begin
  map := aCollName + aCellNumber.ToString;
  objCell := objTable.getCellByPosition(Fields.getIndex(aCollName),
    aCellNumber);
  objCell.Formula := (aFormula);

  Result := self;
end;

procedure TOpenOffice.SetURlFile(const Value: string);
begin
  FURlFile := Value;

  if FURlFile.Trim.IsEmpty then
    exit;

  FURlFile := convertFilePathToUrlFile(FURlFile);
end;

procedure TOpenOffice.startSheet;
begin

  if Assigned( FOnBeforeStartSheet) then
    FOnBeforeStartSheet(self);


  if FURlFile.Trim.IsEmpty then
    FURlFile := NewSheet;

  ValidateSheetName;
  objDocument := objDesktop.loadComponentFromURL(URlFile, SheetName, 0,
    VarArrayOf([]));

  if objDocument.Sheets.hasByName(SheetName) then
    objTable := objDocument.Sheets.getByName(SheetName)
  else
  begin
    objTable := objDocument.createInstance('com.sun.star.sheet.Spreadsheet');
    objDocument.Sheets.insertByName(SheetName, objTable);
  end;

  if Assigned( FOnAfterStartSheet) then
     FOnAfterStartSheet(self);
end;

procedure TOpenOffice.ValidateSheetName;
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
