{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOffice_writer.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }
unit uOpenOffice_writer;

interface

uses
  System.Classes, DB, ActiveX, System.Generics.Collections,
  FireDAC.Comp.Client, Datasnap.DBClient,
  dbWeb, ComObj, XMLDoc, XMLIntf, Vcl.Dialogs, System.Variants, Windows,
  uOpenOffice,
  uOpenOfficeEvents,
  uOpenOfficeSetPrinter,
  uOpenOfficeCollors,
  uOpenOffice.Writer.Table.Model;

type

  TTableWriter = TTableWriter_;
  TLineFieldValue = TLineFieldValue_;
  TOpenOffice_writer = class(TOpenOffice)
  private
    objTextCursor, oText, oCursor: variant;
    FDocName,
    FValueText: string;
    FBoldActive,
    changeForDispatcher: boolean;
    Alphabet: array [0..25] of string;
    FValue: string;
    procedure SetDocName(const Value: string);
  public
  var
    PropsText: array [0 .. 4] of variant;
    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
    procedure startDoc;
    function gotoEndOfSentence : TOpenOffice_writer;
    function setValue(const aText: string): TOpenOffice_writer;
    function getValue: string;
    function CreateTable(ATable: TTableWriter): TOpenOffice_writer; overload;
    function CreateTable(ATable: TFdMemTable): TOpenOffice_writer; overload;
    function CreateTable(ATable: TClientDataset): TOpenOffice_writer; overload;

    property BoldActive : boolean read FBoldActive write FBoldActive;
    property Cursor : variant read oCursor;
    property ValueText : string read FValueText;
    property Value: string read FValue write FValue;
  published
    property ServicesManager: OleVariant read objServiceManager;
    property DocName: string read FDocName write SetDocName;
  end;

  procedure Register;

implementation

{ TOpenOffice_writer }
uses
  System.SysUtils, System.Math;

procedure Register;
begin
  RegisterComponents('DinosOffice', [TOpenOffice_writer]);
end;

constructor TOpenOffice_writer.Create(AOwner: TComponent);
begin
  inherited;
  FBoldActive := false;
  Alphabet[0] := 'A';
  Alphabet[1] := 'B';
  Alphabet[2] := 'C';
  Alphabet[3] := 'D';
  Alphabet[4] := 'E';
  Alphabet[5] := 'F';
  Alphabet[6] := 'G';
  Alphabet[7] := 'H';
  Alphabet[8] := 'I';
  Alphabet[9] := 'J';
  Alphabet[10]:= 'K';
  Alphabet[11]:= 'L';
  Alphabet[12]:= 'M';
  Alphabet[13]:= 'N';
  Alphabet[14]:= 'O';
  Alphabet[15]:= 'P';
  Alphabet[16]:= 'Q';
  Alphabet[17]:= 'R';
  Alphabet[18]:= 'S';
  Alphabet[19]:= 'T';
  Alphabet[20]:= 'U';
  Alphabet[21]:= 'V';
  Alphabet[22]:= 'X';
  Alphabet[23]:= 'W';
  Alphabet[24]:= 'Y';
  Alphabet[25]:= 'Z';
end;

function TOpenOffice_writer.CreateTable(ATable: TTableWriter): TOpenOffice_writer;
var
  oTextTable: Variant;
  oCursor: Variant;
  key, line, linesCol, numRows, numCols: Integer;
  lColl, lOutVal: string;
  lPair: TPair<TTableRowFields, TTableLinesValues>;
  lLines: TList<TLineFieldValue>;
begin
  if not Assigned(ATable) then
    raise Exception.Create('Tabela não carregado. Execute startDoc primeiro.');

  ATable.TableLines.TrimExcess;
  ATable.Table.Add(ATable.TableFields, ATable.TableLines);

  numRows := ATable.TableLines.Count +1;
  numCols := ATable.TableFields.Count;

  oCursor := objDocument.getCurrentController.getViewCursor;

  oTextTable := objDocument.createInstance('com.sun.star.text.TextTable');
  oTextTable.initialize(numRows, numCols);
  linesCol := 0;
  objDocument.getText.insertTextContent(oCursor, oTextTable, False);
  for lPair in ATable.Table do
  begin
    for key := 0 to lPair.Key.Count -1 do //Header
    begin
      if lPair.Key[key] = '' then
        continue;

      oTextTable.getCellByName(Alphabet[key]+ '1').String := lPair.Key[key];
      if ATable.Table.TryGetValue(lPair.Key, lLines) then   //Fields
      begin
        if linesCol > pred(lLines.Count) then
          Break;

        line := 0;
        for lOutVal in lLines[linesCol].Line do
        begin
          oTextTable.getCellByName(Alphabet[line] + IntToStr(linesCol+2)).String := lOutVal;
          Inc(line);
        end;
      end;
      Inc(linesCol);
    end;
  end;
  Result := Self;
end;

function TOpenOffice_writer.CreateTable(ATable: TFdMemTable): TOpenOffice_writer;
var
  lFieldIdx,
  idxFields,
  lRecNo: Integer;
  lTable: TTableWriter;
  lLog: Boolean;
begin
  ATable.DisableControls;
  lTable := TTableWriter.New;
  try
    lLog := ATable.LogChanges;
    lRecNo := ATable.RecNo;
    ATable.LogChanges := False;

    for lFieldIdx := 0 to ATable.Fields.Count-1 do
       lTable.TableFields.Add(ATable.Fields[lFieldIdx].DisplayName);

    ATable.First;
    while not ATable.Eof do
    begin
      for idxFields := 0 to pred(ATable.Fields.Count) do
        lTable.Lines.Line.add(ATable.Fields[idxFields].Value);

      if ATable.RecNo < ATable.RecordCount then
      begin
        lTable.Lines      := TTableWriter.AddNewLines;
        lTable.TableLines.Add(lTable.Lines);
      end;
      ATable.Next;
    end;

    CreateTable(lTable);
  finally
    ATable.EnableControls;
    ATable.LogChanges := lLog;
    ATable.RecNo := lRecNo;
    lTable.Free;
  end;
end;

function TOpenOffice_writer.CreateTable(ATable: TClientDataset): TOpenOffice_writer;
var
  lFieldIdx,
  idxFields,
  lRecNo: Integer;
  lTable: TTableWriter;
  lLog: Boolean;
  lLines: TLineFieldValue;
begin
  ATable.DisableControls;
  lTable := TTableWriter.New;
  try
    lLog := ATable.LogChanges;
    lRecNo := ATable.RecNo;
    ATable.LogChanges := False;

    for lFieldIdx := 0 to ATable.Fields.Count-1 do
       lTable.TableFields.Add(ATable.Fields[lFieldIdx].DisplayName);

    ATable.First;

    while not ATable.Eof do
    begin
      for idxFields := 0 to pred(ATable.Fields.Count) do
        lTable.Lines.Line.add(ATable.Fields[idxFields].Value);

      if ATable.RecNo < ATable.RecordCount then
      begin
        lTable.Lines      := TTableWriter.AddNewLines;
        lTable.TableLines.Add(lTable.Lines);
      end;
      ATable.Next;
    end;

    CreateTable(lTable);
  finally
    ATable.EnableControls;
    ATable.LogChanges := lLog;
    ATable.RecNo := lRecNo;
    lTable.Free;
  end;
end;

destructor TOpenOffice_writer.Destroy;
var
  i: integer;
begin
  for i := 0 to High(propsText) do
    propsText[i] := Unassigned;

  inherited;
end;

function TOpenOffice_writer.getValue: string;
begin
  if assigned(onBeforeGetValue) then
    onBeforeGetValue(self);

  Result := VarToStr(oText.String);

  if assigned(onAfterGetValue) then
    onAfterGetValue(self);
end;


function TOpenOffice_writer.gotoEndOfSentence: TOpenOffice_writer;
begin
  objTextCursor.jumpToEndOfPage;
  oText := objTextCursor.Text;
  oCursor := oText.CreateTextCursor;
  oCursor.gotoEnd(true);
  Result := self;
end;

function TOpenOffice_writer.setValue(const aText: string): TOpenOffice_writer;
begin
  if assigned(onBeforeSetValue) then
    onBeforeSetValue(self);

  oText := objTextCursor.Text;
  oText.InsertString(objTextCursor, '', true);

  oCursor := oText.CreateTextCursor;

  changeForDispatcher := true;

  oText.InsertString(objTextCursor, aText, true);
  oText.InsertControlCharacter(oCursor, 10, true);

  FValueText  := aText;
  Result      := self;

  if assigned(onAfterSetValue) then
    onAfterSetValue(self);
end;

procedure TOpenOffice_writer.startDoc;
begin
  if URlFile.Trim.IsEmpty then
    URlFile := NewFile[integer(TpWriter)];

  LoadDocument(DocName); // cria a instancia do objDocument
  objTextCursor := objDocument.getCurrentController.getViewCursor;

  propsText[0] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[1] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[2] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[3] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[4] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');

  objWriter := objDocument.getCurrentController.getFrame;
  objDispatcher := objServiceManager.createInstance
    ('com.sun.star.frame.DispatchHelper');
end;

procedure TOpenOffice_writer.SetDocName(const Value: string);
begin
  FDocName := Value;
end;

{ TTableWriter }


end.
