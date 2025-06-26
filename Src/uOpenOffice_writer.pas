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
    FobjTextCursor, FoText, FoCursor: variant;
    FDocName,
    FValueText: string;
    FBoldActive,
    changeForDispatcher: boolean;
    Alphabet: array [0..25] of string;
    procedure SetDocName(const Value: string);
  public
  var
    PropsText: array [0 .. 4] of variant;

    function startDoc: TOpenOffice_writer;
    function gotoEndOfSentence : TOpenOffice_writer;
    function gotoStartOfSentence : TOpenOffice_writer;
    function setValue(const aText: string): TOpenOffice_writer;
    function getValue: TOpenOffice_writer;
    function SelectAllText: TOpenOffice_writer;
    function CreateTable(ATable: TTableWriter): TOpenOffice_writer; overload;
    function CreateTable(ATable: TFdMemTable): TOpenOffice_writer; overload;
    function CreateTable(ATable: TClientDataset): TOpenOffice_writer; overload;

    property BoldActive : boolean read FBoldActive write FBoldActive;
    property Cursor : variant read FoCursor;
    property Value : string read FValueText;

    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
  published
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
  FoCursor: Variant;
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

  FoCursor := FobjDocument.getCurrentController.getViewCursor;

  oTextTable := FobjDocument.createInstance('com.sun.star.text.TextTable');
  oTextTable.initialize(numRows, numCols);
  linesCol := 0;
  FobjDocument.getText.insertTextContent(FoCursor, oTextTable, False);
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

function TOpenOffice_writer.getValue: TOpenOffice_writer;
begin
  if assigned(onBeforeGetValue) then
    onBeforeGetValue(self);

  try
    FoText :=  FobjTextCursor.Text;
    FoCursor := FoText.createTextCursor;
    FoCursor.gotoStart(False);
    FoCursor.gotoEnd(True);

    FValueText := FoCursor.getString;

    Result := Self;
  except
    on E: Exception do
    begin
      FValueText := '';
      raise Exception.Create('Erro ao obter texto do documento: ' + E.Message);
    end;
  end;

  if assigned(onAfterGetValue) then
    onAfterGetValue(self);
end;


function TOpenOffice_writer.gotoEndOfSentence: TOpenOffice_writer;
begin
  FobjTextCursor.jumpToEndOfPage;
  FoText := FobjTextCursor.Text;
  FoCursor := FoText.CreateTextCursor;
  FoCursor.gotoEnd(true);
  Result := self;
end;

function TOpenOffice_writer.gotoStartOfSentence: TOpenOffice_writer;
begin
  FoText := FobjTextCursor.Text;
  FoCursor := FoText.CreateTextCursor;
  FoCursor.gotoStart(False);

  Result := self;
end;

function TOpenOffice_writer.setValue(const aText: string): TOpenOffice_writer;
begin
  if Assigned(onBeforeSetValue) then
    onBeforeSetValue(self);

  FoText := FobjTextCursor.Text;
  FoText.InsertString(FobjTextCursor, '', true);

  FoCursor := FoText.CreateTextCursor;

  changeForDispatcher := true;

  FoText.InsertString(FobjTextCursor, aText, true);
  FoText.InsertControlCharacter(FoCursor, 10, true);

  FValueText  := aText;
  Result      := self;

  if assigned(onAfterSetValue) then
    onAfterSetValue(self);
end;

function TOpenOffice_writer.startDoc: TOpenOffice_writer;
begin
  if URlFile.Trim.IsEmpty then
    URlFile := FNewFile[integer(TpWriter)];

  LoadDocument(DocName); // cria a instancia do FobjDocument
  FobjTextCursor := FobjDocument.getCurrentController.getViewCursor;

  propsText[0] := FobjServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[1] := FobjServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[2] := FobjServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[3] := FobjServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[4] := FobjServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');

  FobjWriter := FobjDocument.getCurrentController.getFrame;
  FobjDispatcher := FobjServiceManager.createInstance
    ('com.sun.star.frame.DispatchHelper');

  Result := Self;
end;

function TOpenOffice_writer.SelectAllText: TOpenOffice_writer;
var
  loController, loViewCursor: Variant;
  largs: array[0..0] of Variant;
begin
  try
    loController := FobjDocument.getCurrentController;
    loViewCursor := loController.getViewCursor;

    FoText := FobjDocument.getText;
    FoCursor := FoText.createTextCursor;

    FoCursor.gotoStart(False);
    FoCursor.gotoEnd(True);

    loViewCursor.gotoRange(FoCursor.getStart, False);
    loViewCursor.gotoRange(FoCursor.getEnd, True);

    Result := Self;
  except
    on E: Exception do
      raise Exception.Create('Erro ao selecionar texto: ' + E.Message);
  end;
end;

procedure TOpenOffice_writer.SetDocName(const Value: string);
begin
  FDocName := Value;
end;

{ TTableWriter }


end.
