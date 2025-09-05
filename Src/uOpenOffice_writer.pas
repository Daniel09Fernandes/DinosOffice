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
    FobjTextCursor: Variant;
    FoText: Variant;
    FoCursor: Variant;
    FDocName: string;
    FValueText: string;
    FBoldActive: boolean;
    changeForDispatcher: boolean;
    Alphabet: array [0..25] of string;
    procedure SetDocName(const Value: string);
    function GetTextCursor: Variant;
    function EnsureTextCursor: Boolean;
    function EnsureDocument: Boolean;
  public
  var
    PropsText: array [0 .. 4] of variant;

    function StartDoc: TOpenOffice_writer;
    function GotoEndOfSentence : TOpenOffice_writer;
    function GotoStartOfSentence : TOpenOffice_writer;
    function SetValue(const aText: string): TOpenOffice_writer;
    function GetValue: TOpenOffice_writer;
    function SelectAllText: TOpenOffice_writer;
    function ClearSelection: TOpenOffice_writer;
    function CreateTable(ATable: TTableWriter): TOpenOffice_writer; overload;
    function CreateTable(ATable: TFdMemTable): TOpenOffice_writer; overload;
    function CreateTable(ATable: TClientDataset): TOpenOffice_writer; overload;
    function SelectBetweenText(AStartText, AEndText: string): TOpenOffice_writer;
    function SelectTextRange(AStartPos, AEndPos: Integer): TOpenOffice_writer;
    function ReplaceText(const ASearch, AReplace: string): TOpenOffice_writer;

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
  System.SysUtils, System.Math, StrUtils,
  uOpenOfficeHelper;

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

{code ReplaceText provided by @fabiokruger on Github: https://github.com/Daniel09Fernandes/DinosOffice/discussions/86#discussioncomment-13640228}
function TOpenOffice_writer.ReplaceText(const ASearch, AReplace: string): TOpenOffice_writer;
var
  SearchDescriptor, Found: OleVariant;
begin
  try
    try
      SearchDescriptor := FobjDocument.createSearchDescriptor;
      SearchDescriptor.setSearchString(ASearch);

      SearchDescriptor.SearchCaseSensitive := false;
      SearchDescriptor.SearchWords := false;
      SearchDescriptor.SearchRegularExpression := false;

      Found := FobjDocument.findFirst(SearchDescriptor);

      while not VarIsEmpty(Found) and not VarIsNull(Found) do
      begin
        try
          Found.setString(AReplace);
        except
          on E: Exception do
            OutputDebugString(PChar('Erro ao substituir texto: ' + E.Message));
        end;

        // Encontra a próxima ocorrência
        Found := FobjDocument.findNext(Found.End, SearchDescriptor);
      end;
    Except
      On E: Exception do
        OutputDebugString(PChar('Erro no ReplaceText: ' + E.Message));
    end;
  finally
    Result := Self;
  end;
end;

function TOpenOffice_writer.EnsureTextCursor: Boolean;
begin
  if VarIsEmpty(FobjTextCursor) or VarIsNull(FobjTextCursor) then
    FobjTextCursor := GetTextCursor;

  Result := (not (VarIsEmpty(FobjTextCursor) or VarIsNull(FobjTextCursor)));
end;

function TOpenOffice_writer.EnsureDocument: Boolean;
begin
  Result := (not (VarIsEmpty(FobjDocument) or VarIsNull(FobjDocument)));
  if not Result then
    raise Exception.Create('Documento não está aberto ou foi fechado');
end;

function TOpenOffice_writer.CreateTable(ATable: TTableWriter): TOpenOffice_writer;
var
  oTextTable: Variant;
  FoCursor: Variant;
  key, line, linesCol, numRows, numCols: Integer;
  lOutVal: string;
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

function TOpenOffice_writer.ClearSelection: TOpenOffice_writer;
begin
  if not EnsureDocument then Exit(Self);
  try
    FoCursor := FobjDocument.getCurrentController.getViewCursor;
    if not (VarIsEmpty(FoCursor) or VarIsNull(FoCursor)) then
      FoCursor.collapseToStart;

    SetBold(False);
    Result := Self;
  except
    on E: Exception do
      raise Exception.Create('Erro ao limpar seleção: ' + E.Message);
  end;
end;

function TOpenOffice_writer.CreateTable(ATable: TFdMemTable): TOpenOffice_writer;
var
  lFieldIdx,
  idxFields,
  lRecNo: Integer;
  lTable: TTableWriter;
  lLog: Boolean;
begin
  lLog := False;
  lRecNo := 0;
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
    Result := Self;
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
begin
  lRecNo := 0;
  lLog := False;
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
    Result := Self;
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
  FobjTextCursor := GetTextCursor;
  FoText := FobjTextCursor.Text;
  FoCursor := FoText.CreateTextCursor;
  FoCursor.gotoStart(False);

  Result := self;
end;

function TOpenOffice_writer.setValue(const aText: string): TOpenOffice_writer;
begin
  if not EnsureDocument or not EnsureTextCursor then Exit(Self);

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

function TOpenOffice_writer.SelectTextRange(AStartPos, AEndPos: Integer): TOpenOffice_writer;
var
  oText: Variant;
  oCursorStart, oCursorEnd: Variant;
begin
  if not EnsureDocument then
    Exit(Self);

  try
    oText := FobjDocument.getText;
    oCursorStart := oText.createTextCursorByRange(oText.getStart);
    oCursorEnd := oText.createTextCursorByRange(oText.getStart);
    oCursorStart.goRight(AStartPos, False);
    oCursorEnd.goRight(AEndPos, False);

    FobjDocument.getCurrentController.getViewCursor.gotoRange(oCursorStart.getStart, False);
    FobjDocument.getCurrentController.getViewCursor.gotoRange(oCursorEnd.getEnd, True);

    Result := Self;
  except
    on E: Exception do
      raise Exception.CreateFmt('Erro ao selecionar texto (posições %d a %d): %s',
        [AStartPos, AEndPos, E.Message]);
  end;
end;

function TOpenOffice_writer.SelectBetweenText(AStartText, AEndText: string): TOpenOffice_writer;
var
  lFullText: string;
  lStartPos, lEndPos: Integer;
begin
  if not EnsureDocument then
    Exit(Self);

  try
    SelectAllText;
    lFullText := GetValue.Value;

    lStartPos := Pos(AStartText, lFullText) - Length(AStartText);

    if lStartPos <= 0 then
      lStartPos := 0;

    lEndPos := Pos(AEndText, lFullText) + Length(AEndText)-1;
    if lEndPos = 0 then
      raise Exception.Create('Texto final não encontrado: ' + AEndText);

    SelectTextRange(lStartPos, lEndPos);
    Result := Self;
  except
    on E: Exception do
      raise Exception.Create('Erro ao selecionar entre textos: ' + E.Message);
  end;
end;

function TOpenOffice_writer.GetTextCursor: Variant;
begin
 if not EnsureDocument then
    Exit(Unassigned);

  try
    Result := FobjDocument.getCurrentController.getViewCursor;
    FobjTextCursor := Result;
  except
    on E: Exception do
    begin
      Result := Unassigned;
      raise Exception.Create('Erro ao obter cursor de texto: ' + E.Message);
    end;
  end;
end;

function TOpenOffice_writer.startDoc: TOpenOffice_writer;
begin
  if URlFile.Trim.IsEmpty then
    URlFile := FNewFile[integer(TpWriter)];

  LoadDocument(DocName); // cria a instancia do FobjDocument
  FobjTextCursor := GetTextCursor;

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
