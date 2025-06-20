unit uOpenOffice.Writer.Table.Model;

interface

uses
  System.Generics.Collections, classes, SysUtils;

type
  TLine = TList<string>;

  TLineFieldValue_ = class
  private
    FLine: TLine;
  public
    property Line: TLine read FLine write FLine;
  end;

  TTableRowFields = TList<string>;
  TTableLinesValues = TList<TLineFieldValue_>;

  TTable = TDictionary<TTableRowFields, TTableLinesValues>;

  TTableWriter_ = class
  private
  class var
    FTable: TTable;
    FTableFields: TTableRowFields;
    FTableLines: TTableLinesValues;
    FLines: TLineFieldValue_;

    class procedure FreeTableCascade;
    class function AddNewLine(): TLine;
  public
    class property Table: TTable read FTable;
    class property TableFields: TTableRowFields read FTableFields;
    class property TableLines: TTableLinesValues read FTableLines;
    class property Lines: TLineFieldValue_ read FLines write FLines;

    class function AddNewLines(): TLineFieldValue_;

    class function New: TTableWriter_;
    class destructor Destroy;
  end;

implementation

class function TTableWriter_.AddNewLine: TLine;
begin
  Result := TLine.Create;
end;

class function TTableWriter_.AddNewLines: TLineFieldValue_;
begin
  Result := TLineFieldValue_.Create;
  Result.Line :=  AddNewLine;
end;

class destructor TTableWriter_.Destroy;
begin
  FreeTableCascade;
end;

class procedure TTableWriter_.FreeTableCascade;
begin
  if Assigned(FTableLines) then
  begin
    FTableLines.Clear;
    FreeAndNil(FTableLines);
  end;

  if Assigned(FTableFields) then
  begin
    FTableFields.Clear;
    FreeAndNil(FTableFields);
  end;

  if Assigned(FLines) then
  begin
    FLines.Line.Clear;

   {$IF CompilerVersion >= 29.0}
    if Assigned(Lines.Line) then
      FreeAndNil(Lines.Line);
    {$endif}

    FreeAndNil(FLines);
  end;

  if Assigned(FTable) then
  begin
    FTable.Clear;
    FreeAndNil(FTable);
  end;
end;

class function TTableWriter_.New: TTableWriter_;
begin
  FTable := TTable.Create;
  FTableFields := TTableRowFields.Create;
  FTableLines := TTableLinesValues.Create;
  FLines := TLineFieldValue_.Create;
  FLines.Line := TLine.Create;
  Result := TTableWriter_.Create;
  TableLines.Add(FLines);
end;

end.
