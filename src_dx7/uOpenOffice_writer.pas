unit uOpenOffice_writer;

interface

uses
  Classes, DB, ActiveX, uOpenOffice,
  dbWeb, ComObj, XMLDoc, XMLIntf, Dialogs, Variants, Windows,
  uOpenOfficeEvents, uOpenOfficeSetPrinter, uOpenOfficeCollors;

type

  // TTypeValue = (ftString, ftNumeric);
  TOpenOffice_writer = class(TOpenOffice)
  private
    objTextCursor, oText, oCursor: variant;
    FDocName,
    FValueText: string;
    FBoldActive,
    changeForDispatcher: boolean;
    procedure SetDocName(const Value: string);
  public
    Value: string;
    propsText: array [0 .. 4] of variant;
    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
    procedure startDoc;
    function gotoEndOfSentence : TOpenOffice_writer;
    function setValue(aText: string): TOpenOffice_writer;
    function getValue: string;
    property BoldActive : boolean read FBoldActive write FBoldActive;
    property Cursor : variant read oCursor;
    property ValueText : string read FValueText;
    function setUnderline(aUnderline: boolean): TOpenOffice_writer;
    function setBold(aBold: boolean): TOpenOffice_writer;
    function setFontHeight(aFontHeight: integer): TOpenOffice_writer;
    function setColorText(aColor: TOpenColor) : TOpenOffice_writer;
    function setFontName(aFont : string): TOpenOffice_writer;
  published
    property ServicesManager: OleVariant read objServiceManager;
    property DocName: string read FDocName write SetDocName;
  end;

  procedure Register;

implementation

{ TOpenOffice_writer }
uses
  SysUtils, Math;

procedure Register;
begin
  RegisterComponents('DinosOffice', [TOpenOffice_writer]);
end;

constructor TOpenOffice_writer.Create(AOwner: TComponent);
begin
  inherited;
  FBoldActive := false;
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

function TOpenOffice_writer.setValue(aText: string): TOpenOffice_writer;
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
  if Trim(URlFile) = '' then
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

function TOpenOffice_writer.setUnderline(aUnderline: boolean): TOpenOffice_writer;
begin
  if Trim(ValueText) <> '' then
  begin
    Cursor.CharUnderline := ifthen(aUnderline,1,0);
    setValue(ValueText);
  end;

  Result := self;
end;

function TOpenOffice_writer.setFontHeight(aFontHeight: integer) : TOpenOffice_writer;
begin
  propsText[1].Name := 'FontHeight.Height';
  propsText[1].Value := aFontHeight;
  objDispatcher.executeDispatch(objWriter, '.uno:FontHeight', '', 0, VarArrayOf(propsText));

  Result := self;
end;


function TOpenOffice_writer.setBold(aBold: boolean): TOpenOffice_writer;
var CtrlBold: boolean;
begin
  CtrlBold := false;
  propsText[0].Name := 'bold';
  propsText[0].Value := aBold;
  //Funcionando, porém rever
  if BoldActive and (not aBold) then
  begin
    BoldActive   := false;
    CtrlBold     := true;
    aBold        := true;
  end;

  if ( not BoldActive) and (aBold) then
  begin
    objDispatcher.executeDispatch(objWriter, '.uno:Bold', '', 0,  VarArrayOf(propsText));

    if not CtrlBold then
      BoldActive := true;
  end;

  Result := self;
end;

function TOpenOffice_writer.setFontName(aFont: string): TOpenOffice_writer;
begin
  if trim(ValueText) <> '' then
  begin
    Cursor.CharFontName := aFont;
    setValue(ValueText);
  end;
  Result := self;
end;

function TOpenOffice_writer.setColorText(aColor: TOpenColor): TOpenOffice_writer;
begin
  if  Trim(ValueText) <> '' then
  begin
     Cursor.SetPropertyValue('CharColor', aColor);
     setValue(ValueText);
  end;

  Result := self;
end;

end.
