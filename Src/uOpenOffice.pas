{MIT License

Copyright (c) 2022 Daniel Fernandes

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.}

unit uOpenOffice;

interface

uses ActiveX, System.Classes, Vcl.Dialogs, System.Variants, Windows,
  uOpenOfficeSetPrinter, uOpenOfficeEvents;

  Type
   TTypeOffice   = (TpCalc,TpWriter);
   TTypeLanguage = (TpCalcPTbr,TpCalcEn,TpWriterPTbr,TpWriterEn);

Type
  TOpenOffice = class(TComponent)
  private
    FURlFile: string;
    FSetPrinter         : TSetPrinter;
    FOnBeforePrint      : TBeforePrint;
    FOnBeforeCloseSheet : TBeforeCloseFile;
    FOnAfterCloseSheet  : TAfterCloseFile;

    FOnAfterGetValue   : TAfterGetValue;
    FOnBeforeGetValue  : TBeforeGetValue;
    FOnAfterSetValue   : TAfterSetValue;
    FOnBeforeSetValue  : TBeforeSetValue;

    procedure SetURlFile(const Value: string);
    procedure inicialization;
  protected
    { Protected declarations }
    objCoreReflection,
    objDesktop,
    objServiceManager,
    objDocument,
    objSCalc,
    objWriter,
    objDispatcher,
    objCell, Charts: OleVariant;
    NewFile    : array[0..1] of string;
    function convertFilePathToUrlFile(aFilePath: string): string;
    Property SetPrinter: TSetPrinter read FSetPrinter write FSetPrinter;
    procedure LoadDocument(FileName: string = '');
  published
    property URlFile: string read FURlFile write SetURlFile;
    property OnBeforePrint     : TBeforePrint      read FOnBeforePrint      write FOnBeforePrint;
    property OnBeforeCloseSheet: TBeforeCloseFile read FOnBeforeCloseSheet write FOnBeforeCloseSheet;
    property OnAfterCloseSheet : TAfterCloseFile  read FOnAfterCloseSheet  write FOnAfterCloseSheet;
    property OnBeforeGetValue  : TBeforeGetValue   read FOnBeforeGetValue   write FOnBeforeGetValue;
    property OnAfterGetValue   : TAfterGetValue    read FOnAfterGetValue    write FOnAfterGetValue;
    property OnBeforeSetValue  : TBeforeSetValue   read FOnBeforeSetValue   write FOnBeforeSetValue;
    property OnAfterSetValue   : TAfterSetValue    read FOnAfterSetValue    write FOnAfterSetValue;
  public
    procedure print;
    procedure CloseFile;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure saveFile(aFileName: String);
  end;

implementation

uses
  System.SysUtils, System.Win.ComObj;

{ TOpenOffice }

procedure TOpenOffice.CloseFile;
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

procedure TOpenOffice.SetURlFile(const Value: string);
begin
  FURlFile := Value;

  if FURlFile.Trim.IsEmpty or (FURlFile = NewFile[integer(TpCalc)] )
                           or (FURlFile = NewFile[integer(TpWriter)]) then
    exit;

  FURlFile := convertFilePathToUrlFile(FURlFile);
end;

constructor TOpenOffice.Create(AOwner: TComponent);
begin
  inherited;
  inicialization;
  FSetPrinter := TSetPrinter.Create(nil);
  NewFile[integer(TpCalc)]   := 'private:factory/scalc';
  NewFile[integer(TpWriter)] := 'private:factory/swriter';
end;

destructor TOpenOffice.Destroy;
begin
  inherited;
  FSetPrinter.Free;
  objCoreReflection := Unassigned;
  objDesktop := Unassigned;
  objServiceManager := Unassigned;
  objDocument := Unassigned;
  objSCalc := Unassigned;
  objCell := Unassigned;
end;

procedure TOpenOffice.inicialization;
begin
  try
    // Libre office
    objServiceManager := CreateOleObject('com.sun.star.ServiceManager');
    objCoreReflection := objServiceManager.createInstance('com.sun.star.reflection.CoreReflection');
    objDesktop        := objServiceManager.createInstance('com.sun.star.frame.Desktop');
  except
    messageDlg('Erro(pt-Br):  Instale o LibreOffice para usar o sistema' + #13 +
      #13 + 'Error(En)  :  install  the LibreOffice to use the system',
      TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], 0);
  end;
end;

procedure TOpenOffice.LoadDocument(FileName: string = '');
begin
  if FileName = '' then
      FileName := '_blank';

  objDocument := objDesktop.loadComponentFromURL(URlFile, FileName, 0, VarArrayOf([]));
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

procedure TOpenOffice.saveFile(aFileName: String);
begin
  aFileName := convertFilePathToUrlFile(aFileName);

  if aFileName.Trim.IsEmpty then
    aFileName := URlFile;

  objDocument.storeAsURL(aFileName, VarArrayOf([]));
end;

end.
