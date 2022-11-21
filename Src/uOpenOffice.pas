unit uOpenOffice;

interface

uses ActiveX, System.Classes, Vcl.Dialogs, System.Variants, Windows,
  uOpenOfficeSetPrinter, uOpenOfficeEvents, uInstallLibreOffice,
  System.UITypes;

Type
  TTypeOffice = (TpCalc, TpWriter);
  TTypeLanguage = (TpCalcPTbr, TpCalcEn, TpWriterPTbr, TpWriterEn);

Type
  TOpenOffice = class(TComponent)
  private
    FURlFile: string;
    FSetPrinter: TSetPrinter;
    FOnBeforePrint: TBeforePrint;
    FOnBeforeCloseFile: TBeforeCloseFile;
    FOnAfterCloseFile: TAfterCloseFile;

    FOnAfterGetValue: TAfterGetValue;
    FOnBeforeGetValue: TBeforeGetValue;
    FOnAfterSetValue: TAfterSetValue;
    FOnBeforeSetValue: TBeforeSetValue;

    InstallLibreOffice: TInstallLibreOffice;
    procedure SetURlFile(const Value: string);
    procedure inicialization;
  protected
    { Protected declarations }
    objCoreReflection, objDesktop, objServiceManager, objDocument, objSCalc,
      objWriter, objDispatcher, objCell, Charts: OleVariant;
    NewFile: array [0 .. 1] of string;
    function convertFilePathToUrlFile(aFilePath: string): string;
    Property SetPrinter: TSetPrinter read FSetPrinter write FSetPrinter;
    procedure LoadDocument(FileName: string = '');
  published
    property URlFile: string read FURlFile write SetURlFile;
    property OnBeforePrint: TBeforePrint read FOnBeforePrint
      write FOnBeforePrint;
    property OnBeforeCloseFile: TBeforeCloseFile read FOnBeforeCloseFile
      write FOnBeforeCloseFile;
    property OnAfterCloseFile: TAfterCloseFile read FOnAfterCloseFile
      write FOnAfterCloseFile;
    property OnBeforeGetValue: TBeforeGetValue read FOnBeforeGetValue
      write FOnBeforeGetValue;
    property OnAfterGetValue: TAfterGetValue read FOnAfterGetValue
      write FOnAfterGetValue;
    property OnBeforeSetValue: TBeforeSetValue read FOnBeforeSetValue
      write FOnBeforeSetValue;
    property OnAfterSetValue: TAfterSetValue read FOnAfterSetValue
      write FOnAfterSetValue;
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
  if Assigned(FOnBeforeCloseFile) then
    FOnBeforeCloseFile(self);

  objDocument.close(True);

  if Assigned(FOnAfterCloseFile) then
    FOnAfterCloseFile(self);
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

  if FURlFile.Trim.IsEmpty or (FURlFile = NewFile[integer(TpCalc)]) or
    (FURlFile = NewFile[integer(TpWriter)]) then
    exit;

  FURlFile := convertFilePathToUrlFile(FURlFile);
end;

constructor TOpenOffice.Create(AOwner: TComponent);
begin
  inherited;
  inicialization;
  FSetPrinter := TSetPrinter.Create(nil);
  NewFile[integer(TpCalc)] := 'private:factory/scalc';
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

  if assigned(InstallLibreOffice) then
    freeAndNil(InstallLibreOffice);
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
    if messageDlg('Erro(pt-Br):  Instale o LibreOffice para usar o sistema' +
      #13 + #13 + 'Error(En)  :  install  the LibreOffice to use the system' +
      #13#13 + 'Dowload in: https://www.libreoffice.org/download/download-libreoffice/'
      + #13#13 + '(pt-Br) Deseja instalar a versão mais recente do LibreOffice?'
      + #13 + '(En) Do you want to install the latest version of LibreOffice?',

      TMsgDlgType.mtWarning, [TMsgDlgBtn.mbYes, TMsgDlgBtn.mbNo], 0) = mrYes
    then
    begin
      InstallLibreOffice := TInstallLibreOffice.Create;
      InstallLibreOffice.DownloadLibreOffice;
    end;
  end;
end;

procedure TOpenOffice.LoadDocument(FileName: string = '');
begin
  if FileName = '' then
    FileName := '_blank';

  objDocument := objDesktop.loadComponentFromURL(URlFile, FileName, 0,
    VarArrayOf([]));
end;

procedure TOpenOffice.print;
var
  PaperSize: variant;
  printerProperties: array [0 .. 3] of variant;
begin

  if Assigned(OnBeforePrint) then
  begin
    OnBeforePrint(self, FSetPrinter);

    PaperSize := objServiceManager.Bridge_GetStruct('com.sun.star.awt.Size');

    PaperSize.Width := FSetPrinter.PaperSize_Width;
    PaperSize.Height := FSetPrinter.PaperSize_Height;

    printerProperties[0] := objServiceManager.Bridge_GetStruct
      ('com.sun.star.beans.PropertyValue');
    printerProperties[0].Name := 'Name';
    printerProperties[0].Value := FSetPrinter.PrinterName;

    printerProperties[1] := objServiceManager.Bridge_GetStruct
      ('com.sun.star.beans.PropertyValue');
    printerProperties[1].Name := 'PaperSize';
    printerProperties[1].Value := PaperSize;

    printerProperties[2] := objServiceManager.Bridge_GetStruct
      ('com.sun.star.beans.PropertyValue');
    printerProperties[2].Name := 'Pages';
    printerProperties[2].Value := FSetPrinter.Pages;

    objDocument.Printer := VarArrayOf(printerProperties);

    objDocument.print(VarArrayOf(printerProperties));
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
