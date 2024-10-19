unit uInstallLibreOffice;

interface

uses
  SysUtils, Variants, Classes, Vcl.OleCtrls, SHDocVw, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TInstallLibreOffice = class(TComponent)
  private
    const
      DownloadJS = 'document.getElementsByClassName("dl_download_link")[0].href;';
      URL = 'https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/';
    var
      URL_download: string;
      FWebBrowser: TWebBrowser;
    procedure Download;
    procedure WhenDocIsCompleted(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
  public
    procedure DownloadLibreOffice;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

implementation

uses
  Vcl.Controls, MSHTML, Vcl.Forms;

{ TInstallLibreOffice }

constructor TInstallLibreOffice.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  URL_download := '';
  if not Assigned(FWebBrowser) then
    FWebBrowser := TWebBrowser.Create(self);
end;

destructor TInstallLibreOffice.Destroy;
begin
  if assigned(FWebBrowser) then
    FreeAndNil(FWebBrowser);
  inherited;
end;

procedure TInstallLibreOffice.Download;
begin
  if URL_download <> '' then
  begin
    FWebBrowser.Navigate(URL_download);
    FWebBrowser.OnDocumentComplete := nil; // Reset the event after download
  end
  else
  begin
    writeLn('Erro: URL de download não encontrada!');
  end;
end;

procedure TInstallLibreOffice.DownloadLibreOffice;
begin
  TWinControl(FWebBrowser).Name := 'WebBrowser';
  FWebBrowser.Silent := True; // Don't show JS errors
  FWebBrowser.Visible := false; // Make it visible
  FWebBrowser.HandleNeeded;
  FWebBrowser.Navigate(URL);
  FWebBrowser.RegisterAsBrowser := True;
  FWebBrowser.OnDocumentComplete := WhenDocIsCompleted;
  FWebBrowser.Top := 0;
  FWebBrowser.Left := 0;
  FWebBrowser.Height := 0;
  FWebBrowser.Width := 0;
end;

procedure TInstallLibreOffice.WhenDocIsCompleted(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
var
  Doc: IHTMLDocument2;
  ElementCollection: IHTMLElementCollection;
  Element: IHTMLElement;
  I: Integer;
begin
  try
    Doc := FWebBrowser.Document as IHTMLDocument2;
    ElementCollection := Doc.all.tags('A') as IHTMLElementCollection;

    for I := 0 to ElementCollection.length - 1 do
    begin
      Element := ElementCollection.item(I, 0) as IHTMLElement;
      if (Element._className = 'dl_download_link') then
      begin
        URL_download := Element.getAttribute('href', 0);
        Break;
      end;
    end;

    if URL_download <> '' then
      Download
    else
      writeLn('Erro: URL de download não encontrada!');

  except
    on E: Exception do
      writeLn('Erro ao obter a URL de download: ' + E.Message);
  end;
end;

end.

