unit uInstallLibreOffice;

interface

uses  SysUtils, Variants, Classes, Vcl.OleCtrls, SHDocVw, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TInstallLibreOffice = class(TComponent)
  private
    const
      DonwloadClickJS = ' document.getElementsByClassName("dl_download_link")[0].click(); ';
      URL = 'https://www.libreoffice.org/download/download-libreoffice/';
    var
      URL_download :string;
      FWebBrowser : TWebBrowser;
    procedure download;
    procedure  WhenDocIsCompleted(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
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
  URL_download := 'https://www.libreoffice.org/donate/dl/win-x86_64/versao/pt-BR/LibreOffice_versao_Win_x64.msi';

  if not assigned(FWebBrowser) then
    FWebBrowser := TWebBrowser.Create(self);
end;

destructor TInstallLibreOffice.Destroy;
begin
  FreeAndNil(FWebBrowser);
  inherited;
end;

procedure TInstallLibreOffice.download;
begin
   FWebBrowser.Navigate(URL_download);
   FWebBrowser.OnDocumentComplete := nil;
end;

procedure TInstallLibreOffice.DownloadLibreOffice;
begin
  TWinControl(FWebBrowser).Name   := 'WebBrowser';
  FWebBrowser.Silent := true;  //don't show JS errors
  FWebBrowser.Visible:= false;  //visible...by default true
  FWebBrowser.HandleNeeded;

  FWebBrowser.Navigate(URL);
  FWebBrowser.RegisterAsBrowser:= True;
  FWebBrowser.OnDocumentComplete :=  WhenDocIsCompleted;
  FWebBrowser.Top    := 0;
  FWebBrowser.Left   := 0;
  FWebBrowser.Height := 0;
  FWebBrowser.Width  := 0;
end;

procedure TInstallLibreOffice.WhenDocIsCompleted(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
var
 Doc: IHTMLDocument2;
 body : string;
 versao : string;
 UrlDownload: string;
begin
  Doc := FWebBrowser.Document as IHTMLDocument2;
  body :=   doc.body.innerText;
  versao := Copy(body,pos('is available for the following operating systems/architectures',body)-6,6).Trim;
  UrlDownload := URL_download;
  UrlDownload := UrlDownload.Replace('versao',versao);
  URL_download :=  UrlDownload;
  download;
end;

end.
