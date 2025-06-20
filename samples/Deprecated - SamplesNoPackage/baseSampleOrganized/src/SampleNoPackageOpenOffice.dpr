program SampleNoPackageOpenOffice;

uses
  Vcl.Forms,
  MidasLib,
  Sample.View.Form.SampleNoPackage in '..\view\Sample.View.Form.SampleNoPackage.pas' {FormSampleNoPackage};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormSampleNoPackage, FormSampleNoPackage);
  Application.Run;
end.
