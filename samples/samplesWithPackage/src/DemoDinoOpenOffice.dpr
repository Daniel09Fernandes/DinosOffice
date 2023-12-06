program DemoDinoOpenOffice;

uses
  Vcl.Forms,
  Unit1 in '..\view\Unit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
