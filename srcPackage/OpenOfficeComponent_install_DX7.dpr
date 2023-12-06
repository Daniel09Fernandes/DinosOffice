program OpenOfficeComponent_install_DX7;

uses
  Forms,
  uOpenOffice in '..\src_dx7\uOpenOffice.pas',
  uOpenOffice_calc in '..\src_dx7\uOpenOffice_calc.pas',
  uOpenOffice_writer in '..\src_dx7\uOpenOffice_writer.pas',
  uOpenOfficeCollors in '..\src_dx7\uOpenOfficeCollors.pas',
  uOpenOfficeEvents in '..\src_dx7\uOpenOfficeEvents.pas',
  uOpenOfficeHelper in '..\src_dx7\uOpenOfficeHelper.pas',
  uOpenOfficeSetPrinter in '..\src_dx7\uOpenOfficeSetPrinter.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Run;
end.
