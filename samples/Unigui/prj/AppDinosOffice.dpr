{$define UNIGUI_VCL} // Comment out this line to turn this project into an ISAPI module

{$ifndef UNIGUI_VCL}
library
{$else}
program
{$endif}
  AppDinosOffice;

uses
  uniGUIISAPI,
  Forms,
  Main in '..\Src\Main.pas' {MainForm: TUniForm},
  MainModule in '..\Src\MainModule.pas' {UniMainModule: TUniGUIMainModule},
  ServerModule in '..\Src\ServerModule.pas' {UniServerModule: TUniGUIServerModule};

{$R *.res}

{$ifndef UNIGUI_VCL}
exports
  GetExtensionVersion,
  HttpExtensionProc,
  TerminateExtension;
{$endif}

begin
{$ifdef UNIGUI_VCL}
  ReportMemoryLeaksOnShutdown := True;
  Application.Initialize;
  TUniServerModule.Create(Application);
  Application.Run;
{$endif}
end.
