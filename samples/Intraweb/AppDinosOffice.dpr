program AppDinosOffice;

uses
  Forms,
  IWStart,
  UTF8ContentParser,
  ServerController in 'ServerController.pas',
  Unit1 in 'Unit1.pas',
  UserSessionUnit in 'UserSessionUnit.pas';

{$R *.res}

begin
  TIWStart.Execute(True);
end.
