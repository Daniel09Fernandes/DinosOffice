unit UConvertPDFToSheet;

interface

uses
  SHDocVw, Vcl.Controls;
  Type
   TConvertPDFToSheet = class
     private
       Const
         FUrl = 'https://www.ilovepdf.com/pt/pdf_para_excel';
     public
        procedure callConversor;
        constructor Create;
        destructor Destroy;
   end;

implementation

uses shellApi;
{ TConvertPDFToSheet }

procedure TConvertPDFToSheet.callConversor;
begin
//Provisorio
  ShellExecute(0,0,FUrl,'','',0);
end;

constructor TConvertPDFToSheet.Create;
begin
  
end;

destructor TConvertPDFToSheet.Destroy;
begin

end;

end.
