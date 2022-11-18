unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, uOpenOffice, uOpenOffice_calc, uOpenOfficeCollors,
  uOpenOffice_writer;

type
  TForm1 = class(TForm)
    OpenOffice_calc1: TOpenOffice_calc;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    OpenOffice_writer1: TOpenOffice_writer;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
 Openoffice_calc1.startSheet;

  Openoffice_calc1.SetValue(1,'A', 'STATUS')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opMagenta);

  Openoffice_calc1.SetValue(1,'B', 'VALOR').changeJustify(  fthRIGHT , ftvTOP)
     .SetBorder([bAll], opBrown)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opMagenta);

  Openoffice_calc1.SetValue(2,'B', 109, ftNumeric).SetBorder([bAll], opBrown);
  Openoffice_calc1.SetValue(2,'A', 'AGUA').SetBorder([bAll], opBrown);

  Openoffice_calc1.SetValue(3,'B', 105.55, ftNumeric).SetBorder([bAll], opBrown);
  Openoffice_calc1.SetValue(3,'A', 'LUZ').SetBorder([bAll], opBrown);

  Openoffice_calc1.SetValue(4,'B', 1005.22, ftNumeric);
  Openoffice_calc1.SetValue(4,'A', 'ALUGUEL');

   Openoffice_calc1.SetValue(6,'A', 'Total de linhas');
   Openoffice_calc1.SetValue(6,'B', Openoffice_calc1.CountRow, ftNumeric);

  Openoffice_calc1.SetValue(7,'A', 'Total de Colunas');
  Openoffice_calc1.SetValue(7,'B', Openoffice_calc1.CountCell, ftNumeric);


  Openoffice_calc1.addNewSheet('A Receber',1);

  Openoffice_calc1.SetValue(1,'A', 'VALOR')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true);

  Openoffice_calc1.SetValue(1,'B', 'DESC')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opCiano);

  Openoffice_calc1.SetValue(1,'C', 'SOMA')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opSoftRed);

  Openoffice_calc1.SetValue(1,'H', 'SOMA')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opSoftRed);

  Openoffice_calc1.SetValue(2,'A', 200, ftNumeric);
  Openoffice_calc1.SetValue(2,'B', 'Emprestimo');
  Openoffice_calc1.SetValue(2,'C', 0, ftNumeric);

  Openoffice_calc1.SetValue(3,'A', 369.55, ftNumeric);
  Openoffice_calc1.SetValue(3,'B', 'Dividendos');
  Openoffice_calc1.SetValue(3,'C', 0, ftNumeric);

  Openoffice_calc1.SetValue(4,'A', 1585.22, ftNumeric);
  Openoffice_calc1.SetValue(4,'B', 'ALUGUEL');
  Openoffice_calc1.SetValue(4,'C', 0, ftNumeric);

  Openoffice_calc1.SetValue(8,'A', 1585.22, ftNumeric);
  Openoffice_calc1.SetValue(8,'B', 'Renda extra');
  Openoffice_calc1.SetValue(8,'C', 0, ftNumeric);

  Openoffice_calc1.SetValue(15,'A', 1585.22, ftNumeric);
  Openoffice_calc1.SetValue(15,'B', 'ALUGUEL 2');
  Openoffice_calc1.SetValue(15,'C', 0, ftNumeric);

  Openoffice_calc1.SetValue(17,'A', 'Total de linhas');
  Openoffice_calc1.SetValue(17,'B', Openoffice_calc1.CountRow, ftNumeric);

  Openoffice_calc1.SetValue(19,'A', 'Total de Colunas');
  Openoffice_calc1.SetValue(19,'B', Openoffice_calc1.CountCell, ftNumeric);
  Openoffice_calc1.setFormula(20,'A',  '=A2+A3+A4+A15')
    .setBold(true);


  Openoffice_calc1.positionSheetByName('Planilha1');

  Openoffice_calc1.addChart(ctDefault, 0,3,'A','B','MeuGrafico',0);
  Openoffice_calc1.addChart(ctVertical, 0,3,'A','B','MeuGrafico',0);
  Openoffice_calc1.addChart(ctPie, 0,3,'A','B','MeuGrafico',0);
  Openoffice_calc1.addChart(ctLine, 0,3,'A','B','MeuGrafico',0);
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
 OpenOffice_calc1.CloseFile;
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
begin
 //Openoffice_writer1.URlFile := '';
 OpenOffice_writer1.startDoc;
 Openoffice_writer1.setBold(true)
     .setFontHeight(16)
     .setValue('Título: Apresentando o componente Libre Office writer via Delphi <3'+#13#13);
     Openoffice_writer1.gotoEndOfSentence;

  Openoffice_writer1.setBold(false)
     .setFontHeight(12)
     .setValue('Neste exemplo estou mostrando a criação de documentos via código, de um jeito simples, rápido e fácil.'+#13 );
     Openoffice_writer1.gotoEndOfSentence;

     Openoffice_writer1.setBold(false)
     .setFontHeight(12)
     .setValue('Espero que seja útil e que estejam gostando, lembrando o componente é open source e totalmente free!'+#13#13);
     Openoffice_writer1.gotoEndOfSentence;

     Openoffice_writer1.setBold(false)
     .setFontHeight(12)
     .SetBold(true)
     .setValue('Obrigado a todos pela presença!!!' +#13);
     Openoffice_writer1.gotoEndOfSentence;
end;

procedure TForm1.BitBtn4Click(Sender: TObject);
begin
 OpenOffice_writer1.CloseFile;
end;

end.
