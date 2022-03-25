unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, math, Vcl.Dialogs, Vcl.StdCtrls,  Vcl.ExtCtrls, Vcl.NumberBox, Vcl.Buttons,

  uOpenOffice,
  uOpenOfficeCollors,
  uOpenOfficeHelper,
  uOpenOfficeSetPrinter;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    edtSalvar: TLabeledEdit;
    Button4: TButton;
    edtArq: TLabeledEdit;
    Button5: TButton;
    edtAba: TLabeledEdit;
    edtLinha: TLabeledEdit;
    edtColuna: TLabeledEdit;
    edtValor: TLabeledEdit;
    Button6: TButton;
    Button7: TButton;
    chNumeric: TCheckBox;
    cbQuebraLinha: TCheckBox;
    Button8: TButton;
    Button9: TButton;
    edtPos: TLabeledEdit;
    GroupBox1: TGroupBox;
    RBhCenter: TRadioButton;
    RBhLeft: TRadioButton;
    RBhRight: TRadioButton;
    GroupBox2: TGroupBox;
    RBvTop: TRadioButton;
    RBvBottom: TRadioButton;
    RBvCenter: TRadioButton;
    edtTamanhoFonte: TNumberBox;
    Label1: TLabel;
    CBBold: TCheckBox;
    CBUnderline: TCheckBox;
    CBCorFont: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    CBCorFundo: TComboBox;
    Label4: TLabel;
    cbFontes: TComboBox;
    Graficos: TGroupBox;
    RBDefault: TRadioButton;
    RBVertical: TRadioButton;
    RBPie: TRadioButton;
    RBLine: TRadioButton;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    edtLde: TLabeledEdit;
    edtLAte: TLabeledEdit;
    edtCde: TLabeledEdit;
    edtCAte: TLabeledEdit;
    edtNomeGrafico: TLabeledEdit;
    Button10: TButton;
    Bevel4: TBevel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    OpenOffice1: TOpenOffice;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure OpenOffice1BeforeStartSheet(sender: TObject);
    procedure OpenOffice1BeforePrint(sender: TObject; var SetPrinter: TSetPrinter);
    procedure OpenOffice1BeforeGetValue(sender: TObject);
    procedure OpenOffice1BeforeSetValue(sender: TObject);
    procedure OpenOffice1BeforeCloseSheet(sender: TObject);
    procedure OpenOffice1AfterStartSheet(sender: TObject);
    procedure OpenOffice1AfterSetValue(sender: TObject);
    procedure OpenOffice1AfterGetValue(sender: TObject);
    procedure OpenOffice1AfterCloseSheet(sender: TObject);
  private
    procedure CapturarNomesDeFontes;
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
  ShowMessage(OpenOffice1.CountCell.ToString);
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
  ShowMessage(OpenOffice1.CountRow.ToString)
end;

procedure TForm1.Button10Click(Sender: TObject);
var tpGrafico : TTypeChart;
begin
   if RBDefault.Checked then
     tpGrafico := ctDefault
   else if RBLine.Checked then
     tpGrafico := ctLine
   else if RBPie.Checked then
     tpGrafico := ctPie
   else if RBVertical.Checked then
     tpGrafico := ctVertical;


   OpenOffice1.addChart(tpGrafico, StrToInt(edtLde.Text), StrToInt(edtLAte.Text)
        ,uppercase(edtCde.Text),UpperCase(edtCAte.Text),edtNomeGrafico.Text, strToInt(edtPos.Text));
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
   OpenOffice1.startSheet;

  OpenOffice1.SetValue(1,'A', 'STATUS')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opMagenta);

  OpenOffice1.SetValue(1,'B', 'VALOR').changeJustify(  fthRIGHT , ftvTOP)
     .SetBorder([bAll], opBrown)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opMagenta);

  OpenOffice1.SetValue(2,'B', 109, ftNumeric).SetBorder([bAll], opBrown);
  OpenOffice1.SetValue(2,'A', 'AGUA').SetBorder([bAll], opBrown);

  OpenOffice1.SetValue(3,'B', 105.55, ftNumeric).SetBorder([bAll], opBrown);
  OpenOffice1.SetValue(3,'A', 'LUZ').SetBorder([bAll], opBrown);

  OpenOffice1.SetValue(4,'B', 1005.22, ftNumeric);
  OpenOffice1.SetValue(4,'A', 'ALUGUEL');

   OpenOffice1.SetValue(6,'A', 'Total de linhas');
   OpenOffice1.SetValue(6,'B', OpenOffice1.CountRow, ftNumeric);

  OpenOffice1.SetValue(7,'A', 'Total de Colunas');
  OpenOffice1.SetValue(7,'B', OpenOffice1.CountCell, ftNumeric);


  OpenOffice1.addNewSheet('A Receber',1);

  OpenOffice1.SetValue(1,'A', 'VALOR')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true);

  OpenOffice1.SetValue(1,'B', 'DESC')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opCiano);

  OpenOffice1.SetValue(1,'C', 'SOMA')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opSoftRed);

  OpenOffice1.SetValue(1,'H', 'SOMA')
     .SetBorder([bAll], opBrown)
     .changeJustify(  fthRIGHT , ftvTOP)
     .setBold(true)
     .changeFont('Arial',12)
     .SetUnderline(true)
     .setColor( opWhite ,opSoftRed);

  OpenOffice1.SetValue(2,'A', 200, ftNumeric);
  OpenOffice1.SetValue(2,'B', 'Emprestimo');
  OpenOffice1.SetValue(2,'C', 0, ftNumeric);

  OpenOffice1.SetValue(3,'A', 369.55, ftNumeric);
  OpenOffice1.SetValue(3,'B', 'Dividendos');
  OpenOffice1.SetValue(3,'C', 0, ftNumeric);

  OpenOffice1.SetValue(4,'A', 1585.22, ftNumeric);
  OpenOffice1.SetValue(4,'B', 'ALUGUEL');
  OpenOffice1.SetValue(4,'C', 0, ftNumeric);

  OpenOffice1.SetValue(8,'A', 1585.22, ftNumeric);
  OpenOffice1.SetValue(8,'B', 'Renda extra');
  OpenOffice1.SetValue(8,'C', 0, ftNumeric);

  OpenOffice1.SetValue(15,'A', 1585.22, ftNumeric);
  OpenOffice1.SetValue(15,'B', 'ALUGUEL 2');
  OpenOffice1.SetValue(15,'C', 0, ftNumeric);

  OpenOffice1.SetValue(17,'A', 'Total de linhas');
  OpenOffice1.SetValue(17,'B', OpenOffice1.CountRow, ftNumeric);

  OpenOffice1.SetValue(19,'A', 'Total de Colunas');
  OpenOffice1.SetValue(19,'B', OpenOffice1.CountCell, ftNumeric);
  OpenOffice1.setFormula(20,'A',  '=A2+A3+A4+A15')
    .setBold(true);


  OpenOffice1.positionSheetByName('Planilha1');

  OpenOffice1.addChart(ctDefault, 0,3,'A','B','MeuGrafico',0);
  OpenOffice1.addChart(ctVertical, 0,3,'A','B','MeuGrafico',0);
  OpenOffice1.addChart(ctPie, 0,3,'A','B','MeuGrafico',0);
  OpenOffice1.addChart(ctLine, 0,3,'A','B','MeuGrafico',0);

end;

procedure TForm1.Button2Click(Sender: TObject);
begin
   try
      OpenOffice1.print;
      ShowMessage('Impressão Finalizada!');
   except
      ShowMessage('Erro ao imprimir!');
   end;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  if edtSalvar.Text <> '' then
    OpenOffice1.saveSheet(edtSalvar.Text);
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  OpenOffice1.closeSheet;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
   OpenOffice1.URlFile   := edtArq.Text;
   OpenOffice1.SheetName := edtAba.Text;
   OpenOffice1.startSheet;
end;

procedure TForm1.Button6Click(Sender: TObject);
var tp   : TTypeValue;
    jusH : THoriJustify;
    jusV : TVertJustify;
begin

  if chNumeric.Checked then
    tp := ftNumeric;



  if RBhCenter.Checked then
    jusH := fthCENTER
  else if RBhLeft.Checked then
    jusH := fthLEFT
  else if RBhRight.Checked then
    jusH := fthRIGHT;


  if RBvTop.Checked then
    jusV := ftvTOP
  else if RBvBottom.Checked then
    jusV := ftvBOTTOM
  else if RBvCenter.Checked then
    jusV := ftvCENTER;

  OpenOffice1.SetValue(StrToInt(edtLinha.Text), edtColuna.Text, edtValor.Text, tp,cbQuebraLinha.Checked)
    .setBold(CBBold.Checked)
    .SetUnderline(CBUnderline.Checked)
    .setColor(TOpenColor( CBCorFont.Items.Objects[ CBCorFont.ItemIndex]),TOpenColor( CBCorFundo.Items.Objects[ CBCorFundo.ItemIndex]))
    .changeFont(cbFontes.Text,edtTamanhoFonte.ValueInt)
    .changeJustify(jusH,jusV)
    .SetBorder([bAll], opBlack);
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
  edtValor.Text := OpenOffice1.GetValue( StrToInt(edtLinha.Text), edtColuna.Text).Value;
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
  OpenOffice1.addNewSheet(edtAba.Text, StrToInt(edtPos.Text));
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
  OpenOffice1.positionSheetByName(edtAba.Text);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    CBCorFont.Items.AddObject('Preto', TObject(opBlack));
  CBCorFont.Items.AddObject('Azul'    , TObject(opBlue));
  CBCorFont.Items.AddObject('Verde'   , TObject(opGreen));
  CBCorFont.Items.AddObject('Turquesa', TObject(optTurquesa));
  CBCorFont.Items.AddObject('Vermelho', TObject(opRed));
  CBCorFont.Items.AddObject('Magenta' , TObject(opMagenta));
  CBCorFont.Items.AddObject('Marrom'  , TObject(opBrown));
  CBCorFont.Items.AddObject('Cinza0', TObject(opGray));
  CBCorFont.Items.AddObject('Cinza claro', TObject(opSoftGray));
  CBCorFont.Items.AddObject('Azul claro', TObject(opSoftBlue));
  CBCorFont.Items.AddObject('Cinza 6',TObject(opGreen6));
  CBCorFont.Items.AddObject('Ciano', TObject(opCiano));
  CBCorFont.Items.AddObject('Vermelho claro', TObject(opSoftRed));
  CBCorFont.Items.AddObject('Magenta claro', TObject(opSoftMagenta));
  CBCorFont.Items.AddObject('Amarelo', TObject(opYellow));
  CBCorFont.Items.AddObject('Branco' , TObject(opWhite));
  CBCorFont.Items.AddObject('Cinza 30', TObject(opGray30));
  CBCorFont.Items.AddObject('Salmão', TObject(opSalmon));
  CBCorFont.Items.AddObject('Laranja', TObject(opOrange));
  CBCorFont.Items.AddObject('Laranja 80', TObject(opOrange80));
  CBCorFont.Items.AddObject('Bordo',TObject(opBordo));

  CBCorFundo.Items :=  CBCorFont.Items;

  CBCorFundo.ItemIndex := 15;
  CBCorFont.ItemIndex  := 0;

  edtTamanhoFonte.Value := 10;

  CapturarNomesDeFontes;

  edtSalvar.Text := ExtractFileDir(GetCurrentDir)+'\';
  edtArq.Text    :=   edtSalvar.Text;

end;


procedure TForm1.OpenOffice1AfterCloseSheet(sender: TObject);
begin
 // ShowMessage('Depois de Fechar');
end;

procedure TForm1.OpenOffice1AfterGetValue(sender: TObject);
begin
 //ShowMessage('Depois de pegar o valor');
end;

procedure TForm1.OpenOffice1AfterSetValue(sender: TObject);
begin
 // ShowMessage('Depois de Setar o valor');
end;

procedure TForm1.OpenOffice1AfterStartSheet(sender: TObject);
begin
   // ShowMessage('iniciado');
end;

procedure TForm1.OpenOffice1BeforeCloseSheet(sender: TObject);
begin
//    ShowMessage('Antes de fechar');
end;

procedure TForm1.OpenOffice1BeforeGetValue(sender: TObject);
begin
  //ShowMessage('Antes de pegar o valor');
end;

procedure TForm1.OpenOffice1BeforePrint(sender: TObject; var SetPrinter: TSetPrinter);
begin
  ShowMessage('Iniciando impressão...');
  SetPrinter.PaperSize_Width  := 50000;
  SetPrinter.PaperSize_Height := 50000;
  SetPrinter.PrinterName      := 'Microsoft Print to PDF';
  SetPrinter.Pages            := '1; 3;';
end;

procedure TForm1.OpenOffice1BeforeSetValue(sender: TObject);
begin
 //   ShowMessage('Antes de Setar o valor');
end;

procedure TForm1.OpenOffice1BeforeStartSheet(sender: TObject);
begin
//  ShowMessage(  'Iniciando...'  );
end;

function EnumFontsProc(var LogFont: TLogFont; var TextMetric: TTextMetric;
  FontType: Integer; Data: Pointer): Integer; stdcall;

begin
  TStrings(Data).Add(LogFont.lfFaceName);
  Result := 1;
end;

procedure TForm1.CapturarNomesDeFontes;
var
  DC: HDC;
begin

  DC := GetDC(0);

  EnumFonts(DC, nil, @EnumFontsProc, Pointer(cbFontes.Items));

  ReleaseDC(0, DC);

  cbFontes.Sorted := true;

  cbFontes.ItemIndex := 0;

end;

end.
