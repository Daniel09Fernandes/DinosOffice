object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Demo'
  ClientHeight = 390
  ClientWidth = 558
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 12
    Width = 71
    Height = 13
    Caption = 'tamanho fonte'
  end
  object Label2: TLabel
    Left = 9
    Top = 58
    Width = 61
    Height = 13
    Caption = 'Cor da fonte'
  end
  object Label3: TLabel
    Left = 166
    Top = 58
    Width = 63
    Height = 13
    Caption = 'Cor do fundo'
  end
  object Label4: TLabel
    Left = 90
    Top = 10
    Width = 33
    Height = 13
    Caption = 'Fontes'
  end
  object Bevel1: TBevel
    Left = 337
    Top = 125
    Width = 215
    Height = 80
  end
  object Bevel2: TBevel
    Left = 0
    Top = 8
    Width = 328
    Height = 197
  end
  object Bevel3: TBevel
    Left = 3
    Top = 267
    Width = 552
    Height = 131
  end
  object Bevel4: TBevel
    Left = 5
    Top = 209
    Width = 546
    Height = 50
  end
  object Button1: TButton
    Left = 8
    Top = 355
    Width = 121
    Height = 30
    Caption = 'Criar nova planilha'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 480
    Top = 355
    Width = 65
    Height = 30
    Caption = 'Imprimir '
    TabOrder = 1
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 270
    Top = 355
    Width = 121
    Height = 30
    Caption = 'Salvar'
    TabOrder = 2
    OnClick = Button3Click
  end
  object edtSalvar: TLabeledEdit
    Left = 8
    Top = 328
    Width = 537
    Height = 21
    EditLabel.Width = 47
    EditLabel.Height = 13
    EditLabel.Caption = 'Salvar em'
    TabOrder = 3
  end
  object Button4: TButton
    Left = 402
    Top = 355
    Width = 67
    Height = 30
    Caption = 'Fechar'
    TabOrder = 4
    OnClick = Button4Click
  end
  object edtArq: TLabeledEdit
    Left = 8
    Top = 290
    Width = 537
    Height = 21
    EditLabel.Width = 99
    EditLabel.Height = 13
    EditLabel.Caption = 'Carregar documento'
    TabOrder = 5
  end
  object Button5: TButton
    Left = 139
    Top = 355
    Width = 121
    Height = 30
    Caption = 'Carregar documeto'
    TabOrder = 6
    OnClick = Button5Click
  end
  object edtAba: TLabeledEdit
    Left = 345
    Top = 149
    Width = 121
    Height = 21
    EditLabel.Width = 73
    EditLabel.Height = 13
    EditLabel.Caption = 'Aba da planilha'
    TabOrder = 7
  end
  object edtLinha: TLabeledEdit
    Left = 7
    Top = 146
    Width = 35
    Height = 21
    EditLabel.Width = 25
    EditLabel.Height = 13
    EditLabel.Caption = 'Linha'
    TabOrder = 8
  end
  object edtColuna: TLabeledEdit
    Left = 58
    Top = 146
    Width = 34
    Height = 21
    EditLabel.Width = 33
    EditLabel.Height = 13
    EditLabel.Caption = 'Coluna'
    TabOrder = 9
  end
  object edtValor: TLabeledEdit
    Left = 110
    Top = 146
    Width = 200
    Height = 21
    EditLabel.Width = 24
    EditLabel.Height = 13
    EditLabel.Caption = 'Valor'
    TabOrder = 10
  end
  object Button6: TButton
    Left = 7
    Top = 173
    Width = 89
    Height = 25
    Caption = 'Adicionar Valor'
    TabOrder = 11
    OnClick = Button6Click
  end
  object Button7: TButton
    Left = 110
    Top = 173
    Width = 67
    Height = 25
    Caption = 'Pegar valor'
    TabOrder = 12
    OnClick = Button7Click
  end
  object chNumeric: TCheckBox
    Left = 111
    Top = 106
    Width = 66
    Height = 17
    Caption = 'is Numeric?'
    TabOrder = 13
  end
  object cbQuebraLinha: TCheckBox
    Left = 8
    Top = 106
    Width = 97
    Height = 17
    Caption = 'Quebra de linha'
    TabOrder = 14
  end
  object Button8: TButton
    Left = 345
    Top = 176
    Width = 51
    Height = 25
    Caption = 'add'
    TabOrder = 15
    OnClick = Button8Click
  end
  object Button9: TButton
    Left = 402
    Top = 176
    Width = 51
    Height = 25
    Caption = 'Trocar'
    TabOrder = 16
    OnClick = Button9Click
  end
  object edtPos: TLabeledEdit
    Left = 472
    Top = 149
    Width = 34
    Height = 21
    EditLabel.Width = 36
    EditLabel.Height = 13
    EditLabel.Caption = 'Posi'#231#227'o'
    TabOrder = 17
  end
  object GroupBox1: TGroupBox
    Left = 334
    Top = 8
    Width = 218
    Height = 37
    Caption = 'Alinhamento horizontal'
    TabOrder = 18
    object RBhCenter: TRadioButton
      Left = 79
      Top = 18
      Width = 113
      Height = 17
      Caption = 'Centro'
      TabOrder = 0
    end
    object RBhLeft: TRadioButton
      Left = 3
      Top = 18
      Width = 70
      Height = 17
      Caption = 'Esquerda'
      TabOrder = 1
    end
    object RBhRight: TRadioButton
      Left = 146
      Top = 18
      Width = 113
      Height = 17
      Caption = 'Direita'
      TabOrder = 2
    end
  end
  object GroupBox2: TGroupBox
    Left = 337
    Top = 49
    Width = 88
    Height = 70
    Caption = 'Alinha. Vertical'
    TabOrder = 19
    object RBvTop: TRadioButton
      Left = 16
      Top = 15
      Width = 113
      Height = 17
      Caption = 'Cima'
      TabOrder = 0
    end
    object RBvBottom: TRadioButton
      Left = 16
      Top = 51
      Width = 113
      Height = 17
      Caption = 'Baixo'
      TabOrder = 1
    end
    object RBvCenter: TRadioButton
      Left = 16
      Top = 33
      Width = 113
      Height = 17
      Caption = 'Centro'
      TabOrder = 2
    end
  end
  object edtTamanhoFonte: TNumberBox
    Left = 8
    Top = 29
    Width = 65
    Height = 21
    TabOrder = 20
  end
  object CBBold: TCheckBox
    Left = 183
    Top = 106
    Width = 60
    Height = 17
    Caption = 'Negrito'
    TabOrder = 21
  end
  object CBUnderline: TCheckBox
    Left = 243
    Top = 106
    Width = 82
    Height = 17
    Caption = 'Sublinhado'
    TabOrder = 22
  end
  object CBCorFont: TComboBox
    Left = 8
    Top = 77
    Width = 145
    Height = 21
    TabOrder = 23
    Text = 'opBlack = 0,'
  end
  object CBCorFundo: TComboBox
    Left = 165
    Top = 77
    Width = 145
    Height = 21
    ItemIndex = 0
    TabOrder = 24
    Text = 'opBlack = 0,'
    Items.Strings = (
      'opBlack = 0,'
      'opBlue = 128,'
      'opGreen = 32768,'
      'optTurquesa = 32896,'
      'opRed = 8388608,'
      'opMagenta = 8388736,'
      'opBrown = 8421376,'
      'opGray = 8421504,'
      'opSoftGray = 12632256,'
      'opSoftBlue = 255,'
      'opGreen6 = 4057917,'
      'opCiano = 65535,'
      'opSoftRed = 16711680,'
      'opSoftMagenta = 16711935,'
      'opYellow = 16776960,'
      'opWhite = 16777215,'
      'opGray30 = 11776947,'
      'opSalmon = 26316,'
      'opOrange = 16750950,'
      'opOrange80 = 10066431,'
      'opBordo = 16777164')
  end
  object cbFontes: TComboBox
    Left = 89
    Top = 29
    Width = 221
    Height = 21
    TabOrder = 25
  end
  object Graficos: TGroupBox
    Left = 429
    Top = 49
    Width = 123
    Height = 70
    Caption = 'Tipos de Graficos'
    TabOrder = 26
    object RBDefault: TRadioButton
      Left = 9
      Top = 20
      Width = 56
      Height = 17
      Caption = 'Default'
      Checked = True
      TabOrder = 0
      TabStop = True
    end
    object RBVertical: TRadioButton
      Left = 9
      Top = 41
      Width = 48
      Height = 17
      Caption = 'Vertical'
      TabOrder = 1
    end
    object RBPie: TRadioButton
      Left = 74
      Top = 20
      Width = 116
      Height = 17
      Caption = 'Pie'
      TabOrder = 2
    end
    object RBLine: TRadioButton
      Left = 74
      Top = 41
      Width = 116
      Height = 17
      Caption = 'Line'
      TabOrder = 3
    end
  end
  object edtLde: TLabeledEdit
    Left = 24
    Top = 228
    Width = 34
    Height = 21
    EditLabel.Width = 40
    EditLabel.Height = 13
    EditLabel.Caption = 'Linha de'
    TabOrder = 27
  end
  object edtLAte: TLabeledEdit
    Left = 75
    Top = 228
    Width = 34
    Height = 21
    EditLabel.Width = 41
    EditLabel.Height = 13
    EditLabel.Caption = 'linha ate'
    TabOrder = 28
  end
  object edtCde: TLabeledEdit
    Left = 139
    Top = 228
    Width = 34
    Height = 21
    EditLabel.Width = 48
    EditLabel.Height = 13
    EditLabel.Caption = 'Coluna de'
    TabOrder = 29
  end
  object edtCAte: TLabeledEdit
    Left = 196
    Top = 228
    Width = 34
    Height = 21
    EditLabel.Width = 53
    EditLabel.Height = 13
    EditLabel.Caption = 'Coluna Ate'
    TabOrder = 30
  end
  object edtNomeGrafico: TLabeledEdit
    Left = 267
    Top = 228
    Width = 127
    Height = 21
    EditLabel.Width = 64
    EditLabel.Height = 13
    EditLabel.Caption = 'Nome Grafico'
    TabOrder = 31
  end
  object Button10: TButton
    Left = 410
    Top = 228
    Width = 75
    Height = 21
    Caption = 'Add grafico'
    TabOrder = 32
    OnClick = Button10Click
  end
  object BitBtn1: TBitBtn
    Left = 187
    Top = 173
    Width = 62
    Height = 25
    Caption = 'Qde Coluna'
    TabOrder = 33
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 254
    Top = 173
    Width = 56
    Height = 25
    Caption = 'Qde Linha'
    TabOrder = 34
    OnClick = BitBtn2Click
  end
  object OpenOffice1: TOpenOffice
    Left = 496
    Top = 192
  end
end
