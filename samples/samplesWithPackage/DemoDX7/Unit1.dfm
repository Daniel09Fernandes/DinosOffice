object Form1: TForm1
  Left = 236
  Top = 145
  Width = 404
  Height = 480
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 80
    Top = 272
    Width = 99
    Height = 13
    Caption = 'Caminho para salvar:'
  end
  object BitBtn1: TBitBtn
    Left = 72
    Top = 152
    Width = 97
    Height = 25
    Caption = 'Demo_Calc'
    TabOrder = 0
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 72
    Top = 184
    Width = 97
    Height = 25
    Caption = 'Fechar Planilha'
    TabOrder = 1
    OnClick = BitBtn2Click
  end
  object BitBtn3: TBitBtn
    Left = 216
    Top = 152
    Width = 105
    Height = 25
    Caption = 'Demo_Writer'
    TabOrder = 2
    OnClick = BitBtn3Click
  end
  object BitBtn4: TBitBtn
    Left = 216
    Top = 184
    Width = 105
    Height = 25
    Caption = 'Fechar Doc'
    TabOrder = 3
    OnClick = BitBtn4Click
  end
  object Button1: TButton
    Left = 72
    Top = 216
    Width = 97
    Height = 25
    Caption = 'Salvar'
    TabOrder = 4
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 216
    Top = 216
    Width = 105
    Height = 25
    Caption = 'Salvar'
    TabOrder = 5
    OnClick = Button2Click
  end
  object Edit1: TEdit
    Left = 80
    Top = 288
    Width = 241
    Height = 21
    TabOrder = 6
    Text = 'c:\'
  end
  object CheckBox1: TCheckBox
    Left = 80
    Top = 72
    Width = 241
    Height = 17
    Caption = 'Gerar documento em modo visibilidade?'
    TabOrder = 7
  end
  object OpenOffice_calc1: TOpenOffice_calc
    DocVisible = False
    Left = 112
    Top = 112
  end
  object OpenOffice_writer1: TOpenOffice_writer
    DocVisible = False
    Left = 256
    Top = 112
  end
end
