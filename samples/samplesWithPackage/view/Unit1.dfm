object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Demo'
  ClientHeight = 712
  ClientWidth = 779
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Position = poDesktopCenter
  OnCreate = FormCreate
  TextHeight = 13
  object PageControl2: TPageControl
    Left = 0
    Top = 0
    Width = 779
    Height = 712
    ActivePage = TabSheet2
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 723
    object TabSheet1: TTabSheet
      Caption = 'Calc'
      OnShow = TabSheet1Show
      DesignSize = (
        771
        684)
      object Bevel1: TBevel
        Left = 0
        Top = 51
        Width = 399
        Height = 142
        Style = bsRaised
      end
      object Bevel2: TBevel
        Left = 1
        Top = 279
        Width = 767
        Height = 118
        Anchors = [akLeft, akTop, akRight]
        Style = bsRaised
        ExplicitWidth = 711
      end
      object Bevel3: TBevel
        Left = 405
        Top = 120
        Width = 362
        Height = 72
        Anchors = [akLeft, akTop, akRight]
        Style = bsRaised
        ExplicitWidth = 231
      end
      object Bevel4: TBevel
        Left = -270
        Top = -5
        Width = 669
        Height = 50
        Style = bsRaised
      end
      object Label1: TLabel
        Left = 6
        Top = 5
        Width = 71
        Height = 13
        Caption = 'tamanho fonte'
      end
      object Label2: TLabel
        Left = 6
        Top = 50
        Width = 61
        Height = 13
        Caption = 'Cor da fonte'
      end
      object Label3: TLabel
        Left = 163
        Top = 50
        Width = 63
        Height = 13
        Caption = 'Cor do fundo'
      end
      object Label4: TLabel
        Left = 93
        Top = 5
        Width = 33
        Height = 13
        Caption = 'Fontes'
      end
      object lbl1: TLabel
        Left = 3
        Top = 529
        Width = 3
        Height = 13
      end
      object lbl2: TLabel
        Left = 3
        Top = 548
        Width = 3
        Height = 13
      end
      object BitBtn1: TBitBtn
        Left = 180
        Top = 160
        Width = 62
        Height = 25
        Caption = 'Qde Coluna'
        TabOrder = 0
        OnClick = BitBtn1Click
      end
      object BitBtn2: TBitBtn
        Left = 252
        Top = 160
        Width = 56
        Height = 25
        Caption = 'Qde Linha'
        TabOrder = 1
        OnClick = BitBtn2Click
      end
      object Button1: TButton
        Left = 3
        Top = 363
        Width = 121
        Height = 30
        Caption = 'Criar nova planilha'
        TabOrder = 2
        OnClick = Button1Click
      end
      object Button10: TButton
        Left = 6
        Top = 237
        Width = 75
        Height = 21
        Caption = 'Add grafico'
        TabOrder = 3
        OnClick = Button10Click
      end
      object Button2: TButton
        Left = 263
        Top = 363
        Width = 81
        Height = 30
        Caption = 'Imprimir '
        TabOrder = 4
        OnClick = Button2Click
      end
      object Button3: TButton
        Left = 352
        Top = 363
        Width = 121
        Height = 30
        Caption = 'Salvar'
        TabOrder = 5
        OnClick = Button3Click
      end
      object Button4: TButton
        Left = 480
        Top = 363
        Width = 67
        Height = 30
        Caption = 'Fechar'
        TabOrder = 6
        OnClick = Button4Click
      end
      object Button5: TButton
        Left = 134
        Top = 363
        Width = 121
        Height = 30
        Caption = 'Carregar documeto'
        TabOrder = 7
        OnClick = Button5Click
      end
      object Button6: TButton
        Left = 84
        Top = 160
        Width = 89
        Height = 25
        Caption = 'Adicionar Valor'
        TabOrder = 8
        OnClick = Button6Click
      end
      object Button7: TButton
        Left = 6
        Top = 160
        Width = 75
        Height = 25
        Caption = 'Pegar valor'
        TabOrder = 9
        OnClick = Button7Click
      end
      object Button8: TButton
        Left = 413
        Top = 164
        Width = 57
        Height = 25
        Caption = 'Adicionar'
        TabOrder = 10
        OnClick = Button8Click
      end
      object Button9: TButton
        Left = 472
        Top = 164
        Width = 76
        Height = 25
        Caption = 'Trocar - Nome'
        TabOrder = 11
        OnClick = Button9Click
      end
      object CBBold: TCheckBox
        Left = 180
        Top = 97
        Width = 60
        Height = 17
        Caption = 'Negrito'
        TabOrder = 12
      end
      object CBCorFont: TComboBox
        Left = 6
        Top = 69
        Width = 145
        Height = 21
        TabOrder = 13
        Text = 'opBlack = 0,'
      end
      object CBCorFundo: TComboBox
        Left = 163
        Top = 69
        Width = 145
        Height = 21
        ItemIndex = 0
        TabOrder = 14
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
        Left = 93
        Top = 21
        Width = 294
        Height = 21
        TabOrder = 15
      end
      object cbQuebraLinha: TCheckBox
        Left = 6
        Top = 97
        Width = 97
        Height = 17
        Caption = 'Quebra de linha'
        TabOrder = 16
      end
      object CBUnderline: TCheckBox
        Left = 236
        Top = 97
        Width = 82
        Height = 17
        Caption = 'Sublinhado'
        TabOrder = 17
      end
      object chNumeric: TCheckBox
        Left = 103
        Top = 97
        Width = 76
        Height = 17
        Caption = 'is Numeric?'
        TabOrder = 18
      end
      object edtAba: TLabeledEdit
        Left = 453
        Top = 139
        Width = 170
        Height = 21
        EditLabel.Width = 73
        EditLabel.Height = 13
        EditLabel.Caption = 'Aba da planilha'
        TabOrder = 19
        Text = ''
      end
      object edtArq: TLabeledEdit
        Left = 3
        Top = 337
        Width = 668
        Height = 21
        Anchors = [akLeft, akTop, akRight]
        EditLabel.Width = 99
        EditLabel.Height = 13
        EditLabel.Caption = 'Carregar documento'
        TabOrder = 20
        Text = ''
        ExplicitWidth = 612
      end
      object edtCAte: TLabeledEdit
        Left = 70
        Top = 210
        Width = 53
        Height = 21
        EditLabel.Width = 53
        EditLabel.Height = 13
        EditLabel.Caption = 'Coluna Ate'
        TabOrder = 21
        Text = ''
      end
      object edtCde: TLabeledEdit
        Left = 6
        Top = 210
        Width = 52
        Height = 21
        EditLabel.Width = 48
        EditLabel.Height = 13
        EditLabel.Caption = 'Coluna de'
        TabOrder = 22
        Text = ''
      end
      object edtColuna: TLabeledEdit
        Left = 6
        Top = 133
        Width = 34
        Height = 21
        CharCase = ecUpperCase
        EditLabel.Width = 33
        EditLabel.Height = 13
        EditLabel.Caption = 'Coluna'
        TabOrder = 23
        Text = ''
      end
      object edtLAte: TLabeledEdit
        Left = 186
        Top = 210
        Width = 40
        Height = 21
        EditLabel.Width = 41
        EditLabel.Height = 13
        EditLabel.Caption = 'linha ate'
        NumbersOnly = True
        TabOrder = 24
        Text = ''
      end
      object edtLde: TLabeledEdit
        Left = 134
        Top = 210
        Width = 42
        Height = 21
        EditLabel.Width = 40
        EditLabel.Height = 13
        EditLabel.Caption = 'Linha de'
        NumbersOnly = True
        TabOrder = 25
        Text = ''
      end
      object edtLinha: TLabeledEdit
        Left = 46
        Top = 133
        Width = 35
        Height = 21
        EditLabel.Width = 25
        EditLabel.Height = 13
        EditLabel.Caption = 'Linha'
        NumbersOnly = True
        TabOrder = 26
        Text = ''
      end
      object edtNomeGrafico: TLabeledEdit
        Left = 236
        Top = 210
        Width = 181
        Height = 21
        EditLabel.Width = 64
        EditLabel.Height = 13
        EditLabel.Caption = 'Nome Grafico'
        TabOrder = 27
        Text = ''
      end
      object edtPos: TLabeledEdit
        Left = 413
        Top = 139
        Width = 34
        Height = 21
        EditLabel.Width = 36
        EditLabel.Height = 13
        EditLabel.Caption = 'Posi'#231#227'o'
        TabOrder = 28
        Text = ''
      end
      object edtSalvar: TLabeledEdit
        Left = 3
        Top = 296
        Width = 668
        Height = 21
        Anchors = [akLeft, akTop, akRight]
        EditLabel.Width = 47
        EditLabel.Height = 13
        EditLabel.Caption = 'Salvar em'
        TabOrder = 29
        Text = 'c:\'
        ExplicitWidth = 612
      end
      object edtValor: TLabeledEdit
        Left = 87
        Top = 132
        Width = 221
        Height = 21
        EditLabel.Width = 24
        EditLabel.Height = 13
        EditLabel.Caption = 'Valor'
        TabOrder = 30
        Text = ''
      end
      object Graficos: TGroupBox
        Left = 405
        Top = 44
        Width = 123
        Height = 70
        Caption = 'Tipos de Graficos'
        TabOrder = 31
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
      object GroupBox1: TGroupBox
        Left = 405
        Top = 3
        Width = 362
        Height = 37
        Anchors = [akLeft, akTop, akRight]
        Caption = 'Alinhamento horizontal'
        TabOrder = 33
        ExplicitWidth = 306
        object RBhCenter: TRadioButton
          Left = 87
          Top = 18
          Width = 58
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
          Left = 160
          Top = 18
          Width = 113
          Height = 17
          Caption = 'Direita'
          TabOrder = 2
        end
        object PageControl1: TPageControl
          Left = 88
          Top = 40
          Width = 289
          Height = 193
          TabOrder = 3
        end
      end
      object GroupBox2: TGroupBox
        Left = 534
        Top = 44
        Width = 102
        Height = 70
        Caption = 'Alinha. Vertical'
        TabOrder = 32
        DesignSize = (
          102
          70)
        object RBvTop: TRadioButton
          Left = 32
          Top = 15
          Width = 113
          Height = 17
          Anchors = [akLeft, akTop, akRight]
          Caption = 'Cima'
          TabOrder = 0
        end
        object RBvBottom: TRadioButton
          Left = 32
          Top = 51
          Width = 113
          Height = 17
          Caption = 'Baixo'
          TabOrder = 1
        end
        object RBvCenter: TRadioButton
          Left = 32
          Top = 33
          Width = 113
          Height = 17
          Caption = 'Centro'
          TabOrder = 2
        end
      end
      object edtTamanhoFonte: TEdit
        Left = 6
        Top = 21
        Width = 61
        Height = 21
        NumbersOnly = True
        TabOrder = 34
      end
      object DBGrid1: TDBGrid
        Left = 3
        Top = 403
        Width = 764
        Height = 120
        Anchors = [akLeft, akTop, akRight]
        DataSource = DataSource1
        PopupMenu = PopupMenu1
        TabOrder = 35
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object Button11: TButton
        Left = 554
        Top = 363
        Width = 97
        Height = 30
        Caption = 'Pdf para planilha'
        TabOrder = 36
        OnClick = Button11Click
      end
      object Button12: TButton
        Left = 3
        Top = 534
        Width = 148
        Height = 25
        Caption = 'Planilha para DataSet'
        TabOrder = 37
        OnClick = Button12Click
      end
      object CheckBox1: TCheckBox
        Left = 236
        Top = 235
        Width = 181
        Height = 17
        Caption = 'Visualizar gera'#231#227'o do documento? '
        Checked = True
        State = cbChecked
        TabOrder = 38
        OnClick = CheckBox1Click
      end
      object edtWidth: TLabeledEdit
        Left = 314
        Top = 132
        Width = 44
        Height = 21
        CharCase = ecUpperCase
        EditLabel.Width = 48
        EditLabel.Height = 13
        EditLabel.Caption = 'Cel lWidth'
        EditLabel.Color = clBackground
        EditLabel.ParentColor = False
        ImeName = 'Portuguese (Brazilian ABNT)'
        TabOrder = 39
        Text = '5000'
      end
      object Button13: TButton
        Left = 551
        Top = 164
        Width = 80
        Height = 25
        Caption = 'Trocar - Index'
        TabOrder = 40
        OnClick = Button13Click
      end
      object Button14: TButton
        Left = 636
        Top = 164
        Width = 48
        Height = 25
        Caption = 'Exists'
        TabOrder = 41
        OnClick = Button14Click
      end
      object DBGrid3: TDBGrid
        Left = 3
        Top = 563
        Width = 765
        Height = 120
        Anchors = [akLeft, akTop, akRight, akBottom]
        DataSource = DataSource3
        TabOrder = 42
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object Button16: TButton
        Left = 673
        Top = 337
        Width = 20
        Height = 21
        Anchors = [akTop, akRight]
        Caption = '...'
        TabOrder = 43
        OnClick = Button16Click
        ExplicitLeft = 617
      end
      object Button17: TButton
        Left = 673
        Top = 296
        Width = 20
        Height = 21
        Anchors = [akTop, akRight]
        Caption = '...'
        TabOrder = 44
        OnClick = Button17Click
        ExplicitLeft = 617
      end
      object mListSheet: TMemo
        Left = 566
        Top = 197
        Width = 201
        Height = 76
        TabOrder = 45
      end
      object BtnListSheet: TButton
        Left = 453
        Top = 197
        Width = 107
        Height = 25
        Caption = 'Listar as planilhas'
        TabOrder = 46
        OnClick = BtnListSheetClick
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Writer'
      ImageIndex = 1
      OnShow = TabSheet2Show
      object Label5: TLabel
        Left = 132
        Top = 31
        Width = 73
        Height = 13
        Caption = 'Tamanho fonte'
      end
      object Label6: TLabel
        Left = 235
        Top = 31
        Width = 17
        Height = 13
        Caption = 'Cor'
      end
      object BitBtn3: TBitBtn
        Left = 48
        Top = 129
        Width = 75
        Height = 25
        Caption = 'Iniciar'
        TabOrder = 0
        OnClick = BitBtn3Click
      end
      object BitBtn4: TBitBtn
        Left = 129
        Top = 129
        Width = 75
        Height = 25
        Caption = 'Fechar'
        TabOrder = 1
        OnClick = BitBtn4Click
      end
      object BitBtn5: TBitBtn
        Left = 224
        Top = 129
        Width = 75
        Height = 25
        Caption = 'Salvar'
        TabOrder = 2
        OnClick = BitBtn5Click
      end
      object BitBtn6: TBitBtn
        Left = 305
        Top = 129
        Width = 75
        Height = 25
        Caption = 'Carregar'
        TabOrder = 3
        OnClick = BitBtn6Click
      end
      object mmo: TMemo
        Left = 48
        Top = 160
        Width = 457
        Height = 210
        Lines.Strings = (
          'Ol'#225' mundo')
        TabOrder = 4
      end
      object BitBtn7: TBitBtn
        Left = 48
        Top = 376
        Width = 129
        Height = 25
        Caption = 'Adicionar texto'
        TabOrder = 5
        OnClick = BitBtn7Click
      end
      object edtFontHg: TEdit
        Left = 132
        Top = 50
        Width = 65
        Height = 21
        NumbersOnly = True
        TabOrder = 6
        Text = '10'
      end
      object CB_Bold: TCheckBox
        Left = 48
        Top = 29
        Width = 41
        Height = 17
        Caption = 'Bold'
        TabOrder = 7
      end
      object CBErase: TCheckBox
        Left = 48
        Top = 6
        Width = 161
        Height = 17
        Caption = 'Apar texto antes de escrever'
        Checked = True
        State = cbChecked
        TabOrder = 8
      end
      object BitBtn8: TBitBtn
        Left = 401
        Top = 129
        Width = 104
        Height = 25
        Caption = 'Montar exemplo'
        TabOrder = 9
        OnClick = BitBtn8Click
      end
      object cbUnderline_wrt: TCheckBox
        Left = 48
        Top = 52
        Width = 78
        Height = 17
        Caption = 'UnderLine'
        TabOrder = 10
      end
      object cbColorWriter: TComboBox
        Left = 235
        Top = 50
        Width = 145
        Height = 21
        ItemIndex = 0
        TabOrder = 11
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
      object BitBtn10: TBitBtn
        Left = 195
        Top = 376
        Width = 198
        Height = 25
        Caption = 'Mover cursor para o final da pagina'
        TabOrder = 12
        OnClick = BitBtn10Click
      end
      object edtArqWriter: TLabeledEdit
        Left = 48
        Top = 434
        Width = 457
        Height = 21
        EditLabel.Width = 99
        EditLabel.Height = 13
        EditLabel.Caption = 'Carregar documento'
        TabOrder = 13
        Text = ''
      end
      object DBGrid2: TDBGrid
        Left = 48
        Top = 478
        Width = 577
        Height = 120
        DataSource = DataSource2
        TabOrder = 14
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object Button15: TButton
        Left = 48
        Top = 604
        Width = 149
        Height = 25
        Caption = 'Grid to Doc Table'
        TabOrder = 15
        OnClick = Button15Click
      end
    end
  end
  object ClientDataSet1: TClientDataSet
    PersistDataPacket.Data = {
      4C0000009619E0BD0100000018000000030000000000030000004C0002494404
      00010000000000044E6F6D650100490000000100055749445448020002001400
      05496461646504000100000000000000}
    Active = True
    Aggregates = <>
    Params = <>
    Left = 412
    Top = 368
    object ClientDataSet1ID: TIntegerField
      FieldName = 'ID'
    end
    object ClientDataSet1Nome: TStringField
      FieldName = 'Nome'
    end
    object ClientDataSet1Idade: TIntegerField
      FieldName = 'Idade'
    end
  end
  object DataSource1: TDataSource
    DataSet = ClientDataSet1
    Left = 492
    Top = 368
  end
  object PopupMenu1: TPopupMenu
    Left = 468
    Top = 480
    object Exportarplanilha1: TMenuItem
      Caption = 'Exportar planilha'
      OnClick = Exportarplanilha1Click
    end
  end
  object DataSetProvider1: TDataSetProvider
    UpdateMode = upWhereKeyOnly
    Left = 580
    Top = 368
  end
  object OpenOffice_calc1: TOpenOffice_calc
    DocVisible = True
    Left = 560
    Top = 280
  end
  object OpenOffice_writer1: TOpenOffice_writer
    DocVisible = True
    Left = 468
    Top = 272
  end
  object FDMemWriter: TFDMemTable
    Active = True
    FetchOptions.AssignedValues = [evMode]
    FetchOptions.Mode = fmAll
    ResourceOptions.AssignedValues = [rvPersistent, rvSilentMode]
    ResourceOptions.Persistent = True
    ResourceOptions.SilentMode = True
    UpdateOptions.AssignedValues = [uvCheckRequired, uvAutoCommitUpdates]
    UpdateOptions.CheckRequired = False
    UpdateOptions.AutoCommitUpdates = True
    Left = 628
    Top = 496
    Content = {
      414442531000000053020000FF00010001FF02FF03040016000000460044004D
      0065006D00570072006900740065007200050016000000460044004D0065006D
      00570072006900740065007200060000000000070000080032000000090000FF
      0AFF0B0400080000004E0061006D0065000500080000004E0061006D0065000C
      00010000000E000D000F00140000001000011100011200011300011400011500
      011600080000004E0061006D006500170014000000FEFF0B04000E0000004100
      64006400720065007300730005000E0000004100640064007200650073007300
      0C00020000000E000D000F003200000010000111000112000113000114000115
      000116000E0000004100640064007200650073007300170032000000FEFF0B04
      000A00000045006D00610069006C0005000A00000045006D00610069006C000C
      00030000000E000D000F001E0000001000011100011200011300011400011500
      0116000A00000045006D00610069006C0017001E000000FEFEFF18FEFF19FEFF
      1AFF1B1C0000000000FF1D00000800000044696E6F7344657601000600000042
      72617A696C02001900000064616E69656C2E64696E6F7364657640676D61696C
      2E636F6DFEFEFF1B1C0001000000FF1D00000E00000044616E69656C202D2044
      696E6F730100060000004272617A696C02001000000064616E69656C40676D61
      696C2E636F6DFEFEFF1B1C0002000000FF1D0000040000004A6F736501000600
      00004272617A696C0200100000006A736F736540676D61696C2E2E636F6DFEFE
      FEFEFEFF1EFEFF1F200003000000FF21FEFEFE0E004D0061006E006100670065
      0072001E00550070006400610074006500730052006500670069007300740072
      00790012005400610062006C0065004C006900730074000A005400610062006C
      00650008004E0061006D006500140053006F0075007200630065004E0061006D
      0065000A0054006100620049004400240045006E0066006F0072006300650043
      006F006E00730074007200610069006E00740073001E004D0069006E0069006D
      0075006D0043006100700061006300690074007900180043006800650063006B
      004E006F0074004E0075006C006C00140043006F006C0075006D006E004C0069
      00730074000C0043006F006C0075006D006E00100053006F0075007200630065
      004900440018006400740041006E007300690053007400720069006E00670010
      00440061007400610054007900700065000800530069007A0065001400530065
      006100720063006800610062006C006500120041006C006C006F0077004E0075
      006C006C000800420061007300650014004F0041006C006C006F0077004E0075
      006C006C0012004F0049006E0055007000640061007400650010004F0049006E
      00570068006500720065001A004F0072006900670069006E0043006F006C004E
      0061006D006500140053006F007500720063006500530069007A0065001C0043
      006F006E00730074007200610069006E0074004C006900730074001000560069
      00650077004C006900730074000E0052006F0077004C00690073007400060052
      006F0077000A0052006F0077004900440010004F0072006900670069006E0061
      006C001800520065006C006100740069006F006E004C006900730074001C0055
      007000640061007400650073004A006F00750072006E0061006C001200530061
      007600650050006F0069006E0074000E004300680061006E00670065007300}
    object FDMemWriterName: TStringField
      FieldName = 'Name'
    end
    object FDMemWriterAddress: TStringField
      FieldName = 'Address'
      Size = 50
    end
    object FDMemWriterEmail: TStringField
      FieldName = 'Email'
      Size = 30
    end
  end
  object DataSource2: TDataSource
    DataSet = FDMemWriter
    Left = 596
    Top = 512
  end
  object DataSource3: TDataSource
    DataSet = CdsDados
    Left = 212
    Top = 616
  end
  object CdsDados: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 256
    Top = 632
  end
  object Odlg: TOpenDialog
    Left = 668
    Top = 336
  end
end
