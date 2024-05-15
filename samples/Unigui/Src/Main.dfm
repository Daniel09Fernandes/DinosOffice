object MainForm: TMainForm
  Left = 0
  Top = 0
  ClientHeight = 441
  ClientWidth = 624
  Caption = 'MainForm'
  OldCreateOrder = False
  MonitoredKeys.Keys = <>
  TextHeight = 15
  object UniButton1: TUniButton
    Left = 176
    Top = 160
    Width = 129
    Height = 73
    Hint = ''
    Caption = 'Create Sheet'
    TabOrder = 0
    OnClick = UniButton1Click
  end
  object UniButton2: TUniButton
    Left = 40
    Top = 160
    Width = 121
    Height = 73
    Hint = ''
    Caption = 'Read Sheet'
    TabOrder = 1
    OnClick = UniButton2Click
  end
  object UniButton3: TUniButton
    Left = 328
    Top = 160
    Width = 121
    Height = 73
    Hint = ''
    Caption = 'Close Sheet'
    TabOrder = 2
    OnClick = UniButton3Click
  end
  object UniButton4: TUniButton
    Left = 176
    Top = 32
    Width = 129
    Height = 57
    Hint = ''
    Caption = 'Create Doc Writer'
    TabOrder = 3
    OnClick = UniButton4Click
  end
  object UniButton5: TUniButton
    Left = 328
    Top = 32
    Width = 121
    Height = 57
    Hint = ''
    Caption = 'Close Doc'
    TabOrder = 4
    OnClick = UniButton5Click
  end
  object UniButton6: TUniButton
    Left = 40
    Top = 32
    Width = 121
    Height = 57
    Hint = ''
    Caption = 'Read Doc'
    TabOrder = 5
    OnClick = UniButton2Click
  end
  object up: TUniFileUpload
    Title = 'Upload'
    Messages.Uploading = 'Uploading...'
    Messages.PleaseWait = 'Please Wait'
    Messages.Cancel = 'Cancel'
    Messages.Processing = 'Processing...'
    Messages.UploadError = 'Upload Error'
    Messages.Upload = 'Upload'
    Messages.NoFileError = 'Please select a file'
    Messages.BrowseText = 'Browse...'
    Messages.UploadTimeout = 'Timeout occurred...'
    Messages.MaxSizeError = 'File is bigger than maximum allowed size'
    Messages.MaxFilesError = 'You can upload maximum %d files.'
    ButtonOnly = True
    OnCompleted = upCompleted
    Left = 128
    Top = 112
  end
  object OpenOffice_writer1: TOpenOffice_writer
    DocVisible = False
    Left = 528
    Top = 232
  end
end
