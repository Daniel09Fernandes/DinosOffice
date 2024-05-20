object IWServerController: TIWServerController
  OldCreateOrder = False
  AuthBeforeNewSession = False
  AppName = 'MyApp'
  CacheDir = 'C:\Users\T_DANI~1.ROD\AppData\Local\Temp\'
  ComInitialization = ciMultiThreaded
  Compression.Enabled = True
  Compression.Level = 6
  Description = 'My IntraWeb Application'
  DebugHTML = False
  DisplayName = 'IntraWeb Application'
  Log = loNone
  EnableImageToolbar = False
  ExceptionDisplayMode = smAlert
  HistoryEnabled = False
  JavascriptDebug = False
  PageTransitions = False
  Port = 8888
  RedirectMsgDelay = 0
  ServerResizeTimeout = 0
  ShowLoadingAnimation = True
  SessionTimeout = 10
  SSLOptions.NonSSLRequest = nsAccept
  SSLOptions.Port = 0
  SSLOptions.SSLVersion = sslv3
  Version = '14.0.0'
  AllowMultipleSessionsPerUser = False
  IECompatibilityMode = 'IE=8'
  OnNewSession = IWServerControllerBaseNewSession
  Height = 310
  Width = 342
  object FDGUIxWaitCursor1: TFDGUIxWaitCursor
    Provider = 'Forms'
    Left = 200
    Top = 104
  end
end
