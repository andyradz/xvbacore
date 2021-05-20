VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Backlog
'dodaæ zapis pliku do logu

Private logFile As LoggerFile
Private m_config As LogConfig

Property Let SetConfig(config As LogConfig)
    m_config = config
    logFile.SetConfig = m_config
End Property

'
'
'
Public Function info(logMessage As String) As String
    Call logFile.Append(Prefix & logMessage)
End Function

'
'
'
Public Function warm(logMessage As String) As String
    Call logFile.Append(Prefix & logMessage)
End Function

'
'
'
Public Function error(logMessage As String) As String
    Call logFile.Append(Prefix & logMessage)
End Function

Private Function Prefix() As String
    Prefix = CStr(Now) & "[" & "INFO" & "]" & ":" & "[" & Application.Worksheets.Application & "]" & " - "
End Function

'
'
'
Private Sub Class_Initialize()
    Set logFile = New LoggerFile
End Sub

'
'
'
Private Sub Class_Terminate()

End Sub
