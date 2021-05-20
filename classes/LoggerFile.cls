VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_counter As Long
Dim m_index As Long
Dim m_config As LogConfig
Dim m_buffer As ArrayList

Property Let SetConfig(config As LogConfig)
    m_config = config
End Property


Public Sub Create()

    Dim TextFile As Integer
    Dim filePath As String
    Dim fileName As String
    
    m_counter = m_counter + 1
    
    fileName = CStr(Date) & "_" & CStr(m_counter) & "_"
    
    'What is the file path and name for the new text file?
    filePath = m_config.filePath & m_config.fileName

    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile

    'Open the text file
    Open filePath For Output As TextFile

    Dim idx As Integer: idx = 0

    'Write some lines of text
    For idx = 1 To 1000
      m_index = m_index + 1
      Print #TextFile, "#" & "(" & CStr(m_index) & ")" & CStr(Now)
    Next
      
    'Save & Close Text File
    Close TextFile

End Sub

Public Sub Append(logMessage As String)

    Dim TextFile As Integer

    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile

    'Open the text file
    Open LogPath For Append As TextFile
    
    m_counter = m_counter + 1

    'Write some lines of text
    Print #TextFile, "{" & LogCounter & "}" & ":" & logMessage, """"
      
    'Save & Close Text File
    Close TextFile

End Sub

Public Property Get LogPath() As String
    LogPath = Trim(m_config.filePath) & Trim(m_config.fileName)
End Property


Public Property Get LogCounter() As Long
    LogCounter = m_counter
End Property

Private Sub Class_Initialize()
    Set m_buffer = New ArrayList
End Sub
