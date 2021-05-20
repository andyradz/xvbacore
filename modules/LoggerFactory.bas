Attribute VB_Name = "LoggerFactory"
Option Explicit

Public Type LogConfig
    filePath As String * 255
    fileName As String * 255
    fileBatch As Long
End Type

'Public Enum LogLevel
    'info
    'warm
    'error
'End Enum


Private m_logger As Logger

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Metoda fabrykuj¹ca tworzenie instancji loggera dla ca³ej aplikacji
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateLogger(config As LogConfig) As Logger

    If m_logger Is Nothing Then
        Set m_logger = New Logger
    End If
    
    m_logger.SetConfig = config
    Set CreateLogger = m_logger
        
End Function
