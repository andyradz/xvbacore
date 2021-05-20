VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BubbleSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISortable

Private log As Logger

'
' Podstawowy konstruktor obiektu klasy
'
Private Sub Class_Initialize()
    Dim loggerConfig As LogConfig
    
    With loggerConfig
        .fileName = "log1.txt"
        .filePath = "D:\"
    End With
    
    Set log = LoggerFactory.CreateLogger(loggerConfig)
    
End Sub

'
' Podstawowy destruktor obiektu klasy
'
Private Sub Class_Terminate()
    If Not (log Is Nothing) Then
        Set log = Nothing
    End If
End Sub

'
' Wykonuje sortowanie kolekcji obiektów, korzystaj¹c z algorytmu sortowania b¹belkowego
'
Private Sub ISortable_Sort(items() As Variant)

    Dim loopIdx As Long
    Dim isChanged As Boolean
    Dim swap As ISwapable
        
    loopIdx = 0: isChanged = False: Set swap = New Swapable
            
    Do
        
        isChanged = False
        
        For loopIdx = LBound(items) To UBound(items) - 1
        
            If items(loopIdx) > items(loopIdx + 1) Then
                Call swap.SwapValue(items(loopIdx), items(loopIdx + 1))
                isChanged = True
                log.info ("Value swapped from " & _
                         items(loopIdx) & _
                         " to " & _
                         items(loopIdx + 1))
            End If
    
        Next
    
    Loop Until isChanged = False

End Sub
