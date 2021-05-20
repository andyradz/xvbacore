VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Builder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_name As String


Public Sub Procedure(Optional ByVal Value As Integer)
    
    Dim sum As Integer: sum = -1
    
    If IsMissing(Value) Then
        sum = 9999
    End If
    
End Sub


Public Sub Procedure1(ParamArray items() As Variant)
    If IsEmpty(items()) Then
        Debug.Print "Itesm variable is empty!"
    End If
End Sub


Private Sub Class_Initialize()

End Sub
