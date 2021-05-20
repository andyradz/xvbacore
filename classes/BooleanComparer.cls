VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BooleanComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IComparable

'----------------------------------------------------------------------------------------------------------------------
' Podstawowy konstruktor obiektu klasy
'----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()

End Sub

'----------------------------------------------------------------------------------------------------------------------
' Podstawowy destruktor obiektu klasy
'----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()

End Sub

'----------------------------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------------------------
Private Function IComparable_Compare(leftObject As Variant, rightObject As Variant) As CompareResult
    IComparable_Compare = Unknown
    
    Call Exceptions.InvalidParamType(leftObject, vbBoolean)
    Call Exceptions.InvalidParamType(rightObject, vbBoolean)
            
    If leftObject > rightObject Then
      IComparable_Compare = Greater
    ElseIf (leftObject < rightObject) Then
        IComparable_Compare = Lesser
    Else
        IComparable_Compare = Equals
    End If
    
End Function
