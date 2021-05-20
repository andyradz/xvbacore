VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim pStack As Object

'
' Konstruktor obiektu klasy Stack
'
Private Sub Class_Initialize()
    Set pStack = CreateObject("System.Collections.Stack")
End Sub

'
' Destruktor obiektu klasy Stack
'
Private Sub Class_Terminate()

    If (Not pStack Is Nothing) Then
        pStack.Clear
    End If
    
    Set pStack = Nothing
End Sub

Public Function Push(Value As Variant)
    pStack.Push Value
End Function

Public Function Pop() As Variant
    Pop = pStack.Pop
End Function

Public Function Count() As Long
    Count = pStack.Count
End Function

Public Function ToArray() As Variant()
    ToArray = pStack.ToArray()
End Function

Public Function GetHashCode1() As Integer
    GetHashCode1 = pStack.GetHashCode
End Function

Public Function Clear() As Object
   pStack.Clear
End Function

Public Function Contains(obj As Variant) As Object
  Contains = pStack.Contains(obj)
End Function

