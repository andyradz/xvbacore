VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_list As Object

'
'
'
Public Sub Add(item As Variant)
    Call m_list.Add(item)
End Sub

'
'
'
Public Sub Insert(item As Variant, Optional index As Variant)

    Dim posIdx As Integer: posIdx = 0
    
    If Not IsMissing(index) Then
        posIdx = index
    End If
        
    ' Insert to first position
    m_list.Insert posIdx, item
End Sub

'
'
'
Public Sub Reverse()
    m_list.Reverse
End Sub

'
'
'
Public Function Count() As Integer
    'Return Function
    Count = m_list.Count
End Function

'
'
'
Public Function Contains(ByVal item As Variant) As Boolean
    If IsEmpty(item) Then
        'Return Function
        Contains = False
        Exit Function
    End If
    'Return Function
    Contains = m_list.Contains(item)
End Function

'
'
'
Public Function ToArray() As Variant()
    'Return Function
    ToArray = m_list.ToArray()
End Function

'
' Podstawowy konstruktor obiektu klasy ArrayList
'
Private Sub Class_Initialize()
  Set m_list = CreateObject("System.Collections.ArrayList")
End Sub

'
' Podstawowy destructor obiektu klasy ArrayList
'
Private Sub Class_Terminate()

    If Not m_list Is Nothing Then
        'czyszczenie ca³ej list
        m_list.Clear
    End If
    
    Set m_list = Nothing
End Sub
