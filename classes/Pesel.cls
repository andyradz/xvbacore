VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pesel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_identyfier As String * 11 'dodaæ typ

'
'
'
Private Sub Class_Initialize()

End Sub

'
'
'
Private Sub Class_Terminate()

End Sub

Public Property Let Identyfier(Value As String)
    m_identyfier = Value
End Property


Public Property Get Identyfier() As String
    Identyfier = m_identyfier
End Property

Public Property Get Year() As Integer
    Year = m_pesel.Year
End Property

Public Property Get Sex() As Gender
    Dim genderCode As Integer: genderCode = 0
    
    genderCode = CInt(Right(m_pesel.Gender, 1) Mod 2)
    
    Sex = IIf(0 = genderCode, Gender.Female, Gender.Male)
End Property
