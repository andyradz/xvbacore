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

Private m_pesel as PeselData

'----------------------------------------------------------------------------------------------------------------------
' Podstawowy konstruktor obiektu klasy
'----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
' Podstawowy destruktor obiektu klasy
'----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
End Sub
'----------------------------------------------------------------------------------------------------------------------

Public Property Let Identyfier(ByVal pesel As String)
	m_pesel.pesel = pesel
End Property


Public Property Get Identyfier()As String
    Identyfier = m_pesel.pesel
End Property

Public Property Get Year()As String
    Year = m_pesel.year
End Property

Public Property Get Month()As String
	Month = m_pesel.month	
End Property

Public Property Get Day()As String
	Day = m_pesel.day	
End Property

Public Property Get Sex()As Gender
    Dim genderCode As Integer:genderCode = 0
    
    genderCode = CInt(Right(m_pesel.Gender, 1)Mod 2)
    
    Sex = IIf(0 = genderCode, Gender.Female, Gender.Male)
End Property
