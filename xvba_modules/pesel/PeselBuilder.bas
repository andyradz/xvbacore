Attribute VB_Name = "PeselBuilder"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------
' Podstawowa metoda tworząca instancję identyfikatora Pesel
'----------------------------------------------------------------------------------------------------------------------
Public Function CreatePesel(ByVal Identyfier) As Pesel
	Dim peselEntity As New Pesel
  	peselEntity.Identyfier = Identyfier
  	Set CreatePesel = peselEntity
End Function
'----------------------------------------------------------------------------------------------------------------------