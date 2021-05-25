Attribute VB_Name = "MathUnit"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------
' Funkcja zwraca najmniejszą liczbę całkowitą większą od lub równą dane liczbie
'----------------------------------------------------------------------------------------------------------------------
Public Function Ceil(ByVal number As Variant) As Integer
    Dim Value As Integer: Value = Int(number)
	'___return
    Ceil = IIf(Value <> number, Int(number + 1#), Int(number))
End Function
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
' Funkcja zwraca największą liczbę całkowitą mniejszą od lub równą danej liczbie
'----------------------------------------------------------------------------------------------------------------------
Public Function Floor(ByVal number As Variant) As Integer
    Dim Value As Integer:Value = Int(number)
	'___return
    Floor = IIf(number < Value, Value - 1#, Value)
End Function
'----------------------------------------------------------------------------------------------------------------------
