Attribute VB_Name = "MathUnit"
Option Explicit

'
' Funkcja zwraca najmniejsz¹ liczbê ca³kowit¹ wiêksz¹ od lub równ¹ dane
'
Public Function Ceil(ByVal number As Variant) As Integer
    Dim Value As Integer: Value = Int(number)
    Ceil = IIf(Value <> number, Int(number + 1#), Int(number))
End Function


'
' Funkcja zwraca najwiêksz¹ liczbê ca³kowit¹ mniejsz¹ od lub równ¹ danej
'
Public Function Floor(ByVal number As Variant) As Integer
    Dim Value As Integer: Value = Int(number)
    Floor = IIf(number < Value, Value - 1#, Value)
End Function
