Attribute VB_Name = "MathUnit"
Option Explicit

'
' Funkcja zwraca najmniejsz� liczb� ca�kowit� wi�ksz� od lub r�wn� dane
'
Public Function Ceil(ByVal number As Variant) As Integer
    Dim Value As Integer: Value = Int(number)
    Ceil = IIf(Value <> number, Int(number + 1#), Int(number))
End Function


'
' Funkcja zwraca najwi�ksz� liczb� ca�kowit� mniejsz� od lub r�wn� danej
'
Public Function Floor(ByVal number As Variant) As Integer
    Dim Value As Integer: Value = Int(number)
    Floor = IIf(number < Value, Value - 1#, Value)
End Function
