Attribute VB_Name = "Comparable"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------
' Zbi�r mo�liwych wynik�w por�wnania dw�ch obiekt�w
'----------------------------------------------------------------------------------------------------------------------
Public Enum CompareResult
    Equals = 0   'zawarto�� obiekt�w r�wna
    Lesser = -1  'warto�� obiektu z lewej strony mniejsza ni� warto�� obiektu z prawej strony
    Greater = 1  'warto�� obiektu z lewej strony wi�ksza ni� warto�� obiektu z prawej strony
    Unknown = -2 'nieokre�lony wynik por�wania warto�ci obiekt�w
End Enum
'----------------------------------------------------------------------------------------------------------------------

