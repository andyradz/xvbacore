Attribute VB_Name = "Comparable"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------
' Zbiór mo¿liwych wyników porównania dwóch obiektów
'----------------------------------------------------------------------------------------------------------------------
Public Enum CompareResult
    Equals = 0   'zawartoœæ obiektów równa
    Lesser = -1  'wartoœæ obiektu z lewej strony mniejsza ni¿ wartoœæ obiektu z prawej strony
    Greater = 1  'wartoœæ obiektu z lewej strony wiêksza ni¿ wartoœæ obiektu z prawej strony
    Unknown = -2 'nieokreœlony wynik porówania wartoœci obiektów
End Enum
'----------------------------------------------------------------------------------------------------------------------

