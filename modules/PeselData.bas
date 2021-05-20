Attribute VB_Name = "PeselData"
Option Explicit

'RRMMDDPPPPK

'https://obywatel.gov.pl/pl/dokumenty-i-dane-osobowe/czym-jest-numer-pesel

'Miesi¹c Stulecie
'1800–1899   1900–1999   2000–2099   2100–2199   2200–2299
'Styczeñ        81  01  21  41  61
'Luty           82  02  22  42  62
'Marzec         83  03  23  43  63
'Kwiecieñ       84  04  24  44  64
'Maj            85  05  25  45  65
'Czerwiec       86  06  26  46  66
'Lipiec         87  07  27  47  67
'Sierpieñ       88  08  28  48  68
'Wrzesieñ       89  09  29  49  69
'PaŸdziernik    90  10  30  50  70
'Listopad       91  11  31  51  71
'Grudzieñ       92  12  32  52  72

'----------------------------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------------------------
Public Enum Months
    January = 1
    February = 2
    March = 3
    April = 4
    May = 5
    June = 6
    July = 7
    August = 8
    September = 9
    October = 10
    November = 11
    December = 12
End Enum
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
'Index wieków dla danej epoki
'----------------------------------------------------------------------------------------------------------------------
Private Enum Wiek
    W20
    W21
    W22
    W23
End Enum
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------------------------
Public Type Century
    month As Months
    centuries() As Variant
End Type
'----------------------------------------------------------------------------------------------------------------------

Public Type peselConfig
    weights() As Variant
    centuries() As Century
End Type

'----------------------------------------------------------------------------------------------------------------------
' Struktura sk³adowych numeru identyfikacji PESEL
'----------------------------------------------------------------------------------------------------------------------
Public Type PeselData
    Year As String * 2      'dwie ostatnie cyfry roku
    month As String * 2     'numer miesi¹ca
    day As String * 2       'numer dnia
    Gender As String * 4    'ostatni znak na p³eæ
    checksum As String * 1  'cyfra kontrolna
End Type
'----------------------------------------------------------------------------------------------------------------------


Public Function GetPeselConfig() As peselConfig
    Dim config As peselConfig
    
    config.weights = GetPeselWeights
    config.centuries = GetPeselCenturies
    
    GetPeselConfig = config
End Function

Private Function GetPeselCenturies() As Century()
    Dim centuries(January To December) As Century
        
    With centuries(January)
        .month = January
        .centuries = Array(81, 1, 21, 41, 61)
    End With
    With centuries(February)
        .month = February
        .centuries = Array(82, 2, 22, 42, 62)
    End With
    With centuries(March)
        .month = March
        .centuries = Array(83, 3, 23, 43, 63)
    End With
    With centuries(April)
        .month = April
        .centuries = Array(84, 4, 24, 44, 64)
    End With
    With centuries(May)
        .month = May
        .centuries = Array(85, 5, 25, 45, 65)
    End With
    With centuries(June)
        .month = June
        .centuries = Array(86, 6, 26, 46, 66)
    End With
    With centuries(July)
        .month = July
        .centuries = Array(87, 7, 27, 47, 67)
    End With
    With centuries(August)
        .month = August
        .centuries = Array(88, 8, 28, 48, 68)
    End With
    With centuries(September)
        .month = September
        .centuries = Array(89, 9, 29, 49, 69)
    End With
    With centuries(October)
        .month = October
        .centuries = Array(90, 10, 30, 50, 70)
    End With
    With centuries(November)
        .month = November
        .centuries = Array(91, 11, 31, 51, 71)
    End With
    With centuries(December)
        .month = December
        .centuries = Array(92, 12, 32, 52, 72)
    End With
    
    'return
    GetPeselCenturies = centuries
End Function

Private Function GetPeselWeights() As Variant()
    Dim weights() As Variant: weights = Array(1, 3, 7, 9, 1, 3, 7, 9, 1, 3)
    
    'return
    GetPeselWeights = weights
End Function



