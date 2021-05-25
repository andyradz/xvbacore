Attribute VB_Name = "PeselData"
Option Explicit


'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
' Index epoki dla danych miesięcy
'----------------------------------------------------------------------------------------------------------------------
Private Enum Epoch
	E20
	E21
	E22
	E23
End Enum
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
' Model epoki numeru Pesel
'----------------------------------------------------------------------------------------------------------------------
Public Type Century
	month As Months
	centuries()As Variant
End Type
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
' Funkcja tworzy podstawowe dane konfiguracyjne identyfikatora PESEL
'----------------------------------------------------------------------------------------------------------------------
Public Function GetPeselConfig()As peselConfig
	Dim config As PeselConfig
    
  	config.weights = GetPeselWeights
  	config.centuries = GetPeselCenturies

    '___return
  	GetPeselConfig = config
End Function

'----------------------------------------------------------------------------------------------------------------------
' Funkcja tworzy przeliczniki w epokach dla kontrentych miesięcy
'----------------------------------------------------------------------------------------------------------------------
Private Function GetPeselCenturies()As Century()
	Dim centuries(January To December)As Century
        
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
		
	'___return
	GetPeselCenturies = centuries
End Function

'
' Funkcja tworzy wagi dla poszczególnych pozycji identyfikatora
'
Private Function GetPeselWeights()As Variant()
  Dim weights()As Variant:weights = Array(1, 3, 7, 9, 1, 3, 7, 9, 1, 3)
    
  '___return
  GetPeselWeights = weights
End Function