'----------------------------------------------------------------------------------------------------------------------
' Podstawowa struktura numeru identyfikacji PESEL
'----------------------------------------------------------------------------------------------------------------------
Public Type PeselData
	pesel As String * 11
	year As String * 2      'dwie ostatnie cyfry roku
	month As String * 2     'numer miesiąca
	day As String * 2       'numer dnia
	gender As String * 4    'rodzaj płci
	checksum As String * 1  'cyfra kontrolna
End Type
'----------------------------------------------------------------------------------------------------------------------