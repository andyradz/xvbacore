VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Regon9Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Suma kontrolna
'Dziewi�ta cyfra jest cyfr� kontroln� obliczan� nast�puj�co:
'
'Wagi: 8 9 2 3 4 5 6 7
'
'Pierwsze osiem cyfr REGON: 1 2 3 4 5 6 7 8
'
'Wynik: 1�8 + 2�9 + 3�2 + 4�3 + 5�4 + 6�5 + 7�6 + 8�7 = 192
'
'Cyfra kontrolna: 192 mod 11 = 5
'
'Ca�y REGON: 123456785
'
'W przypadku, gdy dzielenie mod 11 zwraca 10, to za cyfr� kontroln� przyjmuje si� 0.
'
'14-cyfrowy REGON
'Jednostkom organizacyjnym, kt�re maj� swoje siedziby w r�nych wojew�dztwach nadaje si� REGON 14-cyfrowy (identyfikator jednostki lokalnej). Dziewi�� pierwszych cyfr takiego numeru pokrywa si� z REGON-em jednostki macierzystej. Cyfra kontrolna numeru obliczana jest tym samym sposobem, ale wagi s� inne. Poniewa� jeden ze wsp�czynnik�w wagowych wynosi zero to weryfikacja �d�ugiego� REGON-u wymaga weryfikacji 9 cyfr z wagami jak dla 9-cyfrowego numeru, a nast�pnie weryfikacji d�ugiego REGON-u z innymi wagami.
'
'Suma kontrolna
'W przypadku REGON-�w 14-cyfrowych cyfr� kontroln� oblicza si� nast�puj�co:
'
'Wagi: 2 4 8 5 0 9 7 3 6 1 2 4 8
'
'Pierwsze trzyna�cie cyfr REGON: 1 2 3 4 5 6 7 8 5 1 2 3 4
'
'Wynik: 1�2 + 2�4 + 3�8 + 4�5 + 5�0 + 6�9 + 7�7 + 8�3 + 5�6 + 1�1 + 2�2 + 3�4 + 4�8 = 260
'
'Cyfra kontrolna: 260 mod 11 = 7
'
'Ca�y REGON: 12345678512347
