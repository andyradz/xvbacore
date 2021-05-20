Attribute VB_Name = "PeselValidatorTest"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Procedura uruchamia testy jednostkowe walidacji identyfikatora PESEL
' Andrzej Radziszewski
' 2021-05-20
' Released
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TestsSuitePeselValidation()

    Dim peselEntity As Pesel
    Dim peselChecked As New IValidation
    
    Set peselChecked = New PeselValidator
           
    'oczekujemy wartoœci TRUE
    Set peselEntity = CreatePesel("79062601652")
    Debug.Assert peselChecked.Validate(peselEntity.Identyfier)
          
    'oczekujemy wartoœci TRUE
    Set peselEntity = CreatePesel("15211309284")
    Debug.Assert peselChecked.Validate(peselEntity.Identyfier)
    
    'oczekujemy wartoœci FALSE
    Set peselEntity = CreatePesel("52491163833")
    Debug.Assert Not peselChecked.Validate(peselEntity.Identyfier)
        
End Sub
 
 
