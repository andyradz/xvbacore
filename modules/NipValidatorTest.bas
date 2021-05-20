Attribute VB_Name = "NipValidatorTest"
Option Explicit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Procedura uruchamia testy jednostkowe walidacji identyfikatora NIP
' Andrzej Radziszewski
' 2021-05-14
' Released
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TestsSuiteNipValidation()

    Dim nipEntity As Nip
    Dim nipChecked As New NipValidator
       
    'oczekujemy warto�ci TRUE
    nipEntity.Value = "7622654927"
    Debug.Assert nipChecked.Validate(nipEntity)
          
    'oczekujemy warto�ci TRUE
    nipEntity.Value = "1060000062"
    Debug.Assert nipChecked.Validate(nipEntity)
    
    'oczekujemy warto�ci FALSE
    nipEntity.Value = "5249116383"
    Debug.Assert Not nipChecked.Validate(nipEntity)
        
End Sub
