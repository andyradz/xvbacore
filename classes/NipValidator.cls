VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NipValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValidation

Private Sub Class_Initialize()

End Sub


Private Sub Class_Terminate()

End Sub

Private Function IValidation_Validate(identifier As Variant) As Boolean

End Function

Public Function Validate(identifier As Nip) As Boolean

    Const controlNumber = 11

    On Error GoTo Catch_13

    Dim nipValue As String: nipValue = identifier.Value

    Dim weights() As Variant: weights = Array(6, 5, 7, 2, 3, 4, 5, 6, 7)
    
    Dim sum As Integer: sum = 0
    
    Dim checksum As Integer: checksum = Mid(nipValue, 10, 1)
    
    Dim loopIdx As Integer: loopIdx = 0
    
    For loopIdx = LBound(weights) To UBound(weights)
        Dim number As Integer: number = Mid(nipValue, loopIdx + 1, 1)
        sum = sum + (number * weights(loopIdx))
    Next
    
Finally:
    Validate = (sum Mod controlNumber) = checksum
    Exit Function
    
Catch_13:
    Debug.Print Mid(nipValue, loopIdx + 1, 1) & " - " & "Is Not A Number" & " : " & Err.Description
    On Error GoTo -1
    Exit Function
    

End Function

Private Sub RegexReplacingAPattern()
    Dim stringOne As String
    Dim regexOne As Object
    Set regexOne = New RegExp

    regexOne.pattern = "This is the number"
    regexOne.Global = False
    stringOne = "This is the number 718901"

    Debug.Print regexOne.Replace(stringOne, "That is the new number")
End Sub
