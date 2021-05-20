VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PeselValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValidation

Private peselConfig As peselConfig

'
'
'
Private Sub Class_Initialize()
    peselConfig = GetPeselConfig
End Sub

'
'
'
Private Sub Class_Terminate()
    
End Sub

'
'
'
Private Function IValidation_Validate(identifier As Variant) As Boolean
    Dim loopIdx As Integer, sum As Integer, checksum As Integer, controlNumber As Integer
    
    For loopIdx = LBound(peselConfig.weights) To UBound(peselConfig.weights)
        sum = sum + (peselConfig.weights(loopIdx) * CInt(Mid(identifier, loopIdx + 1, 1)))
    Next
    
    sum = sum Mod 10
    checksum = 10 - sum
    controlNumber = CInt(Right(identifier, 1))
    
    IValidation_Validate = (checksum = controlNumber)
    
End Function
