VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Swapable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISwapable


Private Sub ISwapable_SwapObject(factorLeft As Object, factorRight As Object)
 
    If factorLeft Is Nothing Then
        Err.Raise vbError
    End If
    
    If factorRight Is Nothing Then
        Err.Raise vbError
    End If
        
    Dim temp As Object
    
    Set temp = factorLeft
    Set factorLeft = factorRight
    Set factorRight = temp
   

Exiting:
  Set temp = Nothing
End Sub

Private Sub ISwapable_SwapValue(factorLeft As Variant, factorRight As Variant)
    
    Dim temp As Variant: temp = Empty
    
    temp = factorLeft
    factorLeft = factorRight
    factorRight = temp
End Sub
