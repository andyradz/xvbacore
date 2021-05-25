Attribute VB_Name = "MathUnitTest"
Option Explicit

Public Sub TestsSuiteMathFunctions()
    Call TestsCeilFunction
    Call TestsFloorFunction
End Sub

Private Sub TestsFloorFunction()
    Debug.Assert MathUnit.Floor(45.95) = 45
    Debug.Assert MathUnit.Floor(-45.95) = -46
    Debug.Assert MathUnit.Floor(-1001.0912) = -1002
End Sub

Private Sub TestsCeilFunction()
    Debug.Assert MathUnit.Ceil(-5) = -5
    Debug.Assert MathUnit.Ceil(0.95) = 1
    Debug.Assert MathUnit.Ceil(-0.95) = 0
    Debug.Assert MathUnit.Ceil(4) = 4
    Debug.Assert MathUnit.Ceil(7.004) = 8
    Debug.Assert MathUnit.Ceil(-45.99) = -45
    Debug.Assert MathUnit.Ceil(-4.2) = -4
End Sub
