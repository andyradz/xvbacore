VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuickSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISortable

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub


Private Sub sort(items() As Variant, lRange As Long, hRange As Long)
    
    Dim lowBound As Long: lowBound = lRange
    Dim higBound As Long: higBound = hRange
        
    Dim pivotVal As Variant
    Dim vSwap As Variant
    Dim tmpLow As Long
    Dim tmpHi As Long
 
    tmpLow = lowBound
    tmpHi = higBound
    
    pivotVal = items((lowBound + higBound) \ 2)
 
    While (tmpLow <= tmpHi) 'divide
    
        While (items(tmpLow) < pivotVal And tmpLow < higBound)
            tmpLow = tmpLow + 1
        Wend
  
        While (pivotVal < items(tmpHi) And tmpHi > lowBound)
            tmpHi = tmpHi - 1
        Wend
 
        If (tmpLow <= tmpHi) Then
            vSwap = items(tmpLow)
            items(tmpLow) = items(tmpHi)
            items(tmpHi) = vSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
 
    If (lowBound < tmpHi) Then
        Call sort(items, lowBound, tmpHi)  'conquer
    End If
    
    If (tmpLow < higBound) Then
        Call sort(items, tmpLow, higBound)  'conquer
    End If
End Sub


Private Sub ISortable_Sort(items() As Variant)
    'Sorts a one-dimensional VBA array from smallest to largest
    'using a very fast quicksort algorithm variant.
    
    Call sort(items, LBound(items), UBound(items))
End Sub






