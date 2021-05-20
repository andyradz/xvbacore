Attribute VB_Name = "ComparableTest"
Option Explicit

Public Sub TestsSuitCompareTests()
    Call TestsCompareNumber
    Call TestsCompareBoolean
    Call TestsCompareString
    Call TestsCompareNotCompatibility
End Sub

Private Sub TestsCompareNumber()

    Dim comparer As IComparable
    Set comparer = New NumericComparer
            
    Debug.Assert comparer.Compare(0.099012, 0.099012) = Equals
    
    Debug.Assert comparer.Compare(0.099012, 0.099013) = Lesser
       
    Debug.Assert comparer.Compare(0.099013, 0.099011) = Greater
       
    Debug.Assert comparer.Compare(0.099013, 0.099011) = Greater
        
    Debug.Assert comparer.Compare(-0.1, 0.1) = Lesser
    
    Debug.Assert comparer.Compare(2, 1.99999) = Greater
   
End Sub

Private Sub TestsCompareBoolean()

    Dim comparer As IComparable
    Set comparer = New BooleanComparer
        
    Debug.Assert comparer.Compare(True, True) = Equals
    
    Debug.Assert comparer.Compare(False, False) = Equals
        
    Debug.Assert comparer.Compare(False, True) = Greater
        
    Debug.Assert comparer.Compare(True, False) = Lesser

End Sub

Private Sub TestsCompareString()

    Dim comparer As IComparable
    Set comparer = New StringComparer
    
    Debug.Assert comparer.Compare("", "") = Equals
       
    Debug.Assert comparer.Compare("LODY", "lody") = Equals
       
    Debug.Assert comparer.Compare("lody", "") = Greater
    
    Debug.Assert comparer.Compare("", "lody") = Lesser
    
End Sub


Private Sub TestsCompareNotCompatibility()

    Dim comparer As IComparable
    Set comparer = New StringComparer
    
    'On Error GoTo 0
    'Debug.Assert comparer.Compare("", 1) = Unknown
        
End Sub
