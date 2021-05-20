Attribute VB_Name = "ValueFormats"
Option Explicit

'Long Time
'Medium Time
'Short Time

Public Enum DecimalFormats
    ShortDecimal
    MediumDecimal
    LongDecimal
End Enum

Public Enum DateFormats
    ShortDate
    MediumDate
    LongDate
End Enum

Public Enum TimeFormats
    ShortTime
    MediumTime
    LongTime
End Enum


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FormatDecimal(number As Variant, Optional pattern As DecimalFormats = DecimalFormats.ShortDecimal) As String
    On Error GoTo Catch_All
    
    FormatDecimal = ""
    
    Dim patterns As Variant: patterns = Array("#,##0.00", "#,##0.0000", "#,##0.000000")
       
    If IsEmpty(number) Then
        GoTo Finally_IsEmpty
    End If
    
    If varType(number) = vbObject Then
        If number Is Nothing Then
            GoTo Finally_IsNothing
        End If
    End If
    
    If Not (IsNumeric(number)) Then
        GoTo Finally_IsNotNumber
    End If
        
    FormatDecimal = format(number, patterns(pattern))
        
Existing:
    Exit Function
        
Finally_IsEmpty:
    Debug.Print "Passed value is Empty!!!"
    GoTo Existing
    
Finally_IsNothing:
    Debug.Print "Passed value is Nothing!!!"
    GoTo Existing
        
Finally_IsNotNumber:
    Call Err.Raise(vbObjectError + 513, "ValueFormats", "FormatDecimal")
    GoTo Existing
        
Catch_All:
    Debug.Print "[" & "Err.Code=" & Err.number & "]" & " " & Err.Description
                
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FormatDate(instant As Variant, Optional pattern As DateFormats = DateFormats.ShortDate) As String
    On Error GoTo Catch_All
    
    FormatDate = ""
    
    Dim patterns As Variant: patterns = Array("dd.mm.yyyy")
       
    If IsEmpty(instant) Then
        GoTo Finally_IsEmpty
    End If
    
    If varType(instant) = vbObject Then
        If instant Is Nothing Then
            GoTo Finally_IsNothing
        End If
    End If
    
    If Not (IsDate(instant)) Then
        GoTo Finally_IsNotDate
    End If
        
    FormatDate = format(instant, patterns(pattern))
        
Existing:
    Exit Function
        
Finally_IsEmpty:
    Debug.Print "Passed value is Empty!!!"
    GoTo Existing
    
Finally_IsNothing:
    Debug.Print "Passed value is Nothing!!!"
    GoTo Existing
        
Finally_IsNotDate:
    Call Err.Raise(vbObjectError + 513, "ValueFormats", "FormatDate")
    GoTo Existing
        
Catch_All:
    Debug.Print "[" & "Err.Code=" & Err.number & "]" & " " & Err.Description
                
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FormatTime(instant As Variant, Optional pattern As TimeFormats = TimeFormats.ShortTime) As String
    On Error GoTo Catch_All
    
    FormatTime = ""
    
    Dim patterns As Variant: patterns = Array("hh.mm.ss")
       
    If IsEmpty(instant) Then
        GoTo Finally_IsEmpty
    End If
    
    If varType(instant) = vbObject Then
        If instant Is Nothing Then
            GoTo Finally_IsNothing
        End If
    End If
    
    If Not (IsDate(instant)) Then
        GoTo Finally_IsNotTime
    End If
        
    FormatTime = format(instant, patterns(pattern))
        
Existing:
    Exit Function
        
Finally_IsEmpty:
    Debug.Print "Passed value is Empty!!!"
    GoTo Existing
    
Finally_IsNothing:
    Debug.Print "Passed value is Nothing!!!"
    GoTo Existing
        
Finally_IsNotTime:
    Call Err.Raise(vbObjectError + 513, "ValueFormats", "FormatTime")
    GoTo Existing
        
Catch_All:
    Debug.Print "[" & "Err.Code=" & Err.number & "]" & " " & Err.Description
                
End Function
