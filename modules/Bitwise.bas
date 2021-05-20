Attribute VB_Name = "Bitwise"
Option Explicit

'
'
'
Public Function ShiftRight(ByVal Value As Long, ByVal Shift As Byte) As Long
    
    ShiftRight = Value
    
    If Shift > 0 Then
        ShiftRight = Int(ShiftRight / (2 ^ Shift))
    End If
    
End Function

'
'
'
Public Function ShiftLeft(ByVal Value As Long, ByVal Shift As Byte) As Long
        
    ShiftLeft = Value
    
    If Shift > 0 Then
        Dim i As Byte
        Dim m As Long
        
        For i = 1 To Shift
            m = ShiftLeft And &H40000000
    
            ShiftLeft = (ShiftLeft And &H3FFFFFFF) * 2
            
            If m <> 0 Then
                ShiftLeft = ShiftLeft Or &H80000000
            End If
            
        Next i
    End If
    
End Function
