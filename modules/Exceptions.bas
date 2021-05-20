Attribute VB_Name = "Exceptions"
Option Explicit

Const BaseErrorCode = 1
Const InvalidParamTypeErrorCode = -2147483648# + BaseErrorCode


'
' Weryfikacja czy typ Variant jest oczekiwanym typem domyœlnym
'
Public Sub InvalidParamType(ByVal instance As Variant, ByVal Class As VbVarType)

    'Call Reflection.GetCodeModule(Application.VBE.ActiveCodePane.CodeModule.name)

    If Not (varType(instance) = Class) Then
        Call Err.Raise(InvalidParamTypeErrorCode, _
                       "StringComparable", _
                       "Niespójne typy danych w procedurze porównania!")
    End If
        
End Sub

'
' Weryfikacja czy typ Variant jest oczekiwanym typem domyœlnym
'
Public Sub InvalidParamTypes(ByVal instance As Variant, ParamArray Classes() As Variant)
    
    Dim Class As Variant
    
    For Each Class In Classes
    
        If varType(instance) = Class Then
            Exit Sub
        End If
    Next
    
    Call Err.Raise(InvalidParamTypeErrorCode, _
                   "StringComparable", _
                   "Niespójne typy danych w procedurze porównania!")
End Sub
