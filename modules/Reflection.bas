Attribute VB_Name = "Reflection"
Option Explicit

'' returns CodeModule reference needed in the GetFnOrSubName fn
'Public Function GetCodeModule(codeModuleName As String) As VBIDE.CodeModule
'    Dim VBProj As VBIDE.VBProject
'    Dim VBComp As VBIDE.VBComponent
'
'    Set VBProj = ThisWorkbook.VBProject
'    Set VBComp = VBProj.VBComponents(codeModuleName)
'
'    Set GetCodeModule = VBComp.CodeModule
'End Function
'
'' returns the name of the sub where the error occured
'Public Function GetFnOrSubName$(handlerLabel$)
'
'    Dim VBProj As VBIDE.VBProject
'    Dim VBComp As VBIDE.VBComponent
'    Dim CodeMod As VBIDE.CodeModule
'
'    Set VBProj = ThisWorkbook.VBProject
'    Set VBComp = VBProj.VBComponents(Application.VBE.ActiveCodePane.CodeModule.name)
'    Set CodeMod = VBComp.CodeModule
'
'    Dim code$
'    code = CodeMod.Lines(1, CodeMod.CountOfLines)
'
'    Dim handlerAt&
'    handlerAt = InStr(1, code, handlerLabel, vbTextCompare)
'
'    If handlerAt Then
'
'        Dim isFunction&
'        Dim isSub&
'
'        isFunction = InStrRev(Mid$(code, 1, handlerAt), "Function", -1, vbTextCompare)
'        isSub = InStrRev(Mid$(code, 1, handlerAt), "Sub", -1, vbTextCompare)
'
'        If isFunction > isSub Then
'            ' it's a function
'            GetFnOrSubName = Split(Mid$(code, isFunction, 40), "(")(0)
'        Else
'            ' it's a sub
'            GetFnOrSubName = Split(Mid$(code, isSub, 40), "(")(0)
'        End If
'
'    End If
'
'End Function
