Attribute VB_Name = "FileInfo"
Type CustomerRecord    ' Define user-defined type.
    ID As Integer    ' Place this definition in a
    name As String * 20    ' standard module.
    Address As String * 30
End Type

Function GetDirOrFileSize(strFolder As String, Optional strFile As Variant) As Long

'Call Sequence: GetDirOrFileSize("drive\path"[,"filename.ext"])

   Dim lngFSize As Long, lngDSize As Long
   Dim oFO As Object
   Dim oFD As Object
   Dim OFS As Object

   lngFSize = 0
   Set OFS = CreateObject("Scripting.FileSystemObject")

   If strFolder = "" Then strFolder = ActiveWorkbook.Path
   If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
   'Thanks to Jean-Francois Corbett, you can use also OFS.BuildPath(strFolder, strFile)

   If OFS.FolderExists(strFolder) Then
     If Not IsMissing(strFile) Then

       If OFS.FileExists(strFolder & strFile) Then
         Set oFO = OFS.Getfile(strFolder & strFile)
         GetDirOrFileSize = oFO.Size
       End If

       Else
        Set oFD = OFS.GetFolder(strFolder)
        GetDirOrFileSize = oFD.Size
       End If

   End If

End Function   '*** GetDirOrFileSize ***
