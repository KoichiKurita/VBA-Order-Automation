Attribute VB_Name = "modFileUtil"
Option Explicit
'==========================================================
' modFileUtil : ファイル・フォルダ操作の共通関数
'==========================================================

Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(filePath) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExists = False
End Function

' 親フォルダも含めて再帰的に作成
Public Sub EnsureFolder(ByVal folderPath As String)
    Dim parent As String
    Dim pos As Long

    If Len(folderPath) = 0 Then Exit Sub
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    If FolderExists(folderPath) Then Exit Sub

    pos = InStrRev(folderPath, "\")
    If pos > 0 Then
        parent = Left$(folderPath, pos - 1)
        If Len(parent) > 2 And Not FolderExists(parent) Then EnsureFolder parent
    End If
    MkDir folderPath
End Sub

' "Book1.xlsm" → "Book1"
Public Function BaseName(ByVal fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then BaseName = Left$(fileName, pos - 1) Else BaseName = fileName
End Function

' "Book1.xlsm" → ".xlsm"
Public Function Extension(ByVal fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then Extension = Mid$(fileName, pos) Else Extension = ""
End Function

