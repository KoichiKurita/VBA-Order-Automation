Attribute VB_Name = "modBackup"
Option Explicit
'==========================================================
' modBackup : ブックの世代バックアップ
'   Config!BackupDir に タイムスタンプ付きで保存し、
'   Config!BackupGenerations 世代を超えた古いものを自動削除
'==========================================================

Public Sub CreateBackup()
    Const PROC As String = "modBackup.CreateBackup"
    Dim folder As String
    Dim destPath As String
    Dim keep As Long

    If Len(ThisWorkbook.Path) = 0 Then
        RaiseAppError ERR_FILE_NOT_FOUND, PROC, "ブックが未保存のためバックアップできません。先に保存してください。"
    End If

    folder = ResolvePath("BackupDir", "backup")
    EnsureFolder folder

    destPath = folder & "\" & BaseName(ThisWorkbook.Name) & "_" & _
               Format$(Now, "yyyymmdd_hhnnss") & Extension(ThisWorkbook.Name)

    ThisWorkbook.SaveCopyAs destPath
    LogInfo PROC, "バックアップ作成: " & destPath

    keep = CLng(GetConfig("BackupGenerations", 10))
    PurgeOldBackups folder, BaseName(ThisWorkbook.Name) & "_*" & Extension(ThisWorkbook.Name), keep
End Sub

Private Sub PurgeOldBackups(ByVal folder As String, ByVal pattern As String, ByVal keep As Long)
    Const PROC As String = "modBackup.PurgeOldBackups"
    Dim names() As String
    Dim n As Long
    Dim f As String
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    f = Dir$(folder & "\" & pattern)
    Do While Len(f) > 0
        ReDim Preserve names(n)
        names(n) = f
        n = n + 1
        f = Dir$
    Loop
    If n <= keep Then Exit Sub

    ' ファイル名のタイムスタンプで昇順ソート（件数が少ないため単純ソート）
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If names(i) > names(j) Then
                tmp = names(i)
                names(i) = names(j)
                names(j) = tmp
            End If
        Next j
    Next i

    For i = 0 To n - keep - 1
        Kill folder & "\" & names(i)
        LogInfo PROC, "古いバックアップを削除: " & names(i)
    Next i
End Sub


