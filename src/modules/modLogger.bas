Attribute VB_Name = "modLogger"
Option Explicit
'==========================================================
' modLogger : ファイルログ出力
'   形式  : TSV（日時 / レベル / 実行ID / ユーザー / 発生箇所 / メッセージ）
'   保存先: Config!LogDir\app_yyyymmdd.log（日次ローテーション）
'   保持  : Config!LogRetentionDays 日を超えたものを自動削除
'   セキュリティ: 改行・タブを無害化しログ改ざん（ログインジェクション）を防止
'==========================================================

Public Enum LogLevel
    llDebug = 10
    llInfo = 20
    llWarn = 30
    llError = 40
End Enum

Private mRunId As String

' 1回の実行を識別するID（ユーザーへの問い合わせ番号にも使う）
Public Function RunId() As String
    If Len(mRunId) = 0 Then
        Randomize
        mRunId = Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int(Rnd * 10000), "0000")
    End If
    RunId = mRunId
End Function

Public Sub ResetRunId()
    mRunId = vbNullString
End Sub

Public Sub LogDebug(ByVal source As String, ByVal message As String)
    WriteLog llDebug, source, message
End Sub

Public Sub LogInfo(ByVal source As String, ByVal message As String)
    WriteLog llInfo, source, message
End Sub

Public Sub LogWarn(ByVal source As String, ByVal message As String)
    WriteLog llWarn, source, message
End Sub

Public Sub LogError(ByVal source As String, ByVal message As String)
    WriteLog llError, source, message
End Sub

Public Function CurrentLogFile() As String
    Dim folder As String
    folder = ResolvePath("LogDir", "logs")
    EnsureFolder folder
    CurrentLogFile = folder & "\app_" & Format$(Date, "yyyymmdd") & ".log"
End Function

' 保持期間を過ぎたログを削除
Public Sub CleanupOldLogs()
    Dim folder As String
    Dim f As String
    Dim days As Long
    Dim targets As Collection
    Dim item As Variant

    folder = ResolvePath("LogDir", "logs")
    days = CLng(GetConfig("LogRetentionDays", 30))
    If Not FolderExists(folder) Then Exit Sub

    ' 列挙中に削除しないよう、いったん対象を収集
    Set targets = New Collection
    f = Dir$(folder & "\app_*.log")
    Do While Len(f) > 0
        If DateDiff("d", FileDateTime(folder & "\" & f), Now) > days Then targets.Add f
        f = Dir$
    Loop

    For Each item In targets
        Kill folder & "\" & item
    Next item

    If targets.Count > 0 Then
        LogInfo "modLogger.CleanupOldLogs", targets.Count & "件の古いログを削除しました"
    End If
End Sub

Private Sub WriteLog(ByVal level As LogLevel, ByVal source As String, ByVal message As String)
    Dim lineText As String
    Dim fileNo As Integer
    On Error GoTo Fallback

    If level < CurrentLevel() Then Exit Sub

    lineText = Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & _
               LevelName(level) & vbTab & _
               RunId() & vbTab & _
               Environ$("USERNAME") & vbTab & _
               Sanitize(source) & vbTab & _
               Sanitize(message)

    fileNo = FreeFile
    Open CurrentLogFile() For Append As #fileNo
    Print #fileNo, lineText
    Close #fileNo

    Debug.Print lineText
    Exit Sub

Fallback:
    ' ログ出力の失敗で本処理を止めない
    On Error Resume Next
    Close #fileNo
    Debug.Print "[LOG-FAIL] " & Err.Description & " | " & lineText
End Sub

Private Function CurrentLevel() As LogLevel
    Select Case UCase$(CStr(GetConfig("LogLevel", "INFO")))
        Case "DEBUG": CurrentLevel = llDebug
        Case "WARN":  CurrentLevel = llWarn
        Case "ERROR": CurrentLevel = llError
        Case Else:    CurrentLevel = llInfo
    End Select
End Function

Private Function LevelName(ByVal level As LogLevel) As String
    Select Case level
        Case llDebug: LevelName = "DEBUG"
        Case llInfo:  LevelName = "INFO"
        Case llWarn:  LevelName = "WARN"
        Case llError: LevelName = "ERROR"
    End Select
End Function

' 改行・タブを除去し、1イベント=1行を保証する
Private Function Sanitize(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Sanitize = s
End Function


