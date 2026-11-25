Attribute VB_Name = "modAudit"
Option Explicit
'==========================================================
' modAudit : 監査ログ（改ざん検知付き）
'   誰が・いつ・何をしたかを logs\audit.log に記録する
'   （アプリログと違い、保持期間による自動削除はしない）
'
'   改ざん検知：各行に「1つ前の行のハッシュ＋この行の内容」の SHA-256 を付ける（ハッシュチェーン）
'     → 途中の行を書き換える・削除すると、その行以降のハッシュが一致しなくなる
'     → VerifyAuditLog で、改ざんされた最初の行を特定できる
'   形式（TSV）：日時 / ユーザー / 操作 / 結果 / 詳細 / 実行ID / ハッシュ
'
'   限界：末尾の行を削除する・ファイルごと作り直す改ざんは、ハッシュチェーンだけでは検知できない。
'         重要な環境では、別サーバーへの転送や書込専用の保存先と組み合わせる
'==========================================================

Private Const GENESIS_HASH As String = "0000000000000000000000000000000000000000000000000000000000000000"

' 直前のハッシュのキャッシュ（毎回ファイル全体を読まないため）
Private mLastPath As String
Private mLastSize As Long
Private mLastHash As String

Public Function AuditFilePath() As String
    Dim folder As String
    folder = ResolvePath("AuditLogDir", "logs")
    EnsureFolder folder
    AuditFilePath = folder & "\audit.log"
End Function

'----------------------------------------------------------
' 監査ログを1行記録する
'   書込に失敗しても業務処理は止めないが、アプリログに必ずエラーを残す
'----------------------------------------------------------
Public Sub AuditLog(ByVal action As String, ByVal result As String, ByVal detail As String)
    On Error GoTo Fallback
    AuditLogTo AuditFilePath(), action, result, detail
    Exit Sub
Fallback:
    LogError "modAudit.AuditLog", "監査ログの書込に失敗しました: " & Err.Description & " / " & action & " / " & detail
End Sub

' 記録先を指定して1行記録する（テストで一時ファイルを使うために公開）
Public Sub AuditLogTo(ByVal filePath As String, ByVal action As String, ByVal result As String, ByVal detail As String)
    Dim body As String
    Dim prevHash As String
    Dim h As String
    Dim fileNo As Integer

    body = Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & _
           CurrentUserName() & vbTab & _
           Clean(action) & vbTab & _
           Clean(result) & vbTab & _
           Clean(detail) & vbTab & _
           RunId()
    ' ファイルは Shift_JIS で書かれるため、表せない文字はここで「?」にそろえてからハッシュを計算する
    ' （そろえないと、読み戻したときに内容が変わり、検証で不一致になる）
    body = StrConv(StrConv(body, vbFromUnicode), vbUnicode)

    prevHash = LastHash(filePath)
    h = Sha256Hex(prevHash & vbTab & body)

    fileNo = FreeFile
    Open filePath For Append As #fileNo
    Print #fileNo, body & vbTab & h
    Close #fileNo

    mLastPath = filePath
    mLastSize = FileLen(filePath)
    mLastHash = h
End Sub

'----------------------------------------------------------
' UI：監査ログを検証する（管理者）
'----------------------------------------------------------
Public Sub VerifyAuditLog()
    Const PROC As String = "modAudit.VerifyAuditLog"
    Dim badLine As Long
    Dim lineCount As Long
    On Error GoTo EH

    ResetRunId
    ReloadConfig
    RequirePermission actAdmin

    badLine = VerifyAuditFile(AuditFilePath(), lineCount)
    If badLine = 0 Then
        AuditLog "AUDIT_VERIFY", "OK", lineCount & " 行"
        MsgBox "監査ログに改ざんは見つかりませんでした。" & vbCrLf & vbCrLf & _
               "検証した行数：" & Format$(lineCount, "#,##0") & " 行" & vbCrLf & AuditFilePath(), _
               vbInformation, AppTitle()
    Else
        AuditLog "AUDIT_VERIFY", "NG", badLine & " 行目でハッシュ不一致"
        MsgBox "監査ログの " & badLine & " 行目で、ハッシュが一致しません。" & vbCrLf & _
               "この行、またはその直前の行が改ざん・削除された可能性があります。" & vbCrLf & vbCrLf & _
               AuditFilePath(), vbCritical, AppTitle()
    End If
    Exit Sub
EH:
    HandleError PROC
End Sub

' 戻り値：ハッシュが一致しない最初の行番号（0 = 改ざんなし）。lineCount に検証した行数を返す
Public Function VerifyAuditFile(ByVal filePath As String, ByRef lineCount As Long) As Long
    Dim fileNo As Integer
    Dim lineText As String
    Dim lineNo As Long
    Dim p As Long
    Dim body As String
    Dim h As String
    Dim prevHash As String

    lineCount = 0
    If Not FileExists(filePath) Then Exit Function

    prevHash = GENESIS_HASH
    fileNo = FreeFile
    Open filePath For Input As #fileNo
    Do While Not EOF(fileNo)
        Line Input #fileNo, lineText
        lineNo = lineNo + 1
        If Len(lineText) > 0 Then
            p = InStrRev(lineText, vbTab)
            If p = 0 Then
                VerifyAuditFile = lineNo
                Exit Do
            End If
            body = Left$(lineText, p - 1)
            h = Mid$(lineText, p + 1)
            If Sha256Hex(prevHash & vbTab & body) <> h Then
                VerifyAuditFile = lineNo
                Exit Do
            End If
            prevHash = h
            lineCount = lineCount + 1
        End If
    Loop
    Close #fileNo
End Function

'==========================================================
' 内部処理
'==========================================================
' 直前の行のハッシュ。ファイルサイズが前回の書込時と同じならキャッシュを使う
Private Function LastHash(ByVal filePath As String) As String
    Dim fileNo As Integer
    Dim lineText As String
    Dim lastLine As String

    If Not FileExists(filePath) Then
        LastHash = GENESIS_HASH
        Exit Function
    End If
    If filePath = mLastPath And FileLen(filePath) = mLastSize And Len(mLastHash) > 0 Then
        LastHash = mLastHash
        Exit Function
    End If

    fileNo = FreeFile
    Open filePath For Input As #fileNo
    Do While Not EOF(fileNo)
        Line Input #fileNo, lineText
        If Len(lineText) > 0 Then lastLine = lineText
    Loop
    Close #fileNo

    If Len(lastLine) = 0 Then
        LastHash = GENESIS_HASH
    Else
        LastHash = Mid$(lastLine, InStrRev(lastLine, vbTab) + 1)
    End If
End Function

' タブ・改行を空白に置き換え、1件 = 1行・列ずれなしを保証する（ログインジェクション対策）
Private Function Clean(ByVal s As String) As String
    Clean = Replace(Replace(Replace(Replace(s, vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
End Function


