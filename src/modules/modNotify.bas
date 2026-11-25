Attribute VB_Name = "modNotify"
Option Explicit
'==========================================================
' modNotify : 処理結果の通知
'   ・Windows イベントログ（アプリケーション）… 常に記録。既存の監視ツールで拾える
'   ・Webhook（Slack / Discord など）         … Config!NotifyOn の条件に合うときだけ送る
'
'   セキュリティ
'   ・Webhook の URL は「それ自体が送信の権限（秘密情報）」なので、ブックにもソースにも保存しない。
'     Config には環境変数の「名前」だけを書き、URL は環境変数から読む
'   ・https 以外の URL には送らない。ログに URL を出すときは途中を伏せる
'   ・通知に失敗してもバッチ本体の結果は変えない（ログに警告を残すだけ）
'==========================================================

Private Const EVENT_SUCCESS As Long = 0
Private Const EVENT_ERROR As Long = 1
Private Const EVENT_WARNING As Long = 2
Private Const EVENT_INFORMATION As Long = 4

Public Sub NotifyBatchResult(ByVal exitCode As Long, ByVal summary As String)
    Dim title As String
    Dim notifyOn As String

    title = "[" & AppTitle() & "] バッチ結果：" & StatusName(exitCode)

    '--- イベントログ（常に）
    Select Case exitCode
        Case EXIT_OK, EXIT_SKIPPED: WriteEventLog EVENT_INFORMATION, title & vbCrLf & summary
        Case EXIT_WARN:             WriteEventLog EVENT_WARNING, title & vbCrLf & summary
        Case Else:                  WriteEventLog EVENT_ERROR, title & vbCrLf & summary
    End Select

    '--- Webhook（条件付き）
    notifyOn = UCase$(CStr(GetConfig("NotifyOn", "ERROR")))
    Select Case notifyOn
        Case "ALWAYS"
            PostWebhook title & vbLf & summary
        Case "ERROR"
            If exitCode = EXIT_WARN Or exitCode = EXIT_ERROR Then PostWebhook title & vbLf & summary
        Case Else
            ' NONE：送らない
    End Select
End Sub

'----------------------------------------------------------
' 通知のテスト（管理者が Alt+F8 から実行）
'----------------------------------------------------------
Public Sub TestNotification()
    Const PROC As String = "modNotify.TestNotification"
    Dim envName As String
    Dim sent As Boolean
    On Error GoTo EH

    ResetRunId
    ReloadConfig
    RequirePermission actAdmin

    WriteEventLog EVENT_INFORMATION, "[" & AppTitle() & "] 通知のテストです（実行ID：" & RunId() & "）"
    sent = PostWebhook("[" & AppTitle() & "] 通知のテストです（実行ID：" & RunId() & "）")

    envName = CStr(GetConfig("NotifyWebhookEnv", "ORDERTOOL_WEBHOOK_URL"))
    MsgBox "イベントログ：記録しました（イベントビューアー → Windows ログ → アプリケーション、ソース「WSH」）" & vbCrLf & _
           "Webhook：" & IIf(sent, "送信しました", "送信していません（環境変数 " & envName & " が未設定、または送信に失敗。ログを確認してください）"), _
           vbInformation, AppTitle()
    Exit Sub
EH:
    HandleError PROC
End Sub

'==========================================================
' 内部処理
'==========================================================
Private Sub WriteEventLog(ByVal eventType As Long, ByVal message As String)
    On Error GoTo Fail
    CreateObject("WScript.Shell").LogEvent eventType, message
    Exit Sub
Fail:
    LogWarn "modNotify.WriteEventLog", "イベントログに書き込めませんでした: " & Err.Description
End Sub

' 戻り値：送信できたら True
Private Function PostWebhook(ByVal message As String) As Boolean
    Const PROC As String = "modNotify.PostWebhook"
    Dim envName As String
    Dim url As String
    Dim body As String
    Dim http As Object
    On Error GoTo Fail

    envName = CStr(GetConfig("NotifyWebhookEnv", "ORDERTOOL_WEBHOOK_URL"))
    url = Environ$(envName)
    If Len(url) = 0 Then Exit Function                 ' 未設定なら送らない（設定は任意）

    If LCase$(Left$(url, 8)) <> "https://" Then
        LogWarn PROC, "https 以外の URL には送信しません: " & MaskUrl(url)
        Exit Function
    End If

    Select Case UCase$(CStr(GetConfig("NotifyWebhookFormat", "SLACK")))
        Case "DISCORD": body = "{""content"": " & JsonString(Left$(message, 1900)) & "}"   ' Discord は2000文字まで
        Case Else:      body = "{""text"": " & JsonString(message) & "}"                     ' Slack など
    End Select

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 5000, 10000, 10000          ' 通知先が応答しなくてもバッチを長く止めない
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send body

    If http.status >= 200 And http.status < 300 Then
        LogInfo PROC, "通知を送信しました: " & MaskUrl(url)
        PostWebhook = True
    Else
        LogWarn PROC, "通知の送信に失敗しました: HTTP " & http.status & " / " & MaskUrl(url)
    End If
    Exit Function
Fail:
    LogWarn PROC, "通知の送信に失敗しました: " & Err.Description & " / " & MaskUrl(url)
End Function

' URL のうち、ホスト名より後ろ（トークンを含む部分）を伏せる
Public Function MaskUrl(ByVal url As String) As String
    Dim p As Long
    p = InStr(9, url, "/")
    If p = 0 Then
        MaskUrl = url
    Else
        MaskUrl = Left$(url, p) & "***"
    End If
End Function


