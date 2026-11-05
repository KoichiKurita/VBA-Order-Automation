Attribute VB_Name = "modSmokeTest"
Option Explicit
'==========================================================
' modSmokeTest : Phase 1 基盤の動作確認
'==========================================================

' 正常系：設定読込→ログ→バックアップ→ログ掃除
Public Sub Phase1_SmokeTest()
    Const PROC As String = "modSmokeTest.Phase1_SmokeTest"
    Dim guard As clsAppState
    On Error GoTo EH

    Set guard = New clsAppState
    ResetRunId
    ReloadConfig

    LogInfo PROC, "スモークテスト開始"
    LogDebug PROC, "DEBUGログ（LogLevel=DEBUG のときだけ出力される）"
    LogWarn PROC, "改行" & vbLf & "とタブ" & vbTab & "を含む文字列 → 1行に無害化される"

    guard.ShowProgress "バックアップ作成中..."
    CreateBackup
    CleanupOldLogs

    LogInfo PROC, "スモークテスト正常終了"
    Set guard = Nothing

    MsgBox "Phase 1 スモークテスト完了" & vbCrLf & "ログ: " & CurrentLogFile(), vbInformation, AppTitle()
    Exit Sub

EH:
    Set guard = Nothing   ' 先に画面更新等を元に戻してからメッセージ表示
    HandleError PROC
End Sub

' 業務エラー：内容がそのまま表示される
Public Sub Phase1_AppErrorTest()
    Const PROC As String = "modSmokeTest.Phase1_AppErrorTest"
    On Error GoTo EH
    RaiseAppError ERR_VALIDATION, PROC, "入力データに不備があります（テスト）"
    Exit Sub
EH:
    HandleError PROC
End Sub

' システムエラー：汎用メッセージのみ表示、詳細はログ
Public Sub Phase1_UnexpectedErrorTest()
    Const PROC As String = "modSmokeTest.Phase1_UnexpectedErrorTest"
    Dim x As Long
    On Error GoTo EH
    x = 1 / 0
    Exit Sub
EH:
    HandleError PROC
End Sub

