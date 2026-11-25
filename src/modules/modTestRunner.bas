Attribute VB_Name = "modTestRunner"
Option Explicit
'==========================================================
' modTestRunner : 全テストの一括実行
'   Phase 2〜5 の自動テストを順に実行し、結果を1つの画面とファイルにまとめる
'   ・各テストは、一括実行中（IsTestQuiet = True）は個別のダイアログを出さず、
'     RecordTestResult で件数だけを報告する
'   ・あるテストがエラーで中断しても、残りのテストは実行し、中断したことを結果に残す
'   ・外部アドイン（Rubberduck 等）に依存しないため、Excel だけで誰でも実行できる
'==========================================================

Private mQuiet As Boolean
Private mResults As Collection          ' 各要素：Array(テスト名, PASS 件数, FAIL 件数, 備考)

Public Sub RunAllTests()
    Const PROC As String = "modTestRunner.RunAllTests"
    Dim sw As clsStopwatch
    Dim r As Variant
    Dim totalPass As Long
    Dim totalFail As Long
    Dim md As String
    Dim mdPath As String
    Dim lines As String
    On Error GoTo EH

    ResetRunId
    ReloadConfig
    Set mResults = New Collection
    Set sw = New clsStopwatch
    mQuiet = True

    RunSuite "RunPhase2Tests", "Phase 2（CSV取込・検証）"
    RunSuite "RunPhase3Tests", "Phase 3（集計・帳票）"
    RunSuite "RunPhase4Tests", "Phase 4（権限・監査・暗号）"
    RunSuite "RunPhase5Tests", "Phase 5（バッチ・通知）"

    mQuiet = False

    md = "## テスト結果（" & Format$(Now, "yyyy/mm/dd hh:nn") & " / v" & APP_VERSION & "）" & vbCrLf & vbCrLf
    md = md & "| テスト | PASS | FAIL | 備考 |" & vbCrLf & "|---|---:|---:|---|" & vbCrLf
    For Each r In mResults
        md = md & "| " & r(0) & " | " & r(1) & " | " & r(2) & " | " & r(3) & " |" & vbCrLf
        lines = lines & IIf(r(2) = 0 And Len(r(3)) = 0, "OK  ", "NG  ") & r(0) & "：PASS " & r(1) & " / FAIL " & r(2) & _
                IIf(Len(r(3)) > 0, "（" & r(3) & "）", "") & vbCrLf
        totalPass = totalPass + r(1)
        totalFail = totalFail + r(2)
    Next r
    md = md & "| **合計** | **" & totalPass & "** | **" & totalFail & "** | " & Format$(sw.ElapsedSeconds, "0.0") & " 秒 |" & vbCrLf

    mdPath = ResolvePath("OutputDir", "output") & "\test"
    EnsureFolder mdPath
    mdPath = mdPath & "\test_result_" & RunId() & ".md"
    WriteTextFile mdPath, md, False
    Debug.Print md
    LogInfo PROC, "全テスト結果 PASS: " & totalPass & " / FAIL: " & totalFail

    MsgBox "全テストの結果" & vbCrLf & vbCrLf & lines & vbCrLf & _
           "合計：PASS " & totalPass & " / FAIL " & totalFail & vbCrLf & vbCrLf & _
           "結果（Markdown）：" & vbCrLf & mdPath, _
           IIf(totalFail = 0, vbInformation, vbCritical), AppTitle()
    Exit Sub
EH:
    mQuiet = False
    HandleError PROC
End Sub

' 一括実行中か（各テストが個別のダイアログを出すかどうかの判定に使う）
Public Function IsTestQuiet() As Boolean
    IsTestQuiet = mQuiet
End Function

' 各テストから結果を受け取る（単独で実行されたときは何もしない）
Public Sub RecordTestResult(ByVal suiteName As String, ByVal passCount As Long, ByVal failCount As Long)
    If Not mQuiet Then Exit Sub
    If mResults Is Nothing Then Exit Sub
    mResults.Add Array(suiteName, passCount, failCount, "")
End Sub

Private Sub RunSuite(ByVal procName As String, ByVal suiteName As String)
    Dim before As Long
    before = mResults.Count
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!" & procName
    On Error GoTo 0
    ' 結果が報告されなかった = 途中のエラーで中断した
    If mResults.Count = before Then mResults.Add Array(suiteName, 0, 1, "エラーで中断（ログを確認）")
End Sub

