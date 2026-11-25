Attribute VB_Name = "modTestPhase5"
Option Explicit
'==========================================================
' modTestPhase5 : バッチ・通知の部品の自動テスト
'   Alt+F8 → RunPhase5Tests
'==========================================================

Private mPass As Long
Private mFail As Long

Public Sub RunPhase5Tests()
    Dim summary As String
    On Error GoTo EH

    mPass = 0
    mFail = 0
    ResetRunId
    ReloadConfig
    Debug.Print String(50, "-") & vbCrLf & "Phase 5 テスト開始 " & Now

    Test_Json
    Test_FileHelpers
    Test_Misc

    summary = "PASS: " & mPass & " / FAIL: " & mFail
    Debug.Print summary
    LogInfo "modTestPhase5", "テスト結果 " & summary
    MsgBox "Phase 5 テスト結果" & vbCrLf & summary & vbCrLf & "（詳細はイミディエイトウィンドウ）", _
           IIf(mFail = 0, vbInformation, vbCritical), AppTitle()
    Exit Sub
EH:
    HandleError "modTestPhase5.RunPhase5Tests"
End Sub

'----------------------------------------------------------
Private Sub Test_Json()
    AssertEqual "JSON: 通常文字", """abc""", JsonString("abc")
    AssertEqual "JSON: 日本語はそのまま", """受注""", JsonString("受注")
    AssertEqual "JSON: 引用符", """a\""b""", JsonString("a""b")
    AssertEqual "JSON: \ ", """C:\\temp""", JsonString("C:\temp")
    AssertEqual "JSON: 改行・タブ", """a\nb\tc""", JsonString("a" & vbLf & "b" & vbTab & "c")
    AssertEqual "JSON: 制御文字", """\u0001""", JsonString(Chr$(1))
End Sub

Private Sub Test_FileHelpers()
    Dim folder As String
    Dim fileNo As Integer

    folder = Environ$("TEMP") & "\ordertool_test_" & Format$(Now, "yyyymmddhhnnss")
    EnsureFolder folder

    AssertEqual "空き名: 存在しなければそのまま", folder & "\a.csv", UniquePath(folder, "a.csv")
    fileNo = FreeFile
    Open folder & "\a.csv" For Output As #fileNo
    Close #fileNo
    AssertEqual "空き名: 存在すれば _2 を付ける", folder & "\a_2.csv", UniquePath(folder, "a.csv")

    AssertTrue "書込中判定: 作成直後は未完了（60秒）", Not IsFileReady(folder & "\a.csv", 60)
    AssertTrue "書込中判定: 0秒なら完了扱い", IsFileReady(folder & "\a.csv", 0)

    Kill folder & "\a.csv"
    RmDir folder
End Sub

Private Sub Test_Misc()
    AssertEqual "終了コード: 0", "OK", StatusName(EXIT_OK)
    AssertEqual "終了コード: 1", "WARN", StatusName(EXIT_WARN)
    AssertEqual "終了コード: 2", "ERROR", StatusName(EXIT_ERROR)
    AssertEqual "終了コード: 3", "SKIPPED", StatusName(EXIT_SKIPPED)
    AssertEqual "URLの伏せ字", "https://hooks.example.com/***", MaskUrl("https://hooks.example.com/services/T000/B000/XXXX")

    SetBatchMode True
    AssertTrue "バッチモード: 有効", IsBatchMode()
    SetBatchMode False
    AssertTrue "バッチモード: 無効に戻る", Not IsBatchMode()
End Sub

'----------------------------------------------------------
Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If CStr(expected) = CStr(actual) Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & testName
    Else
        mFail = mFail + 1
        Debug.Print "* FAIL  " & testName & "  expected=[" & expected & "] actual=[" & actual & "]"
        LogWarn "modTestPhase5", "FAIL " & testName & " expected=" & expected & " actual=" & actual
    End If
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, True, condition
End Sub


