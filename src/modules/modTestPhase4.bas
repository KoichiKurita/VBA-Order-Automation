Attribute VB_Name = "modTestPhase4"
Option Explicit
'==========================================================
' modTestPhase4 : セキュリティ機能の自動テスト
'   Alt+F8 → RunPhase4Tests（SetupSecurity の実行後に行うこと）
'   ・SHA-256 は公開されている既知の値（テストベクタ）と照合する
'   ・監査ログは一時ファイルで「改ざん → 検知」まで確認する
'==========================================================

Private mPass As Long
Private mFail As Long

Public Sub RunPhase4Tests()
    Dim summary As String
    On Error GoTo EH

    mPass = 0
    mFail = 0
    ResetRunId
    ReloadConfig
    Debug.Print String(50, "-") & vbCrLf & "Phase 4 テスト開始 " & Now

    Test_Sha256
    Test_RandomToken
    Test_Roles
    Test_AuditChain
    Test_CurrentUser

    summary = "PASS: " & mPass & " / FAIL: " & mFail
    Debug.Print summary
    LogInfo "modTestPhase4", "テスト結果 " & summary
    RecordTestResult "Phase 4（権限・監査・暗号）", mPass, mFail
    If Not IsTestQuiet() Then
        MsgBox "Phase 4 テスト結果" & vbCrLf & summary & vbCrLf & "（詳細はイミディエイトウィンドウ）", _
               IIf(mFail = 0, vbInformation, vbCritical), AppTitle()
    End If
    Exit Sub
EH:
    HandleError "modTestPhase4.RunPhase4Tests"
End Sub

'----------------------------------------------------------
Private Sub Test_Sha256()
    Dim b() As Byte

    ' FIPS 180-2 のテストベクタ（ASCII の "abc"）
    b = StrConv("abc", vbFromUnicode)
    AssertEqual "SHA-256: ""abc""", _
        "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", Sha256HexBytes(b)

    ' 空のデータ
    b = ""
    AssertEqual "SHA-256: 空", _
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", Sha256HexBytes(b)

    AssertEqual "SHA-256: 同じ入力は同じ値", Sha256Hex("受注"), Sha256Hex("受注")
    AssertTrue "SHA-256: 1文字違えば別の値", Sha256Hex("受注1") <> Sha256Hex("受注2")
End Sub

Private Sub Test_RandomToken()
    Dim t1 As String
    Dim t2 As String
    t1 = RandomToken(24)
    t2 = RandomToken(24)
    AssertEqual "乱数トークン: 長さ", 24, Len(t1)
    AssertTrue "乱数トークン: 英数字のみ", Not (t1 Like "*[!A-Za-z0-9]*")
    AssertTrue "乱数トークン: 毎回異なる", t1 <> t2
End Sub

Private Sub Test_Roles()
    AssertEqual "ロール名: ADMIN", roleAdmin, RoleFromName("admin")
    AssertEqual "ロール名: 不明な値は権限なし", roleNone, RoleFromName("SUPERUSER")

    AssertTrue "閲覧者: 帳票は可", HasRoleFor(roleViewer, actReport)
    AssertTrue "閲覧者: 取込は不可", Not HasRoleFor(roleViewer, actImport)
    AssertTrue "担当者: 取込は可", HasRoleFor(roleOperator, actImport)
    AssertTrue "担当者: 取消は不可", Not HasRoleFor(roleOperator, actUndoImport)
    AssertTrue "管理者: 管理機能は可", HasRoleFor(roleAdmin, actAdmin)
    AssertTrue "権限なし: 帳票も不可", Not HasRoleFor(roleNone, actReport)
End Sub

Private Sub Test_AuditChain()
    Dim path As String
    Dim lines() As String
    Dim n As Long
    Dim i As Long
    Dim cnt As Long

    path = Environ$("TEMP") & "\audit_test_" & Format$(Now, "yyyymmddhhnnss") & ".log"

    AuditLogTo path, "TEST", "OK", "1行目"
    AuditLogTo path, "TEST", "OK", "2行目"
    AuditLogTo path, "TEST", "OK", "3行目（タブ" & vbTab & "と改行" & vbLf & "を含む）"
    AssertEqual "監査ログ: 改ざんなし", 0, VerifyAuditFile(path, cnt)
    AssertEqual "監査ログ: 行数", 3, cnt

    ' 2行目の内容を書き換える
    lines = ReadLines(path, n)
    lines(2) = Replace(lines(2), "2行目", "2行目（改ざん）")
    WriteLines path, lines, n
    AssertEqual "監査ログ: 書換えを検知（2行目）", 2, VerifyAuditFile(path, cnt)

    ' 元に戻してから、2行目を削除する
    lines(2) = Replace(lines(2), "2行目（改ざん）", "2行目")
    WriteLines path, lines, n
    AssertEqual "監査ログ: 元に戻すと一致", 0, VerifyAuditFile(path, cnt)
    For i = 2 To n - 1
        lines(i) = lines(i + 1)
    Next i
    WriteLines path, lines, n - 1
    AssertEqual "監査ログ: 行の削除を検知（3行目→2行目）", 2, VerifyAuditFile(path, cnt)

    Kill path
End Sub

Private Sub Test_CurrentUser()
    AssertTrue "ユーザー名: 取得できる（" & CurrentUserName() & "）", Len(CurrentUserName()) > 0
    AssertTrue "初期設定: 済み", IsSecuritySetUp()
    AssertEqual "現在のユーザー: 管理者", roleAdmin, CurrentRole()
    AssertEqual "保護パスワード: 24文字", 24, Len(ProtectionPassword())
    AssertTrue "有効な管理者がいる", HasEnabledAdmin()
End Sub

'----------------------------------------------------------
Private Function ReadLines(ByVal path As String, ByRef n As Long) As String()
    Dim arr() As String
    Dim fileNo As Integer
    Dim s As String
    ReDim arr(1 To 100)
    n = 0
    fileNo = FreeFile
    Open path For Input As #fileNo
    Do While Not EOF(fileNo)
        Line Input #fileNo, s
        n = n + 1
        arr(n) = s
    Loop
    Close #fileNo
    ReadLines = arr
End Function

Private Sub WriteLines(ByVal path As String, ByRef arr() As String, ByVal n As Long)
    Dim fileNo As Integer
    Dim i As Long
    fileNo = FreeFile
    Open path For Output As #fileNo
    For i = 1 To n
        Print #fileNo, arr(i)
    Next i
    Close #fileNo
End Sub

Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If CStr(expected) = CStr(actual) Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & testName
    Else
        mFail = mFail + 1
        Debug.Print "* FAIL  " & testName & "  expected=[" & expected & "] actual=[" & actual & "]"
        LogWarn "modTestPhase4", "FAIL " & testName & " expected=" & expected & " actual=" & actual
    End If
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, True, condition
End Sub

