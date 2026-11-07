Attribute VB_Name = "modTestPhase2"
Option Explicit
'==========================================================
' modTestPhase2 : パーサー・検証・セキュリティ関数の自動テスト
'   Alt+F8 → RunPhase2Tests
'   結果はイミディエイトウィンドウ（Ctrl+G）に出力され、失敗はログにも記録
'==========================================================

Private mPass As Long
Private mFail As Long

Public Sub RunPhase2Tests()
    Dim summary As String

    mPass = 0
    mFail = 0
    ResetRunId
    ReloadConfig
    Debug.Print String(50, "-") & vbCrLf & "Phase 2 テスト開始 " & Now

    Test_ParseCsv
    Test_DetectEncoding
    Test_ValidateValue
    Test_Security

    summary = "PASS: " & mPass & " / FAIL: " & mFail
    Debug.Print summary
    LogInfo "modTestPhase2", "テスト結果 " & summary
    MsgBox "Phase 2 テスト結果" & vbCrLf & summary & vbCrLf & "（詳細はイミディエイトウィンドウ）", _
           IIf(mFail = 0, vbInformation, vbCritical), AppTitle()
End Sub

'----------------------------------------------------------
Private Sub Test_ParseCsv()
    Dim d As Variant
    Dim fc() As Long
    Dim errNo As Long

    d = ParseCsv("a,b,c" & vbCrLf & "1,2,3" & vbCrLf, fc)
    AssertEqual "CSV: 行数", 2, UBound(d, 1)
    AssertEqual "CSV: 列数", 3, fc(2)
    AssertEqual "CSV: 値", "3", d(2, 3)

    d = ParseCsv("""a,b"",""x""""y""" & vbLf & """複数" & vbCrLf & "行"",z", fc)
    AssertEqual "CSV: 引用符内のカンマ", "a,b", d(1, 1)
    AssertEqual "CSV: エスケープされた引用符", "x""y", d(1, 2)
    AssertEqual "CSV: 引用符内の改行", "複数" & vbCrLf & "行", d(2, 1)
    AssertEqual "CSV: LF改行の行数", 2, UBound(d, 1)

    d = ParseCsv("a,b" & vbCrLf & vbCrLf & "1,", fc)
    AssertEqual "CSV: 空行を読み飛ばす", 2, UBound(d, 1)
    AssertEqual "CSV: 末尾の空フィールド", 2, fc(2)

    On Error Resume Next
    d = ParseCsv("a,""b", fc)
    errNo = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "CSV: 閉じていない引用符はエラー", ERR_CSV_FORMAT, errNo
End Sub

Private Sub Test_DetectEncoding()
    Dim b() As Byte

    b = StrConv("受注", vbFromUnicode)          ' 日本語WindowsではShift_JISのバイト列
    AssertEqual "文字コード: Shift_JIS", "shift_jis", DetectEncoding(b)

    ReDim b(0 To 2)
    b(0) = &HE3: b(1) = &H81: b(2) = &H82       ' 「あ」のUTF-8
    AssertEqual "文字コード: UTF-8（BOMなし）", "utf-8", DetectEncoding(b)

    ReDim b(0 To 3)
    b(0) = &HEF: b(1) = &HBB: b(2) = &HBF: b(3) = &H41
    AssertEqual "文字コード: UTF-8（BOM付き）", "utf-8", DetectEncoding(b)
End Sub

Private Sub Test_ValidateValue()
    Dim def As ColumnDef
    Dim v As Variant

    ' 整数
    def.Name = "数量": def.Required = True: def.dataType = "INTEGER"
    def.HasMin = True: def.MinValue = 1: def.HasMax = True: def.MaxValue = 9999
    AssertEqual "整数: 正常", "", ValidateValue(def, "10", v)
    AssertEqual "整数: 変換値", 10, v
    AssertEqual "整数: 全角数字を許容", "", ValidateValue(def, "１０", v)
    AssertEqual "整数: 文字はエラー", "整数で入力してください", ValidateValue(def, "abc", v)
    AssertNotEmpty "整数: 範囲外（下限）", ValidateValue(def, "0", v)
    AssertNotEmpty "整数: 範囲外（上限）", ValidateValue(def, "10000", v)
    AssertEqual "必須: 未入力", "必須項目が未入力です", ValidateValue(def, "", v)

    ' 小数
    def.dataType = "DECIMAL": def.HasMin = False: def.HasMax = False
    AssertEqual "小数: 正常", "", ValidateValue(def, "1234.5", v)
    AssertNotEmpty "小数: 桁区切りカンマはエラー", ValidateValue(def, "1,234", v)

    ' 日付
    def.dataType = "DATE"
    AssertEqual "日付: スラッシュ", "", ValidateValue(def, "2026/11/07", v)
    AssertEqual "日付: 変換値", DateSerial(2026, 11, 7), v
    AssertEqual "日付: ハイフン", "", ValidateValue(def, "2026-11-07", v)
    AssertNotEmpty "日付: 存在しない日（2/30）", ValidateValue(def, "2026/02/30", v)
    AssertNotEmpty "日付: 年なし", ValidateValue(def, "11/7", v)

    ' 文字列＋パターン
    def.dataType = "TEXT": def.MaxLength = 6: def.pattern = "C#####": def.Hint = "C00001"
    AssertEqual "文字列: 正常", "", ValidateValue(def, "C00001", v)
    AssertNotEmpty "文字列: パターン不一致", ValidateValue(def, "X00001", v)
    AssertNotEmpty "文字列: 桁数超過", ValidateValue(def, "C000012", v)
    def.pattern = "": def.MaxLength = 0
    AssertNotEmpty "文字列: 制御文字（タブ）", ValidateValue(def, "a" & vbTab & "b", v)

    ' 任意項目
    def.Required = False
    AssertEqual "任意: 未入力はOK", "", ValidateValue(def, "", v)
End Sub

Private Sub Test_Security()
    AssertEqual "数式判定: =", True, IsFormulaLike("=1+1")
    AssertEqual "数式判定: @", True, IsFormulaLike("@SUM(A1)")
    AssertEqual "数式判定: 全角＝", True, IsFormulaLike("＝1+1")
    AssertEqual "数式判定: 通常文字", False, IsFormulaLike("サンプル")
    AssertEqual "数式判定: 空文字", False, IsFormulaLike("")
    AssertEqual "安全化: 先頭に '", "'=1+1", ToSafeCellText("=1+1")
    AssertEqual "安全化: 空文字はそのまま", "", ToSafeCellText("")
End Sub

'----------------------------------------------------------
Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If CStr(expected) = CStr(actual) Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & testName
    Else
        mFail = mFail + 1
        Debug.Print "* FAIL  " & testName & "  expected=[" & expected & "] actual=[" & actual & "]"
        LogWarn "modTestPhase2", "FAIL " & testName & " expected=" & expected & " actual=" & actual
    End If
End Sub

Private Sub AssertNotEmpty(ByVal testName As String, ByVal actual As String)
    If Len(actual) > 0 Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & testName & "  -> " & actual
    Else
        mFail = mFail + 1
        Debug.Print "* FAIL  " & testName & "  （エラーになるべき値が通過）"
        LogWarn "modTestPhase2", "FAIL " & testName
    End If
End Sub


