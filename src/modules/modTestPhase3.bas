Attribute VB_Name = "modTestPhase3"
Option Explicit
'==========================================================
' modTestPhase3 : 集計・CSV出力の自動テスト
'   Alt+F8 → RunPhase3Tests
'   集計結果を、Excelのワークシート関数（SUMPRODUCT / COUNTIFS）で
'   別途計算した値と突き合わせる（独立した検算）
'   ※ Orders にデータが入っている状態で実行すること
'==========================================================

Private mPass As Long
Private mFail As Long

Public Sub RunPhase3Tests()
    Dim summary As String
    On Error GoTo EH

    mPass = 0
    mFail = 0
    ResetRunId
    ReloadConfig
    Debug.Print String(50, "-") & vbCrLf & "Phase 3 テスト開始 " & Now

    Test_ToCsvField
    Test_Stopwatch
    Test_BuildReport

    summary = "PASS: " & mPass & " / FAIL: " & mFail
    Debug.Print summary
    LogInfo "modTestPhase3", "テスト結果 " & summary
    MsgBox "Phase 3 テスト結果" & vbCrLf & summary & vbCrLf & "（詳細はイミディエイトウィンドウ）", _
           IIf(mFail = 0, vbInformation, vbCritical), AppTitle()
    Exit Sub
EH:
    HandleError "modTestPhase3.RunPhase3Tests"
End Sub

'----------------------------------------------------------
Private Sub Test_ToCsvField()
    AssertEqual "CSV出力: 通常文字", "abc", ToCsvField("abc")
    AssertEqual "CSV出力: カンマは囲む", """a,b""", ToCsvField("a,b")
    AssertEqual "CSV出力: 引用符はエスケープ", """x""""y""", ToCsvField("x""y")
    AssertEqual "CSV出力: 改行は囲む", """a" & vbLf & "b""", ToCsvField("a" & vbLf & "b")
    AssertEqual "CSV出力: 数式様の文字列に '", "'=1+1", ToCsvField("=1+1")
    AssertEqual "CSV出力: 数式様＋カンマ", """'=SUM(A1,B1)""", ToCsvField("=SUM(A1,B1)")
    AssertEqual "CSV出力: 負の数値には付けない", "-5", ToCsvField(-5)
    AssertEqual "CSV出力: 日付", "2026/11/07", ToCsvField(DateSerial(2026, 11, 7))
    AssertEqual "CSV出力: 空", "", ToCsvField(Empty)
End Sub

Private Sub Test_Stopwatch()
    Dim sw As clsStopwatch
    Dim t As Double
    Dim t0 As Single

    ' Application.Wait は秒単位で待ち時間がばらつくため、Timer で1秒待つ
    Set sw = New clsStopwatch
    t0 = Timer
    Do While Timer - t0 < 1
        DoEvents
    Loop
    t = sw.ElapsedSeconds
    AssertTrue "ストップウォッチ: 約1秒を計測（" & Format$(t, "0.000") & " 秒）", (t > 0.9 And t < 1.5)
End Sub

Private Sub Test_BuildReport()
    Dim lo As ListObject
    Dim latest As Date
    Dim fromDate As Date
    Dim toDate As Date
    Dim rep As clsReportData
    Dim expCount As Double
    Dim expQty As Double
    Dim expAmount As Double
    Dim cond As String

    Set lo = GetOrdersTable()
    latest = CDate(Application.WorksheetFunction.Max(lo.ListColumns("受注日").DataBodyRange))
    fromDate = DateSerial(Year(latest), Month(latest), 1)
    toDate = DateSerial(Year(latest), Month(latest) + 1, 0)

    ' 期待値：ワークシート関数で独立に計算
    cond = "(tblOrders[受注日]>=" & CLng(fromDate) & ")*(tblOrders[受注日]<=" & CLng(toDate) & ")"
    expCount = Application.Evaluate("SUMPRODUCT(" & cond & ")")
    expQty = Application.Evaluate("SUMPRODUCT(" & cond & "*tblOrders[数量])")
    expAmount = Application.Evaluate("SUMPRODUCT(" & cond & "*tblOrders[数量]*tblOrders[単価])")

    Set rep = BuildReport(fromDate, toDate)

    AssertEqual "集計: 件数（" & Format$(fromDate, "yyyy/mm") & "）", expCount, rep.OrderCount
    AssertEqual "集計: 数量合計", expQty, rep.TotalQty
    AssertEqual "集計: 金額合計", expAmount, rep.TotalAmount

    AssertEqual "集計: 顧客別の合計 = 総額", expAmount, SumCol(rep.ByCustomer, 5)
    AssertEqual "集計: 商品別の合計 = 総額", expAmount, SumCol(rep.ByProduct, 5)
    AssertEqual "集計: 担当者別の合計 = 総額", expAmount, SumCol(rep.ByRep, 5)
    AssertEqual "集計: 日別の合計 = 総額", expAmount, SumCol(rep.ByDate, 5)
    AssertEqual "集計: 顧客別の件数合計", expCount, SumCol(rep.ByCustomer, 3)
    AssertEqual "集計: 構成比の合計 = 100%", "1.000000", Format$(SumCol(rep.ByCustomer, 6), "0.000000")
    AssertTrue "集計: 顧客別は金額の降順", IsSorted(rep.ByCustomer, 5, True)
    AssertTrue "集計: 日別は日付の昇順", IsSorted(rep.ByDate, 1, False)
End Sub

'----------------------------------------------------------
Private Function SumCol(ByVal arr As Variant, ByVal col As Long) As Double
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        SumCol = SumCol + arr(i, col)
    Next i
End Function

Private Function IsSorted(ByVal arr As Variant, ByVal col As Long, ByVal desc As Boolean) As Boolean
    Dim i As Long
    For i = 2 To UBound(arr, 1)
        If desc Then
            If arr(i - 1, col) < arr(i, col) Then Exit Function
        Else
            If arr(i - 1, col) > arr(i, col) Then Exit Function
        End If
    Next i
    IsSorted = True
End Function

Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If CStr(expected) = CStr(actual) Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & testName
    Else
        mFail = mFail + 1
        Debug.Print "* FAIL  " & testName & "  expected=[" & expected & "] actual=[" & actual & "]"
        LogWarn "modTestPhase3", "FAIL " & testName & " expected=" & expected & " actual=" & actual
    End If
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, True, condition
End Sub


