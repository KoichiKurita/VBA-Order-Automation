Attribute VB_Name = "modReport"
Option Explicit
'==========================================================
' modReport : 受注集計レポート（集計・帳票シート・PDF / CSV 出力）
'   UI層   : CreateMonthlyReport
'   処理層 : BuildReport / ExportReportPdf / ExportSummaryCsv
'            （画面操作なし。Phase 5 の無人実行で再利用）
'   高速化 : Orders を配列で一括取得し、Dictionary で1パス集計
'            （セル単位の参照や SUMIF を行数分繰り返さない）
'==========================================================

Private Const REPORT_SHEET As String = "Report"

' 集計に使う列（Schemaの項目名）
Private Const COL_DATE As String = "受注日"
Private Const COL_CUST_CODE As String = "顧客コード"
Private Const COL_CUST_NAME As String = "顧客名"
Private Const COL_PROD_CODE As String = "商品コード"
Private Const COL_PROD_NAME As String = "商品名"
Private Const COL_QTY As String = "数量"
Private Const COL_PRICE As String = "単価"
Private Const COL_REP As String = "担当者"

' 集計結果配列の列
Private Const G_KEY As Long = 1
Private Const G_LABEL As Long = 2
Private Const G_COUNT As Long = 3
Private Const G_QTY As Long = 4
Private Const G_AMOUNT As Long = 5
Private Const G_SHARE As Long = 6
Private Const G_COLS As Long = 6

' 帳票の見た目
Private Const HEADER_COLOR As Long = 7949855      ' RGB(31, 78, 121)
Private Const REPORT_FONT As String = "Meiryo UI"

'----------------------------------------------------------
' UI：Alt+F8 やボタンから呼ぶ
'----------------------------------------------------------
Public Sub CreateMonthlyReport()
    Const PROC As String = "modReport.CreateMonthlyReport"
    Dim ym As String
    Dim fromDate As Date
    Dim toDate As Date
    Dim rep As clsReportData
    Dim pdfPath As String
    Dim csvPath As String
    On Error GoTo EH

    ResetRunId
    ReloadConfig

    ym = AskYearMonth()
    If Len(ym) = 0 Then
        LogInfo PROC, "年月の入力がキャンセルされました"
        Exit Sub
    End If
    fromDate = DateSerial(CLng(Left$(ym, 4)), CLng(Mid$(ym, 6, 2)), 1)
    toDate = DateSerial(Year(fromDate), Month(fromDate) + 1, 0)   ' 月末日

    Set rep = BuildReport(fromDate, toDate)
    pdfPath = ExportReportPdf(rep)
    csvPath = ExportSummaryCsv(rep)

    ThisWorkbook.Worksheets(REPORT_SHEET).Activate
    If MsgBox("レポートを作成しました。" & vbCrLf & vbCrLf & _
              "受注件数：" & Format$(rep.OrderCount, "#,##0") & " 件" & vbCrLf & _
              "受注金額：" & Format$(rep.TotalAmount, "#,##0") & " 円" & vbCrLf & vbCrLf & _
              "PDF：" & pdfPath & vbCrLf & _
              "CSV：" & csvPath & vbCrLf & vbCrLf & _
              "出力フォルダを開きますか？", vbInformation + vbYesNo, AppTitle()) = vbYes Then
        Shell "explorer.exe /select,""" & pdfPath & """", vbNormalFocus
    End If
    Exit Sub

EH:
    HandleError PROC
End Sub

'----------------------------------------------------------
' 処理層：集計して Report シートを作成する
'----------------------------------------------------------
Public Function BuildReport(ByVal fromDate As Date, ByVal toDate As Date) As clsReportData
    Const PROC As String = "modReport.BuildReport"
    Dim sw As clsStopwatch
    Dim guard As clsAppState
    Dim rep As clsReportData
    Dim lo As ListObject
    Dim data As Variant
    Dim iDate As Long, iCustCode As Long, iCustName As Long
    Dim iProdCode As Long, iProdName As Long
    Dim iQty As Long, iPrice As Long, iRep As Long
    Dim n As Long
    Dim i As Long
    Dim cnt As Long
    Dim rowsIdx() As Long
    Dim qtys() As Double
    Dim amounts() As Double
    Dim d As Variant

    Set sw = New clsStopwatch
    Set guard = New clsAppState
    Set rep = New clsReportData
    rep.fromDate = fromDate
    rep.toDate = toDate
    rep.ReportId = RunId()
    LogInfo PROC, "レポート作成開始: " & Format$(fromDate, "yyyy/mm/dd") & " - " & Format$(toDate, "yyyy/mm/dd")

    '--- 1. 一括読込
    guard.ShowProgress "受注データを読み込んでいます..."
    Set lo = GetOrdersTable()
    If lo.DataBodyRange Is Nothing Then RaiseAppError ERR_VALIDATION, PROC, "受注データがありません。先にCSVを取り込んでください。"
    data = lo.DataBodyRange.value

    iDate = ColIndex(lo, COL_DATE)
    iCustCode = ColIndex(lo, COL_CUST_CODE)
    iCustName = ColIndex(lo, COL_CUST_NAME)
    iProdCode = ColIndex(lo, COL_PROD_CODE)
    iProdName = ColIndex(lo, COL_PROD_NAME)
    iQty = ColIndex(lo, COL_QTY)
    iPrice = ColIndex(lo, COL_PRICE)
    iRep = ColIndex(lo, COL_REP)

    '--- 2. 期間で絞り込み、金額を1回だけ計算
    guard.ShowProgress "集計しています..."
    n = UBound(data, 1)
    ReDim rowsIdx(1 To n)
    ReDim qtys(1 To n)
    ReDim amounts(1 To n)
    For i = 1 To n
        d = data(i, iDate)
        If VarType(d) = vbDate Then                ' 空行（キー未入力）は日付が無いので自然に除外
            If d >= fromDate And d <= toDate Then
                cnt = cnt + 1
                rowsIdx(cnt) = i
                qtys(i) = CDbl(data(i, iQty))
                amounts(i) = qtys(i) * CDbl(data(i, iPrice))
                rep.TotalQty = rep.TotalQty + qtys(i)
                rep.TotalAmount = rep.TotalAmount + amounts(i)
            End If
        End If
    Next i

    If cnt = 0 Then
        RaiseAppError ERR_VALIDATION, PROC, "対象期間（" & Format$(fromDate, "yyyy/mm/dd") & " 〜 " & _
                      Format$(toDate, "yyyy/mm/dd") & "）の受注データがありません"
    End If
    rep.OrderCount = cnt

    '--- 3. 軸ごとに集計（金額の降順。日別は日付の昇順）
    rep.ByCustomer = GroupBy(data, rowsIdx, cnt, qtys, amounts, iCustCode, iCustName, rep.TotalAmount, True)
    rep.ByProduct = GroupBy(data, rowsIdx, cnt, qtys, amounts, iProdCode, iProdName, rep.TotalAmount, True)
    rep.ByRep = GroupBy(data, rowsIdx, cnt, qtys, amounts, iRep, iRep, rep.TotalAmount, True)
    rep.ByDate = GroupBy(data, rowsIdx, cnt, qtys, amounts, iDate, iDate, rep.TotalAmount, False)
    data = Empty

    '--- 4. 帳票シート作成
    guard.ShowProgress "帳票を作成しています..."
    RenderReport rep

    rep.ElapsedSeconds = sw.ElapsedSeconds
    LogInfo PROC, "レポート作成終了: " & cnt & " 件 / 金額 " & Format$(rep.TotalAmount, "#,##0") & _
                  " / " & Format$(rep.ElapsedSeconds, "0.000") & " 秒"
    Set BuildReport = rep
End Function

'----------------------------------------------------------
' 処理層：Report シートを PDF 出力
'----------------------------------------------------------
Public Function ExportReportPdf(ByVal rep As clsReportData) As String
    Const PROC As String = "modReport.ExportReportPdf"
    Dim folder As String
    Dim filePath As String

    folder = ResolvePath("OutputDir", "output")
    EnsureFolder folder
    filePath = folder & "\受注集計_" & Format$(rep.fromDate, "yyyymm") & "_" & rep.ReportId & ".pdf"

    ' IncludeDocProperties:=False … 作成者名・会社名などのメタデータをPDFに含めない（情報漏えい対策）
    ThisWorkbook.Worksheets(REPORT_SHEET).ExportAsFixedFormat _
        Type:=xlTypePDF, fileName:=filePath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

    LogInfo PROC, "PDF出力: " & filePath
    ExportReportPdf = filePath
End Function

'----------------------------------------------------------
' 処理層：集計結果を CSV 出力（UTF-8 BOM付き = Excelで文字化けしない）
'   値は ToCsvField を通し、CSVインジェクション対策とエスケープを行う
'----------------------------------------------------------
Public Function ExportSummaryCsv(ByVal rep As clsReportData) As String
    Const PROC As String = "modReport.ExportSummaryCsv"
    Dim lines As Collection
    Dim arr() As String
    Dim i As Long
    Dim folder As String
    Dim filePath As String

    Set lines = New Collection
    lines.Add "区分,コード,名称,件数,数量,金額,構成比"
    AddCsvLines lines, "顧客別", rep.ByCustomer, True
    AddCsvLines lines, "商品別", rep.ByProduct, True
    AddCsvLines lines, "担当者別", rep.ByRep, False
    AddCsvLines lines, "日別", rep.ByDate, False

    ReDim arr(1 To lines.Count)
    For i = 1 To lines.Count
        arr(i) = lines(i)
    Next i

    folder = ResolvePath("OutputDir", "output")
    EnsureFolder folder
    filePath = folder & "\受注集計_" & Format$(rep.fromDate, "yyyymm") & "_" & rep.ReportId & ".csv"
    WriteTextFile filePath, Join(arr, vbCrLf) & vbCrLf, True

    LogInfo PROC, "CSV出力: " & filePath
    ExportSummaryCsv = filePath
End Function

'----------------------------------------------------------
' 共通：Orders テーブルを取得（無ければ業務エラー）
'----------------------------------------------------------
Public Function GetOrdersTable() As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Orders")
    On Error GoTo 0
    If ws Is Nothing Then
        RaiseAppError ERR_VALIDATION, "modReport.GetOrdersTable", "受注データがありません。先にCSVを取り込んでください。"
    End If
    If ws.ListObjects.Count = 0 Then
        RaiseAppError ERR_VALIDATION, "modReport.GetOrdersTable", "Ordersシートにテーブルがありません。"
    End If
    Set GetOrdersTable = ws.ListObjects(1)
End Function

'==========================================================
' 内部処理：集計
'==========================================================
Private Function GroupBy(ByRef data As Variant, ByRef rowsIdx() As Long, ByVal cnt As Long, _
                         ByRef qtys() As Double, ByRef amounts() As Double, _
                         ByVal keyCol As Long, ByVal labelCol As Long, _
                         ByVal total As Double, ByVal sortByAmount As Boolean) As Variant
    Dim dict As Object
    Dim acc() As Variant
    Dim result() As Variant
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim pos As Long
    Dim k As String

    Set dict = CreateObject("Scripting.Dictionary")
    ReDim acc(1 To cnt, 1 To G_COLS)          ' 種類数は最大でも対象行数

    For i = 1 To cnt
        r = rowsIdx(i)
        If VarType(data(r, keyCol)) = vbDate Then
            k = Format$(data(r, keyCol), "yyyy/mm/dd")
        Else
            k = CStr(data(r, keyCol))
        End If
        If Len(k) = 0 Then k = "（未設定）"

        If dict.Exists(k) Then
            pos = dict(k)
        Else
            pos = dict.Count + 1
            dict.Add k, pos
            acc(pos, G_KEY) = k
            acc(pos, G_LABEL) = CStr(data(r, labelCol))
            acc(pos, G_COUNT) = 0
            acc(pos, G_QTY) = 0#
            acc(pos, G_AMOUNT) = 0#
        End If
        acc(pos, G_COUNT) = acc(pos, G_COUNT) + 1
        acc(pos, G_QTY) = acc(pos, G_QTY) + qtys(r)
        acc(pos, G_AMOUNT) = acc(pos, G_AMOUNT) + amounts(r)
    Next i

    ' 実際の種類数に詰め替え、構成比を計算
    ReDim result(1 To dict.Count, 1 To G_COLS)
    For i = 1 To dict.Count
        For c = 1 To G_COLS - 1
            result(i, c) = acc(i, c)
        Next c
        If total <> 0 Then result(i, G_SHARE) = result(i, G_AMOUNT) / total Else result(i, G_SHARE) = 0#
    Next i

    If sortByAmount Then
        SortRows result, G_AMOUNT, True
    Else
        SortRows result, G_KEY, False
    End If
    GroupBy = result
End Function

Private Sub SortRows(ByRef arr() As Variant, ByVal col As Long, ByVal desc As Boolean)
    If UBound(arr, 1) > 1 Then QuickSort arr, col, desc, 1, UBound(arr, 1)
End Sub

Private Sub QuickSort(ByRef arr() As Variant, ByVal col As Long, ByVal desc As Boolean, _
                      ByVal lo As Long, ByVal hi As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant

    i = lo
    j = hi
    pivot = arr((lo + hi) \ 2, col)
    Do While i <= j
        If desc Then
            Do While arr(i, col) > pivot
                i = i + 1
            Loop
            Do While arr(j, col) < pivot
                j = j - 1
            Loop
        Else
            Do While arr(i, col) < pivot
                i = i + 1
            Loop
            Do While arr(j, col) > pivot
                j = j - 1
            Loop
        End If
        If i <= j Then
            SwapRows arr, i, j
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then QuickSort arr, col, desc, lo, j
    If i < hi Then QuickSort arr, col, desc, i, hi
End Sub

Private Sub SwapRows(ByRef arr() As Variant, ByVal a As Long, ByVal b As Long)
    Dim c As Long
    Dim tmp As Variant
    For c = LBound(arr, 2) To UBound(arr, 2)
        tmp = arr(a, c)
        arr(a, c) = arr(b, c)
        arr(b, c) = tmp
    Next c
End Sub

Private Function ColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    ColIndex = lo.ListColumns(colName).Index
    On Error GoTo 0
    If ColIndex = 0 Then RaiseAppError ERR_SCHEMA, "modReport.ColIndex", "Ordersに列がありません：" & colName
End Function

'==========================================================
' 内部処理：帳票
'==========================================================
Private Sub RenderReport(ByVal rep As clsReportData)
    Dim ws As Worksheet
    Dim r As Long
    Dim topN As Long
    Dim lastRow As Long

    Set ws = GetReportSheet()
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop

    With ws.Cells.Font
        .Name = REPORT_FONT
        .size = 10
    End With
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 9

    '--- 見出し
    With ws.Range("A1")
        .value = "受注集計レポート"
        .Font.size = 16
        .Font.Bold = True
    End With
    ws.Range("A2").value = "対象期間：" & Format$(rep.fromDate, "yyyy/mm/dd") & " 〜 " & Format$(rep.toDate, "yyyy/mm/dd")
    ws.Range("A3").value = "出力日時：" & Format$(Now, "yyyy/mm/dd hh:nn") & "　出力者：" & _
                           Environ$("USERNAME") & "　レポートID：" & rep.ReportId
    ws.Range("A3").Font.Color = RGB(89, 89, 89)

    '--- KPI
    r = 5
    WriteKpi ws, r, "受注件数", rep.OrderCount, "#,##0"" 件"""
    WriteKpi ws, r + 1, "数量合計", rep.TotalQty, "#,##0"
    WriteKpi ws, r + 2, "受注金額合計", rep.TotalAmount, "#,##0"" 円"""
    WriteKpi ws, r + 3, "平均受注金額", rep.TotalAmount / rep.OrderCount, "#,##0"" 円"""

    '--- 軸別
    topN = CLng(GetConfig("ReportTopN", 20))
    r = r + 5
    r = WriteSection(ws, r, "顧客別", rep.ByCustomer, "顧客コード", "顧客名", topN)
    r = WriteSection(ws, r, "商品別", rep.ByProduct, "商品コード", "商品名", topN)
    r = WriteSection(ws, r, "担当者別", rep.ByRep, "担当者", "", topN)
    lastRow = WriteDaily(ws, r, rep.ByDate)

    SetupPrint ws, rep, lastRow
End Sub

Private Function GetReportSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = REPORT_SHEET
    End If
    Set GetReportSheet = ws
End Function

Private Sub WriteKpi(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    ByVal value As Double, ByVal numFmt As String)
    ws.Cells(r, 2).value = label
    ws.Cells(r, 2).Font.Bold = True
    With ws.Cells(r, 3)
        .value = value
        .NumberFormatLocal = numFmt
        .HorizontalAlignment = xlLeft
        .Font.size = 12
    End With
End Sub

' 戻り値：次のセクションの開始行
Private Function WriteSection(ByVal ws As Worksheet, ByVal startRow As Long, ByVal title As String, _
                              ByVal arr As Variant, ByVal codeHeader As String, ByVal nameHeader As String, _
                              ByVal topN As Long) As Long
    Dim n As Long
    Dim shown As Long
    Dim i As Long
    Dim j As Long
    Dim out() As Variant
    Dim dataRow As Long
    Dim totalRow As Long
    Dim hasOthers As Boolean

    n = UBound(arr, 1)
    hasOthers = (topN > 0 And n > topN)
    If hasOthers Then shown = topN + 1 Else shown = n

    With ws.Cells(startRow, 1)
        .value = title
        .Font.size = 12
        .Font.Bold = True
    End With
    ws.Cells(startRow + 1, 1).Resize(1, 7).value = Array("順位", codeHeader, nameHeader, "件数", "数量", "金額", "構成比")
    StyleHeader ws.Cells(startRow + 1, 1).Resize(1, 7)

    ReDim out(1 To shown + 1, 1 To 7)
    For i = 1 To shown
        If hasOthers And i = shown Then
            ' 上位N件以外は「その他」にまとめる（取引先が多くても帳票が肥大化しない）
            out(i, 3) = "その他（" & (n - topN) & " 件）"
            out(i, 4) = 0: out(i, 5) = 0#: out(i, 6) = 0#: out(i, 7) = 0#
            For j = topN + 1 To n
                out(i, 4) = out(i, 4) + arr(j, G_COUNT)
                out(i, 5) = out(i, 5) + arr(j, G_QTY)
                out(i, 6) = out(i, 6) + arr(j, G_AMOUNT)
                out(i, 7) = out(i, 7) + arr(j, G_SHARE)
            Next j
        Else
            out(i, 1) = i
            out(i, 2) = ToSafeCellText(CStr(arr(i, G_KEY)))
            If Len(nameHeader) > 0 Then out(i, 3) = ToSafeCellText(CStr(arr(i, G_LABEL)))
            out(i, 4) = arr(i, G_COUNT)
            out(i, 5) = arr(i, G_QTY)
            out(i, 6) = arr(i, G_AMOUNT)
            out(i, 7) = arr(i, G_SHARE)
        End If
    Next i

    ' 合計行
    totalRow = shown + 1
    out(totalRow, 3) = "合計"
    out(totalRow, 4) = 0: out(totalRow, 5) = 0#: out(totalRow, 6) = 0#: out(totalRow, 7) = 0#
    For i = 1 To n
        out(totalRow, 4) = out(totalRow, 4) + arr(i, G_COUNT)
        out(totalRow, 5) = out(totalRow, 5) + arr(i, G_QTY)
        out(totalRow, 6) = out(totalRow, 6) + arr(i, G_AMOUNT)
        out(totalRow, 7) = out(totalRow, 7) + arr(i, G_SHARE)
    Next i

    dataRow = startRow + 2
    ws.Cells(dataRow, 1).Resize(shown + 1, 7).value = out
    ws.Cells(dataRow, 4).Resize(shown + 1, 3).NumberFormat = "#,##0"
    ws.Cells(dataRow, 7).Resize(shown + 1, 1).NumberFormat = "0.0%"
    ws.Cells(dataRow, 1).Resize(shown + 1, 1).HorizontalAlignment = xlCenter
    StyleBody ws.Cells(dataRow, 1).Resize(shown, 7)
    StyleTotal ws.Cells(dataRow + shown, 1).Resize(1, 7)

    WriteSection = dataRow + shown + 2
End Function

' 日別推移の表とグラフ。戻り値：印刷範囲の最終行
Private Function WriteDaily(ByVal ws As Worksheet, ByVal startRow As Long, ByVal arr As Variant) As Long
    Dim n As Long
    Dim i As Long
    Dim out() As Variant
    Dim dataRow As Long
    Dim chartTop As Double
    Dim co As ChartObject

    n = UBound(arr, 1)
    With ws.Cells(startRow, 1)
        .value = "日別推移"
        .Font.size = 12
        .Font.Bold = True
    End With
    ws.Cells(startRow + 1, 2).Resize(1, 5).value = Array("日付", "", "件数", "数量", "金額")
    StyleHeader ws.Cells(startRow + 1, 2).Resize(1, 5)

    ReDim out(1 To n, 1 To 5)
    For i = 1 To n
        out(i, 1) = CDate(arr(i, G_KEY))
        out(i, 3) = arr(i, G_COUNT)
        out(i, 4) = arr(i, G_QTY)
        out(i, 5) = arr(i, G_AMOUNT)
    Next i

    dataRow = startRow + 2
    ws.Cells(dataRow, 2).Resize(n, 5).value = out
    ws.Cells(dataRow, 2).Resize(n, 1).NumberFormat = "yyyy/mm/dd"
    ws.Cells(dataRow, 2).Resize(n, 1).HorizontalAlignment = xlCenter
    ws.Cells(dataRow, 4).Resize(n, 3).NumberFormat = "#,##0"
    StyleBody ws.Cells(dataRow, 2).Resize(n, 5)

    ' グラフ（日別受注金額）
    chartTop = ws.Cells(dataRow + n + 1, 1).Top
    Set co = ws.ChartObjects.Add(ws.Cells(1, 1).Left, chartTop, ws.Range("A1:G1").Width, 220)
    With co.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Name = "受注金額"
            .Values = ws.Cells(dataRow, 6).Resize(n, 1)
            .XValues = ws.Cells(dataRow, 2).Resize(n, 1)
        End With
        .HasTitle = True
        .ChartTitle.text = "日別受注金額"
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    End With

    WriteDaily = co.BottomRightCell.Row
End Function

Private Sub StyleHeader(ByVal rng As Range)
    With rng
        .Interior.Color = HEADER_COLOR
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub StyleBody(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(191, 191, 191)
    End With
End Sub

Private Sub StyleTotal(ByVal rng As Range)
    rng.Font.Bold = True
    rng.Interior.Color = RGB(242, 242, 242)
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .Color = RGB(89, 89, 89)
    End With
End Sub

' 印刷設定（A4縦・横1ページに収める・フッターにページ番号とレポートID）
Private Sub SetupPrint(ByVal ws As Worksheet, ByVal rep As clsReportData, ByVal lastRow As Long)
    On Error GoTo EH
    Application.PrintCommunication = False      ' プリンタとの通信をまとめて行い高速化
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 7)).Address
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .LeftFooter = "レポートID：" & rep.ReportId
        .CenterFooter = "&P / &N"
        .RightFooter = "出力日時：&D &T"
    End With
    Application.PrintCommunication = True
    Exit Sub
EH:
    Application.PrintCommunication = True
    LogWarn "modReport.SetupPrint", "印刷設定に失敗しました（プリンタ未設定の可能性）: " & Err.Description
End Sub

'==========================================================
' 内部処理：入力・CSV
'==========================================================
Private Function AskYearMonth() As String
    Dim def As String
    Dim ans As Variant
    Dim s As String
    Dim m As Long

    def = Format$(LatestOrderDate(), "yyyy/mm")
    Do
        ans = Application.InputBox("集計する年月を入力してください（yyyy/mm）", AppTitle(), def, Type:=2)
        If VarType(ans) = vbBoolean Then Exit Function          ' キャンセル
        s = Replace(StrConv(Trim$(CStr(ans)), vbNarrow), "-", "/")
        If s Like "####/##" Then
            m = CLng(Mid$(s, 6, 2))
            If m >= 1 And m <= 12 Then
                AskYearMonth = s
                Exit Function
            End If
        End If
        MsgBox "yyyy/mm の形式で入力してください（例：" & def & "）", vbExclamation, AppTitle()
    Loop
End Function

' 既定値として最新の受注月を出す（毎回入力する手間を減らす）
Private Function LatestOrderDate() As Date
    Dim lo As ListObject
    Dim v As Double
    On Error GoTo Fallback
    Set lo = GetOrdersTable()
    v = Application.WorksheetFunction.Max(lo.ListColumns(COL_DATE).DataBodyRange)
    If v > 0 Then LatestOrderDate = CDate(v) Else LatestOrderDate = Date
    Exit Function
Fallback:
    LatestOrderDate = Date
End Function

Private Sub AddCsvLines(ByVal lines As Collection, ByVal category As String, _
                        ByVal arr As Variant, ByVal withName As Boolean)
    Dim i As Long
    Dim nameValue As Variant
    For i = 1 To UBound(arr, 1)
        If withName Then nameValue = arr(i, G_LABEL) Else nameValue = ""
        lines.Add ToCsvField(category) & "," & _
                  ToCsvField(arr(i, G_KEY)) & "," & _
                  ToCsvField(nameValue) & "," & _
                  ToCsvField(arr(i, G_COUNT)) & "," & _
                  ToCsvField(arr(i, G_QTY)) & "," & _
                  ToCsvField(arr(i, G_AMOUNT)) & "," & _
                  ToCsvField(Round(arr(i, G_SHARE), 4))
    Next i
End Sub


