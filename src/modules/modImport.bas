Attribute VB_Name = "modImport"
Option Explicit
'==========================================================
' modImport : 受注CSVの取込
'   UI層   : ImportOrdersCsv       … ファイル選択・結果表示（ユーザー操作用）
'   処理層 : ImportOrdersFromFile  … 画面に依存しない（Phase 5 の無人実行で再利用）
'
'   取込の流れ
'     1. ファイルの事前チェック（存在・拡張子・サイズ上限）
'     2. 文字コード判定・パース
'     3. ヘッダーとSchemaの対応付け（列順が違っても取込可）
'     4. 1行ずつ検証（型・必須・形式・範囲・CSVインジェクション・重複）
'     5. エラー一覧の出力 → バックアップ → Ordersへ一括書込
'==========================================================

Private Const ORDERS_SHEET As String = "Orders"
Private Const ORDERS_TABLE As String = "tblOrders"
Private Const ERRORS_SHEET As String = "ImportErrors"
Private Const AUDIT_COLS As Long = 3        ' 取込日時 / 取込ID / 取込ファイル

'----------------------------------------------------------
' UI：Alt+F8 やボタンから呼ぶ
'----------------------------------------------------------
Public Sub ImportOrdersCsv()
    Const PROC As String = "modImport.ImportOrdersCsv"
    Dim filePath As Variant
    Dim res As clsImportResult
    On Error GoTo EH

    ResetRunId
    ReloadConfig

    filePath = SelectCsvFile()
    If VarType(filePath) = vbBoolean Then
        LogInfo PROC, "ファイル選択がキャンセルされました"
        Exit Sub
    End If

    Set res = ImportOrdersFromFile(CStr(filePath))

    If res.ErrorRows > 0 Then
        ThisWorkbook.Worksheets(ERRORS_SHEET).Activate
        MsgBox res.summary() & vbCrLf & vbCrLf & _
               IIf(res.Committed, "エラー行を除いて取り込みました。", "エラーがあるため、取り込みを中止しました。") & vbCrLf & _
               "「" & ERRORS_SHEET & "」シートを確認し、CSVを修正して再度取り込んでください。", _
               vbExclamation, AppTitle()
    Else
        ThisWorkbook.Worksheets(ORDERS_SHEET).Activate
        MsgBox res.summary(), vbInformation, AppTitle()
    End If
    Exit Sub

EH:
    HandleError PROC
End Sub

'----------------------------------------------------------
' 処理層：画面操作（MsgBox・ファイル選択）を一切行わない
'----------------------------------------------------------
Public Function ImportOrdersFromFile(ByVal filePath As String) As clsImportResult
    Const PROC As String = "modImport.ImportOrdersFromFile"
    Dim guard As clsAppState
    Dim res As clsImportResult
    Dim startTime As Double
    Dim text As String
    Dim enc As String
    Dim data As Variant
    Dim fieldCounts() As Long
    Dim defs() As ColumnDef
    Dim colMap() As Long
    Dim keyIdx As Long
    Dim existingKeys As Object
    Dim fileKeys As Object
    Dim outRows() As Variant
    Dim rowVals() As Variant
    Dim errs As Collection
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim rowErrCount As Long
    Dim raw As String
    Dim msg As String
    Dim converted As Variant
    Dim v As Variant
    Dim injectionPolicy As String
    Dim importMode As String
    Dim importedAt As Date
    Dim fileName As String

    startTime = Timer
    Set guard = New clsAppState          ' エラーで抜けても画面更新等は自動で元に戻る
    Set res = New clsImportResult
    res.filePath = filePath
    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    LogInfo PROC, "取込開始: " & filePath

    '--- 1. 事前チェック
    CheckInputFile filePath

    '--- 2. 読込・パース
    guard.ShowProgress "CSVを読み込んでいます..."
    text = ReadTextFile(filePath, enc)
    res.EncodingName = enc
    data = ParseCsv(text, fieldCounts)
    text = vbNullString                  ' 大きな文字列のメモリを解放

    nRows = UBound(data, 1) - 1
    res.TotalRows = nRows
    If nRows <= 0 Then RaiseAppError ERR_CSV_FORMAT, PROC, "データ行がありません（ヘッダーのみ）"
    If nRows > CLng(GetConfig("MaxImportRows", 200000)) Then
        RaiseAppError ERR_FILE_TOO_LARGE, PROC, "行数が上限（" & GetConfig("MaxImportRows", 200000) & " 行）を超えています"
    End If

    '--- 3. ヘッダーとSchemaの対応付け
    defs = LoadSchema()
    nCols = UBound(defs)
    colMap = MapHeader(data, fieldCounts(1), defs)
    keyIdx = FindDef(defs, CStr(GetConfig("ImportKeyColumn")))

    EnsureOrdersTable defs
    Set existingKeys = LoadExistingKeys(defs(keyIdx).Name)
    Set fileKeys = CreateObject("Scripting.Dictionary")

    injectionPolicy = UCase$(CStr(GetConfig("CsvInjectionPolicy", "SANITIZE")))
    importMode = UCase$(CStr(GetConfig("ImportMode", "ALL_OR_NOTHING")))
    importedAt = Now

    ReDim outRows(1 To nRows, 1 To nCols + AUDIT_COLS)
    ReDim rowVals(1 To nCols)
    Set errs = New Collection

    '--- 4. 1行ずつ検証（r はCSVのレコード番号。ヘッダー = 1）
    For r = 2 To nRows + 1
        If r Mod 1000 = 0 Then guard.ShowProgress "検証中... " & Format$(r - 1, "#,##0") & " / " & Format$(nRows, "#,##0")
        rowErrCount = 0

        If fieldCounts(r) <> fieldCounts(1) Then
            AddError errs, r, "(行全体)", "", "列数が一致しません（期待：" & fieldCounts(1) & " / 実際：" & fieldCounts(r) & "）"
            rowErrCount = 1
        Else
            For c = 1 To nCols
                raw = data(r, colMap(c))
                msg = ValidateValue(defs(c), raw, converted)

                ' CSVインジェクション対策（文字列項目のみ。数値の "-5" 等は型検証済み）
                If Len(msg) = 0 And defs(c).dataType = "TEXT" Then
                    If IsFormulaLike(CStr(converted)) Then
                        If injectionPolicy = "REJECT" Then
                            msg = "数式として解釈される文字（= + - @ 等）で始まっています"
                        Else
                            res.FormulaLikeCount = res.FormulaLikeCount + 1
                            LogWarn PROC, "数式様の値を文字列として取込: " & r & "行目 / " & defs(c).Name
                        End If
                    End If
                End If

                If Len(msg) > 0 Then
                    AddError errs, r, defs(c).Name, raw, msg
                    rowErrCount = rowErrCount + 1
                Else
                    rowVals(c) = converted
                End If
            Next c

            ' 重複チェック（二重取込の防止）
            If rowErrCount = 0 Then
                raw = CStr(rowVals(keyIdx))
                If existingKeys.Exists(raw) Then
                    AddError errs, r, defs(keyIdx).Name, raw, "取込済みのデータと重複しています"
                    rowErrCount = 1
                ElseIf fileKeys.Exists(raw) Then
                    AddError errs, r, defs(keyIdx).Name, raw, "ファイル内で重複しています（" & fileKeys(raw) & "行目）"
                    rowErrCount = 1
                Else
                    fileKeys.Add raw, r
                End If
            End If
        End If

        If rowErrCount = 0 Then
            res.ValidRows = res.ValidRows + 1
            For c = 1 To nCols
                v = rowVals(c)
                If defs(c).dataType = "TEXT" Then v = ToSafeCellText(CStr(v))
                outRows(res.ValidRows, c) = v
            Next c
            outRows(res.ValidRows, nCols + 1) = importedAt
            outRows(res.ValidRows, nCols + 2) = ToSafeCellText(RunId())
            outRows(res.ValidRows, nCols + 3) = ToSafeCellText(fileName)
        Else
            res.ErrorRows = res.ErrorRows + 1
        End If
    Next r

    '--- 5. 出力
    guard.ShowProgress "結果を出力しています..."
    WriteErrors errs, CLng(GetConfig("MaxErrorReport", 1000))

    ' ALL_OR_NOTHING：1行でもエラーがあれば取り込まない / PARTIAL：正常行のみ取り込む
    res.Committed = (res.ValidRows > 0) And (res.ErrorRows = 0 Or importMode = "PARTIAL")
    If res.Committed Then
        If CBool(GetConfig("BackupBeforeImport", True)) Then
            guard.ShowProgress "バックアップを作成しています..."
            CreateBackup
        End If
        guard.ShowProgress "Ordersに書き込んでいます..."
        AppendOrders outRows, res.ValidRows, nCols + AUDIT_COLS
        res.ImportedRows = res.ValidRows
    End If

    res.ElapsedSeconds = Timer - startTime
    LogInfo PROC, "取込終了: " & Replace(res.summary(), vbCrLf, " / ")
    Set ImportOrdersFromFile = res
End Function

'==========================================================
' 内部処理
'==========================================================
Private Function SelectCsvFile() As Variant
    Dim initDir As String

    On Error Resume Next             ' 初期フォルダの設定に失敗しても選択はできるようにする
    initDir = ResolvePath("InputDir", "samples")
    If FolderExists(initDir) Then
        ChDrive Left$(initDir, 1)
        ChDir initDir
    End If
    On Error GoTo 0

    SelectCsvFile = Application.GetOpenFilename("CSVファイル (*.csv),*.csv", , "取り込む受注CSVを選択してください")
End Function

Private Sub CheckInputFile(ByVal filePath As String)
    Const PROC As String = "modImport.CheckInputFile"
    Dim limitMB As Double

    If Not FileExists(filePath) Then
        RaiseAppError ERR_FILE_NOT_FOUND, PROC, "ファイルが見つかりません：" & filePath
    End If
    If LCase$(Extension(filePath)) <> ".csv" Then
        RaiseAppError ERR_CSV_FORMAT, PROC, "CSVファイル（.csv）を指定してください"
    End If

    ' 巨大ファイルによるメモリ枯渇・応答停止を防ぐ（DoS対策）
    limitMB = CDbl(GetConfig("MaxImportFileSizeMB", 20))
    If FileLen(filePath) > limitMB * 1024# * 1024# Then
        RaiseAppError ERR_FILE_TOO_LARGE, PROC, "ファイルサイズが上限（" & limitMB & " MB）を超えています"
    End If
End Sub

' 戻り値: map(i) = 項目定義 i がCSVの何列目にあるか
Private Function MapHeader(ByRef data As Variant, ByVal headerCount As Long, ByRef defs() As ColumnDef) As Long()
    Const PROC As String = "modImport.MapHeader"
    Dim idx As Object
    Dim map() As Long
    Dim c As Long
    Dim i As Long
    Dim h As String
    Dim missing As String

    Set idx = CreateObject("Scripting.Dictionary")
    For c = 1 To headerCount
        h = Trim$(CStr(data(1, c)))
        If idx.Exists(h) Then RaiseAppError ERR_CSV_FORMAT, PROC, "ヘッダーに同じ列名があります：" & h
        idx.Add h, c
    Next c

    ReDim map(1 To UBound(defs))
    For i = 1 To UBound(defs)
        If idx.Exists(defs(i).Name) Then
            map(i) = idx(defs(i).Name)
            idx.Remove defs(i).Name
        Else
            If Len(missing) > 0 Then missing = missing & "、"
            missing = missing & defs(i).Name
        End If
    Next i

    If Len(missing) > 0 Then RaiseAppError ERR_CSV_FORMAT, PROC, "CSVに必要な列がありません：" & missing
    If idx.Count > 0 Then LogWarn PROC, "定義外の列を無視しました: " & Join(idx.keys, ", ")
    MapHeader = map
End Function

Private Function FindDef(ByRef defs() As ColumnDef, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To UBound(defs)
        If defs(i).Name = colName Then
            FindDef = i
            Exit Function
        End If
    Next i
    RaiseAppError ERR_SCHEMA, "modImport.FindDef", "Schemaに項目がありません：" & colName & "（Config の ImportKeyColumn を確認してください）"
End Function

' Ordersシート・テーブルが無ければ作成。あれば列構成がSchemaと一致するか確認
Private Sub EnsureOrdersTable(ByRef defs() As ColumnDef)
    Const PROC As String = "modImport.EnsureOrdersTable"
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers() As Variant
    Dim n As Long
    Dim i As Long

    n = UBound(defs)
    ReDim headers(1 To 1, 1 To n + AUDIT_COLS)
    For i = 1 To n
        headers(1, i) = defs(i).Name
    Next i
    headers(1, n + 1) = "取込日時"
    headers(1, n + 2) = "取込ID"
    headers(1, n + 3) = "取込ファイル"

    Set ws = GetOrCreateSheet(ORDERS_SHEET)

    If ws.ListObjects.Count = 0 Then
        ws.Range("A1").Resize(1, n + AUDIT_COLS).value = headers
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(1, n + AUDIT_COLS), , xlYes)
        lo.Name = ORDERS_TABLE
        For i = 1 To n
            Select Case defs(i).dataType
                Case "DATE":    ws.Columns(i).NumberFormat = "yyyy/mm/dd"
                Case "INTEGER": ws.Columns(i).NumberFormat = "#,##0"
            End Select
        Next i
        ws.Columns(n + 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
        LogInfo PROC, "Ordersテーブルを作成しました"
        Exit Sub
    End If

    Set lo = ws.ListObjects(1)
    If lo.ListColumns.Count <> n + AUDIT_COLS Then GoTo Mismatch
    For i = 1 To n + AUDIT_COLS
        If CStr(lo.HeaderRowRange.Cells(1, i).value) <> headers(1, i) Then GoTo Mismatch
    Next i
    Exit Sub

Mismatch:
    RaiseAppError ERR_SCHEMA, PROC, "Ordersシートの列構成がSchemaと一致しません。" & vbCrLf & _
                  "Schemaを変更した場合は、Ordersシートの名前を変更（退避）してから再実行してください。"
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function LoadExistingKeys(ByVal keyName As String) As Object
    Dim keys As Object
    Dim lo As ListObject
    Dim data As Variant
    Dim i As Long
    Dim k As String

    Set keys = CreateObject("Scripting.Dictionary")
    Set lo = ThisWorkbook.Worksheets(ORDERS_SHEET).ListObjects(1)

    If Not lo.DataBodyRange Is Nothing Then
        If lo.ListRows.Count = 1 Then
            k = CStr(lo.ListColumns(keyName).DataBodyRange.value)     ' 1セルだと配列にならない
            If Len(k) > 0 Then keys(k) = True
        Else
            data = lo.ListColumns(keyName).DataBodyRange.value
            For i = 1 To UBound(data, 1)
                k = CStr(data(i, 1))
                If Len(k) > 0 Then keys(k) = True
            Next i
        End If
    End If
    Set LoadExistingKeys = keys
End Function

' 配列をテーブル末尾へ一括書込（セル単位の書込を避け高速化）
Private Sub AppendOrders(ByRef outRows() As Variant, ByVal rowCount As Long, ByVal colCount As Long)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim startRow As Long
    Dim firstCol As Long

    Set ws = ThisWorkbook.Worksheets(ORDERS_SHEET)
    Set lo = ws.ListObjects(1)
    firstCol = lo.Range.Column

    If lo.DataBodyRange Is Nothing Then
        startRow = lo.HeaderRowRange.Row + 1
    ElseIf lo.ListRows.Count = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
        startRow = lo.DataBodyRange.Row                      ' 作成直後の空行を上書き
    Else
        startRow = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count
    End If

    ' outRows は検証OK行が先頭に詰めてあるため、rowCount 行分だけ書き込まれる
    ws.Cells(startRow, firstCol).Resize(rowCount, colCount).value = outRows
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1, 1), ws.Cells(startRow + rowCount - 1, firstCol + colCount - 1))
End Sub

Private Sub AddError(ByVal errs As Collection, ByVal rowNo As Long, ByVal colName As String, _
                     ByVal value As String, ByVal message As String)
    errs.Add Array(rowNo, colName, value, message)
End Sub

' エラー一覧を ImportErrors シートへ出力（毎回クリアして書き直す）
Private Sub WriteErrors(ByVal errs As Collection, ByVal maxReport As Long)
    Dim ws As Worksheet
    Dim out() As Variant
    Dim n As Long
    Dim i As Long
    Dim e As Variant

    Set ws = GetOrCreateSheet(ERRORS_SHEET)
    ws.Cells.Clear
    ws.Range("A1:D1").value = Array("CSV行番号", "列名", "値", "エラー内容")
    ws.Range("A1:D1").Font.Bold = True

    If errs.Count = 0 Then
        ws.Range("A2").value = "エラーはありません（" & Format$(Now, "yyyy/mm/dd hh:nn:ss") & " 取込分）"
        Exit Sub
    End If

    n = errs.Count
    If n > maxReport Then n = maxReport
    ReDim out(1 To n, 1 To 4)

    i = 0
    For Each e In errs
        i = i + 1
        If i > n Then Exit For
        out(i, 1) = e(0)
        out(i, 2) = ToSafeCellText(CStr(e(1)))
        out(i, 3) = ToSafeCellText(CStr(e(2)))     ' CSVの生の値：数式として評価させない
        out(i, 4) = ToSafeCellText(CStr(e(3)))
    Next e

    ws.Range("A2").Resize(n, 4).value = out
    If errs.Count > n Then
        ws.Cells(n + 3, 1).value = "※ エラーが多いため先頭 " & n & " 件のみ表示しています（全 " & errs.Count & " 件）"
    End If

    ws.Columns("A:D").AutoFit
    If ws.Columns("C").ColumnWidth > 50 Then ws.Columns("C").ColumnWidth = 50
End Sub


