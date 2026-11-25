Attribute VB_Name = "modProfile"
Option Explicit
'==========================================================
' modProfile : 取込処理の詳細プロファイル（開発・性能分析用）
'   RunPerformanceTest が生成した output\perf\perf_100000.csv を使い、
'   本番の取込とは別に、処理を部品ごとに分けて計測する
'     1. 読込・パース
'     2. 検証（項目ごとの時間）
'     3. 1行ごとの付帯処理（数式判定・安全化・重複チェック）
'     4. セル書込の方式比較（' 付き文字列 / 文字列書式 "@" ＋ 生の文字列）
'   ※ Orders には書き込まない（作業用シートを作成し、最後に削除する）
'==========================================================

Public Sub ProfileImport()
    Const PROC As String = "modProfile.ProfileImport"
    Dim guard As clsAppState
    Dim sw As clsStopwatch
    Dim csvPath As String
    Dim enc As String
    Dim text As String
    Dim data As Variant
    Dim fc() As Long
    Dim defs() As ColumnDef
    Dim colMap() As Long
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim h As Long
    Dim v As Variant
    Dim msg As String
    Dim md As String
    Dim colTimes() As Double
    Dim tRead As Double, tParse As Double, tRowOps As Double
    Dim tWritePrefixed As Double, tWriteTextFmt As Double
    Dim outPrefixed() As Variant
    Dim outRaw() As Variant
    Dim conv() As Variant
    Dim msgs() As String
    Dim keys As Object
    Dim k As String
    Dim ws As Worksheet
    Dim totalCols As Long
    Dim formulaSafe As String
    Dim mdPath As String
    On Error GoTo EH

    ResetRunId
    ReloadConfig
    RequirePermission actAdmin
    csvPath = ResolvePath("OutputDir", "output") & "\perf\perf_100000.csv"
    If Not FileExists(csvPath) Then
        RaiseAppError ERR_FILE_NOT_FOUND, PROC, "先に RunPerformanceTest を実行して、テストデータを生成してください。"
    End If

    Set guard = New clsAppState
    Set sw = New clsStopwatch

    '--- 1. 読込・パース
    guard.ShowProgress "読込・パース..."
    text = ReadTextFile(csvPath, enc)
    tRead = sw.ElapsedSeconds: sw.StartTimer
    data = ParseCsv(text, fc)
    text = vbNullString
    tParse = sw.ElapsedSeconds

    nRows = UBound(data, 1) - 1
    defs = LoadSchema()
    nCols = UBound(defs)
    totalCols = nCols + 3

    ReDim colMap(1 To nCols)
    For c = 1 To nCols
        For h = 1 To fc(1)
            If data(1, h) = defs(c).Name Then colMap(c) = h
        Next h
    Next c

    '--- 2. 検証（項目ごと。列を外側のループにして列単位で計測）
    ReDim colTimes(1 To nCols)
    ReDim conv(1 To nRows + 1, 1 To nCols)
    ReDim msgs(1 To nRows + 1, 1 To nCols)
    For c = 1 To nCols
        guard.ShowProgress "検証: " & defs(c).Name
        sw.StartTimer
        ValidateColumn defs(c), data, colMap(c), conv, msgs, c
        colTimes(c) = sw.ElapsedSeconds
    Next c

    '--- 3. 1行ごとの付帯処理
    guard.ShowProgress "付帯処理..."
    ReDim outPrefixed(1 To nRows, 1 To totalCols)
    ReDim outRaw(1 To nRows, 1 To totalCols)
    Set keys = CreateObject("Scripting.Dictionary")
    sw.StartTimer
    For r = 1 To nRows
        For c = 1 To nCols
            v = conv(r + 1, c)
            If defs(c).dataType = "TEXT" Then
                If IsFormulaLike(CStr(v)) Then msg = "x"
                outRaw(r, c) = v
                v = ToSafeCellText(CStr(v))
            Else
                outRaw(r, c) = v
            End If
            outPrefixed(r, c) = v
        Next c
        k = CStr(conv(r + 1, 1))
        If Not keys.Exists(k) Then keys.Add k, r
        outPrefixed(r, nCols + 1) = Now: outRaw(r, nCols + 1) = outPrefixed(r, nCols + 1)
        outPrefixed(r, nCols + 2) = "'PROFILE": outRaw(r, nCols + 2) = "PROFILE"
        outPrefixed(r, nCols + 3) = "'perf_100000.csv": outRaw(r, nCols + 3) = "perf_100000.csv"
    Next r
    tRowOps = sw.ElapsedSeconds
    Erase conv
    data = Empty

    '--- 4. セル書込の方式比較（作業用シート）
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "_Profile_" & Format$(Now, "hhnnss")

    guard.ShowProgress "書込A: ' 付き文字列..."
    sw.StartTimer
    ws.Range("A1").Resize(nRows, totalCols).value = outPrefixed
    tWritePrefixed = sw.ElapsedSeconds
    ws.Cells.Clear

    guard.ShowProgress "書込B: 文字列書式＋生の文字列..."
    sw.StartTimer
    For c = 1 To totalCols
        If c > nCols Then
            If c <> nCols + 1 Then ws.Columns(c).NumberFormat = "@"
        ElseIf defs(c).dataType = "TEXT" Then
            ws.Columns(c).NumberFormat = "@"
        End If
    Next c
    ws.Range("A1").Resize(nRows, totalCols).value = outRaw
    tWriteTextFmt = sw.ElapsedSeconds

    ' 方式Bの安全性確認：文字列書式のセルに数式様の値を書いて、数式にならないか
    ws.Cells(1, 1).value = "=1+1"
    If ws.Cells(1, 1).HasFormula Then
        formulaSafe = "NG（数式として評価された）"
    Else
        formulaSafe = "OK（文字列のまま：" & ws.Cells(1, 1).value & "）"
    End If

    ws.Delete
    Set guard = Nothing

    '--- 5. 結果
    md = "## 取込プロファイル（" & Format$(Now, "yyyy/mm/dd hh:nn") & " / " & Format$(nRows, "#,##0") & " 行）" & vbCrLf & vbCrLf
    md = md & "| 区間 | 時間 |" & vbCrLf & "|---|---:|" & vbCrLf
    md = md & "| 読込（文字コード判定＋復号） | " & Format$(tRead, "0.00") & " 秒 |" & vbCrLf
    md = md & "| パース | " & Format$(tParse, "0.00") & " 秒 |" & vbCrLf
    For c = 1 To nCols
        md = md & "| 検証：" & defs(c).Name & "（" & defs(c).dataType & IIf(Len(defs(c).pattern) > 0, "＋形式", "") & _
                  "） | " & Format$(colTimes(c), "0.00") & " 秒 |" & vbCrLf
    Next c
    md = md & "| 付帯処理（数式判定・' 付与・重複チェック） | " & Format$(tRowOps, "0.00") & " 秒 |" & vbCrLf
    md = md & "| 書込A：' 付き文字列 | " & Format$(tWritePrefixed, "0.00") & " 秒 |" & vbCrLf
    md = md & "| 書込B：文字列書式＋生の文字列 | " & Format$(tWriteTextFmt, "0.00") & " 秒 |" & vbCrLf
    md = md & vbCrLf & "方式Bで ""=1+1"" を書き込んだ結果：" & formulaSafe & vbCrLf

    mdPath = ResolvePath("OutputDir", "output") & "\perf\profile_" & RunId() & ".md"
    WriteTextFile mdPath, md, False
    Debug.Print md
    LogInfo PROC, "プロファイル完了: " & mdPath
    MsgBox "プロファイルが完了しました。" & vbCrLf & mdPath, vbInformation, AppTitle()
    Exit Sub

EH:
    HandleError PROC
End Sub


