Attribute VB_Name = "modPerfTest"
Option Explicit
'==========================================================
' modPerfTest : 性能計測（非機能要件の検証）
'   1. テストデータCSVを生成（乱数の種を固定 → 毎回同じデータで再現性を確保）
'   2. 取込 → 集計・帳票作成 → 取消 を Config!PerfTestRepeat 回くり返して計測
'   3. 各回の値と「中央値」を Markdown で出力（docs/performance.md に貼り付ける）
'      ※ 同じ条件でも1回ごとに数秒ぶれるため、1回だけの値では改善前後を比較できない。
'        外れ値に強い中央値で判定する
'   ※ 取り込んだテストデータは各回の計測後に自動で削除する
'==========================================================

Private Const TARGET_IMPORT_SEC As Double = 30     ' 目標：10万行の取込が30秒以内
Private Const TARGET_REPORT_SEC As Double = 10     ' 目標：10万行の集計・帳票作成が10秒以内

Public Sub RunPerformanceTest()
    Const PROC As String = "modPerfTest.RunPerformanceTest"
    Dim rows As Long
    Dim repeatCount As Long
    Dim i As Long
    Dim sw As clsStopwatch
    Dim csvPath As String
    Dim importId As String
    Dim impRes As clsImportResult
    Dim rep As clsReportData
    Dim tGen As Double
    Dim tImport() As Double
    Dim tReport() As Double
    Dim tRemove() As Double
    Dim importDetail() As String
    Dim reportDetail() As String
    Dim md As String
    Dim mdPath As String
    On Error GoTo EH

    ReloadConfig
    rows = CLng(GetConfig("PerfTestRows", 100000))
    repeatCount = CLng(GetConfig("PerfTestRepeat", 3))
    If repeatCount < 1 Then repeatCount = 1
    If MsgBox(Format$(rows, "#,##0") & " 行のテストデータで性能計測を " & repeatCount & " 回行います。" & vbCrLf & _
              "取り込んだテストデータは各回の計測後に自動で削除されます。" & vbCrLf & _
              "（" & repeatCount & "〜" & repeatCount * 2 & " 分ほどかかります）" & vbCrLf & vbCrLf & "実行しますか？", _
              vbQuestion + vbYesNo + vbDefaultButton2, AppTitle()) <> vbYes Then Exit Sub

    ResetRunId
    LogInfo PROC, "性能計測開始: " & rows & " 行 × " & repeatCount & " 回"
    Set sw = New clsStopwatch

    '--- 1. テストデータ生成（参考値。全回で同じファイルを使う）
    Application.StatusBar = "テストデータを生成しています..."
    csvPath = GeneratePerfCsv(rows)
    tGen = sw.ElapsedSeconds

    ReDim tImport(1 To repeatCount)
    ReDim tReport(1 To repeatCount)
    ReDim tRemove(1 To repeatCount)
    ReDim importDetail(1 To repeatCount)
    ReDim reportDetail(1 To repeatCount)

    For i = 1 To repeatCount
        '--- 2. 取込（読込・文字コード判定・検証・書込・バックアップを含む）
        Application.StatusBar = "計測 " & i & " / " & repeatCount & "：取込..."
        sw.StartTimer
        Set impRes = ImportOrdersFromFile(csvPath)
        tImport(i) = sw.ElapsedSeconds
        importDetail(i) = impRes.TimingDetail
        importId = RunId()

        '--- 3. 集計・帳票作成（テストデータの期間 = 2026年10月〜12月）
        If impRes.ImportedRows > 0 Then
            Application.StatusBar = "計測 " & i & " / " & repeatCount & "：集計..."
            sw.StartTimer
            Set rep = BuildReport(DateSerial(2026, 10, 1), DateSerial(2026, 12, 31))
            tReport(i) = sw.ElapsedSeconds
            reportDetail(i) = rep.TimingDetail
        End If

        '--- 4. 後片付け（取消）
        sw.StartTimer
        RemoveImportBatch importId, False
        tRemove(i) = sw.ElapsedSeconds

        If impRes.ImportedRows <> rows Then
            Application.StatusBar = False
            RaiseAppError ERR_VALIDATION, PROC, "テストデータの取込に失敗しました（取込 " & impRes.ImportedRows & _
                          " / エラー行 " & impRes.ErrorRows & "）。ImportErrors シートを確認してください。"
        End If
        LogInfo PROC, "計測 " & i & " 回目: 取込 " & Format$(tImport(i), "0.00") & " 秒 / 集計 " & _
                      Format$(tReport(i), "0.00") & " 秒 / 取消 " & Format$(tRemove(i), "0.00") & " 秒"
    Next i
    Application.StatusBar = False

    '--- 5. 結果出力
    md = BuildResultMarkdown(rows, repeatCount, csvPath, tGen, tImport, tReport, tRemove, _
                             importDetail, reportDetail, rep)
    mdPath = ResolvePath("OutputDir", "output") & "\perf\perf_result_" & RunId() & ".md"
    WriteTextFile mdPath, md, False
    Debug.Print md

    LogInfo PROC, "性能計測終了（中央値）: 取込 " & Format$(Median(tImport), "0.00") & " 秒 / 集計 " & _
                  Format$(Median(tReport), "0.00") & " 秒 / 取消 " & Format$(Median(tRemove), "0.00") & " 秒"

    MsgBox "性能計測が完了しました（" & repeatCount & " 回の中央値）。" & vbCrLf & vbCrLf & _
           "取込　　　：" & Format$(Median(tImport), "0.00") & " 秒" & Verdict(Median(tImport), TARGET_IMPORT_SEC) & vbCrLf & _
           "集計・帳票：" & Format$(Median(tReport), "0.00") & " 秒" & Verdict(Median(tReport), TARGET_REPORT_SEC) & vbCrLf & _
           "取消　　　：" & Format$(Median(tRemove), "0.00") & " 秒" & vbCrLf & vbCrLf & _
           "結果（Markdown）：" & vbCrLf & mdPath, vbInformation, AppTitle()
    Exit Sub

EH:
    Application.StatusBar = False
    HandleError PROC
End Sub

'----------------------------------------------------------
' テストデータCSVを生成。戻り値：ファイルパス
'----------------------------------------------------------
Private Function GeneratePerfCsv(ByVal rows As Long) As String
    Dim custs As Variant
    Dim prods As Variant
    Dim prices As Variant
    Dim reps As Variant
    Dim lines() As String
    Dim i As Long
    Dim p As Long
    Dim baseDate As Date
    Dim folder As String
    Dim filePath As String
    Dim seedReset As Single

    custs = Array("C00001,サンプル商事株式会社", "C00002,テスト工業株式会社", "C00003,架空物産株式会社", _
                  "C00004,ダミー電機株式会社", "C00005,例示食品株式会社")
    prods = Array("P000001,ボールペン（黒）10本入", "P000002,A4コピー用紙 500枚", "P000003,ホッチキス針 No.10", _
                  "P000004,クリアファイル 20枚入", "P000005,付箋 75mm角")
    prices = Array(450, 520, 180, 300, 260)
    reps = Array("山田 太郎", "佐藤 花子", "鈴木 一郎")
    baseDate = DateSerial(2026, 10, 1)

    seedReset = Rnd(-1)                 ' Rnd(負数) → Randomize(数値) の順で呼ぶと乱数列が固定される
    Randomize 20261107                  ' 種を固定：何度実行しても同じデータになる

    ReDim lines(0 To rows)
    lines(0) = "受注番号,受注日,顧客コード,顧客名,商品コード,商品名,数量,単価,担当者,備考"
    For i = 1 To rows
        p = Int(Rnd * 5)
        lines(i) = "ORD-9" & Format$(i, "0000000") & "," & _
                   Format$(baseDate + Int(Rnd * 92), "yyyy/mm/dd") & "," & _
                   custs(Int(Rnd * 5)) & "," & _
                   prods(p) & "," & _
                   (Int(Rnd * 50) + 1) & "," & _
                   prices(p) & "," & _
                   reps(Int(Rnd * 3)) & ","
    Next i

    folder = ResolvePath("OutputDir", "output") & "\perf"
    EnsureFolder folder
    filePath = folder & "\perf_" & rows & ".csv"
    WriteTextFile filePath, Join(lines, vbCrLf) & vbCrLf, True
    GeneratePerfCsv = filePath
End Function

Private Function BuildResultMarkdown(ByVal rows As Long, ByVal repeatCount As Long, ByVal csvPath As String, _
                                     ByVal tGen As Double, ByRef tImport() As Double, ByRef tReport() As Double, _
                                     ByRef tRemove() As Double, ByRef importDetail() As String, _
                                     ByRef reportDetail() As String, ByVal rep As clsReportData) As String
    Dim s As String
    Dim bitness As String
    Dim i As Long
    Dim mImport As Double
    Dim mReport As Double

    #If Win64 Then
        bitness = "64bit"
    #Else
        bitness = "32bit"
    #End If
    mImport = Median(tImport)
    mReport = Median(tReport)

    s = "## 性能計測結果（" & Format$(Now, "yyyy/mm/dd hh:nn") & "）" & vbCrLf & vbCrLf
    s = s & "### 計測条件" & vbCrLf & vbCrLf
    s = s & "| 項目 | 値 |" & vbCrLf & "|---|---|" & vbCrLf
    s = s & "| データ件数 | " & Format$(rows, "#,##0") & " 行 |" & vbCrLf
    s = s & "| CSVサイズ | " & Format$(FileLen(csvPath) / 1024# / 1024#, "0.0") & " MB（UTF-8） |" & vbCrLf
    s = s & "| 計測回数 | " & repeatCount & " 回（判定は中央値） |" & vbCrLf
    s = s & "| Excel | " & Application.Version & "（Build " & Application.Build & " / " & bitness & "） |" & vbCrLf
    s = s & "| OS | " & Application.OperatingSystem & " |" & vbCrLf
    s = s & "| CPU | " & Environ$("PROCESSOR_IDENTIFIER") & "（論理 " & Environ$("NUMBER_OF_PROCESSORS") & " コア） |" & vbCrLf

    s = s & vbCrLf & "### 結果（中央値）" & vbCrLf & vbCrLf
    s = s & "| 処理 | 時間 | 目標 | 判定 |" & vbCrLf & "|---|---:|---:|:---:|" & vbCrLf
    s = s & "| テストデータ生成（参考） | " & Format$(tGen, "0.00") & " 秒 | - | - |" & vbCrLf
    s = s & "| 取込（読込・検証・書込・バックアップ） | " & Format$(mImport, "0.00") & " 秒 | " & _
            TARGET_IMPORT_SEC & " 秒以内 | " & IIf(mImport <= TARGET_IMPORT_SEC, "OK", "NG") & " |" & vbCrLf
    s = s & "| 集計・帳票作成 | " & Format$(mReport, "0.00") & " 秒 | " & _
            TARGET_REPORT_SEC & " 秒以内 | " & IIf(mReport <= TARGET_REPORT_SEC, "OK", "NG") & " |" & vbCrLf
    s = s & "| 取消（テストデータの削除） | " & Format$(Median(tRemove), "0.00") & " 秒 | - | - |" & vbCrLf

    s = s & vbCrLf & "### 各回の値" & vbCrLf & vbCrLf
    s = s & "| 回 | 取込 | 集計・帳票 | 取消 |" & vbCrLf & "|---:|---:|---:|---:|" & vbCrLf
    For i = 1 To repeatCount
        s = s & "| " & i & " | " & Format$(tImport(i), "0.00") & " 秒 | " & Format$(tReport(i), "0.00") & _
                " 秒 | " & Format$(tRemove(i), "0.00") & " 秒 |" & vbCrLf
    Next i

    s = s & vbCrLf & "### 処理時間の内訳（各回）" & vbCrLf & vbCrLf
    For i = 1 To repeatCount
        s = s & "- " & i & " 回目 取込：" & importDetail(i) & vbCrLf
        s = s & "- " & i & " 回目 集計・帳票：" & reportDetail(i) & vbCrLf
    Next i

    If Not rep Is Nothing Then
        s = s & vbCrLf & "集計対象：" & Format$(rep.OrderCount, "#,##0") & " 件 / 受注金額合計 " & _
                Format$(rep.TotalAmount, "#,##0") & " 円" & vbCrLf
    End If
    BuildResultMarkdown = s
End Function

' 中央値（外れ値に強い。並べ替えは件数が少ないので単純な挿入ソート）
Private Function Median(ByRef values() As Double) As Double
    Dim arr() As Double
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Double

    n = UBound(values) - LBound(values) + 1
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = values(LBound(values) + i - 1)
    Next i
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) <= tmp Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i

    If n Mod 2 = 1 Then
        Median = arr((n + 1) \ 2)
    Else
        Median = (arr(n \ 2) + arr(n \ 2 + 1)) / 2
    End If
End Function

Private Function Verdict(ByVal sec As Double, ByVal target As Double) As String
    If sec <= target Then
        Verdict = "（目標 " & target & " 秒以内：OK）"
    Else
        Verdict = "（目標 " & target & " 秒以内：NG）"
    End If
End Function


