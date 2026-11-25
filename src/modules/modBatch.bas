Attribute VB_Name = "modBatch"
Option Explicit
'==========================================================
' modBatch : 無人実行（バッチ）
'   scripts\run_batch.vbs から Application.Run で呼ばれ、画面を一切出さずに最後まで処理する
'
'   流れ
'     1. 二重起動の防止（ロックファイル）
'     2. 受信フォルダ（BatchInboxDir）の CSV を古い順に取込
'          成功   → 処理済みフォルダ（BatchProcessedDir\yyyymm）へ移動
'          エラー → エラーフォルダ（BatchErrorDir）へ移動し、エラー一覧の CSV を横に置く
'          ※ 1ファイルの失敗で全体を止めず、次のファイルへ進む
'     3. 1件以上取り込んだらレポート（PDF / CSV）を作成
'     4. ブックを保存し、結果を JSON・イベントログ・通知（Webhook）に出力
'
'   戻り値（終了コード）：0 = 正常 / 1 = 警告（一部のファイルがエラー） / 2 = 異常終了 / 3 = スキップ（実行中）
'   ※ 取込・集計は、画面操作から使っているものと同じ処理層の関数をそのまま呼び出す
'==========================================================

Public Const EXIT_OK As Long = 0
Public Const EXIT_WARN As Long = 1
Public Const EXIT_ERROR As Long = 2
Public Const EXIT_SKIPPED As Long = 3

' 1回のバッチ実行の結果（JSON・通知の作成に使う）
Private mStartedAt As Date
Private mFiles As Collection          ' 各要素：Array(ファイル名, 状態, 行数, 取込件数, エラー行数, メッセージ)
Private mReportPdf As String
Private mMessages As Collection       ' 全体に関するメッセージ

'----------------------------------------------------------
' バッチの入口（VBScript から呼ぶ）。戻り値は終了コード
'----------------------------------------------------------
Public Function RunBatch() As Long
    Const PROC As String = "modBatch.RunBatch"
    Dim lockPath As String
    Dim lockAcquired As Boolean
    Dim exitCode As Long
    On Error GoTo EH

    Set mFiles = New Collection             ' エラー処理でも使うため最初に用意する
    Set mMessages = New Collection
    mReportPdf = ""
    mStartedAt = Now
    SetBatchMode True
    ResetRunId
    ReloadConfig
    LogInfo PROC, "バッチ開始"

    ApplyProtection                         ' Workbook_Open を動かさずに開くため、ここで保護を掛け直す

    ' 他の人が開いていると読み取り専用になり、取込結果を保存できない
    If ThisWorkbook.ReadOnly Then
        RaiseAppError ERR_BATCH, PROC, "ブックが読み取り専用で開かれています（他のユーザーが使用中の可能性）"
    End If

    '--- 1. 二重起動の防止
    lockPath = ResolvePath("LogDir", "logs") & "\batch.lock"
    If Not AcquireLock(lockPath) Then
        mMessages.Add "前回のバッチが実行中のため、スキップしました"
        exitCode = EXIT_SKIPPED
        GoTo Finish
    End If
    lockAcquired = True

    '--- 2. 取込
    exitCode = ProcessInbox()

    '--- 3. レポート
    If ImportedFileCount() > 0 And CBool(GetConfig("BatchCreateReport", True)) Then
        CreateBatchReport
    End If

    '--- 4. 保存（取り込んだ場合のみ）
    If ImportedFileCount() > 0 Then SaveWorkbook

Finish:
    If lockAcquired Then ReleaseLock lockPath
    FinishBatch exitCode
    RunBatch = exitCode
    SetBatchMode False
    Exit Function

EH:
    ' 想定外のエラー：ログに残し、異常終了として通知する
    mMessages.Add "異常終了：" & Err.Description
    HandleError PROC                         ' バッチモードなので画面は出ない
    On Error Resume Next
    If lockAcquired Then ReleaseLock lockPath
    FinishBatch EXIT_ERROR
    RunBatch = EXIT_ERROR
    SetBatchMode False
End Function

'----------------------------------------------------------
' 手動でバッチを試す（Alt+F8 から実行。結果だけ画面に出す）
'----------------------------------------------------------
Public Sub RunBatchManual()
    Dim code As Long
    code = RunBatch()
    MsgBox "バッチが終了しました。" & vbCrLf & vbCrLf & _
           "結果：" & StatusName(code) & "（終了コード " & code & "）" & vbCrLf & vbCrLf & _
           BuildSummaryText(code), IIf(code = EXIT_OK, vbInformation, vbExclamation), AppTitle()
End Sub

'==========================================================
' 取込
'==========================================================
Private Function ProcessInbox() As Long
    Const PROC As String = "modBatch.ProcessInbox"
    Dim inbox As String
    Dim names() As String
    Dim n As Long
    Dim i As Long
    Dim minAge As Long
    Dim hasError As Boolean

    inbox = ResolvePath("BatchInboxDir", "batch\inbox")
    EnsureFolder inbox
    minAge = CLng(GetConfig("BatchMinFileAgeSec", 60))

    names = ListCsvFiles(inbox, n)
    If n = 0 Then
        mMessages.Add "取り込むファイルはありませんでした"
        LogInfo PROC, "受信フォルダにファイルがありません: " & inbox
        Exit Function                         ' EXIT_OK
    End If

    For i = 1 To n
        If IsFileReady(inbox & "\" & names(i), minAge) Then
            If Not ImportOneFile(inbox, names(i)) Then hasError = True
        Else
            ' 書込中かもしれないファイルは掴まず、次回に回す
            mFiles.Add Array(names(i), "SKIPPED", 0, 0, 0, "更新から " & minAge & " 秒経っていないため次回に取込")
            LogInfo PROC, "書込中の可能性があるためスキップ: " & names(i)
        End If
    Next i

    If hasError Then ProcessInbox = EXIT_WARN
End Function

' 1ファイルを取り込み、結果に応じて移動する。戻り値：成功なら True
Private Function ImportOneFile(ByVal inbox As String, ByVal fileName As String) As Boolean
    Const PROC As String = "modBatch.ImportOneFile"
    Dim src As String
    Dim res As clsImportResult
    Dim dest As String
    Dim destName As String
    Dim errDesc As String

    src = inbox & "\" & fileName
    On Error GoTo EH

    Set res = ImportOrdersFromFile(src)

    If res.Committed And res.ErrorRows = 0 Then
        dest = MoveFile(src, ResolvePath("BatchProcessedDir", "batch\processed") & "\" & Format$(Date, "yyyymm"))
        mFiles.Add Array(fileName, "OK", res.TotalRows, res.ImportedRows, 0, "")
        ImportOneFile = True
    Else
        ' ALL_OR_NOTHING でエラー行あり（取込なし）、または PARTIAL で一部エラー
        dest = MoveFile(src, ResolvePath("BatchErrorDir", "batch\error"))
        destName = BaseName(Mid$(dest, InStrRev(dest, "\") + 1)) & "_errors.csv"
        ExportImportErrors destName, Left$(dest, InStrRev(dest, "\") - 1)
        mFiles.Add Array(fileName, IIf(res.Committed, "PARTIAL", "REJECTED"), res.TotalRows, res.ImportedRows, _
                         res.ErrorRows, "エラー一覧：" & destName)
    End If
    LogInfo PROC, fileName & " → " & dest
    Exit Function

EH:
    ' ファイル単位の失敗（文字コード不正・列不足・権限など）は、記録して次のファイルへ進む
    errDesc = Err.Description
    LogError PROC, fileName & ": " & errDesc
    On Error Resume Next
    dest = MoveFile(src, ResolvePath("BatchErrorDir", "batch\error"))
    mFiles.Add Array(fileName, "FAILED", 0, 0, 0, errDesc)
End Function

' ImportErrors シートの内容をCSVに書き出す（エラーフォルダで原因が分かるように）
Private Sub ExportImportErrors(ByVal fileName As String, ByVal folder As String)
    Dim ws As Worksheet
    Dim data As Variant
    Dim lines() As String
    Dim r As Long
    Dim c As Long
    Dim lineText As String

    Set ws = GetSheetOrNothing("ImportErrors")
    If ws Is Nothing Then Exit Sub
    data = ws.UsedRange.value
    If Not IsArray(data) Then Exit Sub

    ReDim lines(1 To UBound(data, 1))
    For r = 1 To UBound(data, 1)
        lineText = ""
        For c = 1 To UBound(data, 2)
            If c > 1 Then lineText = lineText & ","
            lineText = lineText & ToCsvField(data(r, c))
        Next c
        lines(r) = lineText
    Next r
    WriteTextFile folder & "\" & fileName, Join(lines, vbCrLf) & vbCrLf, True
End Sub

'==========================================================
' レポート・保存
'==========================================================
Private Sub CreateBatchReport()
    Const PROC As String = "modBatch.CreateBatchReport"
    Dim rep As clsReportData
    Dim baseDate As Date
    On Error GoTo EH

    ' CURRENT = 今月 / PREVIOUS = 前月（月初に前月分を締める運用向け）
    If UCase$(CStr(GetConfig("BatchReportMonth", "CURRENT"))) = "PREVIOUS" Then
        baseDate = DateSerial(Year(Date), Month(Date) - 1, 1)
    Else
        baseDate = DateSerial(Year(Date), Month(Date), 1)
    End If

    Set rep = BuildReport(baseDate, DateSerial(Year(baseDate), Month(baseDate) + 1, 0))
    mReportPdf = ExportReportPdf(rep)
    ExportSummaryCsv rep
    Exit Sub
EH:
    ' レポートの失敗（対象月のデータなし等）は取込結果に影響させない
    mMessages.Add "レポートは作成されませんでした：" & Err.Description
    LogWarn PROC, "レポート作成をスキップ: " & Err.Description
End Sub

Private Sub SaveWorkbook()
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False          ' BeforeSave の確認ダイアログを出さない（保護は直前に掛ける）
    ApplyProtection
    ThisWorkbook.Save
    Application.EnableEvents = prevEvents
End Sub

'==========================================================
' 結果の出力
'==========================================================
Private Sub FinishBatch(ByVal exitCode As Long)
    Const PROC As String = "modBatch.FinishBatch"
    Dim jsonPath As String
    Dim summary As String

    On Error Resume Next                       ' 結果の出力に失敗しても、終了コードは返す
    summary = BuildSummaryText(exitCode)

    jsonPath = ResolvePath("OutputDir", "output") & "\batch"
    EnsureFolder jsonPath
    jsonPath = jsonPath & "\batch_" & RunId() & ".json"
    WriteTextFile jsonPath, BuildResultJson(exitCode), False

    LogInfo PROC, "バッチ終了: " & StatusName(exitCode) & " / " & Replace(summary, vbCrLf, " / ")
    AuditLog "BATCH", StatusName(exitCode), "files=" & mFiles.Count & " / imported=" & ImportedFileCount() & _
             " / result=" & jsonPath
    NotifyBatchResult exitCode, summary
End Sub

Public Function StatusName(ByVal exitCode As Long) As String
    Select Case exitCode
        Case EXIT_OK:      StatusName = "OK"
        Case EXIT_WARN:    StatusName = "WARN"
        Case EXIT_SKIPPED: StatusName = "SKIPPED"
        Case Else:         StatusName = "ERROR"
    End Select
End Function

Private Function BuildSummaryText(ByVal exitCode As Long) As String
    Dim s As String
    Dim f As Variant
    Dim m As Variant

    s = "実行ID：" & RunId() & vbCrLf
    s = s & "開始：" & Format$(mStartedAt, "yyyy/mm/dd hh:nn:ss") & "　終了：" & Format$(Now, "hh:nn:ss") & vbCrLf
    If Not mFiles Is Nothing Then
        For Each f In mFiles
            s = s & "・" & f(0) & "：" & f(1)
            If f(1) <> "SKIPPED" And f(1) <> "FAILED" Then s = s & "（" & f(3) & " / " & f(2) & " 件取込）"
            If Len(f(5)) > 0 Then s = s & " " & f(5)
            s = s & vbCrLf
        Next f
    End If
    If Not mMessages Is Nothing Then
        For Each m In mMessages
            s = s & "・" & m & vbCrLf
        Next m
    End If
    If Len(mReportPdf) > 0 Then s = s & "レポート：" & Mid$(mReportPdf, InStrRev(mReportPdf, "\") + 1) & vbCrLf
    BuildSummaryText = s
End Function

' 監視ツールなどから機械的に読めるよう、結果を JSON で残す
Private Function BuildResultJson(ByVal exitCode As Long) As String
    Dim s As String
    Dim f As Variant
    Dim m As Variant
    Dim first As Boolean

    s = "{" & vbLf
    s = s & "  ""runId"": " & JsonString(RunId()) & "," & vbLf
    s = s & "  ""status"": " & JsonString(StatusName(exitCode)) & "," & vbLf
    s = s & "  ""exitCode"": " & exitCode & "," & vbLf
    s = s & "  ""startedAt"": " & JsonString(Format$(mStartedAt, "yyyy-mm-dd\Thh:nn:ss")) & "," & vbLf
    s = s & "  ""finishedAt"": " & JsonString(Format$(Now, "yyyy-mm-dd\Thh:nn:ss")) & "," & vbLf
    s = s & "  ""user"": " & JsonString(CurrentUserName()) & "," & vbLf
    s = s & "  ""files"": ["
    first = True
    For Each f In mFiles
        s = s & IIf(first, vbLf, "," & vbLf)
        s = s & "    {""name"": " & JsonString(CStr(f(0))) & ", ""status"": " & JsonString(CStr(f(1))) & _
                ", ""rows"": " & f(2) & ", ""imported"": " & f(3) & ", ""errorRows"": " & f(4) & _
                ", ""message"": " & JsonString(CStr(f(5))) & "}"
        first = False
    Next f
    s = s & IIf(first, "],", vbLf & "  ],") & vbLf
    s = s & "  ""messages"": ["
    first = True
    For Each m In mMessages
        s = s & IIf(first, "", ", ") & JsonString(CStr(m))
        first = False
    Next m
    s = s & "]," & vbLf
    s = s & "  ""reportPdf"": " & JsonString(mReportPdf) & vbLf
    s = s & "}" & vbLf
    BuildResultJson = s
End Function

' JSON の文字列リテラル（" \ 制御文字をエスケープ）
Public Function JsonString(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 10: out = out & "\n"
            Case 13: out = out & "\r"
            Case 9:  out = out & "\t"
            Case Is < 32: out = out & "\u" & Right$("000" & Hex$(code), 4)
            Case Else: out = out & ch
        End Select
    Next i
    JsonString = """" & out & """"
End Function

Private Function ImportedFileCount() As Long
    Dim f As Variant
    If mFiles Is Nothing Then Exit Function
    For Each f In mFiles
        If f(3) > 0 Then ImportedFileCount = ImportedFileCount + 1
    Next f
End Function

'==========================================================
' ファイル操作
'==========================================================
' フォルダ内の *.csv をファイル名の昇順で返す（日付入りのファイル名なら古い順になる）
Private Function ListCsvFiles(ByVal folder As String, ByRef n As Long) As String()
    Dim names() As String
    Dim f As String
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    ReDim names(1 To 1)
    n = 0
    f = Dir$(folder & "\*.csv")
    Do While Len(f) > 0
        n = n + 1
        ReDim Preserve names(1 To n)
        names(n) = f
        f = Dir$
    Loop
    For i = 1 To n - 1
        For j = i + 1 To n
            If names(i) > names(j) Then
                tmp = names(i): names(i) = names(j): names(j) = tmp
            End If
        Next j
    Next i
    ListCsvFiles = names
End Function

' 最終更新から minAgeSec 秒以上経っているか（コピー・書込中のファイルを掴まないため）
Public Function IsFileReady(ByVal filePath As String, ByVal minAgeSec As Long) As Boolean
    IsFileReady = (DateDiff("s", FileDateTime(filePath), Now) >= minAgeSec)
End Function

' 同名のファイルがあれば _2, _3 … を付けた空いている名前を返す
Public Function UniquePath(ByVal folder As String, ByVal fileName As String) As String
    Dim p As String
    Dim i As Long
    p = folder & "\" & fileName
    i = 1
    Do While FileExists(p)
        i = i + 1
        p = folder & "\" & BaseName(fileName) & "_" & i & Extension(fileName)
    Loop
    UniquePath = p
End Function

' ファイルを移動する（名前の先頭に実行IDを付け、どの実行で処理したかを分かるようにする）。戻り値：移動先
Private Function MoveFile(ByVal src As String, ByVal destFolder As String) As String
    Dim dest As String
    EnsureFolder destFolder
    dest = UniquePath(destFolder, RunId() & "_" & Mid$(src, InStrRev(src, "\") + 1))
    Name src As dest
    MoveFile = dest
End Function

' ロックファイルを作る。実行中の別バッチがあれば False（タイムアウトを過ぎた古いロックは取り直す）
Private Function AcquireLock(ByVal lockPath As String) As Boolean
    Const PROC As String = "modBatch.AcquireLock"
    Dim timeoutMin As Long
    Dim fileNo As Integer

    timeoutMin = CLng(GetConfig("BatchLockTimeoutMin", 60))
    If FileExists(lockPath) Then
        If DateDiff("n", FileDateTime(lockPath), Now) < timeoutMin Then
            LogWarn PROC, "ロックファイルがあるためスキップ: " & lockPath
            Exit Function
        End If
        LogWarn PROC, "古いロックファイルを削除します（前回のバッチが異常終了した可能性）: " & lockPath
        Kill lockPath
    End If

    EnsureFolder Left$(lockPath, InStrRev(lockPath, "\") - 1)
    fileNo = FreeFile
    Open lockPath For Output As #fileNo
    Print #fileNo, RunId() & vbTab & CurrentUserName() & vbTab & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #fileNo
    AcquireLock = True
End Function

Private Sub ReleaseLock(ByVal lockPath As String)
    On Error Resume Next
    If FileExists(lockPath) Then Kill lockPath
End Sub


