Attribute VB_Name = "modImportUndo"
Option Explicit
'==========================================================
' modImportUndo : 取込の取消（取込ID単位）
'   誤ったファイルを取り込んだときに、その取込分だけを安全に取り消す
'   ・取込IDごとに管理しているため、他の取込データには影響しない
'   ・取消前に自動でバックアップ、取消はログに記録（監査証跡）
'   ・確認ダイアログの既定ボタンは「いいえ」（誤操作防止）
'==========================================================

Private Const COL_IMPORT_ID As String = "取込ID"
Private Const COL_IMPORT_FILE As String = "取込ファイル"
Private Const COL_IMPORT_AT As String = "取込日時"

'----------------------------------------------------------
' UI：直近の取込を取り消す
'----------------------------------------------------------
Public Sub UndoLastImport()
    Const PROC As String = "modImportUndo.UndoLastImport"
    Dim lo As ListObject
    Dim data As Variant
    Dim idCol As Long
    Dim fileCol As Long
    Dim atCol As Long
    Dim i As Long
    Dim lastId As String
    Dim fileName As String
    Dim importedAt As Variant
    Dim cnt As Long
    Dim removed As Long
    On Error GoTo EH

    ResetRunId
    ReloadConfig
    RequirePermission actUndoImport       ' 確認ダイアログを出す前に権限を確認する

    Set lo = GetOrdersTable()
    If lo.DataBodyRange Is Nothing Then
        RaiseAppError ERR_VALIDATION, PROC, "取り消せる取込データがありません。"
    End If

    data = lo.DataBodyRange.value
    idCol = lo.ListColumns(COL_IMPORT_ID).Index
    fileCol = lo.ListColumns(COL_IMPORT_FILE).Index
    atCol = lo.ListColumns(COL_IMPORT_AT).Index

    ' 取込IDは「yyyymmdd-hhnnss-乱数」で時系列順に並ぶため、最大値 = 最新
    For i = 1 To UBound(data, 1)
        If CStr(data(i, idCol)) > lastId Then lastId = CStr(data(i, idCol))
    Next i
    If Len(lastId) = 0 Then RaiseAppError ERR_VALIDATION, PROC, "取り消せる取込データがありません。"

    For i = 1 To UBound(data, 1)
        If CStr(data(i, idCol)) = lastId Then
            cnt = cnt + 1
            fileName = CStr(data(i, fileCol))
            importedAt = data(i, atCol)
        End If
    Next i

    If MsgBox("次の取込を取り消します。よろしいですか？" & vbCrLf & vbCrLf & _
              "取込ID　：" & lastId & vbCrLf & _
              "取込日時：" & Format$(importedAt, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
              "ファイル：" & fileName & vbCrLf & _
              "件数　　：" & Format$(cnt, "#,##0") & " 件" & vbCrLf & vbCrLf & _
              "※ 実行前にバックアップを作成します。", _
              vbQuestion + vbYesNo + vbDefaultButton2, AppTitle()) <> vbYes Then
        LogInfo PROC, "取消がキャンセルされました: 取込ID=" & lastId
        Exit Sub
    End If

    removed = RemoveImportBatch(lastId, True)
    MsgBox Format$(removed, "#,##0") & " 件の取込を取り消しました。", vbInformation, AppTitle()
    Exit Sub

EH:
    HandleError PROC
End Sub

'----------------------------------------------------------
' 処理層：指定した取込IDの行を削除する。戻り値：削除件数
'   行を1行ずつ Delete すると10万行では非常に遅いため、
'   「残す行」を配列に詰めて一括で書き戻す
'----------------------------------------------------------
Public Function RemoveImportBatch(ByVal importId As String, Optional ByVal withBackup As Boolean = True) As Long
    Const PROC As String = "modImportUndo.RemoveImportBatch"
    Dim guard As clsAppState
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim data As Variant
    Dim keep() As Variant
    Dim idCol As Long
    Dim n As Long
    Dim nc As Long
    Dim i As Long
    Dim c As Long
    Dim k As Long
    Dim firstRow As Long
    Dim firstCol As Long
    Dim sheetGuard As clsSheetUnlock

    RequirePermission actUndoImport       ' 処理側でも必ず確認する（多層防御）
    Set guard = New clsAppState
    Set lo = GetOrdersTable()
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set ws = lo.parent
    Set sheetGuard = UnlockSheet(ws)          ' テーブルの縮小は保護中にできないため、処理の間だけ解除

    guard.ShowProgress "取消対象を確認しています..."
    data = lo.DataBodyRange.value
    idCol = lo.ListColumns(COL_IMPORT_ID).Index
    n = UBound(data, 1)
    nc = UBound(data, 2)

    ReDim keep(1 To n, 1 To nc)
    For i = 1 To n
        If CStr(data(i, idCol)) <> importId Then
            k = k + 1
            For c = 1 To nc
                keep(k, c) = data(i, c)
            Next c
        End If
    Next i

    RemoveImportBatch = n - k
    If RemoveImportBatch = 0 Then
        LogWarn PROC, "取消対象がありません: 取込ID=" & importId
        Exit Function
    End If

    If withBackup Then
        guard.ShowProgress "バックアップを作成しています..."
        CreateBackup
    End If

    guard.ShowProgress "取り消しています..."
    firstRow = lo.DataBodyRange.Row
    firstCol = lo.Range.Column
    lo.DataBodyRange.ClearContents
    If k > 0 Then
        ' 読み戻した文字列はそのまま書くと "=HYPERLINK(...)" が数式として復活してしまうため、
        ' 文字列項目の列を @ 書式にしてから書き戻す
        ApplyTextFormat lo, firstRow, k
        ws.Cells(firstRow, firstCol).Resize(k, nc).value = keep
    End If

    ' テーブルは最低1行のデータ行が必要なため、全件削除時は空行1行を残す
    If k = 0 Then k = 1
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1, 1), ws.Cells(firstRow + k - 1, firstCol + nc - 1))

    ' テーブルの外に残った行（書式だけが残っている行）を削除する
    ' （残すと使用範囲が10万行のまま記録され、ブックのサイズが何MBも増える）
    If n > k Then ws.Range(ws.rows(firstRow + k), ws.rows(firstRow + n - 1)).Delete

    LogInfo PROC, "取込を取り消しました: 取込ID=" & importId & " / " & RemoveImportBatch & " 件 / 実行者=" & CurrentUserName()
    AuditLog "UNDO_IMPORT", "OK", "importId=" & importId & " / rows=" & RemoveImportBatch
End Function

