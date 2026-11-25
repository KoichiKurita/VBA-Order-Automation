Attribute VB_Name = "modRelease"
Option Explicit
'==========================================================
' modRelease : 配布用ブックの作成
'   開発用のブックには、管理者のユーザー名（_Users）と保護パスワード（_Sys）、
'   テストで取り込んだデータが入っている。配布前にこれらを取り除いたコピーを作る
'
'   ・開発用のブック自体は変更しない（コピーに対して作業する）
'   ・配布用ブックは「セキュリティ未設定」の状態になり、最初に開いた人が
'     初期設定（SetupSecurity）で管理者になる（初回セットアップ方式）
'   ・作成者などの個人情報をファイルのプロパティから削除する
'   出力先：dist\OrderTool.xlsm（dist フォルダは Git の管理対象外）
'==========================================================

Public Sub PrepareRelease()
    Const PROC As String = "modRelease.PrepareRelease"
    Dim distDir As String
    Dim tmpPath As String
    Dim finalPath As String
    Dim pwd As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As Variant
    Dim removed As String
    Dim prevEvents As Boolean
    Dim prevAlerts As Boolean
    On Error GoTo EH

    ResetRunId
    ReloadConfig
    RequirePermission actAdmin

    If MsgBox("配布用のブックを作成します（v" & APP_VERSION & "）。" & vbCrLf & vbCrLf & _
              "・ユーザー一覧（_Users）と保護パスワード（_Sys）を削除" & vbCrLf & _
              "・取り込んだデータ（Orders・ImportErrors・Report）を削除" & vbCrLf & _
              "・グラフシートなど、設計にないシートを削除" & vbCrLf & _
              "・ファイルのプロパティから作成者などの個人情報を削除" & vbCrLf & vbCrLf & _
              "このブック（開発用）は変更しません。実行しますか？", _
              vbQuestion + vbYesNo + vbDefaultButton2, AppTitle()) <> vbYes Then Exit Sub

    distDir = ThisWorkbook.path & "\dist"
    EnsureFolder distDir
    tmpPath = distDir & "\_release_tmp" & Extension(ThisWorkbook.Name)   ' 同じ名前のブックは同時に開けないため、一時的な名前で作る
    finalPath = distDir & "\" & ThisWorkbook.Name
    If FileExists(tmpPath) Then Kill tmpPath
    If FileExists(finalPath) Then Kill finalPath

    pwd = ProtectionPassword()
    ThisWorkbook.SaveCopyAs tmpPath

    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    Application.EnableEvents = False          ' コピーの Workbook_Open（メニュー表示など）を動かさない
    Application.DisplayAlerts = False         ' シート削除の確認を出さない
    Set wb = Workbooks.Open(tmpPath)

    ' 保護パスワードを消す前に、すべての保護を外しておく（消した後では外せなくなる）
    For Each ws In wb.Worksheets
        If ws.ProtectContents Then ws.Unprotect pwd
    Next ws

    ' 秘密情報・テストデータのシートを削除（データのシートは、初回の取込・帳票作成で自動的に作られる）
    For Each sheetName In Array(USERS_SHEET, SYS_SHEET, "Orders", "ImportErrors", "Report")
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(sheetName))
        On Error GoTo EH
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
            ws.Delete
            removed = removed & sheetName & " "
        End If
    Next sheetName

    ' 設計にないグラフシート（F11 キーで誤って作られたもの等）を削除
    Do While wb.Charts.Count > 0
        removed = removed & wb.Charts(1).Name & " "
        wb.Charts(1).Delete
    Loop

    ' 個人情報（作成者・最終更新者など）を保存時に削除する
    wb.RemovePersonalInformation = True
    wb.Worksheets("Config").Activate
    wb.Worksheets("Config").Range("A1").Select

    wb.Save
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts

    Name tmpPath As finalPath
    LogInfo PROC, "配布用ブックを作成: " & finalPath & " / 削除: " & removed
    AuditLog "RELEASE", "OK", "v" & APP_VERSION & " / " & finalPath

    MsgBox "配布用のブックを作成しました。" & vbCrLf & finalPath & vbCrLf & vbCrLf & _
           "削除したシート：" & removed & vbCrLf & vbCrLf & _
           "続けて、配布用のブックを開いて次を行ってください。" & vbCrLf & _
           "1. VBAプロジェクトをロック（ツール → VBAProject のプロパティ → 保護）" & vbCrLf & _
           "2. デジタル署名（ツール → デジタル署名）" & vbCrLf & _
           "3. 保存して閉じる", vbInformation, AppTitle()
    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.source
    On Error Resume Next                        ' 後片付け（コピーを閉じ、Excel の設定を戻す）
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0
    Err.Number = errNum                         ' 退避したエラー情報を戻してから共通のエラー処理へ
    Err.Description = errDesc
    Err.source = errSrc
    HandleError PROC
End Sub
