'==========================================================
' run_batch.vbs : 受注取込バッチの起動スクリプト（タスクスケジューラから実行）
'   使い方：cscript //nologo run_batch.vbs [OrderTool.xlsm のパス]
'           パスを省略すると、このスクリプトの1つ上のフォルダにある OrderTool.xlsm を使う
'   終了コード：0 = 正常 / 1 = 警告 / 2 = 異常終了 / 3 = スキップ（RunBatch の戻り値をそのまま返す）
'
'   ・見えない Excel を起動し、画面を出さずに RunBatch を実行して、必ず Excel を終了する
'     （途中で失敗しても Excel のプロセスを残さない）
'   ・EnableEvents = False で開き、Workbook_Open（メニュー表示）を動かさない
'   ・AutomationSecurity = ByUI：信頼できる場所にあるブックでなければマクロは動かない
'   ・Excel を起動できないなど、Excel の外で起きた失敗は logs\launcher.log に記録する
'==========================================================
Option Explicit

Const MSO_AUTOMATION_SECURITY_BY_UI = 2

Dim fso, bookPath, logPath, xl, wb, exitCode, errMsg

Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count > 0 Then
    bookPath = WScript.Arguments(0)
Else
    bookPath = fso.BuildPath(fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName)), "OrderTool.xlsm")
End If
logPath = fso.BuildPath(fso.GetParentFolderName(bookPath), "logs\launcher.log")

exitCode = 2
On Error Resume Next

If Not fso.FileExists(bookPath) Then
    WriteLog "ERROR", "ブックが見つかりません: " & bookPath
    WScript.Quit 2
End If

Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WriteLog "ERROR", "Excel を起動できません: " & Err.Description
    WScript.Quit 2
End If

xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False
xl.AutomationSecurity = MSO_AUTOMATION_SECURITY_BY_UI

Set wb = xl.Workbooks.Open(bookPath, 0, False)
If Err.Number <> 0 Then
    errMsg = "ブックを開けません: " & Err.Description
    Err.Clear
Else
    exitCode = CLng(xl.Run("'" & wb.Name & "'!RunBatch"))
    If Err.Number <> 0 Then
        errMsg = "RunBatch の実行に失敗しました（マクロが無効の可能性）: " & Err.Description
        exitCode = 2
        Err.Clear
    End If
    wb.Close False                       ' 保存は RunBatch の中で行っている
End If

xl.Quit
Set wb = Nothing
Set xl = Nothing

If Len(errMsg) > 0 Then
    WriteLog "ERROR", errMsg
Else
    WriteLog "INFO", "終了コード " & exitCode
End If
WScript.Quit exitCode

Sub WriteLog(level, message)
    Dim ts
    On Error Resume Next
    If Not fso.FolderExists(fso.GetParentFolderName(logPath)) Then fso.CreateFolder fso.GetParentFolderName(logPath)
    Set ts = fso.OpenTextFile(logPath, 8, True)       ' 8 = 追記
    ts.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
                 FormatDateTime(Now, 4) & ":" & Right("0" & Second(Now), 2) & vbTab & level & vbTab & message
    ts.Close
    WScript.Echo level & ": " & message
End Sub
