Attribute VB_Name = "modProtection"
Option Explicit
'==========================================================
' modProtection : シート保護
'   目的は「手作業による誤った書き換えの防止」（Excel のシート保護は強固なアクセス制御ではない）
'   ・パスワードはソースコードに書かず、初期設定時に乱数で生成して非表示シート _Sys に保持する
'     → 公開リポジトリのソースからパスワードが漏れない
'   ・UserInterfaceOnly:=True で保護し、マクロからのセル書込は許可する
'     ※ この設定はブックを閉じると失われるため、開くたびに再設定する（ThisWorkbook.Workbook_Open）
'   ・テーブルの拡張・グラフの作成など UserInterfaceOnly でも許可されない操作は、
'     clsSheetUnlock で処理の間だけ保護を解除し、終わったら自動で戻す
'==========================================================

Private Const PROTECTED_SHEETS As String = "Config,Schema,Orders,Report,ImportErrors"

' 保護パスワード（未設定なら空文字）
Public Function ProtectionPassword() As String
    Dim ws As Worksheet
    Set ws = GetSheetOrNothing(SYS_SHEET)
    If ws Is Nothing Then Exit Function
    ProtectionPassword = CStr(ws.Range("B1").value)
End Function

'----------------------------------------------------------
' すべての対象シートを保護し、_Users・_Sys を非表示に戻す
'----------------------------------------------------------
Public Sub ApplyProtection()
    Const PROC As String = "modProtection.ApplyProtection"
    Dim pwd As String
    Dim sheetName As Variant
    Dim ws As Worksheet

    pwd = ProtectionPassword()
    If Len(pwd) = 0 Then
        LogWarn PROC, "セキュリティ未設定のため、シートを保護しません"
        Exit Sub
    End If

    For Each sheetName In Split(PROTECTED_SHEETS, ",")
        Set ws = GetSheetOrNothing(CStr(sheetName))
        If Not ws Is Nothing Then ProtectSheet ws, pwd
    Next sheetName

    For Each sheetName In Array(USERS_SHEET, SYS_SHEET)
        Set ws = GetSheetOrNothing(CStr(sheetName))
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVeryHidden
            ProtectSheet ws, pwd
        End If
    Next sheetName
    LogDebug PROC, "シート保護を適用しました"
End Sub

'----------------------------------------------------------
' 1シートを保護する（保護済みでも UserInterfaceOnly を付け直すため Protect を呼ぶ）
'----------------------------------------------------------
Public Sub ProtectSheet(ByVal ws As Worksheet, Optional ByVal pwd As String = "")
    Const PROC As String = "modProtection.ProtectSheet"
    Dim errDesc As String

    If Len(pwd) = 0 Then pwd = ProtectionPassword()
    If Len(pwd) = 0 Then Exit Sub

    On Error Resume Next
    DoProtect ws, pwd
    If Err.Number <> 0 Then                       ' 保護の設定が異なる場合は、いったん解除して掛け直す
        Err.Clear
        ws.Unprotect pwd
        DoProtect ws, pwd
    End If
    errDesc = Err.Description
    On Error GoTo 0

    If Len(errDesc) > 0 Then
        LogError PROC, "シートを保護できません: " & ws.Name & " / " & errDesc
        RaiseAppError ERR_PERMISSION, PROC, "シート「" & ws.Name & "」を保護できません。" & vbCrLf & _
                      "別のパスワードで保護されている可能性があります。"
    End If
End Sub

'----------------------------------------------------------
' 処理の間だけ保護を解除する（戻り値の変数が破棄されると自動で再保護される）
'   使い方：Dim unlock As clsSheetUnlock
'           Set unlock = UnlockSheet(ws)
'----------------------------------------------------------
Public Function UnlockSheet(ByVal ws As Worksheet) As clsSheetUnlock
    Dim u As clsSheetUnlock
    Set u = New clsSheetUnlock
    u.Init ws
    Set UnlockSheet = u
End Function

' シートが無ければ Nothing を返す
Public Function UnlockSheetByName(ByVal sheetName As String) As clsSheetUnlock
    Dim ws As Worksheet
    Set ws = GetSheetOrNothing(sheetName)
    If Not ws Is Nothing Then Set UnlockSheetByName = UnlockSheet(ws)
End Function

Private Sub DoProtect(ByVal ws As Worksheet, ByVal pwd As String)
    ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFormattingColumns:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub


