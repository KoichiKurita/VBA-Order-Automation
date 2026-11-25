Attribute VB_Name = "modAdmin"
Option Explicit
'==========================================================
' modAdmin : 管理者機能（すべて ADMIN のみ実行可。操作は監査ログに記録）
'==========================================================

'----------------------------------------------------------
' 設定（Config・Schema）を編集できるように保護を解除する
'----------------------------------------------------------
Public Sub AdminEditSettings()
    Const PROC As String = "modAdmin.AdminEditSettings"
    Dim sheetName As Variant
    Dim ws As Worksheet
    On Error GoTo EH

    ResetRunId
    RequirePermission actAdmin

    For Each sheetName In Array("Config", "Schema")
        Set ws = GetSheetOrNothing(CStr(sheetName))
        If Not ws Is Nothing Then ws.Unprotect ProtectionPassword()
    Next sheetName
    ThisWorkbook.Worksheets("Config").Activate

    AuditLog "UNPROTECT", "OK", "Config, Schema"
    MsgBox "Config・Schema シートの保護を解除しました。" & vbCrLf & vbCrLf & _
           "編集が終わったら、メニューの「シート保護を再設定する」を押してください。" & vbCrLf & _
           "（ブックの保存時にも自動で保護されます）", vbInformation, AppTitle()
    Exit Sub
EH:
    HandleError PROC
End Sub

'----------------------------------------------------------
' ユーザー一覧（_Users）を表示して編集できるようにする
'----------------------------------------------------------
Public Sub AdminManageUsers()
    Const PROC As String = "modAdmin.AdminManageUsers"
    Dim ws As Worksheet
    On Error GoTo EH

    ResetRunId
    RequirePermission actAdmin

    Set ws = GetSheetOrNothing(USERS_SHEET)
    ws.Unprotect ProtectionPassword()
    ws.Visible = xlSheetVisible
    ws.Activate

    AuditLog "USER_EDIT", "OK", "ユーザー一覧の編集を開始"
    MsgBox "ユーザー一覧を表示しました。" & vbCrLf & vbCrLf & _
           "・Role：ADMIN（管理者）/ OPERATOR（担当者）/ VIEWER（閲覧者）" & vbCrLf & _
           "・Enabled を FALSE にすると、そのユーザーは利用できなくなります" & vbCrLf & _
           "・UserName に * を登録すると、未登録の全員に適用されます（ADMIN は不可）" & vbCrLf & vbCrLf & _
           "編集が終わったら、メニューの「シート保護を再設定する」を押してください。", vbInformation, AppTitle()
    Exit Sub
EH:
    HandleError PROC
End Sub

'----------------------------------------------------------
' すべてのシートを保護し直す（_Users は非表示に戻る）
'----------------------------------------------------------
Public Sub AdminProtectAll()
    Const PROC As String = "modAdmin.AdminProtectAll"
    On Error GoTo EH

    ResetRunId
    RequirePermission actAdmin

    ' 有効な管理者がいない状態で非表示に戻すと、誰も管理できなくなるため止める
    If Not HasEnabledAdmin() Then
        MsgBox "有効な管理者（Role = ADMIN、Enabled = TRUE）が1人もいません。" & vbCrLf & _
               "_Users シートを修正してから、もう一度実行してください。", vbCritical, AppTitle()
        Exit Sub
    End If

    ApplyProtection
    ReloadConfig
    AuditLog "PROTECT", "OK", "全シート"
    MsgBox "すべてのシートを保護しました。", vbInformation, AppTitle()
    Exit Sub
EH:
    HandleError PROC
End Sub

'----------------------------------------------------------
' ログフォルダを開く（ログには操作履歴が含まれるため管理者のみ）
'----------------------------------------------------------
Public Sub AdminOpenLogFolder()
    Const PROC As String = "modAdmin.AdminOpenLogFolder"
    On Error GoTo EH

    ResetRunId
    RequirePermission actAdmin
    Shell "explorer.exe """ & ResolvePath("LogDir", "logs") & """", vbNormalFocus
    Exit Sub
EH:
    HandleError PROC
End Sub


