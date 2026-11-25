VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'==========================================================
' ThisWorkbook : ブックのイベント
'==========================================================

Private Sub Workbook_Open()
    On Error Resume Next            ' 起動時の処理の失敗で、ブックを開けなくならないようにする
    ApplyProtection                 ' UserInterfaceOnly は保存されないため、開くたびに再設定
    AuditLog "OPEN", "OK", ThisWorkbook.Name & "（ロール：" & RoleName(CurrentRole()) & "）"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If Not IsSecuritySetUp() Then Exit Sub

    ' 有効な管理者がいない状態で保存すると、誰も管理できなくなるため止める
    If Not HasEnabledAdmin() Then
        MsgBox "有効な管理者（ADMIN）が1人もいないため、保存できません。" & vbCrLf & _
               "_Users シートを確認してください。", vbCritical, AppTitle()
        Cancel = True
        Exit Sub
    End If
    ApplyProtection                 ' 保護を外したまま保存されないようにする
End Sub


