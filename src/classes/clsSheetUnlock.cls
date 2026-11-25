VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================
' clsSheetUnlock : シート保護の一時解除（clsAppState と同じ RAII パターン）
'   Set unlock = UnlockSheet(ws) で保護を解除し、
'   変数が破棄されたとき（エラーで処理を抜けた場合も含む）に自動で再保護する
'   → 「保護を外したまま」という状態を残さない
'==========================================================

Private mWs As Worksheet

Public Sub Init(ByVal ws As Worksheet)
    Set mWs = ws
    If ws.ProtectContents Then
        On Error GoTo Fail
        ws.Unprotect ProtectionPassword()
    End If
    Exit Sub
Fail:
    Set mWs = Nothing
    RaiseAppError ERR_PERMISSION, "clsSheetUnlock.Init", "シート「" & ws.Name & "」の保護を解除できません。" & vbCrLf & _
                  "管理者メニューの「シート保護を再設定する」を実行してください。"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    ' 解除前の状態にかかわらず保護する（新しく作られたシートも保護される。未設定なら何もしない）
    If Not mWs Is Nothing Then ProtectSheet mWs
End Sub


