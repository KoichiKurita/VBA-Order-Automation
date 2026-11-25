VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenu 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMenu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==========================================================
' frmMenu : メインメニュー（モードレス表示）
'   ・コントロールはコードで生成する
'     （レイアウトがバイナリの .frx ではなくコードに残るため、Git で差分を確認できる）
'   ・ボタンの有効 / 無効はロールで切り替える
'     （処理側でも RequirePermission で必ず確認する＝多層防御。ボタンの無効化だけには頼らない）
'   ・処理中はすべてのボタンを無効化し、二重実行を防ぐ
'==========================================================

Private WithEvents btnImport As MSForms.CommandButton
Attribute btnImport.VB_VarHelpID = -1
Private WithEvents btnReport As MSForms.CommandButton
Attribute btnReport.VB_VarHelpID = -1
Private WithEvents btnUndo As MSForms.CommandButton
Attribute btnUndo.VB_VarHelpID = -1
Private WithEvents btnEditSettings As MSForms.CommandButton
Attribute btnEditSettings.VB_VarHelpID = -1
Private WithEvents btnUsers As MSForms.CommandButton
Attribute btnUsers.VB_VarHelpID = -1
Private WithEvents btnProtect As MSForms.CommandButton
Attribute btnProtect.VB_VarHelpID = -1
Private WithEvents btnVerify As MSForms.CommandButton
Attribute btnVerify.VB_VarHelpID = -1
Private WithEvents btnOpenLog As MSForms.CommandButton
Attribute btnOpenLog.VB_VarHelpID = -1
Private WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1
Private lblUser As MSForms.label
Private lblStatus As MSForms.label

Private mBusy As Boolean

Private Const UI_FONT As String = "Meiryo UI"
Private Const MARGIN As Single = 15
Private Const BTN_W As Single = 290
Private Const BTN_H As Single = 26
Private Const GAP As Single = 6

'----------------------------------------------------------
' 画面の組み立て
'----------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim y As Single

    Me.Caption = AppTitle()
    Me.Width = BTN_W + MARGIN * 2 + (Me.Width - Me.InsideWidth)

    y = 10
    AddLabel "lblTitle", AppTitle(), y, 14, True, 24
    y = y + 26
    Set lblUser = AddLabel("lblUser", "", y, 9, False, 16)
    y = y + 24

    AddLabel "lblSecBiz", "業務", y, 9, True, 16
    y = y + 18
    Set btnImport = AddButton("btnImport", "受注CSVを取り込む", y): y = y + BTN_H + GAP
    Set btnReport = AddButton("btnReport", "月次レポートを作成する（PDF / CSV）", y): y = y + BTN_H + GAP + 8

    AddLabel "lblSecAdmin", "管理者", y, 9, True, 16
    y = y + 18
    Set btnUndo = AddButton("btnUndo", "直近の取込を取り消す", y): y = y + BTN_H + GAP
    Set btnEditSettings = AddButton("btnEditSettings", "設定（Config・Schema）を編集する", y): y = y + BTN_H + GAP
    Set btnUsers = AddButton("btnUsers", "ユーザーを管理する", y): y = y + BTN_H + GAP
    Set btnProtect = AddButton("btnProtect", "シート保護を再設定する", y): y = y + BTN_H + GAP
    Set btnVerify = AddButton("btnVerify", "監査ログを検証する", y): y = y + BTN_H + GAP
    Set btnOpenLog = AddButton("btnOpenLog", "ログフォルダを開く", y): y = y + BTN_H + GAP + 8

    Set btnClose = AddButton("btnClose", "閉じる", y): y = y + BTN_H + GAP
    Set lblStatus = AddLabel("lblStatus", "準備完了", y, 9, False, 30)
    lblStatus.ForeColor = RGB(89, 89, 89)
    lblStatus.WordWrap = True
    y = y + 34

    Me.Height = y + (Me.Height - Me.InsideHeight)
    RefreshState
End Sub

' 処理中は閉じられないようにする
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mBusy Then Cancel = True
End Sub

'----------------------------------------------------------
' ボタン
'----------------------------------------------------------
Private Sub btnImport_Click()
    RunAction "ImportOrdersCsv", "CSV取込"
End Sub
Private Sub btnReport_Click()
    RunAction "CreateMonthlyReport", "レポート作成"
End Sub
Private Sub btnUndo_Click()
    RunAction "UndoLastImport", "取込の取消"
End Sub
Private Sub btnEditSettings_Click()
    RunAction "AdminEditSettings", "設定の編集"
End Sub
Private Sub btnUsers_Click()
    RunAction "AdminManageUsers", "ユーザー管理"
End Sub
Private Sub btnProtect_Click()
    RunAction "AdminProtectAll", "シート保護"
End Sub
Private Sub btnVerify_Click()
    RunAction "VerifyAuditLog", "監査ログの検証"
End Sub
Private Sub btnOpenLog_Click()
    RunAction "AdminOpenLogFolder", "ログフォルダ"
End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub

'----------------------------------------------------------
' 処理側から進捗を受け取る（modUI.ReportProgress 経由）
'----------------------------------------------------------
Public Sub SetStatus(ByVal message As String)
    lblStatus.Caption = message
    Me.Repaint
End Sub

'==========================================================
' 内部処理
'==========================================================
Private Sub RunAction(ByVal procName As String, ByVal actionLabel As String)
    If mBusy Then Exit Sub
    mBusy = True
    RefreshState
    SetStatus actionLabel & "を実行しています..."
    SetProgressSink Me

    ' 各処理は自身でエラー処理（メッセージ表示・ログ）を行う
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!" & procName
    On Error GoTo 0

    ClearProgressSink
    mBusy = False
    RefreshState
    SetStatus actionLabel & "：終了（" & Format$(Now, "hh:nn:ss") & "）"
End Sub

' ロールに応じてボタンの有効 / 無効を切り替える
Private Sub RefreshState()
    Dim role As UserRole
    Dim isAdmin As Boolean

    role = CurrentRole()
    isAdmin = HasRoleFor(role, actAdmin) And Not mBusy
    lblUser.Caption = "ユーザー：" & CurrentUserName() & "　　ロール：" & RoleName(role)

    btnImport.enabled = HasRoleFor(role, actImport) And Not mBusy
    btnReport.enabled = HasRoleFor(role, actReport) And Not mBusy
    btnUndo.enabled = HasRoleFor(role, actUndoImport) And Not mBusy
    btnEditSettings.enabled = isAdmin
    btnUsers.enabled = isAdmin
    btnProtect.enabled = isAdmin
    btnVerify.enabled = isAdmin
    btnOpenLog.enabled = isAdmin
    btnClose.enabled = Not mBusy

    If Not IsSecuritySetUp() Then
        lblStatus.Caption = "セキュリティが未設定です。管理者が SetupSecurity を実行してください。"
    ElseIf role = roleNone Then
        lblStatus.Caption = "ユーザー登録がありません。管理者に登録を依頼してください。"
    End If
End Sub

Private Function AddButton(ByVal ctlName As String, ByVal captionText As String, ByVal topPos As Single) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    Set b = Me.Controls.Add("Forms.CommandButton.1", ctlName)
    With b
        .Caption = captionText
        .Left = MARGIN
        .Top = topPos
        .Width = BTN_W
        .Height = BTN_H
        .Font.Name = UI_FONT
        .Font.size = 10
        .TakeFocusOnClick = False
    End With
    Set AddButton = b
End Function

Private Function AddLabel(ByVal ctlName As String, ByVal captionText As String, ByVal topPos As Single, _
                          ByVal fontSize As Single, ByVal isBold As Boolean, ByVal h As Single) As MSForms.label
    Dim l As MSForms.label
    Set l = Me.Controls.Add("Forms.Label.1", ctlName)
    With l
        .Caption = captionText
        .Left = MARGIN
        .Top = topPos
        .Width = BTN_W
        .Height = h
        .Font.Name = UI_FONT
        .Font.size = fontSize
        .Font.Bold = isBold
    End With
    Set AddLabel = l
End Function


