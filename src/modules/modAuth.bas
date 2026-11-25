Attribute VB_Name = "modAuth"
Option Explicit
'==========================================================
' modAuth : ユーザー認可（ロールベースのアクセス制御）
'   ・ユーザーは Windows のログオン名で識別する
'       環境変数 USERNAME は Excel の起動前に書き換えられるため使わず、Windows API で取得する
'   ・ロール：ADMIN（管理者）> OPERATOR（担当者）> VIEWER（閲覧者）
'   ・未登録ユーザーは何もできない（既定で拒否 = フェイルセーフ）
'   ・権限チェックは画面（ボタンの無効化）と処理（RequirePermission）の両方で行う（多層防御）
'   ・ユーザー一覧は非表示シート _Users のテーブル tblUsers
'       UserName に * を登録すると、未登録の全員にそのロールを与える（デモ・全員閲覧用）
'       ただし * に ADMIN は与えられない（OPERATOR に制限）
'==========================================================

Private Declare PtrSafe Function GetUserNameW Lib "advapi32.dll" ( _
    ByVal lpBuffer As LongPtr, ByRef nSize As Long) As Long

Public Const USERS_SHEET As String = "_Users"
Public Const SYS_SHEET As String = "_Sys"
Private Const USERS_TABLE As String = "tblUsers"

Public Enum UserRole
    roleNone = 0
    roleViewer = 1
    roleOperator = 2
    roleAdmin = 3
End Enum

Public Enum AppAction
    actReport = 1          ' 帳票の作成・出力       … VIEWER 以上
    actImport = 2          ' CSV取込                … OPERATOR 以上
    actUndoImport = 3      ' 取込の取消             … ADMIN
    actAdmin = 4           ' 設定編集・保護・ユーザー管理・監査ログ検証・性能計測 … ADMIN
End Enum

'----------------------------------------------------------
' Windows のログオン名（API で取得）
'----------------------------------------------------------
Public Function CurrentUserName() As String
    Static cached As String
    Dim buf As String
    Dim size As Long

    If Len(cached) = 0 Then
        size = 257
        buf = String$(size, vbNullChar)
        If GetUserNameW(StrPtr(buf), size) <> 0 Then
            cached = Left$(buf, size - 1)           ' size は終端の NULL を含む文字数
        Else
            cached = Environ$("USERNAME")
        End If
    End If
    CurrentUserName = cached
End Function

Public Function IsSecuritySetUp() As Boolean
    IsSecuritySetUp = Not (GetSheetOrNothing(USERS_SHEET) Is Nothing)
End Function

Public Function CurrentRole() As UserRole
    CurrentRole = RoleOf(CurrentUserName())
End Function

' 個別の登録を優先し、なければ * の登録を使う。どちらもなければ権限なし
Public Function RoleOf(ByVal loginName As String) As UserRole
    Dim data As Variant
    Dim i As Long
    Dim wildcard As UserRole

    data = UsersData()
    If IsEmpty(data) Then Exit Function

    For i = 1 To UBound(data, 1)
        If IsEnabled(data(i, 4)) Then
            If StrComp(Trim$(CStr(data(i, 1))), loginName, vbTextCompare) = 0 Then
                RoleOf = RoleFromName(CStr(data(i, 3)))
                Exit Function
            End If
            If Trim$(CStr(data(i, 1))) = "*" Then
                wildcard = RoleFromName(CStr(data(i, 3)))
                If wildcard > roleOperator Then wildcard = roleOperator
            End If
        End If
    Next i
    RoleOf = wildcard
End Function

' 有効な ADMIN（* 以外）が1人以上いるか（全員を外して締め出される事故を防ぐ）
Public Function HasEnabledAdmin() As Boolean
    Dim data As Variant
    Dim i As Long

    data = UsersData()
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If IsEnabled(data(i, 4)) And Trim$(CStr(data(i, 1))) <> "*" Then
            If RoleFromName(CStr(data(i, 3))) = roleAdmin Then
                HasEnabledAdmin = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Function RoleFromName(ByVal roleText As String) As UserRole
    Select Case UCase$(Trim$(roleText))
        Case "ADMIN":    RoleFromName = roleAdmin
        Case "OPERATOR": RoleFromName = roleOperator
        Case "VIEWER":   RoleFromName = roleViewer
        Case Else:       RoleFromName = roleNone
    End Select
End Function

Public Function RoleName(ByVal role As UserRole) As String
    Select Case role
        Case roleAdmin:    RoleName = "管理者"
        Case roleOperator: RoleName = "担当者"
        Case roleViewer:   RoleName = "閲覧者"
        Case Else:         RoleName = "権限なし"
    End Select
End Function

Public Function ActionName(ByVal action As AppAction) As String
    Select Case action
        Case actReport:     ActionName = "帳票の作成"
        Case actImport:     ActionName = "CSV取込"
        Case actUndoImport: ActionName = "取込の取消"
        Case actAdmin:      ActionName = "管理者機能"
    End Select
End Function

' 権限の対応表（画面とテストからも使うため、ロールを引数で受け取る純粋な関数にしている）
Public Function HasRoleFor(ByVal role As UserRole, ByVal action As AppAction) As Boolean
    Dim required As UserRole
    Select Case action
        Case actReport:     required = roleViewer
        Case actImport:     required = roleOperator
        Case actUndoImport: required = roleAdmin
        Case actAdmin:      required = roleAdmin
        Case Else:          required = roleAdmin + 1     ' 未定義の操作は誰にも許可しない
    End Select
    HasRoleFor = (role >= required)
End Function

Public Function HasPermission(ByVal action As AppAction) As Boolean
    HasPermission = HasRoleFor(CurrentRole(), action)
End Function

'----------------------------------------------------------
' 処理の入口で呼ぶ。権限がなければ監査ログに記録して業務エラーにする
'----------------------------------------------------------
Public Sub RequirePermission(ByVal action As AppAction)
    Const PROC As String = "modAuth.RequirePermission"
    Dim role As UserRole

    If Not IsSecuritySetUp() Then
        RaiseAppError ERR_PERMISSION, PROC, "セキュリティの初期設定がされていません。" & vbCrLf & _
                      "管理者が SetupSecurity を実行してください。"
    End If

    role = CurrentRole()
    If HasRoleFor(role, action) Then Exit Sub

    AuditLog "ACCESS_DENIED", "NG", ActionName(action) & "（ロール：" & RoleName(role) & "）"
    If role = roleNone Then
        RaiseAppError ERR_PERMISSION, PROC, "ユーザー登録がありません（" & CurrentUserName() & "）。" & vbCrLf & _
                      "管理者にユーザー登録を依頼してください。"
    Else
        RaiseAppError ERR_PERMISSION, PROC, "「" & ActionName(action) & "」を行う権限がありません。" & vbCrLf & _
                      "（現在のロール：" & RoleName(role) & "）"
    End If
End Sub

'----------------------------------------------------------
' 初期設定：最初に実行した人を ADMIN として登録し、保護パスワードを生成する
'----------------------------------------------------------
Public Sub SetupSecurity()
    Const PROC As String = "modAuth.SetupSecurity"
    Dim wsUsers As Worksheet
    Dim wsSys As Worksheet
    Dim lo As ListObject
    Dim lastWs As Worksheet
    On Error GoTo EH

    ResetRunId
    ReloadConfig

    If IsSecuritySetUp() Then
        MsgBox "セキュリティの初期設定は完了しています。" & vbCrLf & _
               "ユーザーの追加・変更は、メニューの「ユーザーを管理する」から行ってください。", vbInformation, AppTitle()
        Exit Sub
    End If

    If MsgBox("セキュリティの初期設定を行います。" & vbCrLf & vbCrLf & _
              "・現在のユーザー（" & CurrentUserName() & "）を管理者として登録" & vbCrLf & _
              "・シート保護用のパスワードを乱数で生成（ソースコードには含めない）" & vbCrLf & _
              "・各シートを保護" & vbCrLf & vbCrLf & "実行しますか？", _
              vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Sub

    Set lastWs = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)

    ' ユーザー一覧
    Set wsUsers = ThisWorkbook.Worksheets.Add(After:=lastWs)
    wsUsers.Name = USERS_SHEET
    wsUsers.Columns("A:B").NumberFormat = "@"
    wsUsers.Range("A1:E1").value = Array("UserName", "DisplayName", "Role", "Enabled", "Note")
    wsUsers.Range("A2:E2").value = Array(CurrentUserName(), CurrentUserName(), "ADMIN", True, "初期設定で登録")
    Set lo = wsUsers.ListObjects.Add(xlSrcRange, wsUsers.Range("A1:E2"), , xlYes)
    lo.Name = USERS_TABLE
    wsUsers.Columns("A:E").AutoFit

    ' システム情報（保護パスワード）
    Set wsSys = ThisWorkbook.Worksheets.Add(After:=wsUsers)
    wsSys.Name = SYS_SHEET
    wsSys.Columns("B").NumberFormat = "@"
    wsSys.Range("A1").value = "ProtectionPassword"
    wsSys.Range("B1").value = RandomToken(24)               ' 暗号論的に安全な乱数で生成
    wsSys.Range("A2").value = "SetupAt"
    wsSys.Range("B2").value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    wsSys.Range("A3").value = "SetupBy"
    wsSys.Range("B3").value = CurrentUserName()

    wsUsers.Visible = xlSheetVeryHidden          ' VBA からしか表示できない
    wsSys.Visible = xlSheetVeryHidden

    ApplyProtection
    AuditLog "SECURITY_SETUP", "OK", "admin=" & CurrentUserName()
    MsgBox "セキュリティの初期設定が完了しました。" & vbCrLf & vbCrLf & _
           "管理者：" & CurrentUserName() & vbCrLf & _
           "ユーザーの追加は、メニューの「ユーザーを管理する」から行えます。", vbInformation, AppTitle()
    Exit Sub
EH:
    HandleError PROC
End Sub

'==========================================================
' 内部処理
'==========================================================
' tblUsers のデータ部（2次元配列）。シート・データが無ければ Empty
Private Function UsersData() As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = GetSheetOrNothing(USERS_SHEET)
    If ws Is Nothing Then Exit Function
    If ws.ListObjects.Count = 0 Then Exit Function
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function
    UsersData = lo.DataBodyRange.value          ' 5列あるので1行でも2次元配列になる
End Function

Private Function IsEnabled(ByVal v As Variant) As Boolean
    Select Case UCase$(Trim$(CStr(v)))
        Case "TRUE", "1", "Y", "YES", "○"
            IsEnabled = True
    End Select
End Function

Public Function GetSheetOrNothing(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetOrNothing = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function


