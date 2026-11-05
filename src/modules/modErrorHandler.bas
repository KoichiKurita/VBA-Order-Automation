Attribute VB_Name = "modErrorHandler"
Option Explicit
'==========================================================
' modErrorHandler : 共通エラー処理
'   ・業務エラー（想定内）  → 内容をそのままユーザーに表示
'   ・システムエラー（想定外）→ 汎用メッセージのみ表示し、詳細はログへ
'     （内部情報をユーザー画面に出さない＝情報漏えい対策）
'   ・どちらも「問い合わせ番号（実行ID）」を表示し、ログと突合できる
'==========================================================

' 業務エラーコード範囲
Public Const APP_ERR_BASE As Long = vbObjectError + 1000
Public Const APP_ERR_MAX As Long = vbObjectError + 8999

' 業務エラーコード定義
Public Const ERR_CONFIG_NOT_FOUND As Long = vbObjectError + 1001
Public Const ERR_FILE_NOT_FOUND As Long = vbObjectError + 1101
Public Const ERR_VALIDATION As Long = vbObjectError + 1201
Public Const ERR_PERMISSION As Long = vbObjectError + 1301

Public Sub RaiseAppError(ByVal code As Long, ByVal source As String, ByVal message As String)
    Err.Raise code, source, message
End Sub

' 各プロシージャのエラーハンドラから呼ぶ
Public Sub HandleError(ByVal source As String, Optional ByVal notifyUser As Boolean = True)
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim userMsg As String

    ' 他プロシージャ呼び出しでErrがクリアされる前に退避
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.source

    LogError source, "Err=" & errNum & " Src=" & errSrc & " Desc=" & errDesc

    If Not notifyUser Then Exit Sub

    If IsAppError(errNum) Then
        userMsg = errDesc
    Else
        userMsg = "予期しないエラーが発生しました。"
    End If

    MsgBox userMsg & vbCrLf & vbCrLf & _
           "問い合わせ番号: " & RunId() & vbCrLf & _
           "（詳細はログに記録されています）", _
           vbExclamation, AppTitle()
End Sub

Public Function IsAppError(ByVal num As Long) As Boolean
    IsAppError = (num >= APP_ERR_BASE And num <= APP_ERR_MAX)
End Function

Public Function AppTitle() As String
    On Error Resume Next
    AppTitle = CStr(GetConfig("AppName", "VBA Tool"))
    If Len(AppTitle) = 0 Then AppTitle = "VBA Tool"
End Function


