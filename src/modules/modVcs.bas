Attribute VB_Name = "modVcs"
Option Explicit
'==========================================================
' modVcs : Git管理用のソースエクスポート
'   VBAコードは.xlsm（バイナリ）に埋め込まれ差分が見えないため、
'   全モジュールを テキストとして src\ 配下に書き出す。
'   ※事前に「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」を
'     有効化すること（開発端末のみ。セキュリティ上、配布先では不要）
'==========================================================

Private Const VBEXT_STD_MODULE As Long = 1
Private Const VBEXT_CLASS_MODULE As Long = 2
Private Const VBEXT_MS_FORM As Long = 3
Private Const VBEXT_DOCUMENT As Long = 100

Public Sub ExportAllModules()
    Const PROC As String = "modVcs.ExportAllModules"
    Dim root As String
    Dim comp As Object
    Dim subDir As String
    Dim ext As String
    Dim cnt As Long
    On Error GoTo EH

    root = ThisWorkbook.Path & "\src"

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case VBEXT_STD_MODULE:   subDir = "modules":   ext = ".bas"
            Case VBEXT_CLASS_MODULE: subDir = "classes":   ext = ".cls"
            Case VBEXT_MS_FORM:      subDir = "forms":     ext = ".frm"
            Case VBEXT_DOCUMENT:     subDir = "documents": ext = ".cls"
            Case Else:               subDir = ""
        End Select

        If Len(subDir) > 0 Then
            EnsureFolder root & "\" & subDir
            comp.Export root & "\" & subDir & "\" & comp.Name & ext
            cnt = cnt + 1
        End If
    Next comp

    LogInfo PROC, cnt & "件のモジュールをエクスポート: " & root
    MsgBox cnt & "件のモジュールをエクスポートしました。" & vbCrLf & root, vbInformation, AppTitle()
    Exit Sub

EH:
    If Err.Number = 1004 Then
        LogError PROC, "VBProjectへのアクセスが拒否されました"
        MsgBox "VBAプロジェクトへのアクセスが許可されていません。" & vbCrLf & _
               "トラストセンター → マクロの設定 → 「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。", _
               vbExclamation, AppTitle()
    Else
        HandleError PROC
    End If
End Sub


