Attribute VB_Name = "modConfig"
Option Explicit
'==========================================================
' modConfig : 設定値の一元管理
'   Configシートのテーブル(tblConfig: Key / Value / Description)から
'   値を読み込み、Dictionaryにキャッシュする。
'   → コードを変更せずに運用設定を変更できる（運用負荷軽減）
'==========================================================

Private Const CONFIG_SHEET As String = "Config"
Private Const CONFIG_TABLE As String = "tblConfig"

Private mCache As Object 'Scripting.Dictionary（参照設定不要の遅延バインディング）

' 設定値の取得。キーが無く既定値も無い場合は業務エラー
Public Function GetConfig(ByVal key As String, Optional ByVal defaultValue As Variant) As Variant
    If mCache Is Nothing Then LoadConfig

    If mCache.Exists(key) Then
        GetConfig = mCache(key)
    ElseIf Not IsMissing(defaultValue) Then
        GetConfig = defaultValue
    Else
        Err.Raise ERR_CONFIG_NOT_FOUND, "modConfig.GetConfig", _
                  "設定キーが見つかりません: " & key & "（Configシートを確認してください）"
    End If
End Function

' 設定の再読込（Configシート編集後に呼ぶ）
Public Sub ReloadConfig()
    Set mCache = Nothing
    LoadConfig
End Sub

' パス系設定の解決
'   ・%USERPROFILE% を展開
'   ・相対パスはブックの保存フォルダ基準で絶対パス化
Public Function ResolvePath(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    Dim p As String
    If Len(defaultValue) > 0 Then
        p = CStr(GetConfig(key, defaultValue))
    Else
        p = CStr(GetConfig(key))
    End If

    p = Replace(p, "%USERPROFILE%", Environ$("USERPROFILE"))
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)

    If Not (p Like "?:\*" Or p Like "\\*") Then
        p = ThisWorkbook.Path & "\" & p
    End If
    ResolvePath = p
End Function

Private Sub LoadConfig()
    Dim lo As ListObject
    Dim data As Variant
    Dim i As Long
    Dim k As String

    Set mCache = CreateObject("Scripting.Dictionary")
    mCache.CompareMode = vbTextCompare

    Set lo = ThisWorkbook.Worksheets(CONFIG_SHEET).ListObjects(CONFIG_TABLE)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    data = lo.DataBodyRange.Value   ' 配列で一括取得（セル単位アクセスを避ける）
    For i = 1 To UBound(data, 1)
        k = Trim$(CStr(data(i, 1)))
        If Len(k) > 0 Then mCache(k) = data(i, 2)
    Next i
End Sub


