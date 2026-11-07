VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================
' clsImportResult : 取込結果
'   処理層（ImportOrdersFromFile）が画面に依存せず結果を返すためのクラス
'   → Phase 5 の無人実行では、この結果をログ・通知に使う
'==========================================================

Public filePath As String
Public encodingName As String
Public TotalRows As Long          ' データ行数（ヘッダーを除く）
Public ValidRows As Long          ' 検証OKの行数
Public ErrorRows As Long          ' エラーのあった行数
Public ImportedRows As Long       ' 実際に取り込んだ行数
Public FormulaLikeCount As Long   ' 数式様の値を文字列として取り込んだ件数
Public Committed As Boolean       ' Ordersへ書き込んだか
Public ElapsedSeconds As Double

Public Function summary() As String
    Dim s As String
    s = "ファイル　　：" & Mid$(filePath, InStrRev(filePath, "\") + 1) & vbCrLf & _
        "文字コード　：" & encodingName & vbCrLf & _
        "データ件数　：" & Format$(TotalRows, "#,##0") & " 件" & vbCrLf & _
        "取込件数　　：" & Format$(ImportedRows, "#,##0") & " 件" & vbCrLf & _
        "エラー行数　：" & Format$(ErrorRows, "#,##0") & " 件" & vbCrLf
    If FormulaLikeCount > 0 Then
        s = s & "数式様の値　：" & FormulaLikeCount & " 件（文字列として取込）" & vbCrLf
    End If
    s = s & "処理時間　　：" & Format$(ElapsedSeconds, "0.00") & " 秒"
    summary = s
End Function


