Attribute VB_Name = "modUI"
Option Explicit
'==========================================================
' modUI : 画面の表示と、処理から画面への進捗の受け渡し
'   処理層（取込・集計など）は画面を直接知らない。
'   clsAppState.ShowProgress → ReportProgress → 登録された画面（SetStatus）の順に伝える
'   → 画面がない無人実行（Phase 5）でも、同じ処理をそのまま使える
'==========================================================

Private mSink As Object

Public Sub ShowMenu()
    frmMenu.Show vbModeless             ' モードレス：メニューを開いたままシートを操作できる
End Sub

Public Sub SetProgressSink(ByVal sink As Object)
    Set mSink = sink
End Sub

Public Sub ClearProgressSink()
    Set mSink = Nothing
End Sub

Public Sub ReportProgress(ByVal message As String)
    If mSink Is Nothing Then Exit Sub
    On Error Resume Next                ' 画面が閉じられていても処理は止めない
    mSink.SetStatus message
End Sub

