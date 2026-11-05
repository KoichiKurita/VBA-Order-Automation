VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================
' clsAppState : Application設定の退避・復元（RAIIパターン）
'   Set guard = New clsAppState で高速化設定を適用し、
'   変数が破棄されたとき（エラー時を含む）に元の設定へ自動で戻す。
'==========================================================

Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mCalculation As XlCalculation
Private mDisplayAlerts As Boolean
Private mStatusBar As Variant

Private Sub Class_Initialize()
    With Application
        mScreenUpdating = .ScreenUpdating
        mEnableEvents = .EnableEvents
        mCalculation = .Calculation
        mDisplayAlerts = .DisplayAlerts
        mStatusBar = .StatusBar

        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    With Application
        .ScreenUpdating = mScreenUpdating
        .EnableEvents = mEnableEvents
        .Calculation = mCalculation
        .DisplayAlerts = mDisplayAlerts
        .StatusBar = mStatusBar
    End With
End Sub

' 進捗表示（ステータスバー）
Public Sub ShowProgress(ByVal message As String)
    Application.StatusBar = message
    DoEvents
End Sub

