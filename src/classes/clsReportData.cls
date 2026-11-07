VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================
' clsReportData : 集計結果
'   集計（BuildReport）と出力（PDF / CSV）を分離するためのデータクラス
'   ByXxx は 2次元配列 (1 To 件数, 1 To 6)
'     1:キー  2:名称  3:件数  4:数量  5:金額  6:構成比
'==========================================================

Public ReportId As String
Public fromDate As Date
Public toDate As Date
Public OrderCount As Long
Public TotalQty As Double
Public TotalAmount As Double
Public ByCustomer As Variant
Public ByProduct As Variant
Public ByRep As Variant
Public ByDate As Variant
Public ElapsedSeconds As Double

