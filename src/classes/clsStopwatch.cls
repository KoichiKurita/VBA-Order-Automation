VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================
' clsStopwatch : 高精度ストップウォッチ（性能計測用）
'   VBA の Timer は約 1/64 秒精度・日付をまたぐと0に戻るため、
'   Windows API の QueryPerformanceCounter を使う
'   #If VBA7 / PtrSafe により 32bit / 64bit 両方の Office で動作する
'==========================================================

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Private mFreq As Currency
Private mStart As Currency

Private Sub Class_Initialize()
    QueryPerformanceFrequency mFreq
    StartTimer
End Sub

Public Sub StartTimer()
    QueryPerformanceCounter mStart
End Sub

' 経過秒数（Currency型の1/10000スケールは割り算で相殺される）
Public Function ElapsedSeconds() As Double
    Dim nowCount As Currency
    QueryPerformanceCounter nowCount
    ElapsedSeconds = (nowCount - mStart) / mFreq
End Function

