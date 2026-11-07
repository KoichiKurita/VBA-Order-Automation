Attribute VB_Name = "modValidator"
Option Explicit
'==========================================================
' modValidator : 項目定義（Schemaシート）に基づく入力値検証と型変換
'   検証ルールをシートで管理する
'   → 項目追加・桁数や範囲の変更にコード修正が不要（運用負荷軽減）
'==========================================================

Private Const SCHEMA_SHEET As String = "Schema"
Private Const SCHEMA_TABLE As String = "tblSchema"

Public Type ColumnDef
    Name As String
    Required As Boolean
    dataType As String      ' TEXT / INTEGER / DECIMAL / DATE
    MaxLength As Long       ' 0 = 制限なし（TEXTのみ）
    HasMin As Boolean
    MinValue As Variant
    HasMax As Boolean
    MaxValue As Variant
    pattern As String       ' VBA の Like 演算子のパターン（# = 数字1桁）。空ならチェックしない
    Hint As String          ' エラー時に表示する入力例
End Type

'----------------------------------------------------------
' Schemaシートから項目定義を読み込む
'----------------------------------------------------------
Public Function LoadSchema() As ColumnDef()
    Const PROC As String = "modValidator.LoadSchema"
    Dim lo As ListObject
    Dim data As Variant
    Dim defs() As ColumnDef
    Dim i As Long

    Set lo = ThisWorkbook.Worksheets(SCHEMA_SHEET).ListObjects(SCHEMA_TABLE)
    If lo.DataBodyRange Is Nothing Then
        RaiseAppError ERR_SCHEMA, PROC, "Schemaシートに項目定義がありません"
    End If

    data = lo.DataBodyRange.value
    ReDim defs(1 To UBound(data, 1))

    For i = 1 To UBound(data, 1)
        With defs(i)
            .Name = Trim$(CStr(data(i, 1)))
            .Required = ToBool(data(i, 2))
            .dataType = UCase$(Trim$(CStr(data(i, 3))))
            .MaxLength = Val(CStr(data(i, 4)))
            .HasMin = (Len(CStr(data(i, 5))) > 0)
            .MinValue = data(i, 5)
            .HasMax = (Len(CStr(data(i, 6))) > 0)
            .MaxValue = data(i, 6)
            .pattern = CStr(data(i, 7))
            .Hint = CStr(data(i, 8))

            If Len(.Name) = 0 Then
                RaiseAppError ERR_SCHEMA, PROC, "Schemaの " & i & " 行目：項目名が空です"
            End If
            Select Case .dataType
                Case "TEXT", "INTEGER", "DECIMAL", "DATE"
                Case Else
                    RaiseAppError ERR_SCHEMA, PROC, "Schemaの型が不正です：" & .Name & "（" & .dataType & "）"
            End Select
        End With
    Next i

    LoadSchema = defs
End Function

'----------------------------------------------------------
' 1項目の検証と型変換
'   戻り値   : エラーメッセージ（正常なら空文字）
'   converted: 型変換後の値（TEXT=String, INTEGER/DECIMAL=Double, DATE=Date, 未入力=Empty）
'----------------------------------------------------------
Public Function ValidateValue(ByRef def As ColumnDef, ByVal raw As String, ByRef converted As Variant) As String
    Dim v As Variant

    converted = Empty
    raw = Trim$(raw)
    If def.dataType <> "TEXT" Then raw = StrConv(raw, vbNarrow)   ' 全角数字・記号を半角に（入力ゆれの吸収）

    '--- 必須
    If Len(raw) = 0 Then
        If def.Required Then ValidateValue = "必須項目が未入力です"
        Exit Function
    End If

    '--- 型
    Select Case def.dataType
        Case "TEXT"
            If def.MaxLength > 0 And Len(raw) > def.MaxLength Then
                ValidateValue = def.MaxLength & "文字以内で入力してください（" & Len(raw) & "文字）"
                Exit Function
            End If
            If HasControlChar(raw) Then
                ValidateValue = "制御文字（タブ等）が含まれています"
                Exit Function
            End If
            v = raw

        Case "INTEGER"
            If Not IsIntegerText(raw) Then
                ValidateValue = "整数で入力してください"
                Exit Function
            End If
            v = CDbl(raw)

        Case "DECIMAL"
            If Not IsDecimalText(raw) Then
                ValidateValue = "数値で入力してください（桁区切りのカンマは不可）"
                Exit Function
            End If
            v = Val(raw)

        Case "DATE"
            If Not TryParseDate(raw, v) Then
                ValidateValue = "正しい日付（yyyy/mm/dd）で入力してください"
                Exit Function
            End If
    End Select

    '--- 形式（パターン）
    If Len(def.pattern) > 0 Then
        If Not (raw Like def.pattern) Then
            ValidateValue = "形式が不正です"
            If Len(def.Hint) > 0 Then ValidateValue = ValidateValue & "（例：" & def.Hint & "）"
            Exit Function
        End If
    End If

    '--- 範囲
    If def.dataType <> "TEXT" Then
        If def.HasMin Then
            If v < ToTyped(def.dataType, def.MinValue) Then ValidateValue = "範囲外です" & RangeText(def): Exit Function
        End If
        If def.HasMax Then
            If v > ToTyped(def.dataType, def.MaxValue) Then ValidateValue = "範囲外です" & RangeText(def): Exit Function
        End If
    End If

    converted = v
End Function

'==========================================================
' 内部処理
'==========================================================
Private Function ToBool(ByVal v As Variant) As Boolean
    Select Case UCase$(Trim$(CStr(v)))
        Case "TRUE", "1", "Y", "YES", "○", "必須"
            ToBool = True
    End Select
End Function

Private Function HasControlChar(ByVal s As String) As Boolean
    Static pattern As String
    ' 改行（CR/LF）は備考などの複数行入力を許容するため除外
    If Len(pattern) = 0 Then
        pattern = "*[" & Chr$(1) & "-" & Chr$(9) & Chr$(11) & Chr$(12) & Chr$(14) & "-" & Chr$(31) & "]*"
    End If
    HasControlChar = (s Like pattern)
End Function

Private Function IsIntegerText(ByVal s As String) As Boolean
    If Left$(s, 1) = "-" Then s = Mid$(s, 2)
    If Len(s) = 0 Or Len(s) > 15 Then Exit Function
    IsIntegerText = Not (s Like "*[!0-9]*")
End Function

Private Function IsDecimalText(ByVal s As String) As Boolean
    Dim p As Long
    Dim intPart As String
    Dim fracPart As String

    If Left$(s, 1) = "-" Then s = Mid$(s, 2)
    If Len(s) = 0 Or Len(s) > 20 Then Exit Function

    p = InStr(s, ".")
    If p = 0 Then
        IsDecimalText = Not (s Like "*[!0-9]*")
        Exit Function
    End If

    intPart = Left$(s, p - 1)
    fracPart = Mid$(s, p + 1)
    If Len(intPart) = 0 Or Len(fracPart) = 0 Then Exit Function
    IsDecimalText = Not (intPart Like "*[!0-9]*") And Not (fracPart Like "*[!0-9]*")
End Function

' yyyy/m/d, yyyy-mm-dd を厳密に解釈（IsDate は "1/2" 等も日付とみなすため使わない）
Private Function TryParseDate(ByVal s As String, ByRef result As Variant) As Boolean
    Dim parts() As String
    Dim y As Long
    Dim m As Long
    Dim d As Long

    parts = Split(Replace(s, "-", "/"), "/")
    If UBound(parts) <> 2 Then Exit Function
    If Len(parts(0)) <> 4 Then Exit Function
    If Len(parts(1)) < 1 Or Len(parts(1)) > 2 Then Exit Function
    If Len(parts(2)) < 1 Or Len(parts(2)) > 2 Then Exit Function
    If (parts(0) & parts(1) & parts(2)) Like "*[!0-9]*" Then Exit Function

    y = CLng(parts(0))
    m = CLng(parts(1))
    d = CLng(parts(2))
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function

    result = DateSerial(y, m, d)
    ' 2/30 などは DateSerial が翌月に繰り越すため、往復で一致するか確認
    TryParseDate = (Month(result) = m And Day(result) = d)
End Function

Private Function ToTyped(ByVal dataType As String, ByVal v As Variant) As Variant
    If dataType = "DATE" Then
        ToTyped = CDate(v)
    Else
        ToTyped = CDbl(v)
    End If
End Function

Private Function RangeText(ByRef def As ColumnDef) As String
    Dim lo As String
    Dim hi As String
    If def.HasMin Then lo = BoundText(def.dataType, def.MinValue)
    If def.HasMax Then hi = BoundText(def.dataType, def.MaxValue)
    RangeText = "（許容範囲：" & lo & " 〜 " & hi & "）"
End Function

Private Function BoundText(ByVal dataType As String, ByVal v As Variant) As String
    If dataType = "DATE" Then
        BoundText = Format$(CDate(v), "yyyy/mm/dd")
    Else
        BoundText = CStr(v)
    End If
End Function

