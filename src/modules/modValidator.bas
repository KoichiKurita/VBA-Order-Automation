Attribute VB_Name = "modValidator"
Option Explicit
'==========================================================
' modValidator : 項目定義（Schemaシート）に基づく入力値検証と型変換
'   検証ルールをシートで管理する
'   → 項目追加・桁数や範囲の変更にコード修正が不要（運用負荷軽減）
'   高速化（v0.3 性能改善）
'     ・全角→半角変換（StrConv）は、そのままでは解釈できなかった場合だけ行う
'     ・範囲の上限・下限は Schema 読込時に型変換しておき、毎回変換しない
'     ・日付は yyyy/mm/dd 形式を Split なしで解釈する近道を用意
'==========================================================

Private Const SCHEMA_SHEET As String = "Schema"
Private Const SCHEMA_TABLE As String = "tblSchema"

Public Type ColumnDef
    Name As String
    required As Boolean
    dataType As String      ' TEXT / INTEGER / DECIMAL / DATE
    MaxLength As Long       ' 0 = 制限なし（TEXTのみ）
    hasMin As Boolean
    MinValue As Variant     ' LoadSchema で型変換済み（DATE=Date / 数値=Double）
    hasMax As Boolean
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
            .required = ToBool(data(i, 2))
            .dataType = UCase$(Trim$(CStr(data(i, 3))))
            .MaxLength = Val(CStr(data(i, 4)))
            .hasMin = (Len(CStr(data(i, 5))) > 0)
            .MinValue = data(i, 5)
            .hasMax = (Len(CStr(data(i, 6))) > 0)
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

            ' 範囲の比較値を先に型変換しておく（1行ごとの変換を避ける）
            If .dataType <> "TEXT" Then
                On Error GoTo BadBound
                If .hasMin Then .MinValue = ToTyped(.dataType, .MinValue)
                If .hasMax Then .MaxValue = ToTyped(.dataType, .MaxValue)
                On Error GoTo 0
            End If
        End With
    Next i

    LoadSchema = defs
    Exit Function

BadBound:
    RaiseAppError ERR_SCHEMA, PROC, "Schemaの MinValue / MaxValue が型と一致しません：" & defs(i).Name
End Function

'----------------------------------------------------------
' 1項目の検証と型変換
'   戻り値   : エラーメッセージ（正常なら空文字）
'   converted: 型変換後の値（TEXT=String, INTEGER/DECIMAL=Double, DATE=Date, 未入力=Empty）
'----------------------------------------------------------
Public Function ValidateValue(ByRef def As ColumnDef, ByVal raw As String, ByRef converted As Variant) As String
    Dim v As Variant
    Dim narrow As String

    converted = Empty
    raw = Trim$(raw)

    '--- 必須
    If Len(raw) = 0 Then
        If def.required Then ValidateValue = "必須項目が未入力です"
        Exit Function
    End If

    '--- 型
    If def.dataType = "TEXT" Then
        raw = Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf)   ' セル内改行はLFに統一
        If def.MaxLength > 0 And Len(raw) > def.MaxLength Then
            ValidateValue = def.MaxLength & "文字以内で入力してください（" & Len(raw) & "文字）"
            Exit Function
        End If
        If HasControlChar(raw) Then
            ValidateValue = "制御文字（タブ等）が含まれています"
            Exit Function
        End If
        v = raw
    Else
        If Not TryParseTyped(def.dataType, raw, v) Then
            ' 全角数字などの入力ゆれを吸収してから再試行（大半の行は半角なので、変換は失敗時だけ）
            narrow = StrConv(raw, vbNarrow)
            If narrow = raw Then
                ValidateValue = TypeErrorMessage(def.dataType)
                Exit Function
            End If
            raw = narrow
            If Not TryParseTyped(def.dataType, raw, v) Then
                ValidateValue = TypeErrorMessage(def.dataType)
                Exit Function
            End If
        End If
    End If

    '--- 形式（パターン）
    If Len(def.pattern) > 0 Then
        If Not (raw Like def.pattern) Then
            ValidateValue = "形式が不正です"
            If Len(def.Hint) > 0 Then ValidateValue = ValidateValue & "（例：" & def.Hint & "）"
            Exit Function
        End If
    End If

    '--- 範囲（MinValue / MaxValue は LoadSchema で型変換済み）
    If def.dataType <> "TEXT" Then
        If def.hasMin Then
            If v < def.MinValue Then ValidateValue = "範囲外です" & RangeText(def): Exit Function
        End If
        If def.hasMax Then
            If v > def.MaxValue Then ValidateValue = "範囲外です" & RangeText(def): Exit Function
        End If
    End If

    converted = v
End Function

'----------------------------------------------------------
' 列単位の一括検証（v0.3 性能改善）
'   data(2 To 最終行, srcCol) を検証し、converted(r, dstCol) と messages(r, dstCol) に結果を返す
'   ・項目定義をローカル変数に展開し、ループ内では「正常な値」を判定する近道だけを通す
'   ・近道で正常と確定できなかった値だけ ValidateValue に任せる
'     → エラーメッセージや全角変換のロジックは ValidateValue の1か所だけに保たれる
'----------------------------------------------------------
Public Sub ValidateColumn(ByRef def As ColumnDef, ByRef data As Variant, ByVal srcCol As Long, _
                          ByRef converted() As Variant, ByRef messages() As String, ByVal dstCol As Long)
    Dim r As Long
    Dim raw As String
    Dim v As Variant
    Dim ok As Boolean
    Dim isText As Boolean
    Dim dataType As String
    Dim required As Boolean
    Dim maxLen As Long
    Dim pattern As String
    Dim hasMin As Boolean
    Dim hasMax As Boolean
    Dim minV As Variant
    Dim maxV As Variant

    dataType = def.dataType
    isText = (dataType = "TEXT")
    required = def.required
    maxLen = def.MaxLength
    pattern = def.pattern
    hasMin = def.hasMin And Not isText
    hasMax = def.hasMax And Not isText
    minV = def.MinValue
    maxV = def.MaxValue

    For r = 2 To UBound(data, 1)
        raw = Trim$(data(r, srcCol))
        ok = False

        If Len(raw) = 0 Then
            ok = Not required                           ' 任意項目の未入力は正常（Empty のまま）
        ElseIf isText Then
            If InStr(raw, vbCr) = 0 Then
                If maxLen = 0 Or Len(raw) <= maxLen Then
                    If Not HasControlChar(raw) Then
                        If Len(pattern) = 0 Then ok = True Else ok = (raw Like pattern)
                        If ok Then converted(r, dstCol) = raw
                    End If
                End If
            End If
        Else
            ok = TryParseTyped(dataType, raw, v)
            If ok And Len(pattern) > 0 Then ok = (raw Like pattern)
            If ok And hasMin Then ok = Not (v < minV)
            If ok And hasMax Then ok = Not (v > maxV)
            If ok Then converted(r, dstCol) = v
        End If

        If Not ok Then messages(r, dstCol) = ValidateValue(def, data(r, srcCol), converted(r, dstCol))
    Next r
End Sub

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

Private Function TryParseTyped(ByVal dataType As String, ByVal raw As String, ByRef v As Variant) As Boolean
    Select Case dataType
        Case "INTEGER"
            If IsIntegerText(raw) Then v = CDbl(raw): TryParseTyped = True
        Case "DECIMAL"
            If IsDecimalText(raw) Then v = Val(raw): TryParseTyped = True
        Case "DATE"
            TryParseTyped = TryParseDate(raw, v)
    End Select
End Function

Private Function TypeErrorMessage(ByVal dataType As String) As String
    Select Case dataType
        Case "INTEGER": TypeErrorMessage = "整数で入力してください"
        Case "DECIMAL": TypeErrorMessage = "数値で入力してください（桁区切りのカンマは不可）"
        Case "DATE":    TypeErrorMessage = "正しい日付（yyyy/mm/dd）で入力してください"
    End Select
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

    ' 近道：最も多い yyyy/mm/dd（10文字）は Split を使わずに切り出す
    If Len(s) = 10 Then
        If s Like "####[/-]##[/-]##" Then
            y = CLng(Left$(s, 4))
            m = CLng(Mid$(s, 6, 2))
            d = CLng(Right$(s, 2))
            GoTo Validate
        End If
    End If

    parts = Split(Replace(s, "-", "/"), "/")
    If UBound(parts) <> 2 Then Exit Function
    If Len(parts(0)) <> 4 Then Exit Function
    If Len(parts(1)) < 1 Or Len(parts(1)) > 2 Then Exit Function
    If Len(parts(2)) < 1 Or Len(parts(2)) > 2 Then Exit Function
    If (parts(0) & parts(1) & parts(2)) Like "*[!0-9]*" Then Exit Function

    y = CLng(parts(0))
    m = CLng(parts(1))
    d = CLng(parts(2))

Validate:
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
    If def.hasMin Then lo = BoundText(def.dataType, def.MinValue)
    If def.hasMax Then hi = BoundText(def.dataType, def.MaxValue)
    RangeText = "（許容範囲：" & lo & " 〜 " & hi & "）"
End Function

Private Function BoundText(ByVal dataType As String, ByVal v As Variant) As String
    If dataType = "DATE" Then
        BoundText = Format$(v, "yyyy/mm/dd")
    Else
        BoundText = CStr(v)
    End If
End Function


