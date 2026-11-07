Attribute VB_Name = "modCsv"
Option Explicit
'==========================================================
' modCsv : CSVファイルの読込とパース
'   ・文字コード自動判定（BOM → UTF-8妥当性チェック → Shift_JIS）
'   ・RFC 4180 準拠（引用符内のカンマ・改行、"" によるエスケープ）
'   ・改行コード CRLF / LF / CR いずれにも対応
'   ・高速化：文字列をバイト配列に展開して走査し、
'             フィールドは Mid$ で一括切り出し（1文字ずつ連結しない）
'==========================================================

Private Const CH_QUOTE As Long = 34
Private Const CH_COMMA As Long = 44
Private Const CH_LF As Long = 10
Private Const CH_CR As Long = 13

' パース作業用（このモジュール内でのみ使用）
Private mRows() As Variant
Private mRowCount As Long
Private mFields() As String
Private mFieldCount As Long

'----------------------------------------------------------
' ファイルを読み込み、文字列で返す。encodingName に判定結果を返す
'----------------------------------------------------------
Public Function ReadTextFile(ByVal filePath As String, ByRef encodingName As String) As String
    Const PROC As String = "modCsv.ReadTextFile"
    Dim bytes() As Byte
    Dim text As String

    bytes = ReadAllBytes(filePath)
    encodingName = DetectEncoding(bytes)
    text = DecodeBytes(bytes, encodingName)

    If Left$(text, 1) = ChrW$(&HFEFF) Then text = Mid$(text, 2)   ' BOMを除去

    LogDebug PROC, "文字コード判定: " & encodingName & " / " & (UBound(bytes) + 1) & " bytes"
    ReadTextFile = text
End Function

'----------------------------------------------------------
' 文字コード判定
'----------------------------------------------------------
Public Function DetectEncoding(ByRef bytes() As Byte) As String
    Dim n As Long
    n = UBound(bytes) + 1

    If n >= 3 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            DetectEncoding = "utf-8": Exit Function            ' UTF-8 (BOM付き)
        End If
    End If
    If n >= 2 Then
        If bytes(0) = &HFF And bytes(1) = &HFE Then
            DetectEncoding = "unicode": Exit Function          ' UTF-16LE
        End If
    End If

    If IsValidUtf8(bytes) Then
        DetectEncoding = "utf-8"                               ' UTF-8 (BOMなし)
    Else
        DetectEncoding = "shift_jis"
    End If
End Function

'----------------------------------------------------------
' CSV文字列を2次元配列 (1 To 行数, 1 To 最大列数) に変換
'   fieldCounts(r) に各行の実際の列数を返す（列数不一致の検出用）
'   空行は読み飛ばす
'----------------------------------------------------------
Public Function ParseCsv(ByRef text As String, ByRef fieldCounts() As Long) As Variant
    Const PROC As String = "modCsv.ParseCsv"
    Dim chars() As Byte
    Dim n As Long
    Dim i As Long
    Dim c As Long
    Dim inQuotes As Boolean
    Dim quoted As Boolean
    Dim fieldStart As Long
    Dim result() As String
    Dim rowArr As Variant
    Dim r As Long
    Dim k As Long
    Dim maxCols As Long

    n = Len(text)
    If n = 0 Then RaiseAppError ERR_CSV_FORMAT, PROC, "CSVにデータがありません"

    chars = text                    ' UTF-16LE：1文字 = 2バイト
    mRowCount = 0
    ReDim mRows(1 To 1024)
    ReDim mFields(0 To 31)
    mFieldCount = 0

    fieldStart = 1
    i = 1
    Do While i <= n
        ' i文字目のコード（上位バイトが0 = ASCII のときだけ区切り文字の可能性がある）
        If chars(i + i - 1) = 0 Then c = chars(i + i - 2) Else c = -1

        If inQuotes Then
            If c = CH_QUOTE Then
                If i < n Then
                    If chars(i + i) = CH_QUOTE And chars(i + i + 1) = 0 Then
                        i = i + 1                     ' "" はエスケープされた " → 1文字読み飛ばす
                    Else
                        inQuotes = False
                    End If
                Else
                    inQuotes = False
                End If
            End If
        Else
            Select Case c
                Case CH_QUOTE
                    If i = fieldStart Then inQuotes = True: quoted = True
                Case CH_COMMA
                    AddField text, fieldStart, i - 1, quoted
                    fieldStart = i + 1
                    quoted = False
                Case CH_CR, CH_LF
                    AddField text, fieldStart, i - 1, quoted
                    EndRow
                    If c = CH_CR And i < n Then
                        If chars(i + i) = CH_LF And chars(i + i + 1) = 0 Then i = i + 1   ' CRLF
                    End If
                    fieldStart = i + 1
                    quoted = False
            End Select
        End If
        i = i + 1
    Loop

    If inQuotes Then
        RaiseAppError ERR_CSV_FORMAT, PROC, "ダブルクォートが閉じられていません（" & (mRowCount + 1) & "件目のレコード付近）"
    End If

    ' 最終行（末尾に改行がない場合）
    If fieldStart <= n Or mFieldCount > 0 Then
        AddField text, fieldStart, n, quoted
        EndRow
    End If

    If mRowCount = 0 Then RaiseAppError ERR_CSV_FORMAT, PROC, "CSVにデータがありません"

    ' 可変長の行を2次元配列に詰め替え
    For r = 1 To mRowCount
        If UBound(mRows(r)) + 1 > maxCols Then maxCols = UBound(mRows(r)) + 1
    Next r

    ReDim result(1 To mRowCount, 1 To maxCols)
    ReDim fieldCounts(1 To mRowCount)
    For r = 1 To mRowCount
        rowArr = mRows(r)
        fieldCounts(r) = UBound(rowArr) + 1
        For k = 0 To UBound(rowArr)
            result(r, k + 1) = rowArr(k)
        Next k
    Next r

    Erase mRows
    ParseCsv = result
End Function

'----------------------------------------------------------
' 開発用：判定結果と先頭行をイミディエイトウィンドウに表示
'   例) PreviewCsv ThisWorkbook.Path & "\samples\orders_202611_utf8.csv"
'----------------------------------------------------------
Public Sub PreviewCsv(ByVal filePath As String, Optional ByVal maxRows As Long = 5)
    Dim enc As String
    Dim text As String
    Dim d As Variant
    Dim fc() As Long
    Dim r As Long
    Dim c As Long
    Dim lineText As String

    text = ReadTextFile(filePath, enc)
    d = ParseCsv(text, fc)
    Debug.Print "文字コード: " & enc & " / レコード数: " & UBound(d, 1) & "（ヘッダー含む）"

    For r = 1 To UBound(d, 1)
        If r > maxRows Then Exit For
        lineText = ""
        For c = 1 To fc(r)
            lineText = lineText & "[" & Replace(d(r, c), vbCrLf, "<改行>") & "]"
        Next c
        Debug.Print r & ": " & lineText
    Next r
End Sub

'==========================================================
' 内部処理
'==========================================================
Private Function ReadAllBytes(ByVal filePath As String) As Byte()
    Dim fileNo As Integer
    Dim size As Long
    Dim buf() As Byte

    fileNo = FreeFile
    Open filePath For Binary Access Read Shared As #fileNo
    size = LOF(fileNo)
    If size = 0 Then
        Close #fileNo
        RaiseAppError ERR_CSV_FORMAT, "modCsv.ReadAllBytes", "ファイルが空です"
    End If
    ReDim buf(0 To size - 1)
    Get #fileNo, , buf
    Close #fileNo
    ReadAllBytes = buf
End Function

' UTF-8 として正しいバイト列かを判定
Private Function IsValidUtf8(ByRef b() As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim need As Long
    Dim k As Long

    n = UBound(b)
    i = 0
    Do While i <= n
        Select Case b(i)
            Case &H0 To &H7F:  need = 0
            Case &HC2 To &HDF: need = 1
            Case &HE0 To &HEF: need = 2
            Case &HF0 To &HF4: need = 3
            Case Else: Exit Function           ' False
        End Select
        For k = 1 To need
            If i + k > n Then Exit Function
            If (b(i + k) And &HC0) <> &H80 Then Exit Function
        Next k
        i = i + need + 1
    Loop
    IsValidUtf8 = True
End Function

' ADODB.Stream で指定の文字コードとして文字列化
Private Function DecodeBytes(ByRef bytes() As Byte, ByVal charset As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                ' adTypeBinary
    stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2                ' adTypeText
    stm.charset = charset
    DecodeBytes = stm.ReadText(-1)   ' adReadAll
    stm.Close
End Function

Private Sub AddField(ByRef text As String, ByVal startPos As Long, ByVal endPos As Long, ByVal quoted As Boolean)
    Dim s As String

    If endPos >= startPos Then s = Mid$(text, startPos, endPos - startPos + 1)
    If quoted Then
        s = Mid$(s, 2)                                          ' 先頭の "
        If Right$(s, 1) = """" Then s = Left$(s, Len(s) - 1)    ' 末尾の "
        s = Replace(s, """""", """")                            ' "" → "
    End If

    If mFieldCount > UBound(mFields) Then ReDim Preserve mFields(0 To UBound(mFields) * 2 + 1)
    mFields(mFieldCount) = s
    mFieldCount = mFieldCount + 1
End Sub

Private Sub EndRow()
    Dim arr() As String
    Dim k As Long

    ' 空行（フィールドが1つで空文字）は読み飛ばす
    If mFieldCount = 1 Then
        If Len(mFields(0)) = 0 Then
            mFieldCount = 0
            Exit Sub
        End If
    End If

    ReDim arr(0 To mFieldCount - 1)
    For k = 0 To mFieldCount - 1
        arr(k) = mFields(k)
    Next k

    mRowCount = mRowCount + 1
    If mRowCount > UBound(mRows) Then ReDim Preserve mRows(1 To UBound(mRows) * 2)
    mRows(mRowCount) = arr
    mFieldCount = 0
End Sub

'----------------------------------------------------------
' テキストファイルを UTF-8 で書き出す（Phase 3 追加）
'   withBom = True : BOM付き（Excelでダブルクリックしても文字化けしない）
'   withBom = False: BOMなし（Git・他システム連携向け）
'----------------------------------------------------------
Public Sub WriteTextFile(ByVal filePath As String, ByVal text As String, Optional ByVal withBom As Boolean = True)
    Dim stm As Object
    Dim outStm As Object

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                ' adTypeText
    stm.charset = "utf-8"       ' ADODB.Stream は UTF-8 の先頭に BOM(3バイト)を付ける
    stm.Open
    stm.WriteText text

    If withBom Then
        stm.SaveToFile filePath, 2          ' adSaveCreateOverWrite
    Else
        stm.Position = 0
        stm.Type = 1            ' adTypeBinary
        stm.Position = 3        ' BOM を読み飛ばす
        Set outStm = CreateObject("ADODB.Stream")
        outStm.Type = 1
        outStm.Open
        stm.CopyTo outStm
        outStm.SaveToFile filePath, 2
        outStm.Close
    End If
    stm.Close
End Sub

