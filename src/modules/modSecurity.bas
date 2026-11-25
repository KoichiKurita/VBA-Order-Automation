Attribute VB_Name = "modSecurity"
Option Explicit
'==========================================================
' modSecurity : セキュリティ関連の共通関数
'   CSVインジェクション（Formula Injection）対策
'   参考: OWASP "CSV Injection"
'   ・取込時 : 数式として解釈され得る値を検出し、Config!CsvInjectionPolicy に従い
'              REJECT（エラー行にする）または SANITIZE（文字列として取込＋WARNログ）
'   ・書込時 : Orders の文字列項目は、列に文字列書式（@）を設定してから書き込む（modImport.ApplyTextFormat）
'              その他のシートでは ToSafeCellText（先頭に '）を使う
'   ・CSV出力: ToCsvField で数式様の値に ' を付け、RFC 4180 のエスケープを行う（Phase 3）
'==========================================================

' 数式として解釈され得る先頭文字か（全角も含む）
'   先頭1文字の文字コードを数値で比較する（文字列比較より高速。v0.3 性能改善）
Public Function IsFormulaLike(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    IsFormulaLike = IsFormulaCode(AscW(s))
End Function

' 列単位で数式様の値を判定し、flags(r, col) に結果を返す（v0.3 性能改善）
'   セルごとに IsFormulaLike を呼ぶと文字列のコピーが毎回発生するため、列ごとにまとめて判定する
Public Sub MarkFormulaLike(ByRef values() As Variant, ByVal col As Long, ByRef flags() As Boolean)
    Dim r As Long
    For r = LBound(values, 1) To UBound(values, 1)
        If VarType(values(r, col)) = vbString Then
            If Len(values(r, col)) > 0 Then flags(r, col) = IsFormulaCode(AscW(values(r, col)))
        End If
    Next r
End Sub

' 数式の先頭になり得る文字か（判定する文字の一覧はここ1か所だけで管理する）
Private Function IsFormulaCode(ByVal code As Long) As Boolean
    Select Case code
        Case 61, 43, 45, 64, 9, 13, 10                  ' =  +  -  @  タブ  CR  LF
            IsFormulaCode = True
        Case &HFF1D, &HFF0B, &HFF0D, &HFF20             ' ＝ ＋ − ＠（全角）
            IsFormulaCode = True
    End Select
End Function

' セルに「必ず文字列として」書き込むための値に変換する
'   先頭に ' を付けると Excel は値を文字列として格納し、
'   数式・数値・日付への自動変換を行わない（' 自体はセルの値に含まれない）
'   → 数式の実行防止に加え、"00123" の先頭ゼロ消失や "1-2" の日付化も防げる
Public Function ToSafeCellText(ByVal s As String) As String
    If Len(s) = 0 Then
        ToSafeCellText = ""
    Else
        ToSafeCellText = "'" & s
    End If
End Function

' CSV出力用の1フィールドに変換する（Phase 3 追加）
'   ・文字列が数式様なら先頭に ' を付ける（出力したCSVを開いた人の環境で数式を実行させない）
'   ・カンマ / ダブルクォート / 改行を含む場合は "" で囲み、" を "" にエスケープ
'   ・数値は型のまま出力するため、負数 "-5" に ' は付かない
Public Function ToCsvField(ByVal v As Variant) As String
    Dim s As String

    Select Case VarType(v)
        Case vbEmpty, vbNull
            s = ""
        Case vbDate
            s = Format$(v, "yyyy/mm/dd")
        Case vbString
            s = v
            If IsFormulaLike(s) Then s = "'" & s
        Case Else
            s = CStr(v)
    End Select

    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        s = """" & Replace(s, """", """""") & """"
    End If
    ToCsvField = s
End Function

