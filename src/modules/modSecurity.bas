Attribute VB_Name = "modSecurity"
Option Explicit
'==========================================================
' modSecurity : セキュリティ関連の共通関数
'   CSVインジェクション（Formula Injection）対策
'   参考: OWASP "CSV Injection"
'   ・取込時 : 数式として解釈され得る値を検出し、Config!CsvInjectionPolicy に従い
'              REJECT（エラー行にする）または SANITIZE（文字列として取込＋WARNログ）
'   ・書込時 : 文字列項目は必ず ToSafeCellText を通し、数式として評価させない
'   ・CSV出力: ToCsvField で数式様の値に ' を付け、RFC 4180 のエスケープを行う（Phase 3）
'==========================================================

' 数式として解釈され得る先頭文字か（全角も含む）
Public Function IsFormulaLike(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Select Case Left$(s, 1)
        Case "=", "+", "-", "@", vbTab, vbCr, vbLf, "＝", "＋", "−", "＠"
            IsFormulaLike = True
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

