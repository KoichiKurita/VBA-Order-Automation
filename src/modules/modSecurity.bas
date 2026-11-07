Attribute VB_Name = "modSecurity"
Option Explicit
'==========================================================
' modSecurity : セキュリティ関連の共通関数
'   CSVインジェクション（Formula Injection）対策
'   参考: OWASP "CSV Injection"
'   ・取込時 : 数式として解釈され得る値を検出し、Config!CsvInjectionPolicy に従い
'              REJECT（エラー行にする）または SANITIZE（文字列として取込＋WARNログ）
'   ・書込時 : 文字列項目は必ず ToSafeCellText を通し、数式として評価させない
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

