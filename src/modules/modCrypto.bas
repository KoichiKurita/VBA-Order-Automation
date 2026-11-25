Attribute VB_Name = "modCrypto"
Option Explicit
'==========================================================
' modCrypto : 暗号関連（Windows CNG / bcrypt.dll）
'   ・SHA-256 ハッシュ   … 監査ログの改ざん検知に使う
'   ・安全な乱数         … シート保護パスワードの生成に使う
'     （VBA の Rnd は次の値を予測できるため、セキュリティ用途には使わない）
'   ※ PtrSafe / LongPtr を使うため Office 2010 以降（32bit / 64bit 両対応）
'   ※ BCryptHash は Windows 10 以降
'==========================================================

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptHash Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, _
    ByVal pbInput As LongPtr, ByVal cbInput As Long, _
    ByVal pbOutput As LongPtr, ByVal cbOutput As Long) As Long
Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, ByVal pbBuffer As LongPtr, _
    ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long

Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = 2
Private Const ERR_CRYPTO As Long = vbObjectError + 9001

' パスワードに使う文字（0/O、1/l/I など見間違えやすい文字を除外）
Private Const TOKEN_CHARS As String = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789"

'----------------------------------------------------------
' 文字列の SHA-256（VBA の文字列 = UTF-16LE のバイト列に対して計算）。16進小文字64桁
'----------------------------------------------------------
Public Function Sha256Hex(ByVal text As String) As String
    Dim b() As Byte
    b = text
    Sha256Hex = Sha256HexBytes(b)
End Function

'----------------------------------------------------------
' バイト列の SHA-256。16進小文字64桁
'----------------------------------------------------------
Public Function Sha256HexBytes(ByRef data() As Byte) As String
    Dim hAlg As LongPtr
    Dim status As Long
    Dim digest(0 To 31) As Byte
    Dim n As Long
    Dim pIn As LongPtr
    Dim i As Long
    Dim s As String

    status = BCryptOpenAlgorithmProvider(hAlg, StrPtr("SHA256"), 0, 0)
    If status <> 0 Then Err.Raise ERR_CRYPTO, "modCrypto.Sha256HexBytes", "SHA-256 の初期化に失敗しました（0x" & Hex$(status) & "）"

    n = UBound(data) - LBound(data) + 1
    If n > 0 Then pIn = VarPtr(data(LBound(data)))
    status = BCryptHash(hAlg, 0, 0, pIn, n, VarPtr(digest(0)), 32)
    BCryptCloseAlgorithmProvider hAlg, 0
    If status <> 0 Then Err.Raise ERR_CRYPTO, "modCrypto.Sha256HexBytes", "SHA-256 の計算に失敗しました（0x" & Hex$(status) & "）"

    For i = 0 To 31
        s = s & Right$("0" & LCase$(Hex$(digest(i))), 2)
    Next i
    Sha256HexBytes = s
End Function

'----------------------------------------------------------
' 暗号論的に安全な乱数でランダムな文字列を作る
'   偏りが出ないよう、文字数の倍数に収まらないバイトは捨てる（棄却法）
'----------------------------------------------------------
Public Function RandomToken(ByVal length As Long) As String
    Dim buf(0 To 63) As Byte
    Dim nChars As Long
    Dim limit As Long
    Dim i As Long
    Dim s As String

    nChars = Len(TOKEN_CHARS)
    limit = (256 \ nChars) * nChars
    Do While Len(s) < length
        If BCryptGenRandom(0, VarPtr(buf(0)), 64, BCRYPT_USE_SYSTEM_PREFERRED_RNG) <> 0 Then
            Err.Raise ERR_CRYPTO, "modCrypto.RandomToken", "乱数の生成に失敗しました"
        End If
        For i = 0 To 63
            If buf(i) < limit Then
                s = s & Mid$(TOKEN_CHARS, (buf(i) Mod nChars) + 1, 1)
                If Len(s) = length Then Exit For
            End If
        Next i
    Loop
    RandomToken = s
End Function


