Attribute VB_Name = "modMD5"
Option Explicit

Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Dim m_lOnBits(30) As Long
Dim m_l2Power(30) As Long
 
Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    On Error Resume Next

    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    On Error Resume Next

    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4 As Long
    Dim lY4 As Long
    Dim lX8 As Long
    Dim lY8 As Long
    Dim lResult As Long

    On Error Resume Next

    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function

Private Function md5_F(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    md5_F = (x And y) Or ((Not x) And z)
End Function

Private Function md5_G(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    md5_G = (x And z) Or (y And (Not z))
End Function

Private Function md5_H(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    md5_H = (x Xor y Xor z)
End Function

Private Function md5_I(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    md5_I = (y Xor (x Or (Not z)))
End Function

Private Sub md5_FF(ByRef a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal x As Long, ByVal s As Long, ByVal ac As Long)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_GG(ByRef a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal x As Long, ByVal s As Long, ByVal ac As Long)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_HH(ByRef a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal x As Long, ByVal s As Long, ByVal ac As Long)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_II(ByRef a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal x As Long, ByVal s As Long, ByVal ac As Long)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(ByVal sMessage As String) As Long()
    Dim lMessageLength As Long
    Dim lNumberOfWords As Long
    Dim lWordArray() As Long
    Dim lBytePosition As Long
    Dim lByteCount As Long
    Dim lWordCount As Long

    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    
    On Error Resume Next

    lMessageLength = Len(sMessage)
    
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(ByVal lValue As Long)
    Dim lByte As Long
    Dim lCount As Long

    On Error Resume Next

    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next lCount
End Function

Public Function MD5(ByVal sMessage As String) As String
    Dim x() As Long
    Dim k As Long
    Dim AA As Long
    Dim BB As Long
    Dim CC As Long
    Dim DD As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long

    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21

    On Error Resume Next

    m_lOnBits(0) = 1
    m_lOnBits(1) = 3
    m_lOnBits(2) = 7
    m_lOnBits(3) = 15
    m_lOnBits(4) = 31
    m_lOnBits(5) = 63
    m_lOnBits(6) = 127
    m_lOnBits(7) = 255
    m_lOnBits(8) = 511
    m_lOnBits(9) = 1023
    m_lOnBits(10) = 2047
    m_lOnBits(11) = 4095
    m_lOnBits(12) = 8191
    m_lOnBits(13) = 16383
    m_lOnBits(14) = 32767
    m_lOnBits(15) = 65535
    m_lOnBits(16) = 131071
    m_lOnBits(17) = 262143
    m_lOnBits(18) = 524287
    m_lOnBits(19) = 1048575
    m_lOnBits(20) = 2097151
    m_lOnBits(21) = 4194303
    m_lOnBits(22) = 8388607
    m_lOnBits(23) = 16777215
    m_lOnBits(24) = 33554431
    m_lOnBits(25) = 67108863
    m_lOnBits(26) = 134217727
    m_lOnBits(27) = 268435455
    m_lOnBits(28) = 536870911
    m_lOnBits(29) = 1073741823
    m_lOnBits(30) = 2147483647
    
    m_l2Power(0) = 1
    m_l2Power(1) = 2
    m_l2Power(2) = 4
    m_l2Power(3) = 8
    m_l2Power(4) = 16
    m_l2Power(5) = 32
    m_l2Power(6) = 64
    m_l2Power(7) = 128
    m_l2Power(8) = 256
    m_l2Power(9) = 512
    m_l2Power(10) = 1024
    m_l2Power(11) = 2048
    m_l2Power(12) = 4096
    m_l2Power(13) = 8192
    m_l2Power(14) = 16384
    m_l2Power(15) = 32768
    m_l2Power(16) = 65536
    m_l2Power(17) = 131072
    m_l2Power(18) = 262144
    m_l2Power(19) = 524288
    m_l2Power(20) = 1048576
    m_l2Power(21) = 2097152
    m_l2Power(22) = 4194304
    m_l2Power(23) = 8388608
    m_l2Power(24) = 16777216
    m_l2Power(25) = 33554432
    m_l2Power(26) = 67108864
    m_l2Power(27) = 134217728
    m_l2Power(28) = 268435456
    m_l2Power(29) = 536870912
    m_l2Power(30) = 1073741824

    x = ConvertToWordArray(sMessage)

    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476

    For k = 0 To UBound(x) Step 16
        AA = a
        BB = b
        CC = c
        DD = d

        Call md5_FF(a, b, c, d, x(k + 0), S11, &HD76AA478)
        Call md5_FF(d, a, b, c, x(k + 1), S12, &HE8C7B756)
        Call md5_FF(c, d, a, b, x(k + 2), S13, &H242070DB)
        Call md5_FF(b, c, d, a, x(k + 3), S14, &HC1BDCEEE)
        Call md5_FF(a, b, c, d, x(k + 4), S11, &HF57C0FAF)
        Call md5_FF(d, a, b, c, x(k + 5), S12, &H4787C62A)
        Call md5_FF(c, d, a, b, x(k + 6), S13, &HA8304613)
        Call md5_FF(b, c, d, a, x(k + 7), S14, &HFD469501)
        Call md5_FF(a, b, c, d, x(k + 8), S11, &H698098D8)
        Call md5_FF(d, a, b, c, x(k + 9), S12, &H8B44F7AF)
        Call md5_FF(c, d, a, b, x(k + 10), S13, &HFFFF5BB1)
        Call md5_FF(b, c, d, a, x(k + 11), S14, &H895CD7BE)
        Call md5_FF(a, b, c, d, x(k + 12), S11, &H6B901122)
        Call md5_FF(d, a, b, c, x(k + 13), S12, &HFD987193)
        Call md5_FF(c, d, a, b, x(k + 14), S13, &HA679438E)
        Call md5_FF(b, c, d, a, x(k + 15), S14, &H49B40821)

        Call md5_GG(a, b, c, d, x(k + 1), S21, &HF61E2562)
        Call md5_GG(d, a, b, c, x(k + 6), S22, &HC040B340)
        Call md5_GG(c, d, a, b, x(k + 11), S23, &H265E5A51)
        Call md5_GG(b, c, d, a, x(k + 0), S24, &HE9B6C7AA)
        Call md5_GG(a, b, c, d, x(k + 5), S21, &HD62F105D)
        Call md5_GG(d, a, b, c, x(k + 10), S22, &H2441453)
        Call md5_GG(c, d, a, b, x(k + 15), S23, &HD8A1E681)
        Call md5_GG(b, c, d, a, x(k + 4), S24, &HE7D3FBC8)
        Call md5_GG(a, b, c, d, x(k + 9), S21, &H21E1CDE6)
        Call md5_GG(d, a, b, c, x(k + 14), S22, &HC33707D6)
        Call md5_GG(c, d, a, b, x(k + 3), S23, &HF4D50D87)
        Call md5_GG(b, c, d, a, x(k + 8), S24, &H455A14ED)
        Call md5_GG(a, b, c, d, x(k + 13), S21, &HA9E3E905)
        Call md5_GG(d, a, b, c, x(k + 2), S22, &HFCEFA3F8)
        Call md5_GG(c, d, a, b, x(k + 7), S23, &H676F02D9)
        Call md5_GG(b, c, d, a, x(k + 12), S24, &H8D2A4C8A)

        Call md5_HH(a, b, c, d, x(k + 5), S31, &HFFFA3942)
        Call md5_HH(d, a, b, c, x(k + 8), S32, &H8771F681)
        Call md5_HH(c, d, a, b, x(k + 11), S33, &H6D9D6122)
        Call md5_HH(b, c, d, a, x(k + 14), S34, &HFDE5380C)
        Call md5_HH(a, b, c, d, x(k + 1), S31, &HA4BEEA44)
        Call md5_HH(d, a, b, c, x(k + 4), S32, &H4BDECFA9)
        Call md5_HH(c, d, a, b, x(k + 7), S33, &HF6BB4B60)
        Call md5_HH(b, c, d, a, x(k + 10), S34, &HBEBFBC70)
        Call md5_HH(a, b, c, d, x(k + 13), S31, &H289B7EC6)
        Call md5_HH(d, a, b, c, x(k + 0), S32, &HEAA127FA)
        Call md5_HH(c, d, a, b, x(k + 3), S33, &HD4EF3085)
        Call md5_HH(b, c, d, a, x(k + 6), S34, &H4881D05)
        Call md5_HH(a, b, c, d, x(k + 9), S31, &HD9D4D039)
        Call md5_HH(d, a, b, c, x(k + 12), S32, &HE6DB99E5)
        Call md5_HH(c, d, a, b, x(k + 15), S33, &H1FA27CF8)
        Call md5_HH(b, c, d, a, x(k + 2), S34, &HC4AC5665)

        Call md5_II(a, b, c, d, x(k + 0), S41, &HF4292244)
        Call md5_II(d, a, b, c, x(k + 7), S42, &H432AFF97)
        Call md5_II(c, d, a, b, x(k + 14), S43, &HAB9423A7)
        Call md5_II(b, c, d, a, x(k + 5), S44, &HFC93A039)
        Call md5_II(a, b, c, d, x(k + 12), S41, &H655B59C3)
        Call md5_II(d, a, b, c, x(k + 3), S42, &H8F0CCC92)
        Call md5_II(c, d, a, b, x(k + 10), S43, &HFFEFF47D)
        Call md5_II(b, c, d, a, x(k + 1), S44, &H85845DD1)
        Call md5_II(a, b, c, d, x(k + 8), S41, &H6FA87E4F)
        Call md5_II(d, a, b, c, x(k + 15), S42, &HFE2CE6E0)
        Call md5_II(c, d, a, b, x(k + 6), S43, &HA3014314)
        Call md5_II(b, c, d, a, x(k + 13), S44, &H4E0811A1)
        Call md5_II(a, b, c, d, x(k + 4), S41, &HF7537E82)
        Call md5_II(d, a, b, c, x(k + 11), S42, &HBD3AF235)
        Call md5_II(c, d, a, b, x(k + 2), S43, &H2AD7D2BB)
        Call md5_II(b, c, d, a, x(k + 9), S44, &HEB86D391)

        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        d = AddUnsigned(d, DD)
    Next k

    MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
    'MD5 = LCase(WordToHex(b) & WordToHex(c))  'I crop this to fit 16byte database password :D
End Function
