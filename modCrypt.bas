Attribute VB_Name = "modCrypt"
Option Explicit

' 加密解密模块

' AzDG 加密
Public Function AzDG_crypt(ByVal strSourceText As String, Optional ByVal strKey As String = "") As String
    Dim strRandMd5 As String
    Dim intCRCLength As Long
    Dim strTmp As String
    Dim intTextLength As Long
    Dim chrMd5 As String
    Dim i As Long

    Call Randomize

    strRandMd5 = MD5(CInt(Int((32000 * Rnd()))))
    intCRCLength = 0
    strTmp = ""
    intTextLength = Len(strSourceText)
    For i = 1 To intTextLength
        If (intCRCLength > 31) Then
            intCRCLength = 0
        End If
        chrMd5 = Mid(strRandMd5, intCRCLength + 1, 1)
        strTmp = strTmp & (chrMd5 & (Chr(Asc(Mid(strSourceText, i, 1)) Xor Asc(chrMd5))))
        intCRCLength = intCRCLength + 1
    Next
    
    AzDG_crypt = Base64EncodeString(AzDG_encode(strTmp, strKey))
End Function

' AzDG 解密
Public Function AzDG_decrypt(ByVal strSourceText As String, Optional ByVal strKey As String = "") As String
    Dim strTmp As String
    Dim intTextLength As Long
    Dim md5char As String
    Dim i As Long

    strSourceText = AzDG_encode(Base64DecodeString(strSourceText), strKey)
    strTmp = ""
    intTextLength = Len(strSourceText)
    For i = 1 To intTextLength
        md5char = Mid(strSourceText, i, 1)
        i = i + 1
        strTmp = strTmp & Chr(Asc(Mid(strSourceText, i, 1)) Xor Asc(md5char))
    Next
    
    AzDG_decrypt = strTmp
End Function

Private Function AzDG_encode(ByVal strSourceText As String, Optional ByVal strKey As String = "") As String
    Dim strKeyMd5 As String
    Dim intCRCLength As Long
    Dim strTmp As String
    Dim intTextLength As Long
    Dim i As Long
    
    If strKey = "" Then
        ' 全局密钥
        strKeyMd5 = MD5(gstrAzDGPrivateKey)
    Else
        ' 临时密钥
        strKeyMd5 = MD5(strKey)
    End If
    
    intCRCLength = 0
    strTmp = ""
    intTextLength = Len(strSourceText)
    
    For i = 1 To intTextLength
        If (intCRCLength > 31) Then
            intCRCLength = 0
        End If
        strTmp = strTmp & Chr(Asc(Mid(strSourceText, i, 1)) Xor Asc(Mid(strKeyMd5, intCRCLength + 1, 1)))
        intCRCLength = intCRCLength + 1
    Next

    AzDG_encode = strTmp
End Function

' Encipher the text using the pasword.
Public Function Cipher(ByVal Password As String, ByVal FromText As String) As String
Attribute Cipher.VB_UserMemId = 0
    Const MIN_ASC = 32  ' Space.
    Const MAX_ASC = 126 ' ~.
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1

    Dim Offset As Long
    Dim StrLen As Integer
    Dim i As Integer
    Dim ch As Integer
    Dim ToText As String

    ' Initialize the random number generator.
    Offset = NumericPassword(Password)
    Call Rnd(-1)
    Call Randomize(Offset)

    ToText = ""

    ' Encipher the string.
    StrLen = Len(FromText)
    For i = 1 To StrLen
        ch = Asc(Mid(FromText, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            Offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + Offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            ToText = ToText & Chr(ch)
        End If
    Next i
    Cipher = ToText
End Function

' Encipher the text using the pasword.
Public Function Decipher(ByVal Password As String, ByVal FromText As String) As String
    Const MIN_ASC = 32  ' Space.
    Const MAX_ASC = 126 ' ~.
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1

    Dim Offset As Long
    Dim StrLen As Integer
    Dim i As Integer
    Dim ch As Integer
    Dim ToText As String

    ' Initialize the random number generator.
    Offset = NumericPassword(Password)
    Call Rnd(-1)
    Call Randomize(Offset)

    ToText = ""

    ' Encipher the string.
    StrLen = Len(FromText)
    For i = 1 To StrLen
        ch = Asc(Mid(FromText, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            Offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - Offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            ToText = ToText & Chr(ch)
        End If
    Next i
    Decipher = ToText
End Function

' Translate a password into an offset value.
Private Function NumericPassword(ByVal Password As String) As Long
    Dim Value As Long
    Dim ch As Long
    Dim Shift1 As Long
    Dim Shift2 As Long
    Dim i As Integer
    Dim StrLen As Integer

    StrLen = Len(Password)
    For i = 1 To StrLen
        ' Add the next letter.
        ch = Asc(Mid(Password, i, 1))
        Value = Value Xor (ch * 2 ^ Shift1)
        Value = Value Xor (ch * 2 ^ Shift2)

        ' Change the shift offsets.
        Shift1 = (Shift1 + 7) Mod 19
        Shift2 = (Shift2 + 13) Mod 23
    Next i
    NumericPassword = Value
End Function
