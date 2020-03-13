Attribute VB_Name = "MD5"
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: MD5.bas
Rem
Rem Description:
Rem     Contains code to perform basic MD5 hashing of a string.  Lion
Rem     share of code gotten from Internet.
Rem
Rem -------------------------------------------------------------

Dim w1 As String, w2 As String, w3 As String, w4 As String

Function MD5F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal Xin As String, ByVal qdata As String, ByVal rots As Integer)
    MD5F = BigMod32Add(RotLeft(BigMod32Add(BigMod32Add(w, tempstr), BigMod32Add(Xin, qdata)), rots), X)
End Function

Sub MD5F1(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal Xin As String, ByVal qdata As String, ByVal rots As Integer)
Dim tempstr As String

    tempstr = BigXOR(z, BigAND(X, BigXOR(Y, z)))
    w = MD5F(tempstr, w, X, Y, z, Xin, qdata, rots)
End Sub

Sub MD5F2(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal Xin As String, ByVal qdata As String, ByVal rots As Integer)
Dim tempstr As String

    tempstr = BigXOR(Y, BigAND(z, BigXOR(X, Y)))
    w = MD5F(tempstr, w, X, Y, z, Xin, qdata, rots)
End Sub

Sub MD5F3(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal Xin As String, ByVal qdata As String, ByVal rots As Integer)
Dim tempstr As String

    tempstr = BigXOR(X, BigXOR(Y, z))
    w = MD5F(tempstr, w, X, Y, z, Xin, qdata, rots)
End Sub

Sub MD5F4(w As String, ByVal X As String, ByVal Y As String, ByVal z As String, ByVal Xin As String, ByVal qdata As String, ByVal rots As Integer)
Dim tempstr As String

    tempstr = BigXOR(Y, BigOR(X, BigNOT(z)))
    w = MD5F(tempstr, w, X, Y, z, Xin, qdata, rots)
End Sub

Function MD5_Calc(ByVal hashthis As String) As String
ReDim buf(0 To 3) As String
ReDim Xin(0 To 15) As String
Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
Dim a As String, b As String, c As String, d As String

    ' Add padding

    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64

    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If

    hashthis = hashthis + String$(tempnum2, Chr$(0))

    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit

    

    ' Set magic numbers
    buf(0) = "67452301"
    buf(1) = "efcdab89"
    buf(2) = "98badcfe"
    buf(3) = "10325476"

    

    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = buf(0)
        b = buf(1)
        c = buf(2)
        d = buf(3)

        ' Get the 512 bits
        For loopit = 0 To 15
            Xin(loopit) = ""
            For loopinner = 1 To 4
                Xin(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + Xin(loopit)
                If Len(Xin(loopit)) Mod 2 Then Xin(loopit) = "0" + Xin(loopit)
            Next loopinner
        Next loopit

        ' Round 1
        MD5F1 a, b, c, d, Xin(0), "d76aa478", 7
        MD5F1 d, a, b, c, Xin(1), "e8c7b756", 12
        MD5F1 c, d, a, b, Xin(2), "242070db", 17
        MD5F1 b, c, d, a, Xin(3), "c1bdceee", 22
        MD5F1 a, b, c, d, Xin(4), "f57c0faf", 7
        MD5F1 d, a, b, c, Xin(5), "4787c62a", 12
        MD5F1 c, d, a, b, Xin(6), "a8304613", 17
        MD5F1 b, c, d, a, Xin(7), "fd469501", 22
        MD5F1 a, b, c, d, Xin(8), "698098d8", 7
        MD5F1 d, a, b, c, Xin(9), "8b44f7af", 12
        MD5F1 c, d, a, b, Xin(10), "ffff5bb1", 17
        MD5F1 b, c, d, a, Xin(11), "895cd7be", 22
        MD5F1 a, b, c, d, Xin(12), "6b901122", 7
        MD5F1 d, a, b, c, Xin(13), "fd987193", 12
        MD5F1 c, d, a, b, Xin(14), "a679438e", 17
        MD5F1 b, c, d, a, Xin(15), "49b40821", 22

        ' Round 2
        MD5F2 a, b, c, d, Xin(1), "f61e2562", 5
        MD5F2 d, a, b, c, Xin(6), "c040b340", 9
        MD5F2 c, d, a, b, Xin(11), "265e5a51", 14
        MD5F2 b, c, d, a, Xin(0), "e9b6c7aa", 20
        MD5F2 a, b, c, d, Xin(5), "d62f105d", 5
        MD5F2 d, a, b, c, Xin(10), "02441453", 9
        MD5F2 c, d, a, b, Xin(15), "d8a1e681", 14
        MD5F2 b, c, d, a, Xin(4), "e7d3fbc8", 20
        MD5F2 a, b, c, d, Xin(9), "21e1cde6", 5
        MD5F2 d, a, b, c, Xin(14), "c33707d6", 9
        MD5F2 c, d, a, b, Xin(3), "f4d50d87", 14
        MD5F2 b, c, d, a, Xin(8), "455a14ed", 20
        MD5F2 a, b, c, d, Xin(13), "a9e3e905", 5
        MD5F2 d, a, b, c, Xin(2), "fcefa3f8", 9
        MD5F2 c, d, a, b, Xin(7), "676f02d9", 14
        MD5F2 b, c, d, a, Xin(12), "8d2a4c8a", 20

        ' Round 3
        MD5F3 a, b, c, d, Xin(5), "fffa3942", 4
        MD5F3 d, a, b, c, Xin(8), "8771f681", 11
        MD5F3 c, d, a, b, Xin(11), "6d9d6122", 16
        MD5F3 b, c, d, a, Xin(14), "fde5380c", 23
        MD5F3 a, b, c, d, Xin(1), "a4beea44", 4
        MD5F3 d, a, b, c, Xin(4), "4bdecfa9", 11
        MD5F3 c, d, a, b, Xin(7), "f6bb4b60", 16
        MD5F3 b, c, d, a, Xin(10), "bebfbc70", 23
        MD5F3 a, b, c, d, Xin(13), "289b7ec6", 4
        MD5F3 d, a, b, c, Xin(0), "e27fa", 11
        MD5F3 c, d, a, b, Xin(3), "d4ef3085", 16
        MD5F3 b, c, d, a, Xin(6), "04881d05", 23
        MD5F3 a, b, c, d, Xin(9), "d9d4d039", 4
        MD5F3 d, a, b, c, Xin(12), "e6db99e5", 11
        MD5F3 c, d, a, b, Xin(15), "1fa27cf8", 16
        MD5F3 b, c, d, a, Xin(2), "c4ac5665", 23

        ' Round 4
        MD5F4 a, b, c, d, Xin(0), "f4292244", 6
        MD5F4 d, a, b, c, Xin(7), "432aff97", 10
        MD5F4 c, d, a, b, Xin(14), "ab9423a7", 15
        MD5F4 b, c, d, a, Xin(5), "fc93a039", 21
        MD5F4 a, b, c, d, Xin(12), "655b59c3", 6
        MD5F4 d, a, b, c, Xin(3), "8f0ccc92", 10
        MD5F4 c, d, a, b, Xin(10), "ffeff47d", 15
        MD5F4 b, c, d, a, Xin(1), "85845dd1", 21
        MD5F4 a, b, c, d, Xin(8), "6fa87e4f", 6
        MD5F4 d, a, b, c, Xin(15), "fe2ce6e0", 10
        MD5F4 c, d, a, b, Xin(6), "a3014314", 15
        MD5F4 b, c, d, a, Xin(13), "4e0811a1", 21
        MD5F4 a, b, c, d, Xin(4), "f7537e82", 6
        MD5F4 d, a, b, c, Xin(11), "bd3af235", 10
        MD5F4 c, d, a, b, Xin(2), "2ad7d2bb", 15
        MD5F4 b, c, d, a, Xin(9), "eb86d391", 21

        buf(0) = BigAdd(buf(0), a)
        buf(1) = BigAdd(buf(1), b)
        buf(2) = BigAdd(buf(2), c)
        buf(3) = BigAdd(buf(3), d)
    Next loopouter

    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Chr(Val("&H" + Mid$(buf(loopit), 1 + 2 * loopinner, 2)))
        Next loopinner
    Next loopit

    ' And return it
    MD5_Calc = hashthis
End Function

Function BigMod32Add(ByVal value1 As String, ByVal value2 As String) As String
    BigMod32Add = Right$(BigAdd(value1, value2), 8)
End Function

Public Function BigAdd(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value1 = Space$(Abs(tempnum)) + value1
    ElseIf tempnum > 0 Then
        value2 = Space$(Abs(tempnum)) + value2
    End If

    tempnum = 0
    For loopit = Len(value1) To 1 Step -1
        tempnum = tempnum + Val("&H" + Mid$(value1, loopit, 1)) + Val("&H" + Mid$(value2, loopit, 1))
        valueans = Hex$(tempnum Mod 16) + valueans
        tempnum = Int(tempnum / 16)
    Next loopit

    If tempnum <> 0 Then
        valueans = Hex$(tempnum) + valueans
    End If

    BigAdd = Right(valueans, 8)
End Function

Public Function RotLeft(ByVal value1 As String, ByVal rots As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    rots = rots Mod 32
    
    If rots = 0 Then
        RotLeft = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    tempstr = Mid$(value1, rots + 1) + Left$(value1, rots)

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    RotLeft = Right(value1, 8)
End Function

Function BigAND(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) And Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigAND = valueans
End Function

Function BigNOT(ByVal value1 As String) As String
Dim valueans As String
Dim loopit As Integer

    value1 = Right$(value1, 8)
    value1 = String$(8 - Len(value1), "0") + value1
    For loopit = 1 To 8
        valueans = valueans + Hex$(15 Xor Val("&H" + Mid$(value1, loopit, 1)))
    Next loopit

    BigNOT = valueans
End Function

Function BigOR(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Or Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigOR = valueans
End Function

Function BigXOR(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Xor Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigXOR = Right(valueans, 8)
End Function

Public Function GetKey(strName As String) As String
    Dim s1 As String, s2 As String
    Dim i As Integer
    s1 = MD5_Calc(strName & "NathanBrewerBlinnLovesDaddyBunches")
    s2 = ""
    For i = 1 To Len(s1)
        s2 = s2 & Hex(Asc(Mid(s1, i, 1)))
    Next
    GetKey = s2
End Function

Public Function ValidateKey(strName As String, strKey As String) As Boolean
    If strKey = GetKey(strName) Then
        ValidateKey = True
    Else
        ValidateKey = False
    End If
End Function
