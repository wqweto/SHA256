VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form4"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    pvTestMd5
    pvTestSha1
    pvTestBase64
    pvTestChaCha20
    pvTestPoly1305
    pvTestChaCha20Poly1305 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\chacha20_poly1305_test.json"))
End Sub

Private Sub pvTestSha1()
    Debug.Print ToHex(CryptoSha1ByteArray(StrConv("abc", vbFromUnicode)))
    '-> a9993e364706816aba3e25717850c26c9cd0d89d
    Debug.Assert CryptoSha1Text("abc") = "a9993e364706816aba3e25717850c26c9cd0d89d"
    Debug.Print ToHex(CryptoSha1ByteArray(StrConv(vbNullString, vbFromUnicode)))
    '-> da39a3ee5e6b4b0d3255bfef95601890afd80709
End Sub

Private Sub pvTestBase64()
    Dim baInput()       As Byte
    Dim sOutput         As String
    
    baInput = "This is a test"
    sOutput = ToBase64Array(baInput)
    Debug.Print FromBase64Array(sOutput)
End Sub

Private Sub pvTestMd5()
    Debug.Print ToHex(CryptoMd5ByteArray(StrConv(vbNullString, vbFromUnicode)))
    '-> d41d8cd98f00b204e9800998ecf8427e
    Debug.Print ToHex(CryptoMd5ByteArray(StrConv("a", vbFromUnicode)))
    '-> 0cc175b9c0f1b6a831c399e269772661
    Debug.Assert CryptoMd5Text("a") = "0cc175b9c0f1b6a831c399e269772661"
    Debug.Print ToHex(CryptoMd5ByteArray(StrConv("12345678901234567890123456789012345678901234567890123456789012345678901234567890", vbFromUnicode)))
    '-> 57edf4a22be3c955ac49da2e2107b67a
    Debug.Print ToHex(CryptoMd5ByteArray(StrConv("1234567890123456789012345678901234567890123456789012345", vbFromUnicode)))
    '-> c9ccf168914a1bcfc3229f1948e67da0
End Sub

Private Sub pvTestChaCha20()
    Dim baKey()         As Byte
    Dim baNonce()       As Byte
    Dim uCtx            As CryptoChaCha20Context
    Dim baText()        As Byte
    
    baKey = FromHex("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f")
    baNonce = FromHex("0001020304050607")
    ReDim baText(0 To 249)
    CryptoChaCha20Init uCtx, baKey, baNonce
    CryptoChaCha20Cipher uCtx, baText
    Debug.Assert ToHex(baText) = "f798a189f195e66982105ffb640bb7757f579da31602fc93ec01ac56f85ac3c134a4547b733b46413042c9440049176905d3be59ea1c53f15916155c2be8241a38008b9a26bc35941e2444177c8ade6689de95264986d95889fb60e84629c9bd9a5acb1cc118be563eb9b3a4a472f82e09a7e778492b562ef7130e88dfe031c79db9d4f7c7a899151b9a475032b63fc385245fe054e3dd5a97a5f576fe064025d3ce042c566ab2c507b138db853e3d6959660996546cc9c4a6eafdc777c040d70eaf46f76dad3979e5c5360c3317166a1c894c94a371876a94df7628fe4eaaf2ccb27d5aaae0ad7ad0f9d4b6ad3b54098746d4524d38407a6deb"
    '-> f798a189f195e66982105ffb640bb7757f579da31602fc93ec01ac56f85ac3c134a4547b733b46413042c9440049176905d3be59ea1c53f15916155c2be8241a38008b9a26bc35941e2444177c8ade6689de95264986d95889fb60e84629c9bd9a5acb1cc118be563eb9b3a4a472f82e09a7e778492b562ef7130e88dfe031c79db9d4f7c7a899151b9a475032b63fc385245fe054e3dd5a97a5f576fe064025d3ce042c566ab2c507b138db853e3d6959660996546cc9c4a6eafdc777c040d70eaf46f76dad3979e5c5360c3317166a1c894c94a371876a94df7628fe4eaaf2ccb27d5aaae0ad7ad0f9d4b6ad3b54098746d4524d38407a6deb
End Sub

Private Sub pvTestPoly1305()
    Dim uCtx            As CryptoPoly1305Context
    Dim baKey()         As Byte
    Dim baInput()       As Byte
    Dim baOuput()       As Byte
    
    baKey = FromHex("85:d6:be:78:57:55:6d:33:7f:44:52:fe:42:d5:06:a8:01:03:80:8a:fb:0d:b2:fd:4a:bf:f6:af:41:49:f5:1b")
    baInput = StrConv("Cryptographic Forum Research Group", vbFromUnicode)
    CryptoPoly1305Init uCtx, baKey
    CryptoPoly1305Update uCtx, baInput
    CryptoPoly1305Finalize uCtx, baOuput
    Debug.Print ToHex(baOuput)
    '-> a8061dc1305136c6c22b8baf0c0127a9
End Sub

Private Function pvArrayEqual(baFirst() As Byte, baSecond() As Byte) As Boolean
    If UBound(baFirst) = UBound(baSecond) Then
        If UBound(baFirst) < 0 Then
            pvArrayEqual = True
        Else
            pvArrayEqual = (InStrB(baFirst, baSecond) = 1)
        End If
    End If
End Function

Private Sub pvTestChaCha20Poly1305(oJson As Object)
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim baKey()         As Byte
    Dim baIV()          As Byte
    Dim baAad()         As Byte
    Dim baBuffer()      As Byte
    Dim baEncr()        As Byte
    Dim baDecr()        As Byte
    Dim baTag()         As Byte
    Dim baCheckTag()    As Byte
    Dim sResult         As String
    Dim sComment        As String
    
    For Each oGroup In JsonValue(oJson, "testGroups")
        For Each oTest In JsonValue(oGroup, "tests")
            baKey = FromHex(JsonValue(oTest, "key"))
            baIV = FromHex(JsonValue(oTest, "iv"))
            baAad = FromHex(JsonValue(oTest, "aad"))
            baBuffer = FromHex(JsonValue(oTest, "msg"))
            baEncr = FromHex(JsonValue(oTest, "ct"))
            baCheckTag = FromHex(JsonValue(oTest, "tag"))
            sResult = "valid"
            If Not CryptoChaCha20Poly1305Encrypt(baKey, baTag, baBuffer, Nonce:=baIV, AssociatedData:=baAad) Then
                sResult = "invalid"
            End If
            If Not pvArrayEqual(baBuffer, baEncr) Then
                sResult = "invalid"
            End If
            If Not pvArrayEqual(baCheckTag, baTag) Then
                sResult = "invalid"
            End If
            If JsonValue(oTest, "result") <> sResult Then
                sComment = JsonValue(oTest, "comment")
                If IsObject(JsonValue(oTest, "flags")) Then
                    sComment = sComment & ", " & Join(JsonValue(oTest, "flags/*"), ", ")
                End If
                Debug.Print "[-] " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString)
            End If
            baBuffer = FromHex(JsonValue(oTest, "ct"))
            baDecr = FromHex(JsonValue(oTest, "msg"))
            baTag = FromHex(JsonValue(oTest, "tag"))
            sResult = "valid"
            If Not CryptoChaCha20Poly1305Decrypt(baKey, baTag, baBuffer, Nonce:=baIV, AssociatedData:=baAad) Then
                sResult = "invalid"
            End If
            If Not pvArrayEqual(baBuffer, baDecr) Then
                sResult = "invalid"
            End If
            If JsonValue(oTest, "result") <> sResult Then
                sComment = JsonValue(oTest, "comment")
                If IsObject(JsonValue(oTest, "flags")) Then
                    sComment = sComment & ", " & Join(JsonValue(oTest, "flags/*"), ", ")
                End If
                Debug.Print "[-] " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString)
            End If
        Next
    Next
End Sub
