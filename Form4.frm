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
   Begin VB.CommandButton Command1 
      Caption         =   "Form5"
      Height          =   516
      Left            =   504
      TabIndex        =   0
      Top             =   504
      Width           =   1608
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Command1_Click()
    Form5.Show
    Unload Me
End Sub

Private Sub Form_Load()
    pvTestAes
    Exit Sub
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

Private Sub pvTestAes()
    Dim uCtx            As CryptoAesContext
    Dim baKey()         As Byte
    Dim baNonce()       As Byte
    Dim baAad()         As Byte
    Dim baBlock()       As Byte
    Dim baBuffer()      As Byte
    Dim baAppend()      As Byte
    Dim baTag()         As Byte
    Dim uGcmCtx         As CryptoAesGcmContext
    
    baKey = FromHex("00112233445566778899aabbccddeeff")
    CryptoAesInit uCtx, baKey
    baBlock = baKey
    CryptoAesProcess uCtx, False, baBlock
    Debug.Assert ToHex(baBlock) = "b8f21a70bc9cee25249e2761fcbb7a34"
    '-> b8f21a70bc9cee25249e2761fcbb7a34
    CryptoAesProcess uCtx, True, baBlock
    Debug.Assert ToHex(baBlock) = "00112233445566778899aabbccddeeff"
    '-> 00112233445566778899aabbccddeeff

    '--- https://nvlpubs.nist.gov/nistpubs/Legacy/SP/nistspecialpublication800-38a.pdf
    '--- F.2.1 CBC-AES128.Encrypt
    baKey = FromHex("2b7e151628aed2a6abf7158809cf4f3c")
    baNonce = FromHex("000102030405060708090a0b0c0d0e0f")
    CryptoAesInit uCtx, baKey, baNonce
    baBuffer = FromHex("6bc1bee22e409f96e93d7e117393172a")
    CryptoAesCbcEncrypt uCtx, baBuffer, Final:=False
    Debug.Assert ToHex(baBuffer) = "7649abac8119b246cee98e9b12e9197d"
    '-> 7649abac8119b246cee98e9b12e9197d
    baAppend = vbNullString
    'baAppend = StrConv("0123", vbFromUnicode)
    CryptoAesCbcEncrypt uCtx, baAppend
    Debug.Assert ToHex(baAppend) = "8964e0b149c10b7b682e6e39aaeb731c"
    'Debug.Assert ToHex(baAppend) = "09fd79c936a0416df86153e8715da8c1"

    CryptoAesInit uCtx, baKey, baNonce
'    pvConcat baBuffer, baAppend
    If CryptoAesCbcDecrypt(uCtx, baBuffer, Final:=False) Then
        Debug.Assert ToHex(baBuffer) = "6bc1bee22e409f96e93d7e117393172a"
        '-> "00112233445566778899aabbccddeeff"
    Else
        Debug.Assert Len("CryptoAesCbcDecrypt failed") = 0
    End If
    If CryptoAesCbcDecrypt(uCtx, baAppend) Then
        Debug.Assert ToHex(baAppend) = ""
        'Debug.Assert ToHex(baAppend) = "30313233"
    Else
        Debug.Assert Len("CryptoAesCbcDecrypt failed") = 0
    End If

    '--- F.2.3 CBC-AES192.Encrypt
    baKey = FromHex("8e73b0f7da0e6452c810f32b809079e562f8ead2522c6b7b")
    baNonce = FromHex("000102030405060708090a0b0c0d0e0f")
    CryptoAesInit uCtx, baKey, baNonce
    baBuffer = FromHex("6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e5130c81c46a35ce411e5fbc1191a0a52eff69f2445df4f9b17ad2b417be66c3710")
    CryptoAesCbcEncrypt uCtx, baBuffer
    Debug.Assert ToHex(baBuffer) = "4f021db243bc633d7178183a9fa071e8b4d9ada9ad7dedf4e5e738763f69145a571b242012fb7ae07fa9baac3df102e008b0e27988598881d920a9e64f5615cd612ccd79224b350935d45dd6a98f8176"
    CryptoAesInit uCtx, baKey, baNonce
    CryptoAesCbcDecrypt uCtx, baBuffer
    Debug.Assert ToHex(baBuffer) = "6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e5130c81c46a35ce411e5fbc1191a0a52eff69f2445df4f9b17ad2b417be66c3710"

    '--- F.2.5 CBC-AES256.Encrypt
    baKey = FromHex("603deb1015ca71be2b73aef0857d77811f352c073b6108d72d9810a30914dff4")
    baNonce = FromHex("000102030405060708090a0b0c0d0e0f")
    CryptoAesInit uCtx, baKey, baNonce
    baBuffer = FromHex("6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e5130c81c46a35ce411e5fbc1191a0a52eff69f2445df4f9b17ad2b417be66c3710")
    CryptoAesCbcEncrypt uCtx, baBuffer
    Debug.Assert ToHex(baBuffer) = "f58c4c04d6e5f1ba779eabfb5f7bfbd69cfc4e967edb808d679f777bc6702c7d39f23369a9d9bacfa530e26304231461b2eb05e2c39be9fcda6c19078c6a9d1b3f461796d6b0d6b2e0c2a72b4d80e644"
    CryptoAesInit uCtx, baKey, baNonce
    Debug.Assert CryptoAesCbcDecrypt(uCtx, baBuffer)
    Debug.Assert ToHex(baBuffer) = "6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e5130c81c46a35ce411e5fbc1191a0a52eff69f2445df4f9b17ad2b417be66c3710"

    '--- F.5.1 CTR-AES128.Encrypt
    baKey = FromHex("2b7e151628aed2a6abf7158809cf4f3c")
    baNonce = FromHex("f0f1f2f3f4f5f6f7f8f9fafbfcfdfeff")
    CryptoAesInit uCtx, baKey, baNonce
    baBuffer = FromHex("6bc1bee22e409f96e93d7e117393172a")
    CryptoAesCtrCrypt uCtx, baBuffer
    Debug.Assert ToHex(baBuffer) = "874d6191b620e3261bef6864990db6ce"
    CryptoAesInit uCtx, baKey, baNonce
    CryptoAesCtrCrypt uCtx, baBuffer
    Debug.Assert ToHex(baBuffer) = "6bc1bee22e409f96e93d7e117393172a"

    '--- F.5.5 CTR-AES256.Encrypt
    baKey = FromHex("603deb1015ca71be2b73aef0857d77811f352c073b6108d72d9810a30914dff4")
    baNonce = FromHex("f0f1f2f3f4f5f6f7f8f9fafbfcfdfeff")
    CryptoAesInit uCtx, baKey, baNonce
    baBuffer = FromHex("6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e5130c81c46a35ce411e5fbc1191a0a52eff69f2445df4f9b17ad2b417be66c3710")
    CryptoAesCtrCrypt uCtx, baBuffer
    Debug.Assert ToHex(baBuffer) = "601ec313775789a5b7a7f504bbf3d228f443e3ca4d62b59aca84e990cacaf5c52b0930daa23de94ce87017ba2d84988ddfc9c58db67aada613c2dd08457941a6"
    CryptoAesInit uCtx, baKey, baNonce
    CryptoAesCtrCrypt uCtx, baBuffer
    Debug.Assert ToHex(baBuffer) = "6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e5130c81c46a35ce411e5fbc1191a0a52eff69f2445df4f9b17ad2b417be66c3710"

    '--- https://csrc.nist.rip/groups/ST/toolkit/BCM/documents/proposedmodes/gcm/gcm-spec.pdf
    '--- Test Case 1
    baKey = FromHex("00000000000000000000000000000000")
    baNonce = FromHex("000000000000000000000000")
    baAad = vbNullString
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    baBuffer = vbNullString
    CryptoAesGcmEncrypt uGcmCtx, baBuffer, TagSize:=16, Tag:=baTag
    Debug.Assert ToHex(baTag) = "58e2fccefa7e3061367f1d57a4e7455a"
    '-> 58e2fccefa7e3061367f1d57a4e7455a

    '--- Test Case 2
    baKey = FromHex("00000000000000000000000000000000")
    baNonce = FromHex("000000000000000000000000")
    baAad = vbNullString
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    baBuffer = FromHex("00000000000000000000000000000000")
    CryptoAesGcmEncrypt uGcmCtx, baBuffer, TagSize:=16, Tag:=baTag
    Debug.Assert ToHex(baBuffer) = "0388dace60b6a392f328c2b971b2fe78"
    Debug.Assert ToHex(baTag) = "ab6e47d42cec13bdf53a67b21257bddf"
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    CryptoAesGcmDecrypt uGcmCtx, baBuffer, Tag:=baTag
    Debug.Assert ToHex(baBuffer) = "00000000000000000000000000000000"

    '--- Test Case 3
    baKey = FromHex("feffe9928665731c6d6a8f9467308308")
    baNonce = FromHex("cafebabefacedbaddecaf888")
    baAad = vbNullString
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    baBuffer = FromHex("d9313225f88406e5a55909c5aff5269a86a7a9531534f7da2e4c303d8a318a721c3c0c95956809532fcf0e2449a6b525b16aedf5aa0de657ba637b391aafd255")
    CryptoAesGcmEncrypt uGcmCtx, baBuffer, TagSize:=16, Tag:=baTag
    Debug.Assert ToHex(baBuffer) = "42831ec2217774244b7221b784d0d49ce3aa212f2c02a4e035c17e2329aca12e21d514b25466931c7d8f6a5aac84aa051ba30b396a0aac973d58e091473f5985"
    Debug.Assert ToHex(baTag) = "4d5c2af327cd64a62cf35abd2ba6fab4"
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    CryptoAesGcmDecrypt uGcmCtx, baBuffer, Tag:=baTag
    Debug.Assert ToHex(baBuffer) = "d9313225f88406e5a55909c5aff5269a86a7a9531534f7da2e4c303d8a318a721c3c0c95956809532fcf0e2449a6b525b16aedf5aa0de657ba637b391aafd255"

    '--- Test Case 4
    baKey = FromHex("feffe9928665731c6d6a8f9467308308")
    baNonce = FromHex("cafebabefacedbaddecaf888")
    baAad = FromHex("feedfacedeadbeeffeedfacedeadbeefabaddad2")
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    baBuffer = FromHex("d9313225f88406e5a55909c5aff5269a86a7a9531534f7da2e4c303d8a318a721c3c0c95956809532fcf0e2449a6b525b16aedf5aa0de657ba637b39")
    CryptoAesGcmEncrypt uGcmCtx, baBuffer, TagSize:=16, Tag:=baTag
    Debug.Assert ToHex(baBuffer) = "42831ec2217774244b7221b784d0d49ce3aa212f2c02a4e035c17e2329aca12e21d514b25466931c7d8f6a5aac84aa051ba30b396a0aac973d58e091"
    Debug.Assert ToHex(baTag) = "5bc94fbc3221a5db94fae95ae7121a47"
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    CryptoAesGcmDecrypt uGcmCtx, baBuffer, Tag:=baTag
    Debug.Assert ToHex(baBuffer) = "d9313225f88406e5a55909c5aff5269a86a7a9531534f7da2e4c303d8a318a721c3c0c95956809532fcf0e2449a6b525b16aedf5aa0de657ba637b39"

    '--- Test Case 5
    baKey = FromHex("feffe9928665731c6d6a8f9467308308")
    baNonce = FromHex("cafebabefacedbad")
    baAad = FromHex("feedfacedeadbeeffeedfacedeadbeefabaddad2")
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    baBuffer = FromHex("d9313225f88406e5a55909c5aff5269a86a7a9531534f7da2e4c303d8a318a721c3c0c95956809532fcf0e2449a6b525b16aedf5aa0de657ba637b39")
    CryptoAesGcmEncrypt uGcmCtx, baBuffer, TagSize:=16, Tag:=baTag
    Debug.Assert ToHex(baBuffer) = "61353b4c2806934a777ff51fa22a4755699b2a714fcdc6f83766e5f97b6c742373806900e49f24b22b097544d4896b424989b5e1ebac0f07c23f4598"
    Debug.Assert ToHex(baTag) = "3612d2e79e3b0785561be14aaca2fccb"
    CryptoAesGcmInit uGcmCtx, baKey, baNonce, baAad
    CryptoAesGcmDecrypt uGcmCtx, baBuffer, Tag:=baTag
    Debug.Assert ToHex(baBuffer) = "d9313225f88406e5a55909c5aff5269a86a7a9531534f7da2e4c303d8a318a721c3c0c95956809532fcf0e2449a6b525b16aedf5aa0de657ba637b39"
End Sub

Private Sub pvConcat(baBuffer() As Byte, baAppend() As Byte)
    Dim lPos As Long
    lPos = UBound(baBuffer) + 1
    ReDim Preserve baBuffer(0 To lPos + UBound(baAppend)) As Byte
    Call CopyMemory(baBuffer(lPos), baAppend(0), UBound(baAppend) + 1)
End Sub
