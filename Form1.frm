VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pvTestSHA256Managed()
    Dim baInput()       As Byte
    Dim baHash()        As Byte
    
    baInput = FromHex("238dbd3bd53ee4c53c505ca2b56e1756e622aa0c")
    With CreateObject("System.Security.Cryptography.SHA256Managed")
        baHash = .ComputeHash_2(baInput)
        baHash = .ComputeHash_2(baHash)
    End With
    Debug.Print Left$(ToHex(baHash), 8)
End Sub

Private Sub pvTestCAPI()
    Const CALG_SHA_256  As Long = &H800C&
    Dim baInput()       As Byte
    Dim baHash()        As Byte
    
    baInput = FromHex("238dbd3bd53ee4c53c505ca2b56e1756e622aa0c")
    CryptoHash baHash, CALG_SHA_256, baInput
    CryptoHash baHash, CALG_SHA_256, baHash
    Debug.Print Left$(ToHex(baHash), 8)
End Sub

Private Function SimplSHA256(sMessage As String) As String
    Dim baInput()       As Byte
    Dim baHash()        As Byte
    
    With CreateObject("System.Security.Cryptography.SHA256Managed")
        baInput = StrConv(sMessage, vbFromUnicode)
        baHash = .ComputeHash_2(baInput)
        SimplSHA256 = LCase(ToHex(baHash))
    End With
End Function

Private Sub Form_Click1()
    Dim baInput()       As Byte
    Dim baHash()        As Byte
    
    baInput = FromHex("238dbd3bd53ee4c53c505ca2b56e1756e622aa0c")
    baHash = CryptoSha2ByteArray(256, baInput)
    baHash = CryptoSha2ByteArray(256, baHash)
    Debug.Print Left$(ToHex(baHash), 8)
End Sub

Private Sub Form_Click()
    pvTestX25519 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\x25519_test.json"))

    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha512_test.json")), 512
    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha384_test.json")), 384

    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha256_test.json")), 256
    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha224_test.json")), 224
    pvTestHmacSha3 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha3_224_test.json")), 224
    pvTestHmacSha3 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha3_256_test.json")), 256
    pvTestHmacSha3 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha3_384_test.json")), 384
    pvTestHmacSha3 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha3_512_test.json")), 512
End Sub

Private Sub pvTestX25519(oJson As Object)
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim baPub()         As Byte
    Dim baPriv()        As Byte
    Dim baShared()      As Byte
    Dim baBuffer()      As Byte
    Dim sResult         As String
    Dim sComment        As String
    
    For Each oGroup In JsonValue(oJson, "testGroups")
        For Each oTest In JsonValue(oGroup, "tests")
            baPub = FromHex(JsonValue(oTest, "public"))
            baPriv = FromHex(JsonValue(oTest, "private"))
            baShared = FromHex(JsonValue(oTest, "shared"))
            CryptoX25519SharedSecret baBuffer, baPriv, baPub
            sResult = "valid"
            If Not pvArrayEqual(baBuffer, baShared) Then
                sResult = "invalid"
            End If
            If JsonValue(oTest, "result") = "acceptable" And sResult = "valid" Then
                '--- do nothing
            ElseIf JsonValue(oTest, "result") <> sResult Then
                sComment = JsonValue(oTest, "comment")
                If IsObject(JsonValue(oTest, "flags")) Then
                    sComment = sComment & ", " & Join(JsonValue(oTest, "flags/*"), ", ")
                End If
                Debug.Print "[-] " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString)
            End If
        Next
    Next
End Sub

Private Sub pvTestHmacSha2(oJson As Object, ByVal lBitSize As Long)
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim baKey()         As Byte
    Dim baMsg()         As Byte
    Dim baTag()         As Byte
    Dim baBuffer()      As Byte
    Dim sResult         As String
    Dim sComment        As String
    
    For Each oGroup In JsonValue(oJson, "testGroups")
        For Each oTest In JsonValue(oGroup, "tests")
            baKey = FromHex(JsonValue(oTest, "key"))
            baMsg = FromHex(JsonValue(oTest, "msg"))
            baTag = FromHex(JsonValue(oTest, "tag"))
            baBuffer = CryptoHmacSha2ByteArray(lBitSize, baKey, baMsg)
            ReDim Preserve baBuffer(0 To UBound(baTag))
            sResult = "valid"
            If Not pvArrayEqual(baBuffer, baTag) Then
                sResult = "invalid"
            End If
            If JsonValue(oTest, "result") <> sResult Then
                sComment = JsonValue(oTest, "comment")
                Debug.Print "[-] " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString)
            End If
        Next
    Next
End Sub

Private Sub pvTestHmacSha3(oJson As Object, ByVal lBitSize As Long)
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim baKey()         As Byte
    Dim baMsg()         As Byte
    Dim baTag()         As Byte
    Dim baBuffer()      As Byte
    Dim sResult         As String
    Dim sComment        As String
    
    For Each oGroup In JsonValue(oJson, "testGroups")
        For Each oTest In JsonValue(oGroup, "tests")
            baKey = FromHex(JsonValue(oTest, "key"))
            baMsg = FromHex(JsonValue(oTest, "msg"))
            baTag = FromHex(JsonValue(oTest, "tag"))
            baBuffer = CryptoHmacSha3ByteArray(lBitSize, baKey, baMsg)
            ReDim Preserve baBuffer(0 To UBound(baTag))
            sResult = "valid"
            If Not pvArrayEqual(baBuffer, baTag) Then
                sResult = "invalid"
            End If
            If JsonValue(oTest, "result") <> sResult Then
                sComment = JsonValue(oTest, "comment")
                Debug.Print "[-] " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString)
            End If
        Next
    Next
End Sub

Private Function pvArrayEqual(baFirst() As Byte, baSecond() As Byte) As Boolean
    If UBound(baFirst) = UBound(baSecond) Then
        pvArrayEqual = (InStrB(baFirst, baSecond) = 1)
    End If
End Function

Private Sub pvTestSha3()
    Dim sMessage        As String
    Dim baInput()       As Byte
    Dim baHash()        As Byte

'    sMessage = "Rosetta code"
'    baHash = CryptoSha2ByteArray(256, StrConv(sMessage, vbFromUnicode))
'    Debug.Assert LCase$(ToHex(baHash)) = "764faf5c61ac315f1497f9dfa542713965b785e5cc2f707d6468d7d1124cdfcf"
'
'    sMessage = Replace(String(1000, " "), " ", "Rosetta code")
'    baHash = CryptoSha2ByteArray(256, StrConv(sMessage, vbFromUnicode))
'    With New clsSHA256
'        Debug.Assert .SHA256(sMessage) = LCase$(ToHex(baHash))
'    End With
'    Debug.Assert SimplSHA256(sMessage) = LCase$(ToHex(baHash))
    
    baInput = StrConv("abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", vbFromUnicode)
    baHash = CryptoSha3ByteArray(224, baInput)
    Debug.Assert ToHex(baHash) = "543e6868e1666c1a643630df77367ae5a62a85070a51c14cbf665cbc"
    
    baHash = CryptoSha3ByteArray(256, baInput)
    Debug.Assert ToHex(baHash) = "916f6061fe879741ca6469b43971dfdb28b1a32dc36cb3254e812be27aad1d18"
    Debug.Assert CryptoSha3Text(256, "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu") = "916f6061fe879741ca6469b43971dfdb28b1a32dc36cb3254e812be27aad1d18"
    
    baHash = CryptoSha3ByteArray(384, baInput)
    Debug.Assert ToHex(baHash) = "79407d3b5916b59c3e30b09822974791c313fb9ecc849e406f23592d04f625dc8c709b98b43b3852b337216179aa7fc7"
    
    baHash = CryptoSha3ByteArray(512, baInput)
    Debug.Assert ToHex(baHash) = "afebb2ef542e6579c50cad06d2e578f9f8dd6881d7dc824d26360feebf18a4fa73e3261122948efcfd492e74e82e2189ed0fb440d187f382270cb455f21dd185"
    
    baHash = CryptoShakeByteArray(128, 32, baInput, Size:=0)
    Debug.Assert ToHex(baHash) = "7f9c2ba4e88f827d616045507605853ed73b8093f6efbc88eb1a6eacfa66ef26"
    
    baHash = CryptoShakeByteArray(256, 64, baInput, Size:=0)
    Debug.Assert ToHex(baHash) = "46b9dd2b0ba88d13233b3feb743eeb243fcd52ea62b81b82b50c27646ed5762fd75dc4ddd8c0f200cb05019d67b592f6fc821c49479ab48640292eacb3b7c4be"
End Sub

Private Sub pvTestSha512()
    Dim baInput()       As Byte
    Dim baHash()        As Byte
    
    baInput = StrConv("abc", vbFromUnicode)
    baHash = CryptoSha512ByteArray(512, baInput)
    Debug.Print ToHex(baHash)
    '-> ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f
    
    baHash = CryptoSha512ByteArray(384, baInput)
    Debug.Print ToHex(baHash)
    '-> cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7
    
    baHash = CryptoSha2ByteArray(512256, baInput)
    Debug.Print ToHex(baHash)
    '-> 53048e2681941ef99b2e29b76b4c7dabe4c2d0c634fc6d46e0e2f13107e7af23
    
    baHash = CryptoSha2ByteArray(512224, baInput)
    Debug.Print ToHex(baHash)
    '-> 4634270f707b6a54daae7530460842e20e37ed265ceee9a43e8924aa
    
    Debug.Print CryptoSha512Text(512, "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu")
    '-> "8e959b75dae313da8cf4f72814fc143f8f7779c6eb9f7fa17299aeadb6889018501d289e4900f7e4331b99dec4b5433ac7d329eeb6dd26545e96e55b874be909"
End Sub

Private Sub pvTestCryptoX25519()
    Dim dblTimer        As Double
    Dim lIdx            As Long
    Dim baPriv1()       As Byte
    Dim baPub1()        As Byte
    Dim baPriv2()       As Byte
    Dim baPub2()        As Byte
    Dim baShared1()     As Byte
    Dim baShared2()     As Byte
    
    dblTimer = Timer
    For lIdx = 1 To 1
        baPriv1 = FromHex("70076D0A7318A57D3C16C17251B26645DF4C2F87EBC0992AB177FBA51DB92C6A")
        CryptoX25519PublicKey baPub1, baPriv1
        Debug.Print ToHex(baPub1) ' 8520f0098930a754748b7ddcb43ef75a0dbf3a0d26381af4eba4a98eaa9b4e6a
        
        baPriv2 = FromHex("58AB087E624A8A4B79E17F8B83800EE66F3BB1292618B6FD1C2F8B27FF88E06B")
        CryptoX25519PublicKey baPub2, baPriv2
        Debug.Print ToHex(baPub2) ' de9edb7d7b7dc1b4d35b61c2ece435373f8343c85b78674dadfc7e146f882b4f
        
        CryptoX25519SharedSecret baShared1, baPriv1, baPub2
        Debug.Print ToHex(baShared1) ' 4a5d9d5ba4ce2de1728e3bf480350f25e07e21c947d19e3376f09b3c1e161742
        
        CryptoX25519SharedSecret baShared2, baPriv2, baPub1
        Debug.Print ToHex(baShared2) ' 4a5d9d5ba4ce2de1728e3bf480350f25e07e21c947d19e3376f09b3c1e161742
    Next
    Caption = Format$(Timer - dblTimer, "0.000")
End Sub

Private Sub pvTestCryptoEd25519()
    Dim baPriv()        As Byte
    Dim baPub()         As Byte
    Dim baMsg()         As Byte
    Dim baSig()         As Byte
    
    Debug.Assert CryptoSha2Text(256, "Rosetta code") = "764faf5c61ac315f1497f9dfa542713965b785e5cc2f707d6468d7d1124cdfcf"
    baPriv = FromHex("0102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f20")
    CryptoEd25519PublicKey baPub, baPriv
    Debug.Print ToHex(baPub) ' 79b5562e8fe654f94078b112e8a98ba7901f853ae695bed7e0e3910bad049664
    baMsg = StrConv("This is a test", vbFromUnicode)
    CryptoEd25519SignDetached baSig, baPriv, baMsg
    Debug.Print ToHex(baSig) ' b44c90cdc2123a247e645cad07dd83d5256524271b8f91e3d68b853bede61c05df5873f2dacaefd6ea0b17f2eda6dc23b076b718be2b840a8bac65f273527207
    If CryptoEd25519VerifyDetached(baSig, baPub, baMsg) Then
        Debug.Print "Verified message: " & StrConv(baMsg, vbUnicode)
    Else
        Debug.Print "Invalid signature!"
    End If
End Sub

Private Sub Form_Load()
    pvTestSha3
    pvTestSha512
    pvTestCryptoEd25519
    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha512_test.json")), 512
    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha384_test.json")), 384
    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha256_test.json")), 256
End Sub
