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
    pvTestScryptKdf
    Exit Sub
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

Private Sub pvTestPbkdf2HmacSha1()
    Dim dblTimer        As Double
    Dim baOutput()      As Byte
    
    dblTimer = TimerEx
    baOutput = CryptoPbkdf2HmacSha1ByteArray(StrConv("password", vbFromUnicode), StrConv("salt", vbFromUnicode), OutSize:=20, NumIter:=4096)
    Debug.Print Format$(TimerEx - dblTimer, "0.000"), ToHex(baOutput)
    '-> 4b007901b765489abead49d926f721d065a429c1
End Sub

Private Sub pvTestPbkdf2HmacSha2()
    Dim dblTimer        As Double
    Dim baOutput()      As Byte
    
    dblTimer = TimerEx
    baOutput = CryptoPbkdf2HmacSha2ByteArray(224, StrConv("Password", vbFromUnicode), StrConv("sa" & vbNullChar & "lt", vbFromUnicode), OutSize:=256, NumIter:=4096)
    Debug.Print Format$(TimerEx - dblTimer, "0.000"), ToHex(baOutput)
    '-> a329a360c825e12e454ad8633a842a06ba1456907770779d1fa4e0b61a5b1c6ce02e71de74ae433bbf14b907690d008d0cab5b01c976c1e627b027a9a809fd001082c809650344ecfcdebdf0d64b92cb1e869bf91b75517ea36918127b1eccc4cac145fb965071292a6dfa388d8ad893d2541f83a0dac1c55d2d90709963b066de985e92974e87b7d8c0e8026d96684bb0425203919b4792962b065e2b2b815ba888b8428ae51f57a74f637a658e27cf5fbc5593e85f775a1f81660850a723e2eb565f30dfc2cf2973ad57ec95b89c0979c7bab81c11d8987540a32badb2f7bbe4ff21a4f0d91dbd911b88ddd928603fd27b0ede994ee99edd2c04667b82067f
    Debug.Print CryptoPbkdf2HmacSha2Text(224, "Password", "sa" & vbNullChar & "lt", NumIter:=4096)
    dblTimer = TimerEx
    baOutput = CryptoPbkdf2HmacSha2ByteArray(256, StrConv("Password", vbFromUnicode), StrConv("sa" & vbNullChar & "lt", vbFromUnicode), OutSize:=256, NumIter:=4096)
    Debug.Print Format$(TimerEx - dblTimer, "0.000"), ToHex(baOutput)
    '-> 436c82c6af9010bb0fdb274791934ac7dee21745dd11fb57bb90112ab187c495ad82df776ad7cefb606f34fedca59baa5922a57f3e91bc0e11960da7ec87ed0471b456a0808b60dff757b7d313d4068bf8d337a99caede24f3248f87d1bf16892b70b076a07dd163a8a09db788ae34300ff2f2d0a92c9e678186183622a636f4cbce15680dfea46f6d224e51c299d4946aa2471133a649288eef3e4227b609cf203dba65e9fa69e63d35b6ff435ff51664cbd6773d72ebc341d239f0084b004388d6afa504eee6719a7ae1bb9daf6b7628d851fab335f1d13948e8ee6f7ab033a32df447f8d0950809a70066605d6960847ed436fa52cdfbcf261b44d2a87061
    dblTimer = TimerEx
    baOutput = CryptoPbkdf2HmacSha2ByteArray(512, StrConv("Password", vbFromUnicode), StrConv("sa" & vbNullChar & "lt", vbFromUnicode), OutSize:=256, NumIter:=4096)
    Debug.Print Format$(TimerEx - dblTimer, "0.000"), ToHex(baOutput)
    '-> 10176fb32cb98cd7bb31e2bb5c8f6e425c103333a2e496058e3fd2bd88f657485c89ef92daa0668316bc23ebd1ef88f6dd14157b2320b5d54b5f26377c5dc279b1dcdec044bd6f91b166917c80e1e99ef861b1d2c7bce1b961178125fb86867f6db489a2eae0022e7bc9cf421f044319fac765d70cb89b45c214590e2ffb2c2b565ab3b9d07571fde0027b1dc57f8fd25afa842c1056dd459af4074d7510a0c020b914a5e202445d4d3f151070589dd6a2554fc506018c4f001df6239643dc86771286ae4910769d8385531bba57544d63c3640b90c98f1445ebdd129475e02086b600f0beb5b05cc6ca9b3633b452b7dad634e9336f56ec4c3ac0b4fe54ced8
End Sub

Private Sub pvTestScryptKdf()
    Dim dblTimer        As Double
    Dim baOutput()      As Byte
    
    dblTimer = TimerEx
    baOutput = CryptoScryptKdfByteArray(StrConv("", vbFromUnicode), StrConv("", vbFromUnicode), OutSize:=64, Passes:=1, Memory:=16, Parallelism:=1)
    Debug.Print Format$(TimerEx - dblTimer, "0.000"), ToHex(baOutput)
    '-> 77d6576238657b203b19ca42c18a0497f16b4844e3074ae8dfdffa3fede21442fcd0069ded0948f8326a753a0fc81f17e8d3e0fb2e0d3628cf35e20c38d18906
    dblTimer = TimerEx
    baOutput = CryptoScryptKdfByteArray(StrConv("password", vbFromUnicode), StrConv("salt", vbFromUnicode), OutSize:=32)
    Debug.Print Format$(TimerEx - dblTimer, "0.000"), ToHex(baOutput), ToBase64Array(baOutput)
    '-> 745731af4484f323968969eda289aeee005b5903ac561e64a5aca121797bf773      dFcxr0SE8yOWiWntoomu7gBbWQOsVh5kpayhIXl793M=
End Sub

Private Sub pvTestHkdfSha2()
    Dim dblTimer        As Double
    Dim baOutput()      As Byte
    
    baOutput = CryptoHkdfSha2ByteArray(256, FromHex("0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b"), FromHex("000102030405060708090a0b0c"), FromHex("f0f1f2f3f4f5f6f7f8f9"), OutSize:=42)
    Debug.Print ToHex(baOutput)
    '-> 3cb25f25faacd57a90434f64d0362f2a2d2d0a90cf1a5a4c5db02d56ecc4c5bf34007208d5b887185865
    baOutput = CryptoHkdfSha2ByteArray(256, FromHex("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f202122232425262728292a2b2c2d2e2f303132333435363738393a3b3c3d3e3f404142434445464748494a4b4c4d4e4f"), _
        FromHex("606162636465666768696a6b6c6d6e6f707172737475767778797a7b7c7d7e7f808182838485868788898a8b8c8d8e8f909192939495969798999a9b9c9d9e9fa0a1a2a3a4a5a6a7a8a9aaabacadaeaf"), _
        FromHex("b0b1b2b3b4b5b6b7b8b9babbbcbdbebfc0c1c2c3c4c5c6c7c8c9cacbcccdcecfd0d1d2d3d4d5d6d7d8d9dadbdcdddedfe0e1e2e3e4e5e6e7e8e9eaebecedeeeff0f1f2f3f4f5f6f7f8f9fafbfcfdfeff"), OutSize:=82)
    Debug.Print ToHex(baOutput)
    '-> b11e398dc80327a1c8e7f78c596a49344f012eda2d4efad8a050cc4c19afa97c59045a99cac7827271cb41c65e590e09da3275600c2f09b8367793a9aca3db71cc30c58179ec3e87c14c01d5c1f3434f1d87
End Sub

Private Sub pvTestRipeMd160()
    Debug.Print CryptoRipeMd160Text("")
    '-> 9c1185a5c5e9fc54612808977ee8f548b2258d31
    Debug.Print CryptoRipeMd160Text("a")
    '-> 0bdc9d2d256b3ee9daae347be6f4dc835a467ffe
    Debug.Print CryptoRipeMd160Text("abc")
    '-> 8eb208f7e05d987a9b044a8e98c6b087f15a0bfc
    Debug.Print CryptoRipeMd160Text("Rosetta code")
    '-> 1cda558e41e47c3090aafd73ca5651d176f95ca9
End Sub

Private Sub pvTestBlake2s()
    Debug.Print CryptoBlake2sText(256, "")
    '-> 69217a3079908094e11121d042354a7c1f55b6482ca1a51e1b250dfd1ed0eef9
    
    Debug.Print CryptoBlake2sText(256, "abc")
    '-> 508c5e8c327c14e2e1a72ba34eeb452f37458b209ed63a294d999b4c86675982
    
    Debug.Print CryptoBlake2sText(256, "", Key:=FromHex("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f"))
    '-> 48a8997da407876b3d79c0d92325ad3b89cbb754d86ab71aee047ad345fd2c49
    
    Debug.Print CryptoBlake2sText(256, Chr$(0), Key:=FromHex("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f"))
    '-> 40d15fee7c328830166ac3f918650f807e7e01e177258cdc0a39b11f598066f1
End Sub

Private Sub pvTestBlake2b()
    Debug.Print CryptoBlake2bText(512, "")
    '-> 786a02f742015903c6c6fd852552d272912f4740e15847618a86e217f71f5419d25e1031afee585313896444934eb04b903a685b1448b755d56f701afe9be2ce
    
    Debug.Print CryptoBlake2bText(512, "abc")
    '-> ba80a53f981c4d0d6a2797b69f12f6e94c212f14685ac4b74b12bb6fdbffa2d17d87c5392aab792dc252d5de4533cc9518d38aa8dbf1925ab92386edd4009923
    
    Debug.Print CryptoBlake2bText(512, "", Key:=FromHex("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f202122232425262728292a2b2c2d2e2f303132333435363738393a3b3c3d3e3f"))
    '-> 10ebb67700b1868efb4417987acf4690ae9d972fb7a590c2f02871799aaa4786b5e996e8f0f4eb981fc214b005f42d2ff4233499391653df7aefcbc13fc51568

    Debug.Print CryptoBlake2bText(512, Chr$(0), Key:=FromHex("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f202122232425262728292a2b2c2d2e2f303132333435363738393a3b3c3d3e3f"))
    '-> 961f6dd1e4dd30f63901690c512e78e4b45e4742ed197c3c5e45c549fd25f2e4187b0bc9fe30492b16b0d0bc4ef9b0f34c7003fac09a5ef1532e69430234cebd
End Sub

Private Sub pvTestSha1()
    Debug.Print ToHex(CryptoSha1ByteArray(StrConv("abc", vbFromUnicode)))
    '-> a9993e364706816aba3e25717850c26c9cd0d89d
    Debug.Assert CryptoSha1Text("abc") = "a9993e364706816aba3e25717850c26c9cd0d89d"
    Debug.Print ToHex(CryptoSha1ByteArray(StrConv(vbNullString, vbFromUnicode)))
    '-> da39a3ee5e6b4b0d3255bfef95601890afd80709
End Sub

Private Sub pvTestSha2()
    Debug.Assert CryptoSha2Text(256, "Rosetta code") = "764faf5c61ac315f1497f9dfa542713965b785e5cc2f707d6468d7d1124cdfcf"
End Sub

Private Sub pvTestBlake3()
    Debug.Print CryptoBlake3Text("abc")
    '-> 6437b3ac38465133ffb63b75273a8db548c558465d79db03fd359c6cd5bd9d85
    Debug.Print CryptoBlake3Text("abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu")
    '-> 553e1aa2a477cb3166e6ab38c12d59f6c5017f0885aaf079f217da00cfca363f
    Debug.Print CryptoBlake3Text("abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghi")
    '-> 48064d35e3f3ef8047d9cd830165ced72877864a4a97aeb5338192218c7f7767
    Debug.Print CryptoBlake3Text("abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghij")
    '-> 924b59ee54a2bcf9c0783853226b5f286115eb68ae8d816e1ae2c26c64df5326
    Debug.Print CryptoBlake3Text("abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghi")
    '-> 89cdcdf66a829c8d84ae63c755cc9394fb0bbe8e6552639f63d1325e6431a369
    Debug.Print CryptoBlake3Text(Replace(Space(10), " ", "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghiabcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstuabcdefghbcdefghi"))
    '-> 4f86b3c727aea636fd845bc73428ea164e2a0bd43b95735f90829f18197187ce
    Debug.Assert CryptoBlake3Text("abc", Key:="00000000000000000000000000000000") = "91584f1fb59612f3d47c42ae41005f005e09f1f388f7112d4931f39d8790ef83"
    '-> 91584f1fb59612f3d47c42ae41005f005e09f1f388f7112d4931f39d8790ef83
    Debug.Print CryptoBlake3Text(Chr(0), Context:="BLAKE3 2019-12-27 16:29:52 test vectors context", OutSize:=131)
    '-> b3e2e340a117a499c6cf2398a19ee0d29cca2bb7404c73063382693bf66cb06c5827b91bf889b6b97c5477f535361caefca0b5d8c4746441c57617111933158950670f9aa8a05d791daae10ac683cbef8faf897c84e6114a59d2173c3f417023a35d6983f2c7dfa57e7fc559ad751dbfb9ffab39c2ef8c4aafebc9ae973a64f0c76551
End Sub

Private Sub pvTestArgon2()
    Dim baOutput()          As Byte
    
    baOutput = CryptoArgon2KdfByteArray(FromHex("0101010101010101010101010101010101010101010101010101010101010101"), FromHex("02020202020202020202020202020202"), _
        Secret:=FromHex("0303030303030303"), Data:=FromHex("040404040404040404040404"), OutSize:=32, Passes:=3, Memory:=32, Parallelism:=4)
    Debug.Assert ToHex(baOutput) = "c814d9d1dc7f37aa13f0d77f2494bda1c8de6b016dd388d29952a4c4672b6ce8"
    '-- c814d9d1dc7f37aa13f0d77f2494bda1c8de6b016dd388d29952a4c4672b6ce8
    Debug.Print CryptoArgon2KdfText("pasword", "somesalt", Passes:=3, Memory:=4096, Parallelism:=1)
    '-> 957fc0727d83f4060bb0f1071eb590a19a8c448fc0209497ee4f54ca241f3c90
    Debug.Print CryptoArgon2IdKdfText("pasword", "somesalt", Passes:=3, Memory:=4096, Parallelism:=1)
    '-> f55535bfe948710051424c7424b11ba9a13a50239b0459f56ca695ea14bc195e
    Debug.Print CryptoArgon2KdfText("pasword", "somesalt", Memory:=32768, Passes:=1, OutSize:=32, Parallelism:=4)
    '-> afe22ddd167cacbb9ac102ed56143e54fca5042f29e015954f93b885499a7f8d
End Sub

Private Sub pvTestSiphash()
    Dim baKey()         As Byte
    Dim baOutput()      As Byte
    
    
    Debug.Print CryptoSiphash24Text("123", "123")
    '-> 171b9ef8c7d20383
    baKey = FromHex("000102030405060708090A0B0C0D0E0F")
    baOutput = CryptoSiphash24ByteArray(baKey, FromHex("000102030405060708090A0B0C0D0E0F10111213"))
    Debug.Print ToHex(baOutput)
    '-> 98eea21af25cd6be
    
    Debug.Print CryptoHalfSiphash24Text("123", "123")
    '-> 28b6f273 5ea1ceccd121bf6c
    baOutput = CryptoHalfSiphash24ByteArray(baKey, FromHex("00"))
    Debug.Print ToHex(baOutput)
    '-> 27475ab8 be552412f8387315
    baOutput = CryptoHalfSiphash24ByteArray(baKey, FromHex("000102030405060708090A0B0C0D0E0F10111213"))
    Debug.Print ToHex(baOutput)
    '-> e34d1045 1bf52b0a6feea7db
    Debug.Print Hex(CryptoHalfSiphash24Long(baKey, FromHex("000102030405060708090A0B0C0D0E0F10111213")))
    '-> 45104DE3
End Sub

Private Sub Form_Load()
    pvTestSiphash
'    Debug.Print Timer
'    CryptoTestArgon2
'    Debug.Print Timer
'    pvTestArgon2
'    pvTestBlake3
'    pvTestSha1
'    pvTestSha2
'    pvTestBlake2s
'    pvTestBlake2b
'    pvTestPbkdf2HmacSha1
'    pvTestPbkdf2HmacSha2
'    pvTestScryptKdf
    pvTestRipeMd160
    pvTestHkdfSha2
    pvTestSha3
    pvTestSha512
    pvTestCryptoEd25519
    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha512_test.json")), 512
    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha384_test.json")), 384
    pvTestHmacSha2 JsonParseObject(ReadTextFile("C:\Work\Temp\wycheproof\testvectors\hmac_sha256_test.json")), 256
End Sub
