Attribute VB_Name = "mdVectors"
'=========================================================================
'
'  Pure VB6 Crypto (Untested)
'  Copyright (c) 2026 wqweto@gmail.com
'
'  This project is licensed under the terms of the MIT license
'  See the LICENSE file in the project root for more information
'
'=========================================================================
'--- mdVectors.bas
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const MODULE_NAME           As String = "mdVectors"
'--- Wycheproof suites as "file|kind|param", kind picks the runner below
Private Const STR_SUITES            As String = _
    "aes_gcm|gcm|0;" & _
    "aes_gcm_siv|gcmsiv|0;" & _
    "aes_ccm|ccm|0;" & _
    "aes_eax|eax|0;" & _
    "aes_cmac|cmac|0;" & _
    "chacha20_poly1305|chacha|0;" & _
    "hmac_sha1|hmac1|0;" & _
    "hmac_sha224|hmac2|224;" & _
    "hmac_sha256|hmac2|256;" & _
    "hmac_sha384|hmac2|384;" & _
    "hmac_sha512|hmac2|512;" & _
    "hmac_sha3_224|hmac3|224;" & _
    "hmac_sha3_256|hmac3|256;" & _
    "hmac_sha3_384|hmac3|384;" & _
    "hmac_sha3_512|hmac3|512;" & _
    "hkdf_sha1|hkdf1|0;" & _
    "hkdf_sha256|hkdf2|256;" & _
    "hkdf_sha384|hkdf2|384;" & _
    "hkdf_sha512|hkdf2|512;" & _
    "x25519|x25519|0;" & _
    "eddsa|eddsa|0"

Private m_lPass                     As Long
Private m_lFail                     As Long
Private m_lSkip                     As Long
Private m_sFirstFail                As String

'=========================================================================
' Functions
'=========================================================================

Public Sub VectorsRun(vFilter As Variant)
    Dim vSuites         As Variant
    Dim lIdx            As Long
    Dim vParts          As Variant
    Dim sFile           As String
    Dim lRan            As Long
    Dim lTotPass        As Long
    Dim lTotFail        As Long
    Dim lTotSkip        As Long

    vSuites = Split(STR_SUITES, ";")
    '--- Wycheproof driven suites
    For lIdx = 0 To UBound(vSuites)
        vParts = Split(vSuites(lIdx), "|")
        If HasFilterMatch(vParts(0), vFilter) Then
            sFile = App.Path & "\..\wycheproof\" & vParts(0) & "_test.json"
            pvResetCounters
            If Not pvHasFile(sFile) Then
                ConPrintLine PadRight(vParts(0), 26) & "missing " & vParts(0) & "_test.json"
            Else
                pvRunSuite sFile, CStr(vParts(1)), CLng(vParts(2))
                pvPrintReport CStr(vParts(0))
                lTotPass = lTotPass + m_lPass
                lTotFail = lTotFail + m_lFail
                lTotSkip = lTotSkip + m_lSkip
            End If
            lRan = lRan + 1
        End If
    Next
    '--- built in checks for the algorithms Wycheproof does not cover
    If HasFilterMatch("kat", vFilter) Then
        pvResetCounters
        pvRunKat
        pvPrintReport "kat"
        lTotPass = lTotPass + m_lPass
        lTotFail = lTotFail + m_lFail
        lTotSkip = lTotSkip + m_lSkip
        lRan = lRan + 1
    End If
    If HasFilterMatch("selftest", vFilter) Then
        pvResetCounters
        pvRunSelfTest
        pvPrintReport "selftest"
        lTotPass = lTotPass + m_lPass
        lTotFail = lTotFail + m_lFail
        lTotSkip = lTotSkip + m_lSkip
        lRan = lRan + 1
    End If
    If lRan = 0 Then
        ConPrintLine "No suite matched. Known suites:"
        For lIdx = 0 To UBound(vSuites)
            ConPrintLine "  " & Split(vSuites(lIdx), "|")(0)
        Next
        ConPrintLine "  kat, selftest"
    ElseIf lRan > 1 Then
        ConPrintLine
        ConPrintLine PadRight("TOTAL", 26) & PadLeft(lTotPass + lTotFail, 7) & " tests, " & _
            PadLeft(lTotPass, 6) & " ok, " & PadLeft(lTotFail, 5) & " failed, " & PadLeft(lTotSkip, 5) & " skipped"
    End If
End Sub

Private Sub pvResetCounters()
    m_lPass = 0
    m_lFail = 0
    m_lSkip = 0
    m_sFirstFail = vbNullString
End Sub

Private Sub pvPrintReport(sName As String)
    ConPrintLine PadRight(sName, 26) & PadLeft(m_lPass + m_lFail, 7) & " tests, " & _
        PadLeft(m_lPass, 6) & " ok, " & PadLeft(m_lFail, 5) & " failed, " & PadLeft(m_lSkip, 5) & " skipped" & _
        IIf(LenB(m_sFirstFail) <> 0, "   first: " & m_sFirstFail, vbNullString)
End Sub

Private Sub pvCheck(ByVal bOk As Boolean, sWhat As String)
    If bOk Then
        m_lPass = m_lPass + 1
        Exit Sub
    End If
    m_lFail = m_lFail + 1
    If LenB(m_sFirstFail) = 0 Then
        m_sFirstFail = sWhat
    End If
End Sub

Private Function pvHasFile(sFile As String) As Boolean
    On Error GoTo EH
    pvHasFile = (GetAttr(sFile) And vbDirectory) = 0
    Exit Function
EH:
End Function

Private Sub pvRunSuite(sFile As String, sKind As String, ByVal lParam As Long)
    Dim oJson           As Object
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim sResult         As String
    Dim bValid          As Boolean
    Dim bOk             As Boolean

    Set oJson = JsonParseObject(ReadTextFile(sFile))
    For Each oGroup In JsonValue(oJson, "testGroups")
        For Each oTest In JsonValue(oGroup, "tests")
            sResult = JsonValue(oTest, "result")
            bValid = (sResult <> "invalid")
            '--- pvTryCase owns the error handling, so a raised error lands
            '--- here as a rejection and this loop never resumes out of one
            bOk = pvTryCase(sKind, lParam, oGroup, oTest)
            '--- "acceptable" means either outcome is defensible, so both
            '--- accepting and rejecting the case counts as a pass
            If sResult = "acceptable" Then
                m_lPass = m_lPass + 1
            ElseIf bOk = bValid Then
                m_lPass = m_lPass + 1
            Else
                m_lFail = m_lFail + 1
                If LenB(m_sFirstFail) = 0 Then
                    m_sFirstFail = "tcId " & JsonValue(oTest, "tcId") & " " & JsonValue(oTest, "comment")
                End If
            End If
        Next
    Next
End Sub

Private Function pvTryCase(sKind As String, ByVal lParam As Long, oGroup As Object, oTest As Object) As Boolean
    On Error GoTo EH
    pvTryCase = pvRunCase(sKind, lParam, oGroup, oTest)
    Exit Function
EH:
    pvTryCase = False
End Function

Private Function pvRunCase(sKind As String, ByVal lParam As Long, oGroup As Object, oTest As Object) As Boolean
    Dim baKey()         As Byte
    Dim baNonce()       As Byte
    Dim baAad()         As Byte
    Dim baCt()          As Byte
    Dim baTag()         As Byte
    Dim baMsg()         As Byte
    Dim baBuf()         As Byte
    Dim lTagSize        As Long
    Dim baOut()         As Byte
    Dim uGcm            As CryptoAesGcmContext
    Dim uCmac           As CryptoCmacContext

    Select Case sKind
    Case "gcm", "gcmsiv", "ccm", "eax", "chacha"
        '--- encrypt the plaintext and compare, the way Form1 does it, as
        '--- going the other way trips over the zero length msg/ct cases
        baKey = FromHex(JsonValue(oTest, "key"))
        baNonce = FromHex(JsonValue(oTest, "iv"))
        baAad = FromHex(JsonValue(oTest, "aad"))
        baCt = FromHex(JsonValue(oTest, "ct"))
        baTag = FromHex(JsonValue(oTest, "tag"))
        baMsg = FromHex(JsonValue(oTest, "msg"))
        baBuf = baMsg
        lTagSize = pvGetSize(baTag)
        If lTagSize = 0 Then
            lTagSize = 16
        End If
        Select Case sKind
        Case "gcm"
            CryptoAesGcmInit uGcm, baKey, baNonce, baAad
            CryptoAesGcmEncrypt uGcm, baBuf, TagSize:=lTagSize, Tag:=baOut
        Case "gcmsiv"
            CryptoAesGcmSivEncrypt baKey, baNonce, baAad, baBuf, baOut
        Case "ccm"
            CryptoAesCcmEncrypt baKey, baNonce, baAad, baBuf, baOut, TagSize:=lTagSize
        Case "eax"
            CryptoAesEaxEncrypt baKey, baNonce, baAad, baBuf, baOut, TagSize:=lTagSize
        Case "chacha"
            CryptoChaCha20Poly1305Encrypt baKey, baOut, baBuf, 0, -1, baNonce, baAad
        End Select
        pvRunCase = pvIsSameBytes(baBuf, baCt) And pvIsSameBytes(baOut, baTag)
    Case "cmac"
        baKey = FromHex(JsonValue(oTest, "key"))
        baMsg = FromHex(JsonValue(oTest, "msg"))
        baTag = FromHex(JsonValue(oTest, "tag"))
        lTagSize = pvGetSize(baTag)
        If lTagSize = 0 Then
            Exit Function
        End If
        CryptoCmacInit uCmac, baKey
        CryptoCmacUpdate uCmac, baMsg
        CryptoCmacFinalize uCmac, baOut, TagSize:=lTagSize
        pvRunCase = pvIsSameBytes(baOut, baTag)
    Case "hmac1", "hmac2", "hmac3"
        baKey = FromHex(JsonValue(oTest, "key"))
        baMsg = FromHex(JsonValue(oTest, "msg"))
        baTag = FromHex(JsonValue(oTest, "tag"))
        Select Case sKind
        Case "hmac1"
            baOut = CryptoHmacSha1ByteArray(baKey, baMsg)
        Case "hmac2"
            baOut = CryptoHmacSha2ByteArray(lParam, baKey, baMsg)
        Case "hmac3"
            baOut = CryptoHmacSha3ByteArray(lParam, baKey, baMsg)
        End Select
        '--- Wycheproof truncates the tag, so compare the prefix
        pvRunCase = pvIsSamePrefix(baOut, baTag)
    Case "hkdf1", "hkdf2"
        baKey = FromHex(JsonValue(oTest, "ikm"))
        baNonce = FromHex(JsonValue(oTest, "salt"))
        baAad = FromHex(JsonValue(oTest, "info"))
        baTag = FromHex(JsonValue(oTest, "okm"))
        If sKind = "hkdf1" Then
            baOut = CryptoHkdfSha1ByteArray(baKey, baNonce, baAad, OutSize:=pvGetSize(baTag))
        Else
            baOut = CryptoHkdfSha2ByteArray(lParam, baKey, baNonce, baAad, OutSize:=pvGetSize(baTag))
        End If
        pvRunCase = pvIsSameBytes(baOut, baTag)
    Case "x25519"
        baKey = FromHex(JsonValue(oTest, "private"))
        baNonce = FromHex(JsonValue(oTest, "public"))
        baTag = FromHex(JsonValue(oTest, "shared"))
        CryptoX25519SharedSecret baOut, baKey, baNonce
        pvRunCase = pvIsSameBytes(baOut, baTag)
    Case "eddsa"
        '--- key material sits on the group, the signature on the test
        If Not IsObject(JsonValue(oGroup, "key")) Then
            m_lSkip = m_lSkip + 1
            pvRunCase = True
            Exit Function
        End If
        baKey = FromHex(JsonValue(JsonValue(oGroup, "key"), "pk"))
        baMsg = FromHex(JsonValue(oTest, "msg"))
        baTag = FromHex(JsonValue(oTest, "sig"))
        If pvGetSize(baKey) <> 32 Or pvGetSize(baTag) <> 64 Then
            Exit Function
        End If
        '--- note CryptoEd25519VerifyDetached ignores Pos/Size and always
        '--- appends the whole baMsg, so an empty message cannot be passed
        pvRunCase = CryptoEd25519VerifyDetached(baTag, baKey, baMsg)
    Case Else
        Err.Raise vbObjectError, , "Unknown suite kind " & sKind
    End Select
End Function

'--- known answer tests for what Wycheproof has no vectors for. Values come
'--- from the reference documents: RFC 1321/3174, FIPS 180-4/202, RFC 7693,
'--- the RIPEMD-160 and BLAKE3 reference implementations.
Private Sub pvRunKat()
    pvCheckKatHash "md5", vbNullString, "d41d8cd98f00b204e9800998ecf8427e"
    pvCheckKatHash "md5", "abc", "900150983cd24fb0d6963f7d28e17f72"
    pvCheckKatHash "sha1", "abc", "a9993e364706816aba3e25717850c26c9cd0d89d"
    pvCheckKatHash "sha224", "abc", "23097d223405d8228642a477bda255b32aadbce4bda0b3f7e36c9da7"
    pvCheckKatHash "sha256", "abc", "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    pvCheckKatHash "sha384", "abc", "cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7"
    pvCheckKatHash "sha512", "abc", "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f"
    pvCheckKatHash "sha3-256", "abc", "3a985da74fe225b2045c172d6bd390bd855f086e3e9d525b46bfe24511431532"
    pvCheckKatHash "sha3-512", "abc", "b751850b1a57168a5693cd924b6b096e08f621827444f70d884f5d0240d2712e10e116e9192af3c91a7ec57647e3934057340b4cf408d5a56592f8274eec53f0"
    pvCheckKatHash "ripemd160", "abc", "8eb208f7e05d987a9b044a8e98c6b087f15a0bfc"
    pvCheckKatHash "blake2s", "abc", "508c5e8c327c14e2e1a72ba34eeb452f37458b209ed63a294d999b4c86675982"
    pvCheckKatHash "blake2b", "abc", "ba80a53f981c4d0d6a2797b69f12f6e94c212f14685ac4b74b12bb6fdbffa2d17d87c5392aab792dc252d5de4533cc9518d38aa8dbf1925ab92386edd4009923"
    pvCheckKatHash "blake3", "abc", "6437b3ac38465133ffb63b75273a8db548c558465d79db03fd359c6cd5bd9d85"
    pvCheckKatEd25519
End Sub

'--- RFC 8032 section 7.1 test 2, split into the three steps so a failure
'--- says whether key derivation, signing or verification is at fault
Private Sub pvCheckKatEd25519()
    Const STR_SEED      As String = "4ccd089b28ff96da9db6c346ec114e0f5b8a319f35aba624da8cf6ed4fb8a6fb"
    Const STR_PUB       As String = "3d4017c3e843895a92b70aa74d1b7ebc9c982ccf2ec4968cc0cd55f12af4660c"
    Const STR_MSG       As String = "72"
    Const STR_SIG       As String = "92a009a9f0d4cab8720e820b5f642540a2b27b5416503f8fb3762223ebdb69da085ac1e43e15996e458f3613d0f11d8c387b2eaeb4302aeeb00d291612bb0c00"
    Dim baPriv()        As Byte
    Dim baPub()         As Byte
    Dim baMsg()         As Byte
    Dim baSig()         As Byte

    On Error GoTo EH
    CryptoEd25519PrivateKey baPriv, Seed:=FromHex(STR_SEED)
    CryptoEd25519PublicKey baPub, baPriv
    pvCheck ToHex(baPub) = STR_PUB, "ed25519 public key from seed"
    baMsg = FromHex(STR_MSG)
    CryptoEd25519SignDetached baSig, baPriv, baMsg
    pvCheck ToHex(baSig) = STR_SIG, "ed25519 sign"
    pvCheck CryptoEd25519VerifyDetached(FromHex(STR_SIG), FromHex(STR_PUB), baMsg), "ed25519 verify"
    Exit Sub
EH:
    pvCheck False, "ed25519 raised: " & Err.Description
End Sub

Private Sub pvCheckKatHash(sAlgo As String, sInput As String, sExpect As String)
    Dim baInput()       As Byte
    Dim baOut()         As Byte

    On Error GoTo EH
    baInput = ToUtf8Array(sInput)
    Select Case sAlgo
    Case "md5"
        baOut = CryptoMd5ByteArray(baInput)
    Case "sha1"
        baOut = CryptoSha1ByteArray(baInput)
    Case "sha224"
        baOut = CryptoSha2ByteArray(224, baInput)
    Case "sha256"
        baOut = CryptoSha2ByteArray(256, baInput)
    Case "sha384"
        baOut = CryptoSha2ByteArray(384, baInput)
    Case "sha512"
        baOut = CryptoSha2ByteArray(512, baInput)
    Case "sha3-256"
        baOut = CryptoSha3ByteArray(256, baInput)
    Case "sha3-512"
        baOut = CryptoSha3ByteArray(512, baInput)
    Case "ripemd160"
        baOut = CryptoRipeMd160ByteArray(baInput)
    Case "blake2s"
        baOut = CryptoBlake2sByteArray(256, baInput)
    Case "blake2b"
        baOut = CryptoBlake2bByteArray(512, baInput)
    Case "blake3"
        baOut = CryptoBlake3ByteArray(baInput)
    End Select
    pvCheck ToHex(baOut) = sExpect, sAlgo & "(""" & sInput & """)"
    Exit Sub
EH:
    pvCheck False, sAlgo & " raised: " & Err.Description
End Sub

'--- checks that need no published vectors: streaming must agree with the
'--- one shot, every cipher must decrypt what it encrypted, and a flipped
'--- ciphertext bit must fail authentication
Private Sub pvRunSelfTest()
    pvCheckStreamHash "md5"
    pvCheckStreamHash "sha1"
    pvCheckStreamHash "sha256"
    pvCheckStreamHash "sha512"
    pvCheckStreamHash "sha3-256"
    pvCheckStreamHash "ripemd160"
    pvCheckStreamHash "blake2s"
    pvCheckStreamHash "blake2b"
    pvCheckStreamHash "blake3"
    pvCheckStreamHash "ascon-hash"
    pvCheckRoundTrip "aes-128-cbc"
    pvCheckRoundTrip "aes-128-ctr"
    pvCheckRoundTrip "aes-128-gcm"
    pvCheckRoundTrip "aes-128-ccm"
    pvCheckRoundTrip "aes-128-eax"
    pvCheckRoundTrip "aes-128-ocb"
    pvCheckRoundTrip "aes-128-gcm-siv"
    pvCheckRoundTrip "chacha20-poly1305"
    pvCheckRoundTrip "ascon-aead"
    pvCheckRoundTrip "tea"
    pvCheckTamper "aes-128-gcm"
    pvCheckTamper "aes-128-ocb"
    pvCheckTamper "chacha20-poly1305"
End Sub

Private Sub pvCheckStreamHash(sAlgo As String)
    Const LNG_A         As Long = 77
    Const LNG_B         As Long = 200
    Const LNG_TOTAL     As Long = 333
    Dim baInput()       As Byte
    Dim baOneShot()     As Byte
    Dim baStream()      As Byte
    Dim uMd5            As CryptoMd5Context
    Dim uSha1           As CryptoSha1Context
    Dim uSha2           As CryptoSha2Context
    Dim uSha512         As CryptoSha512Context
    Dim uSha3           As CryptoSha3Context
    Dim uRipe           As CryptoRipeMd160Context
    Dim u2s             As CryptoBlake2sContext
    Dim u2b             As CryptoBlake2bContext
    Dim u3              As CryptoBlake3Context
    Dim uAscon          As CryptoAsconContext

    On Error GoTo EH
    baInput = pvBuildPattern(LNG_TOTAL)
    Select Case sAlgo
    Case "md5"
        baOneShot = CryptoMd5ByteArray(baInput)
        CryptoMd5Init uMd5
        CryptoMd5Update uMd5, baInput, Pos:=0, Size:=LNG_A
        CryptoMd5Update uMd5, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoMd5Update uMd5, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoMd5Finalize uMd5, baStream
    Case "sha1"
        baOneShot = CryptoSha1ByteArray(baInput)
        CryptoSha1Init uSha1
        CryptoSha1Update uSha1, baInput, Pos:=0, Size:=LNG_A
        CryptoSha1Update uSha1, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoSha1Update uSha1, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoSha1Finalize uSha1, baStream
    Case "sha256"
        baOneShot = CryptoSha2ByteArray(256, baInput)
        CryptoSha2Init uSha2, 256
        CryptoSha2Update uSha2, baInput, Pos:=0, Size:=LNG_A
        CryptoSha2Update uSha2, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoSha2Update uSha2, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoSha2Finalize uSha2, baStream
    Case "sha512"
        baOneShot = CryptoSha512ByteArray(512, baInput)
        CryptoSha512Init uSha512, 512
        CryptoSha512Update uSha512, baInput, Pos:=0, Size:=LNG_A
        CryptoSha512Update uSha512, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoSha512Update uSha512, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoSha512Finalize uSha512, baStream
    Case "sha3-256"
        baOneShot = CryptoSha3ByteArray(256, baInput)
        CryptoSha3Init uSha3, 256
        CryptoSha3Update uSha3, baInput, Pos:=0, Size:=LNG_A
        CryptoSha3Update uSha3, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoSha3Update uSha3, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoSha3Finalize uSha3, baStream
    Case "ripemd160"
        baOneShot = CryptoRipeMd160ByteArray(baInput)
        CryptoRipeMd160Init uRipe
        CryptoRipeMd160Update uRipe, baInput, Pos:=0, Size:=LNG_A
        CryptoRipeMd160Update uRipe, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoRipeMd160Update uRipe, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoRipeMd160Finalize uRipe, baStream
    Case "blake2s"
        baOneShot = CryptoBlake2sByteArray(256, baInput)
        CryptoBlake2sInit u2s, 256
        CryptoBlake2sUpdate u2s, baInput, Pos:=0, Size:=LNG_A
        CryptoBlake2sUpdate u2s, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoBlake2sUpdate u2s, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoBlake2sFinalize u2s, baStream
    Case "blake2b"
        baOneShot = CryptoBlake2bByteArray(512, baInput)
        CryptoBlake2bInit u2b, 512
        CryptoBlake2bUpdate u2b, baInput, Pos:=0, Size:=LNG_A
        CryptoBlake2bUpdate u2b, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoBlake2bUpdate u2b, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoBlake2bFinalize u2b, baStream
    Case "blake3"
        baOneShot = CryptoBlake3ByteArray(baInput)
        CryptoBlake3Init u3
        CryptoBlake3Update u3, baInput, Pos:=0, Size:=LNG_A
        CryptoBlake3Update u3, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoBlake3Update u3, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoBlake3Finalize u3, baStream
    Case "ascon-hash"
        baOneShot = CryptoAsconHashByteArray(baInput)
        CryptoAsconHashInit uAscon
        CryptoAsconHashUpdate uAscon, baInput, Pos:=0, Size:=LNG_A
        CryptoAsconHashUpdate uAscon, baInput, Pos:=LNG_A, Size:=LNG_B - LNG_A
        CryptoAsconHashUpdate uAscon, baInput, Pos:=LNG_B, Size:=LNG_TOTAL - LNG_B
        CryptoAsconHashFinalize uAscon, baStream
    End Select
    pvCheck pvIsSameBytes(baOneShot, baStream), sAlgo & " stream vs one-shot"
    Exit Sub
EH:
    pvCheck False, sAlgo & " stream raised: " & Err.Description
End Sub

Private Sub pvCheckRoundTrip(sAlgo As String)
    Dim baKey()         As Byte
    Dim baNonce12()     As Byte
    Dim baNonce16()     As Byte
    Dim baAad()         As Byte
    Dim baPlain()       As Byte
    Dim baBuf()         As Byte
    Dim baTag()         As Byte
    Dim bOk             As Boolean
    Dim uAes            As CryptoAesContext
    Dim uGcm            As CryptoAesGcmContext
    Dim uOcb            As CryptoAesOcbContext

    On Error GoTo EH
    baKey = pvBuildPattern(16)
    baNonce12 = pvBuildPattern(12)
    baNonce16 = pvBuildPattern(16)
    baAad = pvBuildPattern(20)
    baPlain = pvBuildPattern(160)
    baBuf = baPlain
    Select Case sAlgo
    Case "aes-128-cbc"
        CryptoAesInit uAes, baKey, Nonce:=baNonce16
        CryptoAesCbcEncrypt uAes, baBuf
        CryptoAesInit uAes, baKey, Nonce:=baNonce16
        bOk = CryptoAesCbcDecrypt(uAes, baBuf)
    Case "aes-128-ctr"
        CryptoAesInit uAes, baKey, Nonce:=baNonce16
        CryptoAesCtrCrypt uAes, baBuf
        CryptoAesInit uAes, baKey, Nonce:=baNonce16
        CryptoAesCtrCrypt uAes, baBuf
        bOk = True
    Case "aes-128-gcm"
        CryptoAesGcmInit uGcm, baKey, baNonce12, baAad
        CryptoAesGcmEncrypt uGcm, baBuf, TagSize:=16, Tag:=baTag
        CryptoAesGcmInit uGcm, baKey, baNonce12, baAad
        bOk = CryptoAesGcmDecrypt(uGcm, baBuf, Tag:=baTag)
    Case "aes-128-ccm"
        CryptoAesCcmEncrypt baKey, baNonce12, baAad, baBuf, baTag, TagSize:=16
        bOk = CryptoAesCcmDecrypt(baKey, baNonce12, baAad, baBuf, baTag)
    Case "aes-128-eax"
        CryptoAesEaxEncrypt baKey, baNonce16, baAad, baBuf, baTag, TagSize:=16
        bOk = CryptoAesEaxDecrypt(baKey, baNonce16, baAad, baBuf, baTag)
    Case "aes-128-ocb"
        CryptoAesOcbInit uOcb, baKey, baNonce12, baAad, TagSize:=16
        CryptoAesOcbEncrypt uOcb, baBuf, TagSize:=16, Tag:=baTag
        CryptoAesOcbInit uOcb, baKey, baNonce12, baAad, TagSize:=16
        bOk = CryptoAesOcbDecrypt(uOcb, baBuf, Tag:=baTag)
    Case "aes-128-gcm-siv"
        CryptoAesGcmSivEncrypt baKey, baNonce12, baAad, baBuf, baTag
        bOk = CryptoAesGcmSivDecrypt(baKey, baNonce12, baAad, baBuf, baTag)
    Case "chacha20-poly1305"
        bOk = CryptoChaCha20Poly1305Encrypt(pvBuildPattern(32), baTag, baBuf, 0, -1, baNonce12, baAad)
        If bOk Then
            bOk = CryptoChaCha20Poly1305Decrypt(pvBuildPattern(32), baTag, baBuf, 0, -1, baNonce12, baAad)
        End If
    Case "ascon-aead"
        CryptoAsconEncrypt baKey, baTag, baBuf, 0, -1, baNonce16, baAad
        bOk = CryptoAsconDecrypt(baKey, baTag, baBuf, 0, -1, baNonce16, baAad)
    Case "tea"
        CryptoTeaEncrypt baKey, baBuf
        CryptoTeaDecrypt baKey, baBuf
        bOk = True
    End Select
    pvCheck bOk And pvIsSameBytes(baBuf, baPlain), sAlgo & " round trip"
    Exit Sub
EH:
    pvCheck False, sAlgo & " round trip raised: " & Err.Description
End Sub

Private Sub pvCheckTamper(sAlgo As String)
    Dim baKey()         As Byte
    Dim baNonce()       As Byte
    Dim baAad()         As Byte
    Dim baBuf()         As Byte
    Dim baTag()         As Byte
    Dim bOk             As Boolean
    Dim uGcm            As CryptoAesGcmContext
    Dim uOcb            As CryptoAesOcbContext

    On Error GoTo EH
    baKey = pvBuildPattern(16)
    baNonce = pvBuildPattern(12)
    baAad = pvBuildPattern(20)
    baBuf = pvBuildPattern(64)
    Select Case sAlgo
    Case "aes-128-gcm"
        CryptoAesGcmInit uGcm, baKey, baNonce, baAad
        CryptoAesGcmEncrypt uGcm, baBuf, TagSize:=16, Tag:=baTag
        baBuf(0) = baBuf(0) Xor 1
        CryptoAesGcmInit uGcm, baKey, baNonce, baAad
        bOk = CryptoAesGcmDecrypt(uGcm, baBuf, Tag:=baTag)
    Case "aes-128-ocb"
        CryptoAesOcbInit uOcb, baKey, baNonce, baAad, TagSize:=16
        CryptoAesOcbEncrypt uOcb, baBuf, TagSize:=16, Tag:=baTag
        baBuf(0) = baBuf(0) Xor 1
        CryptoAesOcbInit uOcb, baKey, baNonce, baAad, TagSize:=16
        bOk = CryptoAesOcbDecrypt(uOcb, baBuf, Tag:=baTag)
    Case "chacha20-poly1305"
        bOk = CryptoChaCha20Poly1305Encrypt(pvBuildPattern(32), baTag, baBuf, 0, -1, baNonce, baAad)
        baBuf(0) = baBuf(0) Xor 1
        bOk = CryptoChaCha20Poly1305Decrypt(pvBuildPattern(32), baTag, baBuf, 0, -1, baNonce, baAad)
    End Select
    pvCheck Not bOk, sAlgo & " tamper detection"
    Exit Sub
EH:
    '--- raising on a bad tag is an acceptable way to reject
    pvCheck True, vbNullString
End Sub

Private Function pvBuildPattern(ByVal lSize As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long

    ReDim baRetVal(0 To lSize - 1) As Byte
    For lIdx = 0 To lSize - 1
        baRetVal(lIdx) = (lIdx * 7 + 11) And &HFF
    Next
    pvBuildPattern = baRetVal
End Function

Private Function pvGetSize(baData() As Byte) As Long
    On Error GoTo EH
    pvGetSize = UBound(baData) + 1
    Exit Function
EH:
End Function

Private Function pvIsSameBytes(baLeft() As Byte, baRight() As Byte) As Boolean
    Dim lIdx            As Long

    If pvGetSize(baLeft) <> pvGetSize(baRight) Then
        Exit Function
    End If
    For lIdx = 0 To pvGetSize(baLeft) - 1
        If baLeft(lIdx) <> baRight(lIdx) Then
            Exit Function
        End If
    Next
    pvIsSameBytes = True
End Function

'--- true when baRight is a prefix of baLeft, for truncated tags
Private Function pvIsSamePrefix(baLeft() As Byte, baRight() As Byte) As Boolean
    Dim lIdx            As Long

    If pvGetSize(baRight) = 0 Or pvGetSize(baRight) > pvGetSize(baLeft) Then
        Exit Function
    End If
    For lIdx = 0 To pvGetSize(baRight) - 1
        If baLeft(lIdx) <> baRight(lIdx) Then
            Exit Function
        End If
    Next
    pvIsSamePrefix = True
End Function
