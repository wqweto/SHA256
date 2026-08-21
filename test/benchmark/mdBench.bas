Attribute VB_Name = "mdBench"
'=========================================================================
'
'  Pure VB6 Crypto (Untested)
'  Copyright (c) 2026 wqweto@gmail.com
'
'  This project is licensed under the terms of the MIT license
'  See the LICENSE file in the project root for more information
'
'=========================================================================
'--- mdBench.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)

'=========================================================================
' API
'=========================================================================

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#End If

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const MODULE_NAME           As String = "mdBench"
'--- block sizes, the ladder openssl speed uses
Private Const STR_SIZES             As String = "16|64|256|1024|8192|65536"
'--- seconds per measurement and per calibration probe
Private Const DBL_TARGET            As Double = 1#
Private Const DBL_PROBE             As Double = 0.05
'--- names in UcsAlgoEnum order, throughput measured in bytes/sec
Private Const STR_BYTES             As String = "md5|sha1|sha224|sha256|sha384|sha512|sha3-256|sha3-512|shake128|ripemd160|blake2s|blake2b|blake3|ascon-hash|siphash24|halfsiphash24|hmac-sha256|cmac-aes128|ghash|poly1305|aes-128-cbc|aes-128-ctr|aes-128-gcm|aes-128-ccm|aes-128-eax|aes-128-ocb|aes-128-gcm-siv|chacha20|chacha20-poly1305|ascon-aead|tea|aes-256-cbc|aes-256-ctr|aes-256-gcm|aes-128-cbc-dec|aes-128-gcm-dec|aes-128-ccm-dec|aes-128-eax-dec|aes-128-ocb-dec|aes-128-gcm-siv-dec|chacha20-poly1305-dec|ascon-aead-dec|tea-dec|aes-256-gcm-dec"
'--- names in UcsOpEnum order, latency measured in operations/sec
Private Const STR_OPS               As String = "x25519-keygen|x25519-derive|ed25519-sign|ed25519-verify|pbkdf2-sha256|hkdf-sha256|argon2id|scrypt"

Private Enum UcsAlgoEnum
    '--- hashes and MACs, fed straight from the input buffer
    ucsAlgoMd5
    ucsAlgoSha1
    ucsAlgoSha224
    ucsAlgoSha256
    ucsAlgoSha384
    ucsAlgoSha512
    ucsAlgoSha3256
    ucsAlgoSha3512
    ucsAlgoShake128
    ucsAlgoRipeMd160
    ucsAlgoBlake2s
    ucsAlgoBlake2b
    ucsAlgoBlake3
    ucsAlgoAsconHash
    ucsAlgoSiphash24
    ucsAlgoHalfSiphash24
    ucsAlgoHmacSha256
    ucsAlgoCmacAes128
    ucsAlgoGhash
    ucsAlgoPoly1305
    '--- ciphers from here on, these need a writable scratch buffer
    ucsAlgoAesCbc
    ucsAlgoAesCtr
    ucsAlgoAesGcm
    ucsAlgoAesCcm
    ucsAlgoAesEax
    ucsAlgoAesOcb
    ucsAlgoAesGcmSiv
    ucsAlgoChaCha20
    ucsAlgoChaCha20Poly1305
    ucsAlgoAsconAead
    ucsAlgoTea
    ucsAlgoAes256Cbc
    ucsAlgoAes256Ctr
    ucsAlgoAes256Gcm
    '--- decryption from here on, these need a real ciphertext to work on
    ucsAlgoAesCbcDec
    ucsAlgoAesGcmDec
    ucsAlgoAesCcmDec
    ucsAlgoAesEaxDec
    ucsAlgoAesOcbDec
    ucsAlgoAesGcmSivDec
    ucsAlgoChaCha20Poly1305Dec
    ucsAlgoAsconAeadDec
    ucsAlgoTeaDec
    ucsAlgoAes256GcmDec
End Enum

Private Enum UcsOpEnum
    ucsOpX25519KeyGen
    ucsOpX25519Derive
    ucsOpEd25519Sign
    ucsOpEd25519Verify
    ucsOpPbkdf2Sha256
    ucsOpHkdfSha256
    ucsOpArgon2Id
    ucsOpScrypt
End Enum

Private m_baKey8()                  As Byte
Private m_baKey16()                 As Byte
Private m_baKey32()                 As Byte
Private m_baNonce12()               As Byte
Private m_baNonce16()               As Byte
Private m_baAad()                   As Byte
Private m_baPriv()                  As Byte
Private m_baPub()                   As Byte

'=========================================================================
' Functions
'=========================================================================

Public Sub BenchList()
    ConPrintLine "Throughput (bytes/sec):"
    pvPrintNames STR_BYTES
    ConPrintLine
    ConPrintLine "Latency (operations/sec):"
    pvPrintNames STR_OPS
End Sub

Public Sub BenchRun(vFilter As Variant)
    Dim vSizes          As Variant
    Dim vNames          As Variant
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim baBuf()         As Byte
    Dim sLine           As String
    Dim dblRate         As Double
    Dim lRan            As Long
    Dim sSizes          As String
    Dim lMax            As Long

    pvInitKeys
    vFilter = pvParseSizeSwitch(vFilter, sSizes)
    If LenB(sSizes) = 0 Then
        sSizes = STR_SIZES
    End If
    vSizes = Split(sSizes, "|")
    '--- the sizes can arrive in any order, so find the largest
    For lIdx = 0 To UBound(vSizes)
        If CLng(vSizes(lIdx)) > lMax Then
            lMax = CLng(vSizes(lIdx))
        End If
    Next
    ReDim baBuf(0 To lMax - 1) As Byte
    For lIdx = 0 To UBound(baBuf)
        baBuf(lIdx) = lIdx And &HFF
    Next
    '--- throughput table, one row per algorithm and one column per size
    vNames = Split(STR_BYTES, "|")
    If pvHasAnyMatch(vNames, vFilter) Then
        sLine = PadRight("type", 22)
        For lJdx = 0 To UBound(vSizes)
            sLine = sLine & PadLeft(pvFormatSize(CLng(vSizes(lJdx))) & " bytes", 13)
        Next
        ConPrintLine sLine
        For lIdx = 0 To UBound(vNames)
            If HasFilterMatch(vNames(lIdx), vFilter) Then
                sLine = PadRight(vNames(lIdx), 22)
                For lJdx = 0 To UBound(vSizes)
                    dblRate = pvMeasureBytes(lIdx, baBuf, CLng(vSizes(lJdx)))
                    If dblRate < 0 Then
                        sLine = sLine & PadLeft("n/a", 13)
                    Else
                        sLine = sLine & PadLeft(FormatRate(dblRate) & "/s", 13)
                    End If
                Next
                ConPrintLine sLine
                lRan = lRan + 1
            End If
        Next
    End If
    '--- latency table for the operations that are not measured per byte
    vNames = Split(STR_OPS, "|")
    If pvHasAnyMatch(vNames, vFilter) Then
        If lRan > 0 Then
            ConPrintLine
        End If
        ConPrintLine PadRight("type", 22) & PadLeft("ops/sec", 13) & PadLeft("usec/op", 13)
        For lIdx = 0 To UBound(vNames)
            If HasFilterMatch(vNames(lIdx), vFilter) Then
                dblRate = pvMeasureOps(lIdx)
                sLine = PadRight(vNames(lIdx), 22)
                If dblRate <= 0 Then
                    sLine = sLine & PadLeft("n/a", 13) & PadLeft("n/a", 13)
                Else
                    sLine = sLine & PadLeft(FormatOps(dblRate), 13) & PadLeft(Format$(1000000# / dblRate, "0"), 13)
                End If
                ConPrintLine sLine
                lRan = lRan + 1
            End If
        Next
    End If
    If lRan = 0 Then
        ConPrintLine "No algorithm matched. Try ""vbcrypto list""."
    End If
End Sub

'--- pulls "-size 1M" or "-size 16,1K,1M" out of the filter list and returns
'--- what is left, so the rest still matches algorithm names as before
Private Function pvParseSizeSwitch(vFilter As Variant, sSizes As String) As Variant
    Dim aRetVal()       As String
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim vItem           As Variant
    Dim lSize           As Long

    sSizes = vbNullString
    If UBound(vFilter) < 0 Then
        pvParseSizeSwitch = vFilter
        Exit Function
    End If
    ReDim aRetVal(0 To UBound(vFilter)) As String
    For lIdx = 0 To UBound(vFilter)
        Select Case LCase$(vFilter(lIdx))
        Case "-size", "-sizes", "-bytes"
            If lIdx < UBound(vFilter) Then
                For Each vItem In Split(Replace(vFilter(lIdx + 1), ",", "|"), "|")
                    lSize = pvParseSize(CStr(vItem))
                    If lSize > 0 Then
                        sSizes = sSizes & IIf(LenB(sSizes) <> 0, "|", vbNullString) & lSize
                    End If
                Next
            End If
        Case Else
            '--- skip the value that follows the switch
            If lIdx = 0 Then
                aRetVal(lCount) = vFilter(lIdx)
                lCount = lCount + 1
            ElseIf Not pvIsSizeSwitch(CStr(vFilter(lIdx - 1))) Then
                aRetVal(lCount) = vFilter(lIdx)
                lCount = lCount + 1
            End If
        End Select
    Next
    If lCount = 0 Then
        pvParseSizeSwitch = Split(vbNullString)
    Else
        ReDim Preserve aRetVal(0 To lCount - 1) As String
        pvParseSizeSwitch = aRetVal
    End If
End Function

Private Function pvIsSizeSwitch(sText As String) As Boolean
    Select Case LCase$(sText)
    Case "-size", "-sizes", "-bytes"
        pvIsSizeSwitch = True
    End Select
End Function

'--- accepts a plain byte count or a K or M suffix
Private Function pvParseSize(sText As String) As Long
    Dim sTemp           As String
    Dim lMult           As Long

    sTemp = Trim$(sText)
    lMult = 1
    Select Case UCase$(Right$(sTemp, 1))
    Case "K"
        lMult = 1024
        sTemp = Left$(sTemp, Len(sTemp) - 1)
    Case "M"
        lMult = 1048576
        sTemp = Left$(sTemp, Len(sTemp) - 1)
    End Select
    If IsNumeric(sTemp) Then
        pvParseSize = CLng(sTemp) * lMult
    End If
End Function

Private Sub pvPrintNames(sList As String)
    Dim vNames          As Variant
    Dim lIdx            As Long
    Dim sLine           As String

    vNames = Split(sList, "|")
    For lIdx = 0 To UBound(vNames)
        sLine = sLine & PadRight(vNames(lIdx), 22)
        If (lIdx + 1) Mod 4 = 0 Then
            ConPrintLine "  " & RTrim$(sLine)
            sLine = vbNullString
        End If
    Next
    If LenB(RTrim$(sLine)) <> 0 Then
        ConPrintLine "  " & RTrim$(sLine)
    End If
End Sub

'--- compact column label, 1048576 becomes "1M"
Private Function pvFormatSize(ByVal lSize As Long) As String
    If lSize >= 1048576 And lSize Mod 1048576 = 0 Then
        pvFormatSize = lSize \ 1048576 & "M"
    ElseIf lSize >= 1024 And lSize Mod 1024 = 0 Then
        pvFormatSize = lSize \ 1024 & "K"
    Else
        pvFormatSize = CStr(lSize)
    End If
End Function

Private Function pvHasAnyMatch(vNames As Variant, vFilter As Variant) As Boolean
    Dim lIdx            As Long

    For lIdx = 0 To UBound(vNames)
        If HasFilterMatch(vNames(lIdx), vFilter) Then
            pvHasAnyMatch = True
            Exit Function
        End If
    Next
End Function

Private Sub pvInitKeys()
    Dim lIdx            As Long

    ReDim m_baKey8(0 To 7) As Byte
    ReDim m_baKey16(0 To 15) As Byte
    ReDim m_baKey32(0 To 31) As Byte
    ReDim m_baNonce12(0 To 11) As Byte
    ReDim m_baNonce16(0 To 15) As Byte
    ReDim m_baAad(0 To 19) As Byte
    For lIdx = 0 To 7
        m_baKey8(lIdx) = lIdx
    Next
    For lIdx = 0 To 15
        m_baKey16(lIdx) = lIdx
        m_baNonce16(lIdx) = 128 + lIdx
    Next
    For lIdx = 0 To 11
        m_baNonce12(lIdx) = 64 + lIdx
    Next
    For lIdx = 0 To 19
        m_baAad(lIdx) = 255 - lIdx
    Next
    For lIdx = 0 To 31
        m_baKey32(lIdx) = lIdx
    Next
    '--- a fixed key pair so the derive benchmark has something to chew on
    CryptoX25519PrivateKey m_baPriv, Seed:=m_baKey32
    CryptoX25519PublicKey m_baPub, m_baPriv
End Sub

'--- calibrate a batch size, then time one full batch. Returns bytes/sec,
'--- or -1 when the algorithm refused the workload.
Private Function pvMeasureBytes(ByVal eAlgo As UcsAlgoEnum, baBuf() As Byte, ByVal lSize As Long) As Double
    Dim lBatch          As Long
    Dim dblStart        As Double
    Dim dblElapsed      As Double

    On Error GoTo EH
    lBatch = 1
    Do
        dblStart = TimerEx
        pvRunBytes eAlgo, baBuf, lSize, lBatch
        dblElapsed = TimerEx - dblStart
        If dblElapsed >= DBL_PROBE Then
            Exit Do
        End If
        If lBatch > &H4000000 Then
            Exit Do
        End If
        lBatch = lBatch * 4
    Loop
    If dblElapsed > 0 Then
        lBatch = pvClampBatch(lBatch * (DBL_TARGET / dblElapsed))
    End If
    dblStart = TimerEx
    pvRunBytes eAlgo, baBuf, lSize, lBatch
    dblElapsed = TimerEx - dblStart
    If dblElapsed <= 0 Then
        pvMeasureBytes = -1
    Else
        pvMeasureBytes = lBatch * CDbl(lSize) / dblElapsed
    End If
    Exit Function
EH:
    pvMeasureBytes = -1
End Function

Private Function pvMeasureOps(ByVal eOp As UcsOpEnum) As Double
    Dim lBatch          As Long
    Dim dblStart        As Double
    Dim dblElapsed      As Double

    On Error GoTo EH
    lBatch = 1
    Do
        dblStart = TimerEx
        pvRunOps eOp, lBatch
        dblElapsed = TimerEx - dblStart
        If dblElapsed >= DBL_PROBE Then
            Exit Do
        End If
        If lBatch > &H100000 Then
            Exit Do
        End If
        lBatch = lBatch * 4
    Loop
    If dblElapsed > 0 Then
        lBatch = pvClampBatch(lBatch * (DBL_TARGET / dblElapsed))
    End If
    dblStart = TimerEx
    pvRunOps eOp, lBatch
    dblElapsed = TimerEx - dblStart
    If dblElapsed <= 0 Then
        pvMeasureOps = -1
    Else
        pvMeasureOps = lBatch / dblElapsed
    End If
    Exit Function
EH:
    pvMeasureOps = -1
End Function

Private Function pvClampBatch(ByVal dblBatch As Double) As Long
    If dblBatch < 1 Then
        pvClampBatch = 1
    ElseIf dblBatch > 100000000# Then
        pvClampBatch = 100000000
    Else
        pvClampBatch = CLng(dblBatch)
    End If
End Function

Private Sub pvRunBytes(ByVal eAlgo As UcsAlgoEnum, baBuf() As Byte, ByVal lSize As Long, ByVal lBatch As Long)
    Dim baWork()        As Byte
    Dim baCipher()      As Byte
    Dim baDecTag()      As Byte
    Dim lIter           As Long
    Dim baOut()         As Byte
    Dim baTag()         As Byte
    Dim uCmac           As CryptoCmacContext
    Dim uGhash          As CryptoGhashContext
    Dim uPoly           As CryptoPoly1305Context
    Dim uAes            As CryptoAesContext
    Dim uGcm            As CryptoAesGcmContext
    Dim uOcb            As CryptoAesOcbContext
    Dim uChaCha         As CryptoChaCha20Context
    Dim bOk             As Boolean

    '--- ciphers work in place, so hand them a scratch copy
    If eAlgo >= ucsAlgoAesCbc Then
        ReDim baWork(0 To lSize - 1) As Byte
        Call CopyMemory(baWork(0), baBuf(0), lSize)
    End If
    '--- decryption needs a ciphertext that authenticates, as a failing tag
    '--- check returns early and would time a fraction of the real work
    If eAlgo >= ucsAlgoAesCbcDec Then
        pvBuildCipherText eAlgo, baBuf, lSize, baCipher, baDecTag
    End If
    For lIter = 1 To lBatch
        Select Case eAlgo
        Case ucsAlgoMd5
            baOut = CryptoMd5ByteArray(baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoSha1
            baOut = CryptoSha1ByteArray(baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoSha224
            baOut = CryptoSha2ByteArray(224, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoSha256
            baOut = CryptoSha2ByteArray(256, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoSha384
            baOut = CryptoSha2ByteArray(384, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoSha512
            baOut = CryptoSha2ByteArray(512, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoSha3256
            baOut = CryptoSha3ByteArray(256, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoSha3512
            baOut = CryptoSha3ByteArray(512, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoShake128
            baOut = CryptoShakeByteArray(128, 32, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoRipeMd160
            baOut = CryptoRipeMd160ByteArray(baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoBlake2s
            baOut = CryptoBlake2sByteArray(256, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoBlake2b
            baOut = CryptoBlake2bByteArray(512, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoBlake3
            baOut = CryptoBlake3ByteArray(baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoAsconHash
            baOut = CryptoAsconHashByteArray(baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoSiphash24
            baOut = CryptoSiphash24ByteArray(m_baKey16, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoHalfSiphash24
            baOut = CryptoHalfSiphash24ByteArray(m_baKey8, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoHmacSha256
            baOut = CryptoHmacSha2ByteArray(256, m_baKey32, baBuf, Pos:=0, Size:=lSize)
        Case ucsAlgoCmacAes128
            CryptoCmacInit uCmac, m_baKey16
            CryptoCmacUpdate uCmac, baBuf, Pos:=0, Size:=lSize
            CryptoCmacFinalize uCmac, baTag
        Case ucsAlgoGhash
            CryptoGhashInit uGhash, m_baKey16
            CryptoGhashUpdate uGhash, baBuf, Pos:=0, Size:=lSize
            CryptoGhashFinalize uGhash, 16, baTag
        Case ucsAlgoPoly1305
            CryptoPoly1305Init uPoly, m_baKey32
            CryptoPoly1305Update uPoly, baBuf, Pos:=0, Size:=lSize
            CryptoPoly1305Finalize uPoly, baOut
        Case ucsAlgoAesCbc
            CryptoAesInit uAes, m_baKey16, Nonce:=m_baNonce16
            CryptoAesCbcEncrypt uAes, baWork, Pos:=0, Size:=lSize, Final:=False
        Case ucsAlgoAesCtr
            CryptoAesInit uAes, m_baKey16, Nonce:=m_baNonce16
            CryptoAesCtrCrypt uAes, baWork, Pos:=0, Size:=lSize
        Case ucsAlgoAesGcm
            CryptoAesGcmInit uGcm, m_baKey16, m_baNonce12, m_baAad
            CryptoAesGcmEncrypt uGcm, baWork, Pos:=0, Size:=lSize, TagSize:=16, Tag:=baTag
        Case ucsAlgoAesCcm
            CryptoAesCcmEncrypt m_baKey16, m_baNonce12, m_baAad, baWork, baTag, TagSize:=16
        Case ucsAlgoAesEax
            CryptoAesEaxEncrypt m_baKey16, m_baNonce16, m_baAad, baWork, baTag, TagSize:=16
        Case ucsAlgoAesOcb
            CryptoAesOcbInit uOcb, m_baKey16, m_baNonce12, m_baAad, TagSize:=16
            CryptoAesOcbEncrypt uOcb, baWork, Pos:=0, Size:=lSize, TagSize:=16, Tag:=baTag
        Case ucsAlgoAesGcmSiv
            CryptoAesGcmSivEncrypt m_baKey16, m_baNonce12, m_baAad, baWork, baTag
        Case ucsAlgoChaCha20
            CryptoChaCha20Init uChaCha, m_baKey32, m_baNonce12
            CryptoChaCha20Cipher uChaCha, baWork, Pos:=0, Size:=lSize
        Case ucsAlgoChaCha20Poly1305
            CryptoChaCha20Poly1305Encrypt m_baKey32, baTag, baWork, 0, lSize, m_baNonce12, m_baAad
        Case ucsAlgoAsconAead
            CryptoAsconEncrypt m_baKey16, baTag, baWork, 0, lSize, m_baNonce16, m_baAad
        Case ucsAlgoTea
            '--- TEA works on whole 8 byte blocks only
            CryptoTeaEncrypt m_baKey16, baWork, Pos:=0, Size:=(lSize \ 8) * 8
        Case ucsAlgoAes256Cbc
            CryptoAesInit uAes, m_baKey32, Nonce:=m_baNonce16
            CryptoAesCbcEncrypt uAes, baWork, Pos:=0, Size:=lSize, Final:=False
        Case ucsAlgoAes256Ctr
            CryptoAesInit uAes, m_baKey32, Nonce:=m_baNonce16
            CryptoAesCtrCrypt uAes, baWork, Pos:=0, Size:=lSize
        Case ucsAlgoAes256Gcm
            CryptoAesGcmInit uGcm, m_baKey32, m_baNonce12, m_baAad
            CryptoAesGcmEncrypt uGcm, baWork, Pos:=0, Size:=lSize, TagSize:=16, Tag:=baTag
        Case ucsAlgoAesCbcDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            CryptoAesInit uAes, m_baKey16, Nonce:=m_baNonce16
            bOk = CryptoAesCbcDecrypt(uAes, baWork, Pos:=0, Size:=lSize, Final:=False)
        Case ucsAlgoAesGcmDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            CryptoAesGcmInit uGcm, m_baKey16, m_baNonce12, m_baAad
            bOk = CryptoAesGcmDecrypt(uGcm, baWork, Pos:=0, Size:=lSize, Tag:=baDecTag)
        Case ucsAlgoAesCcmDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            bOk = CryptoAesCcmDecrypt(m_baKey16, m_baNonce12, m_baAad, baWork, baDecTag)
        Case ucsAlgoAesEaxDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            bOk = CryptoAesEaxDecrypt(m_baKey16, m_baNonce16, m_baAad, baWork, baDecTag)
        Case ucsAlgoAesOcbDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            CryptoAesOcbInit uOcb, m_baKey16, m_baNonce12, m_baAad, TagSize:=16
            bOk = CryptoAesOcbDecrypt(uOcb, baWork, Pos:=0, Size:=lSize, Tag:=baDecTag)
        Case ucsAlgoAesGcmSivDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            bOk = CryptoAesGcmSivDecrypt(m_baKey16, m_baNonce12, m_baAad, baWork, baDecTag)
        Case ucsAlgoChaCha20Poly1305Dec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            bOk = CryptoChaCha20Poly1305Decrypt(m_baKey32, baDecTag, baWork, 0, lSize, m_baNonce12, m_baAad)
        Case ucsAlgoAsconAeadDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            bOk = CryptoAsconDecrypt(m_baKey16, baDecTag, baWork, 0, lSize, m_baNonce16, m_baAad)
        Case ucsAlgoTeaDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            CryptoTeaDecrypt m_baKey16, baWork, Pos:=0, Size:=(lSize \ 8) * 8
        Case ucsAlgoAes256GcmDec
            Call CopyMemory(baWork(0), baCipher(0), lSize)
            CryptoAesGcmInit uGcm, m_baKey32, m_baNonce12, m_baAad
            bOk = CryptoAesGcmDecrypt(uGcm, baWork, Pos:=0, Size:=lSize, Tag:=baDecTag)
        Case Else
            Err.Raise vbObjectError, , "Unknown algorithm " & eAlgo
        End Select
    Next
End Sub

'--- encrypts the sample once and checks the result decrypts, so a broken
'--- or rejected ciphertext shows up as n/a rather than a flattering number
Private Sub pvBuildCipherText(ByVal eAlgo As UcsAlgoEnum, baBuf() As Byte, ByVal lSize As Long, baCipher() As Byte, baDecTag() As Byte)
    Dim baProbe()       As Byte
    Dim bOk             As Boolean
    Dim uAes            As CryptoAesContext
    Dim uGcm            As CryptoAesGcmContext
    Dim uOcb            As CryptoAesOcbContext

    ReDim baCipher(0 To lSize - 1) As Byte
    Call CopyMemory(baCipher(0), baBuf(0), lSize)
    Select Case eAlgo
    Case ucsAlgoAesCbcDec
        CryptoAesInit uAes, m_baKey16, Nonce:=m_baNonce16
        CryptoAesCbcEncrypt uAes, baCipher, Pos:=0, Size:=lSize, Final:=False
    Case ucsAlgoAesGcmDec
        CryptoAesGcmInit uGcm, m_baKey16, m_baNonce12, m_baAad
        CryptoAesGcmEncrypt uGcm, baCipher, Pos:=0, Size:=lSize, TagSize:=16, Tag:=baDecTag
    Case ucsAlgoAesCcmDec
        CryptoAesCcmEncrypt m_baKey16, m_baNonce12, m_baAad, baCipher, baDecTag, TagSize:=16
    Case ucsAlgoAesEaxDec
        CryptoAesEaxEncrypt m_baKey16, m_baNonce16, m_baAad, baCipher, baDecTag, TagSize:=16
    Case ucsAlgoAesOcbDec
        CryptoAesOcbInit uOcb, m_baKey16, m_baNonce12, m_baAad, TagSize:=16
        CryptoAesOcbEncrypt uOcb, baCipher, Pos:=0, Size:=lSize, TagSize:=16, Tag:=baDecTag
    Case ucsAlgoAesGcmSivDec
        CryptoAesGcmSivEncrypt m_baKey16, m_baNonce12, m_baAad, baCipher, baDecTag
    Case ucsAlgoChaCha20Poly1305Dec
        CryptoChaCha20Poly1305Encrypt m_baKey32, baDecTag, baCipher, 0, lSize, m_baNonce12, m_baAad
    Case ucsAlgoAsconAeadDec
        CryptoAsconEncrypt m_baKey16, baDecTag, baCipher, 0, lSize, m_baNonce16, m_baAad
    Case ucsAlgoTeaDec
        CryptoTeaEncrypt m_baKey16, baCipher, Pos:=0, Size:=(lSize \ 8) * 8
    Case ucsAlgoAes256GcmDec
        CryptoAesGcmInit uGcm, m_baKey32, m_baNonce12, m_baAad
        CryptoAesGcmEncrypt uGcm, baCipher, Pos:=0, Size:=lSize, TagSize:=16, Tag:=baDecTag
    End Select
    '--- prove the ciphertext round trips before any of it is timed
    ReDim baProbe(0 To lSize - 1) As Byte
    Call CopyMemory(baProbe(0), baCipher(0), lSize)
    Select Case eAlgo
    Case ucsAlgoAesCbcDec
        CryptoAesInit uAes, m_baKey16, Nonce:=m_baNonce16
        bOk = CryptoAesCbcDecrypt(uAes, baProbe, Pos:=0, Size:=lSize, Final:=False)
    Case ucsAlgoAesGcmDec
        CryptoAesGcmInit uGcm, m_baKey16, m_baNonce12, m_baAad
        bOk = CryptoAesGcmDecrypt(uGcm, baProbe, Pos:=0, Size:=lSize, Tag:=baDecTag)
    Case ucsAlgoAesCcmDec
        bOk = CryptoAesCcmDecrypt(m_baKey16, m_baNonce12, m_baAad, baProbe, baDecTag)
    Case ucsAlgoAesEaxDec
        bOk = CryptoAesEaxDecrypt(m_baKey16, m_baNonce16, m_baAad, baProbe, baDecTag)
    Case ucsAlgoAesOcbDec
        CryptoAesOcbInit uOcb, m_baKey16, m_baNonce12, m_baAad, TagSize:=16
        bOk = CryptoAesOcbDecrypt(uOcb, baProbe, Pos:=0, Size:=lSize, Tag:=baDecTag)
    Case ucsAlgoAesGcmSivDec
        bOk = CryptoAesGcmSivDecrypt(m_baKey16, m_baNonce12, m_baAad, baProbe, baDecTag)
    Case ucsAlgoChaCha20Poly1305Dec
        bOk = CryptoChaCha20Poly1305Decrypt(m_baKey32, baDecTag, baProbe, 0, lSize, m_baNonce12, m_baAad)
    Case ucsAlgoAsconAeadDec
        bOk = CryptoAsconDecrypt(m_baKey16, baDecTag, baProbe, 0, lSize, m_baNonce16, m_baAad)
    Case ucsAlgoTeaDec
        CryptoTeaDecrypt m_baKey16, baProbe, Pos:=0, Size:=(lSize \ 8) * 8
        bOk = True
    Case ucsAlgoAes256GcmDec
        CryptoAesGcmInit uGcm, m_baKey32, m_baNonce12, m_baAad
        bOk = CryptoAesGcmDecrypt(uGcm, baProbe, Pos:=0, Size:=lSize, Tag:=baDecTag)
    End Select
    If Not bOk Then
        Err.Raise vbObjectError, , "Decryption self check failed for " & eAlgo
    End If
End Sub

Private Sub pvRunOps(ByVal eOp As UcsOpEnum, ByVal lBatch As Long)
    Dim baPriv()        As Byte
    Dim baPub()         As Byte
    Dim baSig()         As Byte
    Dim lIter           As Long
    Dim baSecret()      As Byte
    Dim baOut()         As Byte
    Dim bOk             As Boolean

    '--- set up outside the timed loop whatever the operation only consumes
    Select Case eOp
    Case ucsOpEd25519Sign, ucsOpEd25519Verify
        CryptoEd25519PrivateKey baPriv, Seed:=m_baKey32
        CryptoEd25519PublicKey baPub, baPriv
        CryptoEd25519SignDetached baSig, baPriv, m_baKey32
        '--- a failing verify returns early and would be timed as a fast
        '--- success, so make sure the workload is the real one
        If Not CryptoEd25519VerifyDetached(baSig, baPub, m_baKey32) Then
            Err.Raise vbObjectError, , "Ed25519 self check failed"
        End If
    End Select
    For lIter = 1 To lBatch
        Select Case eOp
        Case ucsOpX25519KeyGen
            CryptoX25519PrivateKey baPriv, Seed:=m_baKey32
            CryptoX25519PublicKey baPub, baPriv
        Case ucsOpX25519Derive
            CryptoX25519SharedSecret baSecret, m_baPriv, m_baPub
        Case ucsOpEd25519Sign
            CryptoEd25519SignDetached baSig, baPriv, m_baKey32
        Case ucsOpEd25519Verify
            bOk = CryptoEd25519VerifyDetached(baSig, baPub, m_baKey32)
        Case ucsOpPbkdf2Sha256
            '--- a thousandth of a realistic password setting, for a usable rate
            baOut = CryptoPbkdf2HmacSha2ByteArray(256, m_baKey32, m_baKey16, OutSize:=32, NumIter:=1000)
        Case ucsOpHkdfSha256
            baOut = CryptoHkdfSha2ByteArray(256, m_baKey32, m_baKey16, m_baAad, OutSize:=32)
        Case ucsOpArgon2Id
            '--- deliberately small parameters, this is a rate not a recommendation
            baOut = CryptoArgon2IdKdfByteArray(m_baKey32, m_baKey16, OutSize:=32, Passes:=1, Memory:=1024, Parallelism:=1)
        Case ucsOpScrypt
            baOut = CryptoScryptKdfByteArray(m_baKey32, m_baKey16, OutSize:=32, Passes:=8, Memory:=1024, Parallelism:=1)
        Case Else
            Err.Raise vbObjectError, , "Unknown operation " & eOp
        End Select
    Next
End Sub
