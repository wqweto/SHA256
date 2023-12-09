Attribute VB_Name = "mdAES"
'--- mdAES.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
#End If

Private Const LNG_BLOCKSZ               As Long = 16
Private Const LNG_POW2_1                As Long = 2 ^ 1
Private Const LNG_POW2_2                As Long = 2 ^ 2
Private Const LNG_POW2_3                As Long = 2 ^ 3
Private Const LNG_POW2_4                As Long = 2 ^ 4
Private Const LNG_POW2_7                As Long = 2 ^ 7
Private Const LNG_POW2_8                As Long = 2 ^ 8
Private Const LNG_POW2_16               As Long = 2 ^ 16
Private Const LNG_POW2_23               As Long = 2 ^ 23
Private Const LNG_POW2_24               As Long = 2 ^ 24

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    cElements           As Long
    lLbound             As Long
End Type

Private Type ArrayLong256
    Item(0 To 255)     As Long
End Type

Private Type ArrayLong60
    Item(0 To 59)       As Long
End Type

Private Type AesTables
    Item(0 To 4)        As ArrayLong256
End Type

Private Type AesBlock
    Item(0 To 3)        As Long
End Type

Private m_uEncTables                As AesTables
Private m_uDecTables                As AesTables
Private m_aBlock()                  As AesBlock
Private m_uPeekBlock                As SAFEARRAY1D

Public Type CryptoAesContext
    KeyLen              As Long
    EncKey              As ArrayLong60
    DecKey              As ArrayLong60
    Nonce               As AesBlock
End Type

Private Function BSwap32(ByVal lX As Long) As Long
    BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
                 (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
End Function

Private Sub pvInit(uEncTable As AesTables, uDecTable As AesTables)
    Const FADF_AUTO     As Long = 1
    Dim lIdx            As Long
    Dim uDbl            As ArrayLong256
    Dim uThd            As ArrayLong256
    Dim lX              As Long
    Dim lX2             As Long
    Dim lX4             As Long
    Dim lX8             As Long
    Dim lXInv           As Long
    Dim lS              As Long
    Dim lDec            As Long
    Dim lEnc            As Long
    Dim lTemp           As Long
    Dim pDummy          As LongPtr
    
    '--- double and third tables
    For lIdx = 0 To 255
        #If HasOperators Then
            lTemp = (lIdx << 1) Xor (lIdx >> 7) * 283
        #Else
            lTemp = (lIdx * LNG_POW2_1) Xor (lIdx \ LNG_POW2_7) * 283
        #End If
        uDbl.Item(lIdx) = lTemp
        uThd.Item(lTemp Xor lIdx) = lIdx
    Next
    Do While uEncTable.Item(4).Item(lX) = 0
        '--- sbox
        lS = lXInv Xor lXInv * LNG_POW2_1 Xor lXInv * LNG_POW2_2 Xor lXInv * LNG_POW2_3 Xor lXInv * LNG_POW2_4
        #If HasOperators Then
            lS = (lS >> 8) Xor (lS And 255) Xor 99
        #Else
            lS = (lS \ LNG_POW2_8) Xor (lS And 255) Xor 99
        #End If
        #If HasOperators Then
            uEncTable.Item(4).Item(lX) = lS * &H1010101
            uDecTable.Item(4).Item(lS) = lX * &H1010101
        #Else
            uEncTable.Item(4).Item(lX) = (lS And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lS And LNG_POW2_7) <> 0) * &H80000000 Or lS * &H10101
            uDecTable.Item(4).Item(lS) = (lX And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lX And LNG_POW2_7) <> 0) * &H80000000 Or lX * &H10101
        #End If
        '--- mixcolumns
        lX2 = uDbl.Item(lX)
        lX4 = uDbl.Item(lX2)
        lX8 = uDbl.Item(lX4)
        #If HasOperators Then
            lDec = lX8 * &H1010101 Xor lX4 * &H10001 Xor lX2 * &H101& Xor lX * &H1010100
            lEnc = uDbl.Item(lS) * &H101& Xor lS * &H1010100
        #Else
            lDec = ((lX8 And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lX8 And LNG_POW2_7) <> 0) * &H80000000 Or lX8 * &H10101) _
                Xor lX4 * &H10001 _
                Xor lX2 * &H101& _
                Xor ((lX And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lX And LNG_POW2_7) <> 0) * &H80000000 Or lX * &H10100)
            lEnc = uDbl.Item(lS) * &H101& _
                Xor ((lS And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lS And LNG_POW2_7) <> 0) * &H80000000 Or lS * &H10100)
        #End If
        For lIdx = 0 To 3
            #If HasOperators Then
                lEnc = (lEnc << 24) Xor (lEnc >> 8)
                lDec = (lDec << 24) Xor (lDec >> 8)
            #Else
                lEnc = ((lEnc And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lEnc And LNG_POW2_7) <> 0) * &H80000000) _
                    Xor ((lEnc And &H7FFFFFFF) \ LNG_POW2_8 Or -(lEnc < 0) * LNG_POW2_23)
                lDec = ((lDec And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lDec And LNG_POW2_7) <> 0) * &H80000000) _
                    Xor ((lDec And &H7FFFFFFF) \ LNG_POW2_8 Or -(lDec < 0) * LNG_POW2_23)
            #End If
            uEncTable.Item(lIdx).Item(lX) = lEnc
            uDecTable.Item(lIdx).Item(lS) = lDec
        Next
        If lX2 <> 0 Then
            lX = lX Xor lX2
        Else
            lX = lX Xor 1
        End If
        lXInv = uThd.Item(lXInv)
        If lXInv = 0 Then
            lXInv = 1
        End If
    Loop
    With m_uPeekBlock
        .cDims = 1
        .fFeatures = FADF_AUTO
        .cbElements = 16
        .cLocks = 1
    End With
    Call CopyMemory(ByVal ArrPtr(m_aBlock), VarPtr(m_uPeekBlock), LenB(pDummy))
End Sub

Private Sub pvInitPeek(uArray As SAFEARRAY1D, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    With uArray
        If Size > 0 Then
            .pvData = VarPtr(baBuffer(Pos))
        Else
            .pvData = 0
        End If
        .cElements = Size \ .cbElements
    End With
End Sub

Private Function pvKeySchedule(baKey() As Byte, uSbox As ArrayLong256, uDecTable As AesTables, uEncKey As ArrayLong60, uDecKey As ArrayLong60) As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lRCon           As Long
    Dim lKeyLen         As Long
    Dim lPrev           As Long
    Dim lTemp           As Long
    
    lRCon = 1
    lKeyLen = (UBound(baKey) + 1) \ 4
    If Not (lKeyLen = 4 Or lKeyLen = 6 Or lKeyLen = 8) Then
        Err.Raise vbObjectError, , "Invalid key bit-size for AES (" & lKeyLen * 8 & ")"
    End If
    Call CopyMemory(uEncKey.Item(0), baKey(0), lKeyLen * 4)
    For lIdx = 0 To lKeyLen - 1
        uEncKey.Item(lIdx) = BSwap32(uEncKey.Item(lIdx))
    Next
    For lIdx = lKeyLen To 4 * lKeyLen + 27
        lPrev = uEncKey.Item(lIdx - 1)
        '--- sbox
        If lIdx Mod lKeyLen = 0 Or lIdx Mod lKeyLen = 4 And lKeyLen = 8 Then
            #If HasOperators Then
                lPrev = (uSbox.Item(lPrev >> 24) And &HFF000000) _
                    Xor (uSbox.Item((lPrev >> 16) And 255) And &HFF0000) _
                    Xor (uSbox.Item((lPrev >> 8) And 255) And &HFF00&) _
                    Xor (uSbox.Item(lPrev And 255) And &HFF&)
                If lIdx Mod lKeyLen = 0 Then
                    lPrev = (lPrev << 8) Xor (lPrev >> 24) Xor (lRCon << 24)
                    lRCon = (lRCon << 1) Xor (lRCon >> 7) * 283
                End If
            #Else
                lPrev = (uSbox.Item((lPrev And &H7F000000) \ LNG_POW2_24 Or -(lPrev < 0) * LNG_POW2_7) And &HFF000000) _
                    Xor (uSbox.Item((lPrev And &HFF0000) \ LNG_POW2_16) And &HFF0000) _
                    Xor (uSbox.Item((lPrev And &HFF00&) \ LNG_POW2_8) And &HFF00&) _
                    Xor (uSbox.Item(lPrev And 255) And &HFF&)
                If lIdx Mod lKeyLen = 0 Then
                    lPrev = ((lPrev And (LNG_POW2_23 - 1)) * LNG_POW2_8 Or -((lPrev And LNG_POW2_23) <> 0) * &H80000000) _
                        Xor ((lPrev And &H7FFFFFFF) \ LNG_POW2_24 Or -(lPrev < 0) * LNG_POW2_7) _
                        Xor ((lRCon And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lRCon And LNG_POW2_7) <> 0) * &H80000000)
                    lRCon = lRCon * LNG_POW2_1 Xor (lRCon \ LNG_POW2_7) * 283
                End If
            #End If
        End If
        uEncKey.Item(lIdx) = uEncKey.Item(lIdx - lKeyLen) Xor lPrev
    Next
    pvKeySchedule = lIdx
    For lJdx = 0 To lIdx - 1
        If (lIdx And 3) <> 0 Then
            lPrev = uEncKey.Item(lIdx)
        Else
            lPrev = uEncKey.Item(lIdx - 4)
        End If
        If lIdx <= 4 Or lJdx < 4 Then
            uDecKey.Item(lJdx) = lPrev
        Else
            #If HasOperators Then
                uDecKey.Item(lJdx) = uDecTable.Item(0).Item(uSbox.Item(lPrev >> 24) And &HFF&) _
                    Xor uDecTable.Item(1).Item(uSbox.Item((lPrev >> 16) And 255) And &HFF&) _
                    Xor uDecTable.Item(2).Item(uSbox.Item((lPrev >> 8) And 255) And &HFF&) _
                    Xor uDecTable.Item(3).Item(uSbox.Item(lPrev And 255) And &HFF&)
            #Else
                lTemp = (lPrev And &H7FFFFFFF) \ LNG_POW2_24 Or -(lPrev < 0) * LNG_POW2_7
                uDecKey.Item(lJdx) = uDecTable.Item(0).Item(uSbox.Item(lTemp) And &HFF&) _
                    Xor uDecTable.Item(1).Item(uSbox.Item((lPrev And &HFF0000) \ LNG_POW2_16) And &HFF&) _
                    Xor uDecTable.Item(2).Item(uSbox.Item((lPrev And &HFF00&) \ LNG_POW2_8) And &HFF&) _
                    Xor uDecTable.Item(3).Item(uSbox.Item(lPrev And 255) And &HFF&)
            #End If
        End If
        lIdx = lIdx - 1
    Next
End Function

Private Sub pvCrypt(uInput As AesBlock, uOutput As AesBlock, ByVal bDecrypt As Boolean, uKey As ArrayLong60, ByVal lKeyLen As Long, _
            uT0 As ArrayLong256, uT1 As ArrayLong256, uT2 As ArrayLong256, uT3 As ArrayLong256, uSbox As ArrayLong256)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lKdx            As Long
    Dim lA              As Long
    Dim lB              As Long
    Dim lC              As Long
    Dim lD              As Long
    Dim lTemp1          As Long
    Dim lTemp2          As Long
    Dim lTemp3          As Long

    '--- first round
    lA = uInput.Item(0) Xor uKey.Item(0)
    lB = uInput.Item(1 - bDecrypt * 2) Xor uKey.Item(1)
    lC = uInput.Item(2) Xor uKey.Item(2)
    lD = uInput.Item(3 + bDecrypt * 2) Xor uKey.Item(3)
    '--- inner rounds
    lKdx = 4
    For lIdx = 1 To lKeyLen \ 4 - 2
        #If HasOperators Then
            lTemp1 = uT0.Item(lA >> 24) Xor uT1.Item((lB >> 16) And 255) Xor uT2.Item((lC >> 8) And 255) Xor uT3.Item(lD And 255) Xor uKey.Item(lKdx + 0)
            lTemp2 = uT0.Item(lB >> 24) Xor uT1.Item((lC >> 16) And 255) Xor uT2.Item((lD >> 8) And 255) Xor uT3.Item(lA And 255) Xor uKey.Item(lKdx + 1)
            lTemp3 = uT0.Item(lC >> 24) Xor uT1.Item((lD >> 16) And 255) Xor uT2.Item((lA >> 8) And 255) Xor uT3.Item(lB And 255) Xor uKey.Item(lKdx + 2)
            lD = uT0.Item(lD >> 24) Xor uT1.Item((lA >> 16) And 255) Xor uT2.Item((lB >> 8) And 255) Xor uT3.Item(lC And 255) Xor uKey.Item(lKdx + 3)
        #Else
            lTemp1 = uT0.Item((lA And &H7F000000) \ LNG_POW2_24 Or -(lA < 0) * LNG_POW2_7) _
                Xor uT1.Item((lB And &HFF0000) \ LNG_POW2_16) _
                Xor uT2.Item((lC And &HFF00&) \ LNG_POW2_8) _
                Xor uT3.Item(lD And 255) Xor uKey.Item(lKdx + 0)
            lTemp2 = uT0.Item((lB And &H7F000000) \ LNG_POW2_24 Or -(lB < 0) * LNG_POW2_7) _
                Xor uT1.Item((lC And &HFF0000) \ LNG_POW2_16) _
                Xor uT2.Item((lD And &HFF00&) \ LNG_POW2_8) _
                Xor uT3.Item(lA And 255) Xor uKey.Item(lKdx + 1)
            lTemp3 = uT0.Item((lC And &H7F000000) \ LNG_POW2_24 Or -(lC < 0) * LNG_POW2_7) _
                Xor uT1.Item((lD And &HFF0000) \ LNG_POW2_16) _
                Xor uT2.Item((lA And &HFF00&) \ LNG_POW2_8) _
                Xor uT3.Item(lB And 255) Xor uKey.Item(lKdx + 2)
            lD = uT0.Item((lD And &H7F000000) \ LNG_POW2_24 Or -(lD < 0) * LNG_POW2_7) _
                Xor uT1.Item((lA And &HFF0000) \ LNG_POW2_16) _
                Xor uT2.Item((lB And &HFF00&) \ LNG_POW2_8) _
                Xor uT3.Item(lC And 255) Xor uKey.Item(lKdx + 3)
        #End If
        lKdx = lKdx + 4
        lA = lTemp1: lB = lTemp2: lC = lTemp3
    Next
    '--- last round
    For lIdx = 0 To 3
        If bDecrypt Then
            lJdx = -lIdx And 3
        Else
            lJdx = lIdx
        End If
        #If HasOperators Then
            uOutput.Item(lJdx) = (uSbox.Item(lA >> 24) And &HFF000000) _
                Xor (uSbox.Item((lB >> 16) And 255) And &HFF0000) _
                Xor (uSbox.Item((lC >> 8) And 255) And &HFF00&) _
                Xor (uSbox.Item(lD And 255) And &HFF&) Xor uKey.Item(lKdx)
        #Else
            lTemp1 = (lA And &H7F000000) \ LNG_POW2_24 Or -(lA < 0) * LNG_POW2_7
            uOutput.Item(lJdx) = (uSbox.Item(lTemp1) And &HFF000000) _
                Xor (uSbox.Item((lB And &HFF0000) \ LNG_POW2_16) And &HFF0000) _
                Xor (uSbox.Item((lC And &HFF00&) \ LNG_POW2_8) And &HFF00&) _
                Xor (uSbox.Item(lD And 255) And &HFF&) Xor uKey.Item(lKdx)
        #End If
        lKdx = lKdx + 1
        lTemp1 = lA: lA = lB: lB = lC: lC = lD: lD = lTemp1
    Next
End Sub

Private Sub pvProcess(uCtx As CryptoAesContext, ByVal bEncrypt As Boolean, uInput As AesBlock, uOutput As AesBlock)
    If bEncrypt Then
        pvCrypt uInput, uOutput, False, uCtx.EncKey, uCtx.KeyLen, m_uEncTables.Item(0), m_uEncTables.Item(1), m_uEncTables.Item(2), m_uEncTables.Item(3), m_uEncTables.Item(4)
    Else
        pvCrypt uInput, uOutput, True, uCtx.DecKey, uCtx.KeyLen, m_uDecTables.Item(0), m_uDecTables.Item(1), m_uDecTables.Item(2), m_uDecTables.Item(3), m_uDecTables.Item(4)
    End If
End Sub

Private Function pvUnsignedInc(lValue As Long) As Boolean
    If lValue <> -1 Then
        lValue = (lValue Xor &H80000000) + 1 Xor &H80000000
    Else
        lValue = 0
        '--- signal carry
        pvUnsignedInc = True
    End If
End Function

Public Sub CryptoAesInit(uCtx As CryptoAesContext, baKey() As Byte, Optional Nonce As Variant)
    Dim baNonce()       As Byte
    
    If m_uEncTables.Item(0).Item(0) = 0 Then
        pvInit m_uEncTables, m_uDecTables
    End If
    With uCtx
        .KeyLen = pvKeySchedule(baKey, m_uEncTables.Item(4), m_uDecTables, .EncKey, .DecKey)
        If IsMissing(Nonce) Or IsNumeric(Nonce) Then
            baNonce = vbNullString
        Else
            baNonce = Nonce
        End If
        If UBound(baNonce) <> LNG_BLOCKSZ - 1 Then
            ReDim Preserve baNonce(0 To LNG_BLOCKSZ - 1) As Byte
        End If
        Call CopyMemory(.Nonce, baNonce(0), LNG_BLOCKSZ)
        With .Nonce
            .Item(0) = BSwap32(.Item(0))
            .Item(1) = BSwap32(.Item(1))
            .Item(2) = BSwap32(.Item(2))
            If IsNumeric(Nonce) Then
                .Item(3) = Nonce
            Else
                .Item(3) = BSwap32(.Item(3))
            End If
        End With
    End With
End Sub

Public Sub CryptoAesProcess(uCtx As CryptoAesContext, ByVal Encrypt As Boolean, baBlock() As Byte, Optional ByVal Pos As Long)
    Dim uBlock          As AesBlock
    
    Debug.Assert UBound(baBlock) + 1 >= Pos + LNG_BLOCKSZ
    pvInitPeek m_uPeekBlock, baBlock, Pos, LNG_BLOCKSZ
    With uBlock
        .Item(0) = BSwap32(m_aBlock(0).Item(0))
        .Item(1) = BSwap32(m_aBlock(0).Item(1))
        .Item(2) = BSwap32(m_aBlock(0).Item(2))
        .Item(3) = BSwap32(m_aBlock(0).Item(3))
    End With
    pvProcess uCtx, Encrypt, uBlock, uBlock
    With m_aBlock(0)
        .Item(0) = BSwap32(uBlock.Item(0))
        .Item(1) = BSwap32(uBlock.Item(1))
        .Item(2) = BSwap32(uBlock.Item(2))
        .Item(3) = BSwap32(uBlock.Item(3))
    End With
End Sub

Public Sub CryptoAesCbcEncrypt(uCtx As CryptoAesContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional ByVal Final As Boolean = True)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lNumBlocks      As Long
    Dim uBlock          As AesBlock
    Dim lPad            As Long
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    If Final Then
        lNumBlocks = Size \ LNG_BLOCKSZ
    Else
        If Size Mod LNG_BLOCKSZ <> 0 Then
            Err.Raise vbObjectError, , "Invalid non-final block size for CBC mode (" & Size Mod LNG_BLOCKSZ & ")"
        End If
        lNumBlocks = Size \ LNG_BLOCKSZ - 1
    End If
    pvInitPeek m_uPeekBlock, baBuffer, Pos, Size
    For lIdx = 0 To lNumBlocks
        If lIdx = lNumBlocks And Final Then
            '--- append PKCS#5 padding
            lPad = (LNG_BLOCKSZ - Size Mod LNG_BLOCKSZ) * &H1010101
            uBlock.Item(0) = lPad: uBlock.Item(1) = lPad: uBlock.Item(2) = lPad: uBlock.Item(3) = lPad
            lJdx = lIdx * LNG_BLOCKSZ
            If Size - lJdx > 0 Then
                Call CopyMemory(uBlock, baBuffer(Pos + lJdx), Size - lJdx)
            End If
            ReDim Preserve baBuffer(0 To Pos + lJdx + LNG_BLOCKSZ - 1) As Byte
            pvInitPeek m_uPeekBlock, baBuffer, Pos, lJdx + LNG_BLOCKSZ
            m_aBlock(lIdx) = uBlock
        End If
        With uCtx.Nonce
            .Item(0) = .Item(0) Xor BSwap32(m_aBlock(lIdx).Item(0))
            .Item(1) = .Item(1) Xor BSwap32(m_aBlock(lIdx).Item(1))
            .Item(2) = .Item(2) Xor BSwap32(m_aBlock(lIdx).Item(2))
            .Item(3) = .Item(3) Xor BSwap32(m_aBlock(lIdx).Item(3))
        End With
        pvProcess uCtx, True, uCtx.Nonce, uCtx.Nonce
        With m_aBlock(lIdx)
            .Item(0) = BSwap32(uCtx.Nonce.Item(0))
            .Item(1) = BSwap32(uCtx.Nonce.Item(1))
            .Item(2) = BSwap32(uCtx.Nonce.Item(2))
            .Item(3) = BSwap32(uCtx.Nonce.Item(3))
        End With
    Next
End Sub

Public Function CryptoAesCbcDecrypt(uCtx As CryptoAesContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional ByVal Final As Boolean = True) As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lNumBlocks      As Long
    Dim uInput          As AesBlock
    Dim uBlock          As AesBlock
    Dim lPad            As Long
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    If Size Mod LNG_BLOCKSZ <> 0 Then
        Err.Raise vbObjectError, , "Invalid partial block size for CBC mode (" & Size Mod LNG_BLOCKSZ & ")"
    End If
    lNumBlocks = Size \ LNG_BLOCKSZ - 1
    pvInitPeek m_uPeekBlock, baBuffer, Pos, Size
    For lIdx = 0 To lNumBlocks
        With uInput
            .Item(0) = BSwap32(m_aBlock(lIdx).Item(0))
            .Item(1) = BSwap32(m_aBlock(lIdx).Item(1))
            .Item(2) = BSwap32(m_aBlock(lIdx).Item(2))
            .Item(3) = BSwap32(m_aBlock(lIdx).Item(3))
        End With
        pvProcess uCtx, False, uInput, uBlock
        With uBlock
            .Item(0) = .Item(0) Xor uCtx.Nonce.Item(0)
            .Item(1) = .Item(1) Xor uCtx.Nonce.Item(1)
            .Item(2) = .Item(2) Xor uCtx.Nonce.Item(2)
            .Item(3) = .Item(3) Xor uCtx.Nonce.Item(3)
        End With
        uCtx.Nonce = uInput
        With m_aBlock(lIdx)
            .Item(0) = BSwap32(uBlock.Item(0))
            .Item(1) = BSwap32(uBlock.Item(1))
            .Item(2) = BSwap32(uBlock.Item(2))
            .Item(3) = BSwap32(uBlock.Item(3))
        End With
        If lIdx = lNumBlocks And Final Then
            Pos = Pos + lIdx * LNG_BLOCKSZ
            '--- check and remove PKCS#5 padding
            lPad = baBuffer(Pos + LNG_BLOCKSZ - 1)
            If lPad = 0 Or lPad > LNG_BLOCKSZ Then
                Exit Function
            End If
            For lJdx = 1 To lPad
                If baBuffer(Pos + LNG_BLOCKSZ - lJdx) <> lPad Then
                    Exit Function
                End If
            Next
            Pos = Pos + LNG_BLOCKSZ - lPad
            If Pos = 0 Then
                baBuffer = vbNullString
            Else
                ReDim Preserve baBuffer(0 To Pos - 1) As Byte
            End If
        End If
    Next
    '--- success
    CryptoAesCbcDecrypt = True
End Function

Public Sub CryptoAesCtrCrypt(uCtx As CryptoAesContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lFinal          As Long
    Dim uBlock          As AesBlock
    Dim uTemp           As AesBlock
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    lFinal = Size \ LNG_BLOCKSZ
    pvInitPeek m_uPeekBlock, baBuffer, Pos, Size
    For lIdx = 0 To (Size - 1) \ LNG_BLOCKSZ
        pvProcess uCtx, True, uCtx.Nonce, uBlock
        If lIdx = lFinal Then
            lJdx = lIdx * LNG_BLOCKSZ
            Call CopyMemory(uTemp, baBuffer(Pos + lJdx), Size - lJdx)
            With uTemp
                .Item(0) = .Item(0) Xor BSwap32(uBlock.Item(0))
                .Item(1) = .Item(1) Xor BSwap32(uBlock.Item(1))
                .Item(2) = .Item(2) Xor BSwap32(uBlock.Item(2))
                .Item(3) = .Item(3) Xor BSwap32(uBlock.Item(3))
            End With
            Call CopyMemory(baBuffer(Pos + lJdx), uTemp, Size - lJdx)
        Else
            With m_aBlock(lIdx)
                .Item(0) = .Item(0) Xor BSwap32(uBlock.Item(0))
                .Item(1) = .Item(1) Xor BSwap32(uBlock.Item(1))
                .Item(2) = .Item(2) Xor BSwap32(uBlock.Item(2))
                .Item(3) = .Item(3) Xor BSwap32(uBlock.Item(3))
            End With
        End If
        For lJdx = 3 To 0 Step -1
            If Not pvUnsignedInc(uCtx.Nonce.Item(lJdx)) Then
                Exit For
            End If
        Next
    Next
End Sub
