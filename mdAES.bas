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
Private Const LNG_POLY                  As Long = &H11B
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
    #If Not HasOperators Then
        BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
                  (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
    #Else
        Return ((lX And &H000000FF&) << 24) Or _
               ((lX And &H0000FF00&) << 8) Or _
               ((lX And &H00FF0000&) >> 8) Or _
               ((lX And &HFF000000&) >> 24)
    #End If
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
            lTemp = (lIdx << 1) Xor (lIdx >> 7) * LNG_POLY
        #Else
            lTemp = (lIdx * LNG_POW2_1) Xor (lIdx \ LNG_POW2_7) * LNG_POLY
        #End If
        uDbl.Item(lIdx) = lTemp
        uThd.Item(lTemp Xor lIdx) = lIdx
    Next
    Do While uEncTable.Item(4).Item(lX) = 0
        '--- sbox
        lS = lXInv Xor lXInv * LNG_POW2_1 Xor lXInv * LNG_POW2_2 Xor lXInv * LNG_POW2_3 Xor lXInv * LNG_POW2_4
        #If HasOperators Then
            lS = (lS >> 8) Xor (lS And 255) Xor &H63
        #Else
            lS = (lS \ LNG_POW2_8) Xor (lS And 255) Xor &H63
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
            lDec = lX8 * &H1010101 Xor lX4 * &H1000100 Xor lX2 * &H1010000 Xor lX * &H10101
            lEnc = uDbl.Item(lS) * &H1010000 Xor lS * &H10101
        #Else
            lDec = ((lX8 And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lX8 And LNG_POW2_7) <> 0) * &H80000000 Or lX8 * &H10101) _
                Xor ((lX4 And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lX4 And LNG_POW2_7) <> 0) * &H80000000 Or lX4 * &H100) _
                Xor ((lX2 And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lX2 And LNG_POW2_7) <> 0) * &H80000000 Or lX2 * &H10000) _
                Xor lX * &H10101
            lEnc = ((uDbl.Item(lS) And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((uDbl.Item(lS) And LNG_POW2_7) <> 0) * &H80000000 Or uDbl.Item(lS) * &H10000) _
                Xor lS * &H10101
        #End If
        For lIdx = 0 To 3
            #If HasOperators Then
                lEnc = (lEnc << 8) Xor (lEnc >> 24)
                lDec = (lDec << 8) Xor (lDec >> 24)
            #Else
                lEnc = ((lEnc And (LNG_POW2_23 - 1)) * LNG_POW2_8 Or -((lEnc And LNG_POW2_23) <> 0) * &H80000000) _
                    Xor ((lEnc And &H7FFFFFFF) \ LNG_POW2_24 Or -(lEnc < 0) * LNG_POW2_7)
                lDec = ((lDec And (LNG_POW2_23 - 1)) * LNG_POW2_8 Or -((lDec And LNG_POW2_23) <> 0) * &H80000000) _
                    Xor ((lDec And &H7FFFFFFF) \ LNG_POW2_24 Or -(lDec < 0) * LNG_POW2_7)
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
    
    lKeyLen = (UBound(baKey) + 1) \ 4
    If Not (lKeyLen = 4 Or lKeyLen = 6 Or lKeyLen = 8) Then
        Err.Raise vbObjectError, , "Invalid key bit-size for AES (" & lKeyLen * 32 & ")"
    End If
    lRCon = 1
    Call CopyMemory(uEncKey.Item(0), baKey(0), lKeyLen * 4)
    For lIdx = lKeyLen To 4 * lKeyLen + 27
        lPrev = uEncKey.Item(lIdx - 1)
        '--- sbox
        If lIdx Mod lKeyLen = 0 Then
            #If HasOperators Then
                lPrev = (lPrev << 24) Or (lPrev >> 8)
                lPrev = (uSbox.Item(lPrev And &HFF&) And &HFF&) _
                    Xor (uSbox.Item((lPrev >> 8) And &HFF&) And &HFF00&) _
                    Xor (uSbox.Item((lPrev >> 16) And &HFF&) And &HFF0000) _
                    Xor (uSbox.Item((lPrev >> 24) And &HFF&) And &HFF000000) Xor lRCon
                lRCon = (lRCon << 1) Xor (lRCon >> 7) * LNG_POLY
            #Else
                lPrev = ((lPrev And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lPrev And LNG_POW2_7) <> 0) * &H80000000) _
                    Xor ((lPrev And &H7FFFFFFF) \ LNG_POW2_8 Or -(lPrev < 0) * LNG_POW2_23)
                lPrev = (uSbox.Item(lPrev And &HFF&) And &HFF&) _
                    Xor (uSbox.Item((lPrev And &HFF00&) \ LNG_POW2_8) And &HFF00&) _
                    Xor (uSbox.Item((lPrev And &HFF0000) \ LNG_POW2_16) And &HFF0000) _
                    Xor (uSbox.Item((lPrev And &H7F000000) \ LNG_POW2_24 Or -(lPrev < 0) * LNG_POW2_7) And &HFF000000) Xor lRCon
                lRCon = lRCon * LNG_POW2_1 Xor (lRCon \ LNG_POW2_7) * LNG_POLY
            #End If
        ElseIf lIdx Mod lKeyLen = 4 And lKeyLen > 6 Then
            #If HasOperators Then
                lPrev = (uSbox.Item(lPrev And 255) And &HFF&) _
                    Xor (uSbox.Item((lPrev >> 8) And 255) And &HFF00&) _
                    Xor (uSbox.Item((lPrev >> 16) And 255) And &HFF0000) _
                    Xor (uSbox.Item(lPrev >> 24) And &HFF000000)
            #Else
                lPrev = (uSbox.Item(lPrev And &HFF&) And &HFF&) _
                    Xor (uSbox.Item((lPrev And &HFF00&) \ LNG_POW2_8) And &HFF00&) _
                    Xor (uSbox.Item((lPrev And &HFF0000) \ LNG_POW2_16) And &HFF0000) _
                    Xor (uSbox.Item((lPrev And &H7F000000) \ LNG_POW2_24 Or -(lPrev < 0) * LNG_POW2_7) And &HFF000000)
            #End If
        End If
        uEncKey.Item(lIdx) = uEncKey.Item(lIdx - lKeyLen) Xor lPrev
    Next
    pvKeySchedule = lIdx
    '--- inverse
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
                uDecKey.Item(lJdx) = uDecTable.Item(0).Item(uSbox.Item(lPrev And 255) And &HFF&) _
                    Xor uDecTable.Item(1).Item(uSbox.Item((lPrev >> 8) And 255) And &HFF&) _
                    Xor uDecTable.Item(2).Item(uSbox.Item((lPrev >> 16) And 255) And &HFF&) _
                    Xor uDecTable.Item(3).Item(uSbox.Item(lPrev >> 24) And &HFF&)
            #Else
                lTemp = (lPrev And &H7FFFFFFF) \ LNG_POW2_24 Or -(lPrev < 0) * LNG_POW2_7
                uDecKey.Item(lJdx) = uDecTable.Item(0).Item(uSbox.Item(lPrev And &HFF&) And &HFF&) _
                    Xor uDecTable.Item(1).Item(uSbox.Item((lPrev And &HFF00&) \ LNG_POW2_8) And &HFF&) _
                    Xor uDecTable.Item(2).Item(uSbox.Item((lPrev And &HFF0000) \ LNG_POW2_16) And &HFF&) _
                    Xor uDecTable.Item(3).Item(uSbox.Item(lTemp) And &HFF&)
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
            lTemp1 = uT0.Item(lA And 255) Xor uT1.Item((lB >> 8) And 255) Xor uT2.Item((lC >> 16) And 255) Xor uT3.Item(lD >> 24) Xor uKey.Item(lKdx + 0)
            lTemp2 = uT0.Item(lB And 255) Xor uT1.Item((lC >> 8) And 255) Xor uT2.Item((lD >> 16) And 255) Xor uT3.Item(lA >> 24) Xor uKey.Item(lKdx + 1)
            lTemp3 = uT0.Item(lC And 255) Xor uT1.Item((lD >> 8) And 255) Xor uT2.Item((lA >> 16) And 255) Xor uT3.Item(lB >> 24) Xor uKey.Item(lKdx + 2)
            lD = uT0.Item(lD And 255) Xor uT1.Item((lA >> 8) And 255) Xor uT2.Item((lB >> 16) And 255) Xor uT3.Item(lC >> 24) Xor uKey.Item(lKdx + 3)
        #Else
            lTemp1 = uT0.Item(lA And 255) _
                Xor uT1.Item((lB And &HFF00&) \ LNG_POW2_8) _
                Xor uT2.Item((lC And &HFF0000) \ LNG_POW2_16) _
                Xor uT3.Item((lD And &H7F000000) \ LNG_POW2_24 Or -(lD < 0) * LNG_POW2_7) _
                Xor uKey.Item(lKdx + 0)
            lTemp2 = uT0.Item(lB And 255) _
                Xor uT1.Item((lC And &HFF00&) \ LNG_POW2_8) _
                Xor uT2.Item((lD And &HFF0000) \ LNG_POW2_16) _
                Xor uT3.Item((lA And &H7F000000) \ LNG_POW2_24 Or -(lA < 0) * LNG_POW2_7) _
                Xor uKey.Item(lKdx + 1)
            lTemp3 = uT0.Item(lC And 255) _
                Xor uT1.Item((lD And &HFF00&) \ LNG_POW2_8) _
                Xor uT2.Item((lA And &HFF0000) \ LNG_POW2_16) _
                Xor uT3.Item((lB And &H7F000000) \ LNG_POW2_24 Or -(lB < 0) * LNG_POW2_7) _
                Xor uKey.Item(lKdx + 2)
            lD = uT0.Item(lD And 255) _
                Xor uT1.Item((lA And &HFF00&) \ LNG_POW2_8) _
                Xor uT2.Item((lB And &HFF0000) \ LNG_POW2_16) _
                Xor uT3.Item((lC And &H7F000000) \ LNG_POW2_24 Or -(lC < 0) * LNG_POW2_7) _
                Xor uKey.Item(lKdx + 3)
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
            uOutput.Item(lJdx) = (uSbox.Item(lA And 255) And &HFF&) _
                Xor (uSbox.Item((lB >> 8) And 255) And &HFF00&) _
                Xor (uSbox.Item((lC >> 16) And 255) And &HFF0000) _
                Xor (uSbox.Item(lD >> 24) And &HFF000000) Xor uKey.Item(lKdx)
        #Else
            uOutput.Item(lJdx) = (uSbox.Item(lA And 255) And &HFF&) _
                Xor (uSbox.Item((lB And &HFF00&) \ LNG_POW2_8) And &HFF00&) _
                Xor (uSbox.Item((lC And &HFF0000) \ LNG_POW2_16) And &HFF0000) _
                Xor (uSbox.Item((lD And &H7F000000) \ LNG_POW2_24 Or -(lD < 0) * LNG_POW2_7) And &HFF000000) _
                Xor uKey.Item(lKdx)
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
        If IsNumeric(Nonce) Then
            .Nonce.Item(3) = Nonce
        End If
    End With
End Sub

Public Sub CryptoAesProcess(uCtx As CryptoAesContext, ByVal Encrypt As Boolean, baBlock() As Byte, Optional ByVal Pos As Long)
    Debug.Assert UBound(baBlock) + 1 >= Pos + LNG_BLOCKSZ
    pvInitPeek m_uPeekBlock, baBlock, Pos, LNG_BLOCKSZ
    pvProcess uCtx, Encrypt, m_aBlock(0), m_aBlock(0)
End Sub

Public Sub CryptoAesNextNonce(uCtx As CryptoAesContext, baNonce() As Byte)
    With uCtx
        If UBound(baNonce) <> LNG_BLOCKSZ - 1 Then
            ReDim Preserve baNonce(0 To LNG_BLOCKSZ - 1) As Byte
        End If
        Call CopyMemory(.Nonce, baNonce(0), LNG_BLOCKSZ)
        '--- unsigned inc
        .Nonce.Item(3) = BSwap32((BSwap32(.Nonce.Item(3)) Xor &H80000000) + 1 Xor &H80000000)
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
            With uBlock
                m_aBlock(lIdx).Item(0) = .Item(0)
                m_aBlock(lIdx).Item(1) = .Item(1)
                m_aBlock(lIdx).Item(2) = .Item(2)
                m_aBlock(lIdx).Item(3) = .Item(3)
            End With
        End If
        With uCtx.Nonce
            .Item(0) = .Item(0) Xor m_aBlock(lIdx).Item(0)
            .Item(1) = .Item(1) Xor m_aBlock(lIdx).Item(1)
            .Item(2) = .Item(2) Xor m_aBlock(lIdx).Item(2)
            .Item(3) = .Item(3) Xor m_aBlock(lIdx).Item(3)
        End With
        pvProcess uCtx, True, uCtx.Nonce, uCtx.Nonce
        With uCtx.Nonce
            m_aBlock(lIdx).Item(0) = .Item(0)
            m_aBlock(lIdx).Item(1) = .Item(1)
            m_aBlock(lIdx).Item(2) = .Item(2)
            m_aBlock(lIdx).Item(3) = .Item(3)
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
            .Item(0) = m_aBlock(lIdx).Item(0)
            .Item(1) = m_aBlock(lIdx).Item(1)
            .Item(2) = m_aBlock(lIdx).Item(2)
            .Item(3) = m_aBlock(lIdx).Item(3)
        End With
        pvProcess uCtx, False, uInput, uBlock
        With uBlock
            .Item(0) = .Item(0) Xor uCtx.Nonce.Item(0)
            .Item(1) = .Item(1) Xor uCtx.Nonce.Item(1)
            .Item(2) = .Item(2) Xor uCtx.Nonce.Item(2)
            .Item(3) = .Item(3) Xor uCtx.Nonce.Item(3)
        End With
        With uCtx.Nonce
            .Item(0) = uInput.Item(0)
            .Item(1) = uInput.Item(1)
            .Item(2) = uInput.Item(2)
            .Item(3) = uInput.Item(3)
        End With
        With uBlock
            m_aBlock(lIdx).Item(0) = .Item(0)
            m_aBlock(lIdx).Item(1) = .Item(1)
            m_aBlock(lIdx).Item(2) = .Item(2)
            m_aBlock(lIdx).Item(3) = .Item(3)
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
    If Size = 0 Then
        Exit Sub
    End If
    lFinal = Size \ LNG_BLOCKSZ
    pvInitPeek m_uPeekBlock, baBuffer, Pos, Size
    For lIdx = 0 To (Size - 1) \ LNG_BLOCKSZ
        pvProcess uCtx, True, uCtx.Nonce, uBlock
        If lIdx = lFinal Then
            lJdx = lIdx * LNG_BLOCKSZ
            Call CopyMemory(uTemp, baBuffer(Pos + lJdx), Size - lJdx)
            With uTemp
                .Item(0) = .Item(0) Xor uBlock.Item(0)
                .Item(1) = .Item(1) Xor uBlock.Item(1)
                .Item(2) = .Item(2) Xor uBlock.Item(2)
                .Item(3) = .Item(3) Xor uBlock.Item(3)
            End With
            Call CopyMemory(baBuffer(Pos + lJdx), uTemp, Size - lJdx)
        Else
            With uBlock
                m_aBlock(lIdx).Item(0) = m_aBlock(lIdx).Item(0) Xor .Item(0)
                m_aBlock(lIdx).Item(1) = m_aBlock(lIdx).Item(1) Xor .Item(1)
                m_aBlock(lIdx).Item(2) = m_aBlock(lIdx).Item(2) Xor .Item(2)
                m_aBlock(lIdx).Item(3) = m_aBlock(lIdx).Item(3) Xor .Item(3)
            End With
        End If
        '--- unsigned inc
        uCtx.Nonce.Item(3) = BSwap32((BSwap32(uCtx.Nonce.Item(3)) Xor &H80000000) + 1 Xor &H80000000)
    Next
End Sub
