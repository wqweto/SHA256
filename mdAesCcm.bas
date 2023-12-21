Attribute VB_Name = "mdAesCcm"
'--- mdAesCcm.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)

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
Private Const LNG_POW2_2                As Long = 2 ^ 2
Private Const LNG_POW2_6                As Long = 2 ^ 6

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    cElements           As Long
    lLbound             As Long
End Type

Private Type AesBlock
    Item(0 To 3)        As Long
End Type

Private Function BSwap32(ByVal lX As Long) As Long
    BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
              (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
End Function

Private Function BSwap16(ByVal lX As Long) As Long
    BSwap16 = (lX And &HFF) * &H100 Or (lX And &HFF00&) \ &H100
End Function

Private Function pvComputeLengthSize(baNonce() As Byte, baBuffer() As Byte) As Long
    pvComputeLengthSize = 2
    Do While 2 ^ (8 * pvComputeLengthSize) < UBound(baBuffer) + 1
        pvComputeLengthSize = pvComputeLengthSize + 1
    Loop
    If pvComputeLengthSize < (15 - UBound(baNonce) - 1) Then
        pvComputeLengthSize = (15 - UBound(baNonce) - 1)
    End If
End Function

Private Sub pvMac(uCtx As CryptoAesContext, baInput() As Byte, ByVal lPos As Long, baTag() As Byte)
    Const FADF_AUTO     As Long = 1
    Dim aBlock()        As AesBlock
    Dim uPeekBlock      As SAFEARRAY1D
    Dim uTag            As AesBlock
    Dim pDummy          As LongPtr
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    If UBound(baInput) >= 0 Then
        If lPos + 4 * LNG_BLOCKSZ <= UBound(baInput) + 1 Then
            With uPeekBlock
                .cDims = 1
                .fFeatures = FADF_AUTO
                .cbElements = LenB(uTag)
                .cLocks = 1
                .pvData = VarPtr(baInput(0))
                .cElements = (UBound(baInput) + 1) \ .cbElements
            End With
            Call CopyMemory(ByVal ArrPtr(aBlock), VarPtr(uPeekBlock), LenB(pDummy))
            Call CopyMemory(uTag.Item(0), baTag(0), LNG_BLOCKSZ)
            Do While lPos + LNG_BLOCKSZ <= UBound(baInput) + 1
                uTag.Item(0) = uTag.Item(0) Xor aBlock(lJdx).Item(0)
                uTag.Item(1) = uTag.Item(1) Xor aBlock(lJdx).Item(1)
                uTag.Item(2) = uTag.Item(2) Xor aBlock(lJdx).Item(2)
                uTag.Item(3) = uTag.Item(3) Xor aBlock(lJdx).Item(3)
                CryptoAesProcessPtr uCtx, VarPtr(uTag.Item(0))
                lPos = lPos + LNG_BLOCKSZ
                lJdx = lJdx + 1
            Loop
            Call CopyMemory(baTag(0), uTag.Item(0), LNG_BLOCKSZ)
        End If
        Do While lPos <= UBound(baInput)
            For lIdx = 0 To LNG_BLOCKSZ - 1
                If lPos <= UBound(baInput) Then
                    baTag(lIdx) = baTag(lIdx) Xor baInput(lPos)
                    lPos = lPos + 1
                End If
            Next
            CryptoAesProcess uCtx, baTag
        Loop
    End If
End Sub

Private Sub pvComputeTag(uCtx As CryptoAesContext, baNonce() As Byte, baAad() As Byte, baInput() As Byte, ByVal lLengthSize As Long, baTag() As Byte, ByVal lTagSize As Long)
    Const MAX_SHORT_SIZE As Long = &HFEFF&
    Dim baBlock(0 To LNG_BLOCKSZ - 1) As Byte
    Dim lIdx            As Long
    Dim lSize           As Long
    
    If lTagSize < 4 Or lTagSize > 16 Or (lTagSize And 1) = 1 Then
        Err.Raise vbObjectError, , "Invalid tag size for AES-CCM (" & lTagSize & ")"
    End If
    If UBound(baNonce) + 1 < 7 Or UBound(baNonce) + 1 > 15 Then
        Err.Raise vbObjectError, , "Invalid nonce size for AES-CCM (" & UBound(baNonce) + 1 & ")"
    End If
    ReDim baTag(0 To LNG_BLOCKSZ - 1) As Byte
    '--- [0] = flags
    baTag(0) = -(UBound(baAad) >= 0) * LNG_POW2_6 Or (lTagSize - 2) * LNG_POW2_2 Or (lLengthSize - 1)
    '--- [1 to 15-L] = nonce
    lIdx = UBound(baNonce) + 1
    If lIdx > 15 - lLengthSize Then
        lIdx = 15 - lLengthSize
    End If
    Call CopyMemory(baTag(1), baNonce(0), lIdx)
    '--- [15-L to 15] = big-endian plaintext length
    Call CopyMemory(lIdx, baTag(12), LenB(lIdx))
    lIdx = lIdx Xor BSwap32(UBound(baInput) + 1)
    Call CopyMemory(baTag(12), lIdx, LenB(lIdx))
    CryptoAesProcess uCtx, baTag
    '--- mac the AAD
    lSize = UBound(baAad) + 1
    If lSize > 0 Then
        If lSize < MAX_SHORT_SIZE Then
            lSize = BSwap16(lSize)
            Call CopyMemory(baBlock(0), lSize, 2)
            lIdx = 2
        Else
            lSize = BSwap32(lSize)
            Call CopyMemory(baBlock(2), lSize, 4)
            lSize = BSwap16(MAX_SHORT_SIZE)
            Call CopyMemory(baBlock(0), lSize, 2)
            lIdx = 6
        End If
        lSize = UBound(baAad) + 1
        If lSize > LNG_BLOCKSZ - lIdx Then
            lSize = LNG_BLOCKSZ - lIdx
        End If
        Call CopyMemory(baBlock(lIdx), baAad(0), lSize)
        For lIdx = 0 To LNG_BLOCKSZ - 1
            baTag(lIdx) = baTag(lIdx) Xor baBlock(lIdx)
        Next
        CryptoAesProcess uCtx, baTag
        pvMac uCtx, baAad, lSize, baTag
    End If
    '--- mac the plaintext
    pvMac uCtx, baInput, 0, baTag
End Sub

Public Sub CryptoAesCcmEncrypt(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baBuffer() As Byte, baTag() As Byte, Optional ByVal TagSize As Long = LNG_BLOCKSZ)
    Dim uCtx            As CryptoAesContext
    Dim lLengthSize     As Long
    Dim baBlock(0 To LNG_BLOCKSZ - 1) As Byte
    Dim lTemp           As Long
    
    CryptoAesInit uCtx, baKey
    lLengthSize = pvComputeLengthSize(baNonce, baBuffer)
    pvComputeTag uCtx, baNonce, baAad, baBuffer, lLengthSize, baTag, TagSize
    '--- [0] = flags
    baBlock(0) = 0 * LNG_POW2_6 Or 0 * LNG_POW2_2 Or (lLengthSize - 1)
    '--- [1 to 15-L] = nonce
    lTemp = UBound(baNonce) + 1
    If lTemp > 15 - lLengthSize Then
        lTemp = 15 - lLengthSize
    End If
    Call CopyMemory(baBlock(1), baNonce(0), lTemp)
    CryptoAesSetNonce uCtx, baBlock
    lTemp = (lLengthSize + 3) \ 4
    CryptoAesCtrCrypt uCtx, baTag, CounterWords:=lTemp
    CryptoAesCtrCrypt uCtx, baBuffer, CounterWords:=lTemp
    ReDim Preserve baTag(0 To TagSize - 1) As Byte
End Sub

Public Function CryptoAesCcmDecrypt(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baBuffer() As Byte, baTag() As Byte) As Boolean
    Dim uCtx            As CryptoAesContext
    Dim lLengthSize     As Long
    Dim baBlock(0 To LNG_BLOCKSZ - 1) As Byte
    Dim lTemp           As Long
    Dim lTagSize        As Long
    Dim baTemp()        As Byte
    
    CryptoAesInit uCtx, baKey
    lLengthSize = pvComputeLengthSize(baNonce, baBuffer)
    '--- [0] = flags
    baBlock(0) = 0 * LNG_POW2_6 Or 0 * LNG_POW2_2 Or (lLengthSize - 1)
    '--- [1 to 15-L] = nonce
    lTemp = UBound(baNonce) + 1
    If lTemp > 15 - lLengthSize Then
        lTemp = 15 - lLengthSize
    End If
    Call CopyMemory(baBlock(1), baNonce(0), lTemp)
    CryptoAesSetNonce uCtx, baBlock
    lTagSize = UBound(baTag) + 1
    If lTagSize <> LNG_BLOCKSZ Then
        ReDim Preserve baTag(0 To LNG_BLOCKSZ - 1) As Byte
    End If
    lTemp = (lLengthSize + 3) \ 4
    CryptoAesCtrCrypt uCtx, baTag, CounterWords:=lTemp
    CryptoAesCtrCrypt uCtx, baBuffer, CounterWords:=lTemp
    pvComputeTag uCtx, baNonce, baAad, baBuffer, lLengthSize, baTemp, lTagSize
    If lTagSize <> LNG_BLOCKSZ Then
        ReDim Preserve baTag(0 To lTagSize - 1) As Byte
    End If
    If InStrB(baTemp, baTag) <> 1 Then
        GoTo QH
    End If
    '--- success
    CryptoAesCcmDecrypt = True
QH:
End Function
