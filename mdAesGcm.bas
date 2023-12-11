Attribute VB_Name = "mdAesGcm"
'--- mdAesGcm.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#End If

Private Const LNG_BLOCKSZ               As Long = 16
Private Const LNG_POW2_1                As Long = 2 ^ 1
Private Const LNG_POW2_3                As Long = 2 ^ 3
Private Const LNG_POW2_4                As Long = 2 ^ 4
Private Const LNG_POW2_27               As Long = 2 ^ 27
Private Const LNG_POW2_28               As Long = 2 ^ 28
Private Const LNG_POW2_30               As Long = 2 ^ 30
Private Const LNG_POW2_31               As Long = &H80000000

Private Type ArrayLong4
    Item(0 To 3)        As Long
End Type

Private Type ArrayByte16
    Item(0 To 15)       As Byte
End Type

Private Type ShoupTable
    Item(0 To 15)       As ArrayLong4
End Type

Public Type CryptoGhashContext
    KeyTable            As ShoupTable
    NonceArray          As ArrayByte16
    HashArray           As ArrayByte16
    NPosition           As Long
End Type

Public Type CryptoAesGcmContext
    AesCtx              As CryptoAesContext
    GhashCtx            As CryptoGhashContext
    AadSize             As Currency
    TotalSize           As Currency
End Type

Private m_aReverse(0 To 15)         As Long
Private m_aReduce(0 To 15)          As Long

Private Function BSwap32(ByVal lX As Long) As Long
    BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
                 (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
End Function

Private Sub pvInit()
    Const LNG_POLY1 As Long = &HE1000000
    Const LNG_POLY2 As Long = LNG_POLY1 \ 2 And &H7FFFFFFF
    Const LNG_POLY4 As Long = LNG_POLY2 \ 2
    Const LNG_POLY8 As Long = LNG_POLY4 \ 2
    Dim lIdx            As Long
    
    For lIdx = 0 To 15
        m_aReverse(lIdx) = -((lIdx And 1) <> 0) * 8 Xor -((lIdx And 2) <> 0) * 4 _
                       Xor -((lIdx And 4) <> 0) * 2 Xor -((lIdx And 8) <> 0) * 1
        m_aReduce(lIdx) = -((lIdx And 1) <> 0) * LNG_POLY8 Xor -((lIdx And 2) <> 0) * LNG_POLY4 _
                      Xor -((lIdx And 4) <> 0) * LNG_POLY2 Xor -((lIdx And 8) <> 0) * LNG_POLY1
    Next
End Sub

Private Sub pvPrecompute(baKey() As Byte, uKeyTable As ShoupTable)
    Dim lIdx            As Long
    Dim uOne            As ArrayLong4
    Dim uTemp           As ArrayLong4
    Dim lCarry          As Long
    
    lIdx = UBound(baKey) + 1
    If lIdx > LNG_BLOCKSZ Then
        lIdx = LNG_BLOCKSZ
    End If
    Call CopyMemory(uTemp.Item(0), baKey(0), lIdx)
    With uOne
        .Item(0) = BSwap32(uTemp.Item(3))
        .Item(1) = BSwap32(uTemp.Item(2))
        .Item(2) = BSwap32(uTemp.Item(1))
        .Item(3) = BSwap32(uTemp.Item(0))
    End With
    '--- precompute all multiples of H needed for Shoup's method
    With uKeyTable
        '--- M(1) = H * 1
        lIdx = 1
        .Item(m_aReverse(lIdx)) = uOne
        For lIdx = 2 To UBound(.Item)
            If (lIdx And 1) <> 0 Then
                '--- M(i) = M(i - 1) + M(1)
                uTemp = .Item(m_aReverse(lIdx - 1))
                With uTemp
                    .Item(0) = .Item(0) Xor uOne.Item(0)
                    .Item(1) = .Item(1) Xor uOne.Item(1)
                    .Item(2) = .Item(2) Xor uOne.Item(2)
                    .Item(3) = .Item(3) Xor uOne.Item(3)
                End With
            Else
                '--- M(i) = M(i / 2) * x
                uTemp = .Item(m_aReverse(lIdx \ 2))
                With uTemp
                    lCarry = .Item(0) And 1
                    .Item(0) = (.Item(0) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(0) < 0) * LNG_POW2_30 Or (.Item(1) And 1) * LNG_POW2_31
                    .Item(1) = (.Item(1) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(1) < 0) * LNG_POW2_30 Or (.Item(2) And 1) * LNG_POW2_31
                    .Item(2) = (.Item(2) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(2) < 0) * LNG_POW2_30 Or (.Item(3) And 1) * LNG_POW2_31
                    .Item(3) = (.Item(3) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(3) < 0) * LNG_POW2_30 Xor lCarry * m_aReduce(m_aReverse(1))
                End With
            End If
            .Item(m_aReverse(lIdx)) = uTemp
        Next
    End With
End Sub

Private Sub pvMult(uKeyTable As ShoupTable, uArray As ArrayByte16)
    Dim uBlock          As ArrayLong4
    Dim lIdx            As Long
    Dim lNibble         As Long
    Dim lCarry          As Long
    Dim uResult         As ArrayLong4
    
    With uBlock
        lNibble = uArray.Item(LNG_BLOCKSZ - 1) And &HF
        .Item(0) = uKeyTable.Item(lNibble).Item(0)
        .Item(1) = uKeyTable.Item(lNibble).Item(1)
        .Item(2) = uKeyTable.Item(lNibble).Item(2)
        .Item(3) = uKeyTable.Item(lNibble).Item(3)
        For lIdx = LNG_BLOCKSZ - 1 To 0 Step -1
            If lIdx <> LNG_BLOCKSZ - 1 Then
                '--- mul 16
                lCarry = .Item(0) And &HF
                .Item(0) = (.Item(0) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(0) < 0) * LNG_POW2_27 _
                    Or (.Item(1) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(1) And LNG_POW2_3) <> 0) * LNG_POW2_31
                .Item(1) = (.Item(1) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(1) < 0) * LNG_POW2_27 _
                    Or (.Item(2) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(2) And LNG_POW2_3) <> 0) * LNG_POW2_31
                .Item(2) = (.Item(2) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(2) < 0) * LNG_POW2_27 _
                    Or (.Item(3) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(3) And LNG_POW2_3) <> 0) * LNG_POW2_31
                .Item(3) = (.Item(3) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(3) < 0) * LNG_POW2_27 _
                    Xor m_aReduce(lCarry)
                '--- add lower nibble
                lNibble = uArray.Item(lIdx) And &HF
                .Item(0) = .Item(0) Xor uKeyTable.Item(lNibble).Item(0)
                .Item(1) = .Item(1) Xor uKeyTable.Item(lNibble).Item(1)
                .Item(2) = .Item(2) Xor uKeyTable.Item(lNibble).Item(2)
                .Item(3) = .Item(3) Xor uKeyTable.Item(lNibble).Item(3)
            End If
            '--- mul 16
            lCarry = .Item(0) And &HF
            .Item(0) = (.Item(0) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(0) < 0) * LNG_POW2_27 _
                Or (.Item(1) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(1) And LNG_POW2_3) <> 0) * LNG_POW2_31
            .Item(1) = (.Item(1) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(1) < 0) * LNG_POW2_27 _
                Or (.Item(2) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(2) And LNG_POW2_3) <> 0) * LNG_POW2_31
            .Item(2) = (.Item(2) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(2) < 0) * LNG_POW2_27 _
                Or (.Item(3) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(3) And LNG_POW2_3) <> 0) * LNG_POW2_31
            .Item(3) = (.Item(3) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(3) < 0) * LNG_POW2_27 _
                Xor m_aReduce(lCarry)
            '--- add upper nibble
            lNibble = (uArray.Item(lIdx) \ LNG_POW2_4) And &HF
            .Item(0) = .Item(0) Xor uKeyTable.Item(lNibble).Item(0)
            .Item(1) = .Item(1) Xor uKeyTable.Item(lNibble).Item(1)
            .Item(2) = .Item(2) Xor uKeyTable.Item(lNibble).Item(2)
            .Item(3) = .Item(3) Xor uKeyTable.Item(lNibble).Item(3)
        Next
    End With
    With uResult
        .Item(0) = BSwap32(uBlock.Item(3))
        .Item(1) = BSwap32(uBlock.Item(2))
        .Item(2) = BSwap32(uBlock.Item(1))
        .Item(3) = BSwap32(uBlock.Item(0))
    End With
    LSet uArray = uResult
End Sub

Private Function pvUpdate(uKeyTable As ShoupTable, uArray As ArrayByte16, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long, Optional ByVal Offset As Long) As Long
    Dim lIdx            As Long
    
    With uArray
        For lIdx = 0 To lSize - 1
            .Item(Offset) = .Item(Offset) Xor baInput(lPos + lIdx)
            Offset = Offset + 1
            If Offset = LNG_BLOCKSZ Then
                Offset = 0
                pvMult uKeyTable, uArray
            End If
        Next
    End With
    pvUpdate = Offset
End Function

Public Sub CryptoGhashInit(uCtx As CryptoGhashContext, baKey() As Byte, baNonce() As Byte)
    Dim uArray          As ArrayByte16
    Dim lSize           As Long
    
    If m_aReduce(1) = 0 Then
        pvInit
    End If
    With uCtx
        pvPrecompute baKey, .KeyTable
        .NonceArray = uArray
        .HashArray = uArray
        .NPosition = 0
        lSize = UBound(baNonce) + 1
        If lSize = 12 Then '--- 96 bits
            Call CopyMemory(.NonceArray.Item(0), baNonce(0), lSize)
            .NonceArray.Item(LNG_BLOCKSZ - 1) = 1
        Else
            pvUpdate .KeyTable, .NonceArray, baNonce, 0, lSize
            If lSize Mod LNG_BLOCKSZ <> 0 Then
                pvUpdate .KeyTable, .NonceArray, uArray.Item, 0, LNG_BLOCKSZ - lSize Mod LNG_BLOCKSZ, lSize Mod LNG_BLOCKSZ
            End If
            lSize = BSwap32(lSize * 8)
            Call CopyMemory(uArray.Item(12), lSize, LenB(lSize))
            pvUpdate .KeyTable, .NonceArray, uArray.Item, 0, LNG_BLOCKSZ
        End If
    End With
End Sub

Public Sub CryptoGhashUpdate(uCtx As CryptoGhashContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    uCtx.NPosition = pvUpdate(uCtx.KeyTable, uCtx.HashArray, baInput, Pos, Size, Offset:=uCtx.NPosition)
End Sub

Public Sub CryptoGhashPad(uCtx As CryptoGhashContext)
    If uCtx.NPosition > 0 Then
        pvMult uCtx.KeyTable, uCtx.HashArray
        uCtx.NPosition = 0
    End If
End Sub

Public Sub CryptoGhashFinalize(uCtx As CryptoGhashContext, ByVal lTagSize As Long, baTag() As Byte)
    If lTagSize < 4 Or lTagSize > LNG_BLOCKSZ Then
        Err.Raise vbObjectError, , "Invalid tag size for Ghash (" & lTagSize & ")"
    End If
    With uCtx
        ReDim baTag(0 To lTagSize - 1) As Byte
        Call CopyMemory(baTag(0), .HashArray, lTagSize)
    End With
End Sub

'= AES-GCM ===============================================================

Public Sub CryptoAesGcmInit(uCtx As CryptoAesGcmContext, baKey() As Byte, baNonce() As Byte, baAad() As Byte)
    Dim baHashKey(0 To LNG_BLOCKSZ - 1) As Byte
    
    With uCtx
        CryptoAesInit uCtx.AesCtx, baKey
        '--- encrypt a block of zeroes to create the hashing key
        CryptoAesProcess .AesCtx, True, baHashKey
        CryptoGhashInit .GhashCtx, baHashKey, baNonce
        '--- setup AES counter
        CryptoAesNextNonce .AesCtx, .GhashCtx.NonceArray.Item
        CryptoAesProcess .AesCtx, True, .GhashCtx.NonceArray.Item
        '--- absorb AAD into the hash
        CryptoGhashUpdate .GhashCtx, baAad
        CryptoGhashPad .GhashCtx
        .AadSize = UBound(baAad) + 1
        .TotalSize = 0
    End With
End Sub

Private Function pvAesGcmFinalize(uCtx As CryptoAesGcmContext, ByVal lTagSize As Long, baTag() As Byte)
    Dim cTemp           As Currency
    Dim aTemp(0 To 1)   As Long
    Dim uBlock          As ArrayLong4
    Dim uArray          As ArrayByte16
    Dim lIdx            As Long
    
    With uCtx
        CryptoGhashPad .GhashCtx
        '--- absorb bit-size of AAD and plaintext
        cTemp = .AadSize * 8 / 10000@
        Call CopyMemory(aTemp(0), cTemp, 8)
        uBlock.Item(0) = BSwap32(aTemp(1))
        uBlock.Item(1) = BSwap32(aTemp(0))
        cTemp = .TotalSize * 8 / 10000@
        Call CopyMemory(aTemp(0), cTemp, 8)
        uBlock.Item(2) = BSwap32(aTemp(1))
        uBlock.Item(3) = BSwap32(aTemp(0))
        LSet uArray = uBlock
        CryptoGhashUpdate .GhashCtx, uArray.Item
        '--- finalize hash
        CryptoGhashFinalize .GhashCtx, lTagSize, baTag
        For lIdx = 0 To lTagSize - 1
            baTag(lIdx) = baTag(lIdx) Xor .GhashCtx.NonceArray.Item(lIdx)
        Next
    End With
End Function

Public Sub CryptoAesGcmEncrypt(uCtx As CryptoAesGcmContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional TagSize As Long, Optional Tag As Variant)
    Dim baTag()         As Byte
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    With uCtx
        CryptoAesCtrCrypt .AesCtx, baBuffer, Pos, Size
        CryptoGhashUpdate .GhashCtx, baBuffer, Pos, Size
        .TotalSize = .TotalSize + Size
        If TagSize > 0 Then
            pvAesGcmFinalize uCtx, TagSize, baTag
            Tag = baTag
        End If
    End With
End Sub

Public Function CryptoAesGcmDecrypt(uCtx As CryptoAesGcmContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional Tag As Variant) As Boolean
    Dim baTag()         As Byte
    Dim baCalc()        As Byte
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    With uCtx
        CryptoGhashUpdate .GhashCtx, baBuffer, Pos, Size
        .TotalSize = .TotalSize + Size
        If Not IsMissing(Tag) Then
            baTag = Tag
            pvAesGcmFinalize uCtx, UBound(baTag) + 1, baCalc
            If InStrB(baTag, baCalc) <> 1 Then
                Exit Function
            End If
        End If
        CryptoAesCtrCrypt .AesCtx, baBuffer, Pos, Size
    End With
    '--- success
    CryptoAesGcmDecrypt = True
End Function
