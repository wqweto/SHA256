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
Private Const LNG_POW2_3                As Long = 2 ^ 3
Private Const LNG_POW2_28               As Long = 2 ^ 28

Private Type ArrayLong4
    Item(0 To 3)        As Long
End Type

Private Type ArrayByte16
    Item(0 To 15)       As Byte
End Type

Private Type GmacTable
    Item(0 To 15)       As ArrayLong4
End Type

Private Type CryptoGmacContext
    PreComp             As GmacTable
    StateArray          As ArrayByte16
    NPosition           As Long
    MacArray            As ArrayByte16
End Type

Public Type CryptoAesGcmContext
    AesCtx              As CryptoAesContext
    GmacCtx             As CryptoGmacContext
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
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    lIdx = 1
    For Each vElem In Split("8 4 12 2 10 6 14 1 9 5 13 3 11 7 15")
        m_aReverse(lIdx) = vElem
        lIdx = lIdx + 1
    Next
    lIdx = 1
    For Each vElem In Split("1C20 3840 2460 7080 6CA0 48C0 54E0 E100 FD20 D940 C560 9180 8DA0 A9C0 B5E0")
        m_aReduce(lIdx) = CLng("&H" & vElem & "0000")
        lIdx = lIdx + 1
    Next
End Sub

Private Sub pvMult(uArray As ArrayByte16, uPreComp As GmacTable)
    Dim uBlock          As ArrayLong4
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lCarry          As Long
    
    With uBlock
        For lIdx = LNG_BLOCKSZ - 1 To 0 Step -1
            lJdx = uArray.Item(lIdx) And &HF
            lCarry = .Item(0) And &HF
            .Item(0) = .Item(0) \ 16 Or (.Item(1) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(1) And LNG_POW2_3) <> 0) * &H80000000
            .Item(1) = .Item(1) \ 16 Or (.Item(2) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(2) And LNG_POW2_3) <> 0) * &H80000000
            .Item(2) = .Item(2) \ 16 Or (.Item(3) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(3) And LNG_POW2_3) <> 0) * &H80000000
            .Item(3) = .Item(3) \ 16
            .Item(0) = .Item(0) Xor uPreComp.Item(lJdx).Item(0)
            .Item(1) = .Item(1) Xor uPreComp.Item(lJdx).Item(1)
            .Item(2) = .Item(2) Xor uPreComp.Item(lJdx).Item(2)
            .Item(3) = .Item(3) Xor uPreComp.Item(lJdx).Item(3)
            .Item(3) = .Item(3) Xor m_aReduce(lCarry)
            
            lJdx = (uArray.Item(lIdx) \ 16) And &HF
            lCarry = (.Item(0) \ 16) And &HF
            .Item(0) = .Item(0) \ 16 Or (.Item(1) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(1) And LNG_POW2_3) <> 0) * &H80000000
            .Item(1) = .Item(1) \ 16 Or (.Item(2) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(2) And LNG_POW2_3) <> 0) * &H80000000
            .Item(2) = .Item(2) \ 16 Or (.Item(3) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(3) And LNG_POW2_3) <> 0) * &H80000000
            .Item(3) = .Item(3) \ 16
            .Item(0) = .Item(0) Xor uPreComp.Item(lJdx).Item(0)
            .Item(1) = .Item(1) Xor uPreComp.Item(lJdx).Item(1)
            .Item(2) = .Item(2) Xor uPreComp.Item(lJdx).Item(2)
            .Item(3) = .Item(3) Xor uPreComp.Item(lJdx).Item(3)
            .Item(3) = .Item(3) Xor m_aReduce(lCarry)
        Next
    End With
    With uBlock
        .Item(0) = BSwap32(.Item(0))
        .Item(1) = BSwap32(.Item(1))
        .Item(2) = BSwap32(.Item(2))
        .Item(3) = BSwap32(.Item(3))
    End With
    LSet uArray = uBlock
End Sub

Private Sub pvGmacInit(uCtx As CryptoGmacContext, uH As ArrayByte16)
    Dim uOne            As ArrayLong4
    Dim uTemp           As ArrayLong4
    Dim lIdx            As Long
    Dim lCarry          As Long
    
    '--- init multiplication LUT
    With uCtx.PreComp
        LSet uOne = uH
        '--- M(1) = H * 1
        lIdx = 1
        .Item(m_aReverse(lIdx)) = uOne
        '--- multiples of H for Shoup's method
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
                '--- M(i) = M(i / 2) + M(i / 2)
                uTemp = .Item(m_aReverse(lIdx \ 2))
                With uTemp
                    lCarry = .Item(0) And 1
                    .Item(0) = .Item(0) \ 2 Or -(.Item(1) < 0)
                    .Item(1) = .Item(1) \ 2 Or -(.Item(2) < 0)
                    .Item(2) = .Item(2) \ 2 Or -(.Item(3) < 0)
                    .Item(3) = .Item(3) \ 2 Xor lCarry * m_aReduce(m_aReverse(1))
                End With
            End If
            .Item(m_aReverse(lIdx)) = uTemp
        Next
    End With
End Sub

Private Sub pvGmacReset(uCtx As CryptoGmacContext, baIV() As Byte)
    Dim lPos            As Long
    Dim lSize           As Long
    Dim lJdx            As Long
    Dim uArray          As ArrayByte16
    Dim uSize           As ArrayLong4
    
    If UBound(baIV) + 1 = 12 Then '--- 96 bits
        Call CopyMemory(uArray.Item(0), baIV(0), 12)
        uArray.Item(LNG_BLOCKSZ - 1) = 1
    Else
        Do While lPos < UBound(baIV) + 1
            lSize = UBound(baIV) + 1 - lPos
            If lSize > LNG_BLOCKSZ Then
                lSize = LNG_BLOCKSZ
            End If
            For lJdx = 0 To lSize - 1
                With uArray
                    .Item(lJdx) = .Item(lJdx) Xor baIV(lPos + lJdx)
                End With
            Next
            pvMult uArray, uCtx.PreComp
            lPos = lPos + lSize
        Loop
        uSize.Item(3) = BSwap32((UBound(baIV) + 1) * 8)
        LSet uArray = uSize
    End If
    pvMult uArray, uCtx.PreComp
End Sub

Private Sub pvGmacUpdate(uCtx As CryptoGmacContext, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    lJdx = uCtx.NPosition
    With uCtx.StateArray
        For lIdx = 0 To lSize - 1
            .Item(lJdx) = .Item(lJdx) Xor baInput(lPos + lIdx)
            lJdx = lJdx + 1
            If lJdx = LNG_BLOCKSZ Then
                lJdx = 0
                pvMult uCtx.StateArray, uCtx.PreComp
            End If
        Next
    End With
    uCtx.NPosition = lJdx
End Sub

Private Sub pvGmacFinal(uCtx As CryptoGmacContext, cAadSize As Currency, cTotalSize As Currency, baOutput() As Byte, lOutSize As Long)
    Dim lIdx            As Long
    Dim uBlock          As ArrayLong4
    Dim aTemp(0 To 1)   As Long
    Dim uArray          As ArrayByte16
    
    If lOutSize < 4 Or lOutSize > LNG_BLOCKSZ Then
        Err.Raise vbObjectError, , "Invalid output size for GMAC (" & lOutSize & ")"
    End If
    If uCtx.NPosition > 0 Then
        pvMult uCtx.StateArray, uCtx.PreComp
    End If
    '--- append bit-sizes of AAD and ciphertext
    Call CopyMemory(aTemp(0), cAadSize, 8)
    uBlock.Item(0) = BSwap32(aTemp(1))
    uBlock.Item(1) = BSwap32(aTemp(0))
    Call CopyMemory(aTemp(0), cTotalSize, 8)
    uBlock.Item(2) = BSwap32(aTemp(1))
    uBlock.Item(3) = BSwap32(aTemp(0))
    LSet uArray = uBlock
    With uCtx.StateArray
        For lIdx = 0 To LNG_BLOCKSZ - 1
            .Item(lIdx) = .Item(lIdx) Xor uArray.Item(lIdx)
        Next
    End With
    pvMult uCtx.StateArray, uCtx.PreComp
    For lIdx = 0 To LNG_BLOCKSZ - 1
        With uCtx.MacArray
            .Item(lIdx) = .Item(lIdx) Xor uCtx.StateArray.Item(lIdx)
        End With
    Next
    ReDim baOutput(0 To lOutSize - 1) As Byte
    Call CopyMemory(baOutput(0), uCtx.MacArray, lOutSize)
End Sub

Private Function pvArrayEqual(baFirst() As Byte, baSecond() As Byte) As Boolean
    If UBound(baFirst) = UBound(baSecond) Then
        pvArrayEqual = (InStrB(baFirst, baSecond) = 1)
    End If
End Function

Private Function pvUnsignedInc(lValue As Long) As Boolean
    If lValue <> -1 Then
        lValue = (lValue Xor &H80000000) + 1 Xor &H80000000
    Else
        lValue = 0
        '--- signal carry
        pvUnsignedInc = True
    End If
End Function

Public Sub CryptoAesGcmInit(uCtx As CryptoAesGcmContext, baKey() As Byte, baIV() As Byte, baAad() As Byte)
    Dim uH              As ArrayByte16
    Dim lPos            As Long
    Dim lSize           As Long
    Dim lJdx            As Long
    Dim uArray          As ArrayByte16
    
    If m_aReduce(1) = 0 Then
        pvInit
    End If
    With uCtx
        .GmacCtx.StateArray = uH
        CryptoAesInit .AesCtx, baKey
        CryptoAesProcess .AesCtx, True, uH.Item
        pvGmacInit .GmacCtx, uH
        '--- note: use a public function to set Nonce instead of this hack
        LSet .AesCtx.Nonce = uH
        With .AesCtx.Nonce
            .Item(0) = BSwap32(.Item(0))
            .Item(1) = BSwap32(.Item(1))
            .Item(2) = BSwap32(.Item(2))
            .Item(3) = BSwap32(.Item(3))
            For lJdx = 3 To 0 Step -1
                If Not pvUnsignedInc(.Item(lJdx)) Then
                    Exit For
                End If
            Next
        End With
        CryptoAesProcess .AesCtx, True, uH.Item
        .GmacCtx.MacArray = uH
        pvGmacReset .GmacCtx, baIV
        '--- process AAD
        .AadSize = UBound(baAad) + 1
        Do While lPos < UBound(baAad) + 1
            lSize = UBound(baIV) + 1 - lPos
            If lSize > LNG_BLOCKSZ Then
                lSize = LNG_BLOCKSZ
            End If
            For lJdx = 0 To lSize - 1
                With uArray
                    .Item(lJdx) = .Item(lJdx) Xor baAad(lPos + lJdx)
                End With
            Next
            pvMult uArray, .GmacCtx.PreComp
            lPos = lPos + lSize
        Loop
    End With
End Sub

Public Sub CryptoAesGcmEncrypt(uCtx As CryptoAesGcmContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional TagSize As Long, Optional Tag As Variant)
    Dim baOutTag()      As Byte
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    With uCtx
        CryptoAesCtrCrypt .AesCtx, baBuffer, Pos, Size
        pvGmacUpdate .GmacCtx, baBuffer, Pos, Size
        .TotalSize = .TotalSize + Size
        If TagSize > 0 Then
            pvGmacFinal .GmacCtx, .AadSize * 8 / 10000@, .TotalSize * 8 / 10000@, baOutTag, TagSize
            Tag = baOutTag
        End If
    End With
End Sub

Public Function CryptoAesGcmDecrypt(uCtx As CryptoAesGcmContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional Tag As Variant) As Boolean
    Dim baInTag()       As Byte
    Dim baCalcTag()     As Byte
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    With uCtx
        pvGmacUpdate .GmacCtx, baBuffer, Pos, Size
        .TotalSize = .TotalSize + Size
        If Not IsMissing(Tag) Then
            baInTag = Tag
            pvGmacFinal .GmacCtx, .AadSize * 8 / 10000@, .TotalSize * 8 / 10000@, baCalcTag, UBound(baInTag) + 1
            If Not pvArrayEqual(baInTag, baCalcTag) Then
                Exit Function
            End If
        End If
        CryptoAesCtrCrypt .AesCtx, baBuffer, Pos, Size
    End With
    '--- success
    CryptoAesGcmDecrypt = True
End Function
