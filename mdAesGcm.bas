Attribute VB_Name = "mdAesGcm"
'--- mdAesGcm.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As LongPtr, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As LongPtr, pcbBinary As Long, Optional ByVal pdwSkip As LongPtr, Optional ByVal pdwFlags As LongPtr) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As LongPtr, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As LongPtr, pcbBinary As Long, Optional ByVal pdwSkip As LongPtr, Optional ByVal pdwFlags As LongPtr) As Long
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
    HashArray           As ArrayByte16
    NPosition           As Long
End Type

Public Type CryptoAesGcmContext
    AesCtx              As CryptoAesContext
    GhashCtx            As CryptoGhashContext
    Counter(0 To LNG_BLOCKSZ - 1) As Byte
    AadSize             As Currency
    TotalSize           As Currency
End Type

Private m_aReverse(0 To 15)         As Long
Private m_aReduce(0 To 15)          As Long
Private m_hMulThunk                 As LongPtr

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

Private Sub pvInit()
    Const LNG_POLY1 As Long = &HE1000000 '--- GHASH irreducible polynomial
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
    m_hMulThunk = pvThunkAllocate
End Sub

#If WIN64 = 1 Then
Private Function pvThunkAllocate() As LongPtr

End Function

Private Function pvPatchTrampoline(ByVal Pfn As LongPtr, Optional ByVal Noop As Boolean) As Boolean
    #If Pfn And Noop Then '--- touch
    #End If
    pvPatchTrampoline = True
End Function
#Else
Private Function pvThunkAllocate() As LongPtr
    Const MEM_COMMIT                    As Long = &H1000
    Const MEM_DECOMMIT                  As Long = &H4000
    Const PAGE_EXECUTE_READWRITE        As Long = &H40
    Const CRYPT_STRING_BASE64           As Long = 1
    Const THUNK_SIZE    As Long = 327
    Dim STR_THUNK       As String: STR_THUNK = "VYvsi0UIg+wQU1ZXhcB1IkCNffAzyVMPoovzW4kHiXcEiU8IiVcMi0X4g+AC6QwBAAAPECiLRQzHRfAPDg0Mx0X0CwoJCMdF+AcGBQQPEAjHRfwDAgEADxB18GYPOADOZg84AO4PKNUPKMVmDzpEwQFmDzpE0RBmD+/QDyjdZg86RNkAZg86ROkRDyjCZg9z2ghmD3P4CGYP7+pmD+/YDyjlDyjDZg9y1B9mD3LQH2YPcvMBDyjIZg9z/ARmD3P5BGYP68tmD3PYDA8o2WYPcvUBZg9y8x9mD+vlZg/r4A8owWYPcvAeZg/v2A8owWYPcvAZZg/v2A8o02YPc9sEZg9z+gxmD+/RDyjKDyjCZg9y0QJmD3LQAWYP78gPKMJmD3LQB2YP78hmD+/LZg/vymYP78xmDzgAzg8RCDPAX15bi+VdwggA"  ' 327, 16.12.2023 15:55:18
    Dim lPtr            As LongPtr
        
    lPtr = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If lPtr = 0 Then
        GoTo QH
    End If
    If CryptStringToBinary(StrPtr(STR_THUNK), Len(STR_THUNK), CRYPT_STRING_BASE64, lPtr, THUNK_SIZE) = 0 Then
        GoTo QH
    End If
    pvPatchTrampoline AddressOf pvCallMult
    If pvCallMult(lPtr, 0, 0) = 0 Then '--- checks if PCLMULQDQ instruction supported
        GoTo QH
    End If
    pvPatchTrampoline AddressOf pvMult
    '--- success
    pvThunkAllocate = lPtr
    lPtr = 0
QH:
    If lPtr <> 0 Then
        Call VirtualFree(lPtr, THUNK_SIZE, MEM_DECOMMIT)
    End If
End Function

Private Function pvPatchTrampoline(ByVal Pfn As LongPtr, Optional ByVal Noop As Boolean) As Boolean
    Const PAGE_EXECUTE_READWRITE        As Long = &H40
    Dim bInIDE          As Boolean
 
    If Noop Then
        pvPatchTrampoline = True
        Exit Function
    End If
    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        Call CopyMemory(Pfn, ByVal Pfn + &H16, 4)
    Else
        Call VirtualProtect(Pfn, 8, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0:  58                      pop    eax
    ' 1:  59                      pop    ecx
    ' 2:  50                      push   eax
    ' 3:  ff e1                   jmp    ecx
    ' 5:  90                      nop
    ' 6:  90                      nop
    ' 7:  90                      nop
    Call CopyMemory(ByVal Pfn, -802975883527609.7192@, 8)
    '--- success
    pvPatchTrampoline = True
End Function

Private Function pvSetTrue(bValue As Boolean) As Boolean
    #If TWINBASIC = 0 Then
        bValue = True
    #End If
    pvSetTrue = True
End Function

Private Function pvCallMult(ByVal Pfn As LongPtr, ByVal lPtr1 As LongPtr, ByVal lPtr2 As LongPtr) As Long
    '--- trampoline
End Function
#End If

Private Sub pvPrecompute(baKey() As Byte, uKeyTable As ShoupTable)
    Dim lIdx            As Long
    Dim uOne            As ArrayLong4
    Dim uTemp           As ArrayLong4
    Dim lCarry          As Long
   
    lIdx = UBound(baKey) + 1
    If lIdx > LNG_BLOCKSZ Then
        lIdx = LNG_BLOCKSZ
    End If
    If m_hMulThunk <> 0 Then
        Call CopyMemory(uKeyTable, baKey(0), lIdx)
        Exit Sub
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
        '--- M(1) = H * 1 % POLY
        lIdx = 1
        .Item(m_aReverse(lIdx)) = uOne
        For lIdx = 2 To UBound(.Item)
            If (lIdx And 1) <> 0 Then
                '--- M(i) = M(i - 1) + M(1) % POLY
                uTemp = .Item(m_aReverse(lIdx - 1))
                With uTemp
                    .Item(0) = .Item(0) Xor uOne.Item(0)
                    .Item(1) = .Item(1) Xor uOne.Item(1)
                    .Item(2) = .Item(2) Xor uOne.Item(2)
                    .Item(3) = .Item(3) Xor uOne.Item(3)
                End With
            Else
                '--- M(i) = M(i / 2) * x % POLY
                uTemp = .Item(m_aReverse(lIdx \ 2))
                With uTemp
                    lCarry = .Item(0) And 1
                    #If HasOperators Then
                        .Item(0) = (.Item(0) >> 1) Or (.Item(1) << 31)
                        .Item(1) = (.Item(1) >> 1) Or (.Item(2) << 31)
                        .Item(2) = (.Item(2) >> 1) Or (.Item(3) << 31)
                        .Item(3) = (.Item(3) >> 1) Xor lCarry * m_aReduce(m_aReverse(1))
                    #Else
                        .Item(0) = (.Item(0) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(0) < 0) * LNG_POW2_30 Or (.Item(1) And 1) * LNG_POW2_31
                        .Item(1) = (.Item(1) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(1) < 0) * LNG_POW2_30 Or (.Item(2) And 1) * LNG_POW2_31
                        .Item(2) = (.Item(2) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(2) < 0) * LNG_POW2_30 Or (.Item(3) And 1) * LNG_POW2_31
                        .Item(3) = (.Item(3) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(3) < 0) * LNG_POW2_30 Xor lCarry * m_aReduce(m_aReverse(1))
                    #End If
                End With
            End If
            .Item(m_aReverse(lIdx)) = uTemp
        Next
    End With
End Sub

Private Sub pvMult(ByVal Pfn As LongPtr, uKeyTable As ShoupTable, uArray As ArrayByte16)
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
                #If HasOperators Then
                    .Item(0) = (.Item(0) >> 4) Or (.Item(1) << 28)
                    .Item(1) = (.Item(1) >> 4) Or (.Item(2) << 28)
                    .Item(2) = (.Item(2) >> 4) Or (.Item(3) << 28)
                    .Item(3) = (.Item(3) >> 4) Xor m_aReduce(lCarry)
                #Else
                    .Item(0) = (.Item(0) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(0) < 0) * LNG_POW2_27 _
                        Or (.Item(1) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(1) And LNG_POW2_3) <> 0) * LNG_POW2_31
                    .Item(1) = (.Item(1) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(1) < 0) * LNG_POW2_27 _
                        Or (.Item(2) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(2) And LNG_POW2_3) <> 0) * LNG_POW2_31
                    .Item(2) = (.Item(2) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(2) < 0) * LNG_POW2_27 _
                        Or (.Item(3) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(3) And LNG_POW2_3) <> 0) * LNG_POW2_31
                    .Item(3) = (.Item(3) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(3) < 0) * LNG_POW2_27 _
                        Xor m_aReduce(lCarry)
                #End If
                '--- add lower nibble
                lNibble = uArray.Item(lIdx) And &HF
                .Item(0) = .Item(0) Xor uKeyTable.Item(lNibble).Item(0)
                .Item(1) = .Item(1) Xor uKeyTable.Item(lNibble).Item(1)
                .Item(2) = .Item(2) Xor uKeyTable.Item(lNibble).Item(2)
                .Item(3) = .Item(3) Xor uKeyTable.Item(lNibble).Item(3)
            End If
            '--- mul 16
            lCarry = .Item(0) And &HF
            #If HasOperators Then
                .Item(0) = (.Item(0) >> 4) Or (.Item(1) << 28)
                .Item(1) = (.Item(1) >> 4) Or (.Item(2) << 28)
                .Item(2) = (.Item(2) >> 4) Or (.Item(3) << 28)
                .Item(3) = (.Item(3) >> 4) Xor m_aReduce(lCarry)
            #Else
                .Item(0) = (.Item(0) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(0) < 0) * LNG_POW2_27 _
                    Or (.Item(1) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(1) And LNG_POW2_3) <> 0) * LNG_POW2_31
                .Item(1) = (.Item(1) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(1) < 0) * LNG_POW2_27 _
                    Or (.Item(2) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(2) And LNG_POW2_3) <> 0) * LNG_POW2_31
                .Item(2) = (.Item(2) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(2) < 0) * LNG_POW2_27 _
                    Or (.Item(3) And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((.Item(3) And LNG_POW2_3) <> 0) * LNG_POW2_31
                .Item(3) = (.Item(3) And &H7FFFFFFF) \ LNG_POW2_4 Or -(.Item(3) < 0) * LNG_POW2_27 _
                    Xor m_aReduce(lCarry)
            #End If
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
                #If TWINBASIC = 0 Then
                    Debug.Assert pvPatchTrampoline(AddressOf pvMult, m_hMulThunk = 0)
                #End If
                pvMult m_hMulThunk, uKeyTable, uArray
            End If
        Next
    End With
    pvUpdate = Offset
End Function

Public Sub CryptoGhashInit(uCtx As CryptoGhashContext, baKey() As Byte)
    Dim uEmpty          As ArrayByte16
    
    If m_aReduce(1) = 0 Then
        pvInit
    End If
    With uCtx
        pvPrecompute baKey, .KeyTable
        .HashArray = uEmpty
        .NPosition = 0
    End With
End Sub

Public Sub CryptoGhashGenerCounter(uCtx As CryptoGhashContext, baInput() As Byte, baOutput() As Byte)
    Dim lSize           As Long
    Dim uResult         As ArrayByte16
    Dim uArray          As ArrayByte16
    
    lSize = UBound(baInput) + 1
    If lSize = 12 Then '--- 96 bits
        Call CopyMemory(uResult.Item(0), baInput(0), lSize)
        uResult.Item(LNG_BLOCKSZ - 1) = 1
    Else
        pvUpdate uCtx.KeyTable, uResult, baInput, 0, lSize
        If lSize Mod LNG_BLOCKSZ <> 0 Then
            pvUpdate uCtx.KeyTable, uResult, uArray.Item, 0, LNG_BLOCKSZ - lSize Mod LNG_BLOCKSZ, lSize Mod LNG_BLOCKSZ
        End If
        lSize = BSwap32(lSize * 8)
        Call CopyMemory(uArray.Item(12), lSize, LenB(lSize))
        pvUpdate uCtx.KeyTable, uResult, uArray.Item, 0, LNG_BLOCKSZ
    End If
    If UBound(baOutput) <> LNG_BLOCKSZ - 1 Then
        ReDim baOutput(0 To LNG_BLOCKSZ - 1) As Byte
    End If
    Call CopyMemory(baOutput(0), uResult.Item(0), LNG_BLOCKSZ)
End Sub

Public Sub CryptoGhashUpdate(uCtx As CryptoGhashContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    uCtx.NPosition = pvUpdate(uCtx.KeyTable, uCtx.HashArray, baInput, Pos, Size, Offset:=uCtx.NPosition)
End Sub

Public Sub CryptoGhashPad(uCtx As CryptoGhashContext)
    If uCtx.NPosition > 0 Then
        #If TWINBASIC = 0 Then
            Debug.Assert pvPatchTrampoline(AddressOf pvMult, m_hMulThunk = 0)
        #End If
        pvMult m_hMulThunk, uCtx.KeyTable, uCtx.HashArray
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

Private Function pvFinalize(uCtx As CryptoAesGcmContext, ByVal lTagSize As Long, baTag() As Byte)
    Dim cTemp           As Currency
    Dim aTemp(0 To 1)   As Long
    Dim uBlock          As ArrayLong4
    Dim uArray          As ArrayByte16
    Dim lIdx            As Long
    
    With uCtx
        CryptoGhashPad .GhashCtx
        '--- absorb bit-size of AAD and plaintext
        cTemp = .AadSize * 8@ / 10000@
        Call CopyMemory(aTemp(0), cTemp, 8)
        uBlock.Item(0) = BSwap32(aTemp(1))
        uBlock.Item(1) = BSwap32(aTemp(0))
        cTemp = .TotalSize * 8@ / 10000@
        Call CopyMemory(aTemp(0), cTemp, 8)
        uBlock.Item(2) = BSwap32(aTemp(1))
        uBlock.Item(3) = BSwap32(aTemp(0))
        LSet uArray = uBlock
        CryptoGhashUpdate .GhashCtx, uArray.Item
        '--- finalize hash
        CryptoGhashFinalize .GhashCtx, lTagSize, baTag
        For lIdx = 0 To lTagSize - 1
            baTag(lIdx) = baTag(lIdx) Xor .Counter(lIdx)
        Next
    End With
End Function

Public Sub CryptoAesGcmInit(uCtx As CryptoAesGcmContext, baKey() As Byte, baNonce() As Byte, baAad() As Byte)
    Dim baAuthKey(0 To LNG_BLOCKSZ - 1) As Byte
    
    If UBound(baNonce) + 1 = 0 Then
        Err.Raise vbObjectError, , "Invalid Nonce size for AES-GCM (" & UBound(baNonce) + 1 & ")"
    End If
    With uCtx
        CryptoAesInit uCtx.AesCtx, baKey
        '--- encrypt a block of zeroes to create the hashing key
        CryptoAesProcess .AesCtx, True, baAuthKey
        CryptoGhashInit .GhashCtx, baAuthKey
        CryptoGhashGenerCounter .GhashCtx, baNonce, .Counter
        '--- setup AES counter
        CryptoAesSetNonce .AesCtx, .Counter, CounterWords:=1
        CryptoAesProcess .AesCtx, True, .Counter
        '--- absorb AAD into the hash
        CryptoGhashUpdate .GhashCtx, baAad
        CryptoGhashPad .GhashCtx
        .AadSize = UBound(baAad) + 1
        .TotalSize = 0
    End With
End Sub

Public Sub CryptoAesGcmEncrypt(uCtx As CryptoAesGcmContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional TagSize As Long, Optional Tag As Variant)
    Dim baTag()         As Byte
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    With uCtx
        CryptoAesCtrCrypt .AesCtx, baBuffer, Pos, Size, CounterWords:=1
        CryptoGhashUpdate .GhashCtx, baBuffer, Pos, Size
        .TotalSize = .TotalSize + Size
        If TagSize > 0 Then
            pvFinalize uCtx, TagSize, baTag
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
            pvFinalize uCtx, UBound(baTag) + 1, baCalc
            If InStrB(baTag, baCalc) <> 1 Then
                Exit Function
            End If
        End If
        CryptoAesCtrCrypt .AesCtx, baBuffer, Pos, Size, CounterWords:=1
    End With
    '--- success
    CryptoAesGcmDecrypt = True
End Function

'= POLYVAL ===============================================================

Private Function pvReverseArray(baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim baOutput()      As Byte
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    lJdx = ((Size + LNG_BLOCKSZ - 1) And -LNG_BLOCKSZ) - 1
    ReDim baOutput(0 To lJdx) As Byte
    For lIdx = 0 To Size - 1
        lJdx = (lIdx And -LNG_BLOCKSZ) + (LNG_BLOCKSZ - 1) - (lIdx And (LNG_BLOCKSZ - 1))
        baOutput(lJdx) = baInput(Pos + lIdx)
    Next
    pvReverseArray = baOutput
End Function

Private Function pvMulX(baInput() As Byte) As Byte()
    Dim lIdx            As Long
    Dim uTemp           As ArrayLong4
    Dim lCarry          As Long
    Dim baOutput()      As Byte
    
    If m_aReduce(1) = 0 Then
        pvInit
    End If
    lIdx = UBound(baInput) + 1
    If lIdx > LNG_BLOCKSZ Then
        lIdx = LNG_BLOCKSZ
    End If
    Call CopyMemory(uTemp.Item(0), baInput(0), lIdx)
    With uTemp
        lCarry = .Item(0) And 1
        #If HasOperators Then
            .Item(0) = (.Item(0) >> 1) Or (.Item(1) << 31)
            .Item(1) = (.Item(1) >> 1) Or (.Item(2) << 31)
            .Item(2) = (.Item(2) >> 1) Or (.Item(3) << 31)
            .Item(3) = (.Item(3) >> 1) Xor lCarry * m_aReduce(m_aReverse(1))
        #Else
            .Item(0) = (.Item(0) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(0) < 0) * LNG_POW2_30 Or (.Item(1) And 1) * LNG_POW2_31
            .Item(1) = (.Item(1) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(1) < 0) * LNG_POW2_30 Or (.Item(2) And 1) * LNG_POW2_31
            .Item(2) = (.Item(2) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(2) < 0) * LNG_POW2_30 Or (.Item(3) And 1) * LNG_POW2_31
            .Item(3) = (.Item(3) And &H7FFFFFFF) \ LNG_POW2_1 Or -(.Item(3) < 0) * LNG_POW2_30 Xor lCarry * m_aReduce(m_aReverse(1))
        #End If
    End With
    ReDim baOutput(0 To LNG_BLOCKSZ - 1) As Byte
    Call CopyMemory(baOutput(0), uTemp.Item(0), LNG_BLOCKSZ)
    pvMulX = baOutput
End Function

Public Sub CryptoPolyvalInit(uCtx As CryptoGhashContext, baKey() As Byte)
    CryptoGhashInit uCtx, pvReverseArray(pvMulX(baKey))
End Sub

Public Sub CryptoPolyvalUpdate(uCtx As CryptoGhashContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Const LNG_STEP      As Long = 16 * LNG_BLOCKSZ
    Dim lIdx            As Long
    Dim baTemp(0 To LNG_STEP - 1) As Byte
    Dim lJdx            As Long
    Dim lKdx            As Long
    
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    For lIdx = 0 To Size \ LNG_STEP - 1
        lKdx = Pos + lIdx * LNG_STEP + LNG_STEP - 1
        For lJdx = 0 To LNG_STEP - 1
            baTemp(lJdx) = baInput(lKdx - lJdx)
        Next
        CryptoGhashUpdate uCtx, baTemp
    Next
    If Size > lIdx * LNG_STEP Then
        CryptoGhashUpdate uCtx, pvReverseArray(baInput, Pos + lIdx * LNG_STEP)
    End If
End Sub

Public Sub CryptoPolyvalFinalize(uCtx As CryptoGhashContext, ByVal lTagSize As Long, baTag() As Byte)
    CryptoGhashFinalize uCtx, lTagSize, baTag
    baTag = pvReverseArray(baTag)
End Sub

'= AES-GCM-SIV ===============================================================

Public Sub pvDeriveKeys(uCtx As CryptoAesGcmContext, baKey() As Byte, baNonce() As Byte)
    Const LNG_HALFSZ    As Long = LNG_BLOCKSZ \ 2
    Dim baEncKey()      As Byte
    Dim baAuthKey()     As Byte
    Dim baDerived()     As Byte
    Dim baBlock()       As Byte
    Dim lIdx            As Long
    
    If UBound(baKey) + 1 <> 16 And UBound(baKey) + 1 <> 32 Then
        Err.Raise vbObjectError, , "Invalid key size for AES-GCM-SIV (" & UBound(baKey) + 1 & ")"
    End If
    If UBound(baNonce) + 1 <> 12 Then
        Err.Raise vbObjectError, , "Invalid nonce size for AES-GCM-SIV (" & UBound(baNonce) + 1 & ")"
    End If
    With uCtx
        CryptoAesInit uCtx.AesCtx, baKey
        ReDim baEncKey(0 To UBound(baKey)) As Byte
        ReDim baAuthKey(0 To LNG_BLOCKSZ - 1) As Byte
        ReDim baDerived(0 To LNG_BLOCKSZ + UBound(baEncKey)) As Byte
        ReDim baBlock(0 To LNG_BLOCKSZ - 1) As Byte
        For lIdx = 0 To UBound(baDerived) \ LNG_HALFSZ
            Call CopyMemory(baBlock(0), lIdx, LenB(lIdx))
            Call CopyMemory(baBlock(4), baNonce(0), UBound(baNonce) + 1)
            CryptoAesProcess .AesCtx, True, baBlock
            Call CopyMemory(baDerived(lIdx * LNG_HALFSZ), baBlock(0), LNG_HALFSZ)
        Next
        Call CopyMemory(baAuthKey(0), baDerived(0), LNG_BLOCKSZ)
        Call CopyMemory(baEncKey(0), baDerived(LNG_BLOCKSZ), UBound(baKey) + 1)
        CryptoAesInit uCtx.AesCtx, baEncKey
        CryptoPolyvalInit .GhashCtx, baAuthKey
    End With
End Sub

Public Sub CryptoAesGcmSivEncrypt(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baBuffer() As Byte, baTag() As Byte)
    Dim uCtx            As CryptoAesGcmContext
    Dim baTemp()        As Byte
    Dim cTemp           As Currency
    Dim lIdx            As Long
    
    pvDeriveKeys uCtx, baKey, baNonce
    With uCtx
        CryptoPolyvalUpdate .GhashCtx, baAad
        CryptoPolyvalUpdate .GhashCtx, baBuffer
        ReDim baTemp(0 To LNG_BLOCKSZ - 1) As Byte
        cTemp = (UBound(baAad) + 1) * 8@ / 10000@
        Call CopyMemory(baTemp(0), cTemp, 8)
        cTemp = (UBound(baBuffer) + 1) * 8@ / 10000@
        Call CopyMemory(baTemp(8), cTemp, 8)
        CryptoPolyvalUpdate .GhashCtx, baTemp
        CryptoPolyvalFinalize .GhashCtx, LNG_BLOCKSZ, baTemp
        For lIdx = 0 To UBound(baNonce)
            baTemp(lIdx) = baTemp(lIdx) Xor baNonce(lIdx)
        Next
        baTemp(15) = baTemp(15) And &H7F
        CryptoAesProcess .AesCtx, True, baTemp
        baTag = baTemp
        baTemp(15) = baTemp(15) Or &H80
        CryptoAesSetNonce .AesCtx, baTemp
        CryptoAesCtrCrypt .AesCtx, baBuffer, CounterWords:=-1
    End With
End Sub

Public Function CryptoAesGcmSivDecrypt(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baBuffer() As Byte, baTag() As Byte) As Boolean
    Dim uCtx            As CryptoAesGcmContext
    Dim baTemp()        As Byte
    Dim cTemp           As Currency
    Dim lIdx            As Long
    
    If UBound(baTag) + 1 <> 16 Then
        Err.Raise vbObjectError, , "Invalid tag size for AES-GCM-SIV (" & UBound(baTag) + 1 & ")"
    End If
    pvDeriveKeys uCtx, baKey, baNonce
    With uCtx
        baTemp = baTag
        baTemp(15) = baTemp(15) Or &H80
        CryptoAesSetNonce .AesCtx, baTemp
        CryptoAesCtrCrypt .AesCtx, baBuffer, CounterWords:=-1
        CryptoPolyvalUpdate .GhashCtx, baAad
        CryptoPolyvalUpdate .GhashCtx, baBuffer
        cTemp = (UBound(baAad) + 1) * 8@ / 10000@
        Call CopyMemory(baTemp(0), cTemp, 8)
        cTemp = (UBound(baBuffer) + 1) * 8@ / 10000@
        Call CopyMemory(baTemp(8), cTemp, 8)
        CryptoPolyvalUpdate .GhashCtx, baTemp
        CryptoPolyvalFinalize .GhashCtx, LNG_BLOCKSZ, baTemp
        For lIdx = 0 To UBound(baNonce)
            baTemp(lIdx) = baTemp(lIdx) Xor baNonce(lIdx)
        Next
        baTemp(15) = baTemp(15) And &H7F
        CryptoAesProcess .AesCtx, True, baTemp
        If InStrB(baTemp, baTag) <> 1 Then
            GoTo QH
        End If
    End With
    '--- success
    CryptoAesGcmSivDecrypt = True
QH:
End Function
