Attribute VB_Name = "mdBlake3"
'=========================================================================
'
'  Pure VB6 Crypto (Untested)
'  Copyright (c) 2026 wqweto@gmail.com
'
'  This project is licensed under the terms of the MIT license
'  See the LICENSE file in the project root for more information
'
'=========================================================================
'--- mdBlake3.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_OUT_LEN               As Long = 32
Private Const LNG_KEY_LEN               As Long = 32
Private Const LNG_BLOCK_LEN             As Long = 64
Private Const LNG_CHUNK_LEN             As Long = 1024

Private Type ArrayLong8
    Item(0 To 7) As Long
End Type

Private Type ArrayLong16
    Item(0 To 15) As Long
End Type

Private Enum Blake3Flags
    LNG_CHUNK_START = 2 ^ 0
    LNG_CHUNK_END = 2 ^ 1
    LNG_PARENT = 2 ^ 2
    LNG_ROOT = 2 ^ 3
    LNG_KEYED_HASH = 2 ^ 4
    LNG_DERIVE_KEY_CONTEXT = 2 ^ 5
    LNG_DERIVE_KEY_MATERIAL = 2 ^ 6
End Enum

Private Type Blake3ChunkState
    ChainingValue           As ArrayLong8
    ChunkCounter            As Long
    Block(0 To LNG_BLOCK_LEN - 1) As Byte
    BlockLen                As Byte
    BlocksCompressed        As Byte
    Flags                   As Blake3Flags
End Type

Private Type Blake3Output
    InputChainingValue      As ArrayLong8
    BlockWords              As ArrayLong16
    Counter                 As Currency
    BlockLen                As Byte
    Flags                   As Blake3Flags
End Type

Public Type CryptoBlake3Context
    ChunkState              As Blake3ChunkState
    KeyWords                As ArrayLong8
    CvStack(0 To 53)        As ArrayLong8
    CvStackLen              As Byte
    Flags                   As Blake3Flags
End Type

Private LNG_IV                      As ArrayLong8

#If Not HasOperators Then
Private LNG_POW2(0 To 31)           As Long
Private m_bNoIntegerOverflowChecks  As Boolean

Private Function RotR32(ByVal lX As Long, ByVal lN As Long) As Long
    '--- RotR32 = RShift32(X, n) Or LShift32(X, 32 - n)
    Debug.Assert lN <> 0
    RotR32 = ((lX And &H7FFFFFFF) \ LNG_POW2(lN) - (lX < 0) * LNG_POW2(31 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(32 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * &H80000000)
End Function

Private Function UAdd32(ByVal lX As Long, ByVal lY As Long) As Long
    If (lX Xor lY) >= 0 Then
        UAdd32 = ((lX Xor &H80000000) + lY) Xor &H80000000
    Else
        UAdd32 = lX + lY
    End If
End Function

Private Function pvGetOverflowIgnored(Optional bValue As Boolean = True) As Boolean
    Dim bInIde      As Boolean
    
    If Not bValue Then
        bValue = True
        pvGetOverflowIgnored = True
        Exit Function
    End If
    Debug.Assert pvGetOverflowIgnored(bInIde)
    If bInIde Then
        Exit Function
    End If
    On Error GoTo EH
    If &H8000 - 1 <> 0 Then
        pvGetOverflowIgnored = True
    End If
EH:
End Function

Private Sub pvQuarter32(lA As Long, lB As Long, lC As Long, lD As Long, ByVal lX As Long, ByVal lY As Long)
    If m_bNoIntegerOverflowChecks Then
        lA = lA + lB + lX
        lD = RotR32(lD Xor lA, 16)
        lC = lC + lD
        lB = RotR32(lB Xor lC, 12)
        lA = lA + lB + lY
        lD = RotR32(lD Xor lA, 8)
        lC = lC + lD
        lB = RotR32(lB Xor lC, 7)
    Else
        lA = UAdd32(UAdd32(lA, lB), lX)
        lD = RotR32(lD Xor lA, 16)
        lC = UAdd32(lC, lD)
        lB = RotR32(lB Xor lC, 12)
        lA = UAdd32(UAdd32(lA, lB), lY)
        lD = RotR32(lD Xor lA, 8)
        lC = UAdd32(lC, lD)
        lB = RotR32(lB Xor lC, 7)
    End If
End Sub
#Else
[ IntegerOverflowChecks (False) ]
Private Sub pvQuarter32(lA As Long, lB As Long, lC As Long, lD As Long, ByVal lX As Long, ByVal lY As Long)
    lA = lA + lB + lX
    lD = (lD Xor lA) >> 16 Or (lD Xor lA) << 16
    lC = lC + lD
    lB = (lB Xor lC) >> 12 Or (lB Xor lC) << 20
    lA = lA + lB + lY
    lD = (lD Xor lA) >> 8 Or (lD Xor lA) << 24
    lC = lC + lD
    lB = (lB Xor lC) >> 7 Or (lB Xor lC) << 25
End Sub
#End If

Private Sub pvCompress(uState As ArrayLong8, uBlock As ArrayLong16, ByVal cCounter As Currency, ByVal lBlockLen As Long, ByVal eFlags As Blake3Flags, uRetVal As ArrayLong16, Optional ByVal HalfOnly As Boolean)
    Static S(0 To 1)    As Long
    Dim V0              As Long
    Dim V1              As Long
    Dim V2              As Long
    Dim V3              As Long
    Dim V4              As Long
    Dim V5              As Long
    Dim V6              As Long
    Dim V7              As Long
    Dim V8              As Long
    Dim V9              As Long
    Dim V10             As Long
    Dim V11             As Long
    Dim V12             As Long
    Dim V13             As Long
    Dim V14             As Long
    Dim V15             As Long
    
    With uState
        V0 = .Item(0):  V1 = .Item(1)
        V2 = .Item(2):  V3 = .Item(3)
        V4 = .Item(4):  V5 = .Item(5)
        V6 = .Item(6):  V7 = .Item(7)
    End With
    With LNG_IV
        V8 = .Item(0):  V9 = .Item(1)
        V10 = .Item(2): V11 = .Item(3)
    End With
    cCounter = cCounter / 10000@
    Call CopyMemory(S(0), cCounter, 8)
    V12 = S(0)
    V13 = S(1)
    V14 = lBlockLen
    V15 = eFlags
    With uBlock
        '--- Round 1
        pvQuarter32 V0, V4, V8, V12, .Item(0), .Item(1)
        pvQuarter32 V1, V5, V9, V13, .Item(2), .Item(3)
        pvQuarter32 V2, V6, V10, V14, .Item(4), .Item(5)
        pvQuarter32 V3, V7, V11, V15, .Item(6), .Item(7)
        pvQuarter32 V0, V5, V10, V15, .Item(8), .Item(9)
        pvQuarter32 V1, V6, V11, V12, .Item(10), .Item(11)
        pvQuarter32 V2, V7, V8, V13, .Item(12), .Item(13)
        pvQuarter32 V3, V4, V9, V14, .Item(14), .Item(15)
        '--- Round 2
        pvQuarter32 V0, V4, V8, V12, .Item(2), .Item(6)
        pvQuarter32 V1, V5, V9, V13, .Item(3), .Item(10)
        pvQuarter32 V2, V6, V10, V14, .Item(7), .Item(0)
        pvQuarter32 V3, V7, V11, V15, .Item(4), .Item(13)
        pvQuarter32 V0, V5, V10, V15, .Item(1), .Item(11)
        pvQuarter32 V1, V6, V11, V12, .Item(12), .Item(5)
        pvQuarter32 V2, V7, V8, V13, .Item(9), .Item(14)
        pvQuarter32 V3, V4, V9, V14, .Item(15), .Item(8)
        '--- Round 3
        pvQuarter32 V0, V4, V8, V12, .Item(3), .Item(4)
        pvQuarter32 V1, V5, V9, V13, .Item(10), .Item(12)
        pvQuarter32 V2, V6, V10, V14, .Item(13), .Item(2)
        pvQuarter32 V3, V7, V11, V15, .Item(7), .Item(14)
        pvQuarter32 V0, V5, V10, V15, .Item(6), .Item(5)
        pvQuarter32 V1, V6, V11, V12, .Item(9), .Item(0)
        pvQuarter32 V2, V7, V8, V13, .Item(11), .Item(15)
        pvQuarter32 V3, V4, V9, V14, .Item(8), .Item(1)
        '--- Round 4
        pvQuarter32 V0, V4, V8, V12, .Item(10), .Item(7)
        pvQuarter32 V1, V5, V9, V13, .Item(12), .Item(9)
        pvQuarter32 V2, V6, V10, V14, .Item(14), .Item(3)
        pvQuarter32 V3, V7, V11, V15, .Item(13), .Item(15)
        pvQuarter32 V0, V5, V10, V15, .Item(4), .Item(0)
        pvQuarter32 V1, V6, V11, V12, .Item(11), .Item(2)
        pvQuarter32 V2, V7, V8, V13, .Item(5), .Item(8)
        pvQuarter32 V3, V4, V9, V14, .Item(1), .Item(6)
        '--- Round 5
        pvQuarter32 V0, V4, V8, V12, .Item(12), .Item(13)
        pvQuarter32 V1, V5, V9, V13, .Item(9), .Item(11)
        pvQuarter32 V2, V6, V10, V14, .Item(15), .Item(10)
        pvQuarter32 V3, V7, V11, V15, .Item(14), .Item(8)
        pvQuarter32 V0, V5, V10, V15, .Item(7), .Item(2)
        pvQuarter32 V1, V6, V11, V12, .Item(5), .Item(3)
        pvQuarter32 V2, V7, V8, V13, .Item(0), .Item(1)
        pvQuarter32 V3, V4, V9, V14, .Item(6), .Item(4)
        '--- Round 6
        pvQuarter32 V0, V4, V8, V12, .Item(9), .Item(14)
        pvQuarter32 V1, V5, V9, V13, .Item(11), .Item(5)
        pvQuarter32 V2, V6, V10, V14, .Item(8), .Item(12)
        pvQuarter32 V3, V7, V11, V15, .Item(15), .Item(1)
        pvQuarter32 V0, V5, V10, V15, .Item(13), .Item(3)
        pvQuarter32 V1, V6, V11, V12, .Item(0), .Item(10)
        pvQuarter32 V2, V7, V8, V13, .Item(2), .Item(6)
        pvQuarter32 V3, V4, V9, V14, .Item(4), .Item(7)
        '--- Round 7
        pvQuarter32 V0, V4, V8, V12, .Item(11), .Item(15)
        pvQuarter32 V1, V5, V9, V13, .Item(5), .Item(0)
        pvQuarter32 V2, V6, V10, V14, .Item(1), .Item(9)
        pvQuarter32 V3, V7, V11, V15, .Item(8), .Item(6)
        pvQuarter32 V0, V5, V10, V15, .Item(14), .Item(10)
        pvQuarter32 V1, V6, V11, V12, .Item(2), .Item(12)
        pvQuarter32 V2, V7, V8, V13, .Item(3), .Item(4)
        pvQuarter32 V3, V4, V9, V14, .Item(7), .Item(13)
    End With
    With uRetVal
        .Item(0) = V0 Xor V8: .Item(1) = V1 Xor V9
        .Item(2) = V2 Xor V10: .Item(3) = V3 Xor V11
        .Item(4) = V4 Xor V12: .Item(5) = V5 Xor V13
        .Item(6) = V6 Xor V14: .Item(7) = V7 Xor V15
        If Not HalfOnly Then
            .Item(8) = V8 Xor uState.Item(0)
            .Item(9) = V9 Xor uState.Item(1)
            .Item(10) = V10 Xor uState.Item(2)
            .Item(11) = V11 Xor uState.Item(3)
            .Item(12) = V12 Xor uState.Item(4)
            .Item(13) = V13 Xor uState.Item(5)
            .Item(14) = V14 Xor uState.Item(6)
            .Item(15) = V15 Xor uState.Item(7)
        End If
    End With
End Sub

Private Sub pvUpdateChunk(uChunk As Blake3ChunkState, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Dim eStartFlag      As Blake3Flags
    Dim lRemaining      As Long
    
    With uChunk
        Do While lSize > 0
            If .BlockLen = LNG_BLOCK_LEN Then
                eStartFlag = -(.BlocksCompressed = 0) * LNG_CHUNK_START
                #If HasOperators Then
                    pvCompress .ChainingValue, VarPtr(.Block(0)), .ChunkCounter, .BlockLen, .Flags Or eStartFlag, VarPtr(.ChainingValue), HalfOnly:=True
                #Else
                    Static uTemp As ArrayLong16
                    Call CopyMemory(uTemp, .Block(0), LNG_BLOCK_LEN)
                    pvCompress .ChainingValue, uTemp, .ChunkCounter, .BlockLen, .Flags Or eStartFlag, uTemp
                    Call CopyMemory(.ChainingValue, uTemp, LNG_BLOCK_LEN \ 2)
                #End If
                .BlocksCompressed = .BlocksCompressed + 1
                .BlockLen = 0
            End If
            lRemaining = LNG_BLOCK_LEN - .BlockLen
            If lRemaining > lSize Then
                lRemaining = lSize
            End If
            Call CopyMemory(.Block(.BlockLen), baInput(lPos), lRemaining)
            .BlockLen = .BlockLen + lRemaining
            lPos = lPos + lRemaining
            lSize = lSize - lRemaining
        Loop
    End With
End Sub

Private Sub pvGetChunkOutput(uChunk As Blake3ChunkState, uOutput As Blake3Output)
    Dim eStartFlag      As Blake3Flags
    
    With uChunk
        uOutput.InputChainingValue = .ChainingValue
        If .BlockLen > 0 Then
            Call CopyMemory(uOutput.BlockWords, .Block(0), .BlockLen)
        End If
        uOutput.Counter = .ChunkCounter
        uOutput.BlockLen = .BlockLen
        eStartFlag = -(.BlocksCompressed = 0) * LNG_CHUNK_START
        uOutput.Flags = .Flags Or eStartFlag Or LNG_CHUNK_END
    End With
End Sub

Private Function pvGetChunkLen(uChunk As Blake3ChunkState) As Long
    With uChunk
        pvGetChunkLen = .BlocksCompressed * LNG_BLOCK_LEN + .BlockLen
    End With
End Function

Private Sub pvMakeParentOutput(uLeft As ArrayLong8, uRight As ArrayLong8, uKeyWords As ArrayLong8, ByVal eFlags As Blake3Flags, uOutput As Blake3Output)
    With uOutput
        .InputChainingValue = uKeyWords
        Call CopyMemory(.BlockWords.Item(0), uLeft, LNG_BLOCK_LEN \ 2)
        Call CopyMemory(.BlockWords.Item(8), uRight, LNG_BLOCK_LEN \ 2)
        .Counter = 0
        .BlockLen = LNG_BLOCK_LEN
        .Flags = eFlags Or LNG_PARENT
    End With
End Sub

Private Sub pvGetChainingValue(uOutput As Blake3Output, uRetVal As ArrayLong8)
    With uOutput
        #If HasOperators Then
            pvCompress .InputChainingValue, .BlockWords, .Counter, .BlockLen, .Flags, VarPtr(uRetVal), HalfOnly:=True
        #Else
            Static uTemp As ArrayLong16
            pvCompress .InputChainingValue, .BlockWords, .Counter, .BlockLen, .Flags, uTemp
            Call CopyMemory(uRetVal, uTemp, LNG_BLOCK_LEN \ 2)
        #End If
    End With
End Sub

Private Sub pvGetRootBytes(uOutput As Blake3Output, baOutput() As Byte, ByVal lOutSize As Long)
    Dim uTemp           As ArrayLong16
    Dim cCounter        As Currency
    Dim lPos            As Long
    Dim lRemaining      As Long
    
    With uOutput
        ReDim baOutput(0 To lOutSize - 1) As Byte
        Do While lPos < lOutSize
            pvCompress .InputChainingValue, .BlockWords, cCounter, .BlockLen, .Flags Or LNG_ROOT, uTemp
            lRemaining = lOutSize - lPos
            If lRemaining > LNG_BLOCK_LEN Then
                lRemaining = LNG_BLOCK_LEN
            End If
            Call CopyMemory(baOutput(lPos), uTemp, lRemaining)
            lPos = lPos + lRemaining
            cCounter = cCounter + 1
        Loop
    End With
End Sub

Private Sub pvInitHasher(uCtx As CryptoBlake3Context, ByVal lKeyPtr As LongPtr, Optional ByVal eFlags As Blake3Flags)
    Call FillMemory(uCtx, LenB(uCtx), 0)
    With uCtx
        Call CopyMemory(.KeyWords, ByVal lKeyPtr, LNG_KEY_LEN)
        .Flags = eFlags
        .ChunkState.ChainingValue = .KeyWords
        .ChunkState.Flags = .Flags
    End With
End Sub

Public Sub CryptoBlake3Init(uCtx As CryptoBlake3Context, Optional Key As Variant, Optional Context As Variant)
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim baKey()         As Byte
    Dim baContext()     As Byte
    
    If LNG_IV.Item(0) = 0 Then
        For Each vElem In Split("6A09E667 BB67AE85 3C6EF372 A54FF53A 510E527F 9B05688C 1F83D9AB 5BE0CD19")
            LNG_IV.Item(lIdx) = "&H" & vElem
            lIdx = lIdx + 1
        Next
        #If Not HasOperators Then
            LNG_POW2(0) = 1
            For lIdx = 1 To 30
                LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
            Next
            LNG_POW2(31) = &H80000000
            m_bNoIntegerOverflowChecks = pvGetOverflowIgnored
        #End If
    End If
    With uCtx
        If Not IsMissing(Key) Then
            If IsArray(Key) Then
                baKey = Key
            Else
                baKey = ToUtf8Array(CStr(Key))
            End If
            ReDim Preserve baKey(0 To LNG_KEY_LEN - 1) As Byte
            pvInitHasher uCtx, VarPtr(baKey(0)), LNG_KEYED_HASH
        ElseIf Not IsMissing(Context) Then
            If IsArray(Context) Then
                baContext = Context
            Else
                baContext = ToUtf8Array(CStr(Context))
            End If
            pvInitHasher uCtx, VarPtr(LNG_IV), LNG_DERIVE_KEY_CONTEXT
            CryptoBlake3Update uCtx, baContext
            CryptoBlake3Finalize uCtx, baKey
            pvInitHasher uCtx, VarPtr(baKey(0)), LNG_DERIVE_KEY_MATERIAL
        Else
            pvInitHasher uCtx, VarPtr(LNG_IV)
        End If
    End With
End Sub

Public Sub CryptoBlake3Update(uCtx As CryptoBlake3Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim uOutput         As Blake3Output
    Dim uRight          As ArrayLong8
    Dim lTotalChunks    As Long
    Dim lRemaining      As Long
    
    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        Do While Size > 0
            If pvGetChunkLen(.ChunkState) = LNG_CHUNK_LEN Then
                pvGetChunkOutput .ChunkState, uOutput
                pvGetChainingValue uOutput, uRight
                lTotalChunks = .ChunkState.ChunkCounter + 1
                Do While (lTotalChunks And 1) = 0
                    .CvStackLen = .CvStackLen - 1
                    pvMakeParentOutput .CvStack(.CvStackLen), uRight, .KeyWords, .Flags, uOutput
                    pvGetChainingValue uOutput, uRight
                    lTotalChunks = lTotalChunks \ 2
                Loop
                .CvStack(.CvStackLen) = uRight
                .CvStackLen = .CvStackLen + 1
                .ChunkState.ChainingValue = .KeyWords
                .ChunkState.ChunkCounter = .ChunkState.ChunkCounter + 1
                .ChunkState.BlockLen = 0
                .ChunkState.BlocksCompressed = 0
                .ChunkState.Flags = .Flags
            End If
            lRemaining = LNG_CHUNK_LEN - pvGetChunkLen(.ChunkState)
            If lRemaining > Size Then
                lRemaining = Size
            End If
            pvUpdateChunk .ChunkState, baInput, Pos, lRemaining
            Pos = Pos + lRemaining
            Size = Size - lRemaining
        Loop
    End With
End Sub

Public Sub CryptoBlake3Finalize(uCtx As CryptoBlake3Context, baOutput() As Byte, Optional ByVal OutSize As Long)
    Dim uOutput         As Blake3Output
    Dim uRight          As ArrayLong8
    
    With uCtx
        pvGetChunkOutput .ChunkState, uOutput
        Do While .CvStackLen > 0
            pvGetChainingValue uOutput, uRight
            .CvStackLen = .CvStackLen - 1
            pvMakeParentOutput .CvStack(.CvStackLen), uRight, .KeyWords, .Flags, uOutput
        Loop
        If OutSize <= 0 Then
            OutSize = LNG_OUT_LEN
        End If
        pvGetRootBytes uOutput, baOutput, OutSize
    End With
End Sub

Public Function CryptoBlake3ByteArray(baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional Key As Variant, Optional Context As Variant, Optional OutSize As Long) As Byte()
    Dim uCtx            As CryptoBlake3Context
    
    CryptoBlake3Init uCtx, Key:=Key, Context:=Context
    CryptoBlake3Update uCtx, baInput, Pos, Size
    CryptoBlake3Finalize uCtx, CryptoBlake3ByteArray, OutSize:=OutSize
End Function

Private Function ToUtf8Array(sText As String) As Byte()
    Const CP_UTF8       As Long = 65001
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    ReDim baRetVal(0 To 4 * Len(sText)) As Byte
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), UBound(baRetVal) + 1, 0, 0)
    If lSize > 0 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function

Private Function ToHex(baData() As Byte) As String
    Dim lIdx            As Long
    Dim sByte           As String
    
    ToHex = String$(UBound(baData) * 2 + 2, 48)
    For lIdx = 0 To UBound(baData)
        sByte = LCase$(Hex$(baData(lIdx)))
        Mid$(ToHex, lIdx * 2 + 3 - Len(sByte)) = sByte
    Next
End Function

Public Function CryptoBlake3Text(sText As String, Optional Key As Variant, Optional Context As Variant, Optional OutSize As Long) As String
    CryptoBlake3Text = ToHex(CryptoBlake3ByteArray(ToUtf8Array(sText), Key:=Key, Context:=Context, OutSize:=OutSize))
End Function
