Attribute VB_Name = "mdArgon2"
'--- mdArgon2.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)
#Const HasLongLong = (VBA7 <> 0 And WIN64 <> 0 Or TWINBASIC <> 0)
#Const DebugMode = False

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_BLOCKSZ               As Long = 1024
Private Const LNG_ARRAYSZ               As Long = LNG_BLOCKSZ \ 8
Private Const LNG_HASHSZ                As Long = 64
Private Const LNG_SYNC_POINTS           As Long = 4
Private Const LNG_VERSION               As Long = &H13

Private Enum Argon2ModeEnum
    LNG_MODE_Argon2d = 0
    LNG_MODE_Argon2i
    LNG_MODE_Argon2id
End Enum

Private Type ArrayLong128
#If HasLongLong Then
    Item(0 To LNG_ARRAYSZ - 1)  As LongLong
#Else
    Item(0 To LNG_ARRAYSZ - 1)  As Variant
#End If
End Type

#If Not HasOperators Then
#If HasLongLong Then
Private LNG_ZERO                    As LongLong
Private LNG_POW2(0 To 63)           As LongLong
Private LNG_SIGN_BIT                As LongLong ' 2 ^ 63
Private LNG_UINT_MAX                As LongLong ' 2 ^ 32 - 1
#Else
Private LNG_ZERO                    As Variant
Private LNG_POW2(0 To 63)           As Variant
Private LNG_SIGN_BIT                As Variant
Private LNG_UINT_MAX                As Variant
#End If

#If HasLongLong Then
Private Function RotR64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RotR64(lX As Variant, ByVal lN As Long) As Variant
#End If
    '--- RotR64 = RShift64(X, n) Or LShift64(X, 64 - n)
    Debug.Assert lN <> 0
    RotR64 = ((lX And (-1 Xor LNG_SIGN_BIT)) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(63 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(64 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * LNG_SIGN_BIT)
End Function

#If HasLongLong Then
Private Function RShift64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RShift64(lX As Variant, ByVal lN As Long) As Variant
#End If
    If lN = 0 Then
        RShift64 = lX
    Else
        RShift64 = (lX And (-1 Xor LNG_SIGN_BIT)) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(63 - lN)
    End If
End Function

#If HasLongLong Then
Private Function UAdd64(ByVal lX As LongLong, ByVal lY As LongLong) As LongLong
#Else
Private Function UAdd64(lX As Variant, lY As Variant) As Variant
#End If
    If (lX Xor lY) >= 0 Then
        UAdd64 = ((lX Xor LNG_SIGN_BIT) + lY) Xor LNG_SIGN_BIT
    Else
        UAdd64 = lX + lY
    End If
End Function

#If HasLongLong Then
Private Function UMul64(ByVal lX As LongLong, ByVal lY As LongLong) As LongLong
#Else
Private Function UMul64(lX As Variant, lY As Variant) As Variant
#End If
    UMul64 = (lX And &H7FFFFFFF) * (lY And LNG_UINT_MAX)
    If (lX And LNG_POW2(31)) <> 0 Then
        UMul64 = UAdd64(UMul64, (lX And LNG_POW2(31)) * (lY And LNG_UINT_MAX))
    End If
End Function

#If HasLongLong Then
Private Sub pvQuarter64(lA As LongLong, lB As LongLong, lC As LongLong, lD As LongLong)
    Dim lX              As LongLong
#Else
Private Sub pvQuarter64(lA As Variant, lB As Variant, lC As Variant, lD As Variant)
    Dim lX              As Variant
#End If
    lX = UMul64(lA, lB)
    lA = UAdd64(UAdd64(UAdd64(lA, lB), lX), lX)
    lD = RotR64(lD Xor lA, 32)
    lX = UMul64(lC, lD)
    lC = UAdd64(UAdd64(UAdd64(lC, lD), lX), lX)
    lB = RotR64(lB Xor lC, 24)
    lX = UMul64(lA, lB)
    lA = UAdd64(UAdd64(UAdd64(lA, lB), lX), lX)
    lD = RotR64(lD Xor lA, 16)
    lX = UMul64(lC, lD)
    lC = UAdd64(UAdd64(UAdd64(lC, lD), lX), lX)
    lB = RotR64(lB Xor lC, 63)
End Sub
#Else
Private Const LNG_ZERO              As LongLong = 0
Private Const LNG_UINT_MAX          As LongLong = 2 ^ 32 - 1

[ IntegerOverflowChecks (False) ]
Private Sub pvQuarter64(lA As LongLong, lB As LongLong, lC As LongLong, lD As LongLong)
    Dim lX              As LongLong
    
    lX = (lA And LNG_UINT_MAX) * (lB And LNG_UINT_MAX)
    lA = lA + lB + lX + lX
    lD = (lD Xor lA) >> 32 Or (lD Xor lA) << 32
    lX = (lC And LNG_UINT_MAX) * (lD And LNG_UINT_MAX)
    lC = lC + lD + lX + lX
    lB = (lB Xor lC) >> 24 Or (lB Xor lC) << 40
    lX = (lA And LNG_UINT_MAX) * (lB And LNG_UINT_MAX)
    lA = lA + lB + lX + lX
    lD = (lD Xor lA) >> 16 Or (lD Xor lA) << 48
    lX = (lC And LNG_UINT_MAX) * (lD And LNG_UINT_MAX)
    lC = lC + lD + lX + lX
    lB = (lB Xor lC) >> 63 Or (lB Xor lC) << 1
End Sub
#End If

#If Not HasLongLong Then
Private Function CLngLng(vValue As Variant) As Variant
    Const VT_I8 As Long = &H14
    Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
End Function
#End If

Private Sub pvFillBlock(uA As ArrayLong128, uB As ArrayLong128, uOut As ArrayLong128, Optional ByVal IsXor As Boolean)
    Static W            As ArrayLong128
    Dim lIdx            As Long
    
    With W
        For lIdx = 0 To LNG_ARRAYSZ - 1
            W.Item(lIdx) = uA.Item(lIdx) Xor uB.Item(lIdx)
        Next
        For lIdx = 0 To LNG_ARRAYSZ - 1 Step 16
            pvQuarter64 .Item(lIdx + 0), .Item(lIdx + 4), .Item(lIdx + 8), .Item(lIdx + 12)
            pvQuarter64 .Item(lIdx + 1), .Item(lIdx + 5), .Item(lIdx + 9), .Item(lIdx + 13)
            pvQuarter64 .Item(lIdx + 2), .Item(lIdx + 6), .Item(lIdx + 10), .Item(lIdx + 14)
            pvQuarter64 .Item(lIdx + 3), .Item(lIdx + 7), .Item(lIdx + 11), .Item(lIdx + 15)
            pvQuarter64 .Item(lIdx + 0), .Item(lIdx + 5), .Item(lIdx + 10), .Item(lIdx + 15)
            pvQuarter64 .Item(lIdx + 1), .Item(lIdx + 6), .Item(lIdx + 11), .Item(lIdx + 12)
            pvQuarter64 .Item(lIdx + 2), .Item(lIdx + 7), .Item(lIdx + 8), .Item(lIdx + 13)
            pvQuarter64 .Item(lIdx + 3), .Item(lIdx + 4), .Item(lIdx + 9), .Item(lIdx + 14)
        Next
        For lIdx = 0 To LNG_ARRAYSZ \ 8 - 1 Step 2
            pvQuarter64 .Item(lIdx + 0), .Item(lIdx + 32), .Item(lIdx + 64), .Item(lIdx + 96)
            pvQuarter64 .Item(lIdx + 1), .Item(lIdx + 33), .Item(lIdx + 65), .Item(lIdx + 97)
            pvQuarter64 .Item(lIdx + 16), .Item(lIdx + 48), .Item(lIdx + 80), .Item(lIdx + 112)
            pvQuarter64 .Item(lIdx + 17), .Item(lIdx + 49), .Item(lIdx + 81), .Item(lIdx + 113)
            pvQuarter64 .Item(lIdx + 0), .Item(lIdx + 33), .Item(lIdx + 80), .Item(lIdx + 113)
            pvQuarter64 .Item(lIdx + 1), .Item(lIdx + 48), .Item(lIdx + 81), .Item(lIdx + 96)
            pvQuarter64 .Item(lIdx + 16), .Item(lIdx + 49), .Item(lIdx + 64), .Item(lIdx + 97)
            pvQuarter64 .Item(lIdx + 17), .Item(lIdx + 32), .Item(lIdx + 65), .Item(lIdx + 112)
        Next
        If IsXor Then
            For lIdx = 0 To LNG_ARRAYSZ - 1
                uOut.Item(lIdx) = uOut.Item(lIdx) Xor uA.Item(lIdx) Xor uB.Item(lIdx) Xor W.Item(lIdx)
            Next
        Else
            For lIdx = 0 To LNG_ARRAYSZ - 1
                uOut.Item(lIdx) = uA.Item(lIdx) Xor uB.Item(lIdx) Xor W.Item(lIdx)
            Next
        End If
    End With
End Sub

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
#If HasLongLong Then
Private Function pvIndexAlpha(ByVal lRandom As LongLong, ByVal lLanes As Long, ByVal lSegments As Long, ByVal lThreads As Long, ByVal lN As Long, ByVal lSlice As Long, ByVal lLane As Long, ByVal lIndex As Long) As Long
    Dim lP              As LongLong
#Else
Private Function pvIndexAlpha(lRandom As Variant, ByVal lLanes As Long, ByVal lSegments As Long, ByVal lThreads As Long, ByVal lN As Long, ByVal lSlice As Long, ByVal lLane As Long, ByVal lIndex As Long) As Long
    Dim lP              As Variant
#End If
    Dim lRefLane        As Long
    Dim lM              As Long
    Dim lS              As Long
    
    #If HasOperators Then
        lRefLane = CLng((lRandom >> 32) Mod lThreads)
    #Else
        lRefLane = CLng(RShift64(lRandom, 32) Mod lThreads)
    #End If
    If lN = 0 And lSlice = 0 Then
        lRefLane = lLane
    End If
    lM = 3 * lSegments
    lS = ((lSlice + 1) Mod LNG_SYNC_POINTS) * lSegments
    If lLane = lRefLane Then
        lM = lM + lIndex
    End If
    If lN = 0 Then
        lM = lSlice * lSegments
        lS = 0
        If lSlice = 0 Or lLane = lRefLane Then
            lM = lM + lIndex
        End If
    End If
    If lIndex = 0 Or lLane = lRefLane Then
        lM = lM - 1
    End If
    '--- phi
    #If HasOperators Then
        lP = lRandom And LNG_UINT_MAX
        lP = (lP * lP) >> 32
        lP = (lP * lM) >> 32
    #Else
        lP = RShift64(UMul64(lRandom, lRandom), 32)
        lP = RShift64(UMul64(lP, lM), 32)
    #End If
    pvIndexAlpha = lRefLane * lLanes + CLng((lS + lM - lP - 1) Mod lLanes)
End Function

Private Sub pvExtendedHash(baInput() As Byte, baOutput() As Byte)
    Dim baTemp()        As Byte
    Dim uHash           As CryptoBlake2bContext
    Dim lOutSize        As Long
    Dim lOutPos         As Long
    Dim lRemaining      As Long
        
    lOutSize = UBound(baOutput) + 1
    ReDim baTemp(0 To 3) As Byte
    Call CopyMemory(baTemp(0), lOutSize, 4)
    If lOutSize < LNG_HASHSZ Then
        CryptoBlake2bInit uHash, lOutSize * 8
        CryptoBlake2bUpdate uHash, baTemp
        CryptoBlake2bUpdate uHash, baInput
        CryptoBlake2bFinalize uHash, baOutput
    Else
        CryptoBlake2bInit uHash, LNG_HASHSZ * 8
        CryptoBlake2bUpdate uHash, baTemp
        CryptoBlake2bUpdate uHash, baInput
        CryptoBlake2bFinalize uHash, baTemp
        Call CopyMemory(baOutput(0), baTemp(0), UBound(baTemp) + 1)
        lOutPos = LNG_HASHSZ \ 2
        lRemaining = lOutSize - LNG_HASHSZ \ 2
        Do While lRemaining > LNG_HASHSZ
            CryptoBlake2bInit uHash, LNG_HASHSZ * 8
            CryptoBlake2bUpdate uHash, baTemp
            CryptoBlake2bFinalize uHash, baTemp
            Call CopyMemory(baOutput(lOutPos), baTemp(0), LNG_HASHSZ \ 2)
            lOutPos = lOutPos + LNG_HASHSZ \ 2
            lRemaining = lRemaining - LNG_HASHSZ \ 2
        Loop
        CryptoBlake2bInit uHash, lRemaining * 8
        CryptoBlake2bUpdate uHash, baTemp
        CryptoBlake2bFinalize uHash, baTemp
        Call CopyMemory(baOutput(lOutPos), baTemp(0), lRemaining)
    End If
End Sub

Private Sub pvInitHash(baPassword() As Byte, baSalt() As Byte, baSecret() As Byte, baData() As Byte, ByVal lPasses As Long, ByVal lMemory As Long, ByVal lThreads As Long, ByVal lKeySize As Long, ByVal eMode As Argon2ModeEnum, baOutput() As Byte)
    Dim uHash           As CryptoBlake2bContext
    Dim baTemp(0 To 3)      As Byte
    
    CryptoBlake2bInit uHash, 512
    Call CopyMemory(baTemp(0), lThreads, 4)
    CryptoBlake2bUpdate uHash, baTemp
    Call CopyMemory(baTemp(0), lKeySize, 4)
    CryptoBlake2bUpdate uHash, baTemp
    Call CopyMemory(baTemp(0), lMemory, 4)
    CryptoBlake2bUpdate uHash, baTemp
    Call CopyMemory(baTemp(0), lPasses, 4)
    CryptoBlake2bUpdate uHash, baTemp
    Call CopyMemory(baTemp(0), LNG_VERSION, 4)
    CryptoBlake2bUpdate uHash, baTemp
    Call CopyMemory(baTemp(0), eMode, 4)
    CryptoBlake2bUpdate uHash, baTemp
    Call CopyMemory(baTemp(0), UBound(baPassword) + 1, 4)
    CryptoBlake2bUpdate uHash, baTemp
    CryptoBlake2bUpdate uHash, baPassword
    Call CopyMemory(baTemp(0), UBound(baSalt) + 1, 4)
    CryptoBlake2bUpdate uHash, baTemp
    CryptoBlake2bUpdate uHash, baSalt
    Call CopyMemory(baTemp(0), UBound(baSecret) + 1, 4)
    CryptoBlake2bUpdate uHash, baTemp
    CryptoBlake2bUpdate uHash, baSecret
    Call CopyMemory(baTemp(0), UBound(baData) + 1, 4)
    CryptoBlake2bUpdate uHash, baTemp
    CryptoBlake2bUpdate uHash, baData
    CryptoBlake2bFinalize uHash, baOutput
    #If DebugMode Then
        Debug.Print "Pre-hash digest: " & ToHex(baOutput)
    #End If
    ReDim Preserve baOutput(0 To UBound(baOutput) + 8) As Byte '--- 8 bytes overallocated
End Sub

Private Sub pvInitBlocks(baHash() As Byte, ByVal lMemory As Long, ByVal lThreads As Long, uOutput() As ArrayLong128)
    Static baTemp(0 To LNG_BLOCKSZ - 1) As Byte
    Dim lLane           As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lKdx            As Long
    
    Debug.Assert UBound(baHash) + 1 >= LNG_HASHSZ + 8
    ReDim uOutput(0 To lMemory - 1) As ArrayLong128
    For lLane = 0 To lThreads - 1
        lJdx = lLane * (lMemory \ lThreads)
        Call CopyMemory(baHash(LNG_HASHSZ + 4), lLane, 4)
        For lKdx = 0 To 1
            Call CopyMemory(baHash(LNG_HASHSZ), lKdx, 4)
            pvExtendedHash baHash, baTemp
            With uOutput(lJdx + lKdx)
                #If HasLongLong Then
                    Call CopyMemory(.Item(0), baTemp(0), LNG_BLOCKSZ)
                #Else
                    Static B2(0 To LNG_ARRAYSZ - 1) As Currency
                    Dim lTemp As Variant
                    Call CopyMemory(B2(0), baTemp(0), LNG_BLOCKSZ)
                    lTemp = LNG_ZERO
                    For lIdx = 0 To LNG_ARRAYSZ - 1
                        Call CopyMemory(ByVal VarPtr(lTemp) + 8, B2(lIdx), 8)
                        .Item(lIdx) = lTemp
                    Next
                #End If
            End With
        Next
    Next
End Sub

#If DebugMode Then
Private Function Hex64(vValue As Variant) As String
    Static S(0 To 1) As Long
    Call CopyMemory(S(0), ByVal VarPtr(vValue) + 8, 8)
    Hex64 = Right$("0000000" & Hex(S(1)), 8) & Right$("0000000" & Hex$(S(0)), 8)
End Function
#End If

Private Sub pvProcessBlocks(uBlocks() As ArrayLong128, ByVal lPasses As Long, ByVal lMemory As Long, ByVal lThreads As Long, ByVal eMode As Argon2ModeEnum)
#If HasLongLong Then
    Dim lRandom         As LongLong
#Else
    Dim lRandom         As Variant
#End If
    Dim lLanes          As Long
    Dim lSegments       As Long
    Dim lN              As Long
    Dim lSlice          As Long
    Dim lLane           As Long
    Dim uAddresses      As ArrayLong128
    Dim uInput          As ArrayLong128
    Dim uZero           As ArrayLong128
    Dim lIndex          As Long
    Dim lOffset         As Long
    Dim lPrev           As Long
    Dim lNewOffset      As Long
    
    lLanes = lMemory \ lThreads
    lSegments = lLanes \ LNG_SYNC_POINTS
    For lN = 0 To lPasses - 1
        For lSlice = 0 To LNG_SYNC_POINTS - 1
            For lLane = 0 To lThreads - 1
                If eMode = LNG_MODE_Argon2i Or eMode = LNG_MODE_Argon2id And lN = 0 And lSlice < LNG_SYNC_POINTS \ 2 Then
                    uInput.Item(0) = LNG_ZERO + lN
                    uInput.Item(1) = LNG_ZERO + lLane
                    uInput.Item(2) = LNG_ZERO + lSlice
                    uInput.Item(3) = LNG_ZERO + lMemory
                    uInput.Item(4) = LNG_ZERO + lPasses
                    uInput.Item(5) = LNG_ZERO + eMode
                End If
                uInput.Item(6) = LNG_ZERO
                If lN = 0 And lSlice = 0 Then
                    '--- first 2 blocks already generated
                    lIndex = 2
                    If eMode = LNG_MODE_Argon2i Or eMode = LNG_MODE_Argon2id Then
                        uInput.Item(6) = uInput.Item(6) + 1
                        pvFillBlock uInput, uZero, uAddresses
                        pvFillBlock uAddresses, uZero, uAddresses
                    End If
                Else
                    lIndex = 0
                End If
                lOffset = lLane * lLanes + lSlice * lSegments + lIndex
                Do While lIndex < lSegments
                    lPrev = lOffset - 1
                    If lIndex = 0 And lSlice = 0 Then
                        lPrev = lPrev + lLanes
                    End If
                    If eMode = LNG_MODE_Argon2i Or eMode = LNG_MODE_Argon2id And lN = 0 And lSlice < LNG_SYNC_POINTS \ 2 Then
                        If lIndex Mod LNG_ARRAYSZ = 0 Then
                            uInput.Item(6) = uInput.Item(6) + 1
                            pvFillBlock uInput, uZero, uAddresses
                            pvFillBlock uAddresses, uZero, uAddresses
                        End If
                        lRandom = uAddresses.Item(lIndex Mod LNG_ARRAYSZ)
                    Else
                        lRandom = uBlocks(lPrev).Item(0)
                    End If
                    lNewOffset = pvIndexAlpha(lRandom, lLanes, lSegments, lThreads, lN, lSlice, lLane, lIndex)
                    pvFillBlock uBlocks(lPrev), uBlocks(lNewOffset), uBlocks(lOffset), IsXor:=(lN > 0)
                    lIndex = lIndex + 1
                    lOffset = lOffset + 1
                Loop
            Next
        Next
        #If DebugMode Then
            Debug.Print "After pass " & lN
            For lIndex = 0 To 7
                Debug.Print "Block " & lIndex & ": " & LCase$(Hex64(uBlocks(lIndex).Item(0)))
            Next
            For lIndex = lMemory - 16 To lMemory - 1
                Debug.Print "Block " & lIndex & ": " & LCase$(Hex64(uBlocks(lIndex).Item(0)))
            Next
        #End If
    Next
End Sub

Private Sub pvExtractKey(uBlocks() As ArrayLong128, ByVal lMemory As Long, ByVal lThreads As Long, ByVal lOutSize As Long, baOutput() As Byte)
    Static baTemp(0 To LNG_BLOCKSZ - 1) As Byte
    Dim lLanes          As Long
    Dim lLane           As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    If lOutSize = 0 Then
        baOutput = vbNullString
        Exit Sub
    End If
    lLanes = lMemory \ lThreads
    With uBlocks(lMemory - 1)
        For lLane = 0 To lThreads - 2
            lJdx = lLane * lLanes + lLanes - 1
            For lIdx = 0 To LNG_ARRAYSZ - 1
                .Item(lIdx) = .Item(lIdx) Xor uBlocks(lJdx).Item(lIdx)
            Next
        Next
        #If HasLongLong Then
            Call CopyMemory(baTemp(0), .Item(0), LNG_BLOCKSZ)
        #Else
            For lIdx = 0 To LNG_ARRAYSZ - 1
                Call CopyMemory(baTemp(lIdx * 8), ByVal VarPtr(.Item(lIdx)) + 8, 8)
            Next
        #End If
        ReDim baOutput(0 To lOutSize - 1) As Byte
        pvExtendedHash baTemp, baOutput
    End With
End Sub

Private Sub pvDeriveKey( _
            ByVal eMode As Argon2ModeEnum, baPassword() As Byte, baSalt() As Byte, _
            Secret As Variant, Data As Variant, _
            ByVal lPasses As Long, ByVal lMemory As Long, ByVal lThreads As Long, _
            ByVal lOutSize As Long, baOutput() As Byte)
    Dim lIdx            As Long
    Dim baHash()        As Byte
    Dim uBlocks()       As ArrayLong128
    Dim baSecret()      As Byte
    Dim baData()        As Byte
    
    #If Not HasOperators Then
        If LNG_POW2(0) = 0 Then
            LNG_POW2(0) = CLngLng(1)
            For lIdx = 1 To 63
                LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
            Next
            LNG_SIGN_BIT = LNG_POW2(63)
            LNG_UINT_MAX = LNG_POW2(32) - 1
            LNG_ZERO = CLngLng(0)
        End If
    #End If
    If lPasses < 1 Then
        Err.Raise vbObjectError, , "Invalid CPU cost for Argon2 (" & lPasses & ")"
    End If
    If lThreads < 1 Then
        Err.Raise vbObjectError, , "Invalid parallelism for Argon2 (" & lThreads & ")"
    End If
    If lOutSize <= 0 Then
        lOutSize = 32
    End If
    If IsMissing(Secret) Then
        baSecret = vbNullString
    ElseIf IsArray(Secret) Then
        baSecret = Secret
    Else
        baSecret = ToUtf8Array(CStr(Secret))
    End If
    If IsMissing(Data) Then
        baData = vbNullString
    ElseIf IsArray(Data) Then
        baData = Data
    Else
        baData = ToUtf8Array(CStr(Data))
    End If
    pvInitHash baPassword, baSalt, baSecret, baData, lPasses, lMemory, lThreads, lOutSize, eMode, baHash
    lMemory = (lMemory \ (LNG_SYNC_POINTS * lThreads)) * (LNG_SYNC_POINTS * lThreads)
    If lMemory < 2 * LNG_SYNC_POINTS * lThreads Then
        lMemory = 2 * LNG_SYNC_POINTS * lThreads
    End If
    pvInitBlocks baHash, lMemory, lThreads, uBlocks
    pvProcessBlocks uBlocks, lPasses, lMemory, lThreads, eMode
    pvExtractKey uBlocks, lMemory, lThreads, lOutSize, baOutput
End Sub

Public Function CryptoArgon2KdfByteArray(baPass() As Byte, baSalt() As Byte, _
            Optional ByVal OutSize As Long, _
            Optional Secret As Variant, _
            Optional Data As Variant, _
            Optional ByVal Passes As Long = 1, _
            Optional ByVal Memory As Long = 64& * 1024, _
            Optional ByVal Parallelism As Long = 4) As Byte()
    pvDeriveKey LNG_MODE_Argon2i, baPass, baSalt, Secret, Data, Passes, Memory, Parallelism, OutSize, CryptoArgon2KdfByteArray
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

Public Function CryptoArgon2KdfText(sPass As String, sSalt As String, _
            Optional ByVal OutSize As Long, _
            Optional Secret As Variant, _
            Optional Data As Variant, _
            Optional ByVal Passes As Long = 1, _
            Optional ByVal Memory As Long = 64& * 1024, _
            Optional ByVal Parallelism As Long = 4) As String
    CryptoArgon2KdfText = ToHex(CryptoArgon2KdfByteArray(ToUtf8Array(sPass), ToUtf8Array(sSalt), OutSize:=OutSize, Secret:=Secret, Data:=Data, Passes:=Passes, Memory:=Memory, Parallelism:=Parallelism))
End Function

Public Function CryptoArgon2IdKdfByteArray(baPass() As Byte, baSalt() As Byte, _
            Optional ByVal OutSize As Long, _
            Optional Secret As Variant, _
            Optional Data As Variant, _
            Optional ByVal Passes As Long = 1, _
            Optional ByVal Memory As Long = 64& * 1024, _
            Optional ByVal Parallelism As Long = 4) As Byte()
    pvDeriveKey LNG_MODE_Argon2id, baPass, baSalt, Secret, Data, Passes, Memory, Parallelism, OutSize, CryptoArgon2IdKdfByteArray
End Function

Public Function CryptoArgon2IdKdfText(sPass As String, sSalt As String, _
            Optional ByVal OutSize As Long, _
            Optional Secret As Variant, _
            Optional Data As Variant, _
            Optional ByVal Passes As Long = 1, _
            Optional ByVal Memory As Long = 64& * 1024, _
            Optional ByVal Parallelism As Long = 4) As String
    CryptoArgon2IdKdfText = ToHex(CryptoArgon2IdKdfByteArray(ToUtf8Array(sPass), ToUtf8Array(sSalt), OutSize:=OutSize, Secret:=Secret, Data:=Data, Passes:=Passes, Memory:=Memory, Parallelism:=Parallelism))
End Function

Private Sub pvTestVectors(ByVal SeqNo As Long, ByVal Mode As Argon2ModeEnum, ByVal Passes As Long, ByVal Memory As Long, ByVal Parallelism As Long, HASH As String)
    Dim baEmpty() As Byte
    Dim baOutput() As Byte
    
    baEmpty = vbNullString
    pvDeriveKey Mode, ToUtf8Array("password"), ToUtf8Array("somesalt"), baEmpty, baEmpty, Passes, Memory, Parallelism, Len(HASH) \ 2, baOutput
    If ToHex(baOutput) <> HASH Then
        Debug.Print "Test " & SeqNo & " - got: " & ToHex(baOutput) & " want: " & HASH
    End If
End Sub

Public Sub CryptoTestArgon2()
    pvTestVectors 1, LNG_MODE_Argon2i, Passes:=1, Memory:=64, Parallelism:=1, HASH:="b9c401d1844a67d50eae3967dc28870b22e508092e861a37"
    pvTestVectors 2, LNG_MODE_Argon2d, Passes:=1, Memory:=64, Parallelism:=1, HASH:="8727405fd07c32c78d64f547f24150d3f2e703a89f981a19"
    pvTestVectors 3, LNG_MODE_Argon2id, Passes:=1, Memory:=64, Parallelism:=1, HASH:="655ad15eac652dc59f7170a7332bf49b8469be1fdb9c28bb"
    pvTestVectors 4, LNG_MODE_Argon2i, Passes:=2, Memory:=64, Parallelism:=1, HASH:="8cf3d8f76a6617afe35fac48eb0b7433a9a670ca4a07ed64"
    pvTestVectors 5, LNG_MODE_Argon2d, Passes:=2, Memory:=64, Parallelism:=1, HASH:="3be9ec79a69b75d3752acb59a1fbb8b295a46529c48fbb75"
    pvTestVectors 6, LNG_MODE_Argon2id, Passes:=2, Memory:=64, Parallelism:=1, HASH:="068d62b26455936aa6ebe60060b0a65870dbfa3ddf8d41f7"
    pvTestVectors 7, LNG_MODE_Argon2i, Passes:=2, Memory:=64, Parallelism:=2, HASH:="2089f3e78a799720f80af806553128f29b132cafe40d059f"
    pvTestVectors 8, LNG_MODE_Argon2d, Passes:=2, Memory:=64, Parallelism:=2, HASH:="68e2462c98b8bc6bb60ec68db418ae2c9ed24fc6748a40e9"
    pvTestVectors 9, LNG_MODE_Argon2id, Passes:=2, Memory:=64, Parallelism:=2, HASH:="350ac37222f436ccb5c0972f1ebd3bf6b958bf2071841362"
    pvTestVectors 10, LNG_MODE_Argon2i, Passes:=3, Memory:=256, Parallelism:=2, HASH:="f5bbf5d4c3836af13193053155b73ec7476a6a2eb93fd5e6"
    pvTestVectors 11, LNG_MODE_Argon2d, Passes:=3, Memory:=256, Parallelism:=2, HASH:="f4f0669218eaf3641f39cc97efb915721102f4b128211ef2"
    pvTestVectors 12, LNG_MODE_Argon2id, Passes:=3, Memory:=256, Parallelism:=2, HASH:="4668d30ac4187e6878eedeacf0fd83c5a0a30db2cc16ef0b"
    pvTestVectors 13, LNG_MODE_Argon2i, Passes:=4, Memory:=4096, Parallelism:=4, HASH:="a11f7b7f3f93f02ad4bddb59ab62d121e278369288a0d0e7"
    pvTestVectors 14, LNG_MODE_Argon2d, Passes:=4, Memory:=4096, Parallelism:=4, HASH:="935598181aa8dc2b720914aa6435ac8d3e3a4210c5b0fb2d"
    pvTestVectors 15, LNG_MODE_Argon2id, Passes:=4, Memory:=4096, Parallelism:=4, HASH:="145db9733a9f4ee43edf33c509be96b934d505a4efb33c5a"
    pvTestVectors 16, LNG_MODE_Argon2i, Passes:=4, Memory:=1024, Parallelism:=8, HASH:="0cdd3956aa35e6b475a7b0c63488822f774f15b43f6e6e17"
    pvTestVectors 17, LNG_MODE_Argon2d, Passes:=4, Memory:=1024, Parallelism:=8, HASH:="83604fc2ad0589b9d055578f4d3cc55bc616df3578a896e9"
    pvTestVectors 18, LNG_MODE_Argon2id, Passes:=4, Memory:=1024, Parallelism:=8, HASH:="8dafa8e004f8ea96bf7c0f93eecf67a6047476143d15577f"
    pvTestVectors 19, LNG_MODE_Argon2i, Passes:=2, Memory:=64, Parallelism:=3, HASH:="5cab452fe6b8479c8661def8cd703b611a3905a6d5477fe6"
    pvTestVectors 20, LNG_MODE_Argon2d, Passes:=2, Memory:=64, Parallelism:=3, HASH:="22474a423bda2ccd36ec9afd5119e5c8949798cadf659f51"
    pvTestVectors 21, LNG_MODE_Argon2id, Passes:=2, Memory:=64, Parallelism:=3, HASH:="4a15b31aec7c2590b87d1f520be7d96f56658172deaa3079"
    pvTestVectors 22, LNG_MODE_Argon2i, Passes:=3, Memory:=1024, Parallelism:=6, HASH:="d236b29c2b2a09babee842b0dec6aa1e83ccbdea8023dced"
    pvTestVectors 23, LNG_MODE_Argon2d, Passes:=3, Memory:=1024, Parallelism:=6, HASH:="a3351b0319a53229152023d9206902f4ef59661cdca89481"
    pvTestVectors 24, LNG_MODE_Argon2id, Passes:=3, Memory:=1024, Parallelism:=6, HASH:="1640b932f4b60e272f5d2207b9a9c626ffa1bd88d2349016"
End Sub
