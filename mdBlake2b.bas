Attribute VB_Name = "mdBlake2b"
'--- mdBlake2b.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
Private Declare PtrSafe Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_BLOCKSZ               As Long = 128
Private Const LNG_ROUNDS                As Long = 12

Public Type CryptoBlake2bContext
#If HasPtrSafe Then
    H0                  As LongLong
    H1                  As LongLong
    H2                  As LongLong
    H3                  As LongLong
    H4                  As LongLong
    H5                  As LongLong
    H6                  As LongLong
    H7                  As LongLong
#Else
    H0                  As Variant
    H1                  As Variant
    H2                  As Variant
    H3                  As Variant
    H4                  As Variant
    H5                  As Variant
    H6                  As Variant
    H7                  As Variant
#End If
    Partial(0 To LNG_BLOCKSZ - 1) As Byte
    NPartial            As Long
    NInput              As Currency
    OutSize             As Long
End Type

#If HasPtrSafe Then
Private LNG_ZERO                    As LongLong
Private LNG_IV(0 To 7)              As LongLong
#Else
Private LNG_ZERO                    As Variant
Private LNG_IV(0 To 7)              As Variant
#End If
Private LNG_SIGMA(0 To 15, 0 To LNG_ROUNDS - 1)  As Long

#If Not HasOperators Then
#If HasPtrSafe Then
Private LNG_POW2(0 To 63)           As LongLong
Private LNG_SIGN_BIT                As LongLong ' 2 ^ 63
#Else
Private LNG_POW2(0 To 63)           As Variant
Private LNG_SIGN_BIT                As Variant
#End If

#If HasPtrSafe Then
Private Function RotR64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RotR64(lX As Variant, ByVal lN As Long) As Variant
#End If
    '--- RotR64 = RShift64(X, n) Or LShift64(X, 64 - n)
    Debug.Assert lN <> 0
    RotR64 = ((lX And (-1 Xor LNG_SIGN_BIT)) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(63 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(64 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * LNG_SIGN_BIT)
End Function

#If HasPtrSafe Then
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

#If HasPtrSafe Then
Private Sub pvQuarter64(lA As LongLong, lB As LongLong, lC As LongLong, lD As LongLong, ByVal lX As LongLong, ByVal lY As LongLong)
#Else
Private Sub pvQuarter64(lA As Variant, lB As Variant, lC As Variant, lD As Variant, ByVal lX As Variant, ByVal lY As Variant)
#End If
    lA = UAdd64(UAdd64(lA, lB), lX)
    lD = RotR64(lD Xor lA, 32)
    lC = UAdd64(lC, lD)
    lB = RotR64(lB Xor lC, 24)
    lA = UAdd64(UAdd64(lA, lB), lY)
    lD = RotR64(lD Xor lA, 16)
    lC = UAdd64(lC, lD)
    lB = RotR64(lB Xor lC, 63)
End Sub
#Else
[ IntegerOverflowChecks (False) ]
Private Sub pvQuarter64(lA As LongLong, lB As LongLong, lC As LongLong, lD As LongLong, ByVal lX As LongLong, ByVal lY As LongLong)
    lA = lA + lB + lX
    lD = (lD Xor lA) >> 32 Or (lD Xor lA) << 32
    lC = lC + lD
    lB = (lB Xor lC) >> 24 or (lB Xor lC) << 40
    lA = lA + lB + lY
    lD = (lD Xor lA) >> 16 or (lD Xor lA) << 48
    lC = lC + lD
    lB = (lB Xor lC) >> 63 or (lB Xor lC) << 1
End Sub
#End If

#If Not HasPtrSafe Then
Private Function CLngLng(vValue As Variant) As Variant
    Const VT_I8 As Long = &H14
    Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
End Function
#End If

Private Sub pvCompress(uCtx As CryptoBlake2bContext, Optional ByVal IsLast As Boolean)
#If HasPtrSafe Then
    Static B(0 To 15)   As LongLong
    Dim V0              As LongLong
    Dim V1              As LongLong
    Dim V2              As LongLong
    Dim V3              As LongLong
    Dim V4              As LongLong
    Dim V5              As LongLong
    Dim V6              As LongLong
    Dim V7              As LongLong
    Dim V8              As LongLong
    Dim V9              As LongLong
    Dim V10             As LongLong
    Dim V11             As LongLong
    Dim V12             As LongLong
    Dim V13             As LongLong
    Dim V14             As LongLong
    Dim V15             As LongLong
    Dim S0              As LongLong
#Else
    Static B(0 To 15)   As Variant
    Dim V0              As Variant
    Dim V1              As Variant
    Dim V2              As Variant
    Dim V3              As Variant
    Dim V4              As Variant
    Dim V5              As Variant
    Dim V6              As Variant
    Dim V7              As Variant
    Dim V8              As Variant
    Dim V9              As Variant
    Dim V10             As Variant
    Dim V11             As Variant
    Dim V12             As Variant
    Dim V13             As Variant
    Dim V14             As Variant
    Dim V15             As Variant
    Dim S0              As Variant
#End If
    Dim cTemp           As Currency
    Dim lIdx            As Long

    With uCtx
        If .NPartial < LNG_BLOCKSZ Then
            Call FillMemory(.Partial(.NPartial), LNG_BLOCKSZ - .NPartial, 0)
        End If
        #If HasPtrSafe Then
            Call CopyMemory(B(0), .Partial(0), LNG_BLOCKSZ)
        #Else
            For lIdx = 0 To UBound(B)
                B(lIdx) = LNG_ZERO
                Call CopyMemory(ByVal VarPtr(B(lIdx)) + 8, .Partial(8 * lIdx), 8)
            Next
        #End If
        V0 = .H0: V1 = .H1
        V2 = .H2: V3 = .H3
        V4 = .H4: V5 = .H5
        V6 = .H6: V7 = .H7
        V8 = LNG_IV(0): V9 = LNG_IV(1)
        V10 = LNG_IV(2): V11 = LNG_IV(3)
        V12 = LNG_IV(4): V13 = LNG_IV(5)
        V14 = LNG_IV(6): V15 = LNG_IV(7)
        .NInput = .NInput + .NPartial
        .NPartial = 0
        cTemp = .NInput / 10000@
        #If HasPtrSafe Then
            Call CopyMemory(S0, cTemp, 8)
        #Else
            S0 = LNG_ZERO
            Call CopyMemory(ByVal VarPtr(S0) + 8, cTemp, 8)
        #End If
        V12 = V12 Xor S0
        If IsLast Then
            V14 = Not V14
        End If
        For lIdx = 0 To LNG_ROUNDS - 1
            pvQuarter64 V0, V4, V8, V12, B(LNG_SIGMA(0, lIdx)), B(LNG_SIGMA(1, lIdx))
            pvQuarter64 V1, V5, V9, V13, B(LNG_SIGMA(2, lIdx)), B(LNG_SIGMA(3, lIdx))
            pvQuarter64 V2, V6, V10, V14, B(LNG_SIGMA(4, lIdx)), B(LNG_SIGMA(5, lIdx))
            pvQuarter64 V3, V7, V11, V15, B(LNG_SIGMA(6, lIdx)), B(LNG_SIGMA(7, lIdx))
            pvQuarter64 V0, V5, V10, V15, B(LNG_SIGMA(8, lIdx)), B(LNG_SIGMA(9, lIdx))
            pvQuarter64 V1, V6, V11, V12, B(LNG_SIGMA(10, lIdx)), B(LNG_SIGMA(11, lIdx))
            pvQuarter64 V2, V7, V8, V13, B(LNG_SIGMA(12, lIdx)), B(LNG_SIGMA(13, lIdx))
            pvQuarter64 V3, V4, V9, V14, B(LNG_SIGMA(14, lIdx)), B(LNG_SIGMA(15, lIdx))
        Next
        .H0 = .H0 Xor V0 Xor V8
        .H1 = .H1 Xor V1 Xor V9
        .H2 = .H2 Xor V2 Xor V10
        .H3 = .H3 Xor V3 Xor V11
        .H4 = .H4 Xor V4 Xor V12
        .H5 = .H5 Xor V5 Xor V13
        .H6 = .H6 Xor V6 Xor V14
        .H7 = .H7 Xor V7 Xor V15
    End With
End Sub

Public Sub CryptoBlake2bInit(uCtx As CryptoBlake2bContext, ByVal lBitSize As Long, Optional Key As Variant)
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim baKey()         As Byte
    Dim lKeySize        As Long
    
    If LNG_IV(0) = 0 Then
        LNG_ZERO = CLngLng(0)
        For Each vElem In Split("6A09E667F3BCC908 BB67AE8584CAA73B 3C6EF372FE94F82B A54FF53A5F1D36F1 510E527FADE682D1 9B05688C2B3E6C1F 1F83D9ABFB41BD6B 5BE0CD19137E2179")
            LNG_IV(lIdx) = CLngLng(CStr("&H" & vElem))
            lIdx = lIdx + 1
        Next
        lIdx = 0
        For Each vElem In Split("0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 " & _
                                "14 10 4 8 9 15 13 6 1 12 0 2 11 7 5 3 " & _
                                "11 8 12 0 5 2 15 13 10 14 3 6 7 1 9 4 " & _
                                "7 9 3 1 13 12 11 14 2 6 5 10 4 0 15 8 " & _
                                "9 0 5 7 2 4 10 15 14 1 11 12 6 8 3 13 " & _
                                "2 12 6 10 0 11 8 3 4 13 7 5 15 14 1 9 " & _
                                "12 5 1 15 14 13 4 10 0 7 6 3 9 2 8 11 " & _
                                "13 11 7 14 12 1 3 9 5 0 15 4 8 6 2 10 " & _
                                "6 15 14 9 11 3 0 8 12 2 13 7 1 4 10 5 " & _
                                "10 2 8 4 7 6 1 5 15 11 9 14 3 12 13 0")
            LNG_SIGMA(lIdx And 15, lIdx \ 16) = vElem
            lIdx = lIdx + 1
        Next
        '--- copy rows 10 & 11 from rows 0 & 1
        Call CopyMemory(LNG_SIGMA(0, 10), LNG_SIGMA(0, 0), 2 * 64)
        #If Not HasOperators Then
            LNG_POW2(0) = CLngLng(1)
            For lIdx = 1 To 63
                LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
            Next
            LNG_SIGN_BIT = LNG_POW2(63)
        #End If
    End If
    If lBitSize <= 0 Or lBitSize > 512 Or (lBitSize And 7) <> 0 Then
        Err.Raise vbObjectError, , "Invalid bit-size for BLAKE2b (" & lBitSize & ")"
    End If
    If Not IsMissing(Key) Then
        If IsArray(Key) Then
            baKey = Key
        Else
            baKey = ToUtf8Array(CStr(Key))
        End If
        lKeySize = UBound(baKey) + 1
    End If
    If lKeySize > 64 Then
        Err.Raise vbObjectError, , "Key for BLAKE2b-MAC must be up to 64 bytes (" & lKeySize & ")"
    End If
    With uCtx
        #If HasPtrSafe Then
            Call CopyMemory(.H0, LNG_IV(0), 8 * 8)
        #Else
            Call CopyMemory(.H0, LNG_IV(0), 8 * 16)
        #End If
        .OutSize = lBitSize \ 8
        .H0 = .H0 Xor &H1010000 Xor (lKeySize * &H100) Xor .OutSize
        .NPartial = 0
        .NInput = 0
        If lKeySize > 0 Then
            Call CopyMemory(.Partial(0), baKey(0), lKeySize)
            Call FillMemory(.Partial(lKeySize), LNG_BLOCKSZ - lKeySize, 0)
            .NPartial = LNG_BLOCKSZ
        End If
    End With
End Sub

Public Sub CryptoBlake2bUpdate(uCtx As CryptoBlake2bContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lIdx            As Long
    
    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        If .NPartial > 0 And .NPartial < LNG_BLOCKSZ And Size > 0 Then
            lIdx = LNG_BLOCKSZ - .NPartial
            If lIdx > Size Then
                lIdx = Size
            End If
            Call CopyMemory(.Partial(.NPartial), baInput(Pos), lIdx)
            .NPartial = .NPartial + lIdx
            Pos = Pos + lIdx
            Size = Size - lIdx
        End If
        Do While Size > 0
            If .NPartial <> 0 Then
                '--- do nothing
            ElseIf Size >= LNG_BLOCKSZ Then
                Call CopyMemory(.Partial(0), baInput(Pos), LNG_BLOCKSZ)
                .NPartial = LNG_BLOCKSZ
                Pos = Pos + LNG_BLOCKSZ
                Size = Size - LNG_BLOCKSZ
            Else
                Call CopyMemory(.Partial(0), baInput(Pos), Size)
                .NPartial = Size
                Exit Do
            End If
            pvCompress uCtx
        Loop
    End With
End Sub

Public Sub CryptoBlake2bFinalize(uCtx As CryptoBlake2bContext, baOutput() As Byte)
    With uCtx
        pvCompress uCtx, IsLast:=True
        ReDim baOutput(0 To .OutSize - 1) As Byte
        #If HasPtrSafe Then
            Call CopyMemory(baOutput(0), .H0, .OutSize)
        #Else
            Call CopyMemory(.Partial(0), ByVal VarPtr(.H0) + 8, 8)
            Call CopyMemory(.Partial(8), ByVal VarPtr(.H1) + 8, 8)
            Call CopyMemory(.Partial(16), ByVal VarPtr(.H2) + 8, 8)
            Call CopyMemory(.Partial(24), ByVal VarPtr(.H3) + 8, 8)
            If .OutSize > 32 Then
                Call CopyMemory(.Partial(32), ByVal VarPtr(.H4) + 8, 8)
                Call CopyMemory(.Partial(40), ByVal VarPtr(.H5) + 8, 8)
                Call CopyMemory(.Partial(48), ByVal VarPtr(.H6) + 8, 8)
                Call CopyMemory(.Partial(56), ByVal VarPtr(.H7) + 8, 8)
            End If
            Call CopyMemory(baOutput(0), .Partial(0), .OutSize)
        #End If
    End With
    Call FillMemory(uCtx, LenB(uCtx), 0)
End Sub

Public Function CryptoBlake2bByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional Key As Variant) As Byte()
    Dim uCtx            As CryptoBlake2bContext
    
    CryptoBlake2bInit uCtx, lBitSize, Key:=Key
    CryptoBlake2bUpdate uCtx, baInput, Pos, Size
    CryptoBlake2bFinalize uCtx, CryptoBlake2bByteArray
End Function

Private Function ToUtf8Array(sText As String) As Byte()
    Const CP_UTF8       As Long = 65001
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
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
        If Len(sByte) = 1 Then
            Mid$(ToHex, lIdx * 2 + 2, 1) = sByte
        Else
            Mid$(ToHex, lIdx * 2 + 1, 2) = sByte
        End If
    Next
End Function

Public Function CryptoBlake2bText(ByVal lBitSize As Long, sText As String, Optional Key As Variant) As String
    CryptoBlake2bText = ToHex(CryptoBlake2bByteArray(lBitSize, ToUtf8Array(sText), Key:=Key))
End Function
