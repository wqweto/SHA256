Attribute VB_Name = "mdBlake2s"
'--- mdBlake2s.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_BLOCKSZ               As Long = 64
Private Const LNG_ROUNDS                As Long = 10

Public Type CryptoBlake2sContext
    H0                  As Long
    H1                  As Long
    H2                  As Long
    H3                  As Long
    H4                  As Long
    H5                  As Long
    H6                  As Long
    H7                  As Long
    Partial(0 To LNG_BLOCKSZ - 1) As Byte
    NPartial            As Long
    NInput              As Currency
    OutSize             As Long
End Type

Private LNG_IV(0 To 7)              As Long
Private LNG_SIGMA(0 To 15, 0 To LNG_ROUNDS - 1)  As Long

#If Not HasOperators Then
Private LNG_POW2(0 To 31)           As Long

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

Private Sub pvQuarter32(lA As Long, lB As Long, lC As Long, lD As Long, ByVal lX As Long, ByVal lY As Long)
    lA = UAdd32(UAdd32(lA, lB), lX)
    lD = RotR32(lD Xor lA, 16)
    lC = UAdd32(lC, lD)
    lB = RotR32(lB Xor lC, 12)
    lA = UAdd32(UAdd32(lA, lB), lY)
    lD = RotR32(lD Xor lA, 8)
    lC = UAdd32(lC, lD)
    lB = RotR32(lB Xor lC, 7)
End Sub
#Else
[ IntegerOverflowChecks (False) ]
Private Sub pvQuarter32(lA As Long, lB As Long, lC As Long, lD As Long, ByVal lX As Long, ByVal lY As Long)
    lA = lA + lB + lX
    lD = (lD Xor lA) >> 16 Or (lD Xor lA) << 16
    lC = lC + lD
    lB = (lB Xor lC) >> 12 or (lB Xor lC) << 20
    lA = lA + lB + lY
    lD = (lD Xor lA) >> 8 or (lD Xor lA) << 24
    lC = lC + lD
    lB = (lB Xor lC) >> 7 or (lB Xor lC) << 25
End Sub
#End If

Private Sub pvCompress(uCtx As CryptoBlake2sContext, Optional ByVal IsLast As Boolean)
    Static B(0 To 15)   As Long
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
    Dim cTemp           As Currency
    Dim lIdx            As Long

    With uCtx
        If .NPartial < LNG_BLOCKSZ Then
            Call FillMemory(.Partial(.NPartial), LNG_BLOCKSZ - .NPartial, 0)
        End If
        Call CopyMemory(B(0), .Partial(0), LNG_BLOCKSZ)
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
        Call CopyMemory(S(0), cTemp, 8)
        V12 = V12 Xor S(0)
        V13 = V13 Xor S(1)
        If IsLast Then
            V14 = Not V14
        End If
        For lIdx = 0 To LNG_ROUNDS - 1
            pvQuarter32 V0, V4, V8, V12, B(LNG_SIGMA(0, lIdx)), B(LNG_SIGMA(1, lIdx))
            pvQuarter32 V1, V5, V9, V13, B(LNG_SIGMA(2, lIdx)), B(LNG_SIGMA(3, lIdx))
            pvQuarter32 V2, V6, V10, V14, B(LNG_SIGMA(4, lIdx)), B(LNG_SIGMA(5, lIdx))
            pvQuarter32 V3, V7, V11, V15, B(LNG_SIGMA(6, lIdx)), B(LNG_SIGMA(7, lIdx))
            pvQuarter32 V0, V5, V10, V15, B(LNG_SIGMA(8, lIdx)), B(LNG_SIGMA(9, lIdx))
            pvQuarter32 V1, V6, V11, V12, B(LNG_SIGMA(10, lIdx)), B(LNG_SIGMA(11, lIdx))
            pvQuarter32 V2, V7, V8, V13, B(LNG_SIGMA(12, lIdx)), B(LNG_SIGMA(13, lIdx))
            pvQuarter32 V3, V4, V9, V14, B(LNG_SIGMA(14, lIdx)), B(LNG_SIGMA(15, lIdx))
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

Public Sub CryptoBlake2sInit(uCtx As CryptoBlake2sContext, ByVal lBitSize As Long, Optional Key As Variant)
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim baKey()         As Byte
    Dim lKeySize        As Long
    
    If LNG_IV(0) = 0 Then
        For Each vElem In Split("6A09E667 BB67AE85 3C6EF372 A54FF53A 510E527F 9B05688C 1F83D9AB 5BE0CD19")
            LNG_IV(lIdx) = "&H" & vElem
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
        #If Not HasOperators Then
            LNG_POW2(0) = 1
            For lIdx = 1 To 30
                LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
            Next
            LNG_POW2(31) = &H80000000
        #End If
    End If
    If lBitSize <= 0 Or lBitSize > 256 Or (lBitSize And 7) <> 0 Then
        Err.Raise vbObjectError, , "Invalid bit-size for BLAKE2s (" & lBitSize & ")"
    End If
    If Not IsMissing(Key) Then
        If IsArray(Key) Then
            baKey = Key
        Else
            baKey = ToUtf8Array(CStr(Key))
        End If
        lKeySize = UBound(baKey) + 1
    End If
    If lKeySize > 32 Then
        Err.Raise vbObjectError, , "Key for BLAKE2s-MAC must be up to 32 bytes (" & lKeySize & ")"
    End If
    With uCtx
        Call CopyMemory(.H0, LNG_IV(0), 8 * 4)
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

Public Sub CryptoBlake2sUpdate(uCtx As CryptoBlake2sContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
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
                '--- do nothng
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

Public Sub CryptoBlake2sFinalize(uCtx As CryptoBlake2sContext, baOutput() As Byte)
    With uCtx
        pvCompress uCtx, IsLast:=True
        ReDim baOutput(0 To .OutSize - 1) As Byte
        Call CopyMemory(baOutput(0), .H0, .OutSize)
    End With
    Call FillMemory(uCtx, LenB(uCtx), 0)
End Sub

Public Function CryptoBlake2sByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional Key As Variant) As Byte()
    Dim uCtx            As CryptoBlake2sContext
    
    CryptoBlake2sInit uCtx, lBitSize, Key:=Key
    CryptoBlake2sUpdate uCtx, baInput, Pos, Size
    CryptoBlake2sFinalize uCtx, CryptoBlake2sByteArray
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

Public Function CryptoBlake2sText(ByVal lBitSize As Long, sText As String, Optional Key As Variant) As String
    CryptoBlake2sText = ToHex(CryptoBlake2sByteArray(lBitSize, ToUtf8Array(sText), Key:=Key))
End Function
