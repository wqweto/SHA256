Attribute VB_Name = "mdScryptKdf"
'--- mdScryptKdf.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Type ArrayLong16
    Item(0 To 15)           As Long
End Type

#If Not HasOperators Then
Private LNG_POW2(0 To 31)           As Long

Private Function RotL32(ByVal lX As Long, ByVal lN As Long) As Long
    '--- RotL32 = LShift(X, n) Or RShift(X, 32 - n)
    Debug.Assert lN <> 0
    RotL32 = ((lX And (LNG_POW2(31 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(31 - lN)) <> 0) * LNG_POW2(31)) Or _
        ((lX And (LNG_POW2(31) Xor -1)) \ LNG_POW2(32 - lN) Or -(lX < 0) * LNG_POW2(lN - 1))
End Function

Private Function UAdd32(ByVal lX As Long, ByVal lY As Long) As Long
    If (lX Xor lY) >= 0 Then
        UAdd32 = ((lX Xor &H80000000) + lY) Xor &H80000000
    Else
        UAdd32 = lX + lY
    End If
End Function

Private Sub Op32(X As ArrayLong16, ByVal lIdx As Long, ByVal lA As Long, ByVal lB As Long, ByVal lShift As Long)
    X.Item(lIdx) = X.Item(lIdx) Xor RotL32(UAdd32(X.Item(lA), X.Item(lB)), lShift)
End Sub
#End If

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Private Sub pvSalsa20Core(B() As Byte)
    Dim B32             As ArrayLong16
    Dim X               As ArrayLong16
    Dim lIdx            As Long
    
    Debug.Assert UBound(B) + 1 >= 64
    Call CopyMemory(B32, B(0), 64)
    X = B32
    For lIdx = 0 To 3
        #If HasOperators Then
            Dim lTemp As Long
            With X
                lTemp = .Item(0) + .Item(12): .Item(4) = .Item(4) Xor (lTemp << 7 Or lTemp >> 25)
                lTemp = .Item(4) + .Item(0): .Item(8) = .Item(8) Xor (lTemp << 9 Or lTemp >> 23)
                lTemp = .Item(8) + .Item(4): .Item(12) = .Item(12) Xor (lTemp << 13 Or lTemp >> 19)
                lTemp = .Item(12) + .Item(8): .Item(0) = .Item(0) Xor (lTemp << 18 Or lTemp >> 14)
                lTemp = .Item(5) + .Item(1): .Item(9) = .Item(9) Xor (lTemp << 7 Or lTemp >> 25)
                lTemp = .Item(9) + .Item(5): .Item(13) = .Item(13) Xor (lTemp << 9 Or lTemp >> 23)
                lTemp = .Item(13) + .Item(9): .Item(1) = .Item(1) Xor (lTemp << 13 Or lTemp >> 19)
                lTemp = .Item(1) + .Item(13): .Item(5) = .Item(5) Xor (lTemp << 18 Or lTemp >> 14)
                lTemp = .Item(10) + .Item(6): .Item(14) = .Item(14) Xor (lTemp << 7 Or lTemp >> 25)
                lTemp = .Item(14) + .Item(10): .Item(2) = .Item(2) Xor (lTemp << 9 Or lTemp >> 23)
                lTemp = .Item(2) + .Item(14): .Item(6) = .Item(6) Xor (lTemp << 13 Or lTemp >> 19)
                lTemp = .Item(6) + .Item(2): .Item(10) = .Item(10) Xor (lTemp << 18 Or lTemp >> 14)
                lTemp = .Item(15) + .Item(11): .Item(3) = .Item(3) Xor (lTemp << 7 Or lTemp >> 25)
                lTemp = .Item(3) + .Item(15): .Item(7) = .Item(7) Xor (lTemp << 9 Or lTemp >> 23)
                lTemp = .Item(7) + .Item(3): .Item(11) = .Item(11) Xor (lTemp << 13 Or lTemp >> 19)
                lTemp = .Item(11) + .Item(7): .Item(15) = .Item(15) Xor (lTemp << 18 Or lTemp >> 14)
                lTemp = .Item(0) + .Item(3): .Item(1) = .Item(1) Xor (lTemp << 7 Or lTemp >> 25)
                lTemp = .Item(1) + .Item(0): .Item(2) = .Item(2) Xor (lTemp << 9 Or lTemp >> 23)
                lTemp = .Item(2) + .Item(1): .Item(3) = .Item(3) Xor (lTemp << 13 Or lTemp >> 19)
                lTemp = .Item(3) + .Item(2): .Item(0) = .Item(0) Xor (lTemp << 18 Or lTemp >> 14)
                lTemp = .Item(5) + .Item(4): .Item(6) = .Item(6) Xor (lTemp << 7 Or lTemp >> 25)
                lTemp = .Item(6) + .Item(5): .Item(7) = .Item(7) Xor (lTemp << 9 Or lTemp >> 23)
                lTemp = .Item(7) + .Item(6): .Item(4) = .Item(4) Xor (lTemp << 13 Or lTemp >> 19)
                lTemp = .Item(4) + .Item(7): .Item(5) = .Item(5) Xor (lTemp << 18 Or lTemp >> 14)
                lTemp = .Item(10) + .Item(9): .Item(11) = .Item(11) Xor (lTemp << 7 Or lTemp >> 25)
                lTemp = .Item(11) + .Item(10): .Item(8) = .Item(8) Xor (lTemp << 9 Or lTemp >> 23)
                lTemp = .Item(8) + .Item(11): .Item(9) = .Item(9) Xor (lTemp << 13 Or lTemp >> 19)
                lTemp = .Item(9) + .Item(8): .Item(10) = .Item(10) Xor (lTemp << 18 Or lTemp >> 14)
                lTemp = .Item(15) + .Item(14): .Item(12) = .Item(12) Xor (lTemp << 7 Or lTemp >> 25)
                lTemp = .Item(12) + .Item(15): .Item(13) = .Item(13) Xor (lTemp << 9 Or lTemp >> 23)
                lTemp = .Item(13) + .Item(12): .Item(14) = .Item(14) Xor (lTemp << 13 Or lTemp >> 19)
                lTemp = .Item(14) + .Item(13): .Item(15) = .Item(15) Xor (lTemp << 18 Or lTemp >> 14)
            End With
        #Else
            '--- Operate on columns
            Op32 X, 4, 0, 12, 7: Op32 X, 8, 4, 0, 9
            Op32 X, 12, 8, 4, 13: Op32 X, 0, 12, 8, 18
            Op32 X, 9, 5, 1, 7: Op32 X, 13, 9, 5, 9
            Op32 X, 1, 13, 9, 13: Op32 X, 5, 1, 13, 18
            Op32 X, 14, 10, 6, 7: Op32 X, 2, 14, 10, 9
            Op32 X, 6, 2, 14, 13: Op32 X, 10, 6, 2, 18
            Op32 X, 3, 15, 11, 7: Op32 X, 7, 3, 15, 9
            Op32 X, 11, 7, 3, 13: Op32 X, 15, 11, 7, 18
            '--- Operate on rows
            Op32 X, 1, 0, 3, 7: Op32 X, 2, 1, 0, 9
            Op32 X, 3, 2, 1, 13: Op32 X, 0, 3, 2, 18
            Op32 X, 6, 5, 4, 7: Op32 X, 7, 6, 5, 9
            Op32 X, 4, 7, 6, 13: Op32 X, 5, 4, 7, 18
            Op32 X, 11, 10, 9, 7: Op32 X, 8, 11, 10, 9
            Op32 X, 9, 8, 11, 13: Op32 X, 10, 9, 8, 18
            Op32 X, 12, 15, 14, 7: Op32 X, 13, 12, 15, 9
            Op32 X, 14, 13, 12, 13: Op32 X, 15, 14, 13, 18
        #End If
    Next
    For lIdx = 0 To 15
        #If HasOperators Then
            B32.Item(lIdx) += X.Item(lIdx)
        #Else
            B32.Item(lIdx) = UAdd32(B32.Item(lIdx), X.Item(lIdx))
        #End If
    Next
    Call CopyMemory(B(0), B32, 64)
End Sub

Private Sub pvBlockMix(B() As Byte, ByVal lR As Long, TempY() As Byte)
    Dim X(0 To 63)      As Byte
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    Debug.Assert UBound(B) + 1 >= 2 * lR * 64
    Debug.Assert UBound(TempY) + 1 >= 2 * lR * 64
    Call CopyMemory(X(0), B((2 * lR - 1) * 64), 64)
    For lIdx = 0 To 2 * lR - 1
        For lJdx = 0 To 63
            X(lJdx) = X(lJdx) Xor B(lIdx * 64 + lJdx)
        Next
        pvSalsa20Core X
        Call CopyMemory(TempY(lIdx * 64), X(0), 64)
    Next
    For lIdx = 0 To lR - 1
        Call CopyMemory(B((0 + lIdx) * 64), TempY((2 * lIdx + 0) * 64), 64)
    Next
    For lIdx = 0 To lR - 1
        Call CopyMemory(B((lR + lIdx) * 64), TempY((2 * lIdx + 1) * 64), 64)
    Next
End Sub

Private Function pvIntegerify(X() As Byte, ByVal lR As Long) As Long
    Debug.Assert UBound(X) + 1 >= 2 * lR * 64
    Call CopyMemory(pvIntegerify, X((2 * lR - 1) * 64), 4)
    pvIntegerify = pvIntegerify And &H7FFFFFFF
End Function

Private Sub pvROMix(X() As Byte, ByVal lR As Long, ByVal lN As Long, TempV() As Byte, TempY() As Byte)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lK              As Long
    Dim lBlockSize      As Long
    
    lBlockSize = 128 * lR
    Debug.Assert UBound(X) + 1 >= lBlockSize
    Debug.Assert UBound(TempY) + 1 >= lBlockSize
    Debug.Assert UBound(TempV) + 1 >= lN * lBlockSize
    For lIdx = 0 To lN - 1
        Call CopyMemory(TempV(lIdx * lBlockSize), X(0), lBlockSize)
        pvBlockMix X, lR, TempY
    Next
    For lIdx = 0 To lN - 1
        lK = pvIntegerify(X, lR) And (lN - 1)
        For lJdx = 0 To lBlockSize - 1
            X(lJdx) = X(lJdx) Xor TempV(lK * lBlockSize + lJdx)
        Next
        pvBlockMix X, lR, TempY
    Next
End Sub

Public Function CryptoScryptKdfByteArray(baPass() As Byte, baSalt() As Byte, _
            Optional ByVal OutSize As Long, _
            Optional ByVal CpuCost As Long = 16384, _
            Optional ByVal MemoryCost As Long = 8, _
            Optional ByVal Parallel As Long = 1) As Byte()
    Dim lN              As Long: lN = CpuCost
    Dim lR              As Long: lR = MemoryCost
    Dim lP              As Long: lP = Parallel
    Dim TempV()         As Byte
    Dim TempY()         As Byte
    Dim X()             As Byte
    Dim B()             As Byte
    Dim lIdx            As Long
    
    Debug.Assert (lN And (lN - 1)) = 0          '-- must be power of 2
    Debug.Assert CDbl(lP) * lR <= 2 ^ 30
    #If Not HasOperators Then
        If LNG_POW2(0) = 0 Then
            LNG_POW2(0) = 1
            For lIdx = 1 To 30
                LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
            Next
            LNG_POW2(31) = &H80000000
        End If
    #End If
    ReDim TempV(0 To lN * 128 * lR - 1) As Byte
    ReDim TempY(0 To 128 * lR - 1) As Byte
    ReDim X(0 To 128 * lR - 1) As Byte
    B = CryptoPbkdf2HmacSha2ByteArray(256, baPass, baSalt, OutSize:=lP * 128 * lR, NumIter:=1)
    For lIdx = 0 To lP - 1
        Call CopyMemory(X(0), B(lIdx * 128 * lR), 128 * lR)
        pvROMix X, lR, lN, TempV, TempY
        Call CopyMemory(B(lIdx * 128 * lR), X(0), 128 * lR)
    Next
    Erase TempV
    Erase TempY
    CryptoScryptKdfByteArray = CryptoPbkdf2HmacSha2ByteArray(256, baPass, B, OutSize:=OutSize, NumIter:=1)
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

Public Function CryptoScryptKdfText(sPass As String, sSalt As String, _
            Optional ByVal OutSize As Long, _
            Optional ByVal CpuCost As Long = 16384, _
            Optional ByVal MemoryCost As Long = 8, _
            Optional ByVal Parallel As Long = 1) As String
    CryptoScryptKdfText = ToHex(CryptoScryptKdfByteArray(ToUtf8Array(sPass), ToUtf8Array(sSalt), OutSize:=OutSize, CpuCost:=CpuCost, MemoryCost:=MemoryCost, Parallel:=Parallel))
End Function
