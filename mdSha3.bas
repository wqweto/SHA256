Attribute VB_Name = "mdSha3"
'--- mdSha3.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const LargeAddressAware = (Win64 = 0 And VBA7 = 0 And VBA6 = 0 And VBA5 = 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare PtrSafe Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    cElements           As Long
    lLbound             As Long
End Type

Private Const LNG_ROUNDS                As Long = 24
Private Const LNG_WORDS                 As Long = 25

#If HasPtrSafe Then
    Private LNG_POW2(0 To 63)       As LongLong
    Private LNG_ROUND_C(0 To 23)    As LongLong
#Else
    Private LNG_POW2(0 To 63)       As Variant
    Private LNG_ROUND_C(0 To 23)    As Variant
#End If

Public Type CryptoSha3Context
    DigestSize          As Long
    Capacity            As Long
    Absorbed            As Long
    #If HasPtrSafe Then
        Words(0 To LNG_WORDS - 1) As LongLong
    #Else
        Words(0 To LNG_WORDS - 1) As Variant
    #End If
    Bytes()             As Byte
    PeekArray           As SAFEARRAY1D
End Type

#If Not HasOperators Then
#If HasPtrSafe Then
Private Function ROTL64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function ROTL64(lX As Variant, ByVal lN As Long) As Variant
#End If
    '--- ROTL64 = LShift(X, n) Or RShift(X, 64 - n)
    Debug.Assert lN <> 0
    ROTL64 = ((lX And (LNG_POW2(63 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(63 - lN)) <> 0) * LNG_POW2(63)) Or _
        ((lX And (LNG_POW2(63) Xor -1)) \ LNG_POW2(64 - lN) Or -(lX < 0) * LNG_POW2(lN - 1))
End Function
#End If

Private Sub Keccak(uCtx As CryptoSha3Context)
    #If HasPtrSafe Then
        Static C(0 To 4) As LongLong
        Dim vTemp       As LongLong
        Dim aTemp()     As LongLong
    #Else
        Static C(0 To 4) As Variant
        Dim vTemp       As Variant
        Dim aTemp()     As Variant
    #End If
    Dim lRound          As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    With uCtx
    For lRound = 0 To LNG_ROUNDS - 1
        '--- Theta
        For lIdx = 0 To 4
            C(lIdx) = .Words(lIdx) Xor .Words(lIdx + 5) Xor .Words(lIdx + 10) Xor .Words(lIdx + 15) Xor .Words(lIdx + 20)
        Next
        For lIdx = 0 To 4
            #If HasOperators Then
                vTemp = C((lIdx + 4) Mod 5) Xor (C((lIdx + 1) Mod 5) << 1 Or C((lIdx + 1) Mod 5) >> 63)
            #Else
                vTemp = C((lIdx + 4) Mod 5) Xor ROTL64(C((lIdx + 1) Mod 5), 1)
            #End If
            For lJdx = 0 To 24 Step 5
                .Words(lIdx + lJdx) = .Words(lIdx + lJdx) Xor vTemp
            Next
        Next
        '--- Rho & Pi
        aTemp = .Words
        #If HasOperators Then
            .Words(10) = (aTemp(1) << 1) Or (aTemp(1) >> (64 - 1))
            .Words(20) = (aTemp(2) << 62) Or (aTemp(2) >> (64 - 62))
            .Words(5) = (aTemp(3) << 28) Or (aTemp(3) >> (64 - 28))
            .Words(15) = (aTemp(4) << 27) Or (aTemp(4) >> (64 - 27))
            .Words(16) = (aTemp(5) << 36) Or (aTemp(5) >> (64 - 36))
            .Words(1) = (aTemp(6) << 44) Or (aTemp(6) >> (64 - 44))
            .Words(11) = (aTemp(7) << 6) Or (aTemp(7) >> (64 - 6))
            .Words(21) = (aTemp(8) << 55) Or (aTemp(8) >> (64 - 55))
            .Words(6) = (aTemp(9) << 20) Or (aTemp(9) >> (64 - 20))
            .Words(7) = (aTemp(10) << 3) Or (aTemp(10) >> (64 - 3))
            .Words(17) = (aTemp(11) << 10) Or (aTemp(11) >> (64 - 10))
            .Words(2) = (aTemp(12) << 43) Or (aTemp(12) >> (64 - 43))
            .Words(12) = (aTemp(13) << 25) Or (aTemp(13) >> (64 - 25))
            .Words(22) = (aTemp(14) << 39) Or (aTemp(14) >> (64 - 39))
            .Words(23) = (aTemp(15) << 41) Or (aTemp(15) >> (64 - 41))
            .Words(8) = (aTemp(16) << 45) Or (aTemp(16) >> (64 - 45))
            .Words(18) = (aTemp(17) << 15) Or (aTemp(17) >> (64 - 15))
            .Words(3) = (aTemp(18) << 21) Or (aTemp(18) >> (64 - 21))
            .Words(13) = (aTemp(19) << 8) Or (aTemp(19) >> (64 - 8))
            .Words(14) = (aTemp(20) << 18) Or (aTemp(20) >> (64 - 18))
            .Words(24) = (aTemp(21) << 2) Or (aTemp(21) >> (64 - 2))
            .Words(9) = (aTemp(22) << 61) Or (aTemp(22) >> (64 - 61))
            .Words(19) = (aTemp(23) << 56) Or (aTemp(23) >> (64 - 56))
            .Words(4) = (aTemp(24) << 14) Or (aTemp(24) >> (64 - 14))
        #Else
            .Words(10) = ROTL64(aTemp(1), 1)
            .Words(20) = ROTL64(aTemp(2), 62)
            .Words(5) = ROTL64(aTemp(3), 28)
            .Words(15) = ROTL64(aTemp(4), 27)
            .Words(16) = ROTL64(aTemp(5), 36)
            .Words(1) = ROTL64(aTemp(6), 44)
            .Words(11) = ROTL64(aTemp(7), 6)
            .Words(21) = ROTL64(aTemp(8), 55)
            .Words(6) = ROTL64(aTemp(9), 20)
            .Words(7) = ROTL64(aTemp(10), 3)
            .Words(17) = ROTL64(aTemp(11), 10)
            .Words(2) = ROTL64(aTemp(12), 43)
            .Words(12) = ROTL64(aTemp(13), 25)
            .Words(22) = ROTL64(aTemp(14), 39)
            .Words(23) = ROTL64(aTemp(15), 41)
            .Words(8) = ROTL64(aTemp(16), 45)
            .Words(18) = ROTL64(aTemp(17), 15)
            .Words(3) = ROTL64(aTemp(18), 21)
            .Words(13) = ROTL64(aTemp(19), 8)
            .Words(14) = ROTL64(aTemp(20), 18)
            .Words(24) = ROTL64(aTemp(21), 2)
            .Words(9) = ROTL64(aTemp(22), 61)
            .Words(19) = ROTL64(aTemp(23), 56)
            .Words(4) = ROTL64(aTemp(24), 14)
        #End If
        '--- Chi
        For lJdx = 0 To 24 Step 5
            For lIdx = 0 To 4
                C(lIdx) = .Words(lIdx + lJdx)
            Next
            For lIdx = 0 To 4
                .Words(lIdx + lJdx) = .Words(lIdx + lJdx) Xor (Not C((lIdx + 1) Mod 5) And C((lIdx + 2) Mod 5))
            Next
        Next
        '--- Iota
        .Words(0) = .Words(0) Xor LNG_ROUND_C(lRound)
    Next
    End With
End Sub

#If HasPtrSafe Then
    Private Function PeekByte(uCtx As CryptoSha3Context, ByVal lOffset As Long) As Long
        #If uCtx Then '--- silence MZ-Tools
        #End If
        PeekByte = lOffset Mod 200
    End Function
#Else
    Private Function PeekByte(uCtx As CryptoSha3Context, ByVal lOffset As Long) As Long
        #If LargeAddressAware Then
            uCtx.PeekArray.pvData = (VarPtr(uCtx.Words(lOffset \ 8)) Xor &H80000000) + 8 Xor &H80000000
        #Else
            uCtx.PeekArray.pvData = VarPtr(uCtx.Words(lOffset \ 8)) + 8
        #End If
        PeekByte = lOffset Mod 8
    End Function
    
    Private Function CLngLng(vValue As Variant) As Variant
        Const VT_I8 As Long = &H14
        Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
    End Function
#End If

Public Sub CryptoSha3Init(uCtx As CryptoSha3Context, ByVal lBitSize As Long)
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    If LNG_POW2(0) = 0 Then
        LNG_POW2(0) = CLngLng(1)
        For lIdx = 1 To 63
            LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
        Next
        lIdx = 0
        For Each vElem In Split("1 8082 800000000000808A 8000000080008000 808B 80000001 8000000080008081 8000000000008009 8A 88 80008009 8000000A 8000808B 800000000000008B 8000000000008089 8000000000008003 8000000000008002 8000000000000080 800A 800000008000000A 8000000080008081 8000000000008080 80000001 8000000080008008")
            LNG_ROUND_C(lIdx) = CLngLng(CStr("&H" & vElem))
            #If HasPtrSafe Then
                Debug.Assert Hex$(LNG_ROUND_C(lIdx)) = vElem
            #End If
            lIdx = lIdx + 1
        Next
    End If
    With uCtx
        .DigestSize = (lBitSize + 7) \ 8
        .Capacity = LNG_WORDS * 8 - 2 * .DigestSize
        .Words(0) = CLngLng(0)
        For lIdx = 1 To UBound(.Words)
            .Words(lIdx) = .Words(0)
        Next
        If .PeekArray.cDims = 0 Then
            With .PeekArray
                .cDims = 1
                .fFeatures = 1 ' FADF_AUTO
                .cbElements = 1
                .cLocks = 1
                #If HasPtrSafe Then
                    .pvData = VarPtr(uCtx.Words(0))
                    .cElements = LNG_WORDS * 8
                #Else
                    .cElements = 8
                #End If
            End With
            Dim pDummy As LongPtr
            Call CopyMemory(ByVal ArrPtr(.Bytes), VarPtr(.PeekArray), LenB(pDummy))
        End If
    End With
End Sub

Public Sub CryptoSha3Update(uCtx As CryptoSha3Context, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lIdx            As Long
    Dim lOffset         As Long
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    With uCtx
        lOffset = PeekByte(uCtx, .Absorbed)
        For lIdx = Pos To Size - 1
            .Bytes(lOffset) = .Bytes(lOffset) Xor baBuffer(lIdx)
            If .Absorbed = .Capacity - 1 Then
                Keccak uCtx
                .Absorbed = 0
                lOffset = PeekByte(uCtx, .Absorbed)
            Else
                .Absorbed = .Absorbed + 1
                If lOffset = UBound(.Bytes) Then
                    lOffset = PeekByte(uCtx, .Absorbed)
                Else
                    lOffset = lOffset + 1
                End If
            End If
        Next
    End With
End Sub

Public Sub CryptoSha3Finalize(uCtx As CryptoSha3Context, baOutput() As Byte, Optional ByVal OutSize As Long, Optional ByVal LFSR As Long)
    Dim lIdx            As Long
    Dim lOffset         As Long
    Dim uEmpty          As CryptoSha3Context
    
    With uCtx
        If OutSize = 0 Then
            OutSize = .DigestSize
        End If
        If LFSR = 0 Then
            LFSR = &H6
        End If
        ReDim baOutput(0 To OutSize - 1) As Byte
        lOffset = PeekByte(uCtx, .Absorbed)
        .Bytes(lOffset) = .Bytes(lOffset) Xor LFSR
        lOffset = PeekByte(uCtx, .Capacity - 1)
        .Bytes(lOffset) = .Bytes(lOffset) Xor &H80
        For lIdx = 0 To UBound(baOutput)
            If lIdx Mod .Capacity = 0 Then
                Keccak uCtx
                lOffset = PeekByte(uCtx, 0)
            End If
            baOutput(lIdx) = .Bytes(lOffset)
            If lOffset = UBound(.Bytes) Then
                lOffset = PeekByte(uCtx, lIdx + 1)
            Else
                lOffset = lOffset + 1
            End If
        Next
    End With
    uCtx = uEmpty
End Sub

Public Function CryptoSha3ByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha3Context
    
    CryptoSha3Init uCtx, lBitSize
    CryptoSha3Update uCtx, baInput, Pos, Size
    CryptoSha3Finalize uCtx, CryptoSha3ByteArray
End Function

Public Function CryptoSha3Text(ByVal lBitSize As Long, sText As String) As String
    Const CP_UTF8       As Long = 65001
    Dim uCtx            As CryptoSha3Context
    Dim lSize           As Long
    Dim baInput()       As Byte
    Dim baOutput()      As Byte
    Dim aSplit()        As String
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baInput(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baInput(0), lSize, 0, 0)
    Else
        baInput = vbNullString
    End If
    CryptoSha3Init uCtx, lBitSize
    CryptoSha3Update uCtx, baInput, 0, lSize
    CryptoSha3Finalize uCtx, baOutput
    ReDim aSplit(0 To UBound(baOutput)) As String
    For lSize = 0 To UBound(aSplit)
        aSplit(lSize) = Right$("0" & Hex$(baOutput(lSize)), 2)
    Next
    CryptoSha3Text = LCase$(Join(aSplit, vbNullString))
End Function

Public Function CryptoKeccakByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha3Context
    
    CryptoSha3Init uCtx, lBitSize
    CryptoSha3Update uCtx, baInput, Pos, Size
    CryptoSha3Finalize uCtx, CryptoKeccakByteArray, uCtx.DigestSize, &H1
End Function

Public Function CryptoShakeByteArray(ByVal lBitSize As Long, ByVal lOutSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha3Context
    
    CryptoSha3Init uCtx, lBitSize
    CryptoSha3Update uCtx, baInput, Pos, Size
    CryptoSha3Finalize uCtx, CryptoShakeByteArray, lOutSize, &H1F
End Function

Public Function CryptoHmacSha3ByteArray(ByVal lBitSize As Long, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Const INNER_PAD     As Long = &H36
    Const OUTER_PAD     As Long = &H5C
    Dim lPadSize        As Long
    Dim lIdx            As Long
    Dim baPass()        As Byte
    Dim baPad()         As Byte
    Dim baHash()        As Byte
    
    '--- pad size is equal to sponge capacity
    lPadSize = LNG_WORDS * 8 - 2 * ((lBitSize + 7) \ 8)
    If UBound(baKey) < lPadSize Then
        baPass = baKey
    Else
        baPass = CryptoSha3ByteArray(lBitSize, baKey)
    End If
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    ReDim baPad(0 To Size + lPadSize - 1) As Byte
    For lIdx = 0 To UBound(baPass)
        baPad(lIdx) = baPass(lIdx) Xor INNER_PAD
    Next
    For lIdx = lIdx To lPadSize - 1
        baPad(lIdx) = INNER_PAD
    Next
    For lIdx = 0 To Size - 1
        baPad(lPadSize + lIdx) = baInput(Pos + lIdx)
    Next
    baHash = CryptoSha3ByteArray(lBitSize, baPad)
    Size = UBound(baHash) + 1
    ReDim baPad(0 To Size + lPadSize - 1) As Byte
    For lIdx = 0 To UBound(baPass)
        baPad(lIdx) = baPass(lIdx) Xor OUTER_PAD
    Next
    For lIdx = lIdx To lPadSize - 1
        baPad(lIdx) = OUTER_PAD
    Next
    For lIdx = 0 To Size - 1
        baPad(lPadSize + lIdx) = baHash(lIdx)
    Next
    CryptoHmacSha3ByteArray = CryptoSha3ByteArray(lBitSize, baPad)
End Function
