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
#If Not HasOperators Then
    Private LNG_POW2(0 To 63)       As LongLong
    Private LNG_SIGN_BIT            As LongLong ' 2 ^ 63
#End If
    Private LNG_ROUND_C(0 To 23)    As LongLong
#Else
    Private LNG_POW2(0 To 63)       As Variant
    Private LNG_SIGN_BIT            As Variant
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
Private Function RotL64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RotL64(lX As Variant, ByVal lN As Long) As Variant
#End If
    '--- RotL64 = LShift(X, n) Or RShift(X, 64 - n)
    Debug.Assert lN <> 0
    RotL64 = ((lX And (LNG_POW2(63 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(63 - lN)) <> 0) * LNG_SIGN_BIT) Or _
        ((lX And (LNG_SIGN_BIT Xor -1)) \ LNG_POW2(64 - lN) Or -(lX < 0) * LNG_POW2(lN - 1))
End Function
#End If

#If Not HasPtrSafe Then
Private Function CLngLng(vValue As Variant) As Variant
    Const VT_I8 As Long = &H14
    Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
End Function

Private Function PeekByte(uCtx As CryptoSha3Context, ByVal lOffset As Long) As Long
    #If LargeAddressAware Then
        uCtx.PeekArray.pvData = (VarPtr(uCtx.Words(lOffset \ 8)) Xor &H80000000) + 8 Xor &H80000000
    #Else
        uCtx.PeekArray.pvData = VarPtr(uCtx.Words(lOffset \ 8)) + 8
    #End If
    PeekByte = lOffset Mod 8
End Function
#End If

Private Sub Keccak(uCtx As CryptoSha3Context)
#If HasPtrSafe Then
    Static C(0 To 4) As LongLong
    Static T(0 To LNG_WORDS - 1) As LongLong
    Dim lTemp       As LongLong
#Else
    Static C(0 To 4) As Variant
    Static T(0 To LNG_WORDS - 1) As Variant
    Dim lTemp       As Variant
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
                lTemp = C((lIdx + 1) Mod 5)
                lTemp = C((lIdx + 4) Mod 5) Xor (lTemp << 1 Or lTemp >> 63)
            #Else
                lTemp = C((lIdx + 4) Mod 5) Xor RotL64(C((lIdx + 1) Mod 5), 1)
            #End If
            For lJdx = 0 To 24 Step 5
                .Words(lIdx + lJdx) = .Words(lIdx + lJdx) Xor lTemp
            Next
        Next
        '--- Rho & Pi
        #If HasPtrSafe Then
            Call CopyMemory(T(0), .Words(0), LNG_WORDS * 8)
        #Else
            Call CopyMemory(T(0), .Words(0), LNG_WORDS * 16)
        #End If
        #If HasOperators Then
            .Words(10) = (T(1) << 1) Or (T(1) >> (64 - 1))
            .Words(20) = (T(2) << 62) Or (T(2) >> (64 - 62))
            .Words(5) = (T(3) << 28) Or (T(3) >> (64 - 28))
            .Words(15) = (T(4) << 27) Or (T(4) >> (64 - 27))
            .Words(16) = (T(5) << 36) Or (T(5) >> (64 - 36))
            .Words(1) = (T(6) << 44) Or (T(6) >> (64 - 44))
            .Words(11) = (T(7) << 6) Or (T(7) >> (64 - 6))
            .Words(21) = (T(8) << 55) Or (T(8) >> (64 - 55))
            .Words(6) = (T(9) << 20) Or (T(9) >> (64 - 20))
            .Words(7) = (T(10) << 3) Or (T(10) >> (64 - 3))
            .Words(17) = (T(11) << 10) Or (T(11) >> (64 - 10))
            .Words(2) = (T(12) << 43) Or (T(12) >> (64 - 43))
            .Words(12) = (T(13) << 25) Or (T(13) >> (64 - 25))
            .Words(22) = (T(14) << 39) Or (T(14) >> (64 - 39))
            .Words(23) = (T(15) << 41) Or (T(15) >> (64 - 41))
            .Words(8) = (T(16) << 45) Or (T(16) >> (64 - 45))
            .Words(18) = (T(17) << 15) Or (T(17) >> (64 - 15))
            .Words(3) = (T(18) << 21) Or (T(18) >> (64 - 21))
            .Words(13) = (T(19) << 8) Or (T(19) >> (64 - 8))
            .Words(14) = (T(20) << 18) Or (T(20) >> (64 - 18))
            .Words(24) = (T(21) << 2) Or (T(21) >> (64 - 2))
            .Words(9) = (T(22) << 61) Or (T(22) >> (64 - 61))
            .Words(19) = (T(23) << 56) Or (T(23) >> (64 - 56))
            .Words(4) = (T(24) << 14) Or (T(24) >> (64 - 14))
        #Else
            .Words(10) = RotL64(T(1), 1)
            .Words(20) = RotL64(T(2), 62)
            .Words(5) = RotL64(T(3), 28)
            .Words(15) = RotL64(T(4), 27)
            .Words(16) = RotL64(T(5), 36)
            .Words(1) = RotL64(T(6), 44)
            .Words(11) = RotL64(T(7), 6)
            .Words(21) = RotL64(T(8), 55)
            .Words(6) = RotL64(T(9), 20)
            .Words(7) = RotL64(T(10), 3)
            .Words(17) = RotL64(T(11), 10)
            .Words(2) = RotL64(T(12), 43)
            .Words(12) = RotL64(T(13), 25)
            .Words(22) = RotL64(T(14), 39)
            .Words(23) = RotL64(T(15), 41)
            .Words(8) = RotL64(T(16), 45)
            .Words(18) = RotL64(T(17), 15)
            .Words(3) = RotL64(T(18), 21)
            .Words(13) = RotL64(T(19), 8)
            .Words(14) = RotL64(T(20), 18)
            .Words(24) = RotL64(T(21), 2)
            .Words(9) = RotL64(T(22), 61)
            .Words(19) = RotL64(T(23), 56)
            .Words(4) = RotL64(T(24), 14)
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

Public Sub CryptoSha3Init(uCtx As CryptoSha3Context, ByVal lBitSize As Long)
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim pDummy          As LongPtr
    
    If LNG_ROUND_C(0) = 0 Then
        For Each vElem In Split("1 8082 800000000000808A 8000000080008000 808B 80000001 8000000080008081 8000000000008009 8A 88 80008009 8000000A 8000808B 800000000000008B 8000000000008089 8000000000008003 8000000000008002 8000000000000080 800A 800000008000000A 8000000080008081 8000000000008080 80000001 8000000080008008")
            LNG_ROUND_C(lIdx) = CLngLng(CStr("&H" & vElem))
            #If HasPtrSafe Then
                Debug.Assert Hex$(LNG_ROUND_C(lIdx)) = vElem
            #End If
            lIdx = lIdx + 1
        Next
        #If Not HasOperators Then
            LNG_POW2(0) = CLngLng(1)
            For lIdx = 1 To 63
                LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
            Next
            LNG_SIGN_BIT = LNG_POW2(63)
        #End If
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
        #If HasPtrSafe Then
            lOffset = .Absorbed Mod 200
        #Else
            lOffset = PeekByte(uCtx, .Absorbed)
        #End If
        For lIdx = Pos To Size - 1
            .Bytes(lOffset) = .Bytes(lOffset) Xor baBuffer(lIdx)
            If .Absorbed = .Capacity - 1 Then
                Keccak uCtx
                .Absorbed = 0
                #If HasPtrSafe Then
                    lOffset = .Absorbed Mod 200
                #Else
                    lOffset = PeekByte(uCtx, .Absorbed)
                #End If
            Else
                .Absorbed = .Absorbed + 1
                If lOffset = UBound(.Bytes) Then
                    #If HasPtrSafe Then
                        lOffset = .Absorbed Mod 200
                    #Else
                        lOffset = PeekByte(uCtx, .Absorbed)
                    #End If
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
    Dim pDummy          As LongPtr
    Dim uEmpty          As CryptoSha3Context
    
    With uCtx
        If OutSize = 0 Then
            OutSize = .DigestSize
        End If
        If LFSR = 0 Then
            LFSR = &H6
        End If
        ReDim baOutput(0 To OutSize - 1) As Byte
        #If HasPtrSafe Then
            lOffset = .Absorbed Mod 200
        #Else
            lOffset = PeekByte(uCtx, .Absorbed)
        #End If
        .Bytes(lOffset) = .Bytes(lOffset) Xor LFSR
        #If HasPtrSafe Then
            lOffset = (.Capacity - 1) Mod 200
        #Else
            lOffset = PeekByte(uCtx, .Capacity - 1)
        #End If
        .Bytes(lOffset) = .Bytes(lOffset) Xor &H80
        For lIdx = 0 To UBound(baOutput)
            If lIdx Mod .Capacity = 0 Then
                Keccak uCtx
                #If HasPtrSafe Then
                    lOffset = 0
                #Else
                    lOffset = PeekByte(uCtx, 0)
                #End If
            End If
            baOutput(lIdx) = .Bytes(lOffset)
            If lOffset = UBound(.Bytes) Then
                #If HasPtrSafe Then
                    lOffset = (lIdx + 1) Mod 200
                #Else
                    lOffset = PeekByte(uCtx, lIdx + 1)
                #End If
            Else
                lOffset = lOffset + 1
            End If
        Next
        Call CopyMemory(ByVal ArrPtr(.Bytes), pDummy, LenB(pDummy))
    End With
    uCtx = uEmpty
End Sub

Public Function CryptoSha3ByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha3Context
    
    CryptoSha3Init uCtx, lBitSize
    CryptoSha3Update uCtx, baInput, Pos, Size
    CryptoSha3Finalize uCtx, CryptoSha3ByteArray
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

Public Function CryptoSha3Text(ByVal lBitSize As Long, sText As String) As String
    CryptoSha3Text = ToHex(CryptoSha3ByteArray(lBitSize, ToUtf8Array(sText)))
End Function

Public Function CryptoKeccakByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha3Context
    
    CryptoSha3Init uCtx, lBitSize
    CryptoSha3Update uCtx, baInput, Pos, Size
    CryptoSha3Finalize uCtx, CryptoKeccakByteArray, uCtx.DigestSize, &H1
End Function

Public Function CryptoKeccakText(ByVal lBitSize As Long, sText As String) As String
    CryptoKeccakText = ToHex(CryptoKeccakByteArray(lBitSize, ToUtf8Array(sText)))
End Function

Public Function CryptoShakeByteArray(ByVal lBitSize As Long, ByVal lOutSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha3Context
    
    CryptoSha3Init uCtx, lBitSize
    CryptoSha3Update uCtx, baInput, Pos, Size
    CryptoSha3Finalize uCtx, CryptoShakeByteArray, lOutSize, &H1F
End Function

Public Function CryptoShakeText(ByVal lBitSize As Long, ByVal lOutSize As Long, sText As String) As String
    CryptoShakeText = ToHex(CryptoShakeByteArray(lBitSize, lOutSize, ToUtf8Array(sText)))
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
    ReDim baPad(0 To lPadSize + Size - 1) As Byte
    For lIdx = 0 To UBound(baPass)
        baPad(lIdx) = baPass(lIdx) Xor INNER_PAD
    Next
    For lIdx = lIdx To lPadSize - 1
        baPad(lIdx) = INNER_PAD
    Next
    If Size > 0 Then
        Call CopyMemory(baPad(lPadSize), baInput(Pos), Size)
    End If
    baHash = CryptoSha3ByteArray(lBitSize, baPad)
    Size = UBound(baHash) + 1
    ReDim baPad(0 To lPadSize + Size - 1) As Byte
    For lIdx = 0 To UBound(baPass)
        baPad(lIdx) = baPass(lIdx) Xor OUTER_PAD
    Next
    For lIdx = lIdx To lPadSize - 1
        baPad(lIdx) = OUTER_PAD
    Next
    Call CopyMemory(baPad(lPadSize), baHash(0), Size)
    CryptoHmacSha3ByteArray = CryptoSha3ByteArray(lBitSize, baPad)
End Function

Public Function CryptoHmacSha3Text(ByVal lBitSize As Long, sKey As String, sText As String) As String
    CryptoHmacSha3Text = ToHex(CryptoHmacSha3ByteArray(lBitSize, ToUtf8Array(sKey), ToUtf8Array(sText)))
End Function

Private Function BSwap32(ByVal lX As Long) As Long
    BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
                 (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
End Function

Public Function CryptoPbkdf2HmacSha3ByteArray(ByVal lBitSize As Long, baPass() As Byte, baSalt() As Byte, _
            Optional ByVal OutSize As Long, _
            Optional ByVal NumIter As Long = 10000) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lKdx            As Long
    Dim lHashSize       As Long
    Dim baInit()        As Byte
    Dim baHmac()        As Byte
    Dim baTemp()        As Byte
    Dim lRemaining      As Long
    
    If NumIter <= 0 Then
        baRetVal = vbNullString
    Else
        If OutSize <= 0 Then
            OutSize = (lBitSize + 7) \ 8
        End If
        ReDim baRetVal(0 To OutSize - 1) As Byte
        baInit = baSalt
        ReDim Preserve baInit(0 To LenB(CStr(baInit)) + 3) As Byte
        lHashSize = (lBitSize + 7) \ 8
        For lIdx = 0 To (OutSize + lHashSize - 1) \ lHashSize - 1
            Call CopyMemory(baInit(UBound(baInit) - 3), BSwap32(lIdx + 1), 4)
            baTemp = baInit
            ReDim baHmac(0 To lHashSize - 1) As Byte
            For lJdx = 0 To NumIter - 1
                baTemp = CryptoHmacSha3ByteArray(lBitSize, baPass, baTemp)
                For lKdx = 0 To UBound(baTemp)
                    baHmac(lKdx) = baHmac(lKdx) Xor baTemp(lKdx)
                Next
            Next
            lRemaining = OutSize - lIdx * lHashSize
            If lRemaining > lHashSize Then
                lRemaining = lHashSize
            End If
            Call CopyMemory(baRetVal(lIdx * lHashSize), baHmac(0), lRemaining)
        Next
    End If
    CryptoPbkdf2HmacSha3ByteArray = baRetVal
End Function

Public Function CryptoPbkdf2HmacSha3Text(ByVal lBitSize As Long, sPass As String, sSalt As String, _
            Optional ByVal OutSize As Long, _
            Optional ByVal NumIter As Long = 10000) As String
    CryptoPbkdf2HmacSha3Text = ToHex(CryptoPbkdf2HmacSha3ByteArray(lBitSize, ToUtf8Array(sPass), ToUtf8Array(sSalt), NumIter:=NumIter, OutSize:=OutSize))
End Function

Public Function CryptoHkdfSha3ByteArray(ByVal lBitSize As Long, baIKM() As Byte, baSalt() As Byte, baInfo() As Byte, Optional ByVal OutSize As Long) As Byte()
    Dim lHashSize       As Long
    Dim baRetVal()      As Byte
    Dim baKey()         As Byte
    Dim baPad()         As Byte
    Dim baHash()        As Byte
    Dim lIdx            As Long
    Dim lRemaining      As Long
    
    lHashSize = (lBitSize + 7) \ 8
    If OutSize <= 0 Then
        OutSize = lHashSize
    End If
    ReDim baRetVal(0 To OutSize - 1) As Byte
    baKey = CryptoHmacSha3ByteArray(lBitSize, baSalt, baIKM)
    ReDim baPad(0 To lHashSize + UBound(baInfo) + 1) As Byte
    If UBound(baInfo) >= 0 Then
        Call CopyMemory(baPad(lHashSize), baInfo(0), UBound(baInfo) + 1)
    End If
    For lIdx = 0 To (OutSize + lHashSize - 1) \ lHashSize - 1
        baPad(UBound(baPad)) = (lIdx + 1) And &HFF
        baHash = CryptoHmacSha3ByteArray(lBitSize, baKey, baPad, Pos:=-(lIdx = 0) * lHashSize)
        Call CopyMemory(baPad(0), baHash(0), lHashSize)
        lRemaining = OutSize - lIdx * lHashSize
        If lRemaining > lHashSize Then
            lRemaining = lHashSize
        End If
        Call CopyMemory(baRetVal(lIdx * lHashSize), baHash(0), lRemaining)
    Next
    CryptoHkdfSha3ByteArray = baRetVal
End Function

Public Function CryptoHkdfSha3Text(ByVal lBitSize As Long, sIKM As String, sSalt As String, sInfo As String, Optional ByVal OutSize As Long) As String
    CryptoHkdfSha3Text = ToHex(CryptoHkdfSha3ByteArray(lBitSize, ToUtf8Array(sIKM), ToUtf8Array(sSalt), ToUtf8Array(sInfo), OutSize:=OutSize))
End Function
