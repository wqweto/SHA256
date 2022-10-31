Attribute VB_Name = "mdSha1"
'--- mdSha1.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_BLOCKSZ               As Long = 64
Private Const LNG_ROUNDS                As Long = 80
Private Const LNG_HASHSZ                As Long = 20

Public Type CryptoSha1Context
    H0                  As Long
    H1                  As Long
    H2                  As Long
    H3                  As Long
    H4                  As Long
    Partial(0 To LNG_BLOCKSZ - 1) As Byte
    NPartial            As Long
    NInput              As Currency
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
#End If

Private Function BSwap32(ByVal lX As Long) As Long
    BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
                 (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
End Function

Public Sub CryptoSha1Init(uCtx As CryptoSha1Context)
    #If Not HasOperators Then
        Dim lIdx            As Long
        
        If LNG_POW2(0) = 0 Then
            LNG_POW2(0) = 1
            For lIdx = 1 To 30
                LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
            Next
            LNG_POW2(31) = &H80000000
        End If
    #End If
    With uCtx
        .H0 = &H67452301: .H1 = &HEFCDAB89: .H2 = &H98BADCFE: .H3 = &H10325476: .H4 = &HC3D2E1F0
        .NPartial = 0
        .NInput = 0
    End With
End Sub

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Public Sub CryptoSha1Update(uCtx As CryptoSha1Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Static W(0 To LNG_ROUNDS - 1) As Long
    Static B(0 To 15)   As Long
    Dim lIdx            As Long
    Dim lA              As Long
    Dim lB              As Long
    Dim lC              As Long
    Dim lD              As Long
    Dim lE              As Long
    Dim lTemp           As Long
    Dim lK              As Long
    
    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        .NInput = .NInput + Size
        If .NPartial > 0 Then
            lTemp = LNG_BLOCKSZ - .NPartial
            If lTemp > Size Then
                lTemp = Size
            End If
            Call CopyMemory(.Partial(.NPartial), baInput(Pos), lTemp)
            .NPartial = .NPartial + lTemp
            Pos = Pos + lTemp
            Size = Size - lTemp
        End If
        Do While Size > 0 Or .NPartial = LNG_BLOCKSZ
            If .NPartial <> 0 Then
                Call CopyMemory(B(0), .Partial(0), LNG_BLOCKSZ)
                .NPartial = 0
            ElseIf Size >= LNG_BLOCKSZ Then
                Call CopyMemory(B(0), baInput(Pos), LNG_BLOCKSZ)
                Pos = Pos + LNG_BLOCKSZ
                Size = Size - LNG_BLOCKSZ
            Else
                Call CopyMemory(.Partial(0), baInput(Pos), Size)
                .NPartial = Size
                Exit Do
            End If
            '--- sha1 step
            lA = .H0: lB = .H1: lC = .H2: lD = .H3: lE = .H4
            For lIdx = 0 To LNG_ROUNDS - 1
                If lIdx < 16 Then
                    W(lIdx) = BSwap32(B(lIdx))
                Else
                    #If HasOperators Then
                        lTemp = W(lIdx - 3) Xor W(lIdx - 8) Xor W(lIdx - 14) Xor W(lIdx - 16)
                        W(lIdx) = (lTemp << 1 Or lTemp >> 31)
                    #Else
                        W(lIdx) = RotL32(W(lIdx - 3) Xor W(lIdx - 8) Xor W(lIdx - 14) Xor W(lIdx - 16), 1)
                    #End If
                End If
                Select Case lIdx
                Case 0 To 19
                    lTemp = (lB And lC) Or ((Not lB) And lD)
                    lK = &H5A827999
                Case 20 To 39
                    lTemp = lB Xor lC Xor lD
                    lK = &H6ED9EBA1
                Case 40 To 59
                    lTemp = (lB And lC) Or (lB And lD) Or (lC And lD)
                    lK = &H8F1BBCDC
                Case 60 To 79
                    lTemp = lB Xor lC Xor lD
                    lK = &HCA62C1D6
                End Select
                #If HasOperators Then
                    lTemp += (lA << 5 or lA >> 27) + lE + lK + W(lIdx)
                #Else
                    lTemp = UAdd32(UAdd32(UAdd32(UAdd32(lTemp, RotL32(lA, 5)), lE), lK), W(lIdx))
                #End If
                lE = lD
                lD = lC
                #If HasOperators Then
                    lC = (lB << 30 Or lB >> 2)
                #Else
                    lC = RotL32(lB, 30)
                #End If
                lB = lA
                lA = lTemp
            Next
            #If HasOperators Then
                .H0 += lA: .H1 += lB: .H2 += lC: .H3 += lD: .H4 += lE
            #Else
                .H0 = UAdd32(.H0, lA): .H1 = UAdd32(.H1, lB): .H2 = UAdd32(.H2, lC): .H3 = UAdd32(.H3, lD): .H4 = UAdd32(.H4, lE)
            #End If
        Loop
    End With
End Sub

Public Sub CryptoSha1Finalize(uCtx As CryptoSha1Context, baOutput() As Byte)
    Static B(0 To 4)    As Long
    Dim P(0 To LNG_BLOCKSZ + 9) As Byte
    Dim lSize           As Long
    
    With uCtx
        lSize = LNG_BLOCKSZ - .NPartial
        If lSize < 9 Then
            lSize = lSize + LNG_BLOCKSZ
        End If
        P(0) = &H80
        .NInput = .NInput / 10000@ * 8
        Call CopyMemory(B(0), .NInput, 8)
        Call CopyMemory(P(lSize - 4), BSwap32(B(0)), 4)
        Call CopyMemory(P(lSize - 8), BSwap32(B(1)), 4)
        CryptoSha1Update uCtx, P, Size:=lSize
        Debug.Assert .NPartial = 0
        B(0) = BSwap32(.H0): B(1) = BSwap32(.H1): B(2) = BSwap32(.H2): B(3) = BSwap32(.H3): B(4) = BSwap32(.H4)
        ReDim baOutput(0 To LNG_HASHSZ - 1) As Byte
        Call CopyMemory(baOutput(0), B(0), UBound(baOutput) + 1)
    End With
End Sub

Public Function CryptoSha1ByteArray(baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha1Context
    
    CryptoSha1Init uCtx
    CryptoSha1Update uCtx, baInput, Pos, Size
    CryptoSha1Finalize uCtx, CryptoSha1ByteArray
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

Public Function CryptoSha1Text(sText As String) As String
    CryptoSha1Text = ToHex(CryptoSha1ByteArray(ToUtf8Array(sText)))
End Function

Public Function CryptoHmacSha1ByteArray(baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Const INNER_PAD     As Long = &H36
    Const OUTER_PAD     As Long = &H5C
    Dim lPadSize        As Long
    Dim lIdx            As Long
    Dim baPass()        As Byte
    Dim baPad()         As Byte
    Dim baHash()        As Byte
    
    lPadSize = LNG_BLOCKSZ
    If UBound(baKey) < lPadSize Then
        baPass = baKey
    Else
        baPass = CryptoSha1ByteArray(baKey)
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
    baHash = CryptoSha1ByteArray(baPad)
    Size = UBound(baHash) + 1
    ReDim baPad(0 To lPadSize + Size - 1) As Byte
    For lIdx = 0 To UBound(baPass)
        baPad(lIdx) = baPass(lIdx) Xor OUTER_PAD
    Next
    For lIdx = lIdx To lPadSize - 1
        baPad(lIdx) = OUTER_PAD
    Next
    Call CopyMemory(baPad(lPadSize), baHash(0), Size)
    CryptoHmacSha1ByteArray = CryptoSha1ByteArray(baPad)
End Function

Public Function CryptoHmacSha1Text(sKey As String, sText As String) As String
    CryptoHmacSha1Text = ToHex(CryptoHmacSha1ByteArray(ToUtf8Array(sKey), ToUtf8Array(sText)))
End Function

Public Function CryptoPbkdf2HmacSha1ByteArray(baPass() As Byte, baSalt() As Byte, _
            Optional ByVal OutSize As Long, _
            Optional ByVal NumIter As Long = 10000) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lKdx            As Long
    Dim baInit()        As Byte
    Dim baHmac()        As Byte
    Dim baTemp()        As Byte
    Dim lRemaining      As Long
    
    If NumIter <= 0 Then
        baRetVal = vbNullString
    Else
        If OutSize <= 0 Then
            OutSize = LNG_HASHSZ
        End If
        ReDim baRetVal(0 To OutSize - 1) As Byte
        baInit = baSalt
        ReDim Preserve baInit(0 To LenB(CStr(baInit)) + 3) As Byte
        For lIdx = 0 To (OutSize + LNG_HASHSZ - 1) \ LNG_HASHSZ - 1
            Call CopyMemory(baInit(UBound(baInit) - 3), BSwap32(lIdx + 1), 4)
            baTemp = baInit
            ReDim baHmac(0 To LNG_HASHSZ - 1) As Byte
            For lJdx = 0 To NumIter - 1
                baTemp = CryptoHmacSha1ByteArray(baPass, baTemp)
                For lKdx = 0 To UBound(baTemp)
                    baHmac(lKdx) = baHmac(lKdx) Xor baTemp(lKdx)
                Next
            Next
            lRemaining = OutSize - lIdx * LNG_HASHSZ
            If lRemaining > LNG_HASHSZ Then
                lRemaining = LNG_HASHSZ
            End If
            Call CopyMemory(baRetVal(lIdx * LNG_HASHSZ), baHmac(0), lRemaining)
        Next
    End If
    CryptoPbkdf2HmacSha1ByteArray = baRetVal
End Function

Public Function CryptoPbkdf2HmacSha1Text(sPass As String, sSalt As String, _
            Optional ByVal OutSize As Long, _
            Optional ByVal NumIter As Long = 10000) As String
    CryptoPbkdf2HmacSha1Text = ToHex(CryptoPbkdf2HmacSha1ByteArray(ToUtf8Array(sPass), ToUtf8Array(sSalt), NumIter:=NumIter, OutSize:=OutSize))
End Function

Public Function CryptoHkdfSha1ByteArray(baIKM() As Byte, baSalt() As Byte, baInfo() As Byte, Optional ByVal OutSize As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim baKey()         As Byte
    Dim baPad()         As Byte
    Dim baHash()        As Byte
    Dim lIdx            As Long
    Dim lRemaining      As Long
    
    If OutSize <= 0 Then
        OutSize = LNG_HASHSZ
    End If
    ReDim baRetVal(0 To OutSize - 1) As Byte
    baKey = CryptoHmacSha1ByteArray(baSalt, baIKM)
    ReDim baPad(0 To LNG_HASHSZ + UBound(baInfo) + 1) As Byte
    If UBound(baInfo) >= 0 Then
        Call CopyMemory(baPad(LNG_HASHSZ), baInfo(0), UBound(baInfo) + 1)
    End If
    For lIdx = 0 To (OutSize + LNG_HASHSZ - 1) \ LNG_HASHSZ - 1
        baPad(UBound(baPad)) = (lIdx + 1) And &HFF
        baHash = CryptoHmacSha1ByteArray(baKey, baPad, Pos:=-(lIdx = 0) * LNG_HASHSZ)
        Call CopyMemory(baPad(0), baHash(0), LNG_HASHSZ)
        lRemaining = OutSize - lIdx * LNG_HASHSZ
        If lRemaining > LNG_HASHSZ Then
            lRemaining = LNG_HASHSZ
        End If
        Call CopyMemory(baRetVal(lIdx * LNG_HASHSZ), baHash(0), lRemaining)
    Next
    CryptoHkdfSha1ByteArray = baRetVal
End Function

Public Function CryptoHkdfSha1Text(sIKM As String, sSalt As String, sInfo As String, Optional ByVal OutSize As Long) As String
    CryptoHkdfSha1Text = ToHex(CryptoHkdfSha1ByteArray(ToUtf8Array(sIKM), ToUtf8Array(sSalt), ToUtf8Array(sInfo), OutSize:=OutSize))
End Function
