Attribute VB_Name = "mdSha2"
'--- mdSha2.bas
Option Explicit
DefObj A-Z

#Const HasSha512 = (CRYPT_HAS_SHA512 <> 0)
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
Private Const LNG_ROUNDS                As Long = 64

Public Type CryptoSha2Context
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
    BitSize             As Long
End Type

Private LNG_K(0 To LNG_ROUNDS - 1)  As Long

#If Not HasOperators Then
Private LNG_POW2(0 To 31)           As Long

Private Function ROTR32(ByVal lX As Long, ByVal lN As Long) As Long
    '--- ROTR32 = RShift(X, n) Or LShift(X, 32 - n)
    Debug.Assert lN <> 0
    ROTR32 = ((lX And &H7FFFFFFF) \ LNG_POW2(lN) - (lX < 0) * LNG_POW2(31 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(32 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * &H80000000)
End Function

'Private Function LShift(ByVal lX As Long, ByVal lN As Long) As Long
'    If lN = 0 Then
'        LShift = lX
'    Else
'        LShift = (lX And (LNG_POW2(31 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(31 - lN)) <> 0) * &H80000000
'    End If
'End Function

Private Function RShift(ByVal lX As Long, ByVal lN As Long) As Long
    If lN = 0 Then
        RShift = lX
    Else
        RShift = (lX And &H7FFFFFFF) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(31 - lN)
    End If
End Function

Private Function UAdd(ByVal lX As Long, ByVal lY As Long) As Long
    If (lX Xor lY) >= 0 Then
        UAdd = ((lX Xor &H80000000) + lY) Xor &H80000000
    Else
        UAdd = lX + lY
    End If
End Function

Private Function ByteSwap(ByVal lX As Long) As Long
    ByteSwap = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or (lX And &H7F000000) \ &H1000000 Or _
                -((lX And &H80) <> 0) * &H80000000 Or -((lX And &H80000000) <> 0) * &H80
End Function

Private Function Ch(ByVal lX As Long, ByVal lY As Long, ByVal lZ As Long) As Long
    Ch = (lX And lY) Xor ((Not lX) And lZ)
End Function

Private Function Maj(ByVal lX As Long, ByVal lY As Long, ByVal lZ As Long) As Long
    Maj = (lX And lY) Xor (lX And lZ) Xor (lY And lZ)
End Function

Private Function BigSigma0(ByVal lX As Long) As Long
    BigSigma0 = ROTR32(lX, 2) Xor ROTR32(lX, 13) Xor ROTR32(lX, 22)
End Function

Private Function BigSigma1(ByVal lX As Long) As Long
    BigSigma1 = ROTR32(lX, 6) Xor ROTR32(lX, 11) Xor ROTR32(lX, 25)
End Function

Private Function SmallSigma0(ByVal lX As Long) As Long
    SmallSigma0 = ROTR32(lX, 7) Xor ROTR32(lX, 18) Xor RShift(lX, 3)
End Function

Private Function SmallSigma1(ByVal lX As Long) As Long
    SmallSigma1 = ROTR32(lX, 17) Xor ROTR32(lX, 19) Xor RShift(lX, 10)
End Function
#End If

Public Sub CryptoSha2Init(uCtx As CryptoSha2Context, ByVal lBitSize As Long)
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    If LNG_K(0) = 0 Then
        '--- K: first 32 bits of the fractional parts of the cube roots of the first 64 primes
        For Each vElem In Split("428A2F98 71374491 B5C0FBCF E9B5DBA5 3956C25B 59F111F1 923F82A4 AB1C5ED5 D807AA98 12835B01 243185BE 550C7DC3 72BE5D74 80DEB1FE 9BDC06A7 C19BF174 E49B69C1 EFBE4786 FC19DC6 240CA1CC 2DE92C6F 4A7484AA 5CB0A9DC 76F988DA 983E5152 A831C66D B00327C8 BF597FC7 C6E00BF3 D5A79147 6CA6351 14292967 27B70A85 2E1B2138 4D2C6DFC 53380D13 650A7354 766A0ABB 81C2C92E 92722C85 A2BFE8A1 A81A664B C24B8B70 C76C51A3 D192E819 D6990624 F40E3585 106AA070 19A4C116 1E376C08 2748774C 34B0BCB5 391C0CB3 4ED8AA4A 5B9CCA4F 682E6FF3 748F82EE 78A5636F 84C87814 8CC70208 90BEFFFA A4506CEB BEF9A3F7 C67178F2")
            LNG_K(lIdx) = "&H" & vElem
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
    With uCtx
        Select Case lBitSize
        Case 224
            .H0 = &HC1059ED8: .H1 = &H367CD507: .H2 = &H3070DD17: .H3 = &HF70E5939
            .H4 = &HFFC00B31: .H5 = &H68581511: .H6 = &H64F98FA7: .H7 = &HBEFA4FA4
        Case 256
            .H0 = &H6A09E667: .H1 = &HBB67AE85: .H2 = &H3C6EF372: .H3 = &HA54FF53A
            .H4 = &H510E527F: .H5 = &H9B05688C: .H6 = &H1F83D9AB: .H7 = &H5BE0CD19
        Case Else
            Err.Raise vbObjectError, , "Invalid bit-size for SHA-2 (" & lBitSize & ")"
        End Select
        .NPartial = 0
        .NInput = 0
        .BitSize = lBitSize
    End With
End Sub

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Public Sub CryptoSha2Update(uCtx As CryptoSha2Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Static W(0 To LNG_ROUNDS - 1) As Long
    Static B(0 To 15)   As Long
    Dim lIdx            As Long
    Dim lA              As Long
    Dim lB              As Long
    Dim lC              As Long
    Dim lD              As Long
    Dim lE              As Long
    Dim lF              As Long
    Dim lG              As Long
    Dim lH              As Long
    Dim lT1             As Long
    Dim lT2             As Long
    Dim lX              As Long
    Dim lSigma1         As Long
    Dim lSigma0         As Long
    Dim lCh             As Long
    Dim lMaj            As Long
    
    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        .NInput = .NInput + Size
        If .NPartial > 0 Then
            lIdx = LNG_BLOCKSZ - .NPartial
            If lIdx > Size Then
                lIdx = Size
            End If
            Call CopyMemory(.Partial(.NPartial), baInput(Pos), lIdx)
            .NPartial = .NPartial + lIdx
            Pos = Pos + lIdx
            Size = Size - lIdx
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
            '--- sha2 step
            lA = .H0: lB = .H1: lC = .H2: lD = .H3
            lE = .H4: lF = .H5: lG = .H6: lH = .H7
            For lIdx = 0 To LNG_ROUNDS - 1
                If lIdx < 16 Then
                    W(lIdx) = ByteSwap(B(lIdx))
                Else
                    #If HasOperators Then
                        lX = W(lIdx - 2)
                        lSigma1 = (lX >> 17 Or lX << 15) Xor (lX >> 19 Or lX << 13) Xor (lX >> 10)
                        lX = W(lIdx - 15)
                        lSigma0 = (lX >> 7 Or lX << 25) Xor (lX >> 18 Or lX << 14) Xor (lX >> 3)
                        W(lIdx) = lSigma1 + W(lIdx - 7) + lSigma0 + W(lIdx - 16)
                    #Else
                        W(lIdx) = UAdd(UAdd(UAdd(SmallSigma1(W(lIdx - 2)), W(lIdx - 7)), SmallSigma0(W(lIdx - 15))), W(lIdx - 16))
                    #End If
                End If
                #If HasOperators Then
                    lSigma1 = (lE >> 6 Or lE << 26) Xor (lE >> 11 Or lE << 21) Xor (lE >> 25 Or lE << 7)
                    lSigma0 = (lA >> 2 Or lA << 30) Xor (lA >> 13 Or lA << 19) Xor (lA >> 22 Or lA << 10)
                    lCh = (lE And lF) Xor ((Not lE) And lG)
                    lMaj = (lA And lB) Xor (lA And lC) Xor (lB And lC)
                    lT1 = lH + lSigma1 + lCh + K(lIdx) + W(lIdx)
                    lT2 = lSigma0 + lMaj
                #Else
                    lT1 = UAdd(UAdd(UAdd(UAdd(lH, BigSigma1(lE)), Ch(lE, lF, lG)), LNG_K(lIdx)), W(lIdx))
                    lT2 = UAdd(BigSigma0(lA), Maj(lA, lB, lC))
                #End If
                lH = lG
                lG = lF
                lF = lE
                #If HasOperators Then
                    lE = lD + lT1
                #Else
                    lE = UAdd(lD, lT1)
                #End If
                lD = lC
                lC = lB
                lB = lA
                #If HasOperators Then
                    lA = lT1 + lT2
                #Else
                    lA = UAdd(lT1, lT2)
                #End If
            Next
            #If HasOperators Then
                .H0 += lA: .H1 += lB: .H2 += lC: .H3 += lD
                .H4 += lE: .H5 += lF: .H6 += lG: .H7 += lH
            #Else
                .H0 = UAdd(.H0, lA): .H1 = UAdd(.H1, lB): .H2 = UAdd(.H2, lC): .H3 = UAdd(.H3, lD)
                .H4 = UAdd(.H4, lE): .H5 = UAdd(.H5, lF): .H6 = UAdd(.H6, lG): .H7 = UAdd(.H7, lH)
            #End If
        Loop
    End With
End Sub

Public Sub CryptoSha2Finalize(uCtx As CryptoSha2Context, baOutput() As Byte)
    Static B(0 To 7)    As Long
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
        Call CopyMemory(P(lSize - 4), ByteSwap(B(0)), 4)
        Call CopyMemory(P(lSize - 8), ByteSwap(B(1)), 4)
        CryptoSha2Update uCtx, P, Size:=lSize
        Debug.Assert .NPartial = 0
        B(0) = ByteSwap(.H0): B(1) = ByteSwap(.H1): B(2) = ByteSwap(.H2): B(3) = ByteSwap(.H3)
        B(4) = ByteSwap(.H4): B(5) = ByteSwap(.H5): B(6) = ByteSwap(.H6): B(7) = ByteSwap(.H7)
        ReDim baOutput(0 To (.BitSize + 7) \ 8 - 1) As Byte
        Call CopyMemory(baOutput(0), B(0), UBound(baOutput) + 1)
    End With
End Sub

Public Function CryptoSha2ByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha2Context
    
    Select Case lBitSize
#If HasSha512 Then
    Case 512, 384, 512256, 512224
        CryptoSha2ByteArray = CryptoSha512ByteArray(lBitSize Mod 1000, baInput, Pos, Size)
#End If
    Case Else
        CryptoSha2Init uCtx, lBitSize
        CryptoSha2Update uCtx, baInput, Pos, Size
        CryptoSha2Finalize uCtx, CryptoSha2ByteArray
    End Select
End Function

Public Function CryptoSha2Text(ByVal lBitSize As Long, sText As String) As String
    Const CP_UTF8       As Long = 65001
    Dim uCtx            As CryptoSha2Context
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
    Select Case lBitSize
#If HasSha512 Then
    Case 512, 384, 512256, 512224
        baOutput = CryptoSha512ByteArray(lBitSize Mod 1000, baInput)
#End If
    Case Else
        CryptoSha2Init uCtx, lBitSize
        CryptoSha2Update uCtx, baInput, 0, lSize
        CryptoSha2Finalize uCtx, baOutput
    End Select
    ReDim aSplit(0 To UBound(baOutput)) As String
    For lSize = 0 To UBound(aSplit)
        aSplit(lSize) = Right$("0" & Hex$(baOutput(lSize)), 2)
    Next
    CryptoSha2Text = LCase$(Join(aSplit, vbNullString))
End Function

Public Function CryptoHmacSha2ByteArray(ByVal lBitSize As Long, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Const INNER_PAD     As Long = &H36
    Const OUTER_PAD     As Long = &H5C
    Dim lPadSize        As Long
    Dim lIdx            As Long
    Dim baPass()        As Byte
    Dim baPad()         As Byte
    Dim baHash()        As Byte
    
    lPadSize = IIf(lBitSize > 256, LNG_BLOCKSZ * 2, LNG_BLOCKSZ)
    If UBound(baKey) < lPadSize Then
        baPass = baKey
    Else
        baPass = CryptoSha2ByteArray(lBitSize, baKey)
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
    baHash = CryptoSha2ByteArray(lBitSize, baPad)
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
    CryptoHmacSha2ByteArray = CryptoSha2ByteArray(lBitSize, baPad)
End Function
