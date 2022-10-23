Attribute VB_Name = "mdSha2Old"
'--- mdSha2.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

#If Not HasOperators Then
Private LNG_POW2(0 To 31)           As Long

Private Function ROTR32(ByVal lX As Long, ByVal lN As Long) As Long
    '--- ROTR32 = RShift(X, n) Or LShift(X, 32 - n)
    Debug.Assert lN <> 0
    ROTR32 = ((lX And &H7FFFFFFF) \ LNG_POW2(lN) - (lX < 0) * LNG_POW2(31 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(32 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * &H80000000)
End Function

Private Function LShift(ByVal lX As Long, ByVal lN As Long) As Long
    If lN = 0 Then
        LShift = lX
    Else
        LShift = (lX And (LNG_POW2(31 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(31 - lN)) <> 0) * &H80000000
    End If
End Function

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

Private Function bSwap(ByVal lX As Long) As Long
    bSwap = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or (lX And &H7F000000) \ &H1000000 Or _
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

Private Sub ToBigEndian(aRetVal() As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Dim lIdx            As Long
    Dim lOutSize        As Long
    
    If lSize < 0 Then
        lSize = UBound(baBuffer) + 1 - lPos
    End If
    lOutSize = ((lSize + 8) \ 64 + 1) * 16
    ReDim aRetVal(0 To lOutSize - 1) As Long
    If lSize > 0 Then
        Call CopyMemory(aRetVal(0), baBuffer(0), lSize)
    End If
    Call CopyMemory(ByVal VarPtr(aRetVal(0)) + lSize, &H80, 1)
    #If HasOperators Then
        For lIdx = 0 To lOutSize - 3
            Dim lX As Long = aRetVal(lIdx)
            aRetVal(lIdx) = (lX >> 24) Or (lX >> 8) And &HFF00& Or (lX << 8) And &HFF0000 Or (lX << 24)
        Next
        aRetVal(lOutSize - 1) = (lSize << 3)
        aRetVal(lOutSize - 2) = (lSize >> 29)
    #Else
        For lIdx = 0 To lOutSize - 3
            aRetVal(lIdx) = bSwap(aRetVal(lIdx))
        Next
        aRetVal(lOutSize - 1) = LShift(lSize, 3)
        aRetVal(lOutSize - 2) = RShift(lSize, 29)
    #End If
End Sub

Private Sub FromBigEndian(baRetVal() As Byte, aInput() As Long, ByVal lPos As Long, ByVal lSize As Long)
    Dim lIdx            As Long
    Dim lWord           As Long
    
    If lSize < 0 Then
        lSize = UBound(aInput) + 1 - lPos
    End If
    ReDim baRetVal(0 To lSize * 4 - 1) As Byte
    For lIdx = 0 To lSize - 1
        lWord = aInput(lPos + lIdx)
        #If HasOperators Then
            baRetVal(4 * lIdx + 0) = (lWord >> 24)
            baRetVal(4 * lIdx + 1) = (lWord >> 16) And &HFF&
            baRetVal(4 * lIdx + 2) = (lWord >> 8) And &HFF&
            baRetVal(4 * lIdx + 3) = lWord And &HFF&
        #Else
            baRetVal(4 * lIdx + 0) = RShift(lWord, 24) And &HFF&
            baRetVal(4 * lIdx + 1) = (lWord And &HFF0000) \ &H10000 And &HFF&
            baRetVal(4 * lIdx + 2) = (lWord And &HFF00) \ &H100& And &HFF&
            baRetVal(4 * lIdx + 3) = lWord And &HFF&
        #End If
    Next
End Sub

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Private Sub CalcSha2(baOutput() As Byte, ByVal lOutPos As Long, ByVal lOutSize As Long, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long, H() As Long)
    Static K(0 To 63)   As Long
    Static W(0 To 63)   As Long
    Dim M()             As Long
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
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim vElem           As Variant
    Dim lX              As Long
    Dim lSigma1         As Long
    Dim lSigma0         As Long
    Dim lCh             As Long
    Dim lMaj            As Long
    
    If K(0) = 0 Then
        #If Not HasOperators Then
            For lIdx = 0 To 30
                LNG_POW2(lIdx) = 2& ^ lIdx
            Next
            LNG_POW2(31) = &H80000000
        #End If
        '--- K: first 32 bits of the fractional parts of the cube roots of the first 64 primes
        For Each vElem In Split("428A2F98 71374491 B5C0FBCF E9B5DBA5 3956C25B 59F111F1 923F82A4 AB1C5ED5 D807AA98 12835B01 243185BE 550C7DC3 72BE5D74 80DEB1FE 9BDC06A7 C19BF174 E49B69C1 EFBE4786 FC19DC6 240CA1CC 2DE92C6F 4A7484AA 5CB0A9DC 76F988DA 983E5152 A831C66D B00327C8 BF597FC7 C6E00BF3 D5A79147 6CA6351 14292967 27B70A85 2E1B2138 4D2C6DFC 53380D13 650A7354 766A0ABB 81C2C92E 92722C85 A2BFE8A1 A81A664B C24B8B70 C76C51A3 D192E819 D6990624 F40E3585 106AA070 19A4C116 1E376C08 2748774C 34B0BCB5 391C0CB3 4ED8AA4A 5B9CCA4F 682E6FF3 748F82EE 78A5636F 84C87814 8CC70208 90BEFFFA A4506CEB BEF9A3F7 C67178F2")
            K(lJdx) = "&H" & vElem
            lJdx = lJdx + 1
        Next
    End If
    ToBigEndian M, baInput, lPos, lSize
    For lIdx = 0 To UBound(M) Step 16
        lA = H(0): lB = H(1): lC = H(2): lD = H(3)
        lE = H(4): lF = H(5): lG = H(6): lH = H(7)
        For lJdx = 0 To 63
            If lJdx < 16 Then
                W(lJdx) = M(lJdx + lIdx)
            Else
                #If HasOperators Then
                    lX = W(lJdx - 2)
                    lSigma1 = (lX >> 17 Or lX << 15) Xor (lX >> 19 Or lX << 13) Xor (lX >> 10)
                    lX = W(lJdx - 15)
                    lSigma0 = (lX >> 7 Or lX << 25) Xor (lX >> 18 Or lX << 14) Xor (lX >> 3)
                    W(lJdx) = lSigma1 + W(lJdx - 7) + lSigma0 + W(lJdx - 16)
                #Else
                    W(lJdx) = UAdd(UAdd(UAdd(SmallSigma1(W(lJdx - 2)), W(lJdx - 7)), SmallSigma0(W(lJdx - 15))), W(lJdx - 16))
                #End If
            End If
            #If HasOperators Then
                lSigma1 = (lE >> 6 Or lE << 26) Xor (lE >> 11 Or lE << 21) Xor (lE >> 25 Or lE << 7)
                lSigma0 = (lA >> 2 Or lA << 30) Xor (lA >> 13 Or lA << 19) Xor (lA >> 22 Or lA << 10)
                lCh = (lE And lF) Xor ((Not lE) And lG)
                lMaj = (lA And lB) Xor (lA And lC) Xor (lB And lC)
                lT1 = lH + lSigma1 + lCh + K(lJdx) + W(lJdx)
                lT2 = lSigma0 + lMaj
            #Else
                lT1 = UAdd(UAdd(UAdd(UAdd(lH, BigSigma1(lE)), Ch(lE, lF, lG)), K(lJdx)), W(lJdx))
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
            H(0) += lA: H(1) += lB: H(2) += lC: H(3) += lD
            H(4) += lE: H(5) += lF: H(6) += lG: H(7) += lH
        #Else
            H(0) = UAdd(lA, H(0)): H(1) = UAdd(lB, H(1)): H(2) = UAdd(lC, H(2)): H(3) = UAdd(lD, H(3))
            H(4) = UAdd(lE, H(4)): H(5) = UAdd(lF, H(5)): H(6) = UAdd(lG, H(6)): H(7) = UAdd(lH, H(7))
        #End If
    Next
    FromBigEndian baOutput, H, lOutPos, (lOutSize + 3) \ 4
End Sub

Public Sub CryptoSha2(ByVal lBitSize As Long, baOutput() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim H(0 To 7)       As Long

    Select Case lBitSize
#If HasSha512 Then
    Case 512, 384, 512256, 512224
        CryptoSha512 lBitSize Mod 1000, baOutput, baInput, Pos, Size
        Exit Sub
#End If
    Case 224
        H(0) = &HC1059ED8
        H(1) = &H367CD507
        H(2) = &H3070DD17
        H(3) = &HF70E5939
        H(4) = &HFFC00B31
        H(5) = &H68581511
        H(6) = &H64F98FA7
        H(7) = &HBEFA4FA4
    Case 256
        H(0) = &H6A09E667
        H(1) = &HBB67AE85
        H(2) = &H3C6EF372
        H(3) = &HA54FF53A
        H(4) = &H510E527F
        H(5) = &H9B05688C
        H(6) = &H1F83D9AB
        H(7) = &H5BE0CD19
    Case Else
        Err.Raise vbObjectError, , "Invalid bit-size for SHA-2 (" & lBitSize & ")"
    End Select
    CalcSha2 baOutput, 0, (lBitSize + 7) \ 8, baInput, Pos, Size, H
End Sub

Public Sub CryptoHmacSha2(ByVal lBitSize As Long, baOutput() As Byte, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Const INNER_PAD     As Long = &H36
    Const OUTER_PAD     As Long = &H5C
    Dim lPadSize        As Long
    Dim lIdx            As Long
    Dim baPass()        As Byte
    Dim baPad()         As Byte
    Dim baHash()        As Byte
    
    lPadSize = IIf(lBitSize > 256, 128, 64)
    If UBound(baKey) < lPadSize Then
        baPass = baKey
    Else
        CryptoSha2 lBitSize, baPass, baKey
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
    CryptoSha2 lBitSize, baHash, baPad
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
    CryptoSha2 lBitSize, baOutput, baPad
End Sub

Public Function CryptoSha2ByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    CryptoSha2 lBitSize, CryptoSha2ByteArray, baInput, Pos, Size
End Function

Public Function CryptoSha2Text(ByVal lBitSize As Long, sText As String) As String

End Function

Public Function CryptoHmacSha2ByteArray(ByVal lBitSize As Long, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    CryptoHmacSha2 lBitSize, CryptoHmacSha2ByteArray, baKey, baInput, Pos, Size
End Function
