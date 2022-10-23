Attribute VB_Name = "mdSha512Old"
'--- mdSha512.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0) Or (TWINBASIC <> 0)

#If Not HasPtrSafe Then
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
#End If

#If HasPtrSafe Then
    Private LNG_POW2(0 To 63)       As LongLong
#Else
    Private LNG_POW2(0 To 63)       As Variant
#End If

#If HasPtrSafe Then
Private Function ROTR64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function ROTR64(lX As Variant, ByVal lN As Long) As Variant
#End If
    '--- ROTR64 = RShift(X, n) Or LShift(X, 64 - n)
    Debug.Assert lN <> 0
    ROTR64 = ((lX And (-1 Xor LNG_POW2(63))) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(63 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(64 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * LNG_POW2(63))
End Function

#If HasPtrSafe Then
Private Function LShift(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function LShift(lX As Variant, ByVal lN As Long) As Variant
#End If
    If lN = 0 Then
        LShift = lX
    Else
        LShift = (lX And (LNG_POW2(63 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(63 - lN)) <> 0) * LNG_POW2(63)
    End If
End Function

#If HasPtrSafe Then
Private Function RShift(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RShift(lX As Variant, ByVal lN As Long) As Variant
#End If
    If lN = 0 Then
        RShift = lX
    Else
        RShift = (lX And (-1 Xor LNG_POW2(63))) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(63 - lN)
    End If
End Function

#If HasPtrSafe Then
Private Function UAdd(ByVal lX As LongLong, ByVal lY As LongLong) As LongLong
#Else
Private Function UAdd(lX As Variant, lY As Variant) As Variant
#End If
    If (lX Xor lY) > 0 Then
        UAdd = ((lX Xor LNG_POW2(63)) + lY) Xor LNG_POW2(63)
    Else
        UAdd = lX + lY
    End If
End Function

#If HasPtrSafe Then
Private Function Ch(ByVal lX As LongLong, ByVal lY As LongLong, ByVal lZ As LongLong) As LongLong
#Else
Private Function Ch(lX As Variant, lY As Variant, ByVal lZ As Variant) As Variant
#End If
    Ch = (lX And lY) Xor ((Not lX) And lZ)
End Function

#If HasPtrSafe Then
Private Function Maj(ByVal lX As LongLong, ByVal lY As LongLong, ByVal lZ As LongLong) As LongLong
#Else
Private Function Maj(lX As Variant, lY As Variant, lZ As Variant) As Variant
#End If
    Maj = (lX And lY) Xor (lX And lZ) Xor (lY And lZ)
End Function

#If HasPtrSafe Then
Private Function BigSigma0(ByVal lX As LongLong) As LongLong
#Else
Private Function BigSigma0(lX As Variant) As Variant
#End If
    BigSigma0 = ROTR64(lX, 28) Xor ROTR64(lX, 34) Xor ROTR64(lX, 39)
End Function

#If HasPtrSafe Then
Private Function BigSigma1(ByVal lX As LongLong) As LongLong
#Else
Private Function BigSigma1(lX As Variant) As Variant
#End If
    BigSigma1 = ROTR64(lX, 14) Xor ROTR64(lX, 18) Xor ROTR64(lX, 41)
End Function

#If HasPtrSafe Then
Private Function SmallSigma0(ByVal lX As LongLong) As LongLong
#Else
Private Function SmallSigma0(lX As Variant) As Variant
#End If
    SmallSigma0 = ROTR64(lX, 1) Xor ROTR64(lX, 8) Xor RShift(lX, 7)
End Function

#If HasPtrSafe Then
Private Function SmallSigma1(ByVal lX As LongLong) As LongLong
#Else
Private Function SmallSigma1(lX As Variant) As Variant
#End If
    SmallSigma1 = ROTR64(lX, 19) Xor ROTR64(lX, 61) Xor RShift(lX, 6)
End Function

#If HasPtrSafe Then
Private Sub ToBigEndian(aRetVal() As LongLong, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long)
#Else
Private Sub ToBigEndian(aRetVal() As Variant, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long)
#End If
    Dim lIdx            As Long
    Dim lOutSize        As Long
    Dim lOutIdx         As Long
    Dim lOffset         As Long
    
    If lSize < 0 Then
        lSize = UBound(baBuffer) + 1 - lPos
    End If
    lOutSize = ((lSize + 16) \ 128 + 1) * 16
    ReDim aRetVal(0 To lOutSize - 1)
    For lIdx = 0 To lSize - 1
        lOutIdx = lIdx \ 8
        lOffset = 56 - (lIdx Mod 8) * 8
        aRetVal(lOutIdx) = aRetVal(lOutIdx) Or LShift(baBuffer(lPos + lIdx), lOffset)
    Next
    lOutIdx = lIdx \ 8
    lOffset = 56 - (lIdx Mod 8) * 8
    aRetVal(lOutIdx) = aRetVal(lOutIdx) Or LShift(&H80, lOffset)
    aRetVal(lOutSize - 1) = LShift(lSize, 3)
End Sub

#If HasPtrSafe Then
Private Sub FromBigEndian(baRetVal() As Byte, aInput() As LongLong, ByVal lPos As Long, ByVal lSize As Long)
    Dim lWord           As LongLong
#Else
Private Sub FromBigEndian(baRetVal() As Byte, aInput() As Variant, ByVal lPos As Long, ByVal lSize As Long)
    Dim lWord           As Variant
#End If
    Dim lIdx            As Long
    
    If lSize < 0 Then
        lSize = UBound(aInput) + 1 - lPos
    End If
    ReDim baRetVal(0 To lSize * 8 - 1) As Byte
    For lIdx = 0 To lSize - 1
        lWord = aInput(lPos + lIdx)
        baRetVal(8 * lIdx + 0) = CByte(RShift(lWord, 56) And &HFF&)
        baRetVal(8 * lIdx + 1) = CByte(RShift(lWord, 48) And &HFF&)
        baRetVal(8 * lIdx + 2) = CByte(RShift(lWord, 40) And &HFF&)
        baRetVal(8 * lIdx + 3) = CByte(RShift(lWord, 32) And &HFF&)
        baRetVal(8 * lIdx + 4) = CByte(RShift(lWord, 24) And &HFF&)
        baRetVal(8 * lIdx + 5) = CByte((lWord And &HFF0000) \ &H10000 And &HFF&)
        baRetVal(8 * lIdx + 6) = CByte((lWord And &HFF00) \ &H100& And &HFF&)
        baRetVal(8 * lIdx + 7) = CByte(lWord And &HFF&)
    Next
End Sub

#If Not HasPtrSafe Then
    Private Function CLngLng(vValue As Variant) As Variant
        Const VT_I8 As Long = &H14
        Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
    End Function
#End If

#If HasPtrSafe Then
Private Sub CalcSha512(baOutput() As Byte, ByVal lOutPos As Long, ByVal lOutSize As Long, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long, H() As LongLong)
    Static K(0 To 79)   As LongLong
    Static W(0 To 79)   As LongLong
    Dim M()             As LongLong
    Dim lA              As LongLong
    Dim lB              As LongLong
    Dim lC              As LongLong
    Dim lD              As LongLong
    Dim lE              As LongLong
    Dim lF              As LongLong
    Dim lG              As LongLong
    Dim lH              As LongLong
    Dim lT1             As LongLong
    Dim lT2             As LongLong
#Else
Private Sub CalcSha512(baOutput() As Byte, ByVal lOutPos As Long, ByVal lOutSize As Long, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long, H() As Variant)
    Static K(0 To 79)   As Variant
    Static W(0 To 79)   As Variant
    Dim M()             As Variant
    Dim lA              As Variant
    Dim lB              As Variant
    Dim lC              As Variant
    Dim lD              As Variant
    Dim lE              As Variant
    Dim lF              As Variant
    Dim lG              As Variant
    Dim lH              As Variant
    Dim lT1             As Variant
    Dim lT2             As Variant
#End If
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim vElem           As Variant
    
    If LNG_POW2(0) = 0 Then
        LNG_POW2(0) = CLngLng(1)
        For lIdx = 1 To 63
            LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
        Next
        '--- K: first 64 bits of the fractional parts of the cube roots of the first 80 primes
        For Each vElem In Split("428A2F98D728AE22 7137449123EF65CD B5C0FBCFEC4D3B2F E9B5DBA58189DBBC 3956C25BF348B538 59F111F1B605D019 923F82A4AF194F9B AB1C5ED5DA6D8118 D807AA98A3030242 12835B0145706FBE 243185BE4EE4B28C 550C7DC3D5FFB4E2 72BE5D74F27B896F 80DEB1FE3B1696B1 9BDC06A725C71235 C19BF174CF692694 E49B69C19EF14AD2 EFBE4786384F25E3 0FC19DC68B8CD5B5 240CA1CC77AC9C65 2DE92C6F592B0275 4A7484AA6EA6E483 5CB0A9DCBD41FBD4 76F988DA831153B5 983E5152EE66DFAB A831C66D2DB43210 B00327C898FB213F BF597FC7BEEF0EE4 C6E00BF33DA88FC2 D5A79147930AA725 06CA6351E003826F 142929670A0E6E70 27B70A8546D22FFC 2E1B21385C26C926 4D2C6DFC5AC42AED 53380D139D95B3DF 650A73548BAF63DE 766A0ABB3C77B2A8 81C2C92E47EDAEE6 92722C851482353B " & _
                                "A2BFE8A14CF10364 A81A664BBC423001 C24B8B70D0F89791 C76C51A30654BE30 D192E819D6EF5218 D69906245565A910 F40E35855771202A 106AA07032BBD1B8 19A4C116B8D2D0C8 1E376C085141AB53 2748774CDF8EEB99 34B0BCB5E19B48A8 391C0CB3C5C95A63 4ED8AA4AE3418ACB 5B9CCA4F7763E373 682E6FF3D6B2B8A3 748F82EE5DEFB2FC 78A5636F43172F60 84C87814A1F0AB72 8CC702081A6439EC 90BEFFFA23631E28 A4506CEBDE82BDE9 BEF9A3F7B2C67915 C67178F2E372532B CA273ECEEA26619C D186B8C721C0C207 EADA7DD6CDE0EB1E F57D4F7FEE6ED178 06F067AA72176FBA 0A637DC5A2C898A6 113F9804BEF90DAE 1B710B35131C471B 28DB77F523047D84 32CAAB7B40C72493 3C9EBE0A15C9BEBC 431D67C49C100D4C 4CC5D4BECB3E42B6 597F299CFC657E2A 5FCB6FAB3AD6FAEC 6C44198C4A475817")
            K(lJdx) = CLngLng(CStr("&H" & vElem))
            lJdx = lJdx + 1
        Next
    End If
    ToBigEndian M, baInput, lPos, lSize
    For lIdx = 0 To UBound(M) Step 16
        lA = H(0)
        lB = H(1)
        lC = H(2)
        lD = H(3)
        lE = H(4)
        lF = H(5)
        lG = H(6)
        lH = H(7)
        For lJdx = 0 To 79
            If lJdx < 16 Then
                W(lJdx) = M(lJdx + lIdx)
            Else
                W(lJdx) = UAdd(UAdd(UAdd(SmallSigma1(W(lJdx - 2)), W(lJdx - 7)), SmallSigma0(W(lJdx - 15))), W(lJdx - 16))
            End If
'            Debug.Print "W(" & lJdx + lIdx & ")=" & W(lJdx)
            lT1 = UAdd(UAdd(UAdd(UAdd(lH, BigSigma1(lE)), Ch(lE, lF, lG)), K(lJdx)), W(lJdx))
            lT2 = UAdd(BigSigma0(lA), Maj(lA, lB, lC))
            lH = lG
            lG = lF
            lF = lE
            lE = UAdd(lD, lT1)
            lD = lC
            lC = lB
            lB = lA
            lA = UAdd(lT1, lT2)
        Next
        H(0) = UAdd(lA, H(0))
        H(1) = UAdd(lB, H(1))
        H(2) = UAdd(lC, H(2))
        H(3) = UAdd(lD, H(3))
        H(4) = UAdd(lE, H(4))
        H(5) = UAdd(lF, H(5))
        H(6) = UAdd(lG, H(6))
        H(7) = UAdd(lH, H(7))
    Next
    FromBigEndian baOutput, H, lOutPos, (lOutSize + 7) \ 8
    If UBound(baOutput) <> lOutSize - 1 Then
        ReDim Preserve baOutput(0 To lOutSize - 1) As Byte
    End If
End Sub

Public Sub CryptoSha512(ByVal lBitSize As Long, baOutput() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
#If HasPtrSafe Then
    Dim H(0 To 7)       As LongLong
#Else
    Dim H(0 To 7)       As Variant
#End If
    Dim sIV             As String
    Dim vElem           As Variant
    Dim lIdx            As Long

    Select Case lBitSize
    Case 224
        sIV = "8C3D37C819544DA2 73E1996689DCD4D6 1DFAB7AE32FF9C82 679DD514582F9FCF F6D2B697BD44DA8 77E36F7304C48942 3F9D85A86A1D36C8 1112E6AD91D692A1"
    Case 256
        sIV = "22312194FC2BF72C 9F555FA3C84C64C2 2393B86B6F53B151 963877195940EABD 96283EE2A88EFFE3 BE5E1E2553863992 2B0199FC2C85B8AA EB72DDC81C52CA2"
    Case 384
        sIV = "CBBB9D5DC1059ED8 629A292A367CD507 9159015A3070DD17 152FECD8F70E5939 67332667FFC00B31 8EB44A8768581511 DB0C2E0D64F98FA7 47B5481DBEFA4FA4"
    Case 512
        sIV = "6A09E667F3BCC908 BB67AE8584CAA73B 3C6EF372FE94F82B A54FF53A5F1D36F1 510E527FADE682D1 9B05688C2B3E6C1F 1F83D9ABFB41BD6B 5BE0CD19137E2179"
    Case Else
        Err.Raise vbObjectError, , "Invalid bit-size for SHA-512 (" & lBitSize & ")"
    End Select
    For Each vElem In Split(sIV)
        H(lIdx) = CLngLng(CStr("&H" & vElem))
        #If HasPtrSafe Then
            Debug.Assert Hex$(H(lIdx)) = vElem
        #End If
        lIdx = lIdx + 1
    Next
    CalcSha512 baOutput, 0, (lBitSize + 7) \ 8, baInput, Pos, Size, H
End Sub

