Attribute VB_Name = "mdSha512"
'--- mdSha512.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
#End If

Private Const LNG_BLOCKSZ               As Long = 128
Private Const LNG_ROUNDS                As Long = 80

Public Type CryptoSha512Context
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
    BitSize             As Long
End Type

#If HasPtrSafe Then
    Private LNG_K(0 To LNG_ROUNDS - 1) As LongLong
    Private LNG_POW2(0 To 63)       As LongLong
#Else
    Private LNG_K(0 To LNG_ROUNDS - 1) As Variant
    Private LNG_POW2(0 To 63)       As Variant
#End If

#If HasPtrSafe Then
Private Function RotR64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RotR64(lX As Variant, ByVal lN As Long) As Variant
#End If
    '--- RotR64 = RShift64(X, n) Or LShift64(X, 64 - n)
    Debug.Assert lN <> 0
    RotR64 = ((lX And (-1 Xor LNG_POW2(63))) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(63 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(64 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * LNG_POW2(63))
End Function

#If HasPtrSafe Then
Private Function LShift64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function LShift64(lX As Variant, ByVal lN As Long) As Variant
#End If
    If lN = 0 Then
        LShift64 = lX
    Else
        LShift64 = (lX And (LNG_POW2(63 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(63 - lN)) <> 0) * LNG_POW2(63)
    End If
End Function

#If HasPtrSafe Then
Private Function RShift64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RShift64(lX As Variant, ByVal lN As Long) As Variant
#End If
    If lN = 0 Then
        RShift64 = lX
    Else
        RShift64 = (lX And (-1 Xor LNG_POW2(63))) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(63 - lN)
    End If
End Function

#If HasPtrSafe Then
Private Function UAdd64(ByVal lX As LongLong, ByVal lY As LongLong) As LongLong
#Else
Private Function UAdd64(lX As Variant, lY As Variant) As Variant
#End If
    If (lX Xor lY) >= 0 Then
        UAdd64 = ((lX Xor LNG_POW2(63)) + lY) Xor LNG_POW2(63)
    Else
        UAdd64 = lX + lY
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
    BigSigma0 = RotR64(lX, 28) Xor RotR64(lX, 34) Xor RotR64(lX, 39)
End Function

#If HasPtrSafe Then
Private Function BigSigma1(ByVal lX As LongLong) As LongLong
#Else
Private Function BigSigma1(lX As Variant) As Variant
#End If
    BigSigma1 = RotR64(lX, 14) Xor RotR64(lX, 18) Xor RotR64(lX, 41)
End Function

#If HasPtrSafe Then
Private Function SmallSigma0(ByVal lX As LongLong) As LongLong
#Else
Private Function SmallSigma0(lX As Variant) As Variant
#End If
    SmallSigma0 = RotR64(lX, 1) Xor RotR64(lX, 8) Xor RShift64(lX, 7)
End Function

#If HasPtrSafe Then
Private Function SmallSigma1(ByVal lX As LongLong) As LongLong
#Else
Private Function SmallSigma1(lX As Variant) As Variant
#End If
    SmallSigma1 = RotR64(lX, 19) Xor RotR64(lX, 61) Xor RShift64(lX, 6)
End Function

Private Function BSwap32(ByVal lX As Long) As Long
    BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
                 (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
End Function

#If HasPtrSafe Then
Private Function BSwap64(ByVal lX As LongLong) As LongLong
#Else
Private Function BSwap64(ByVal lX As Variant) As Variant
#End If
    Dim lA As Long
    lA = BSwap32(CLng(lX And &H7FFFFFFF))
    BSwap64 = lA And &H7FFFFFFF Or -((lA < 0) <> 0) * LNG_POW2(31) Or -((lX And LNG_POW2(31)) <> 0) * &H80
End Function

#If Not HasPtrSafe Then
    Private Function CLngLng(vValue As Variant) As Variant
        Const VT_I8 As Long = &H14
        Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
    End Function
#End If

Public Sub CryptoSha512Init(uCtx As CryptoSha512Context, ByVal lBitSize As Long)
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim vSplit          As Variant
    
    If LNG_K(0) = 0 Then
        '--- K: first 64 bits of the fractional parts of the cube roots of the first 80 primes
        For Each vElem In Split("428A2F98D728AE22 7137449123EF65CD B5C0FBCFEC4D3B2F E9B5DBA58189DBBC 3956C25BF348B538 59F111F1B605D019 923F82A4AF194F9B AB1C5ED5DA6D8118 D807AA98A3030242 12835B0145706FBE 243185BE4EE4B28C 550C7DC3D5FFB4E2 72BE5D74F27B896F 80DEB1FE3B1696B1 9BDC06A725C71235 C19BF174CF692694 E49B69C19EF14AD2 EFBE4786384F25E3 0FC19DC68B8CD5B5 240CA1CC77AC9C65 2DE92C6F592B0275 4A7484AA6EA6E483 5CB0A9DCBD41FBD4 76F988DA831153B5 983E5152EE66DFAB A831C66D2DB43210 B00327C898FB213F BF597FC7BEEF0EE4 C6E00BF33DA88FC2 D5A79147930AA725 06CA6351E003826F 142929670A0E6E70 27B70A8546D22FFC 2E1B21385C26C926 4D2C6DFC5AC42AED 53380D139D95B3DF 650A73548BAF63DE 766A0ABB3C77B2A8 81C2C92E47EDAEE6 92722C851482353B " & _
                                "A2BFE8A14CF10364 A81A664BBC423001 C24B8B70D0F89791 C76C51A30654BE30 D192E819D6EF5218 D69906245565A910 F40E35855771202A 106AA07032BBD1B8 19A4C116B8D2D0C8 1E376C085141AB53 2748774CDF8EEB99 34B0BCB5E19B48A8 391C0CB3C5C95A63 4ED8AA4AE3418ACB 5B9CCA4F7763E373 682E6FF3D6B2B8A3 748F82EE5DEFB2FC 78A5636F43172F60 84C87814A1F0AB72 8CC702081A6439EC 90BEFFFA23631E28 A4506CEBDE82BDE9 BEF9A3F7B2C67915 C67178F2E372532B CA273ECEEA26619C D186B8C721C0C207 EADA7DD6CDE0EB1E F57D4F7FEE6ED178 06F067AA72176FBA 0A637DC5A2C898A6 113F9804BEF90DAE 1B710B35131C471B 28DB77F523047D84 32CAAB7B40C72493 3C9EBE0A15C9BEBC 431D67C49C100D4C 4CC5D4BECB3E42B6 597F299CFC657E2A 5FCB6FAB3AD6FAEC 6C44198C4A475817")
            LNG_K(lIdx) = CLngLng(CStr("&H" & vElem))
            lIdx = lIdx + 1
        Next
        LNG_POW2(0) = CLngLng(1)
        For lIdx = 1 To 63
            LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
        Next
    End If
    With uCtx
        Select Case lBitSize Mod 1000
        Case 224
            vSplit = Split("8C3D37C819544DA2 73E1996689DCD4D6 1DFAB7AE32FF9C82 679DD514582F9FCF F6D2B697BD44DA8 77E36F7304C48942 3F9D85A86A1D36C8 1112E6AD91D692A1")
        Case 256
            vSplit = Split("22312194FC2BF72C 9F555FA3C84C64C2 2393B86B6F53B151 963877195940EABD 96283EE2A88EFFE3 BE5E1E2553863992 2B0199FC2C85B8AA EB72DDC81C52CA2")
        Case 384
            vSplit = Split("CBBB9D5DC1059ED8 629A292A367CD507 9159015A3070DD17 152FECD8F70E5939 67332667FFC00B31 8EB44A8768581511 DB0C2E0D64F98FA7 47B5481DBEFA4FA4")
        Case 512
            vSplit = Split("6A09E667F3BCC908 BB67AE8584CAA73B 3C6EF372FE94F82B A54FF53A5F1D36F1 510E527FADE682D1 9B05688C2B3E6C1F 1F83D9ABFB41BD6B 5BE0CD19137E2179")
        Case Else
            Err.Raise vbObjectError, , "Invalid bit-size for SHA-512 (" & lBitSize & ")"
        End Select
        .H0 = CLngLng(CStr("&H" & vSplit(0))): .H1 = CLngLng(CStr("&H" & vSplit(1))): .H2 = CLngLng(CStr("&H" & vSplit(2))): .H3 = CLngLng(CStr("&H" & vSplit(3)))
        .H4 = CLngLng(CStr("&H" & vSplit(4))): .H5 = CLngLng(CStr("&H" & vSplit(5))): .H6 = CLngLng(CStr("&H" & vSplit(6))): .H7 = CLngLng(CStr("&H" & vSplit(7)))
        .NPartial = 0
        .NInput = 0
        .BitSize = lBitSize
    End With
End Sub

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Public Sub CryptoSha512Update(uCtx As CryptoSha512Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Static B(0 To 31)   As Long
#If HasPtrSafe Then
    Static W(0 To LNG_ROUNDS - 1) As LongLong
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
    Static W(0 To LNG_ROUNDS - 1) As Variant
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
            '--- sha512 step
            lA = .H0: lB = .H1: lC = .H2: lD = .H3
            lE = .H4: lF = .H5: lG = .H6: lH = .H7
            For lIdx = 0 To LNG_ROUNDS - 1
                If lIdx < 16 Then
                    W(lIdx) = BSwap64(CLngLng(B(lIdx * 2 + 1))) Or LShift64(BSwap64(CLngLng(B(lIdx * 2))), 32)
                Else
                    W(lIdx) = UAdd64(UAdd64(UAdd64(SmallSigma1(W(lIdx - 2)), W(lIdx - 7)), SmallSigma0(W(lIdx - 15))), W(lIdx - 16))
                End If
                lT1 = UAdd64(UAdd64(UAdd64(UAdd64(lH, BigSigma1(lE)), Ch(lE, lF, lG)), LNG_K(lIdx)), W(lIdx))
                lT2 = UAdd64(BigSigma0(lA), Maj(lA, lB, lC))
                lH = lG
                lG = lF
                lF = lE
                lE = UAdd64(lD, lT1)
                lD = lC
                lC = lB
                lB = lA
                lA = UAdd64(lT1, lT2)
            Next
            .H0 = UAdd64(.H0, lA): .H1 = UAdd64(.H1, lB): .H2 = UAdd64(.H2, lC): .H3 = UAdd64(.H3, lD)
            .H4 = UAdd64(.H4, lE): .H5 = UAdd64(.H5, lF): .H6 = UAdd64(.H6, lG): .H7 = UAdd64(.H7, lH)
        Loop
    End With
End Sub

#If HasPtrSafe Then
Private Function pvToLong(ByVal lX As LongLong, lHi As Long, lLo As Long) As Long
    Dim lA              As LongLong
#Else
Private Function pvToLong(ByVal lX As Variant, lHi As Long, lLo As Long) As Long
    Dim lA              As Variant
#End If
    lA = BSwap64(RShift64(lX, 32))
    lHi = CLng(lA And &H7FFFFFFF) Or -((lA And LNG_POW2(31)) <> 0) * &H80000000
    lA = BSwap64(lX)
    lLo = CLng(lA And &H7FFFFFFF) Or -((lA And LNG_POW2(31)) <> 0) * &H80000000
End Function

Public Sub CryptoSha512Finalize(uCtx As CryptoSha512Context, baOutput() As Byte)
    Static B(0 To 15)   As Long
    Dim P(0 To LNG_BLOCKSZ + 17) As Byte
    Dim lSize           As Long
    
    With uCtx
        lSize = LNG_BLOCKSZ - .NPartial
        If lSize < 17 Then
            lSize = lSize + LNG_BLOCKSZ
        End If
        P(0) = &H80
        .NInput = .NInput / 10000@ * 8
        Call CopyMemory(B(0), .NInput, 8)
        Call CopyMemory(P(lSize - 4), BSwap32(B(0)), 4)
        Call CopyMemory(P(lSize - 8), BSwap32(B(1)), 4)
        CryptoSha512Update uCtx, P, Size:=lSize
        Debug.Assert .NPartial = 0
        pvToLong .H0, B(0), B(1)
        pvToLong .H1, B(2), B(3)
        pvToLong .H2, B(4), B(5)
        pvToLong .H3, B(6), B(7)
        pvToLong .H4, B(8), B(9)
        pvToLong .H5, B(10), B(11)
        pvToLong .H6, B(12), B(13)
        pvToLong .H7, B(14), B(15)
        ReDim baOutput(0 To (.BitSize + 7) \ 8 - 1) As Byte
        Call CopyMemory(baOutput(0), B(0), UBound(baOutput) + 1)
    End With
End Sub

Public Function CryptoSha512ByteArray(ByVal lBitSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSha512Context
    
    CryptoSha512Init uCtx, lBitSize
    CryptoSha512Update uCtx, baInput, Pos, Size
    CryptoSha512Finalize uCtx, CryptoSha512ByteArray
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

Public Function CryptoSha512Text(ByVal lBitSize As Long, sText As String) As String
    CryptoSha512Text = ToHex(CryptoSha512ByteArray(lBitSize, ToUtf8Array(sText)))
End Function
