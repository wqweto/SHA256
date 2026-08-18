Attribute VB_Name = "mdSha512Sliced"
'=========================================================================
'
'  Pure VB6 Crypto (Untested)
'  Copyright (c) 2026 wqweto@gmail.com
'
'  This project is licensed under the terms of the MIT license
'  See the LICENSE file in the project root for more information
'
'=========================================================================
'--- mdSha512Sliced.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_STATESZ               As Long = 64
Private Const LNG_BLOCKSZ               As Long = 128
Private Const LNG_ROUNDS                As Long = 80
Private Const LNG_POW2_1                As Long = 2 ^ 1
Private Const LNG_POW2_2                As Long = 2 ^ 2
Private Const LNG_POW2_3                As Long = 2 ^ 3
Private Const LNG_POW2_4                As Long = 2 ^ 4
Private Const LNG_POW2_5                As Long = 2 ^ 5
Private Const LNG_POW2_6                As Long = 2 ^ 6
Private Const LNG_POW2_7                As Long = 2 ^ 7
Private Const LNG_POW2_8                As Long = 2 ^ 8
Private Const LNG_POW2_9                As Long = 2 ^ 9
Private Const LNG_POW2_12               As Long = 2 ^ 12
Private Const LNG_POW2_13               As Long = 2 ^ 13
Private Const LNG_POW2_14               As Long = 2 ^ 14
Private Const LNG_POW2_17               As Long = 2 ^ 17
Private Const LNG_POW2_18               As Long = 2 ^ 18
Private Const LNG_POW2_19               As Long = 2 ^ 19
Private Const LNG_POW2_22               As Long = 2 ^ 22
Private Const LNG_POW2_23               As Long = 2 ^ 23
Private Const LNG_POW2_24               As Long = 2 ^ 24
Private Const LNG_POW2_25               As Long = 2 ^ 25
Private Const LNG_POW2_26               As Long = 2 ^ 26
Private Const LNG_POW2_27               As Long = 2 ^ 27
Private Const LNG_POW2_28               As Long = 2 ^ 28
Private Const LNG_POW2_29               As Long = 2 ^ 29
Private Const LNG_POW2_30               As Long = 2 ^ 30
Private Const LNG_POW2_31               As Long = &H80000000

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    cElements           As Long
    lLbound             As Long
End Type

Private Type ArrayLong16
    Item(0 To LNG_STATESZ \ 4 - 1) As Long
End Type

Private Type ArrayLong32
    Item(0 To LNG_BLOCKSZ \ 4 - 1) As Long
End Type

Public Type CryptoSha512Context
    State               As ArrayLong16
    Block               As ArrayLong32
    Bytes()             As Byte                 '--- overlaying Block or State arrays above
    ArrayBytes          As SAFEARRAY1D
    NPartial            As Long
    NInput              As Currency
    BitSize             As Long
End Type

Private LNG_K(0 To 2 * LNG_ROUNDS - 1) As Long
Private m_bNoIntegerOverflowChecks As Boolean

Private Function BSwap32(ByVal lX As Long) As Long
    #If Not HasOperators Then
        BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
                  (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
    #Else
        Return ((lX And &H000000FF&) << 24) Or _
               ((lX And &H0000FF00&) << 8) Or _
               ((lX And &H00FF0000&) >> 8) Or _
               ((lX And &HFF000000&) >> 24)
    #End If
End Function

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Private Sub pvAdd64(lAL As Long, lAH As Long, ByVal lBL As Long, ByVal lBH As Long)
    Dim lSign           As Long
    
    #If Not HasOperators Then
        If m_bNoIntegerOverflowChecks Then
            lAL = lAL + lBL
            lAH = lAH + lBH
            If (lAL And &H80000000) <> 0 Then
                lSign = 1
            Else
                lSign = 0
            End If
            If (lBL And &H80000000) <> 0 Then
                lSign = lSign - 1
            End If
            Select Case True
            Case lSign < 0, lSign = 0 And (lAL And &H7FFFFFFF) < (lBL And &H7FFFFFFF)
                lAH = lAH + 1
            End Select
        Else
            If (lAL Xor lBL) >= 0 Then
                lAL = ((lAL Xor &H80000000) + lBL) Xor &H80000000
            Else
                lAL = lAL + lBL
            End If
            If (lAH Xor lBH) >= 0 Then
                lAH = ((lAH Xor &H80000000) + lBH) Xor &H80000000
            Else
                lAH = lAH + lBH
            End If
            If (lAL And &H80000000) <> 0 Then
                lSign = 1
            End If
            If (lBL And &H80000000) <> 0 Then
                lSign = lSign - 1
            End If
            Select Case True
            Case lSign < 0, lSign = 0 And (lAL And &H7FFFFFFF) < (lBL And &H7FFFFFFF)
                If lAH >= 0 Then
                    lAH = ((lAH Xor &H80000000) + 1) Xor &H80000000
                Else
                    lAH = lAH + 1
                End If
            End Select
        End If
    #Else
        lAL += lBL
        lAH += lBH
        lSign = (lAL >> 31) - (lBL >> 31)
        If lSign < 0 Or lSign = 0 And (lAL And &H7FFFFFFF) < (lBL And &H7FFFFFFF) Then
            lAH += 1
        End If
    #End If
End Sub

Private Function pvSum0L(ByVal lX As Long, ByVal lY As Long) As Long
    #If Not HasOperators Then
        pvSum0L = ((lX And (LNG_POW2_6 - 1)) * LNG_POW2_25 Or -((lX And LNG_POW2_6) <> 0) * &H80000000) _
            Xor ((lX And (LNG_POW2_1 - 1)) * LNG_POW2_30 Or -((lX And LNG_POW2_1) <> 0) * &H80000000) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_28 Or -(lX < 0) * LNG_POW2_3) _
            Xor ((lY And &H7FFFFFFF) \ LNG_POW2_7 Or -(lY < 0) * LNG_POW2_24) _
            Xor ((lY And &H7FFFFFFF) \ LNG_POW2_2 Or -(lY < 0) * LNG_POW2_29) _
            Xor ((lY And (LNG_POW2_27 - 1)) * LNG_POW2_4 Or -((lY And LNG_POW2_27) <> 0) * &H80000000)
    #Else
        Return (lX << 25) Xor (lX << 30) Xor (lX >> 28) Xor (lY >> 7) Xor (lY >> 2) Xor (lY << 4)
    #End If
End Function

Private Function pvSum1L(ByVal lX As Long, ByVal lY As Long) As Long
    #If Not HasOperators Then
        pvSum1L = ((lX And (LNG_POW2_8 - 1)) * LNG_POW2_23 Or -((lX And LNG_POW2_8) <> 0) * &H80000000) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_14 Or -(lX < 0) * LNG_POW2_17) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_18 Or -(lX < 0) * LNG_POW2_13) _
            Xor ((lY And &H7FFFFFFF) \ LNG_POW2_9 Or -(lY < 0) * LNG_POW2_22) _
            Xor ((lY And (LNG_POW2_13 - 1)) * LNG_POW2_18 Or -((lY And LNG_POW2_13) <> 0) * &H80000000) _
            Xor ((lY And (LNG_POW2_17 - 1)) * LNG_POW2_14 Or -((lY And LNG_POW2_17) <> 0) * &H80000000)
    #Else
        Return (lX << 23) Xor (lX >> 14) Xor (lX >> 18) Xor (lY >> 9) Xor (lY << 18) Xor (lY << 14)
    #End If
End Function

Private Function pvSig0L(ByVal lX As Long, ByVal lY As Long) As Long
    #If Not HasOperators Then
        pvSig0L = ((lX And &H7FFFFFFF) \ LNG_POW2_1 Or -(lX < 0) * LNG_POW2_30) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_7 Or -(lX < 0) * LNG_POW2_24) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_8 Or -(lX < 0) * LNG_POW2_23) _
            Xor ((lY And 0) * LNG_POW2_31 Or -((lY And 1) <> 0) * &H80000000) _
            Xor ((lY And (LNG_POW2_6 - 1)) * LNG_POW2_25 Or -((lY And LNG_POW2_6) <> 0) * &H80000000) _
            Xor ((lY And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lY And LNG_POW2_7) <> 0) * &H80000000)
    #Else
        Return (lX >> 1) Xor (lX >> 7) Xor (lX >> 8) Xor (lY << 31) Xor (lY << 25) Xor (lY << 24)
    #End If
End Function
  
Private Function pvSig0H(ByVal lX As Long, ByVal lY As Long) As Long
    #If Not HasOperators Then
        pvSig0H = ((lX And &H7FFFFFFF) \ LNG_POW2_1 Or -(lX < 0) * LNG_POW2_30) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_7 Or -(lX < 0) * LNG_POW2_24) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_8 Or -(lX < 0) * LNG_POW2_23) _
            Xor ((lY And 0) * LNG_POW2_31 Or -((lY And 1) <> 0) * &H80000000) _
            Xor ((lY And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lY And LNG_POW2_7) <> 0) * &H80000000)
    #Else
        Return (lX >> 1) Xor (lX >> 7) Xor (lX >> 8) Xor (lY << 31) Xor (lY << 24)
    #End If
End Function

Private Function pvSig1L(ByVal lX As Long, ByVal lY As Long) As Long
    #If Not HasOperators Then
        pvSig1L = ((lX And (LNG_POW2_28 - 1)) * LNG_POW2_3 Or -((lX And LNG_POW2_28) <> 0) * &H80000000) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_6 Or -(lX < 0) * LNG_POW2_25) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_19 Or -(lX < 0) * LNG_POW2_12) _
            Xor ((lY And &H7FFFFFFF) \ LNG_POW2_29 Or -(lY < 0) * LNG_POW2_2) _
            Xor ((lY And (LNG_POW2_5 - 1)) * LNG_POW2_26 Or -((lY And LNG_POW2_5) <> 0) * &H80000000) _
            Xor ((lY And (LNG_POW2_18 - 1)) * LNG_POW2_13 Or -((lY And LNG_POW2_18) <> 0) * &H80000000)
    #Else
        Return (lX << 3) Xor (lX >> 6) Xor (lX >> 19) Xor (lY >> 29) Xor (lY << 26) Xor (lY << 13)
    #End If
End Function

Private Function pvSig1H(ByVal lX As Long, ByVal lY As Long) As Long
    #If Not HasOperators Then
        pvSig1H = ((lX And (LNG_POW2_28 - 1)) * LNG_POW2_3 Or -((lX And LNG_POW2_28) <> 0) * &H80000000) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_6 Or -(lX < 0) * LNG_POW2_25) _
            Xor ((lX And &H7FFFFFFF) \ LNG_POW2_19 Or -(lX < 0) * LNG_POW2_12) _
            Xor ((lY And &H7FFFFFFF) \ LNG_POW2_29 Or -(lY < 0) * LNG_POW2_2) _
            Xor ((lY And (LNG_POW2_18 - 1)) * LNG_POW2_13 Or -((lY And LNG_POW2_18) <> 0) * &H80000000)
    #Else
        Return (lX << 3) Xor (lX >> 6) Xor (lX >> 19) Xor (lY >> 29) Xor (lY << 13)
    #End If
End Function

Private Sub pvRound( _
            ByVal lX00 As Long, ByVal lX01 As Long, ByVal lX02 As Long, ByVal lX03 As Long, ByVal lX04 As Long, ByVal lX05 As Long, lX06 As Long, lX07 As Long, _
            ByVal lX08 As Long, ByVal lX09 As Long, ByVal lX10 As Long, ByVal lX11 As Long, ByVal lX12 As Long, ByVal lX13 As Long, lX14 As Long, lX15 As Long, _
            uArray As ArrayLong32, ByVal lIdx As Long, ByVal lJdx As Long)
    pvAdd64 lX14, lX15, uArray.Item(lIdx), uArray.Item(lIdx + 1)
    pvAdd64 lX14, lX15, LNG_K(lJdx + lIdx), LNG_K(lJdx + lIdx + 1)
    pvAdd64 lX14, lX15, lX12 Xor (lX08 And (lX10 Xor lX12)), lX13 Xor (lX09 And (lX11 Xor lX13))
    pvAdd64 lX14, lX15, pvSum1L(lX08, lX09), pvSum1L(lX09, lX08)
    pvAdd64 lX06, lX07, lX14, lX15
    pvAdd64 lX14, lX15, pvSum0L(lX00, lX01), pvSum0L(lX01, lX00)
    pvAdd64 lX14, lX15, ((lX00 Or lX04) And lX02) Or (lX04 And lX00), ((lX01 Or lX05) And lX03) Or (lX05 And lX01)
End Sub

Private Sub pvStore(uArray As ArrayLong32, ByVal lIdx As Long)
    Dim lTL             As Long
    Dim lTH             As Long
    Dim lUL             As Long
    Dim lUH             As Long
    
    With uArray
        lTL = .Item(lIdx)
        lTH = .Item(lIdx + 1)
        pvAdd64 lTL, lTH, .Item((lIdx + 18) And &H1F), .Item((lIdx + 19) And &H1F)
        lUL = pvSig0L(.Item((lIdx + 2) And &H1F), .Item((lIdx + 3) And &H1F))
        lUH = pvSig0H(.Item((lIdx + 3) And &H1F), .Item((lIdx + 2) And &H1F))
        pvAdd64 lTL, lTH, lUL, lUH
        lUL = pvSig1L(.Item((lIdx + 28) And &H1F), .Item((lIdx + 29) And &H1F))
        lUH = pvSig1H(.Item((lIdx + 29) And &H1F), .Item((lIdx + 28) And &H1F))
        pvAdd64 lTL, lTH, lUL, lUH
        .Item(lIdx) = lTL
        .Item(lIdx + 1) = lTH
    End With
End Sub

Private Function pvGetOverflowIgnored(Optional bValue As Boolean = True) As Boolean
    Dim bInIde      As Boolean
    
    If Not bValue Then
        bValue = True
        pvGetOverflowIgnored = True
        Exit Function
    End If
    Debug.Assert pvGetOverflowIgnored(bInIde)
    If bInIde Then
        Exit Function
    End If
    On Error GoTo EH
    If &H8000 - 1 <> 0 Then
        pvGetOverflowIgnored = True
    End If
EH:
End Function

Public Sub CryptoSha512Init(uCtx As CryptoSha512Context, ByVal lBitSize As Long)
    Const FADF_AUTO     As Long = 1
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim pDummy          As LongPtr
    
    If LNG_K(0) = 0 Then
        '--- K: first 64 bits of the fractional parts of the cube roots of the first 80 primes
        For Each vElem In Split("D728AE22 428A2F98 23EF65CD 71374491 EC4D3B2F B5C0FBCF 8189DBBC E9B5DBA5 F348B538 3956C25B B605D019 59F111F1 AF194F9B 923F82A4 DA6D8118 AB1C5ED5 A3030242 D807AA98 45706FBE 12835B01 4EE4B28C 243185BE D5FFB4E2 550C7DC3 F27B896F 72BE5D74 3B1696B1 80DEB1FE 25C71235 9BDC06A7 CF692694 C19BF174 9EF14AD2 E49B69C1 384F25E3 EFBE4786 8B8CD5B5 0FC19DC6 77AC9C65 240CA1CC 592B0275 2DE92C6F 6EA6E483 4A7484AA BD41FBD4 5CB0A9DC 831153B5 76F988DA EE66DFAB 983E5152 2DB43210 A831C66D 98FB213F B00327C8 BEEF0EE4 BF597FC7 3DA88FC2 C6E00BF3 930AA725 D5A79147 E003826F 06CA6351 0A0E6E70 14292967 46D22FFC 27B70A85 5C26C926 2E1B2138 5AC42AED 4D2C6DFC 9D95B3DF 53380D13 8BAF63DE 650A7354 3C77B2A8 766A0ABB 47EDAEE6 81C2C92E 1482353B 92722C85 " & _
                                "4CF10364 A2BFE8A1 BC423001 A81A664B D0F89791 C24B8B70 0654BE30 C76C51A3 D6EF5218 D192E819 5565A910 D6990624 5771202A F40E3585 32BBD1B8 106AA070 B8D2D0C8 19A4C116 5141AB53 1E376C08 DF8EEB99 2748774C E19B48A8 34B0BCB5 C5C95A63 391C0CB3 E3418ACB 4ED8AA4A 7763E373 5B9CCA4F D6B2B8A3 682E6FF3 5DEFB2FC 748F82EE 43172F60 78A5636F A1F0AB72 84C87814 1A6439EC 8CC70208 23631E28 90BEFFFA DE82BDE9 A4506CEB B2C67915 BEF9A3F7 E372532B C67178F2 EA26619C CA273ECE 21C0C207 D186B8C7 CDE0EB1E EADA7DD6 EE6ED178 F57D4F7F 72176FBA 06F067AA A2C898A6 0A637DC5 BEF90DAE 113F9804 131C471B 1B710B35 23047D84 28DB77F5 40C72493 32CAAB7B 15C9BEBC 3C9EBE0A 9C100D4C 431D67C4 CB3E42B6 4CC5D4BE FC657E2A 597F299C 3AD6FAEC 5FCB6FAB 4A475817 6C44198C")
            LNG_K(lIdx) = "&H" & vElem
            lIdx = lIdx + 1
        Next
        m_bNoIntegerOverflowChecks = pvGetOverflowIgnored
    End If
    With uCtx
        Select Case lBitSize Mod 1000
        Case 224
            vSplit = Split("19544DA2 8C3D37C8 89DCD4D6 73E19966 32FF9C82 1DFAB7AE 582F9FCF 679DD514 7BD44DA8 F6D2B69 04C48942 77E36F73 6A1D36C8 3F9D85A8 91D692A1 1112E6AD")
        Case 256
            vSplit = Split("FC2BF72C 22312194 C84C64C2 9F555FA3 6F53B151 2393B86B 5940EABD 96387719 A88EFFE3 96283EE2 53863992 BE5E1E25 2C85B8AA 2B0199FC 81C52CA2 EB72DDC")
        Case 384
            vSplit = Split("C1059ED8 CBBB9D5D 367CD507 629A292A 3070DD17 9159015A F70E5939 152FECD8 FFC00B31 67332667 68581511 8EB44A87 64F98FA7 DB0C2E0D BEFA4FA4 47B5481D")
        Case 512
            vSplit = Split("F3BCC908 6A09E667 84CAA73B BB67AE85 FE94F82B 3C6EF372 5F1D36F1 A54FF53A ADE682D1 510E527F 2B3E6C1F 9B05688C FB41BD6B 1F83D9AB 137E2179 5BE0CD19")
        Case Else
            Err.Raise vbObjectError, , "Invalid bit-size for SHA-512 (" & lBitSize & ")"
        End Select
        lIdx = 0
        For Each vElem In vSplit
            .State.Item(lIdx) = "&H" & vElem
            lIdx = lIdx + 1
        Next
        .NPartial = 0
        .NInput = 0
        .BitSize = lBitSize
        With .ArrayBytes
            .cDims = 1
            .fFeatures = FADF_AUTO
            .cbElements = 1
            .cLocks = 1
            .pvData = VarPtr(uCtx.Block.Item(0))
            .cElements = LNG_BLOCKSZ \ .cbElements
        End With
        Call CopyMemory(ByVal ArrPtr(.Bytes), VarPtr(.ArrayBytes), LenB(pDummy))
    End With
End Sub

Public Sub CryptoSha512Update(uCtx As CryptoSha512Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lAL             As Long
    Dim lAH             As Long
    Dim lBL             As Long
    Dim lBH             As Long
    Dim lCL             As Long
    Dim lCh             As Long
    Dim lDL             As Long
    Dim lDH             As Long
    Dim lEL             As Long
    Dim lEH             As Long
    Dim lFL             As Long
    Dim lFH             As Long
    Dim lGL             As Long
    Dim lGH             As Long
    Dim lHL             As Long
    Dim lHH             As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        .NInput = .NInput + Size
        If .NPartial > 0 And Size > 0 Then
            lIdx = LNG_BLOCKSZ - .NPartial
            If lIdx > Size Then
                lIdx = Size
            End If
            Call CopyMemory(.Bytes(.NPartial), baInput(Pos), lIdx)
            .NPartial = .NPartial + lIdx
            Pos = Pos + lIdx
            Size = Size - lIdx
        End If
        Do While Size > 0 Or .NPartial = LNG_BLOCKSZ
            If .NPartial <> 0 Then
                .NPartial = 0
            ElseIf Size >= LNG_BLOCKSZ Then
                Call CopyMemory(.Bytes(0), baInput(Pos), LNG_BLOCKSZ)
                Pos = Pos + LNG_BLOCKSZ
                Size = Size - LNG_BLOCKSZ
            Else
                Call CopyMemory(.Bytes(0), baInput(Pos), Size)
                .NPartial = Size
                Exit Do
            End If
            '--- sha512 step
            For lIdx = 0 To UBound(.Block.Item) Step 2
                lAL = BSwap32(.Block.Item(lIdx))
                .Block.Item(lIdx) = BSwap32(.Block.Item(lIdx + 1))
                .Block.Item(lIdx + 1) = lAL
            Next
            lAL = .State.Item(0): lAH = .State.Item(1)
            lBL = .State.Item(2): lBH = .State.Item(3)
            lCL = .State.Item(4): lCh = .State.Item(5)
            lDL = .State.Item(6): lDH = .State.Item(7)
            lEL = .State.Item(8): lEH = .State.Item(9)
            lFL = .State.Item(10): lFH = .State.Item(11)
            lGL = .State.Item(12): lGH = .State.Item(13)
            lHL = .State.Item(14): lHH = .State.Item(15)
            lIdx = 0
            Do While lIdx < 2 * LNG_ROUNDS
                lJdx = 0
                Do While lJdx < LNG_BLOCKSZ \ 4
                    pvRound lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, .Block, lJdx + 0, lIdx
                    pvRound lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, .Block, lJdx + 2, lIdx
                    pvRound lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, .Block, lJdx + 4, lIdx
                    pvRound lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, .Block, lJdx + 6, lIdx
                    pvRound lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, .Block, lJdx + 8, lIdx
                    pvRound lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, .Block, lJdx + 10, lIdx
                    pvRound lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, .Block, lJdx + 12, lIdx
                    pvRound lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, .Block, lJdx + 14, lIdx
                    lJdx = lJdx + 16
                Loop
                lIdx = lIdx + 32
                If lIdx >= 2 * LNG_ROUNDS Then
                    Exit Do
                End If
                For lJdx = 0 To 30 Step 2
                    pvStore .Block, lJdx
                Next
            Loop
            pvAdd64 .State.Item(0), .State.Item(1), lAL, lAH
            pvAdd64 .State.Item(2), .State.Item(3), lBL, lBH
            pvAdd64 .State.Item(4), .State.Item(5), lCL, lCh
            pvAdd64 .State.Item(6), .State.Item(7), lDL, lDH
            pvAdd64 .State.Item(8), .State.Item(9), lEL, lEH
            pvAdd64 .State.Item(10), .State.Item(11), lFL, lFH
            pvAdd64 .State.Item(12), .State.Item(13), lGL, lGH
            pvAdd64 .State.Item(14), .State.Item(15), lHL, lHH
        Loop
    End With
End Sub

Public Sub CryptoSha512Finalize(uCtx As CryptoSha512Context, baOutput() As Byte)
    Static B(0 To 1)    As Long
    Dim baPad()         As Byte
    Dim lIdx            As Long
    Dim pDummy          As LongPtr
    
    With uCtx
        lIdx = LNG_BLOCKSZ - .NPartial
        If lIdx < 17 Then
            lIdx = lIdx + LNG_BLOCKSZ
        End If
        ReDim baPad(0 To lIdx - 1) As Byte
        baPad(0) = &H80
        .NInput = .NInput / 10000@ * 8
        Call CopyMemory(B(0), .NInput, 8)
        Call CopyMemory(baPad(lIdx - 4), BSwap32(B(0)), 4)
        Call CopyMemory(baPad(lIdx - 8), BSwap32(B(1)), 4)
        CryptoSha512Update uCtx, baPad
        Debug.Assert .NPartial = 0
        ReDim baOutput(0 To (.BitSize + 7) \ 8 - 1) As Byte
        .ArrayBytes.pvData = VarPtr(.State.Item(0))
        For lIdx = 0 To UBound(baOutput)
            baOutput(lIdx) = .Bytes(lIdx + 7 - 2 * (lIdx And 7))
        Next
        Call CopyMemory(ByVal ArrPtr(.Bytes), pDummy, LenB(pDummy))
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

Public Function CryptoSha512Text(ByVal lBitSize As Long, sText As String) As String
    CryptoSha512Text = ToHex(CryptoSha512ByteArray(lBitSize, ToUtf8Array(sText)))
End Function
