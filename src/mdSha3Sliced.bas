Attribute VB_Name = "mdSha3Sliced"
'--- mdSha3Sliced.bas
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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_STATESZ               As Long = 200
Private Const LNG_ROUNDS                As Long = 24
Private Const LNG_POW2_1                As Long = 2 ^ 1
Private Const LNG_POW2_2                As Long = 2 ^ 2
Private Const LNG_POW2_3                As Long = 2 ^ 3
Private Const LNG_POW2_4                As Long = 2 ^ 4
Private Const LNG_POW2_5                As Long = 2 ^ 5
Private Const LNG_POW2_6                As Long = 2 ^ 6
Private Const LNG_POW2_7                As Long = 2 ^ 7
Private Const LNG_POW2_8                As Long = 2 ^ 8
Private Const LNG_POW2_9                As Long = 2 ^ 9
Private Const LNG_POW2_10               As Long = 2 ^ 10
Private Const LNG_POW2_11               As Long = 2 ^ 11
Private Const LNG_POW2_12               As Long = 2 ^ 12
Private Const LNG_POW2_13               As Long = 2 ^ 13
Private Const LNG_POW2_14               As Long = 2 ^ 14
Private Const LNG_POW2_15               As Long = 2 ^ 15
Private Const LNG_POW2_16               As Long = 2 ^ 16
Private Const LNG_POW2_17               As Long = 2 ^ 17
Private Const LNG_POW2_18               As Long = 2 ^ 18
Private Const LNG_POW2_19               As Long = 2 ^ 19
Private Const LNG_POW2_20               As Long = 2 ^ 20
Private Const LNG_POW2_21               As Long = 2 ^ 21
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

Private Type ArrayLong50
    Item(0 To LNG_STATESZ \ 4 - 1) As Long
End Type

Public Type CryptoSha3Context
    State               As ArrayLong50
    Bytes()             As Byte                     '--- overlaying State array above
    PeekArray           As SAFEARRAY1D
    DigestSize          As Long
    Capacity            As Long
    Absorbed            As Long
End Type

Private LNG_RC(0 To 2 * LNG_ROUNDS - 1)     As Long

Private Function pvSeparate(ByVal lX As Long) As Long
    Dim lTemp           As Long
    
    #If Not HasOperators Then
        lTemp = (((lX And &H7FFFFFFF) \ LNG_POW2_1 Or -(lX < 0) * LNG_POW2_30) Xor lX) And &H22222222
        lX = (lX Xor lTemp) Xor ((lTemp And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lTemp And LNG_POW2_30) <> 0) * &H80000000)
        lTemp = (((lX And &H7FFFFFFF) \ LNG_POW2_2 Or -(lX < 0) * LNG_POW2_29) Xor lX) And &HC0C0C0C
        lX = (lX Xor lTemp) Xor ((lTemp And (LNG_POW2_29 - 1)) * LNG_POW2_2 Or -((lTemp And LNG_POW2_29) <> 0) * &H80000000)
        lTemp = (((lX And &H7FFFFFFF) \ LNG_POW2_4 Or -(lX < 0) * LNG_POW2_27) Xor lX) And &HF000F0
        lX = (lX Xor lTemp) Xor ((lTemp And (LNG_POW2_27 - 1)) * LNG_POW2_4 Or -((lTemp And LNG_POW2_27) <> 0) * &H80000000)
        lTemp = (((lX And &H7FFFFFFF) \ LNG_POW2_8 Or -(lX < 0) * LNG_POW2_23) Xor lX) And &HFF00&
        pvSeparate = (lX Xor lTemp) Xor ((lTemp And (LNG_POW2_23 - 1)) * LNG_POW2_8 Or -((lTemp And LNG_POW2_23) <> 0) * &H80000000)
    #Else
        lTemp = ((lX >> 1) Xor lX) And &H22222222
        lX = (lX Xor lTemp) Xor (lTemp << 1)
        lTemp = ((lX >> 2) Xor lX) And &HC0C0C0C
        lX = (lX Xor lTemp) Xor (lTemp << 2)
        lTemp = ((lX >> 4) Xor lX) And &HF000F0
        lX = (lX Xor lTemp) Xor (lTemp << 4)
        lTemp = ((lX >> 8) Xor lX) And &HFF00&
        pvSeparate = (lX Xor lTemp) Xor (lTemp << 8)
    #End If
End Function

Private Function pvCombine(ByVal lX As Long) As Long
    Dim lTemp           As Long
    
    #If Not HasOperators Then
        lTemp = (((lX And &H7FFFFFFF) \ LNG_POW2_15 Or -(lX < 0) * LNG_POW2_16) Xor lX) And &HAAAA&
        lX = (lX Xor lTemp) Xor ((lTemp And (LNG_POW2_16 - 1)) * LNG_POW2_15 Or -((lTemp And LNG_POW2_16) <> 0) * &H80000000)
        lTemp = (((lX And &H7FFFFFFF) \ LNG_POW2_14 Or -(lX < 0) * LNG_POW2_17) Xor lX) And &HCCCC&
        lX = (lX Xor lTemp) Xor ((lTemp And (LNG_POW2_17 - 1)) * LNG_POW2_14 Or -((lTemp And LNG_POW2_17) <> 0) * &H80000000)
        lTemp = (((lX And &H7FFFFFFF) \ LNG_POW2_12 Or -(lX < 0) * LNG_POW2_19) Xor lX) And &HF0F0&
        lX = (lX Xor lTemp) Xor ((lTemp And (LNG_POW2_19 - 1)) * LNG_POW2_12 Or -((lTemp And LNG_POW2_19) <> 0) * &H80000000)
        lTemp = (((lX And &H7FFFFFFF) \ LNG_POW2_8 Or -(lX < 0) * LNG_POW2_23) Xor lX) And &HFF00&
        pvCombine = (lX Xor lTemp) Xor ((lTemp And (LNG_POW2_23 - 1)) * LNG_POW2_8 Or -((lTemp And LNG_POW2_23) <> 0) * &H80000000)
    #Else
        lTemp = ((lX >> 15) Xor lX) And &HAAAA&
        lX = (lX Xor lTemp) Xor (lTemp << 15)
        lTemp = ((lX >> 14) Xor lX) And &HCCCC&
        lX = (lX Xor lTemp) Xor (lTemp << 14)
        lTemp = ((lX >> 12) Xor lX) And &HF0F0&
        lX = (lX Xor lTemp) Xor (lTemp << 12)
        lTemp = ((lX >> 8) Xor lX) And &HFF00&
        pvCombine = (lX Xor lTemp) Xor (lTemp << 8)
    #End If
End Function

Private Sub pvToSliced(uState As ArrayLong50)
    Dim lIdx            As Long
    Dim lT0             As Long
    Dim lT1             As Long
    
    With uState
    For lIdx = 0 To UBound(.Item) Step 2
        lT0 = pvSeparate(.Item(lIdx))
        lT1 = pvSeparate(.Item(lIdx + 1))
        #If Not HasOperators Then
            .Item(lIdx) = (lT0 And &HFFFF&) Or ((lT1 And (LNG_POW2_15 - 1)) * LNG_POW2_16 Or -((lT1 And LNG_POW2_15) <> 0) * &H80000000)
            .Item(lIdx + 1) = (lT1 And &HFFFF0000) Or ((lT0 And &H7FFFFFFF) \ LNG_POW2_16 Or -(lT0 < 0) * LNG_POW2_15)
        #Else
            .Item(lIdx) = (lT0 And &HFFFF&) Or (lT1 << 16)
            .Item(lIdx + 1) = (lT1 And &HFFFF0000) Or (lT0 >> 16)
        #End If
    Next
    End With
End Sub

Private Sub pvFromSliced(uState As ArrayLong50)
    Dim lIdx            As Long
    Dim lT0             As Long
    Dim lT1             As Long
    Dim lX              As Long
    
    With uState
    For lIdx = 0 To UBound(.Item) Step 2
        lT0 = pvCombine(.Item(lIdx))
        lT1 = pvCombine(.Item(lIdx + 1))
        #If Not HasOperators Then
            lX = lT1 And &H55555555
            .Item(lIdx) = ((lX And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lX And LNG_POW2_30) <> 0) * &H80000000) Or (lT0 And &H55555555)
            lX = lT0 And &HAAAAAAAA
            .Item(lIdx + 1) = ((lX And &H7FFFFFFF) \ LNG_POW2_1 Or -(lX < 0) * LNG_POW2_30) Or (lT1 And &HAAAAAAAA)
        #Else
            .Item(lIdx) = ((lT1 And &H55555555) << 1) Or (lT0 And &H55555555)
            .Item(lIdx + 1) = ((lT0 And &HAAAAAAAA) >> 1) Or (lT1 And &HAAAAAAAA)
        #End If
    Next
    End With
End Sub

Private Sub Keccak(uState As ArrayLong50)
    Dim lT0             As Long
    Dim lT1             As Long
    Dim lT2             As Long
    Dim lT3             As Long
    Dim lT4             As Long
    Dim lT5             As Long
    Dim lT6             As Long
    Dim lT7             As Long
    Dim lT8             As Long
    Dim lT9             As Long
    Dim lU0             As Long
    Dim lU1             As Long
    Dim lU2             As Long
    Dim lU3             As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    pvToSliced uState
    With uState
        lU0 = .Item(40)
        lU1 = .Item(41)
        lT2 = .Item(42)
        lT3 = .Item(43)
        lT4 = .Item(44)
        lT5 = .Item(45)
        lT6 = .Item(46)
        lT7 = .Item(47)
        lT8 = .Item(48)
        lT9 = .Item(49)
        
        For lIdx = 0 To 2 * LNG_ROUNDS - 1 Step 2
            '--- Theta
            For lJdx = 0 To UBound(.Item) - 10 Step 10
                lU0 = lU0 Xor .Item(lJdx)
                lU1 = lU1 Xor .Item(lJdx + 1)
                lT2 = lT2 Xor .Item(lJdx + 2)
                lT3 = lT3 Xor .Item(lJdx + 3)
                lT4 = lT4 Xor .Item(lJdx + 4)
                lT5 = lT5 Xor .Item(lJdx + 5)
                lT6 = lT6 Xor .Item(lJdx + 6)
                lT7 = lT7 Xor .Item(lJdx + 7)
                lT8 = lT8 Xor .Item(lJdx + 8)
                lT9 = lT9 Xor .Item(lJdx + 9)
            Next
            #If Not HasOperators Then
                lT0 = ((lT5 And &H7FFFFFFF) \ LNG_POW2_31 - (lT5 < 0)) Or _
                    ((lT5 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lT5 And LNG_POW2_30) <> 0) * &H80000000) Xor lU0
                lT1 = lU1 Xor lT4
                lT4 = ((lT9 And &H7FFFFFFF) \ LNG_POW2_31 - (lT9 < 0)) Or _
                    ((lT9 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lT9 And LNG_POW2_30) <> 0) * &H80000000) Xor lT4
                lT5 = lT5 Xor lT8
                lT8 = ((lT3 And &H7FFFFFFF) \ LNG_POW2_31 - (lT3 < 0)) Or _
                    ((lT3 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lT3 And LNG_POW2_30) <> 0) * &H80000000) Xor lT8
                lT9 = lT9 Xor lT2
                lT2 = ((lT7 And &H7FFFFFFF) \ LNG_POW2_31 - (lT7 < 0)) Or _
                    ((lT7 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lT7 And LNG_POW2_30) <> 0) * &H80000000) Xor lT2
                lT3 = lT3 Xor lT6
                lT6 = ((lU1 And &H7FFFFFFF) \ LNG_POW2_31 - (lU1 < 0)) Or _
                    ((lU1 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lU1 And LNG_POW2_30) <> 0) * &H80000000) Xor lT6
                lT7 = lT7 Xor lU0
            #Else
                lT0 = lU0 Xor (lT5 >> 31 Or lT5 << 1)
                lT1 = lU1 Xor lT4
                lT4 = lT4 Xor (lT9 >> 31 Or lT9 << 1)
                lT5 = lT5 Xor lT8
                lT8 = lT8 Xor (lT3 >> 31 Or lT3 << 1)
                lT9 = lT9 Xor lT2
                lT2 = lT2 Xor (lT7 >> 31 Or lT7 << 1)
                lT3 = lT3 Xor lT6
                lT6 = lT6 Xor (lU1 >> 31 Or lU1 << 1)
                lT7 = lT7 Xor lU0
            #End If
            
            '--- (Theta) Rho Pi
            lU0 = .Item(0) Xor lT8
            lU1 = .Item(1) Xor lT9
            .Item(0) = lU0
            .Item(1) = lU1
            lU2 = .Item(2) Xor lT0
            lU3 = .Item(3) Xor lT1
            lU0 = .Item(12) Xor lT0
            lU1 = .Item(13) Xor lT1
            #If Not HasOperators Then
                .Item(2) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_10 - (lU0 < 0) * LNG_POW2_21) Or _
                    ((lU0 And (LNG_POW2_9 - 1)) * LNG_POW2_22 Or -((lU0 And LNG_POW2_9) <> 0) * &H80000000)
                .Item(3) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_10 - (lU1 < 0) * LNG_POW2_21) Or _
                    ((lU1 And (LNG_POW2_9 - 1)) * LNG_POW2_22 Or -((lU1 And LNG_POW2_9) <> 0) * &H80000000)
            #Else
                .Item(2) = (lU0 >> 10 Or lU0 << 22)
                .Item(3) = (lU1 >> 10 Or lU1 << 22)
            #End If
            lU0 = .Item(18) Xor lT6
            lU1 = .Item(19) Xor lT7
            #If Not HasOperators Then
                .Item(12) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_22 - (lU0 < 0) * LNG_POW2_9) Or _
                    ((lU0 And (LNG_POW2_21 - 1)) * LNG_POW2_10 Or -((lU0 And LNG_POW2_21) <> 0) * &H80000000)
                .Item(13) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_22 - (lU1 < 0) * LNG_POW2_9) Or _
                    ((lU1 And (LNG_POW2_21 - 1)) * LNG_POW2_10 Or -((lU1 And LNG_POW2_21) <> 0) * &H80000000)
            #Else
                .Item(12) = (lU0 >> 22 Or lU0 << 10)
                .Item(13) = (lU1 >> 22 Or lU1 << 10)
            #End If
            lU0 = .Item(44) Xor lT2
            lU1 = .Item(45) Xor lT3
            #If Not HasOperators Then
                .Item(18) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_1 - (lU1 < 0) * LNG_POW2_30) Or _
                    ((lU1 And 0) * LNG_POW2_31 Or -((lU1 And 1) <> 0) * &H80000000)
                .Item(19) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_2 - (lU0 < 0) * LNG_POW2_29) Or _
                    ((lU0 And (LNG_POW2_1 - 1)) * LNG_POW2_30 Or -((lU0 And LNG_POW2_1) <> 0) * &H80000000)
            #Else
                .Item(18) = (lU1 >> 1 Or lU1 << 31)
                .Item(19) = (lU0 >> 2 Or lU0 << 30)
            #End If
            lU0 = .Item(28) Xor lT6
            lU1 = .Item(29) Xor lT7
            #If Not HasOperators Then
                .Item(44) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_12 - (lU1 < 0) * LNG_POW2_19) Or _
                    ((lU1 And (LNG_POW2_11 - 1)) * LNG_POW2_20 Or -((lU1 And LNG_POW2_11) <> 0) * &H80000000)
                .Item(45) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_13 - (lU0 < 0) * LNG_POW2_18) Or _
                    ((lU0 And (LNG_POW2_12 - 1)) * LNG_POW2_19 Or -((lU0 And LNG_POW2_12) <> 0) * &H80000000)
            #Else
                .Item(44) = (lU1 >> 12 Or lU1 << 20)
                .Item(45) = (lU0 >> 13 Or lU0 << 19)
            #End If
            lU0 = .Item(40) Xor lT8
            lU1 = .Item(41) Xor lT9
            #If Not HasOperators Then
                .Item(28) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_23 - (lU0 < 0) * LNG_POW2_8) Or _
                    ((lU0 And (LNG_POW2_22 - 1)) * LNG_POW2_9 Or -((lU0 And LNG_POW2_22) <> 0) * &H80000000)
                .Item(29) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_23 - (lU1 < 0) * LNG_POW2_8) Or _
                    ((lU1 And (LNG_POW2_22 - 1)) * LNG_POW2_9 Or -((lU1 And LNG_POW2_22) <> 0) * &H80000000)
            #Else
                .Item(28) = (lU0 >> 23 Or lU0 << 9)
                .Item(29) = (lU1 >> 23 Or lU1 << 9)
            #End If
            lU0 = .Item(4) Xor lT2
            lU1 = .Item(5) Xor lT3
            #If Not HasOperators Then
                .Item(40) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_1 - (lU0 < 0) * LNG_POW2_30) Or _
                    ((lU0 And 0) * LNG_POW2_31 Or -((lU0 And 1) <> 0) * &H80000000)
                .Item(41) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_1 - (lU1 < 0) * LNG_POW2_30) Or _
                    ((lU1 And 0) * LNG_POW2_31 Or -((lU1 And 1) <> 0) * &H80000000)
            #Else
                .Item(40) = (lU0 >> 1 Or lU0 << 31)
                .Item(41) = (lU1 >> 1 Or lU1 << 31)
            #End If
            lU0 = .Item(24) Xor lT2
            lU1 = .Item(25) Xor lT3
            #If Not HasOperators Then
                .Item(4) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_10 - (lU1 < 0) * LNG_POW2_21) Or _
                    ((lU1 And (LNG_POW2_9 - 1)) * LNG_POW2_22 Or -((lU1 And LNG_POW2_9) <> 0) * &H80000000)
                .Item(5) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_11 - (lU0 < 0) * LNG_POW2_20) Or _
                    ((lU0 And (LNG_POW2_10 - 1)) * LNG_POW2_21 Or -((lU0 And LNG_POW2_10) <> 0) * &H80000000)
            #Else
                .Item(4) = (lU1 >> 10 Or lU1 << 22)
                .Item(5) = (lU0 >> 11 Or lU0 << 21)
            #End If
            lU0 = .Item(26) Xor lT4
            lU1 = .Item(27) Xor lT5
            #If Not HasOperators Then
                .Item(24) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_19 - (lU1 < 0) * LNG_POW2_12) Or _
                    ((lU1 And (LNG_POW2_18 - 1)) * LNG_POW2_13 Or -((lU1 And LNG_POW2_18) <> 0) * &H80000000)
                .Item(25) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_20 - (lU0 < 0) * LNG_POW2_11) Or _
                    ((lU0 And (LNG_POW2_19 - 1)) * LNG_POW2_12 Or -((lU0 And LNG_POW2_19) <> 0) * &H80000000)
            #Else
                .Item(24) = (lU1 >> 19 Or lU1 << 13)
                .Item(25) = (lU0 >> 20 Or lU0 << 12)
            #End If
            lU0 = .Item(38) Xor lT6
            lU1 = .Item(39) Xor lT7
            #If Not HasOperators Then
                .Item(26) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_28 - (lU0 < 0) * LNG_POW2_3) Or _
                    ((lU0 And (LNG_POW2_27 - 1)) * LNG_POW2_4 Or -((lU0 And LNG_POW2_27) <> 0) * &H80000000)
                .Item(27) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_28 - (lU1 < 0) * LNG_POW2_3) Or _
                    ((lU1 And (LNG_POW2_27 - 1)) * LNG_POW2_4 Or -((lU1 And LNG_POW2_27) <> 0) * &H80000000)
            #Else
                .Item(26) = (lU0 >> 28 Or lU0 << 4)
                .Item(27) = (lU1 >> 28 Or lU1 << 4)
            #End If
            lU0 = .Item(46) Xor lT4
            lU1 = .Item(47) Xor lT5
            #If Not HasOperators Then
                .Item(38) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_4 - (lU0 < 0) * LNG_POW2_27) Or _
                    ((lU0 And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((lU0 And LNG_POW2_3) <> 0) * &H80000000)
                .Item(39) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_4 - (lU1 < 0) * LNG_POW2_27) Or _
                    ((lU1 And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((lU1 And LNG_POW2_3) <> 0) * &H80000000)
            #Else
                .Item(38) = (lU0 >> 4 Or lU0 << 28)
                .Item(39) = (lU1 >> 4 Or lU1 << 28)
            #End If
            lU0 = .Item(30) Xor lT8
            lU1 = .Item(31) Xor lT9
            #If Not HasOperators Then
                .Item(46) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_11 - (lU1 < 0) * LNG_POW2_20) Or _
                    ((lU1 And (LNG_POW2_10 - 1)) * LNG_POW2_21 Or -((lU1 And LNG_POW2_10) <> 0) * &H80000000)
                .Item(47) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_12 - (lU0 < 0) * LNG_POW2_19) Or _
                    ((lU0 And (LNG_POW2_11 - 1)) * LNG_POW2_20 Or -((lU0 And LNG_POW2_11) <> 0) * &H80000000)
            #Else
                .Item(46) = (lU1 >> 11 Or lU1 << 21)
                .Item(47) = (lU0 >> 12 Or lU0 << 20)
            #End If
            lU0 = .Item(8) Xor lT6
            lU1 = .Item(9) Xor lT7
            #If Not HasOperators Then
                .Item(30) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_18 - (lU1 < 0) * LNG_POW2_13) Or _
                    ((lU1 And (LNG_POW2_17 - 1)) * LNG_POW2_14 Or -((lU1 And LNG_POW2_17) <> 0) * &H80000000)
                .Item(31) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_19 - (lU0 < 0) * LNG_POW2_12) Or _
                    ((lU0 And (LNG_POW2_18 - 1)) * LNG_POW2_13 Or -((lU0 And LNG_POW2_18) <> 0) * &H80000000)
            #Else
                .Item(30) = (lU1 >> 18 Or lU1 << 14)
                .Item(31) = (lU0 >> 19 Or lU0 << 13)
            #End If
            lU0 = .Item(48) Xor lT6
            lU1 = .Item(49) Xor lT7
            #If Not HasOperators Then
                .Item(8) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_25 - (lU0 < 0) * LNG_POW2_6) Or _
                    ((lU0 And (LNG_POW2_24 - 1)) * LNG_POW2_7 Or -((lU0 And LNG_POW2_24) <> 0) * &H80000000)
                .Item(9) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_25 - (lU1 < 0) * LNG_POW2_6) Or _
                    ((lU1 And (LNG_POW2_24 - 1)) * LNG_POW2_7 Or -((lU1 And LNG_POW2_24) <> 0) * &H80000000)
            #Else
                .Item(8) = (lU0 >> 25 Or lU0 << 7)
                .Item(9) = (lU1 >> 25 Or lU1 << 7)
            #End If
            lU0 = .Item(42) Xor lT0
            lU1 = .Item(43) Xor lT1
            #If Not HasOperators Then
                .Item(48) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_31 - (lU0 < 0)) Or _
                    ((lU0 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lU0 And LNG_POW2_30) <> 0) * &H80000000)
                .Item(49) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_31 - (lU1 < 0)) Or _
                    ((lU1 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lU1 And LNG_POW2_30) <> 0) * &H80000000)
            #Else
                .Item(48) = (lU0 >> 31 Or lU0 << 1)
                .Item(49) = (lU1 >> 31 Or lU1 << 1)
            #End If
            lU0 = .Item(16) Xor lT4
            lU1 = .Item(17) Xor lT5
            #If Not HasOperators Then
                .Item(42) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_4 - (lU1 < 0) * LNG_POW2_27) Or _
                    ((lU1 And (LNG_POW2_3 - 1)) * LNG_POW2_28 Or -((lU1 And LNG_POW2_3) <> 0) * &H80000000)
                .Item(43) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_5 - (lU0 < 0) * LNG_POW2_26) Or _
                    ((lU0 And (LNG_POW2_4 - 1)) * LNG_POW2_27 Or -((lU0 And LNG_POW2_4) <> 0) * &H80000000)
            #Else
                .Item(42) = (lU1 >> 4 Or lU1 << 28)
                .Item(43) = (lU0 >> 5 Or lU0 << 27)
            #End If
            lU0 = .Item(32) Xor lT0
            lU1 = .Item(33) Xor lT1
            #If Not HasOperators Then
                .Item(16) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_9 - (lU1 < 0) * LNG_POW2_22) Or _
                    ((lU1 And (LNG_POW2_8 - 1)) * LNG_POW2_23 Or -((lU1 And LNG_POW2_8) <> 0) * &H80000000)
                .Item(17) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_10 - (lU0 < 0) * LNG_POW2_21) Or _
                    ((lU0 And (LNG_POW2_9 - 1)) * LNG_POW2_22 Or -((lU0 And LNG_POW2_9) <> 0) * &H80000000)
            #Else
                .Item(16) = (lU1 >> 9 Or lU1 << 23)
                .Item(17) = (lU0 >> 10 Or lU0 << 22)
            #End If
            lU0 = .Item(10) Xor lT8
            lU1 = .Item(11) Xor lT9
            #If Not HasOperators Then
                .Item(32) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_14 - (lU0 < 0) * LNG_POW2_17) Or _
                    ((lU0 And (LNG_POW2_13 - 1)) * LNG_POW2_18 Or -((lU0 And LNG_POW2_13) <> 0) * &H80000000)
                .Item(33) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_14 - (lU1 < 0) * LNG_POW2_17) Or _
                    ((lU1 And (LNG_POW2_13 - 1)) * LNG_POW2_18 Or -((lU1 And LNG_POW2_13) <> 0) * &H80000000)
            #Else
                .Item(32) = (lU0 >> 14 Or lU0 << 18)
                .Item(33) = (lU1 >> 14 Or lU1 << 18)
            #End If
            lU0 = .Item(6) Xor lT4
            lU1 = .Item(7) Xor lT5
            #If Not HasOperators Then
                .Item(10) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_18 - (lU0 < 0) * LNG_POW2_13) Or _
                    ((lU0 And (LNG_POW2_17 - 1)) * LNG_POW2_14 Or -((lU0 And LNG_POW2_17) <> 0) * &H80000000)
                .Item(11) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_18 - (lU1 < 0) * LNG_POW2_13) Or _
                    ((lU1 And (LNG_POW2_17 - 1)) * LNG_POW2_14 Or -((lU1 And LNG_POW2_17) <> 0) * &H80000000)
            #Else
                .Item(10) = (lU0 >> 18 Or lU0 << 14)
                .Item(11) = (lU1 >> 18 Or lU1 << 14)
            #End If
            lU0 = .Item(36) Xor lT4
            lU1 = .Item(37) Xor lT5
            #If Not HasOperators Then
                .Item(6) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_21 - (lU1 < 0) * LNG_POW2_10) Or _
                    ((lU1 And (LNG_POW2_20 - 1)) * LNG_POW2_11 Or -((lU1 And LNG_POW2_20) <> 0) * &H80000000)
                .Item(7) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_22 - (lU0 < 0) * LNG_POW2_9) Or _
                    ((lU0 And (LNG_POW2_21 - 1)) * LNG_POW2_10 Or -((lU0 And LNG_POW2_21) <> 0) * &H80000000)
            #Else
                .Item(6) = (lU1 >> 21 Or lU1 << 11)
                .Item(7) = (lU0 >> 22 Or lU0 << 10)
            #End If
            lU0 = .Item(34) Xor lT2
            lU1 = .Item(35) Xor lT3
            #If Not HasOperators Then
                .Item(36) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_24 - (lU1 < 0) * LNG_POW2_7) Or _
                    ((lU1 And (LNG_POW2_23 - 1)) * LNG_POW2_8 Or -((lU1 And LNG_POW2_23) <> 0) * &H80000000)
                .Item(37) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_25 - (lU0 < 0) * LNG_POW2_6) Or _
                    ((lU0 And (LNG_POW2_24 - 1)) * LNG_POW2_7 Or -((lU0 And LNG_POW2_24) <> 0) * &H80000000)
            #Else
                .Item(36) = (lU1 >> 24 Or lU1 << 8)
                .Item(37) = (lU0 >> 25 Or lU0 << 7)
            #End If
            lU0 = .Item(22) Xor lT0
            lU1 = .Item(23) Xor lT1
            #If Not HasOperators Then
                .Item(34) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_27 - (lU0 < 0) * LNG_POW2_4) Or _
                    ((lU0 And (LNG_POW2_26 - 1)) * LNG_POW2_5 Or -((lU0 And LNG_POW2_26) <> 0) * &H80000000)
                .Item(35) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_27 - (lU1 < 0) * LNG_POW2_4) Or _
                    ((lU1 And (LNG_POW2_26 - 1)) * LNG_POW2_5 Or -((lU1 And LNG_POW2_26) <> 0) * &H80000000)
            #Else
                .Item(34) = (lU0 >> 27 Or lU0 << 5)
                .Item(35) = (lU1 >> 27 Or lU1 << 5)
            #End If
            lU0 = .Item(14) Xor lT2
            lU1 = .Item(15) Xor lT3
            #If Not HasOperators Then
                .Item(22) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_29 - (lU0 < 0) * LNG_POW2_2) Or _
                    ((lU0 And (LNG_POW2_28 - 1)) * LNG_POW2_3 Or -((lU0 And LNG_POW2_28) <> 0) * &H80000000)
                .Item(23) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_29 - (lU1 < 0) * LNG_POW2_2) Or _
                    ((lU1 And (LNG_POW2_28 - 1)) * LNG_POW2_3 Or -((lU1 And LNG_POW2_28) <> 0) * &H80000000)
            #Else
                .Item(22) = (lU0 >> 29 Or lU0 << 3)
                .Item(23) = (lU1 >> 29 Or lU1 << 3)
            #End If
            lU0 = .Item(20) Xor lT8
            lU1 = .Item(21) Xor lT9
            #If Not HasOperators Then
                .Item(14) = ((lU1 And &H7FFFFFFF) \ LNG_POW2_30 - (lU1 < 0) * LNG_POW2_1) Or _
                    ((lU1 And (LNG_POW2_29 - 1)) * LNG_POW2_2 Or -((lU1 And LNG_POW2_29) <> 0) * &H80000000)
                .Item(15) = ((lU0 And &H7FFFFFFF) \ LNG_POW2_31 - (lU0 < 0)) Or _
                    ((lU0 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lU0 And LNG_POW2_30) <> 0) * &H80000000)
                .Item(20) = ((lU3 And &H7FFFFFFF) \ LNG_POW2_31 - (lU3 < 0)) Or _
                    ((lU3 And (LNG_POW2_30 - 1)) * LNG_POW2_1 Or -((lU3 And LNG_POW2_30) <> 0) * &H80000000)
            #Else
                .Item(14) = (lU1 >> 30 Or lU1 << 2)
                .Item(15) = (lU0 >> 31 Or lU0 << 1)
                .Item(20) = (lU3 >> 31 Or lU3 << 1)
            #End If
            .Item(21) = lU2
            
            '--- Chi
            For lJdx = 0 To UBound(.Item) Step 10
                lU0 = .Item(lJdx + 0)
                lT2 = .Item(lJdx + 2)
                lT4 = .Item(lJdx + 4)
                lT6 = .Item(lJdx + 6)
                lT8 = .Item(lJdx + 8)
                lU1 = .Item(lJdx + 1)
                lT3 = .Item(lJdx + 3)
                lT5 = .Item(lJdx + 5)
                lT7 = .Item(lJdx + 7)
                lT9 = .Item(lJdx + 9)
                lT0 = (lT8 And Not lT6)
                lT1 = (lT9 And Not lT7)
                lT8 = lT8 Xor (lT2 And Not lU0)
                lT9 = lT9 Xor (lT3 And Not lU1)
                lT2 = lT2 Xor (lT6 And Not lT4)
                lT3 = lT3 Xor (lT7 And Not lT5)
                lT6 = lT6 Xor (lU0 And Not lT8)
                lT7 = lT7 Xor (lU1 And Not lT9)
                lU0 = lU0 Xor (lT4 And Not lT2)
                lU1 = lU1 Xor (lT5 And Not lT3)
                lT4 = lT4 Xor lT0
                lT5 = lT5 Xor lT1
                .Item(lJdx + 0) = lU0
                .Item(lJdx + 2) = lT2
                .Item(lJdx + 4) = lT4
                .Item(lJdx + 6) = lT6
                .Item(lJdx + 8) = lT8
                .Item(lJdx + 1) = lU1
                .Item(lJdx + 3) = lT3
                .Item(lJdx + 5) = lT5
                .Item(lJdx + 7) = lT7
                .Item(lJdx + 9) = lT9
            Next
            
            '--- Iota
            .Item(0) = .Item(0) Xor LNG_RC(lIdx)
            .Item(1) = .Item(1) Xor LNG_RC(lIdx + 1)
        Next
    End With
    pvFromSliced uState
End Sub

Public Sub CryptoSha3Init(uCtx As CryptoSha3Context, ByVal lBitSize As Long)
    Const FADF_AUTO     As Long = 1
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim pDummy          As LongPtr
    
    If LNG_RC(0) = 0 Then
        For Each vElem In Split("1 0 0 89 0 8000008B 0 80008080 1 8B 1 8000 1 80008088 1 80000082 0 B 0 A 1 8082 0 8003 1 808B 1 8000000B 1 8000008A 1 80000081 0 80000081 0 80000008 0 83 0 80008003 1 80008088 0 80000088 1 8000 0 80008082")
            LNG_RC(lIdx) = "&H" & vElem
            lIdx = lIdx + 1
        Next
    End If
    With uCtx
        .DigestSize = (lBitSize + 7) \ 8
        .Capacity = LNG_STATESZ - 2 * .DigestSize
        With .PeekArray
            .cDims = 1
            .fFeatures = FADF_AUTO
            .cbElements = 1
            .cLocks = 1
            .pvData = VarPtr(uCtx.State.Item(0))
            .cElements = LNG_STATESZ \ .cbElements
        End With
        Call CopyMemory(ByVal ArrPtr(.Bytes), VarPtr(.PeekArray), LenB(pDummy))
    End With
End Sub

Public Sub CryptoSha3Update(uCtx As CryptoSha3Context, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lIdx            As Long
    
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    With uCtx
        For lIdx = Pos To Size - 1
            .Bytes(.Absorbed) = .Bytes(.Absorbed) Xor baBuffer(lIdx)
            If .Absorbed = .Capacity - 1 Then
                Keccak .State
                .Absorbed = 0
            Else
                .Absorbed = .Absorbed + 1
            End If
        Next
    End With
End Sub

Public Sub CryptoSha3Finalize(uCtx As CryptoSha3Context, baOutput() As Byte, Optional ByVal OutSize As Long, Optional ByVal Delimiter As Long)
    Dim lIdx            As Long
    Dim lOffset         As Long
    Dim pDummy          As LongPtr
    Dim uEmpty          As CryptoSha3Context
    
    With uCtx
        If OutSize = 0 Then
            OutSize = .DigestSize
        End If
        ReDim baOutput(0 To OutSize - 1) As Byte
        If Delimiter = 0 Then
            Delimiter = &H6
        End If
        .Bytes(.Absorbed) = .Bytes(.Absorbed) Xor Delimiter
        lOffset = .Capacity - 1
        .Bytes(lOffset) = .Bytes(lOffset) Xor &H80
        For lIdx = 0 To UBound(baOutput)
            If lOffset = .Capacity - 1 Then
                Keccak .State
                lOffset = 0
            End If
            baOutput(lIdx) = .Bytes(lOffset)
            lOffset = lOffset + 1
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
    lPadSize = LNG_STATESZ - 2 * ((lBitSize + 7) \ 8)
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
