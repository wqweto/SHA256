Attribute VB_Name = "mdSha3"
'--- mdSha3.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const LargeAddressAware = (Win64 = 0 And VBA7 = 0 And VBA6 = 0 And VBA5 = 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare PtrSafe Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
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

Private Const LNG_ROUNDS            As Long = 24
Private Const LNG_SPONGE_WORDS      As Long = 25

#If HasPtrSafe Then
    Private LNG_POW2(0 To 63)       As LongLong
    Private LNG_ROUND_C(0 To 23)    As LongLong
#Else
    Private LNG_POW2(0 To 63)       As Variant
    Private LNG_ROUND_C(0 To 23)    As Variant
#End If

Private Type HashState
    DigestSize      As Long
    Capacity        As Long
    Absorbed        As Long
    #If HasPtrSafe Then
        Words(0 To LNG_SPONGE_WORDS - 1) As LongLong
    #Else
        Words(0 To LNG_SPONGE_WORDS - 1) As Variant
    #End If
    Bytes()         As Byte
    PeekArray       As SAFEARRAY1D
End Type

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

Private Sub Keccak(uState As HashState)
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
    
    With uState
    For lRound = 0 To LNG_ROUNDS - 1
        '--- Theta
        For lIdx = 0 To 4
            C(lIdx) = .Words(lIdx) Xor .Words(lIdx + 5) Xor .Words(lIdx + 10) Xor .Words(lIdx + 15) Xor .Words(lIdx + 20)
        Next
        For lIdx = 0 To 4
            vTemp = C((lIdx + 4) Mod 5) Xor ROTL64(C((lIdx + 1) Mod 5), 1)
            For lJdx = 0 To 24 Step 5
                .Words(lIdx + lJdx) = .Words(lIdx + lJdx) Xor vTemp
            Next
        Next
        '--- Rho & Pi
        aTemp = .Words
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

Private Sub Absorb(uState As HashState, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Dim lIdx            As Long
    Dim lOffset         As Long
    
    If lSize < 0 Then
        lSize = UBound(baBuffer) + 1 - lPos
    End If
    With uState
        lOffset = PeekByte(uState, .Absorbed)
        For lIdx = lPos To lSize - 1
            .Bytes(lOffset) = .Bytes(lOffset) Xor baBuffer(lIdx)
            If .Absorbed = .Capacity - 1 Then
                Keccak uState
                .Absorbed = 0
            Else
                .Absorbed = .Absorbed + 1
            End If
            #If HasPtrSafe Then
                lOffset = lOffset + 1
            #Else
                If lOffset = 7 Then
                    lOffset = PeekByte(uState, .Absorbed)
                Else
                    lOffset = lOffset + 1
                End If
            #End If
        Next
    End With
End Sub

Private Sub Squeeze(uState As HashState, baOutput() As Byte, ByVal lOutSize As Long, ByVal lLFSR As Long)
    Dim lIdx            As Long
    Dim lOffset         As Long
    Dim uEmpty          As HashState
    
    With uState
        ReDim baOutput(0 To lOutSize - 1) As Byte
        lOffset = PeekByte(uState, .Absorbed)
        .Bytes(lOffset) = .Bytes(lOffset) Xor lLFSR
        lOffset = PeekByte(uState, .Capacity - 1)
        .Bytes(lOffset) = .Bytes(lOffset) Xor &H80
        lOffset = PeekByte(uState, 0)
        For lIdx = 0 To UBound(baOutput)
            If lIdx Mod .Capacity = 0 Then
                Keccak uState
            End If
            baOutput(lIdx) = .Bytes(lOffset)
            #If HasPtrSafe Then
                lOffset = lOffset + 1
            #Else
                If lOffset = 7 Then
                    lOffset = PeekByte(uState, lIdx + 1)
                Else
                    lOffset = lOffset + 1
                End If
            #End If
        Next
    End With
    uState = uEmpty
End Sub

Private Sub Init(uState As HashState, ByVal lBitSize As Long)
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
    With uState
        .DigestSize = (lBitSize + 7) \ 8
        .Capacity = LNG_SPONGE_WORDS * 8 - 2 * .DigestSize
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
                    .pvData = VarPtr(uState.Words(0))
                    .cElements = LNG_SPONGE_WORDS * 8
                #Else
                    .cElements = 8
                #End If
            End With
            Dim pDummy As LongPtr
            Call CopyMemory(ByVal ArrPtr(.Bytes), VarPtr(.PeekArray), LenB(pDummy))
        End If
    End With
End Sub

#If HasPtrSafe Then
    Private Function PeekByte(uState As HashState, ByVal lOffset As Long) As Long
        PeekByte = lOffset
    End Function
#Else
    Private Function PeekByte(uState As HashState, ByVal lOffset As Long) As Long
        #If LargeAddressAware Then
            uState.PeekArray.pvData = (VarPtr(uState.Words(lOffset \ 8)) Xor &H80000000) + 8 Xor &H80000000
        #Else
            uState.PeekArray.pvData = VarPtr(uState.Words(lOffset \ 8)) + 8
        #End If
        PeekByte = lOffset Mod 8
    End Function
    
    Private Function CLngLng(vValue As Variant) As Variant
        Const VT_I8 As Long = &H14
        Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
    End Function
#End If

Public Sub CryptoSha3(ByVal lBitSize As Long, baOutput() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim uState          As HashState
    
    Init uState, lBitSize
    Absorb uState, baInput, Pos, Size
    Squeeze uState, baOutput, uState.DigestSize, &H6
End Sub

Public Sub CryptoKeccak(ByVal lBitSize As Long, baOutput() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim uState          As HashState
    
    Init uState, lBitSize
    Absorb uState, baInput, Pos, Size
    Squeeze uState, baOutput, uState.DigestSize, &H1
End Sub

Public Sub CryptoShake(ByVal lBitSize As Long, baOutput() As Byte, ByVal lOutSize As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim uState          As HashState
    
    Init uState, lBitSize
    Absorb uState, baInput, Pos, Size
    Squeeze uState, baOutput, lOutSize, &H1F
End Sub

Public Sub CryptoHmacSha3(ByVal lBitSize As Long, baOutput() As Byte, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Const INNER_PAD     As Long = &H36
    Const OUTER_PAD     As Long = &H5C
    Dim lPadSize        As Long
    Dim lIdx            As Long
    Dim baPass()        As Byte
    Dim baPad()         As Byte
    Dim baHash()        As Byte
    
    '--- pad size is equal to sponge capacity
    lPadSize = LNG_SPONGE_WORDS * 8 - 2 * ((lBitSize + 7) \ 8)
    If UBound(baKey) < lPadSize Then
        baPass = baKey
    Else
        CryptoSha3 lBitSize, baPass, baKey
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
    CryptoSha3 lBitSize, baHash, baPad
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
    CryptoSha3 lBitSize, baOutput, baPad
End Sub
