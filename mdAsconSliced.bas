Attribute VB_Name = "mdAsconSliced"
'--- mdAsconSliced.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)
#Const DebugState = False

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_KEYSZ                 As Long = 16
Private Const LNG_LONGKEYSZ             As Long = 20
Private Const LNG_NONCESZ               As Long = 16
Private Const LNG_TAGSZ                 As Long = 16
Private Const LNG_HASHSZ                As Long = 32
Private Const LNG_ROUNDS                As Long = 12
Private Const LNG_STATESZ               As Long = 40
Private Const LNG_BLOCKSZ               As Long = 8

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    cElements           As Long
    lLbound             As Long
End Type

Public Type CryptoAsconSlicedContext
    State(0 To LNG_STATESZ \ 8 - 1) As Currency
    RoundsItermediate   As Long
    RoundsFinal         As Long
    Rate                As Long
    Bytes()             As Byte
    Words()             As Long
    ArrayBytes          As SAFEARRAY1D
    ArrayWords          As SAFEARRAY1D
    Partial(0 To LNG_BLOCKSZ - 1) As Byte
    NPartial            As Long
End Type

Private LNG_RC(0 To 23)             As Long
Private m_aPeek()                   As Long
Private m_uArrayPeek                As SAFEARRAY1D

#If Not HasOperators Then
Private LNG_POW2(0 To 31)           As Long

Private Function RotR32(ByVal lX As Long, ByVal lN As Long) As Long
    '--- RotR32 = RShift32(X, n) Or LShift32(X, 32 - n)
    Debug.Assert lN <> 0
    RotR32 = ((lX And &H7FFFFFFF) \ LNG_POW2(lN) - (lX < 0) * LNG_POW2(31 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(32 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * &H80000000)
End Function

Private Function BSwap32(ByVal lX As Long) As Long
    BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
              (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
End Function

Private Function LShift32(ByVal lX As Long, ByVal lN As Long) As Long
    If lN = 0 Then
        LShift32 = lX
    Else
        LShift32 = (lX And (LNG_POW2(31 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(31 - lN)) <> 0) * &H80000000
    End If
End Function

Private Function RShift32(ByVal lX As Long, ByVal lN As Long) As Long
    If lN = 0 Then
        RShift32 = lX
    Else
        RShift32 = (lX And &H7FFFFFFF) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(31 - lN)
    End If
End Function
#Else
Private Function BSwap32(ByVal lX As Long) As Long
    Return ((lX And &H000000FF&) << 24) Or _
           ((lX And &H0000FF00&) << 8) Or _
           ((lX And &H00FF0000&) >> 8) Or _
           ((lX And &HFF000000&) >> 24)
End Function
#End If

Private Function pvBitPermute(ByVal lX As Long, ByVal lMask As Long, ByVal lShift As Long) As Long
    Dim lTemp           As Long
    
    #If Not HasOperators Then
        lTemp = (RShift32(lX, lShift) Xor lX) And lMask
        pvBitPermute = (lX Xor lTemp) Xor LShift32(lTemp, lShift)
    #Else
        lTemp = ((lX >> lShift) Xor lX) And lMask
        pvBitPermute = (lX Xor lTemp) Xor (lTemp << lShift)
    #End If
End Function

Private Function pvSeparate(ByVal lX As Long) As Long
'    Dim lTemp           As Long
'
'    lTemp = (RShift32(lX, 1) Xor lX) And &H22222222
'    lX = (lX Xor lTemp) Xor LShift32(lTemp, 1)
'    lTemp = (RShift32(lX, 2) Xor lX) And &HC0C0C0C
'    lX = (lX Xor lTemp) Xor LShift32(lTemp, 2)
'    lTemp = (RShift32(lX, 4) Xor lX) And &HF000F0
'    lX = (lX Xor lTemp) Xor LShift32(lTemp, 4)
'    lTemp = (RShift32(lX, 8) Xor lX) And &HFF00&
'    pvSeparate = (lX Xor lTemp) Xor LShift32(lTemp, 8)
    lX = pvBitPermute(lX, &H22222222, 1)
    lX = pvBitPermute(lX, &HC0C0C0C, 2)
    lX = pvBitPermute(lX, &HF000F0, 4)
    pvSeparate = pvBitPermute(lX, &HFF00&, 8)
End Function

Private Function pvCombine(ByVal lX As Long) As Long
    lX = pvBitPermute(lX, &HAAAA&, 15)
    lX = pvBitPermute(lX, &HCCCC&, 14)
    lX = pvBitPermute(lX, &HF0F0&, 12)
    pvCombine = pvBitPermute(lX, &HFF00&, 8)
End Function

Private Sub pvToSliced(uCtx As CryptoAsconSlicedContext)
    Dim lIdx            As Long
    Dim lHigh           As Long
    Dim lLow            As Long
    
    With uCtx
        For lIdx = 0 To UBound(.Words) Step 2
            lHigh = pvSeparate(BSwap32(.Words(lIdx)))
            lLow = pvSeparate(BSwap32(.Words(lIdx + 1)))
            #If Not HasOperators Then
                .Words(lIdx) = LShift32(lHigh, 16) Or (lLow And &HFFFF&)
                .Words(lIdx + 1) = (lHigh And &HFFFF0000) Or RShift32(lLow, 16)
            #Else
                .Words(lIdx) = (lHigh << 16) Or (lLow And &HFFFF&)
                .Words(lIdx + 1) = (lHigh And &HFFFF0000) Or (lLow >> 16)
            #End If
        Next
    End With
End Sub

Private Sub pvFromSliced(uCtx As CryptoAsconSlicedContext)
    Dim lIdx            As Long
    Dim lHigh           As Long
    Dim lLow            As Long
    
    With uCtx
        For lIdx = 0 To UBound(.Words) Step 2
            #If Not HasOperators Then
                lHigh = RShift32(.Words(lIdx), 16) Or (.Words(lIdx + 1) And &HFFFF0000)
                lLow = (.Words(lIdx) And &HFFFF&) Or LShift32(.Words(lIdx + 1), 16)
            #Else
                lHigh = (.Words(lIdx) >> 16) Or (.Words(lIdx + 1) And &HFFFF0000)
                lLow = (.Words(lIdx) And &HFFFF&) Or (.Words(lIdx + 1) << 16)
            #End If
            .Words(lIdx) = BSwap32(pvCombine(lHigh))
            .Words(lIdx + 1) = BSwap32(pvCombine(lLow))
        Next
    End With
End Sub

Private Sub pvAbsorbSliced(uCtx As CryptoAsconSlicedContext, ByVal lHigh As Long, ByVal lLow As Long, ByVal lOffset As Long)
#If DebugState Then
    Dim lTemp0      As Long: lTemp0 = lHigh
    Dim lTemp1      As Long: lTemp1 = lLow
#End If
    lOffset = 2 * lOffset
    With uCtx
        lHigh = pvSeparate(BSwap32(lHigh))
        lLow = pvSeparate(BSwap32(lLow))
        #If Not HasOperators Then
            .Words(lOffset) = .Words(lOffset) Xor (LShift32(lHigh, 16) Or (lLow And &HFFFF&))
            .Words(lOffset + 1) = .Words(lOffset + 1) Xor ((lHigh And &HFFFF0000) Or RShift32(lLow, 16))
        #Else
            .Words(lOffset) = .Words(lOffset) Xor ((lHigh << 16) Or (lLow And &HFFFF&))
            .Words(lOffset + 1) = .Words(lOffset + 1) Xor ((lHigh And &HFFFF0000) Or (lLow >> 16))
        #End If
    End With
    #If DebugState Then
        Debug.Print pvDumpState(uCtx), "sliced absorb " & Right$("00000000" & Hex(lTemp0), 8) & " " & Right$("00000000" & Hex(lTemp1), 8), lOffset
    #End If
End Sub

Private Sub pvSqueezeSliced(uCtx As CryptoAsconSlicedContext, lHigh As Long, lLow As Long, ByVal lOffset As Long)
    lOffset = 2 * lOffset
    With uCtx
        #If Not HasOperators Then
            lHigh = RShift32(.Words(lOffset), 16) Or (.Words(lOffset + 1) And &HFFFF0000)
            lLow = (.Words(lOffset) And &HFFFF&) Or LShift32(.Words(lOffset + 1), 16)
        #Else
            lHigh = (.Words(lOffset) >> 16) Or (.Words(lOffset + 1) And &HFFFF0000)
            lLow = (.Words(lOffset) And &HFFFF&) Or (.Words(lOffset + 1) << 16)
        #End If
        lHigh = BSwap32(pvCombine(lHigh))
        lLow = BSwap32(pvCombine(lLow))
    End With
End Sub

Private Sub pvDecryptSliced(uCtx As CryptoAsconSlicedContext, lHigh As Long, lLow As Long, ByVal lOffset As Long)
    Dim lHigh2      As Long
    Dim lLow2       As Long
    
    lOffset = 2 * lOffset
    With uCtx
        lHigh2 = pvSeparate(BSwap32(lHigh))
        lLow2 = pvSeparate(BSwap32(lLow))
        #If Not HasOperators Then
            lHigh = lHigh2 Xor RShift32(.Words(lOffset), 16) Or (.Words(lOffset + 1) And &HFFFF0000)
            lLow = lLow2 Xor (.Words(lOffset) And &HFFFF&) Or LShift32(.Words(lOffset + 1), 16)
        #Else
            lHigh = lHigh2 Xor (.Words(lOffset) >> 16) Or (.Words(lOffset + 1) And &HFFFF0000)
            lLow = lLow2 Xor (.Words(lOffset) And &HFFFF&) Or (.Words(lOffset + 1) << 16)
        #End If
        lHigh = BSwap32(pvCombine(lHigh))
        lLow = BSwap32(pvCombine(lLow))
        #If Not HasOperators Then
            .Words(lOffset) = (LShift32(lHigh2, 16) Or (lLow2 And &HFFFF&))
            .Words(lOffset + 1) = ((lHigh2 And &HFFFF0000) Or RShift32(lLow2, 16))
        #Else
            .Words(lOffset) = ((lHigh2 << 16) Or (lLow2 And &HFFFF&))
            .Words(lOffset + 1) = ((lHigh2 And &HFFFF0000) Or (lLow2 >> 16))
        #End If
    End With
    #If DebugState Then
        Debug.Print pvDumpState(uCtx), "sliced decrypt " & Right$("00000000" & Hex(lHigh), 8) & " " & Right$("00000000" & Hex(lLow), 8), lOffset
    #End If
End Sub

Private Sub pvPermuteSliced(uCtx As CryptoAsconSlicedContext, ByVal lRounds As Long)
    Dim S0_e            As Long
    Dim S0_o            As Long
    Dim S1_e            As Long
    Dim S1_o            As Long
    Dim S2_e            As Long
    Dim S2_o            As Long
    Dim S3_e            As Long
    Dim S3_o            As Long
    Dim S4_e            As Long
    Dim S4_o            As Long
    Dim lTemp0          As Long
    Dim lTemp1          As Long
    Dim lIdx            As Long

    With uCtx
        S0_e = .Words(0)
        S0_o = .Words(1)
        S1_e = .Words(2)
        S1_o = .Words(3)
        S2_e = .Words(4)
        S2_o = .Words(5)
        S3_e = .Words(6)
        S3_o = .Words(7)
        S4_e = .Words(8)
        S4_o = .Words(9)
        For lIdx = LNG_ROUNDS - lRounds To LNG_ROUNDS - 1
            '--- round constant
            S2_e = S2_e Xor LNG_RC(2 * lIdx)
            S2_o = S2_o Xor LNG_RC(2 * lIdx + 1)
            '--- substitution layer (high)
            S0_e = S0_e Xor S4_e
            S4_e = S4_e Xor S3_e
            S2_e = S2_e Xor S1_e
            lTemp0 = S0_e And Not S4_e
            S0_e = S0_e Xor (S2_e And Not S1_e)
            S2_e = S2_e Xor (S4_e And Not S3_e)
            S4_e = S4_e Xor (S1_e And Not S0_e)
            S1_e = S1_e Xor (S3_e And Not S2_e)
            S3_e = S3_e Xor lTemp0
            S1_e = S1_e Xor S0_e
            S0_e = S0_e Xor S4_e
            S3_e = S3_e Xor S2_e
            S2_e = Not S2_e
            '--- substitution layer (low)
            S0_o = S0_o Xor S4_o
            S4_o = S4_o Xor S3_o
            S2_o = S2_o Xor S1_o
            lTemp0 = S0_o And Not S4_o
            S0_o = S0_o Xor (S2_o And Not S1_o)
            S2_o = S2_o Xor (S4_o And Not S3_o)
            S4_o = S4_o Xor (S1_o And Not S0_o)
            S1_o = S1_o Xor (S3_o And Not S2_o)
            S3_o = S3_o Xor lTemp0
            S1_o = S1_o Xor S0_o
            S0_o = S0_o Xor S4_o
            S3_o = S3_o Xor S2_o
            S2_o = Not S2_o
            '--- linear diffusion layer
            #If Not HasOperators Then
                lTemp0 = S0_e Xor RotR32(S0_o, 4)
                lTemp1 = S0_o Xor RotR32(S0_e, 5)
                S0_e = S0_e Xor RotR32(lTemp1, 9)
                S0_o = S0_o Xor RotR32(lTemp0, 10)
                lTemp0 = S1_e Xor RotR32(S1_e, 11)
                lTemp1 = S1_o Xor RotR32(S1_o, 11)
                S1_e = S1_e Xor RotR32(lTemp1, 19)
                S1_o = S1_o Xor RotR32(lTemp0, 20)
                lTemp0 = S2_e Xor RotR32(S2_o, 2)
                lTemp1 = S2_o Xor RotR32(S2_e, 3)
                S2_e = S2_e Xor lTemp1
                S2_o = S2_o Xor RotR32(lTemp0, 1)
                lTemp0 = S3_e Xor RotR32(S3_o, 3)
                lTemp1 = S3_o Xor RotR32(S3_e, 4)
                S3_e = S3_e Xor RotR32(lTemp0, 5)
                S3_o = S3_o Xor RotR32(lTemp1, 5)
                lTemp0 = S4_e Xor RotR32(S4_e, 17)
                lTemp1 = S4_o Xor RotR32(S4_o, 17)
                S4_e = S4_e Xor RotR32(lTemp1, 3)
                S4_o = S4_o Xor RotR32(lTemp0, 4)
            #Else
                lTemp0 = S0_e Xor (S0_o >> 4 Or S0_o << 28)
                lTemp1 = S0_o Xor (S0_e >> 5 Or S0_e << 27)
                S0_e = S0_e Xor (lTemp1 >> 9 Or lTemp1 << 23)
                S0_o = S0_o Xor (lTemp0 >> 10 Or lTemp0 << 22)
                lTemp0 = S1_e Xor (S1_e >> 11 Or S1_e << 21)
                lTemp1 = S1_o Xor (S1_o >> 11 Or S1_o << 21)
                S1_e = S1_e Xor (lTemp1 >> 19 Or lTemp1 << 13)
                S1_o = S1_o Xor (lTemp0 >> 20 Or lTemp0 << 12)
                lTemp0 = S2_e Xor (S2_o >> 2 Or S2_o << 30)
                lTemp1 = S2_o Xor (S2_e >> 3 Or S2_e << 29)
                S2_e = S2_e Xor lTemp1
                S2_o = S2_o Xor (lTemp0 >> 1 Or lTemp0 << 31)
                lTemp0 = S3_e Xor (S3_o >> 3 Or S3_o << 29)
                lTemp1 = S3_o Xor (S3_e >> 4 Or S3_e << 28)
                S3_e = S3_e Xor (lTemp0 >> 5 Or lTemp0 << 27)
                S3_o = S3_o Xor (lTemp1 >> 5 Or lTemp1 << 27)
                lTemp0 = S4_e Xor (S4_e >> 17 Or S4_e << 15)
                lTemp1 = S4_o Xor (S4_o >> 17 Or S4_o << 15)
                S4_e = S4_e Xor (lTemp1 >> 3 Or lTemp1 << 29)
                S4_o = S4_o Xor (lTemp0 >> 4 Or lTemp0 << 28)
            #End If
        Next
        .Words(0) = S0_e
        .Words(1) = S0_o
        .Words(2) = S1_e
        .Words(3) = S1_o
        .Words(4) = S2_e
        .Words(5) = S2_o
        .Words(6) = S3_e
        .Words(7) = S3_o
        .Words(8) = S4_e
        .Words(9) = S4_o
    End With
    #If DebugState Then
        Debug.Print pvDumpState(uCtx), "sliced permute " & lRounds
    #End If
End Sub

Private Function pvDumpState(uCtx As CryptoAsconSlicedContext) As String
    pvFromSliced uCtx
    pvDumpState = ToHex(uCtx.Bytes)
    pvToSliced uCtx
End Function

Private Sub pvInit(uCtx As CryptoAsconSlicedContext)
    Const FADF_AUTO     As Long = 1
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim pDummy          As LongPtr
    
    #If Not HasOperators Then
        If LNG_POW2(0) = 0 Then
            LNG_POW2(0) = 1
            For lIdx = 1 To 30
                LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
            Next
            LNG_POW2(31) = &H80000000
        End If
    #End If
    If LNG_RC(0) = 0 Then
        lIdx = 0
        For Each vElem In Split("12 12 9 12 12 9 9 9 6 12 3 12 6 9 3 9 12 6 9 6 12 3 9 3")
            LNG_RC(lIdx) = vElem
            lIdx = lIdx + 1
        Next
        With m_uArrayPeek
            .cDims = 1
            .fFeatures = FADF_AUTO
            .cbElements = 4
            .cLocks = 1
        End With
        Call CopyMemory(ByVal ArrPtr(m_aPeek), VarPtr(m_uArrayPeek), LenB(pDummy))
    End If
    With uCtx
        With .ArrayBytes
            .cDims = 1
            .fFeatures = FADF_AUTO
            .cbElements = 1
            .cLocks = 1
            .pvData = VarPtr(uCtx.State(0))
            .cElements = LNG_STATESZ \ .cbElements
        End With
        Call CopyMemory(ByVal ArrPtr(.Bytes), VarPtr(.ArrayBytes), LenB(pDummy))
        With .ArrayWords
            .cDims = 1
            .fFeatures = FADF_AUTO
            .cbElements = 4
            .cLocks = 1
            .pvData = VarPtr(uCtx.State(0))
            .cElements = LNG_STATESZ \ .cbElements
        End With
        Call CopyMemory(ByVal ArrPtr(.Words), VarPtr(.ArrayWords), LenB(pDummy))
    End With
End Sub

Private Sub pvInitPeek(uArray As SAFEARRAY1D, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    With uArray
        .pvData = VarPtr(baInput(Pos))
        .cElements = Size \ .cbElements
    End With
End Sub

Private Sub pvInitHash(uCtx As CryptoAsconSlicedContext, Optional AsconVariant As String)
    Dim sState          As Variant
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    pvInit uCtx
    With uCtx
        Select Case LCase$(AsconVariant)
        Case "ascon-hash", vbNullString
            .RoundsItermediate = LNG_ROUNDS
            sState = "446318142388178.635 14863613160486.9771 712324061313542.0084 -166521396747559.9293 467505948832861.778"
        Case "ascon-hasha"
            .RoundsItermediate = 8
            sState = "-647381232885581.2351 -634115870784097.1149 549226995250965.9182 902277108517712.4566 -867907184661769.5071"
        Case "ascon-xof"
            .RoundsItermediate = LNG_ROUNDS
            sState = "164502388182400.9909 231616784492634.5515 173919820479251.3382 89321191666631.817 -529072205218721.0161"
        Case "ascon-xofa"
            .RoundsItermediate = 8
            sState = "364579992601713.466 362688130062775.4445 296372296757763.8391 656682645757712.1828 458221163737440.5544"
        Case Else
            Err.Raise vbObjectError, , "Invalid variant for Ascon hash (" & AsconVariant & ")"
        End Select
        .Rate = 8
        .RoundsFinal = LNG_ROUNDS
        '--- init state
        lIdx = 0
        For Each vElem In Split(sState)
            .State(lIdx) = vElem
            lIdx = lIdx + 1
        Next
        pvToSliced uCtx
    End With
End Sub

Private Sub pvInitAead(uCtx As CryptoAsconSlicedContext, baKey() As Byte, Nonce As Variant, AssociatedData As Variant, AsconVariant As String)
    Dim baNonce()       As Byte
    Dim baAad()         As Byte
    Dim lIdx            As Long
    Dim lSize           As Long
    
    pvInit uCtx
    If IsMissing(Nonce) Then
        baNonce = vbNullString
    Else
        baNonce = Nonce
    End If
    ReDim Preserve baNonce(0 To LNG_NONCESZ - 1) As Byte
    If IsMissing(AssociatedData) Then
        baAad = vbNullString
    Else
        baAad = AssociatedData
    End If
    With uCtx
        Select Case LCase$(AsconVariant)
        Case "ascon-128", vbNullString
            .RoundsItermediate = LNG_ROUNDS \ 2
            .Rate = 8
            .State(0) = 10146.624@
            Debug.Assert UBound(baKey) + 1 = LNG_KEYSZ
            ReDim Preserve baKey(0 To LNG_KEYSZ - 1) As Byte
        Case "ascon-128a"
            .RoundsItermediate = 8
            .Rate = 16
            .State(0) = 13503.7056@
            Debug.Assert UBound(baKey) + 1 = LNG_KEYSZ
            ReDim Preserve baKey(0 To LNG_KEYSZ - 1) As Byte
        Case "ascon-80pq"
            .RoundsItermediate = LNG_ROUNDS \ 2
            .Rate = 8
            .State(0) = 10146.6272@
            Debug.Assert UBound(baKey) + 1 = LNG_LONGKEYSZ
            ReDim Preserve baKey(0 To LNG_LONGKEYSZ - 1) As Byte
        Case Else
            Err.Raise vbObjectError, , "Invalid variant for Ascon AEAD (" & AsconVariant & ")"
        End Select
        .RoundsFinal = LNG_ROUNDS
        '--- init state
        For lIdx = 1 To UBound(.State)
            .State(lIdx) = 0
        Next
        lSize = UBound(baKey) + 1
        Call CopyMemory(.Bytes(LNG_STATESZ - LNG_NONCESZ - lSize), baKey(0), lSize)
        Call CopyMemory(.Bytes(LNG_STATESZ - LNG_NONCESZ), baNonce(0), LNG_NONCESZ)
        pvToSliced uCtx
        pvPermuteSliced uCtx, .RoundsFinal
        pvInitPeek m_uArrayPeek, baKey
        If UBound(baKey) + 1 = LNG_KEYSZ Then
            pvAbsorbSliced uCtx, m_aPeek(0), m_aPeek(1), 3
            pvAbsorbSliced uCtx, m_aPeek(2), m_aPeek(3), 4
        Else
            pvAbsorbSliced uCtx, 0, m_aPeek(0), 2
            pvAbsorbSliced uCtx, m_aPeek(1), m_aPeek(2), 3
            pvAbsorbSliced uCtx, m_aPeek(3), m_aPeek(4), 4
        End If
        '--- process associated data
        If UBound(baAad) >= 0 Then
            pvUpdate uCtx, baAad, 0, UBound(baAad) + 1, Final:=.RoundsItermediate
        End If
        '--- separator
        .Words(8) = .Words(8) Xor 1
    End With
End Sub

Private Sub pvUpdate(uCtx As CryptoAsconSlicedContext, baInput() As Byte, ByVal Pos As Long, ByVal Size As Long, Optional ByVal Encrypt As Boolean, Optional ByVal Decrypt As Boolean, Optional ByVal Final As Long, Optional Key As Variant)
    Dim aLongs(0 To 3)  As Long
    Dim lIdx            As Long
    Dim baKey()         As Byte
    Dim aTemp(0 To 3)   As Long

    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    With uCtx
        If Size > 0 Then
            pvInitPeek m_uArrayPeek, baInput, Pos, Size
            If .Rate = 8 Then
                For lIdx = 0 To UBound(m_aPeek) - 1 Step 2
                    If Decrypt Then
                        pvDecryptSliced uCtx, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                    Else
                        pvAbsorbSliced uCtx, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                        If Encrypt Then
                            pvSqueezeSliced uCtx, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                        End If
                    End If
                    pvPermuteSliced uCtx, .RoundsItermediate
                Next
            Else
                For lIdx = 0 To UBound(m_aPeek) - 3 Step 4
                    If Decrypt Then
                        pvDecryptSliced uCtx, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                        pvDecryptSliced uCtx, m_aPeek(lIdx + 2), m_aPeek(lIdx + 3), 1
                    Else
                        pvAbsorbSliced uCtx, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                        pvAbsorbSliced uCtx, m_aPeek(lIdx + 2), m_aPeek(lIdx + 3), 1
                        If Encrypt Then
                            pvSqueezeSliced uCtx, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                            pvSqueezeSliced uCtx, m_aPeek(lIdx + 2), m_aPeek(lIdx + 3), 1
                        End If
                    End If
                    pvPermuteSliced uCtx, .RoundsItermediate
                Next
            End If
        End If
        If Final > 0 Then
            lIdx = lIdx * 4
            If Size - lIdx > 0 Then
                Call CopyMemory(aLongs(0), baInput(lIdx), Size - lIdx)
                If Decrypt Then
                    pvSqueezeSliced uCtx, aTemp(0), aTemp(1), 0
                    If .Rate > 8 Then
                        pvSqueezeSliced uCtx, aTemp(2), aTemp(3), 1
                    End If
                    Call FillMemory(ByVal VarPtr(aTemp(0)) + Size - lIdx, 16 - Size + lIdx, 0)
                    aLongs(0) = aLongs(0) Xor aTemp(0)
                    aLongs(1) = aLongs(1) Xor aTemp(1)
                    If .Rate > 8 Then
                        aLongs(2) = aLongs(2) Xor aTemp(2)
                        aLongs(3) = aLongs(3) Xor aTemp(3)
                    End If
                End If
            End If
            Call CopyMemory(ByVal VarPtr(aLongs(0)) + Size - lIdx, &H80&, 1)
            pvAbsorbSliced uCtx, aLongs(0), aLongs(1), 0
            If .Rate > 8 Then
                pvAbsorbSliced uCtx, aLongs(2), aLongs(3), 1
            End If
            If Size - lIdx > 0 Then
                If Encrypt Then
                    pvSqueezeSliced uCtx, aLongs(0), aLongs(1), 0
                    If .Rate > 8 Then
                        pvSqueezeSliced uCtx, aLongs(2), aLongs(3), 1
                    End If
                End If
                If Encrypt Or Decrypt Then
                    Call CopyMemory(baInput(lIdx), aLongs(0), Size - lIdx)
                End If
            End If
            If Not IsMissing(Key) Then
                baKey = Key
                pvInitPeek m_uArrayPeek, baKey
                pvAbsorbSliced uCtx, m_aPeek(0), m_aPeek(1), 1
                pvAbsorbSliced uCtx, m_aPeek(2), m_aPeek(3), 2
                If UBound(baKey) + 1 > LNG_KEYSZ Then
                    pvAbsorbSliced uCtx, m_aPeek(4), 0, 3
                End If
            End If
            pvPermuteSliced uCtx, Final
        Else
            '--- ToDo: preserve Partial
            Debug.Assert False
        End If
    End With
End Sub

Private Sub pvFinalizeHash(uCtx As CryptoAsconSlicedContext, baOutput() As Byte, Optional ByVal OutSize As Long)
    Dim lIdx            As Long
    Dim aLongs(0 To 1)  As Long
    Dim lTemp           As Long
    Dim pDummy          As LongPtr
    Dim uEmpty          As CryptoAsconSlicedContext

    If OutSize <= 0 Then
        OutSize = LNG_HASHSZ
    End If
    ReDim baOutput(0 To OutSize - 1) As Byte
    For lIdx = 0 To OutSize - 1 Step 8
        pvSqueezeSliced uCtx, aLongs(0), aLongs(1), 0
        lTemp = OutSize - lIdx
        If lTemp > 8 Then
            lTemp = 8
            pvPermuteSliced uCtx, uCtx.RoundsItermediate
        End If
        Call CopyMemory(baOutput(lIdx), aLongs(0), lTemp)
    Next
    With uCtx
        Call CopyMemory(ByVal ArrPtr(.Bytes), pDummy, LenB(pDummy))
        Call CopyMemory(ByVal ArrPtr(.Words), pDummy, LenB(pDummy))
    End With
    uCtx = uEmpty
End Sub

Private Sub pvFinalizeAead(uCtx As CryptoAsconSlicedContext, baKey() As Byte, baTag() As Byte)
    Dim pDummy          As LongPtr
    Dim uEmpty          As CryptoAsconSlicedContext

    With uCtx
        pvInitPeek m_uArrayPeek, baKey
        If UBound(baKey) + 1 = LNG_KEYSZ Then
            pvAbsorbSliced uCtx, m_aPeek(0), m_aPeek(1), 3
            pvAbsorbSliced uCtx, m_aPeek(2), m_aPeek(3), 4
        Else
            pvAbsorbSliced uCtx, 0, m_aPeek(0), 2
            pvAbsorbSliced uCtx, m_aPeek(1), m_aPeek(2), 3
            pvAbsorbSliced uCtx, m_aPeek(3), m_aPeek(4), 4
        End If
        ReDim baTag(0 To LNG_TAGSZ - 1) As Byte
        pvInitPeek m_uArrayPeek, baTag
        pvSqueezeSliced uCtx, m_aPeek(0), m_aPeek(1), 3
        pvSqueezeSliced uCtx, m_aPeek(2), m_aPeek(3), 4
        Call CopyMemory(ByVal ArrPtr(.Bytes), pDummy, LenB(pDummy))
        Call CopyMemory(ByVal ArrPtr(.Words), pDummy, LenB(pDummy))
    End With
    uCtx = uEmpty
End Sub

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

'Public Sub CryptoAsconHashInit(uCtx As CryptoAsconSlicedContext, Optional AsconVariant As String)
'    pvInitHash uCtx, AsconVariant
'End Sub
'
'Public Sub CryptoAsconHashUpdate(uCtx As CryptoAsconSlicedContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
'    pvUpdate uCtx, baInput, Pos, Size
'End Sub
'
'Public Sub CryptoAsconHashFinalize(uCtx As CryptoAsconSlicedContext, baOutput() As Byte, Optional ByVal OutSize As Long)
'    pvFinalizeHash uCtx, baOutput, OutSize
'End Sub

Public Function CryptoAsconHashByteArraySliced(baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional AsconVariant As String, Optional OutSize As Long) As Byte()
    Dim uCtx            As CryptoAsconSlicedContext
    
    pvInitHash uCtx, AsconVariant
    pvUpdate uCtx, baInput, Pos, Size, Final:=uCtx.RoundsFinal
    pvFinalizeHash uCtx, CryptoAsconHashByteArraySliced, OutSize
End Function

Public Function CryptoAsconHashTextSliced(sText As String, Optional AsconVariant As String) As String
    CryptoAsconHashTextSliced = ToHex(CryptoAsconHashByteArraySliced(ToUtf8Array(sText), AsconVariant:=AsconVariant))
End Function

Public Sub CryptoAsconEncryptSliced(baKey() As Byte, baTag() As Byte, _
            baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, _
            Optional Nonce As Variant, Optional AssociatedData As Variant, Optional AsconVariant As String)
    Dim uCtx            As CryptoAsconSlicedContext
    
    pvInitAead uCtx, baKey, Nonce, AssociatedData, AsconVariant
    pvUpdate uCtx, baInput, Pos, Size, Encrypt:=True, Final:=uCtx.RoundsFinal, Key:=baKey
    pvFinalizeAead uCtx, baKey, baTag
End Sub

Public Function CryptoAsconDecryptSliced(baKey() As Byte, baTag() As Byte, _
            baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, _
            Optional Nonce As Variant, Optional AssociatedData As Variant, Optional AsconVariant As String) As Boolean
    Dim uCtx            As CryptoAsconSlicedContext
    Dim baTemp()        As Byte

    pvInitAead uCtx, baKey, Nonce, AssociatedData, AsconVariant
    pvUpdate uCtx, baInput, Pos, Size, Decrypt:=True, Final:=uCtx.RoundsFinal, Key:=baKey
    pvFinalizeAead uCtx, baKey, baTemp
    If UBound(baTemp) = UBound(baTag) Then
        CryptoAsconDecryptSliced = (InStrB(baTemp, baTag) = 1)
    End If
End Function
