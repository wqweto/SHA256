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

Private Type ArrayLong10
    Item(0 To 9)        As Long
End Type

Public Type CryptoAsconSlicedContext
    State               As ArrayLong10
    Bytes()             As Byte                     '--- overlaying Words array above
    ArrayBytes          As SAFEARRAY1D
    Absorbed            As Long
    RoundsItermediate   As Long
    RoundsFinal         As Long
    Rate                As Long
    Key()               As Byte                     '--- only for AEAD
    Encrypt             As Boolean                  '--- only for AEAD
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

Private Function pvBitPerm32(ByVal lX As Long, ByVal lMask As Long, ByVal lN As Long) As Long
    Dim lTemp           As Long
    
    #If Not HasOperators Then
        lTemp = (RShift32(lX, lN) Xor lX) And lMask
        pvBitPerm32 = (lX Xor lTemp) Xor LShift32(lTemp, lN)
    #Else
        lTemp = ((lX >> lN) Xor lX) And lMask
        pvBitPerm32 = (lX Xor lTemp) Xor (lTemp << lN)
    #End If
End Function

Private Function pvSeparate(ByVal lX As Long) As Long
    lX = pvBitPerm32(lX, &H22222222, 1)
    lX = pvBitPerm32(lX, &HC0C0C0C, 2)
    lX = pvBitPerm32(lX, &HF000F0, 4)
    pvSeparate = pvBitPerm32(lX, &HFF00&, 8)
End Function

Private Function pvCombine(ByVal lX As Long) As Long
    lX = pvBitPerm32(lX, &HAAAA&, 15)
    lX = pvBitPerm32(lX, &HCCCC&, 14)
    lX = pvBitPerm32(lX, &HF0F0&, 12)
    pvCombine = pvBitPerm32(lX, &HFF00&, 8)
End Function

Private Sub pvToSliced(uState As ArrayLong10)
    Dim lIdx            As Long
    Dim lHigh           As Long
    Dim lLow            As Long
    
    With uState
        For lIdx = 0 To UBound(.Item) Step 2
            lHigh = pvSeparate(BSwap32(.Item(lIdx)))
            lLow = pvSeparate(BSwap32(.Item(lIdx + 1)))
            #If Not HasOperators Then
                .Item(lIdx) = LShift32(lHigh, 16) Or (lLow And &HFFFF&)
                .Item(lIdx + 1) = (lHigh And &HFFFF0000) Or RShift32(lLow, 16)
            #Else
                .Item(lIdx) = (lHigh << 16) Or (lLow And &HFFFF&)
                .Item(lIdx + 1) = (lHigh And &HFFFF0000) Or (lLow >> 16)
            #End If
        Next
    End With
End Sub

Private Sub pvFromSliced(uState As ArrayLong10)
    Dim lIdx            As Long
    Dim lHigh           As Long
    Dim lLow            As Long
    
    With uState
        For lIdx = 0 To UBound(.Item) Step 2
            #If Not HasOperators Then
                lHigh = RShift32(.Item(lIdx), 16) Or (.Item(lIdx + 1) And &HFFFF0000)
                lLow = (.Item(lIdx) And &HFFFF&) Or LShift32(.Item(lIdx + 1), 16)
            #Else
                lHigh = (.Item(lIdx) >> 16) Or (.Item(lIdx + 1) And &HFFFF0000)
                lLow = (.Item(lIdx) And &HFFFF&) Or (.Item(lIdx + 1) << 16)
            #End If
            .Item(lIdx) = BSwap32(pvCombine(lHigh))
            .Item(lIdx + 1) = BSwap32(pvCombine(lLow))
        Next
    End With
End Sub

Private Sub pvAbsorbSliced(uState As ArrayLong10, ByVal lHigh As Long, ByVal lLow As Long, ByVal lOffset As Long)
#If DebugState Then
    Dim lTemp0      As Long: lTemp0 = lHigh
    Dim lTemp1      As Long: lTemp1 = lLow
#End If
    lOffset = 2 * lOffset
    With uState
        lHigh = pvSeparate(BSwap32(lHigh))
        lLow = pvSeparate(BSwap32(lLow))
        #If Not HasOperators Then
            .Item(lOffset) = .Item(lOffset) Xor (LShift32(lHigh, 16) Or (lLow And &HFFFF&))
            .Item(lOffset + 1) = .Item(lOffset + 1) Xor ((lHigh And &HFFFF0000) Or RShift32(lLow, 16))
        #Else
            .Item(lOffset) = .Item(lOffset) Xor ((lHigh << 16) Or (lLow And &HFFFF&))
            .Item(lOffset + 1) = .Item(lOffset + 1) Xor ((lHigh And &HFFFF0000) Or (lLow >> 16))
        #End If
    End With
    #If DebugState Then
        Debug.Print pvDumpState(uCtx), "sliced absorb " & Right$("00000000" & Hex(lTemp0), 8) & " " & Right$("00000000" & Hex(lTemp1), 8), lOffset
    #End If
End Sub

Private Sub pvSqueezeSliced(uState As ArrayLong10, lHigh As Long, lLow As Long, ByVal lOffset As Long)
    lOffset = 2 * lOffset
    With uState
        #If Not HasOperators Then
            lHigh = RShift32(.Item(lOffset), 16) Or (.Item(lOffset + 1) And &HFFFF0000)
            lLow = (.Item(lOffset) And &HFFFF&) Or LShift32(.Item(lOffset + 1), 16)
        #Else
            lHigh = (.Item(lOffset) >> 16) Or (.Item(lOffset + 1) And &HFFFF0000)
            lLow = (.Item(lOffset) And &HFFFF&) Or (.Item(lOffset + 1) << 16)
        #End If
        lHigh = BSwap32(pvCombine(lHigh))
        lLow = BSwap32(pvCombine(lLow))
    End With
End Sub

Private Sub pvDecryptSliced(uState As ArrayLong10, lHigh As Long, lLow As Long, ByVal lOffset As Long)
    Dim lHigh2      As Long
    Dim lLow2       As Long
    
    lOffset = 2 * lOffset
    With uState
        lHigh2 = pvSeparate(BSwap32(lHigh))
        lLow2 = pvSeparate(BSwap32(lLow))
        #If Not HasOperators Then
            lHigh = lHigh2 Xor RShift32(.Item(lOffset), 16) Or (.Item(lOffset + 1) And &HFFFF0000)
            lLow = lLow2 Xor (.Item(lOffset) And &HFFFF&) Or LShift32(.Item(lOffset + 1), 16)
        #Else
            lHigh = lHigh2 Xor (.Item(lOffset) >> 16) Or (.Item(lOffset + 1) And &HFFFF0000)
            lLow = lLow2 Xor (.Item(lOffset) And &HFFFF&) Or (.Item(lOffset + 1) << 16)
        #End If
        lHigh = BSwap32(pvCombine(lHigh))
        lLow = BSwap32(pvCombine(lLow))
        #If Not HasOperators Then
            .Item(lOffset) = (LShift32(lHigh2, 16) Or (lLow2 And &HFFFF&))
            .Item(lOffset + 1) = ((lHigh2 And &HFFFF0000) Or RShift32(lLow2, 16))
        #Else
            .Item(lOffset) = ((lHigh2 << 16) Or (lLow2 And &HFFFF&))
            .Item(lOffset + 1) = ((lHigh2 And &HFFFF0000) Or (lLow2 >> 16))
        #End If
    End With
    #If DebugState Then
        Debug.Print pvDumpState(uCtx), "sliced decrypt " & Right$("00000000" & Hex(lHigh), 8) & " " & Right$("00000000" & Hex(lLow), 8), lOffset
    #End If
End Sub

Private Sub pvPermuteSliced(uState As ArrayLong10, ByVal lRounds As Long)
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

    With uState
        S0_e = .Item(0)
        S0_o = .Item(1)
        S1_e = .Item(2)
        S1_o = .Item(3)
        S2_e = .Item(4)
        S2_o = .Item(5)
        S3_e = .Item(6)
        S3_o = .Item(7)
        S4_e = .Item(8)
        S4_o = .Item(9)
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
        .Item(0) = S0_e
        .Item(1) = S0_o
        .Item(2) = S1_e
        .Item(3) = S1_o
        .Item(4) = S2_e
        .Item(5) = S2_o
        .Item(6) = S3_e
        .Item(7) = S3_o
        .Item(8) = S4_e
        .Item(9) = S4_o
    End With
    #If DebugState Then
        Debug.Print pvDumpState(uCtx), "sliced permute " & lRounds
    #End If
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

Private Function pvDumpState(uCtx As CryptoAsconSlicedContext) As String
    Dim uCopy           As ArrayLong10
    
    uCopy = uCtx.State
    pvFromSliced uCtx.State
    pvDumpState = ToHex(uCtx.Bytes)
    uCtx.State = uCopy
End Function

Private Sub pvInit(uCtx As CryptoAsconSlicedContext)
    Const FADF_AUTO     As Long = 1
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim uEmpty          As CryptoAsconSlicedContext
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
    uCtx = uEmpty
    With uCtx
        With .ArrayBytes
            .cDims = 1
            .fFeatures = FADF_AUTO
            .cbElements = 1
            .cLocks = 1
            .pvData = VarPtr(uCtx.State.Item(0))
            .cElements = LNG_STATESZ \ .cbElements
        End With
        Call CopyMemory(ByVal ArrPtr(.Bytes), VarPtr(.ArrayBytes), LenB(pDummy))
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
            sState = "AA9893EE 3DF067DB 3118B28B 2100FC6 DB928AB4 62DAD598 21991843 E8E3F8B8 C9A58F34 40E125D5"
        Case "ascon-hasha"
            .RoundsItermediate = LNG_ROUNDS \ 2 + 2
            sState = "94014701 A62865FC 8AC38E73 A7FFADC0 29E3C82E 4C38766C 4DA5F6D6 7D37527F A2423CA1 878DBE23"
        Case "ascon-xof"
            .RoundsItermediate = LNG_ROUNDS
            sState = "3B277EB5 16D44C81 2504512B 2024AE62 76A7A366 1822DF8D 7A0AAD5A C655381 320E3E4F B6939453"
        Case "ascon-xofa"
            .RoundsItermediate = LNG_ROUNDS \ 2 + 2
            sState = "68659044 32987BB7 AE6C8DCD 32554553 2721B5F7 29214256 E1856824 5B220DDE E35CCBA8 3F974934"
        Case Else
            Err.Raise vbObjectError, , "Invalid variant for Ascon hash (" & AsconVariant & ")"
        End Select
        .Rate = LNG_BLOCKSZ
        .RoundsFinal = LNG_ROUNDS
        .Key = vbNullString
        '--- init state
        For Each vElem In Split(sState)
            .State.Item(lIdx) = "&H" & vElem
            lIdx = lIdx + 1
        Next
        pvToSliced .State
    End With
End Sub

Private Sub pvInitAead(uCtx As CryptoAsconSlicedContext, baKey() As Byte, Nonce As Variant, AssociatedData As Variant, AsconVariant As String, Optional ByVal Encrypt As Boolean)
    Dim baNonce()       As Byte
    Dim baAad()         As Byte
    Dim lKeySize        As Long
    
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
            lKeySize = LNG_KEYSZ
            .RoundsItermediate = LNG_ROUNDS \ 2
            .Rate = LNG_BLOCKSZ
            .State.Item(0) = &H60C4080
        Case "ascon-128a"
            lKeySize = LNG_KEYSZ
            .RoundsItermediate = LNG_ROUNDS \ 2 + 2
            .Rate = 2 * LNG_BLOCKSZ
            .State.Item(0) = &H80C8080
        Case "ascon-80pq"
            lKeySize = LNG_KEYSZ + 4
            .RoundsItermediate = LNG_ROUNDS \ 2
            .Rate = LNG_BLOCKSZ
            .State.Item(0) = &H60C40A0
        Case Else
            Err.Raise vbObjectError, , "Invalid variant for Ascon AEAD (" & AsconVariant & ")"
        End Select
        .RoundsFinal = LNG_ROUNDS
        .Key = baKey
        .Encrypt = Encrypt
        Debug.Assert UBound(.Key) + 1 = lKeySize
        If UBound(.Key) + 1 <> lKeySize Then
            ReDim Preserve .Key(0 To lKeySize - 1) As Byte
        End If
        '--- init state
        Call CopyMemory(.Bytes(LNG_STATESZ - LNG_NONCESZ - lKeySize), .Key(0), lKeySize)
        Call CopyMemory(.Bytes(LNG_STATESZ - LNG_NONCESZ), baNonce(0), LNG_NONCESZ)
        pvToSliced .State
        pvPermuteSliced .State, .RoundsFinal
        pvInitPeek m_uArrayPeek, .Key
        If lKeySize = LNG_KEYSZ Then
            pvAbsorbSliced .State, m_aPeek(0), m_aPeek(1), 3
            pvAbsorbSliced .State, m_aPeek(2), m_aPeek(3), 4
        Else
            pvAbsorbSliced .State, 0, m_aPeek(0), 2
            pvAbsorbSliced .State, m_aPeek(1), m_aPeek(2), 3
            pvAbsorbSliced .State, m_aPeek(3), m_aPeek(4), 4
        End If
        '--- process associated data
        If UBound(baAad) >= 0 Then
            pvUpdate uCtx, baAad, 0, UBound(baAad) + 1, Final:=.RoundsItermediate
            .Absorbed = 0
        End If
        '--- separator
        .State.Item(8) = .State.Item(8) Xor 1
    End With
End Sub

Private Sub pvUpdate(uCtx As CryptoAsconSlicedContext, baInput() As Byte, ByVal Pos As Long, ByVal Size As Long, Optional ByVal Aead As Boolean, Optional ByVal Final As Long)
    Dim aTemp(0 To 3)   As Long
    Dim aLongs(0 To 3)  As Long
    Dim lIdx            As Long
    Dim lTemp           As Long
    Dim bEncrypt        As Boolean
    Dim bDecrypt        As Boolean

    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    With uCtx
        If Aead Then
            bEncrypt = .Encrypt
            bDecrypt = Not .Encrypt
        End If
        If Size > 0 And .Absorbed > 0 Then
            lTemp = .Rate - .Absorbed
            If lTemp > Size Then
                lTemp = Size
            End If
            Debug.Assert UBound(baInput) + 1 >= Pos + lTemp
            Call CopyMemory(ByVal VarPtr(aTemp(0)) + .Absorbed, baInput(Pos), lTemp)
            pvAbsorbSliced .State, aTemp(0), aTemp(1), 0
            If .Rate > LNG_BLOCKSZ Then
                pvAbsorbSliced .State, aTemp(2), aTemp(3), 1
            End If
            .Absorbed = .Absorbed + lTemp
            If .Absorbed = .Rate Then
                pvPermuteSliced .State, .RoundsItermediate
                .Absorbed = 0
            End If
            Pos = Pos + lTemp
            Size = Size - lTemp
        End If
        If Size > 0 Then
            pvInitPeek m_uArrayPeek, baInput, Pos, Size
            If .Rate = LNG_BLOCKSZ Then
                For lIdx = 0 To UBound(m_aPeek) - 1 Step 2
                    If bDecrypt Then
                        pvDecryptSliced .State, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                    Else
                        pvAbsorbSliced .State, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                        If bEncrypt Then
                            pvSqueezeSliced .State, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                        End If
                    End If
                    pvPermuteSliced .State, .RoundsItermediate
                Next
            Else
                For lIdx = 0 To UBound(m_aPeek) - 3 Step 4
                    If bDecrypt Then
                        pvDecryptSliced .State, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                        pvDecryptSliced .State, m_aPeek(lIdx + 2), m_aPeek(lIdx + 3), 1
                    Else
                        pvAbsorbSliced .State, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                        pvAbsorbSliced .State, m_aPeek(lIdx + 2), m_aPeek(lIdx + 3), 1
                        If bEncrypt Then
                            pvSqueezeSliced .State, m_aPeek(lIdx + 0), m_aPeek(lIdx + 1), 0
                            pvSqueezeSliced .State, m_aPeek(lIdx + 2), m_aPeek(lIdx + 3), 1
                        End If
                    End If
                    pvPermuteSliced .State, .RoundsItermediate
                Next
            End If
            .Absorbed = Size - lIdx * 4
            lIdx = Pos + lIdx * 4
            If .Absorbed > 0 Then
                Debug.Assert UBound(baInput) + 1 >= lIdx + .Absorbed
                Call CopyMemory(aLongs(0), baInput(lIdx), .Absorbed)
            End If
        End If
        Debug.Assert .Absorbed < .Rate
        If Final > 0 Then
            If .Absorbed > 0 And bDecrypt Then
                pvSqueezeSliced .State, aTemp(0), aTemp(1), 0
                If .Rate > LNG_BLOCKSZ Then
                    pvSqueezeSliced .State, aTemp(2), aTemp(3), 1
                End If
                Call FillMemory(ByVal VarPtr(aTemp(0)) + .Absorbed, .Rate - .Absorbed, 0)
                aLongs(0) = aLongs(0) Xor aTemp(0)
                aLongs(1) = aLongs(1) Xor aTemp(1)
                If .Rate > LNG_BLOCKSZ Then
                    aLongs(2) = aLongs(2) Xor aTemp(2)
                    aLongs(3) = aLongs(3) Xor aTemp(3)
                End If
            End If
            Call CopyMemory(ByVal VarPtr(aLongs(0)) + .Absorbed, &H80&, 1)
            pvAbsorbSliced .State, aLongs(0), aLongs(1), 0
            If .Rate > LNG_BLOCKSZ Then
                pvAbsorbSliced .State, aLongs(2), aLongs(3), 1
            End If
            If Aead Then
                If .Absorbed > 0 Then
                    If bEncrypt Then
                        pvSqueezeSliced .State, aLongs(0), aLongs(1), 0
                        If .Rate > LNG_BLOCKSZ Then
                            pvSqueezeSliced .State, aLongs(2), aLongs(3), 1
                        End If
                    End If
                    Debug.Assert UBound(baInput) + 1 >= lIdx + .Absorbed
                    Call CopyMemory(baInput(lIdx), aLongs(0), .Absorbed)
                End If
                pvInitPeek m_uArrayPeek, .Key
                pvAbsorbSliced .State, m_aPeek(0), m_aPeek(1), 1
                pvAbsorbSliced .State, m_aPeek(2), m_aPeek(3), 2
                If UBound(.Key) + 1 > LNG_KEYSZ Then
                    pvAbsorbSliced .State, m_aPeek(4), 0, 3
                End If
            End If
            pvPermuteSliced .State, Final
        ElseIf .Absorbed > 0 Then
            Call CopyMemory(aLongs(0), baInput(lIdx), .Absorbed)
            pvAbsorbSliced .State, aLongs(0), aLongs(1), 0
            If .Rate > LNG_BLOCKSZ Then
                pvAbsorbSliced .State, aLongs(2), aLongs(3), 1
            End If
        End If
    End With
End Sub

Private Sub pvFinalizeHash(uCtx As CryptoAsconSlicedContext, baOutput() As Byte, Optional ByVal OutSize As Long)
    Dim aTemp(0 To 1)   As Long
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim pDummy          As LongPtr
    Dim uEmpty          As CryptoAsconSlicedContext

    If OutSize <= 0 Then
        OutSize = LNG_HASHSZ
    End If
    ReDim baOutput(0 To OutSize - 1) As Byte
    With uCtx
        For lIdx = 0 To OutSize - 1 Step LNG_BLOCKSZ
            pvSqueezeSliced .State, aTemp(0), aTemp(1), 0
            lSize = OutSize - lIdx
            If lSize > LNG_BLOCKSZ Then
                lSize = LNG_BLOCKSZ
                pvPermuteSliced .State, .RoundsItermediate
            End If
            Call CopyMemory(baOutput(lIdx), aTemp(0), lSize)
        Next
        Call CopyMemory(ByVal ArrPtr(.Bytes), pDummy, LenB(pDummy))
    End With
    uCtx = uEmpty
End Sub

Private Sub pvFinalizeAead(uCtx As CryptoAsconSlicedContext, baOutput() As Byte)
    Dim lIdx            As Long
    Dim pDummy          As LongPtr
    Dim uEmpty          As CryptoAsconSlicedContext
    
    With uCtx
        pvInitPeek m_uArrayPeek, .Key
        If UBound(.Key) + 1 = LNG_KEYSZ Then
            pvAbsorbSliced .State, m_aPeek(0), m_aPeek(1), 3
            pvAbsorbSliced .State, m_aPeek(2), m_aPeek(3), 4
        Else
            pvAbsorbSliced .State, 0, m_aPeek(0), 2
            pvAbsorbSliced .State, m_aPeek(1), m_aPeek(2), 3
            pvAbsorbSliced .State, m_aPeek(3), m_aPeek(4), 4
        End If
        ReDim baOutput(0 To LNG_TAGSZ - 1) As Byte
        pvInitPeek m_uArrayPeek, baOutput
        pvSqueezeSliced .State, m_aPeek(0), m_aPeek(1), 3
        pvSqueezeSliced .State, m_aPeek(2), m_aPeek(3), 4
        '--- wipe key
        For lIdx = 0 To UBound(.Key)
            .Key(lIdx) = 0
        Next
        Call CopyMemory(ByVal ArrPtr(.Bytes), pDummy, LenB(pDummy))
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

Public Sub CryptoAsconSlicedHashInit(uCtx As CryptoAsconSlicedContext, Optional AsconVariant As String)
    pvInitHash uCtx, AsconVariant
End Sub

Public Sub CryptoAsconSlicedHashUpdate(uCtx As CryptoAsconSlicedContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    pvUpdate uCtx, baInput, Pos, Size
End Sub

Public Sub CryptoAsconSlicedHashFinalize(uCtx As CryptoAsconSlicedContext, baOutput() As Byte, Optional ByVal OutSize As Long)
    pvUpdate uCtx, uCtx.Bytes, 0, 0, Final:=uCtx.RoundsFinal
    pvFinalizeHash uCtx, baOutput, OutSize
End Sub

Public Function CryptoAsconSlicedHashByteArray(baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional AsconVariant As String, Optional OutSize As Long) As Byte()
    Dim uCtx            As CryptoAsconSlicedContext
    
    pvInitHash uCtx, AsconVariant
    pvUpdate uCtx, baInput, Pos, Size, Final:=uCtx.RoundsFinal
    pvFinalizeHash uCtx, CryptoAsconSlicedHashByteArray, OutSize
End Function

Public Function CryptoAsconSlicedHashText(sText As String, Optional AsconVariant As String) As String
    CryptoAsconSlicedHashText = ToHex(CryptoAsconSlicedHashByteArray(ToUtf8Array(sText), AsconVariant:=AsconVariant))
End Function

Public Sub CryptoAsconSlicedEncrypt(baKey() As Byte, baTag() As Byte, _
            baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, _
            Optional Nonce As Variant, Optional AssociatedData As Variant, Optional AsconVariant As String)
    Dim uCtx            As CryptoAsconSlicedContext
    
    pvInitAead uCtx, baKey, Nonce, AssociatedData, AsconVariant, Encrypt:=True
    pvUpdate uCtx, baInput, Pos, Size, Aead:=True, Final:=uCtx.RoundsFinal
    pvFinalizeAead uCtx, baTag
End Sub

Public Function CryptoAsconSlicedDecrypt(baKey() As Byte, baTag() As Byte, _
            baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, _
            Optional Nonce As Variant, Optional AssociatedData As Variant, Optional AsconVariant As String) As Boolean
    Dim uCtx            As CryptoAsconSlicedContext
    Dim baTemp()        As Byte

    pvInitAead uCtx, baKey, Nonce, AssociatedData, AsconVariant
    pvUpdate uCtx, baInput, Pos, Size, Aead:=True, Final:=uCtx.RoundsFinal
    pvFinalizeAead uCtx, baTemp
    If UBound(baTemp) = UBound(baTag) Then
        CryptoAsconSlicedDecrypt = (InStrB(baTemp, baTag) = 1)
    End If
End Function
