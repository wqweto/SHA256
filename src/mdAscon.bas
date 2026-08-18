Attribute VB_Name = "mdAscon"
'--- mdAscon.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)
#Const DebugState = False

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare PtrSafe Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_KEYSZ                 As Long = 16
Private Const LNG_LONGKEYSZ             As Long = 20
Private Const LNG_NONCESZ               As Long = 16
Private Const LNG_TAGSZ                 As Long = 16
Private Const LNG_ROUNDS                As Long = 12
Private Const LNG_STATESZ               As Long = 40

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    cElements           As Long
    lLbound             As Long
End Type

Public Type CryptoAsconContext
    Words(0 To LNG_STATESZ \ 8 - 1) As Currency
    Bytes()             As Byte                     '--- overlaying Words array above
    ArrayBytes          As SAFEARRAY1D
    Absorbed            As Long
    RoundsItermediate   As Long
    RoundsFinal         As Long
    Rate                As Long
End Type

#If Not HasOperators Then
#If HasPtrSafe Then
Private LNG_POW2(0 To 63)           As LongLong
Private LNG_SIGN_BIT                As LongLong ' 2 ^ 63
#Else
Private LNG_POW2(0 To 63)           As Variant
Private LNG_SIGN_BIT                As Variant
#End If

#If HasPtrSafe Then
Private Function RotR64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RotR64(lX As Variant, ByVal lN As Long) As Variant
#End If
    '--- RotR64 = RShift64(X, n) Or LShift64(X, 64 - n)
    Debug.Assert lN <> 0
    RotR64 = ((lX And (-1 Xor LNG_SIGN_BIT)) \ LNG_POW2(lN) Or -(lX < 0) * LNG_POW2(63 - lN)) Or _
        ((lX And (LNG_POW2(lN - 1) - 1)) * LNG_POW2(64 - lN) Or -((lX And LNG_POW2(lN - 1)) <> 0) * LNG_SIGN_BIT)
End Function

#If Not HasPtrSafe Then
Private Function CLngLng(vValue As Variant) As Variant
    Const VT_I8 As Long = &H14
    Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
End Function

Private Function ToLngLng(ByVal cValue As Currency) As Variant
    Const VT_I8 As Long = &H14
    Static B(0 To 1)    As Long
    Dim lIdx            As Long
    
    Call CopyMemory(B(0), cValue, 8)
    lIdx = B(0)
    B(0) = BSwap32(B(1))
    B(1) = BSwap32(lIdx)
    Call VariantChangeType(ToLngLng, ToLngLng, 0, VT_I8)
    #If LargeAddressAware Then
        Call CopyMemory(ByVal (VarPtr(ToLngLng) Xor &H80000000) + 8 Xor &H80000000, B(0), 8)
    #Else
        Call CopyMemory(ByVal VarPtr(ToLngLng) + 8, B(0), 8)
    #End If
End Function

Private Function FromLngLng(lValue As Variant) As Currency
    Static B(0 To 1)    As Long
    Dim lIdx            As Long
    
    #If LargeAddressAware Then
        Call CopyMemory(B(0), ByVal (VarPtr(lValue) Xor &H80000000) + 8 Xor &H80000000, 8)
    #Else
        Call CopyMemory(B(0), ByVal VarPtr(lValue) + 8, 8)
    #End If
    lIdx = B(0)
    B(0) = BSwap32(B(1))
    B(1) = BSwap32(lIdx)
    Call CopyMemory(FromLngLng, B(0), 8)
End Function
#Else
Private Function ToLngLng(ByVal cValue As Currency) As LongLong
    Const VT_I8 As Long = &H14
    Static B(0 To 1)    As Long
    Dim lIdx            As Long
    
    Call CopyMemory(B(0), cValue, 8)
    lIdx = B(0)
    B(0) = BSwap32(B(1))
    B(1) = BSwap32(lIdx)
    Call CopyMemory(ToLngLng, B(0), 8)
End Function

Private Function FromLngLng(ByVal lValue As LongLong) As Currency
    Static B(0 To 1)    As Long
    Dim lIdx            As Long
    
    Call CopyMemory(B(0), lValue, 8)
    lIdx = B(0)
    B(0) = BSwap32(B(1))
    B(1) = BSwap32(lIdx)
    Call CopyMemory(FromLngLng, B(0), 8)
End Function
#End If

Private Function BSwap32(ByVal lX As Long) As Long
    BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or _
              (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * &H80000000
End Function

Private Sub pvPermute(uCtx As CryptoAsconContext, ByVal lRounds As Long)
#If HasPtrSafe Then
    Dim S0              As LongLong
    Dim S1              As LongLong
    Dim S2              As LongLong
    Dim S3              As LongLong
    Dim S4              As LongLong
    Dim lTemp           As LongLong
#Else
    Dim S0              As Variant
    Dim S1              As Variant
    Dim S2              As Variant
    Dim S3              As Variant
    Dim S4              As Variant
    Dim lTemp           As Variant
#End If
    Dim lIdx            As Long

    #If DebugState Then
        Debug.Print ToHex(uCtx.Bytes), "before permute " & lRounds
    #End If
    With uCtx
        S0 = ToLngLng(.Words(0))
        S1 = ToLngLng(.Words(1))
        S2 = ToLngLng(.Words(2))
        S3 = ToLngLng(.Words(3))
        S4 = ToLngLng(.Words(4))
        For lIdx = LNG_ROUNDS - lRounds To LNG_ROUNDS - 1
            '--- round constant
            S2 = S2 Xor (&HF0 - lIdx * 15)
            '--- substitution layer
            S0 = S0 Xor S4
            S4 = S4 Xor S3
            S2 = S2 Xor S1
            lTemp = S0 And Not S4
            S0 = S0 Xor (S2 And Not S1)
            S2 = S2 Xor (S4 And Not S3)
            S4 = S4 Xor (S1 And Not S0)
            S1 = S1 Xor (S3 And Not S2)
            S3 = S3 Xor lTemp
            S1 = S1 Xor S0
            S0 = S0 Xor S4
            S3 = S3 Xor S2
            S2 = Not S2
            '--- linear diffusion layer
            S0 = S0 Xor RotR64(S0, 19) Xor RotR64(S0, 28)
            S1 = S1 Xor RotR64(S1, 61) Xor RotR64(S1, 39)
            S2 = S2 Xor RotR64(S2, 1) Xor RotR64(S2, 6)
            S3 = S3 Xor RotR64(S3, 10) Xor RotR64(S3, 17)
            S4 = S4 Xor RotR64(S4, 7) Xor RotR64(S4, 41)
            #If DebugPermutation Then
                .Words(0) = FromLngLng(S0)
                .Words(1) = FromLngLng(S1)
                .Words(2) = FromLngLng(S2)
                .Words(3) = FromLngLng(S3)
                .Words(4) = FromLngLng(S4)
                Debug.Print ToHex(uCtx.Bytes)
            #End If
        Next
        .Words(0) = FromLngLng(S0)
        .Words(1) = FromLngLng(S1)
        .Words(2) = FromLngLng(S2)
        .Words(3) = FromLngLng(S3)
        .Words(4) = FromLngLng(S4)
    End With
    #If DebugState Then
        Debug.Print ToHex(uCtx.Bytes), "after permute " & lRounds
    #End If
End Sub
#Else
Private Type ArrayLongLong5
    Item0               As LongLong
    Item1               As LongLong
    Item2               As LongLong
    Item3               As LongLong
    Item4               As LongLong
End Type

Private Function BSwap64(ByVal lX As LongLong) As LongLong
    Return ((lX And &H00000000000000FF^) << 56) Or _
           ((lX And &H000000000000FF00^) << 40) Or _
           ((lX And &H0000000000FF0000^) << 24) Or _
           ((lX And &H00000000FF000000^) << 8) Or _
           ((lX And &H000000FF00000000^) >> 8) Or _
           ((lX And &H0000FF0000000000^) >> 24) Or _
           ((lX And &H00FF000000000000^) >> 40) Or _
           ((lX And &HFF00000000000000^) >> 56)
End Function

Private Sub pvAssign(uArray As ArrayLongLong5, S0 As LongLong, S1 As LongLong, S2 As LongLong, S3 As LongLong, S4 As LongLong)
    With uArray
        S0 = BSwap64(.Item0)
        S1 = BSwap64(.Item1)
        S2 = BSwap64(.Item2)
        S3 = BSwap64(.Item3)
        S4 = BSwap64(.Item4)
    End With
End Sub

Private Sub pvUnassign(uArray As ArrayLongLong5, ByVal S0 As LongLong, ByVal S1 As LongLong, ByVal S2 As LongLong, ByVal S3 As LongLong, ByVal S4 As LongLong)
    With uArray
        .Item0 = BSwap64(S0)
        .Item1 = BSwap64(S1)
        .Item2 = BSwap64(S2)
        .Item3 = BSwap64(S3)
        .Item4 = BSwap64(S4)
    End With
End Sub

Private Sub pvPermute(uCtx As CryptoAsconContext, ByVal lRounds As Long)
    Dim S0              As LongLong
    Dim S1              As LongLong
    Dim S2              As LongLong
    Dim S3              As LongLong
    Dim S4              As LongLong
    Dim lTemp           As LongLong
    Dim lIdx            As Long

    With uCtx
        pvAssign VarPtr(.Bytes(0)), S0, S1, S2, S3, S4
        For lIdx = LNG_ROUNDS - lRounds To LNG_ROUNDS - 1
            '--- round constant
            S2 = S2 Xor (&HF0 - lIdx * 15)
            '--- substitution layer
            S0 = S0 Xor S4
            S4 = S4 Xor S3
            S2 = S2 Xor S1
            lTemp = S0 And Not S4
            S0 = S0 Xor (S2 And Not S1)
            S2 = S2 Xor (S4 And Not S3)
            S4 = S4 Xor (S1 And Not S0)
            S1 = S1 Xor (S3 And Not S2)
            S3 = S3 Xor lTemp
            S1 = S1 Xor S0
            S0 = S0 Xor S4
            S3 = S3 Xor S2
            S2 = Not S2
            '--- linear diffusion layer
            lTemp = S0 Xor (S0 >> 9 Or S0 << 55)
            S0 = S0 Xor (lTemp >> 19 Or lTemp << 45)
            lTemp = S1 Xor (S1 >> 22 Or S1 << 42)
            S1 = S1 Xor (lTemp >> 39 Or lTemp << 25)
            lTemp = S2 Xor (S2 >> 5 Or S2 << 59)
            S2 = S2 Xor (lTemp >> 1 Or lTemp << 63)
            lTemp = S3 Xor (S3 >> 7 Or S3 << 57)
            S3 = S3 Xor (lTemp >> 10 Or lTemp << 54)
            lTemp = S4 Xor (S4 >> 34 Or S4 << 30)
            S4 = S4 Xor (lTemp >> 7 Or lTemp << 57)
        Next
        pvUnassign VarPtr(.Bytes(0)), S0, S1, S2, S3, S4
    End With
End Sub
#End If

Private Sub pvInit(uCtx As CryptoAsconContext)
    Const FADF_AUTO     As Long = 1
    Dim lIdx            As Long
    Dim pDummy          As LongPtr
    
    #If Not HasOperators Then
        If LNG_POW2(0) = 0 Then
            LNG_POW2(0) = CLngLng(1)
            For lIdx = 1 To 63
                LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
            Next
            LNG_SIGN_BIT = LNG_POW2(63)
        End If
    #End If
    With uCtx
        If .ArrayBytes.cDims = 0 Then
            With .ArrayBytes
                .cDims = 1
                .fFeatures = FADF_AUTO
                .cbElements = 1
                .cLocks = 1
                .pvData = VarPtr(uCtx.Words(0))
                .cElements = LNG_STATESZ
            End With
            Call CopyMemory(ByVal ArrPtr(.Bytes), VarPtr(.ArrayBytes), LenB(pDummy))
        End If
    End With
End Sub

Private Sub pvInitHash(uCtx As CryptoAsconContext, Optional AsconVariant As String)
    Dim sState          As Variant
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    pvInit uCtx
    With uCtx
        .Absorbed = 0
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
            .Words(lIdx) = vElem
            lIdx = lIdx + 1
        Next
    End With
End Sub

Private Sub pvInitAead(uCtx As CryptoAsconContext, baKey() As Byte, Nonce As Variant, AssociatedData As Variant, AsconVariant As String)
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
        .Absorbed = 0
        Select Case LCase$(AsconVariant)
        Case "ascon-128", vbNullString
            .RoundsItermediate = LNG_ROUNDS \ 2
            .Rate = 8
            .Words(0) = 10146.624@
            Debug.Assert UBound(baKey) + 1 = LNG_KEYSZ
            ReDim Preserve baKey(0 To LNG_KEYSZ - 1) As Byte
        Case "ascon-128a"
            .RoundsItermediate = 8
            .Rate = 16
            .Words(0) = 13503.7056@
            Debug.Assert UBound(baKey) + 1 = LNG_KEYSZ
            ReDim Preserve baKey(0 To LNG_KEYSZ - 1) As Byte
        Case "ascon-80pq"
            .RoundsItermediate = LNG_ROUNDS \ 2
            .Rate = 8
            .Words(0) = 10146.6272@
            Debug.Assert UBound(baKey) + 1 = LNG_LONGKEYSZ
            ReDim Preserve baKey(0 To LNG_LONGKEYSZ - 1) As Byte
        Case Else
            Err.Raise vbObjectError, , "Invalid variant for Ascon AEAD (" & AsconVariant & ")"
        End Select
        .RoundsFinal = LNG_ROUNDS
        '--- init state
        For lIdx = 1 To UBound(.Words)
            .Words(lIdx) = 0
        Next
        lSize = UBound(baKey) + 1
        Call CopyMemory(.Bytes(LNG_STATESZ - LNG_NONCESZ - lSize), baKey(0), lSize)
        Call CopyMemory(.Bytes(LNG_STATESZ - LNG_NONCESZ), baNonce(0), LNG_NONCESZ)
        pvPermute uCtx, .RoundsFinal
        lSize = LNG_STATESZ - UBound(baKey) - 1
        For lIdx = 0 To UBound(baKey)
            .Bytes(lSize + lIdx) = .Bytes(lSize + lIdx) Xor baKey(lIdx)
        Next
        '--- process associated data
        If UBound(baAad) >= 0 Then
            pvUpdate uCtx, baAad, 0, UBound(baAad) + 1
            .Bytes(.Absorbed) = .Bytes(.Absorbed) Xor &H80
            pvPermute uCtx, .RoundsItermediate
            .Absorbed = 0
        End If
        '--- separator
        .Bytes(LNG_STATESZ - 1) = .Bytes(LNG_STATESZ - 1) Xor 1
    End With
End Sub

Private Sub pvUpdate(uCtx As CryptoAsconContext, baInput() As Byte, ByVal Pos As Long, ByVal Size As Long, Optional ByVal Encrypt As Boolean, Optional ByVal Decrypt As Boolean)
    Dim lIdx            As Long
    Dim lTemp           As Long
    
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    With uCtx
        For lIdx = 0 To Size - 1
            If Decrypt Then
                lTemp = .Bytes(.Absorbed) Xor baInput(Pos + lIdx)
                .Bytes(.Absorbed) = baInput(Pos + lIdx)
                baInput(Pos + lIdx) = lTemp
            Else
                .Bytes(.Absorbed) = .Bytes(.Absorbed) Xor baInput(Pos + lIdx)
                If Encrypt Then
                    baInput(Pos + lIdx) = .Bytes(.Absorbed)
                End If
            End If
            If .Absorbed = .Rate - 1 Then
                pvPermute uCtx, .RoundsItermediate
                .Absorbed = 0
            Else
                .Absorbed = .Absorbed + 1
            End If
        Next
    End With
End Sub

Private Sub pvFinalizeHash(uCtx As CryptoAsconContext, baOutput() As Byte, Optional ByVal OutSize As Long)
    Dim lIdx            As Long
    Dim pDummy          As LongPtr
    Dim uEmpty          As CryptoAsconContext
    
    With uCtx
        .Bytes(.Absorbed) = .Bytes(.Absorbed) Xor &H80
        pvPermute uCtx, .RoundsFinal
        .Absorbed = 0
        If OutSize = 0 Then
            OutSize = 32
        End If
        ReDim baOutput(0 To OutSize - 1) As Byte
        For lIdx = 0 To UBound(baOutput)
            baOutput(lIdx) = .Bytes(.Absorbed)
            If .Absorbed = .Rate - 1 Then
                pvPermute uCtx, .RoundsItermediate
                .Absorbed = 0
            Else
                .Absorbed = .Absorbed + 1
            End If
        Next
        Call CopyMemory(ByVal ArrPtr(.Bytes), pDummy, LenB(pDummy))
    End With
    uCtx = uEmpty
End Sub

Private Sub pvFinalizeAead(uCtx As CryptoAsconContext, baKey() As Byte, baTag() As Byte)
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim pDummy          As LongPtr
    Dim uEmpty          As CryptoAsconContext
    
    With uCtx
        .Bytes(.Absorbed) = .Bytes(.Absorbed) Xor &H80
        lSize = .Rate
        For lIdx = 0 To UBound(baKey)
            .Bytes(lSize + lIdx) = .Bytes(lSize + lIdx) Xor baKey(lIdx)
        Next
        pvPermute uCtx, .RoundsFinal
        lSize = LNG_STATESZ - UBound(baKey) - 1
        For lIdx = 0 To UBound(baKey)
            .Bytes(lSize + lIdx) = .Bytes(lSize + lIdx) Xor baKey(lIdx)
        Next
        ReDim baTag(0 To LNG_TAGSZ - 1) As Byte
        Call CopyMemory(baTag(0), .Bytes(LNG_STATESZ - LNG_TAGSZ), LNG_TAGSZ)
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

Public Sub CryptoAsconHashInit(uCtx As CryptoAsconContext, Optional AsconVariant As String)
    pvInitHash uCtx, AsconVariant
End Sub

Public Sub CryptoAsconHashUpdate(uCtx As CryptoAsconContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    pvUpdate uCtx, baInput, Pos, Size
End Sub

Public Sub CryptoAsconHashFinalize(uCtx As CryptoAsconContext, baOutput() As Byte, Optional ByVal OutSize As Long)
    pvFinalizeHash uCtx, baOutput, OutSize
End Sub

Public Function CryptoAsconHashByteArray(baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional AsconVariant As String) As Byte()
    Dim uCtx            As CryptoAsconContext
    
    pvInitHash uCtx, AsconVariant
    pvUpdate uCtx, baInput, Pos, Size
    pvFinalizeHash uCtx, CryptoAsconHashByteArray
End Function

Public Function CryptoAsconHashText(sText As String, Optional AsconVariant As String) As String
    CryptoAsconHashText = ToHex(CryptoAsconHashByteArray(ToUtf8Array(sText), AsconVariant:=AsconVariant))
End Function

Public Sub CryptoAsconEncrypt(baKey() As Byte, baTag() As Byte, _
            baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, _
            Optional Nonce As Variant, Optional AssociatedData As Variant, Optional AsconVariant As String)
    Dim uCtx            As CryptoAsconContext
    
    pvInitAead uCtx, baKey, Nonce, AssociatedData, AsconVariant
    pvUpdate uCtx, baInput, Pos, Size, Encrypt:=True
    pvFinalizeAead uCtx, baKey, baTag
End Sub

Public Function CryptoAsconDecrypt(baKey() As Byte, baTag() As Byte, _
            baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, _
            Optional Nonce As Variant, Optional AssociatedData As Variant, Optional AsconVariant As String) As Boolean
    Dim uCtx            As CryptoAsconContext
    Dim baTemp()        As Byte
    
    pvInitAead uCtx, baKey, Nonce, AssociatedData, AsconVariant
    pvUpdate uCtx, baInput, Pos, Size, Decrypt:=True
    pvFinalizeAead uCtx, baKey, baTemp
    If UBound(baTemp) = UBound(baTag) Then
        CryptoAsconDecrypt = (InStrB(baTemp, baTag) = 1)
    End If
End Function
