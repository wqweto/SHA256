Attribute VB_Name = "mdChaCha20Poly1305"
'--- mdChaCha20Poly1305.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

Private Const LNG_KEYSZ                 As Long = 32
Private Const LNG_BLOCKSZ               As Long = 64
Private Const LNG_NONCESZ               As Long = 12
Private Const LNG_MACKEYSZ              As Long = 32
Private Const LNG_MACBLOCKSZ            As Long = 16
Private Const LNG_POW2_6                As Long = 2 ^ 6
Private Const LNG_POW2_7                As Long = 2 ^ 7
Private Const LNG_POW2_8                As Long = 2 ^ 8
Private Const LNG_POW2_11               As Long = 2 ^ 11
Private Const LNG_POW2_12               As Long = 2 ^ 12
Private Const LNG_POW2_15               As Long = 2 ^ 15
Private Const LNG_POW2_16               As Long = 2 ^ 16
Private Const LNG_POW2_19               As Long = 2 ^ 19
Private Const LNG_POW2_20               As Long = 2 ^ 20
Private Const LNG_POW2_23               As Long = 2 ^ 23
Private Const LNG_POW2_24               As Long = 2 ^ 24
Private Const LNG_POW2_25               As Long = 2 ^ 25
Private Const LNG_POW2_31               As Long = &H80000000

Private Type ArrayLong17
    Item(0 To 16)       As Long
End Type

Public Type CryptoChaCha20Context
    Constant(0 To 3)    As Long
    Key(0 To 7)         As Long
    Nonce(0 To 3)       As Long
    Block(0 To 63)      As Byte
    NBlock              As Long
    NCounter            As Long
End Type

Public Type CryptoPoly1305Context
    H                   As ArrayLong17
    R                   As ArrayLong17
    S(0 To 15)          As Byte
    Partial(0 To 15)    As Byte
    NPartial            As Long
End Type

#If Not HasOperators Then
Private Function UAdd32(ByVal lX As Long, ByVal lY As Long) As Long
    If (lX Xor lY) >= 0 Then
        UAdd32 = ((lX Xor &H80000000) + lY) Xor &H80000000
    Else
        UAdd32 = lX + lY
    End If
End Function

Private Sub pvChaCha20Quarter(lA As Long, lB As Long, lC As Long, lD As Long)
    If (lA Xor lB) >= 0 Then
        lA = ((lA Xor &H80000000) + lB) Xor &H80000000
    Else
        lA = lA + lB
    End If
    lD = lD Xor lA
    lD = ((lD And (LNG_POW2_15 - 1)) * LNG_POW2_16 Or -((lD And LNG_POW2_15) <> 0) * LNG_POW2_31) Or _
         ((lD And (LNG_POW2_31 Xor -1)) \ LNG_POW2_16 Or -(lD < 0) * LNG_POW2_15)
    If (lC Xor lD) >= 0 Then
        lC = ((lC Xor &H80000000) + lD) Xor &H80000000
    Else
        lC = lC + lD
    End If
    lB = lB Xor lC
    lB = ((lB And (LNG_POW2_19 - 1)) * LNG_POW2_12 Or -((lB And LNG_POW2_19) <> 0) * LNG_POW2_31) Or _
         ((lB And (LNG_POW2_31 Xor -1)) \ LNG_POW2_20 Or -(lB < 0) * LNG_POW2_11)
    If (lA Xor lB) >= 0 Then
        lA = ((lA Xor &H80000000) + lB) Xor &H80000000
    Else
        lA = lA + lB
    End If
    lD = lD Xor lA
    lD = ((lD And (LNG_POW2_23 - 1)) * LNG_POW2_8 Or -((lD And LNG_POW2_23) <> 0) * LNG_POW2_31) Or _
         ((lD And (LNG_POW2_31 Xor -1)) \ LNG_POW2_24 Or -(lD < 0) * LNG_POW2_7)
    If (lC Xor lD) >= 0 Then
        lC = ((lC Xor &H80000000) + lD) Xor &H80000000
    Else
        lC = lC + lD
    End If
    lB = lB Xor lC
    lB = ((lB And (LNG_POW2_24 - 1)) * LNG_POW2_7 Or -((lB And LNG_POW2_24) <> 0) * LNG_POW2_31) Or _
         ((lB And (LNG_POW2_31 Xor -1)) \ LNG_POW2_25 Or -(lB < 0) * LNG_POW2_6)
End Sub
#Else
[ IntegerOverflowChecks (False) ]
Private Sub pvChaCha20Quarter(lA As Long, lB As Long, lC As Long, lD As Long)
    lA += lB: lD = ((lD Xor lA) << 16) Or ((lD Xor lA) >> 16)
    lC += lD: lB = ((lB Xor lC) << 12) Or ((lB Xor lC) >> 20)
    lA += lB: lD = ((lD Xor lA) << 8) Or ((lD Xor lA) >> 24)
    lC += lD: lB = ((lB Xor lC) << 7) Or ((lB Xor lC) >> 25)
End Sub
#End If

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Private Sub pvChaCha20Core(uCtx As CryptoChaCha20Context, baOutput() As Byte)
    Static lZ(0 To 15)  As Long
    Static lX(0 To 15)  As Long
    Dim lIdx            As Long
    
    Call CopyMemory(lZ(0), uCtx.Constant(0), 16 * 4)
    Call CopyMemory(lX(0), uCtx.Constant(0), 16 * 4)
    For lIdx = 0 To 9
        pvChaCha20Quarter lZ(0), lZ(4), lZ(8), lZ(12)
        pvChaCha20Quarter lZ(1), lZ(5), lZ(9), lZ(13)
        pvChaCha20Quarter lZ(2), lZ(6), lZ(10), lZ(14)
        pvChaCha20Quarter lZ(3), lZ(7), lZ(11), lZ(15)
        pvChaCha20Quarter lZ(0), lZ(5), lZ(10), lZ(15)
        pvChaCha20Quarter lZ(1), lZ(6), lZ(11), lZ(12)
        pvChaCha20Quarter lZ(2), lZ(7), lZ(8), lZ(13)
        pvChaCha20Quarter lZ(3), lZ(4), lZ(9), lZ(14)
    Next
    For lIdx = 0 To 15
        #If Not HasOperators Then
            lX(lIdx) = UAdd32(lX(lIdx), lZ(lIdx))
        #Else
            lX(lIdx) += lZ(lIdx)
        #End If
    Next
    Call CopyMemory(baOutput(0), lX(0), 16 * 4)
End Sub

Public Sub CryptoChaCha20Init(uCtx As CryptoChaCha20Context, baKey() As Byte, baNonce() As Byte, Optional ByVal NCounter As Long = 4)
    Dim sConstant       As String
    Dim baFull(0 To 15) As Byte
    
    Debug.Assert UBound(baKey) + 1 = 16 Or UBound(baKey) + 1 = 32
    With uCtx
        If UBound(baKey) = 31 Then
            Call CopyMemory(.Key(0), baKey(0), 32)
            sConstant = "expand 32-byte k"
        Else
            Call CopyMemory(.Key(0), baKey(0), 16)
            Call CopyMemory(.Key(4), baKey(0), 16)
            sConstant = "expand 16-byte k"
        End If
        Call CopyMemory(.Constant(0), ByVal sConstant, Len(sConstant))
        If UBound(baNonce) >= UBound(baFull) Then
            Call CopyMemory(baFull(0), baNonce(0), UBound(baFull) + 1)
        ElseIf UBound(baNonce) >= 0 Then
            Call CopyMemory(baFull(15 - UBound(baNonce)), baNonce(0), UBound(baNonce) + 1)
        End If
        Call CopyMemory(.Nonce(0), baFull(0), 16)
        .NBlock = 0
        .NCounter = NCounter '--- part of Nonce that get incremented after pvChaCha20Core (in DWORDs)
    End With
End Sub

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Public Sub CryptoChaCha20Cipher(uCtx As CryptoChaCha20Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lOffset         As Long
    Dim lTaken          As Long
    Dim lIdx            As Long
    
    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        Do While Size > 0
            If .NBlock = 0 Then
                pvChaCha20Core uCtx, .Block
                For lIdx = 0 To .NCounter - 1
                    #If Not HasOperators Then
                        uCtx.Nonce(lIdx) = UAdd32(uCtx.Nonce(lIdx), 1)
                    #Else
                        uCtx.Nonce(lIdx) += 1
                    #End If
                    If uCtx.Nonce(lIdx) <> 0 Then
                        Exit For
                    End If
                Next
                .NBlock = LNG_BLOCKSZ
            End If
            lOffset = LNG_BLOCKSZ - .NBlock
            lTaken = .NBlock
            If Size < lTaken Then
                lTaken = Size
            End If
            For lIdx = 0 To lTaken - 1
                baInput(Pos) = baInput(Pos) Xor .Block(lOffset)
                Pos = Pos + 1
                lOffset = lOffset + 1
            Next
            .NBlock = .NBlock - lTaken
            Size = Size - lTaken
        Loop
    End With
End Sub

'= Poly1305 ==============================================================

Private Sub pvPoly1305Add(uX As ArrayLong17, uY As ArrayLong17)
    Dim lIdx            As Long
    Dim lCarry          As Long
    
    For lIdx = 0 To 16
        lCarry = lCarry + uX.Item(lIdx) + uY.Item(lIdx)
        uX.Item(lIdx) = lCarry And &HFF
        lCarry = lCarry \ &H100
    Next
End Sub

Private Sub pvPoly1305Mul(uX As ArrayLong17, uY As ArrayLong17)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lAccum          As Long
    Dim uR              As ArrayLong17
    
    For lIdx = 0 To 16
        For lJdx = 0 To 16
            If lJdx <= lIdx Then
                lAccum = lAccum + uX.Item(lJdx) * uY.Item(lIdx - lJdx)
            Else
                lAccum = lAccum + 320 * uX.Item(lJdx) * uY.Item(lIdx - lJdx + 17)
            End If
        Next
        uR.Item(lIdx) = lAccum
        lAccum = 0
    Next
    pvPoly1305MinReduce uR
    uX = uR
End Sub

Private Sub pvPoly1305MinReduce(uX As ArrayLong17)
    Dim lIdx            As Long
    Dim lCarry          As Long
    
    For lIdx = 0 To 15
        lCarry = lCarry + uX.Item(lIdx)
        uX.Item(lIdx) = lCarry And &HFF
        lCarry = lCarry \ &H100
    Next
    lCarry = lCarry + uX.Item(16)
    uX.Item(16) = lCarry And 3
    lCarry = 5 * (lCarry \ 4)
    For lIdx = 0 To 15
        lCarry = lCarry + uX.Item(lIdx)
        uX.Item(lIdx) = lCarry And &HFF
        lCarry = lCarry \ &H100
    Next
    uX.Item(16) = lCarry + uX.Item(16)
End Sub

Private Sub pvPoly1305FullReduce(uX As ArrayLong17)
    Dim lIdx            As Long
    Dim uSub            As ArrayLong17
    Dim uNeg            As ArrayLong17 '-> -(2^130-5)
    Dim lMask           As Long
    
    uSub = uX
    uNeg.Item(0) = 5
    uNeg.Item(16) = &HFC
    pvPoly1305Add uSub, uNeg
    lMask = (uSub.Item(16) And &H80) <> 0
    For lIdx = 0 To 16
        uX.Item(lIdx) = (uX.Item(lIdx) And lMask) Or (uSub.Item(lIdx) And Not lMask)
    Next
End Sub

Private Sub pvPoly1305Block(uCtx As CryptoPoly1305Context, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Dim lIdx            As Long
    Dim uX              As ArrayLong17
    
    For lIdx = 0 To lSize - 1
        uX.Item(lIdx) = baBuffer(lPos + lIdx)
    Next
    uX.Item(lSize) = 1
    pvPoly1305Add uCtx.H, uX
    pvPoly1305Mul uCtx.H, uCtx.R
End Sub

Public Sub CryptoPoly1305Init(uCtx As CryptoPoly1305Context, baKey() As Byte)
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) + 1 = LNG_KEYSZ
    With uCtx
        For lIdx = 0 To UBound(.H.Item)
            .H.Item(lIdx) = 0
            Select Case lIdx
            Case 3, 7, 11, 15
                .R.Item(lIdx) = baKey(lIdx) And &HF
            Case 4, 8, 12
                .R.Item(lIdx) = baKey(lIdx) And &HFC
            Case 16
                .R.Item(lIdx) = 0
            Case Else
                .R.Item(lIdx) = baKey(lIdx)
            End Select
        Next
        Call CopyMemory(.S(0), baKey(LNG_KEYSZ \ 2), LNG_KEYSZ \ 2)
        .NPartial = 0
    End With
End Sub

Public Sub CryptoPoly1305Update(uCtx As CryptoPoly1305Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lTaken          As Long
    
    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        If .NPartial > 0 And Size > 0 Then
            lTaken = LNG_MACBLOCKSZ - .NPartial
            If lTaken > Size Then
                lTaken = Size
            End If
            Call CopyMemory(.Partial(.NPartial), baInput(Pos), lTaken)
            Pos = Pos + lTaken
            Size = Size - lTaken
            .NPartial = .NPartial + lTaken
            If .NPartial = LNG_MACBLOCKSZ Then
                pvPoly1305Block uCtx, .Partial, 0, .NPartial
                .NPartial = 0
            End If
        End If
        Do While Size >= LNG_MACBLOCKSZ
            Debug.Assert .NPartial = 0
            pvPoly1305Block uCtx, baInput, Pos, LNG_MACBLOCKSZ
            Pos = Pos + LNG_MACBLOCKSZ
            Size = Size - LNG_MACBLOCKSZ
        Loop
        If Size > 0 Then
            lTaken = LNG_MACBLOCKSZ - .NPartial
            If lTaken > Size Then
                lTaken = Size
            End If
            Call CopyMemory(.Partial(.NPartial), baInput(Pos), lTaken)
            .NPartial = .NPartial + lTaken
            Debug.Assert .NPartial < LNG_MACBLOCKSZ
        End If
    End With
End Sub

Public Sub CryptoPoly1305Finalize(uCtx As CryptoPoly1305Context, baOutput() As Byte)
    Dim lIdx            As Long
    Dim uX              As ArrayLong17
    
    With uCtx
        If .NPartial > 0 Then
            pvPoly1305Block uCtx, .Partial, 0, .NPartial
        End If
        For lIdx = 0 To LNG_MACBLOCKSZ - 1
            uX.Item(lIdx) = .S(lIdx)
        Next
        pvPoly1305FullReduce .H
        pvPoly1305Add .H, uX
        ReDim baOutput(0 To LNG_MACBLOCKSZ - 1) As Byte
        For lIdx = 0 To LNG_MACBLOCKSZ - 1
            baOutput(lIdx) = .H.Item(lIdx)
        Next
    End With
End Sub

'= ChaCha20Poly1305 ======================================================

Private Function Process(baKey() As Byte, Nonce As Variant, AssociatedData As Variant, baTag() As Byte, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long, ByVal Encrypt As Boolean) As Boolean
    Dim uChaCha         As CryptoChaCha20Context
    Dim uPoly           As CryptoPoly1305Context
    Dim baNonce()       As Byte
    Dim baAad()         As Byte
    Dim baMacKey(0 To LNG_MACKEYSZ - 1) As Byte
    Dim baPad(0 To LNG_MACBLOCKSZ - 1) As Byte
    Dim baTemp()        As Byte
    
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
    If lSize < 0 Then
        lSize = UBound(baBuffer) + 1 - lPos
    End If
    CryptoChaCha20Init uChaCha, baKey, baNonce, 1
    CryptoChaCha20Cipher uChaCha, baMacKey
    CryptoPoly1305Init uPoly, baMacKey
    '--- discard 32 bytes from chacha20 key stream
    CryptoChaCha20Cipher uChaCha, baMacKey
    If Encrypt Then
        '--- encrypt then MAC
        CryptoChaCha20Cipher uChaCha, baBuffer, Pos:=lPos, Size:=lSize
    End If
    '--- ADD || pad(AAD)
    CryptoPoly1305Update uPoly, baAad
    CryptoPoly1305Update uPoly, baPad, Size:=(LNG_MACBLOCKSZ - (UBound(baAad) + 1) And (LNG_MACBLOCKSZ - 1)) And (LNG_MACBLOCKSZ - 1)
    '--- cipher || pad(cipher)
    CryptoPoly1305Update uPoly, baBuffer, Pos:=lPos, Size:=lSize
    CryptoPoly1305Update uPoly, baPad, Size:=(LNG_MACBLOCKSZ - lSize And (LNG_MACBLOCKSZ - 1)) And (LNG_MACBLOCKSZ - 1)
    '--- len_64(aad) || len_64(cipher)
    Call CopyMemory(baPad(0), UBound(baAad) + 1, 4)
    Call CopyMemory(baPad(8), lSize, 4)
    CryptoPoly1305Update uPoly, baPad
    '--- MAC complete
    If Encrypt Then
        CryptoPoly1305Finalize uPoly, baTag
    Else
        CryptoPoly1305Finalize uPoly, baTemp
        '--- decrypt only if tag matches
        If UBound(baTag) <> UBound(baTemp) Or InStrB(baTag, baTemp) <> 1 Then
            GoTo QH
        End If
        CryptoChaCha20Cipher uChaCha, baBuffer, Pos:=lPos, Size:=lSize
    End If
    '--- success
    Process = True
QH:
End Function

Public Function CryptoChaCha20Poly1305Encrypt(baKey() As Byte, baTag() As Byte, _
            baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, _
            Optional Nonce As Variant, Optional AssociatedData As Variant) As Boolean
    CryptoChaCha20Poly1305Encrypt = Process(baKey, Nonce, AssociatedData, baTag, baBuffer, Pos, Size, Encrypt:=True)
End Function

Public Function CryptoChaCha20Poly1305Decrypt(baKey() As Byte, baTag() As Byte, _
            baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, _
            Optional Nonce As Variant, Optional AssociatedData As Variant) As Boolean
    CryptoChaCha20Poly1305Decrypt = Process(baKey, Nonce, AssociatedData, baTag, baBuffer, Pos, Size, Encrypt:=False)
End Function
