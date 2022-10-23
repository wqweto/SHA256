Attribute VB_Name = "mdChaCha20Poly1305"
'--- mdChaCha20Poly1305.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

Private LNG_POW2(0 To 31)   As Long

Public Type CryptoChaCha20Context
    Constant(0 To 3)    As Long
    Key(0 To 7)         As Long
    Nonce(0 To 3)       As Long
    Block(0 To 63)      As Byte
    NBlock              As Long
    NCounter            As Long
End Type

Private Type FieldElement
    Item(0 To 16)       As Long
End Type

Public Type CryptoPoly1305Context
    H                   As FieldElement
    R                   As FieldElement
    S(0 To 15)          As Byte
    Partial(0 To 15)    As Byte
    NPartial            As Long
End Type

Private Function ROTL32(ByVal lX As Long, ByVal lN As Long) As Long
    '--- ROTL32 = LShift(X, n) Or RShift(X, 32 - n)
    Debug.Assert lN <> 0
    ROTL32 = ((lX And (LNG_POW2(31 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(31 - lN)) <> 0) * LNG_POW2(31)) Or _
        ((lX And (LNG_POW2(31) Xor -1)) \ LNG_POW2(32 - lN) Or -(lX < 0) * LNG_POW2(lN - 1))
End Function

Private Function UAdd(ByVal lX As Long, ByVal lY As Long) As Long
    If (lX Xor lY) >= 0 Then
        UAdd = ((lX Xor &H80000000) + lY) Xor &H80000000
    Else
        UAdd = lX + lY
    End If
End Function

Private Sub pvInit()
    Dim lIdx            As Long
    
    If LNG_POW2(0) = 0 Then
        LNG_POW2(0) = 1
        For lIdx = 1 To 30
            LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
        Next
        LNG_POW2(31) = &H80000000
    End If
End Sub

Private Sub pvChaCha20Quarter(lA As Long, lB As Long, lC As Long, lD As Long)
    lA = UAdd(lA, lB): lD = ROTL32(lD Xor lA, 16)
    lC = UAdd(lC, lD): lB = ROTL32(lB Xor lC, 12)
    lA = UAdd(lA, lB): lD = ROTL32(lD Xor lA, 8)
    lC = UAdd(lC, lD): lB = ROTL32(lB Xor lC, 7)
End Sub

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
        lX(lIdx) = UAdd(lX(lIdx), lZ(lIdx))
    Next
    Call CopyMemory(baOutput(0), lX(0), 16 * 4)
End Sub

Public Sub CryptoChaCha20Init(uCtx As CryptoChaCha20Context, baKey() As Byte, baNonce() As Byte, Optional ByVal NCounter As Long = 4)
    Dim sConstant       As String
    Dim baFull(0 To 15) As Byte
    
    Debug.Assert UBound(baKey) + 1 = 16 Or UBound(baKey) + 1 = 32
    With uCtx
        pvInit
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

Public Sub CryptoChaCha20Cipher(uCtx As CryptoChaCha20Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Const BLOCKSZ       As Long = 64
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
                    uCtx.Nonce(lIdx) = UAdd(uCtx.Nonce(lIdx), 1)
                    If uCtx.Nonce(lIdx) <> 0 Then
                        Exit For
                    End If
                Next
                .NBlock = BLOCKSZ
            End If
            lOffset = BLOCKSZ - .NBlock
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

Private Sub pvPoly1305Add(uX As FieldElement, uY As FieldElement)
    Dim lIdx            As Long
    Dim lCarry          As Long
    
    For lIdx = 0 To 16
        lCarry = lCarry + uX.Item(lIdx) + uY.Item(lIdx)
        uX.Item(lIdx) = lCarry And &HFF
        lCarry = lCarry \ &H100
    Next
End Sub

Private Sub pvPoly1305Mul(uX As FieldElement, uY As FieldElement)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lAccum          As Long
    Dim uR              As FieldElement
    
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

Private Sub pvPoly1305MinReduce(uX As FieldElement)
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

Private Sub pvPoly1305FullReduce(uX As FieldElement)
    Dim lIdx            As Long
    Dim uSub            As FieldElement
    Dim uNeg            As FieldElement '-> -(2^130-5)
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
    Dim uX              As FieldElement
    
    For lIdx = 0 To lSize - 1
        uX.Item(lIdx) = baBuffer(lPos + lIdx)
    Next
    uX.Item(lSize) = 1
    pvPoly1305Add uCtx.H, uX
    pvPoly1305Mul uCtx.H, uCtx.R
End Sub

Public Sub CryptoPoly1305Init(uCtx As CryptoPoly1305Context, baKey() As Byte)
    Const KEYSZ         As Long = 32
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) + 1 = KEYSZ
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
        Call CopyMemory(.S(0), baKey(KEYSZ \ 2), KEYSZ \ 2)
        .NPartial = 0
    End With
End Sub

Public Sub CryptoPoly1305Update(uCtx As CryptoPoly1305Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Const BLOCKSZ       As Long = 16
    Dim lTaken          As Long
    
    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        If .NPartial > 0 Then
            lTaken = BLOCKSZ - .NPartial
            If lTaken > Size Then
                lTaken = Size
            End If
            Call CopyMemory(.Partial(.NPartial), baInput(Pos), lTaken)
            Pos = Pos + lTaken
            Size = Size - lTaken
            .NPartial = .NPartial + lTaken
            If .NPartial = BLOCKSZ Then
                pvPoly1305Block uCtx, .Partial, 0, .NPartial
                .NPartial = 0
            End If
        End If
        Do While Size >= BLOCKSZ
            Debug.Assert .NPartial = 0
            pvPoly1305Block uCtx, baInput, Pos, BLOCKSZ
            Pos = Pos + BLOCKSZ
            Size = Size - BLOCKSZ
        Loop
        If Size > 0 Then
            lTaken = BLOCKSZ - .NPartial
            If lTaken > Size Then
                lTaken = Size
            End If
            Call CopyMemory(.Partial(.NPartial), baInput(Pos), lTaken)
            .NPartial = .NPartial + lTaken
            Debug.Assert .NPartial < BLOCKSZ
        End If
    End With
End Sub

Public Sub CryptoPoly1305Finish(uCtx As CryptoPoly1305Context, baOutput() As Byte)
    Const BLOCKSZ       As Long = 16
    Dim lIdx            As Long
    Dim uX              As FieldElement
    
    With uCtx
        If .NPartial > 0 Then
            pvPoly1305Block uCtx, .Partial, 0, .NPartial
        End If
        For lIdx = 0 To BLOCKSZ - 1
            uX.Item(lIdx) = .S(lIdx)
        Next
        pvPoly1305FullReduce .H
        pvPoly1305Add .H, uX
        ReDim baOutput(0 To BLOCKSZ - 1) As Byte
        For lIdx = 0 To BLOCKSZ - 1
            baOutput(lIdx) = .H.Item(lIdx)
        Next
    End With
End Sub

'= ChaCha20Poly130 =======================================================

Private Function Process(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baTag() As Byte, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long, ByVal Encrypt As Boolean) As Boolean
    Dim uChaCha         As CryptoChaCha20Context
    Dim uPoly           As CryptoPoly1305Context
    Dim baPolyKey(0 To 31) As Byte
    Dim baPad(0 To 15)  As Byte
    Dim baTemp()        As Byte
    
    If UBound(baNonce) + 1 <> 12 Then
        GoTo QH
    End If
    If lSize < 0 Then
        lSize = UBound(baBuffer) + 1 - lPos
    End If
    CryptoChaCha20Init uChaCha, baKey, baNonce, 1
    CryptoChaCha20Cipher uChaCha, baPolyKey
    CryptoPoly1305Init uPoly, baPolyKey
    '--- discard 32 bytes from chacha20 key stream
    CryptoChaCha20Cipher uChaCha, baPolyKey
    If Encrypt Then
        '--- encrypt then MAC
        CryptoChaCha20Cipher uChaCha, baBuffer, Pos:=lPos, Size:=lSize
    End If
    '--- ADD || pad(AAD)
    CryptoPoly1305Update uPoly, baAad
    CryptoPoly1305Update uPoly, baPad, Size:=(16 - (UBound(baAad) + 1) And 15) And 15
    '--- cipher || pad(cipher)
    CryptoPoly1305Update uPoly, baBuffer, Pos:=lPos, Size:=lSize
    CryptoPoly1305Update uPoly, baPad, Size:=(16 - lSize And 15) And 15
    '--- len_64(aad) || len_64(cipher)
    Call CopyMemory(baPad(0), UBound(baAad) + 1, 4)
    Call CopyMemory(baPad(8), lSize, 4)
    CryptoPoly1305Update uPoly, baPad
    '--- MAC complete
    If Encrypt Then
        CryptoPoly1305Finish uPoly, baTag
    Else
        CryptoPoly1305Finish uPoly, baTemp
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

Public Function CryptoChaCha20Poly1305Encrypt(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baTag() As Byte, _
            baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    CryptoChaCha20Poly1305Encrypt = Process(baKey, baNonce, baAad, baTag, baBuffer, Pos, Size, Encrypt:=True)
End Function

Public Function CryptoChaCha20Poly1305Decrypt(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baTag() As Byte, _
            baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    CryptoChaCha20Poly1305Decrypt = Process(baKey, baNonce, baAad, baTag, baBuffer, Pos, Size, Encrypt:=False)
End Function
