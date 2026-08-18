Attribute VB_Name = "mdAesOcb"
'=========================================================================
'
'  Pure VB6 Crypto (Untested)
'  Copyright (c) 2026 wqweto@gmail.com
'
'  This project is licensed under the terms of the MIT license
'  See the LICENSE file in the project root for more information
'
'=========================================================================
'--- mdAesOcb.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr
#End If

Private Const LNG_BLOCKSZ               As Long = 16

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As LongPtr
    cElements           As Long
    lLbound             As Long
End Type

Private Type ArrayLong4
    Item(0 To 3)        As Long
End Type

Public Type CryptoAesOcbContext
    AesCtx              As CryptoAesContext
    K1                  As ArrayLong4
    K2                  As ArrayLong4
    L()                 As ArrayLong4
    NumLookups          As Long
    Offset              As ArrayLong4
    Checksum            As ArrayLong4
    Sum                 As ArrayLong4
    NumBlocks           As Long
End Type

Private Sub pvShift(baInput() As Byte, ByVal lPos As Long, ByVal lBits As Long, uOutput As ArrayLong4)
    Dim lPow1           As Long
    Dim lPow2           As Long
    Dim lIdx            As Long
    Dim lNext           As Long
    Dim lCarry          As Long
    Dim baOutput(0 To LNG_BLOCKSZ - 1) As Byte
    
    lPow1 = 2 ^ (8 - lBits)
    lPow2 = 2 ^ lBits
    lCarry = baInput(lPos + LNG_BLOCKSZ) \ lPow1
    For lIdx = LNG_BLOCKSZ - 1 To 0 Step -1
        lNext = baInput(lPos + lIdx) \ lPow1
        baOutput(lIdx) = (baInput(lPos + lIdx) * lPow2) And &HFF Or lCarry
        lCarry = lNext
    Next
    Call CopyMemory(uOutput, baOutput(0), LNG_BLOCKSZ)
End Sub

Private Sub pvDouble(uInput As ArrayLong4, uOutput As ArrayLong4)
    Const LNG_POLY      As Long = &H87
    Dim baInput(0 To LNG_BLOCKSZ - 1) As Byte
    Dim baOutput(0 To LNG_BLOCKSZ - 1) As Byte
    Dim lIdx            As Long
    Dim lTemp           As Long
    Dim lCarry          As Long

    Call CopyMemory(baInput(0), uInput, LNG_BLOCKSZ)
    For lIdx = LNG_BLOCKSZ - 1 To 0 Step -1
        lTemp = baInput(lIdx)
        baOutput(lIdx) = (lTemp * 2) And &HFF Or lCarry
        lCarry = -((lTemp And &H80) <> 0)
    Next
    baOutput(LNG_BLOCKSZ - 1) = baOutput(LNG_BLOCKSZ - 1) Xor lCarry * LNG_POLY
    Call CopyMemory(uOutput, baOutput(0), LNG_BLOCKSZ)
End Sub

Private Function pvNtz(ByVal lBlock As Long) As Long
    '--- find first not-zero bit
    Do While (lBlock And 1) = 0
        pvNtz = pvNtz + 1
        lBlock = lBlock \ 2
    Loop
End Function

Private Sub pvLookupL(uCtx As CryptoAesOcbContext, ByVal lBlock As Long, uOutput As ArrayLong4)
    Dim lNtz            As Long
    
    lNtz = pvNtz(lBlock)
    With uCtx
        If lNtz > UBound(.L) Then
            ReDim Preserve .L(0 To lNtz + 3) As ArrayLong4
        End If
        Do While .NumLookups < lNtz
            pvDouble .L(.NumLookups), .L(.NumLookups + 1)
            .NumLookups = .NumLookups + 1
        Loop
        uOutput = .L(lNtz)
    End With
End Sub

Public Function pvProcess(uCtx As CryptoAesOcbContext, ByVal bDecrypt As Boolean, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long, ByVal lTagSize As Long, Tag As Variant) As Boolean
    Const FADF_AUTO     As Long = 1
    Dim aBlock()        As ArrayLong4
    Dim uPeekBlock      As SAFEARRAY1D
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim uTemp           As ArrayLong4
    Dim baCalcTag()     As Byte
    Dim uPad            As ArrayLong4
    Dim uChecksumPad    As ArrayLong4
    Dim lNtz            As Long
    
    If lSize < 0 Then
        lSize = UBound(baBuffer) + 1 - lPos
    End If
    With uCtx
        If lSize >= LNG_BLOCKSZ Then
            With uPeekBlock
                .cDims = 1
                .fFeatures = FADF_AUTO
                .cbElements = LNG_BLOCKSZ
                .cLocks = 1
                .pvData = VarPtr(baBuffer(lPos))
                .cElements = lSize \ .cbElements
            End With
            Call CopyMemory(ByVal ArrPtr(aBlock), VarPtr(uPeekBlock), 4)
        End If
        lIdx = .NumBlocks + lSize \ LNG_BLOCKSZ
        Do While lIdx > 0
            lIdx = lIdx \ 2
            lNtz = lNtz + 1
        Loop
        If lNtz > UBound(.L) Then
            ReDim Preserve .L(0 To lNtz + 3) As ArrayLong4
        End If
        Do While .NumLookups < lNtz
            pvDouble .L(.NumLookups), .L(.NumLookups + 1)
            .NumLookups = .NumLookups + 1
        Loop
        For lJdx = 0 To lSize \ LNG_BLOCKSZ - 1
            .NumBlocks = .NumBlocks + 1
            lNtz = pvNtz(.NumBlocks)
            .Offset.Item(0) = .Offset.Item(0) Xor .L(lNtz).Item(0)
            .Offset.Item(1) = .Offset.Item(1) Xor .L(lNtz).Item(1)
            .Offset.Item(2) = .Offset.Item(2) Xor .L(lNtz).Item(2)
            .Offset.Item(3) = .Offset.Item(3) Xor .L(lNtz).Item(3)
            If Not bDecrypt Then
                .Checksum.Item(0) = .Checksum.Item(0) Xor aBlock(lJdx).Item(0)
                .Checksum.Item(1) = .Checksum.Item(1) Xor aBlock(lJdx).Item(1)
                .Checksum.Item(2) = .Checksum.Item(2) Xor aBlock(lJdx).Item(2)
                .Checksum.Item(3) = .Checksum.Item(3) Xor aBlock(lJdx).Item(3)
            End If
            aBlock(lJdx).Item(0) = aBlock(lJdx).Item(0) Xor .Offset.Item(0)
            aBlock(lJdx).Item(1) = aBlock(lJdx).Item(1) Xor .Offset.Item(1)
            aBlock(lJdx).Item(2) = aBlock(lJdx).Item(2) Xor .Offset.Item(2)
            aBlock(lJdx).Item(3) = aBlock(lJdx).Item(3) Xor .Offset.Item(3)
            CryptoAesProcessPtr .AesCtx, VarPtr(aBlock(lJdx)), Decrypt:=bDecrypt
            aBlock(lJdx).Item(0) = aBlock(lJdx).Item(0) Xor .Offset.Item(0)
            aBlock(lJdx).Item(1) = aBlock(lJdx).Item(1) Xor .Offset.Item(1)
            aBlock(lJdx).Item(2) = aBlock(lJdx).Item(2) Xor .Offset.Item(2)
            aBlock(lJdx).Item(3) = aBlock(lJdx).Item(3) Xor .Offset.Item(3)
            If bDecrypt Then
                .Checksum.Item(0) = .Checksum.Item(0) Xor aBlock(lJdx).Item(0)
                .Checksum.Item(1) = .Checksum.Item(1) Xor aBlock(lJdx).Item(1)
                .Checksum.Item(2) = .Checksum.Item(2) Xor aBlock(lJdx).Item(2)
                .Checksum.Item(3) = .Checksum.Item(3) Xor aBlock(lJdx).Item(3)
            End If
            lPos = lPos + LNG_BLOCKSZ
        Next
        If lTagSize > 0 Then
            lSize = lSize Mod LNG_BLOCKSZ
            If lSize > 0 Then
                For lIdx = 0 To 3
                    .Offset.Item(lIdx) = .Offset.Item(lIdx) Xor .K1.Item(lIdx)
                Next
                uTemp = .Offset
                CryptoAesProcessPtr .AesCtx, VarPtr(uTemp)
                Call CopyMemory(uPad, baBuffer(lPos), lSize)
                Call CopyMemory(ByVal VarPtr(uPad) + lSize, &H80, 1)
                For lIdx = 0 To 3
                    If Not bDecrypt Then
                        uChecksumPad.Item(lIdx) = .Checksum.Item(lIdx) Xor uPad.Item(lIdx)
                    End If
                    uPad.Item(lIdx) = uPad.Item(lIdx) Xor uTemp.Item(lIdx)
                    If bDecrypt Then
                        Call CopyMemory(ByVal VarPtr(uPad) + lSize, &H80, 1)
                        uChecksumPad.Item(lIdx) = .Checksum.Item(lIdx) Xor uPad.Item(lIdx)
                    End If
                Next
                Call CopyMemory(baBuffer(lPos), uPad, lSize)
                Call CopyMemory(.Checksum, uChecksumPad, lSize + 1)
            End If
            pvFinalize uCtx, lTagSize, baCalcTag
            If bDecrypt Then
                If InStrB(baCalcTag, Tag) <> 1 Then
                    GoTo QH
                End If
            Else
                Tag = baCalcTag
            End If
        Else
            Debug.Assert lSize Mod 16 = 0
        End If
    End With
    '--- success
    pvProcess = True
QH:
End Function

Private Sub pvFinalize(uCtx As CryptoAesOcbContext, ByVal lTagSize As Long, baTag() As Byte)
    Dim uTemp           As ArrayLong4
    Dim lIdx            As Long
    
    With uCtx
        If lTagSize < 1 Or lTagSize > LNG_BLOCKSZ Then
            Err.Raise vbObjectError, , "Invalid tag size for AES-OCB (" & lTagSize & ")"
        End If
        For lIdx = 0 To 3
            uTemp.Item(lIdx) = .Offset.Item(lIdx) Xor .Checksum.Item(lIdx) Xor .K2.Item(lIdx)
        Next
        CryptoAesProcessPtr .AesCtx, VarPtr(uTemp)
        For lIdx = 0 To 3
            uTemp.Item(lIdx) = uTemp.Item(lIdx) Xor .Sum.Item(lIdx)
        Next
        ReDim baTag(0 To lTagSize - 1) As Byte
        Call CopyMemory(baTag(0), uTemp, lTagSize)
    End With
End Sub

Public Sub CryptoAesOcbInit(uCtx As CryptoAesOcbContext, baKey() As Byte, baNonce() As Byte, baAad() As Byte, Optional ByVal TagSize As Long = LNG_BLOCKSZ)
    Dim uEmpty          As ArrayLong4
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim lBottom         As Long
    Dim baKtop(0 To LNG_BLOCKSZ - 1) As Byte
    Dim baStretch(0 To LNG_BLOCKSZ + LNG_BLOCKSZ \ 2 - 1) As Byte
    Dim lJdx            As Long
    Dim lBlock          As Long
    Dim lPos            As Long
    Dim uLookup         As ArrayLong4
    Dim uOffset         As ArrayLong4
    Dim uTemp           As ArrayLong4
    Dim uAad            As ArrayLong4
    
    With uCtx
        CryptoAesInit .AesCtx, baKey
        .K1 = uEmpty
        CryptoAesProcessPtr .AesCtx, VarPtr(.K1)
        pvDouble .K1, .K2
        .NumLookups = 4
        ReDim .L(0 To .NumLookups) As ArrayLong4
        pvDouble .K2, .L(0)
        For lIdx = 1 To .NumLookups
            pvDouble .L(lIdx - 1), .L(lIdx)
        Next
        .Offset = uEmpty
        .Checksum = uEmpty
        .Sum = uEmpty
        .NumBlocks = 0
        '--- setup IV
        lSize = UBound(baNonce) + 1
        If lSize > LNG_BLOCKSZ - 1 Or lSize <= 0 Then
            Err.Raise vbObjectError, , "Invalid Nonce size for AES-OCB (" & lSize & ")"
        End If
        '--- Nonce = num2str(TAGLEN mod 128,7) || zeros(120-bitlen(N)) || 1 || N
        baKtop(0) = (TagSize * 8 Mod 128) * 2
        baKtop(LNG_BLOCKSZ - lSize - 1) = baKtop(LNG_BLOCKSZ - lSize - 1) Or 1
        Call CopyMemory(baKtop(LNG_BLOCKSZ - lSize), baNonce(0), lSize)
        '--- bottom = str2num(Nonce[123..128])
        lBottom = baKtop(LNG_BLOCKSZ - 1) And &H3F
        '--- Ktop = ENCIPHER(K, Nonce[1..122] || zeros(6))
        baKtop(LNG_BLOCKSZ - 1) = baKtop(LNG_BLOCKSZ - 1) And &HC0
        CryptoAesProcess .AesCtx, baKtop
        '--- Stretch = Ktop || (Ktop[1..64] xor Ktop[9..72])
        Call CopyMemory(baStretch(0), baKtop(0), LNG_BLOCKSZ)
        For lIdx = 0 To LNG_BLOCKSZ \ 2 - 1
            baStretch(lIdx + LNG_BLOCKSZ) = baKtop(lIdx) Xor baKtop(lIdx + 1)
        Next
        '--- Offset_0 = Stretch[1+bottom..128+bottom]
        pvShift baStretch, lBottom \ 8, lBottom Mod 8, .Offset
        '--- setup AAD
        lSize = UBound(baAad) + 1
        For lJdx = 0 To lSize \ LNG_BLOCKSZ - 1
            lBlock = lBlock + 1
            pvLookupL uCtx, lBlock, uLookup
            Call CopyMemory(uAad, baAad(lPos), LNG_BLOCKSZ)
            For lIdx = 0 To 3
                uOffset.Item(lIdx) = uOffset.Item(lIdx) Xor uLookup.Item(lIdx)
                uTemp.Item(lIdx) = uAad.Item(lIdx) Xor uOffset.Item(lIdx)
            Next
            CryptoAesProcessPtr .AesCtx, VarPtr(uTemp)
            For lIdx = 0 To 3
                .Sum.Item(lIdx) = .Sum.Item(lIdx) Xor uTemp.Item(lIdx)
            Next
            lPos = lPos + LNG_BLOCKSZ
        Next
        lSize = lSize Mod LNG_BLOCKSZ
        If lSize > 0 Then
            uAad = uEmpty
            Call CopyMemory(uAad, baAad(lPos), lSize)
            Call CopyMemory(ByVal VarPtr(uAad) + lSize, &H80, 1)
            For lIdx = 0 To 3
                uTemp.Item(lIdx) = uAad.Item(lIdx) Xor uOffset.Item(lIdx) Xor .K1.Item(lIdx)
            Next
            CryptoAesProcessPtr .AesCtx, VarPtr(uTemp)
            For lIdx = 0 To 3
                .Sum.Item(lIdx) = .Sum.Item(lIdx) Xor uTemp.Item(lIdx)
            Next
        End If
    End With
End Sub

Public Sub CryptoAesOcbEncrypt(uCtx As CryptoAesOcbContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional ByVal TagSize As Long, Optional Tag As Variant)
    pvProcess uCtx, False, baBuffer, Pos, Size, TagSize, Tag
End Sub

Public Function CryptoAesOcbDecrypt(uCtx As CryptoAesOcbContext, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1, Optional Tag As Variant) As Boolean
    CryptoAesOcbDecrypt = pvProcess(uCtx, True, baBuffer, Pos, Size, LNG_BLOCKSZ, Tag)
End Function
