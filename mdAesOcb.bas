Attribute VB_Name = "mdAesOcb"
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

Private Type ArrayByte16
    Item(0 To 15)       As Byte
End Type

Public Type CryptoAesOcbContext
    AesCtx              As CryptoAesContext
    K1                  As ArrayByte16
    K2                  As ArrayByte16
    L()                 As ArrayByte16
    lCount              As Long
    Offset              As ArrayByte16
    Checksum            As ArrayByte16
    Sum                 As ArrayByte16
    NumBlocks           As Long
End Type

Private Sub pvShift(baInput() As Byte, ByVal lPos As Long, ByVal lBits As Long, uOutput As ArrayByte16)
    Dim lPow1           As Long
    Dim lPow2           As Long
    Dim lIdx            As Long
    Dim lNext           As Long
    Dim lCarry          As Long
        
    lPow1 = 2 ^ (8 - lBits)
    lPow2 = 2 ^ lBits
    lCarry = baInput(lPos + LNG_BLOCKSZ) \ lPow1
    For lIdx = LNG_BLOCKSZ - 1 To 0 Step -1
        lNext = baInput(lPos + lIdx) \ lPow1
        uOutput.Item(lIdx) = (baInput(lPos + lIdx) * lPow2) And &HFF Or lCarry
        lCarry = lNext
    Next
End Sub

Private Sub pvDouble(uInput As ArrayByte16, uOutput As ArrayByte16)
    Const LNG_POLY      As Long = &H87
    Dim lIdx            As Long
    Dim lTemp           As Long
    Dim lCarry          As Long

    For lIdx = LNG_BLOCKSZ - 1 To 0 Step -1
        lTemp = uInput.Item(lIdx)
        uOutput.Item(lIdx) = (lTemp * 2) And &HFF Or lCarry
        lCarry = -((lTemp And &H80) <> 0)
    Next
    uOutput.Item(LNG_BLOCKSZ - 1) = uOutput.Item(LNG_BLOCKSZ - 1) Xor lCarry * LNG_POLY
End Sub

Private Sub pvLookupL(uCtx As CryptoAesOcbContext, ByVal lBlock As Long, uOutput As ArrayByte16)
    Dim lNtz            As Long
    
    '--- find first not-zero bit
    Do While (lBlock And 1) = 0
        lNtz = lNtz + 1
        lBlock = lBlock \ 2
    Loop
    With uCtx
        If lNtz > UBound(.L) Then
            ReDim Preserve .L(0 To lNtz + 3) As ArrayByte16
        End If
        Do While .lCount < lNtz
            pvDouble .L(.lCount), .L(.lCount + 1)
            .lCount = .lCount + 1
        Loop
        uOutput = .L(lNtz)
    End With
End Sub

Public Function pvProcess(uCtx As CryptoAesOcbContext, ByVal bDecrypt As Boolean, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long, ByVal lTagSize As Long, Tag As Variant) As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim uLookup         As ArrayByte16
    Dim baTemp(0 To LNG_BLOCKSZ - 1) As Byte
    Dim baCalcTag()     As Byte
    
    If lSize < 0 Then
        lSize = UBound(baBuffer) + 1 - lPos
    End If
    With uCtx
        For lJdx = 0 To lSize \ LNG_BLOCKSZ - 1
            .NumBlocks = .NumBlocks + 1
            pvLookupL uCtx, .NumBlocks, uLookup
            For lIdx = 0 To LNG_BLOCKSZ - 1
                .Offset.Item(lIdx) = .Offset.Item(lIdx) Xor uLookup.Item(lIdx)
                If Not bDecrypt Then
                    .Checksum.Item(lIdx) = .Checksum.Item(lIdx) Xor baBuffer(lPos + lIdx)
                End If
                baBuffer(lPos + lIdx) = baBuffer(lPos + lIdx) Xor .Offset.Item(lIdx)
            Next
            CryptoAesProcessPtr .AesCtx, VarPtr(baBuffer(lPos)), Decrypt:=bDecrypt
            For lIdx = 0 To LNG_BLOCKSZ - 1
                baBuffer(lPos + lIdx) = baBuffer(lPos + lIdx) Xor .Offset.Item(lIdx)
                If bDecrypt Then
                    .Checksum.Item(lIdx) = .Checksum.Item(lIdx) Xor baBuffer(lPos + lIdx)
                End If
            Next
            lPos = lPos + LNG_BLOCKSZ
        Next
        If lTagSize > 0 Then
            lSize = lSize Mod LNG_BLOCKSZ
            If lSize > 0 Then
                For lIdx = 0 To LNG_BLOCKSZ - 1
                    .Offset.Item(lIdx) = .Offset.Item(lIdx) Xor .K1.Item(lIdx)
                Next
                Call CopyMemory(baTemp(0), .Offset.Item(0), LNG_BLOCKSZ)
                CryptoAesProcess .AesCtx, baTemp
                For lIdx = 0 To lSize - 1
                    If Not bDecrypt Then
                        .Checksum.Item(lIdx) = .Checksum.Item(lIdx) Xor baBuffer(lPos + lIdx)
                    End If
                    baBuffer(lPos + lIdx) = baBuffer(lPos + lIdx) Xor baTemp(lIdx)
                    If bDecrypt Then
                        .Checksum.Item(lIdx) = .Checksum.Item(lIdx) Xor baBuffer(lPos + lIdx)
                    End If
                Next
                .Checksum.Item(lSize) = .Checksum.Item(lSize) Xor &H80
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
    Dim baTemp(0 To LNG_BLOCKSZ - 1) As Byte
    Dim lIdx            As Long
    
    With uCtx
        If lTagSize < 1 Or lTagSize > LNG_BLOCKSZ Then
            Err.Raise vbObjectError, , "Invalid tag size for AES-OCB (" & lTagSize & ")"
        End If
        For lIdx = 0 To LNG_BLOCKSZ - 1
            baTemp(lIdx) = .Offset.Item(lIdx) Xor .Checksum.Item(lIdx) Xor .K2.Item(lIdx)
        Next
        CryptoAesProcess .AesCtx, baTemp
        ReDim baTag(0 To lTagSize - 1) As Byte
        For lIdx = 0 To lTagSize - 1
            baTag(lIdx) = baTemp(lIdx) Xor .Sum.Item(lIdx)
        Next
    End With
End Sub

Public Sub CryptoAesOcbInit(uCtx As CryptoAesOcbContext, baKey() As Byte, baNonce() As Byte, baAad() As Byte, Optional ByVal TagSize As Long = LNG_BLOCKSZ)
    Dim uEmpty          As ArrayByte16
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim lBottom         As Long
    Dim baKtop(0 To LNG_BLOCKSZ - 1) As Byte
    Dim baStretch(0 To LNG_BLOCKSZ + LNG_BLOCKSZ \ 2 - 1) As Byte
    Dim lJdx            As Long
    Dim lBlock          As Long
    Dim lPos            As Long
    Dim uLookup         As ArrayByte16
    Dim baOffset(0 To LNG_BLOCKSZ - 1) As Byte
    Dim baTemp(0 To LNG_BLOCKSZ - 1) As Byte
    
    With uCtx
        CryptoAesInit .AesCtx, baKey
        .K1 = uEmpty
        CryptoAesProcess .AesCtx, .K1.Item
        pvDouble .K1, .K2
        .lCount = 0
        ReDim .L(0 To .lCount) As ArrayByte16
        pvDouble .K2, .L(0)
        For lIdx = 1 To .lCount
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
            For lIdx = 0 To LNG_BLOCKSZ - 1
                baOffset(lIdx) = baOffset(lIdx) Xor uLookup.Item(lIdx)
                baTemp(lIdx) = baAad(lPos + lIdx) Xor baOffset(lIdx)
            Next
            CryptoAesProcess .AesCtx, baTemp
            For lIdx = 0 To LNG_BLOCKSZ - 1
                .Sum.Item(lIdx) = .Sum.Item(lIdx) Xor baTemp(lIdx)
            Next
            lPos = lPos + 16
        Next
        lSize = lSize Mod LNG_BLOCKSZ
        If lSize > 0 Then
            For lIdx = 0 To LNG_BLOCKSZ - 1
                If lIdx < lSize Then
                    baTemp(lIdx) = baAad(lPos + lIdx) Xor baOffset(lIdx) Xor .K1.Item(lIdx)
                Else
                    baTemp(lIdx) = baOffset(lIdx) Xor .K1.Item(lIdx)
                End If
            Next
            baTemp(lSize) = baTemp(lSize) Xor &H80
            CryptoAesProcess .AesCtx, baTemp
            For lIdx = 0 To LNG_BLOCKSZ - 1
                .Sum.Item(lIdx) = .Sum.Item(lIdx) Xor baTemp(lIdx)
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
