Attribute VB_Name = "mdAesEax"
'--- mdAesEax.bas
Option Explicit
DefObj A-Z

Private Const LNG_BLOCKSZ               As Long = 16

Private Type ArrayByte16
    Item(0 To 15)       As Byte
End Type

Public Type CryptoCmacContext
    AesCtx              As CryptoAesContext
    K1                  As ArrayByte16
    K2                  As ArrayByte16
    HashArray           As ArrayByte16
    NPosition           As Long
End Type

'= CMAC ==================================================================

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

Public Sub CryptoCmacInit(uCtx As CryptoCmacContext, baKey() As Byte)
    Dim uEmpty          As ArrayByte16
    
    With uCtx
        .HashArray = uEmpty
        .NPosition = 0
        CryptoAesInit .AesCtx, baKey
        CryptoAesProcess .AesCtx, uEmpty.Item
        pvDouble uEmpty, .K1
        pvDouble .K1, .K2
    End With
End Sub

Public Sub CryptoCmacReset(uCtx As CryptoCmacContext)
    Dim uEmpty          As ArrayByte16
    
    With uCtx
        .HashArray = uEmpty
        .NPosition = 0
    End With
End Sub

Public Sub CryptoCmacUpdate(uCtx As CryptoCmacContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lIdx            As Long
    
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    With uCtx
        For lIdx = 0 To Size - 1
            If .NPosition = LNG_BLOCKSZ Then
                CryptoAesProcess .AesCtx, .HashArray.Item
                .NPosition = 0
            End If
            .HashArray.Item(.NPosition) = .HashArray.Item(.NPosition) Xor baInput(Pos + lIdx)
            .NPosition = .NPosition + 1
        Next
    End With
End Sub

Public Sub CryptoCmacFinalize(uCtx As CryptoCmacContext, baTag() As Byte, Optional ByVal TagSize As Long = LNG_BLOCKSZ)
    Dim uKey            As ArrayByte16
    Dim lIdx            As Long
    
    If TagSize < 4 Or TagSize > LNG_BLOCKSZ Then
        Err.Raise vbObjectError, , "Invalid tag size for CMAC (" & TagSize & ")"
    End If
    With uCtx
        If .NPosition = LNG_BLOCKSZ Then
            uKey = .K1
        Else
            .HashArray.Item(.NPosition) = .HashArray.Item(.NPosition) Xor &H80
            uKey = .K2
        End If
        For lIdx = 0 To LNG_BLOCKSZ - 1
            .HashArray.Item(lIdx) = .HashArray.Item(lIdx) Xor uKey.Item(lIdx)
        Next
        CryptoAesProcess .AesCtx, .HashArray.Item
        .NPosition = 0
        baTag = .HashArray.Item
        ReDim Preserve baTag(0 To TagSize - 1) As Byte
    End With
End Sub

'= AES-EAX ===============================================================

Private Sub pvMac(uCtx As CryptoCmacContext, ByVal lStep As Long, baInput() As Byte, baTag() As Byte)
    Dim uBlock          As ArrayByte16
    
    CryptoCmacReset uCtx
    uBlock.Item(LNG_BLOCKSZ - 1) = lStep
    CryptoCmacUpdate uCtx, uBlock.Item
    CryptoCmacUpdate uCtx, baInput
    CryptoCmacFinalize uCtx, baTag
End Sub

Public Sub CryptoAesEaxEncrypt(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baBuffer() As Byte, baTag() As Byte, Optional ByVal TagSize As Long = LNG_BLOCKSZ)
    Dim uAesCtx         As CryptoAesContext
    Dim uCtx            As CryptoCmacContext
    Dim baTagNonce()    As Byte
    Dim baTagAad()      As Byte
    Dim lIdx            As Long
    
    If TagSize < 4 Or TagSize > LNG_BLOCKSZ Then
        Err.Raise vbObjectError, , "Invalid tag size for EAX (" & TagSize & ")"
    End If
    CryptoCmacInit uCtx, baKey
    pvMac uCtx, 0, baNonce, baTagNonce
    pvMac uCtx, 1, baAad, baTagAad
    CryptoAesInit uAesCtx, baKey, baTagNonce
    CryptoAesCtrCrypt uAesCtx, baBuffer
    CryptoCmacInit uCtx, baKey
    pvMac uCtx, 2, baBuffer, baTag
    For lIdx = 0 To LNG_BLOCKSZ - 1
        baTag(lIdx) = baTag(lIdx) Xor baTagNonce(lIdx) Xor baTagAad(lIdx)
    Next
    ReDim Preserve baTag(0 To TagSize - 1) As Byte
End Sub

Public Function CryptoAesEaxDecrypt(baKey() As Byte, baNonce() As Byte, baAad() As Byte, baBuffer() As Byte, baTag() As Byte) As Boolean
    Dim uCtx            As CryptoCmacContext
    Dim baTagNonce()    As Byte
    Dim baTagAad()      As Byte
    Dim baCalc()        As Byte
    Dim uAesCtx         As CryptoAesContext
    Dim lIdx            As Long
    
    CryptoCmacInit uCtx, baKey
    pvMac uCtx, 0, baNonce, baTagNonce
    pvMac uCtx, 1, baAad, baTagAad
    pvMac uCtx, 2, baBuffer, baCalc
    For lIdx = 0 To LNG_BLOCKSZ - 1
        baCalc(lIdx) = baCalc(lIdx) Xor baTagNonce(lIdx) Xor baTagAad(lIdx)
    Next
    If UBound(baTag) <> UBound(baCalc) Then
        ReDim Preserve baCalc(0 To UBound(baTag)) As Byte
    End If
    If InStrB(baTag, baCalc) <> 1 Then
        Exit Function
    End If
    CryptoAesInit uAesCtx, baKey, baTagNonce
    CryptoAesCtrCrypt uAesCtx, baBuffer
    '--- success
    CryptoAesEaxDecrypt = True
End Function
