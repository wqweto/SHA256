Attribute VB_Name = "mdHalfSiphash"
'--- mdHalfSiphash.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Const LNG_BLOCKSZ       As Long = 4
Private Const LNG_KEYSZ         As Long = 8
Private Const LNG_POW2_4        As Long = 2 ^ 4
Private Const LNG_POW2_5        As Long = 2 ^ 5
Private Const LNG_POW2_6        As Long = 2 ^ 6
Private Const LNG_POW2_7        As Long = 2 ^ 7
Private Const LNG_POW2_8        As Long = 2 ^ 8
Private Const LNG_POW2_12       As Long = 2 ^ 12
Private Const LNG_POW2_13       As Long = 2 ^ 13
Private Const LNG_POW2_15       As Long = 2 ^ 15
Private Const LNG_POW2_16       As Long = 2 ^ 16
Private Const LNG_POW2_18       As Long = 2 ^ 18
Private Const LNG_POW2_19       As Long = 2 ^ 19
Private Const LNG_POW2_23       As Long = 2 ^ 23
Private Const LNG_POW2_24       As Long = 2 ^ 24
Private Const LNG_POW2_25       As Long = 2 ^ 25
Private Const LNG_POW2_26       As Long = 2 ^ 26
Private Const LNG_POW2_27       As Long = 2 ^ 27
Private Const LNG_POW2_31       As Long = &H80000000
    
Private Type ArrayLong128
    Item(0 To 127)      As Long
End Type

Public Type CryptoHalfSiphashContext
    V0                  As Long
    V1                  As Long
    V2                  As Long
    V3                  As Long
    Partial(0 To LNG_BLOCKSZ - 1) As Byte
    NPartial            As Long
    NInput              As Currency
    UpdateIters         As Long
    FinalizeIters       As Long
    OutSize             As Long
End Type

#If Not HasOperators Then
Private Sub pvCompress(uCtx As CryptoHalfSiphashContext, ByVal lRounds As Long)
    With uCtx
        Do While lRounds > 0
'            .V0 = UAdd32(.V0, .V1)
'            .V1 = RotL32(.V1, 5) Xor .V0
'            .V2 = UAdd32(.V2, .V3)
'            .V3 = RotL32(.V3, 8) Xor .V2
'            .V0 = RotL32(.V0, 16)
            If (.V0 Xor .V1) >= 0 Then
                .V0 = ((.V0 Xor &H80000000) + .V1) Xor &H80000000
            Else
                .V0 = .V0 + .V1
            End If
            .V1 = ((.V1 And (LNG_POW2_26 - 1)) * LNG_POW2_5 Or -((.V1 And LNG_POW2_26) <> 0) * LNG_POW2_31) Or _
                ((.V1 And (LNG_POW2_31 Xor -1)) \ LNG_POW2_27 Or -(.V1 < 0) * LNG_POW2_4) Xor .V0
            If (.V2 Xor .V3) >= 0 Then
                .V2 = ((.V2 Xor &H80000000) + .V3) Xor &H80000000
            Else
                .V2 = .V2 + .V3
            End If
            .V3 = ((.V3 And (LNG_POW2_23 - 1)) * LNG_POW2_8 Or -((.V3 And LNG_POW2_23) <> 0) * LNG_POW2_31) Or _
                ((.V3 And (LNG_POW2_31 Xor -1)) \ LNG_POW2_24 Or -(.V3 < 0) * LNG_POW2_7) Xor .V2
            .V0 = ((.V0 And (LNG_POW2_15 - 1)) * LNG_POW2_16 Or -((.V0 And LNG_POW2_15) <> 0) * LNG_POW2_31) Or _
                ((.V0 And (LNG_POW2_31 Xor -1)) \ LNG_POW2_16 Or -(.V0 < 0) * LNG_POW2_15)
            

'            .V2 = UAdd32(.V2, .V1)
'            .V1 = RotL32(.V1, 13) Xor .V2
'            .V0 = UAdd32(.V0, .V3)
'            .V3 = RotL32(.V3, 7) Xor .V0
'            .V2 = RotL32(.V2, 16)
            If (.V2 Xor .V1) >= 0 Then
                .V2 = ((.V2 Xor &H80000000) + .V1) Xor &H80000000
            Else
                .V2 = .V2 + .V1
            End If
            .V1 = ((.V1 And (LNG_POW2_18 - 1)) * LNG_POW2_13 Or -((.V1 And LNG_POW2_18) <> 0) * LNG_POW2_31) Or _
                ((.V1 And (LNG_POW2_31 Xor -1)) \ LNG_POW2_19 Or -(.V1 < 0) * LNG_POW2_12) Xor .V2
            If (.V0 Xor .V3) >= 0 Then
                .V0 = ((.V0 Xor &H80000000) + .V3) Xor &H80000000
            Else
                .V0 = .V0 + .V3
            End If
            .V3 = ((.V3 And (LNG_POW2_24 - 1)) * LNG_POW2_7 Or -((.V3 And LNG_POW2_24) <> 0) * LNG_POW2_31) Or _
                ((.V3 And (LNG_POW2_31 Xor -1)) \ LNG_POW2_25 Or -(.V3 < 0) * LNG_POW2_6) Xor .V0
            .V2 = ((.V2 And (LNG_POW2_15 - 1)) * LNG_POW2_16 Or -((.V2 And LNG_POW2_15) <> 0) * LNG_POW2_31) Or _
                ((.V2 And (LNG_POW2_31 Xor -1)) \ LNG_POW2_16 Or -(.V2 < 0) * LNG_POW2_15)

            lRounds = lRounds - 1
        Loop
    End With
End Sub
#Else
[ IntegerOverflowChecks (False) ]
Private Sub pvCompress(uCtx As CryptoHalfSiphashContext, ByVal lRounds As Long)
    With uCtx
        Do While lRounds > 0
            .V0 += .V1
            .V1 = ((.V1 << 5) Or (.V1 >> 27)) Xor .V0
            .V2 += .V3
            .V3 = ((.V3 << 8) Or (.V3 >> 24)) Xor .V2
            .V0 = (.V0 << 16) Or (.V0 >> 16)
            
            .V2 += .V1
            .V1 = ((.V1 << 13) Or (.V1 >> 19)) Xor .V2
            .V0 += .V3
            .V3 = ((.V3 << 7) Or (.V3 >> 25)) Xor .V0
            .V2 = (.V2 << 16) Or (.V2 >> 16)
            
            lRounds -= 1
        Loop
    End With
End Sub
#End If

Private Function pvCompressArray(uCtx As CryptoHalfSiphashContext, ByVal lSize As Long, uBlock As ArrayLong128, NotUsed As Long) As Long
    Dim lIdx            As Long
    
    With uCtx
        For lIdx = 0 To lSize - 1
            .V3 = .V3 Xor uBlock.Item(lIdx)
            pvCompress uCtx, .UpdateIters
            .V0 = .V0 Xor uBlock.Item(lIdx)
        Next
    End With
End Function

Public Sub CryptoHalfSiphashInit(uCtx As CryptoHalfSiphashContext, baKey() As Byte, _
            Optional ByVal UpdateIters As Long = 2, _
            Optional ByVal FinalizeIters As Long = 4, _
            Optional ByVal OutSize As Long = 4)
    Static K(0 To 1)    As Long
    Dim lIdx            As Long
    
    If UBound(baKey) + 1 < LNG_KEYSZ Then
        K(0) = 0: K(1) = 0
        If UBound(baKey) >= 0 Then
            Call CopyMemory(K(0), baKey(0), UBound(baKey) + 1)
        End If
    Else
        Call CopyMemory(K(0), baKey(0), LNG_KEYSZ)
    End If
    With uCtx
        If OutSize > 4 Then
            lIdx = &HEE
        Else
            lIdx = 0
        End If
        .V0 = K(0)
        .V1 = K(1) Xor lIdx
        .V2 = &H6C796765 Xor K(0)
        .V3 = &H74656462 Xor K(1)
        .NPartial = 0
        .NInput = 0
        .UpdateIters = UpdateIters
        .FinalizeIters = FinalizeIters
        .OutSize = OutSize
    End With
End Sub

Public Sub CryptoHalfSiphashUpdate(uCtx As CryptoHalfSiphashContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim B               As Long
    Dim lIdx            As Long
    Dim lBlocks         As Long

    With uCtx
        If Size < 0 Then
            Size = UBound(baInput) + 1 - Pos
        End If
        .NInput = .NInput + Size
        If .NPartial > 0 And Size > 0 Then
            lIdx = LNG_BLOCKSZ - .NPartial
            If lIdx > Size Then
                lIdx = Size
            End If
            Call CopyMemory(.Partial(.NPartial), baInput(Pos), lIdx)
            .NPartial = .NPartial + lIdx
            Pos = Pos + lIdx
            Size = Size - lIdx
        End If
        Do While Size > 0 Or .NPartial = LNG_BLOCKSZ
            If .NPartial <> 0 Then
                Call CopyMemory(B, .Partial(0), LNG_BLOCKSZ)
                .NPartial = 0
                .V3 = .V3 Xor B
                pvCompress uCtx, .UpdateIters
                .V0 = .V0 Xor B
            ElseIf Size >= 128 * LNG_BLOCKSZ Then
                lBlocks = Size \ (128 * LNG_BLOCKSZ)
                For lIdx = 0 To lBlocks - 1
                    Call CallWindowProc(AddressOf pvCompressArray, VarPtr(uCtx), 128, VarPtr(baInput(Pos)), VarPtr(0))
                    Pos = Pos + 128 * LNG_BLOCKSZ
                    Size = Size - 128 * LNG_BLOCKSZ
                Next
            ElseIf Size >= LNG_BLOCKSZ Then
                lIdx = Size \ LNG_BLOCKSZ
                Call CallWindowProc(AddressOf pvCompressArray, VarPtr(uCtx), lIdx, VarPtr(baInput(Pos)), VarPtr(0))
                Pos = Pos + lIdx * LNG_BLOCKSZ
                Size = Size - lIdx * LNG_BLOCKSZ
            Else
                Call CopyMemory(.Partial(0), baInput(Pos), Size)
                .NPartial = Size
                Exit Do
            End If
        Loop
    End With
End Sub

Public Sub CryptoHalfSiphashFinalize(uCtx As CryptoHalfSiphashContext, baOutput() As Byte)
    Dim B               As Long
    Dim lIdx            As Long
    
    With uCtx
        ReDim baOutput(0 To .OutSize - 1) As Byte
        #If HasOperators Then
            B = CLng(.NInput) << 24
        #Else
            B = .NInput And &HFF
'            B = RotL32(B, 24)
            B = ((B And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((B And LNG_POW2_7) <> 0) * LNG_POW2_31) Or _
                ((B And (LNG_POW2_31 Xor -1)) \ LNG_POW2_8 Or -(B < 0) * LNG_POW2_23)
        #End If
        Call CopyMemory(B, .Partial(0), .NPartial)
        .V3 = .V3 Xor B
        pvCompress uCtx, .UpdateIters
        .V0 = .V0 Xor B
        If .OutSize > 4 Then
            lIdx = &HEE
        Else
            lIdx = &HFF
        End If
        .V2 = .V2 Xor lIdx
        pvCompress uCtx, .FinalizeIters
        B = .V1 Xor .V3
        If .OutSize < 4 Then
            lIdx = .OutSize
        Else
            lIdx = 4
        End If
        Call CopyMemory(baOutput(0), B, lIdx)
        If .OutSize > 4 Then
            .V1 = .V1 Xor &HDD
            pvCompress uCtx, .FinalizeIters
            B = .V1 Xor .V3
            If .OutSize < 8 Then
                lIdx = .OutSize - 4
            Else
                lIdx = 4
            End If
            Call CopyMemory(baOutput(4), B, lIdx)
        End If
    End With
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

Public Function CryptoHalfSiphash24ByteArray(baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoHalfSiphashContext
    
    CryptoHalfSiphashInit uCtx, baKey, UpdateIters:=2, FinalizeIters:=4
    CryptoHalfSiphashUpdate uCtx, baInput, Pos, Size
    CryptoHalfSiphashFinalize uCtx, CryptoHalfSiphash24ByteArray
End Function

Public Function CryptoHalfSiphash24Long(baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Long
    Dim baOuput()       As Byte
    
    baOuput = CryptoHalfSiphash24ByteArray(baKey, baInput, Pos, Size)
    Call CopyMemory(CryptoHalfSiphash24Long, baOuput(0), 4)
End Function

Public Function CryptoHalfSiphash24Text(sKey As String, sText As String) As String
    CryptoHalfSiphash24Text = ToHex(CryptoHalfSiphash24ByteArray(ToUtf8Array(sKey), ToUtf8Array(sText)))
End Function

Public Function CryptoHalfSiphash13ByteArray(baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoHalfSiphashContext
    
    CryptoHalfSiphashInit uCtx, baKey, UpdateIters:=1, FinalizeIters:=3
    CryptoHalfSiphashUpdate uCtx, baInput, Pos, Size
    CryptoHalfSiphashFinalize uCtx, CryptoHalfSiphash13ByteArray
End Function

Public Function CryptoHalfSiphash13Long(baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Long
    Dim baOuput()       As Byte
    
    baOuput = CryptoHalfSiphash13ByteArray(baKey, baInput, Pos, Size)
    Call CopyMemory(CryptoHalfSiphash13Long, baOuput(0), 4)
End Function

Public Function CryptoHalfSiphash13Text(sKey As String, sText As String) As String
    CryptoHalfSiphash13Text = ToHex(CryptoHalfSiphash13ByteArray(ToUtf8Array(sKey), ToUtf8Array(sText)))
End Function
