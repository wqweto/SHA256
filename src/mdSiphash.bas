Attribute VB_Name = "mdSiphash"
'=========================================================================
'
'  Pure VB6 Crypto (Untested)
'  Copyright (c) 2026 wqweto@gmail.com
'
'  This project is licensed under the terms of the MIT license
'  See the LICENSE file in the project root for more information
'
'=========================================================================
'--- mdSiphash.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_BLOCKSZ       As Long = 8
Private Const LNG_KEYSZ         As Long = 16

Public Type CryptoSiphashContext
#If HasPtrSafe Then
    V0                  As LongLong
    V1                  As LongLong
    V2                  As LongLong
    V3                  As LongLong
#Else
    V0                  As Variant
    V1                  As Variant
    V2                  As Variant
    V3                  As Variant
#End If
    Partial(0 To LNG_BLOCKSZ - 1) As Byte
    NPartial            As Long
    NInput              As Currency
    UpdateIters         As Long
    FinalizeIters       As Long
    OutSize             As Long
End Type

#If HasPtrSafe Then
#If Not HasOperators Then
    Private LNG_POW2(0 To 63)       As LongLong
    Private LNG_SIGN_BIT            As LongLong ' 2 ^ 63
#End If
    Private LNG_ZERO                As LongLong
    Private LNG_IV(0 To 3)          As LongLong
#Else
    Private LNG_POW2(0 To 63)       As Variant
    Private LNG_SIGN_BIT            As Variant
    Private LNG_ZERO                As Variant
    Private LNG_IV(0 To 3)          As Variant
#End If

#If Not HasOperators Then
#If HasPtrSafe Then
Private Function RotL64(ByVal lX As LongLong, ByVal lN As Long) As LongLong
#Else
Private Function RotL64(lX As Variant, ByVal lN As Long) As Variant
#End If
    '--- RotL64 = LShift(X, n) Or RShift(X, 64 - n)
    Debug.Assert lN <> 0
    RotL64 = ((lX And (LNG_POW2(63 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(63 - lN)) <> 0) * LNG_SIGN_BIT) Or _
        ((lX And (LNG_SIGN_BIT Xor -1)) \ LNG_POW2(64 - lN) Or -(lX < 0) * LNG_POW2(lN - 1))
End Function

#If HasPtrSafe Then
Private Function UAdd64(ByVal lX As LongLong, ByVal lY As LongLong) As LongLong
#Else
Private Function UAdd64(lX As Variant, lY As Variant) As Variant
#End If
    If (lX Xor lY) >= 0 Then
        UAdd64 = ((lX Xor LNG_SIGN_BIT) + lY) Xor LNG_SIGN_BIT
    Else
        UAdd64 = lX + lY
    End If
End Function

Private Sub pvCompress(uCtx As CryptoSiphashContext, ByVal lRounds As Long)
    With uCtx
        Do While lRounds > 0
            .V0 = UAdd64(.V0, .V1)
            .V2 = UAdd64(.V2, .V3)
            .V1 = RotL64(.V1, 13)
            .V3 = RotL64(.V3, 16)
            .V1 = .V1 Xor .V0
            .V3 = .V3 Xor .V2
            .V0 = RotL64(.V0, 32)
            
            .V2 = UAdd64(.V2, .V1)
            .V0 = UAdd64(.V0, .V3)
            .V1 = RotL64(.V1, 17)
            .V3 = RotL64(.V3, 21)
            .V1 = .V1 Xor .V2
            .V3 = .V3 Xor .V0
            .V2 = RotL64(.V2, 32)
            lRounds = lRounds - 1
        Loop
    End With
End Sub
#Else
[ IntegerOverflowChecks (False) ]
Private Sub pvCompress(uCtx As CryptoSiphashContext, ByVal lRounds As Long)
    With uCtx
        Do While lRounds > 0
            .V0 += .V1
            .V2 += .V3
            .V1 = (.V1 << 13) Or (.V1 >> 51)
            .V3 = (.V3 << 16) Or (.V3 >> 48)
            .V1 = .V1 Xor .V0
            .V3 = .V3 Xor .V2
            .V0 = (.V0 << 32) Or (.V0 >> 32)
            .V2 += .V1
            .V0 += .V3
            .V1 = (.V1 << 17) Or (.V1 >> 47)
            .V3 = (.V3 << 21) Or (.V3 >> 43)
            .V1 = .V1 Xor .V2
            .V3 = .V3 Xor .V0
            .V2 = (.V2 << 32) Or (.V2 >> 32)
            lRounds -= 1
        Loop
    End With
End Sub
#End If

#If Not HasPtrSafe Then
Private Function CLngLng(vValue As Variant) As Variant
    Const VT_I8 As Long = &H14
    Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
End Function
#End If

Public Sub CryptoSiphashInit(uCtx As CryptoSiphashContext, baKey() As Byte, _
            Optional ByVal UpdateIters As Long = 2, _
            Optional ByVal FinalizeIters As Long = 4, _
            Optional ByVal OutSize As Long = 8)
#If HasPtrSafe Then
    Static K(0 To 1)    As LongLong
#Else
    Static K(0 To 1)    As Variant
#End If
    Dim lIdx            As Long
    
    If LNG_IV(0) = 0 Then
        LNG_IV(0) = CLngLng("&H736f6d6570736575")
        LNG_IV(1) = CLngLng("&H646f72616e646f6d")
        LNG_IV(2) = CLngLng("&H6c7967656e657261")
        LNG_IV(3) = CLngLng("&H7465646279746573")
        LNG_ZERO = CLngLng(0)
        #If Not HasOperators Then
            LNG_POW2(0) = CLngLng(1)
            For lIdx = 1 To 63
                LNG_POW2(lIdx) = CVar(LNG_POW2(lIdx - 1)) * 2
            Next
            LNG_SIGN_BIT = LNG_POW2(63)
        #End If
    End If
    If UBound(baKey) + 1 < LNG_KEYSZ Then
        K(0) = LNG_ZERO: K(1) = LNG_ZERO
        #If HasPtrSafe Then
            If UBound(baKey) >= 0 Then
                Call CopyMemory(K(0), baKey(0), UBound(baKey) + 1)
            End If
        #Else
            lIdx = UBound(baKey) + 1
            If lIdx > 0 Then
                Call CopyMemory(ByVal VarPtr(K(0)) + 8, baKey(0), IIf(lIdx > 8, 8, lIdx))
            End If
            lIdx = UBound(baKey) - 7
            If lIdx > 0 Then
                Call CopyMemory(ByVal VarPtr(K(1)) + 8, baKey(8), lIdx)
            End If
        #End If
    Else
        #If HasPtrSafe Then
            Call CopyMemory(K(0), baKey(0), LNG_KEYSZ)
        #Else
            K(0) = LNG_ZERO: K(1) = LNG_ZERO
            Call CopyMemory(ByVal VarPtr(K(0)) + 8, baKey(0), 8)
            Call CopyMemory(ByVal VarPtr(K(1)) + 8, baKey(8), 8)
        #End If
    End If
    With uCtx
        If OutSize > 8 Then
            lIdx = &HEE
        Else
            lIdx = 0
        End If
        .V0 = LNG_IV(0) Xor K(0)
        .V1 = LNG_IV(1) Xor K(1) Xor lIdx
        .V2 = LNG_IV(2) Xor K(0)
        .V3 = LNG_IV(3) Xor K(1)
        .NPartial = 0
        .NInput = 0
        .UpdateIters = UpdateIters
        .FinalizeIters = FinalizeIters
        .OutSize = OutSize
    End With
End Sub

Public Sub CryptoSiphashUpdate(uCtx As CryptoSiphashContext, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
#If HasPtrSafe Then
    Static B            As LongLong
#Else
    Static B            As Variant
#End If
    Dim lIdx            As Long

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
                #If HasPtrSafe Then
                    Call CopyMemory(B, .Partial(0), LNG_BLOCKSZ)
                #Else
                    B = LNG_ZERO
                    Call CopyMemory(ByVal VarPtr(B) + 8, .Partial(0), LNG_BLOCKSZ)
                #End If
                .NPartial = 0
            ElseIf Size >= LNG_BLOCKSZ Then
                #If HasPtrSafe Then
                    Call CopyMemory(B, baInput(Pos), LNG_BLOCKSZ)
                #Else
                    B = LNG_ZERO
                    Call CopyMemory(ByVal VarPtr(B) + 8, baInput(Pos), LNG_BLOCKSZ)
                #End If
                Pos = Pos + LNG_BLOCKSZ
                Size = Size - LNG_BLOCKSZ
            Else
                Call CopyMemory(.Partial(0), baInput(Pos), Size)
                .NPartial = Size
                Exit Do
            End If
            .V3 = .V3 Xor B
            pvCompress uCtx, .UpdateIters
            .V0 = .V0 Xor B
        Loop
    End With
End Sub

Public Sub CryptoSiphashFinalize(uCtx As CryptoSiphashContext, baOutput() As Byte)
#If HasPtrSafe Then
    Static B            As LongLong
#Else
    Static B            As Variant
#End If
    Dim lIdx            As Long
    
    With uCtx
        ReDim baOutput(0 To .OutSize - 1) As Byte
        #If HasOperators Then
            B = CLngLng(.NInput) << 56
        #Else
            B = RotL64(CLngLng(.NInput) And &HFF, 56)
        #End If
        #If HasPtrSafe Then
            Call CopyMemory(B, .Partial(0), .NPartial)
        #Else
            Call CopyMemory(ByVal VarPtr(B) + 8, .Partial(0), .NPartial)
        #End If
        .V3 = .V3 Xor B
        pvCompress uCtx, .UpdateIters
        .V0 = .V0 Xor B
        If .OutSize > 8 Then
            lIdx = &HEE
        Else
            lIdx = &HFF
        End If
        .V2 = .V2 Xor lIdx
        pvCompress uCtx, .FinalizeIters
        B = .V0 Xor .V1 Xor .V2 Xor .V3
        If .OutSize < 8 Then
            lIdx = .OutSize
        Else
            lIdx = 8
        End If
        #If HasPtrSafe Then
            Call CopyMemory(baOutput(0), B, lIdx)
        #Else
            Call CopyMemory(baOutput(0), ByVal VarPtr(B) + 8, lIdx)
        #End If
        If .OutSize > 8 Then
            .V1 = .V1 Xor &HDD
            pvCompress uCtx, .FinalizeIters
            B = .V0 Xor .V1 Xor .V2 Xor .V3
            If .OutSize < 16 Then
                lIdx = .OutSize - 8
            Else
                lIdx = 8
            End If
            #If HasPtrSafe Then
                Call CopyMemory(baOutput(8), B, lIdx)
            #Else
                Call CopyMemory(baOutput(8), ByVal VarPtr(B) + 8, lIdx)
            #End If
        End If
    End With
End Sub

Public Function CryptoSiphash24ByteArray(baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSiphashContext
    
    CryptoSiphashInit uCtx, baKey, UpdateIters:=2, FinalizeIters:=4
    CryptoSiphashUpdate uCtx, baInput, Pos, Size
    CryptoSiphashFinalize uCtx, CryptoSiphash24ByteArray
End Function

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

Public Function CryptoSiphash24Text(sKey As String, sText As String) As String
    CryptoSiphash24Text = ToHex(CryptoSiphash24ByteArray(ToUtf8Array(sKey), ToUtf8Array(sText)))
End Function

Public Function CryptoSiphash13ByteArray(baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoSiphashContext
    
    CryptoSiphashInit uCtx, baKey, UpdateIters:=1, FinalizeIters:=3
    CryptoSiphashUpdate uCtx, baInput, Pos, Size
    CryptoSiphashFinalize uCtx, CryptoSiphash13ByteArray
End Function

Public Function CryptoSiphash13Text(sKey As String, sText As String) As String
    CryptoSiphash13Text = ToHex(CryptoSiphash13ByteArray(ToUtf8Array(sKey), ToUtf8Array(sText)))
End Function
