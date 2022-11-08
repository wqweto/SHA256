Attribute VB_Name = "md5"
'--- md5.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const LNG_BLOCKSZ               As Long = 64
Private Const LNG_ROUNDS                As Long = 64

Public Type CryptoMd5Context
    H0                  As Long
    H1                  As Long
    H2                  As Long
    H3                  As Long
    Partial(0 To LNG_BLOCKSZ - 1) As Byte
    NPartial            As Long
    NInput              As Currency
End Type

Private LNG_POW2(0 To 31)           As Long
Private S(0 To 15)                  As Long
Private K(0 To LNG_ROUNDS - 1)      As Long

#If Not HasOperators Then
Private Function RotL32(ByVal lX As Long, ByVal lN As Long) As Long
    '--- RotL32 = LShift(X, n) Or RShift(X, 32 - n)
    Debug.Assert lN <> 0
    RotL32 = ((lX And (LNG_POW2(31 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(31 - lN)) <> 0) * LNG_POW2(31)) Or _
        ((lX And (LNG_POW2(31) Xor -1)) \ LNG_POW2(32 - lN) Or -(lX < 0) * LNG_POW2(lN - 1))
End Function

Private Function UAdd32(ByVal lX As Long, ByVal lY As Long) As Long
    If (lX Xor lY) >= 0 Then
        UAdd32 = ((lX Xor &H80000000) + lY) Xor &H80000000
    Else
        UAdd32 = lX + lY
    End If
End Function
#End If

Public Sub CryptoMd5Init(uCtx As CryptoMd5Context)
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    If S(0) = 0 Then
        For Each vElem In Split("7 12 17 22 5 9 14 20 4 11 16 23 6 10 15 21")
            S(lIdx) = vElem
            lIdx = lIdx + 1
        Next
        For lIdx = 0 To UBound(K)
            vElem = Abs(Sin(lIdx + 1)) * 4294967296#
            K(lIdx) = Int(IIf(vElem > 2147483648#, vElem - 4294967296#, vElem))
        Next
        #If Not HasOperators Then
            LNG_POW2(0) = 1
            For lIdx = 1 To 30
                LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
            Next
            LNG_POW2(31) = &H80000000
        #End If
    End If
    With uCtx
        .H0 = &H67452301: .H1 = &HEFCDAB89: .H2 = &H98BADCFE: .H3 = &H10325476
        .NPartial = 0
        .NInput = 0
    End With
End Sub

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
Public Sub CryptoMd5Update(uCtx As CryptoMd5Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Static B(0 To 15)   As Long
    Dim lIdx            As Long
    Dim lA              As Long
    Dim lB              As Long
    Dim lC              As Long
    Dim lD              As Long
    Dim lR              As Long
    Dim lE              As Long
    Dim lS              As Long
    Dim lTemp           As Long
    Dim lBufIdx         As Long
    
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
                Call CopyMemory(B(0), .Partial(0), LNG_BLOCKSZ)
                .NPartial = 0
            ElseIf Size >= LNG_BLOCKSZ Then
                Call CopyMemory(B(0), baInput(Pos), LNG_BLOCKSZ)
                Pos = Pos + LNG_BLOCKSZ
                Size = Size - LNG_BLOCKSZ
            Else
                Call CopyMemory(.Partial(0), baInput(Pos), Size)
                .NPartial = Size
                Exit Do
            End If
            '--- md5 step
            lA = .H0: lB = .H1: lC = .H2: lD = .H3
            For lIdx = 0 To LNG_ROUNDS - 1
                lR = lIdx \ 16
                Select Case lR
                Case 0
                    lE = (lB And lC) Or (Not lB And lD)
                    lBufIdx = lIdx
                Case 1
                    lE = (lB And lD) Or (lC And Not lD)
                    lBufIdx = (lIdx * 5 + 1) And 15
                Case 2
                    lE = lB Xor lC Xor lD
                    lBufIdx = (lIdx * 3 + 5) And 15
                Case 3
                    lE = lC Xor (lB Or Not lD)
                    lBufIdx = (lIdx * 7) And 15
                End Select
                lS = S((lR * 4) Or (lIdx And 3))
                #If HasOperators Then
                    lE += lA + K(lIdx) + B(lBufIdx)
                #Else
                    lE = UAdd32(UAdd32(UAdd32(lE, lA), K(lIdx)), B(lBufIdx))
                #End If
                lTemp = lD
                lD = lC
                lC = lB
                #If HasOperators Then
                    lB += (lE << lS) Or (lE >> (32 - lS))
                #Else
                    lB = UAdd32(lB, RotL32(lE, lS))
                #End If
                lA = lTemp
            Next
            #If HasOperators Then
                .H0 += lA: .H1 += lB: .H2 += lC: .H3 += lD
            #Else
                .H0 = UAdd32(.H0, lA): .H1 = UAdd32(.H1, lB): .H2 = UAdd32(.H2, lC): .H3 = UAdd32(.H3, lD)
            #End If
        Loop
    End With
End Sub

Public Sub CryptoMd5Finalize(uCtx As CryptoMd5Context, baOutput() As Byte)
    Dim P(0 To LNG_BLOCKSZ + 9) As Byte
    Dim lSize           As Long
    
    With uCtx
        lSize = LNG_BLOCKSZ - .NPartial
        If lSize < 9 Then
            lSize = lSize + LNG_BLOCKSZ
        End If
        P(0) = &H80
        .NInput = .NInput / 10000@ * 8
        Call CopyMemory(P(lSize - 8), .NInput, 8)
        CryptoMd5Update uCtx, P, Size:=lSize
        Debug.Assert .NPartial = 0
        ReDim baOutput(0 To 15) As Byte
        Call CopyMemory(baOutput(0), .H0, UBound(baOutput) + 1)
    End With
End Sub

Public Function CryptoMd5ByteArray(baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim uCtx            As CryptoMd5Context
    
    CryptoMd5Init uCtx
    CryptoMd5Update uCtx, baInput, Pos, Size
    CryptoMd5Finalize uCtx, CryptoMd5ByteArray
End Function

Private Function ToUtf8Array(sText As String) As Byte()
    Const CP_UTF8       As Long = 65001
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
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
        If Len(sByte) = 1 Then
            Mid$(ToHex, lIdx * 2 + 2, 1) = sByte
        Else
            Mid$(ToHex, lIdx * 2 + 1, 2) = sByte
        End If
    Next
End Function

Public Function CryptoMd5Text(sText As String) As String
    CryptoMd5Text = ToHex(CryptoMd5ByteArray(ToUtf8Array(sText)))
End Function
