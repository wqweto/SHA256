Attribute VB_Name = "mdTea"
'--- mdTea.bas -- Wheeler & Needham’s Tiny Encryption Algorithm
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#End If

Private Const LNG_KEYSZ             As Long = 16
Private Const LNG_BLOCKSZ           As Long = 4
Private Const LNG_DELTA             As Long = &H9E3779B9
#If Not HasOperators Then
Private LNG_POW2(0 To 31)           As Long
Private m_bNoIntegerOverflowChecks  As Boolean

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

Private Function UAdd32(ByVal lX As Long, ByVal lY As Long) As Long
    If (lX Xor lY) >= 0 Then
        UAdd32 = ((lX Xor &H80000000) + lY) Xor &H80000000
    Else
        UAdd32 = lX + lY
    End If
End Function

Private Function pvGetOverflowIgnored(Optional bValue As Boolean = True) As Boolean
    Dim bInIde      As Boolean
    
    If Not bValue Then
        bValue = True
        pvGetOverflowIgnored = True
        Exit Function
    End If
    Debug.Assert pvGetOverflowIgnored(bInIde)
    If bInIde Then
        Exit Function
    End If
    On Error GoTo EH
    If &H8000 - 1 <> 0 Then
        pvGetOverflowIgnored = True
    End If
EH:
End Function

Private Sub pvInit()
    Dim lIdx            As Long
    
    If LNG_POW2(0) = 0 Then
        LNG_POW2(0) = 1
        For lIdx = 1 To 30
            LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
        Next
        LNG_POW2(31) = &H80000000
        m_bNoIntegerOverflowChecks = pvGetOverflowIgnored
    End If
End Sub
#End If

Public Sub CryptoTeaEncrypt(baKey() As Byte, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lIdx            As Long
    Dim aKey(0 To LNG_KEYSZ \ 4 - 1) As Long
    Dim aBuffer()       As Long
    Dim lRound          As Long
    Dim lN              As Long
    Dim lY              As Long
    Dim lZ              As Long
    Dim lMx             As Long
    Dim lE              As Long
    Dim lSum            As Long
    
    #If Not HasOperators Then
        pvInit
    #End If
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    If Size Mod LNG_BLOCKSZ <> 0 Then
        Err.Raise vbObjectError, , "Invalid block size for TEA (" & Size Mod LNG_BLOCKSZ & ")"
    End If
    ReDim aBuffer(0 To Size \ 4 + 1) As Long
    Call CopyMemory(aBuffer(0), baBuffer(Pos), Size)
    lIdx = UBound(baKey) + 1
    If lIdx > LNG_KEYSZ Then
        lIdx = LNG_KEYSZ
    End If
    Call CopyMemory(aKey(0), baKey(0), lIdx)
    lN = Size \ LNG_BLOCKSZ
    If lN < 2 Then
        lN = 2
    End If
    lZ = aBuffer(lN - 1)
    For lRound = 1 To 6 + 52 \ lN
        #If HasOperators Then
            lSum = lSum + LNG_DELTA
        #Else
            If m_bNoIntegerOverflowChecks Then
                lSum = lSum + LNG_DELTA
            Else
                lSum = UAdd32(lSum, LNG_DELTA)
            End If
        #End If
        lE = ((lSum And &HFFFF&) \ LNG_BLOCKSZ) And 3
        For lIdx = 0 To lN - 1
            lY = aBuffer((lIdx + 1) Mod lN)
            #If HasOperators Then
                lMx = (((lZ >> 5) Xor (lY << 2)) + ((lY >> 3) Xor (lZ << 4))) Xor ((lSum Xor lY) + (aKey((lIdx And 3) Xor lE) Xor lZ))
                lZ = aBuffer(lIdx) + lMx
            #Else
                If m_bNoIntegerOverflowChecks Then
                    lMx = ((RShift32(lZ, 5) Xor LShift32(lY, 2)) + (RShift32(lY, 3) Xor LShift32(lZ, 4))) Xor ((lSum Xor lY) + (aKey((lIdx And 3) Xor lE) Xor lZ))
                    lZ = aBuffer(lIdx) + lMx
                Else
                    lMx = UAdd32(RShift32(lZ, 5) Xor LShift32(lY, 2), RShift32(lY, 3) Xor LShift32(lZ, 4)) Xor UAdd32(lSum Xor lY, aKey(lIdx And 3 Xor lE) Xor lZ)
                    lZ = UAdd32(aBuffer(lIdx), lMx)
                End If
            #End If
            aBuffer(lIdx) = lZ
        Next
    Next
    Call CopyMemory(baBuffer(0), aBuffer(0), UBound(baBuffer) + 1)
End Sub

Public Sub CryptoTeaDecrypt(baKey() As Byte, baBuffer() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim lIdx            As Long
    Dim aKey(0 To LNG_KEYSZ \ 4 - 1) As Long
    Dim aBuffer()       As Long
    Dim lRound          As Long
    Dim lN              As Long
    Dim lY              As Long
    Dim lZ              As Long
    Dim lMx             As Long
    Dim lE              As Long
    Dim lSum            As Long
    
    #If Not HasOperators Then
        pvInit
    #End If
    If Size < 0 Then
        Size = UBound(baBuffer) + 1 - Pos
    End If
    If Size Mod LNG_BLOCKSZ <> 0 Then
        Err.Raise vbObjectError, , "Invalid block size for TEA (" & Size Mod LNG_BLOCKSZ & ")"
    End If
    ReDim aBuffer(0 To Size \ 4 + 1) As Long
    Call CopyMemory(aBuffer(0), baBuffer(Pos), Size)
    lIdx = UBound(baKey) + 1
    If lIdx > LNG_KEYSZ Then
        lIdx = LNG_KEYSZ
    End If
    Call CopyMemory(aKey(0), baKey(0), lIdx)
    lN = Size \ LNG_BLOCKSZ
    If lN < 2 Then
        lN = 2
    End If
    lY = aBuffer(0)
    #If HasOperators Then
        lSum = (6 + 52 \ lN) * LNG_DELTA
    #Else
        If m_bNoIntegerOverflowChecks Then
            lSum = (6 + 52 \ lN) * LNG_DELTA
        Else
            For lRound = 1 To 6 + 52 \ lN
                lSum = UAdd32(lSum, LNG_DELTA)
            Next
        End If
    #End If
    For lRound = 1 To 6 + 52 \ lN
        lE = ((lSum And &HFFFF&) \ LNG_BLOCKSZ) And 3
        For lIdx = lN - 1 To 0 Step -1
            lZ = aBuffer((lIdx + lN - 1) Mod lN)
            #If HasOperators Then
                lMx = (((lZ >> 5) Xor (lY << 2)) + ((lY >> 3) Xor (lZ << 4))) Xor ((lSum Xor lY) + (aKey((lIdx And 3) Xor lE) Xor lZ))
                lY = aBuffer(lIdx) - lMx
            #Else
                If m_bNoIntegerOverflowChecks Then
                    lMx = ((RShift32(lZ, 5) Xor LShift32(lY, 2)) + (RShift32(lY, 3) Xor LShift32(lZ, 4))) Xor ((lSum Xor lY) + (aKey((lIdx And 3) Xor lE) Xor lZ))
                    lY = aBuffer(lIdx) - lMx
                Else
                    lMx = UAdd32(RShift32(lZ, 5) Xor LShift32(lY, 2), RShift32(lY, 3) Xor LShift32(lZ, 4)) Xor UAdd32(lSum Xor lY, aKey(lIdx And 3 Xor lE) Xor lZ)
                    lY = UAdd32(aBuffer(lIdx), -lMx)
                End If
            #End If
            aBuffer(lIdx) = lY
        Next
        #If HasOperators Then
            lSum = lSum - LNG_DELTA
        #Else
            If m_bNoIntegerOverflowChecks Then
                lSum = lSum - LNG_DELTA
            Else
                lSum = UAdd32(lSum, -LNG_DELTA)
            End If
        #End If
    Next
    Call CopyMemory(baBuffer(0), aBuffer(0), UBound(baBuffer) + 1)
End Sub
