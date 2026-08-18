Attribute VB_Name = "mdCurve25519"
'--- mdCurve25519.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)
#Const HasSha512 = (CRYPT_HAS_SHA512 <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long
#End If

Private Const LNG_ELEMSZ            As Long = 16
Private Const LNG_KEYSZ             As Long = 32
Private Const LNG_HASHSZ            As Long = 64 '--- SHA-512
Private Const LNG_HALFHASHSZ        As Long = LNG_HASHSZ \ 2
Private Const LNG_POW16             As Long = 2 ^ 16

#If HasPtrSafe Then
    Private m_lZero             As LongLong
#Else
    Private m_lZero             As Variant
#End If
Private LNG_POW2(0 To 7)        As Long
Private EmptyByteArray()        As Byte
Private m_gf0                   As GF25519Element
Private m_gf1                   As GF25519Element
Private m_gfD                   As GF25519Element
Private m_gfD2                  As GF25519Element
Private m_gfX                   As GF25519Element
Private m_gfY                   As GF25519Element
Private m_gfI                   As GF25519Element
Private m_aL                    As ArrayLong64

Private Type GF25519Element
#If HasPtrSafe Then
    Item(0 To LNG_ELEMSZ - 1) As LongLong
#Else
    Item(0 To LNG_ELEMSZ - 1) As Variant
#End If
End Type

Private Type XyztPoint
    gfX                     As GF25519Element
    gfY                     As GF25519Element
    gfZ                     As GF25519Element
    gfT                     As GF25519Element
End Type

Private Type ArrayLong64
#If HasPtrSafe Then
    Item(0 To 63)           As LongLong
#Else
    Item(0 To 63)           As Variant
#End If
End Type

#If Not HasPtrSafe Then
    Private Function CLngLng(vValue As Variant) As Variant
        Const VT_I8 As Long = &H14
        Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
    End Function
#End If

Private Sub pvInit(Optional ByVal Extended As Boolean)
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    If LNG_POW2(0) = 0 Then
        LNG_POW2(0) = 1
        For lIdx = 1 To UBound(LNG_POW2)
            LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
        Next
        EmptyByteArray = vbNullString
        m_lZero = CLngLng(0)
    End If
    If m_gf1.Item(0) = 0 And Extended Then
        pvGF25519Assign m_gf0, "0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0"
        pvGF25519Assign m_gf1, "1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0"
        pvGF25519Assign m_gfD, "78A3 1359 4DCA 75EB D8AB 4141 0A4D 0070 E898 7779 4079 8CC7 FE73 2B6F 6CEE 5203"
        pvGF25519Assign m_gfD2, "F159 26B2 9B94 EBD6 B156 8283 149A 00E0 D130 EEF3 80F2 198E FCE7 56DF D9DC 2406"
        pvGF25519Assign m_gfX, "D51A 8F25 2D60 C956 A7B2 9525 C760 692C DC5C FDD6 E231 C0A4 53FE CD6E 36D3 2169"
        pvGF25519Assign m_gfY, "6658 6666 6666 6666 6666 6666 6666 6666 6666 6666 6666 6666 6666 6666 6666 6666"
        pvGF25519Assign m_gfI, "A0B0 4A0E 1B27 C4EE E478 AD2F 1806 2F43 D7A7 3DFB 0099 2B4D DF0B 4FC1 2480 2B83"
        lIdx = 0
        For Each vElem In Split("ED D3 F5 5C 1A 63 12 58 D6 9C F7 A2 DE F9 DE 14 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 10")
            m_aL.Item(lIdx) = CLngLng(CStr("&H" & vElem))
            lIdx = lIdx + 1
        Next
    End If
End Sub

Private Sub pvGF25519Sel(uA As GF25519Element, uB As GF25519Element, ByVal bSwap As Boolean)
    Dim lIdx            As Long
#If HasPtrSafe Then
    Dim lTemp           As LongLong
#Else
    Dim lTemp           As Variant
#End If
    
    For lIdx = 0 To LNG_ELEMSZ - 1
        lTemp = (uA.Item(lIdx) Xor uB.Item(lIdx)) And bSwap
        uA.Item(lIdx) = uA.Item(lIdx) Xor lTemp
        uB.Item(lIdx) = uB.Item(lIdx) Xor lTemp
    Next
End Sub

Private Sub pvGF25519Car(uRetVal As GF25519Element)
    Dim lIdx            As Long
    Dim lNext           As Long
#If HasPtrSafe Then
    Dim lCarry          As LongLong
#Else
    Dim lCarry          As Variant
#End If
    
    For lIdx = 0 To LNG_ELEMSZ - 1
        uRetVal.Item(lIdx) = uRetVal.Item(lIdx) + LNG_POW16
        lCarry = (uRetVal.Item(lIdx) And -LNG_POW16) \ LNG_POW16
        uRetVal.Item(lIdx) = uRetVal.Item(lIdx) - lCarry * LNG_POW16
        If lIdx = LNG_ELEMSZ - 1 Then
            lCarry = 38 * (lCarry - 1)
        Else
            lCarry = lCarry - 1
        End If
        lNext = (lIdx + 1) Mod LNG_ELEMSZ
        uRetVal.Item(lNext) = uRetVal.Item(lNext) + lCarry
    Next
End Sub

Private Sub pvGF25519Add(uRetVal As GF25519Element, uA As GF25519Element, uB As GF25519Element)
    Dim lIdx            As Long
    
    For lIdx = 0 To LNG_ELEMSZ - 1
        uRetVal.Item(lIdx) = uA.Item(lIdx) + uB.Item(lIdx)
    Next
End Sub

Private Sub pvGF25519Sub(uRetVal As GF25519Element, uA As GF25519Element, uB As GF25519Element)
    Dim lIdx            As Long
    
    For lIdx = 0 To LNG_ELEMSZ - 1
        uRetVal.Item(lIdx) = uA.Item(lIdx) - uB.Item(lIdx)
    Next
End Sub

Private Sub pvGF25519Mul(uRetVal As GF25519Element, uA As GF25519Element, uB As GF25519Element)
#If HasPtrSafe Then
    Static aTemp(0 To LNG_ELEMSZ * 2 - 1) As LongLong
#Else
    Static aTemp(0 To LNG_ELEMSZ * 2 - 1) As Variant
#End If
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    For lIdx = 0 To UBound(aTemp)
        aTemp(lIdx) = CLng(0)
    Next
    For lIdx = 0 To LNG_ELEMSZ - 1
        For lJdx = 0 To LNG_ELEMSZ - 1
            aTemp(lIdx + lJdx) = aTemp(lIdx + lJdx) + uA.Item(lIdx) * uB.Item(lJdx)
        Next
    Next
    For lIdx = 0 To LNG_ELEMSZ - 1
        If lIdx < LNG_ELEMSZ - 1 Then
            uRetVal.Item(lIdx) = aTemp(lIdx) + 38 * aTemp(lIdx + LNG_ELEMSZ)
        Else
            uRetVal.Item(lIdx) = aTemp(lIdx)
        End If
    Next
    pvGF25519Car uRetVal
    pvGF25519Car uRetVal
End Sub

Private Sub pvGF25519Sqr(uRetVal As GF25519Element, uA As GF25519Element)
    pvGF25519Mul uRetVal, uA, uA
End Sub

Private Sub pvGF25519Inv(uRetVal As GF25519Element, uA As GF25519Element)
    Dim uTemp           As GF25519Element
    Dim lIdx            As Long
    
    uTemp = uA
    For lIdx = 253 To 0 Step -1
        pvGF25519Mul uTemp, uTemp, uTemp
        If lIdx <> 2 And lIdx <> 4 Then
            pvGF25519Mul uTemp, uTemp, uA
        End If
    Next
    uRetVal = uTemp
End Sub

Private Sub pvGF25519Pow2523(uRetVal As GF25519Element, uA As GF25519Element)
    Dim uTemp           As GF25519Element
    Dim lIdx            As Long
    
    uTemp = uA
    For lIdx = 250 To 0 Step -1
        pvGF25519Sqr uTemp, uTemp
        If lIdx <> 1 Then
            pvGF25519Mul uTemp, uTemp, uA
        End If
    Next
    uRetVal = uTemp
End Sub

Private Function pvGF25519Neq(uA As GF25519Element, uB As GF25519Element) As Boolean
    Dim baA()           As Byte
    Dim baB()           As Byte
    Dim lIdx            As Long
    Dim lAccum            As Long
    
    pvGF25519Pack baA, uA
    pvGF25519Pack baB, uB
    For lIdx = 0 To UBound(baA)
        lAccum = lAccum Or (baA(lIdx) Xor baB(lIdx))
    Next
    pvGF25519Neq = lAccum <> 0
End Function

Private Sub pvGF25519Unpack(uRetVal As GF25519Element, baInput() As Byte)
    Dim aTemp(0 To LNG_ELEMSZ - 1) As Integer
    Dim lIdx            As Long

    If UBound(baInput) >= 0 Then
        Debug.Assert (UBound(aTemp) + 1) * 2 >= UBound(baInput) + 1
        Call CopyMemory(aTemp(0), baInput(0), UBound(baInput) + 1)
    End If
    For lIdx = 0 To LNG_ELEMSZ - 1
        If aTemp(lIdx) < 0 Then
            uRetVal.Item(lIdx) = m_lZero + LNG_POW16 + aTemp(lIdx)
        Else
            uRetVal.Item(lIdx) = m_lZero + aTemp(lIdx)
        End If
    Next
End Sub

Private Sub pvGF25519Pack(baRetVal() As Byte, uA As GF25519Element)
    Dim lRetry          As Long
    Dim lIdx            As Long
    Dim uTemp           As GF25519Element
    Dim lFlag           As Long
    
    ReDim baRetVal(0 To LNG_KEYSZ - 1) As Byte
    For lRetry = 0 To 1
        uTemp.Item(0) = uA.Item(0) - &HFFED&
        For lIdx = 1 To LNG_ELEMSZ - 1
            lFlag = -((uTemp.Item(lIdx - 1) And LNG_POW16) <> 0)
            If lIdx = LNG_ELEMSZ - 1 Then
                lFlag = &H7FFF& + lFlag
            Else
                lFlag = &HFFFF& + lFlag
            End If
            uTemp.Item(lIdx) = uA.Item(lIdx) - lFlag
            uTemp.Item(lIdx - 1) = uTemp.Item(lIdx - 1) And &HFFFF&
        Next
        lFlag = -((uTemp.Item(LNG_ELEMSZ - 1) And LNG_POW16) <> 0)
        pvGF25519Sel uA, uTemp, lFlag = 0
    Next
    For lIdx = 0 To LNG_ELEMSZ - 1
        lFlag = CLng(uA.Item(lIdx) And LNG_POW16 - 1)
        Call CopyMemory(baRetVal(2 * lIdx), lFlag, 2)
    Next
End Sub

Private Sub pvGF25519Clamp(baPriv() As Byte)
    baPriv(0) = baPriv(0) And &HF8
    baPriv(31) = baPriv(31) And &H7F Or &H40
End Sub

Private Sub pvGF25519Assign(uRetVal As GF25519Element, sText As String)
    Dim vElem           As Variant
    Dim lIdx            As Long

    For Each vElem In Split(sText)
        uRetVal.Item(lIdx) = CLngLng(CStr("&H" & vElem))
        lIdx = lIdx + 1
    Next
End Sub

Private Sub pvGF25519ScalarMult(baRetVal() As Byte, baPriv() As Byte, baPub() As Byte)
    Dim baKey()         As Byte
    Dim uX              As GF25519Element
    Dim uA              As GF25519Element
    Dim uB              As GF25519Element
    Dim uC              As GF25519Element
    Dim uD              As GF25519Element
    Dim uE              As GF25519Element
    Dim uF              As GF25519Element
    Dim uG              As GF25519Element
    Dim lIdx            As Long
    Dim lFlag           As Long
    Dim lPrev           As Long
    
    baKey = baPriv
    pvGF25519Clamp baKey
    pvGF25519Unpack uA, EmptyByteArray
    pvGF25519Unpack uX, baPub
    uB = uX
    uC = uA
    uD = uA
    uG = uA
    uG.Item(0) = uG.Item(0) + &HDB41&
    uG.Item(1) = uG.Item(1) + 1
    uA.Item(0) = uG.Item(1)         ' a[0] = 1
    uD.Item(0) = uG.Item(1)         ' d[0] = 1
    
    For lIdx = 254 To 0 Step -1
        lPrev = lFlag
        lFlag = (baKey(lIdx \ 8) \ LNG_POW2(lIdx And 7)) And 1
        pvGF25519Sel uA, uB, lFlag Xor lPrev
        pvGF25519Sel uC, uD, lFlag Xor lPrev
        pvGF25519Add uE, uA, uC  ' e = a + c
        pvGF25519Sub uA, uA, uC  ' a = a - c
        pvGF25519Add uC, uB, uD  ' c = b + d
        pvGF25519Sub uB, uB, uD  ' b = b - d
        pvGF25519Mul uD, uE, uE  ' d = e * e
        pvGF25519Mul uF, uA, uA  ' f = a * a
        pvGF25519Mul uA, uC, uA  ' a = c * a
        pvGF25519Mul uC, uB, uE  ' c = b * e
        pvGF25519Add uE, uA, uC  ' e = a + c
        pvGF25519Sub uA, uA, uC  ' a = a - c
        pvGF25519Mul uB, uA, uA  ' b = a * a
        pvGF25519Sub uC, uD, uF  ' c = d - f
        pvGF25519Mul uA, uC, uG  ' a = c * g
        pvGF25519Add uA, uA, uD  ' a = a + d
        pvGF25519Mul uC, uC, uA  ' c = c * a
        pvGF25519Mul uA, uD, uF  ' a = d * f
        pvGF25519Mul uD, uB, uX  ' d = b * x
        pvGF25519Mul uB, uE, uE  ' b = e * e
    Next
    pvGF25519Inv uC, uC
    pvGF25519Mul uX, uA, uC
    pvGF25519Pack baRetVal, uX
End Sub

Private Sub pvGF25519ScalarBase(baRetVal() As Byte, baPriv() As Byte)
    Dim baBase(0 To LNG_KEYSZ - 1) As Byte
    
    baBase(0) = 9
    pvGF25519ScalarMult baRetVal, baPriv, baBase
End Sub

Public Sub CryptoX25519PrivateKey(baRetVal() As Byte, Optional Seed As Variant)
    If Not IsMissing(Seed) Then
        baRetVal = Seed
        ReDim Preserve baRetVal(0 To LNG_KEYSZ - 1) As Byte
    Else
        ReDim baRetVal(0 To LNG_KEYSZ - 1) As Byte
        Call RtlGenRandom(baRetVal(0), UBound(baRetVal) + 1)
    End If
    pvGF25519Clamp baRetVal
End Sub

Public Sub CryptoX25519PublicKey(baRetVal() As Byte, baPriv() As Byte)
    pvInit
    pvGF25519ScalarBase baRetVal, baPriv
End Sub

Public Sub CryptoX25519SharedSecret(baRetVal() As Byte, baPriv() As Byte, baPub() As Byte)
    pvInit
    pvGF25519ScalarMult baRetVal, baPriv, baPub
End Sub

'= XyztPoint =============================================================

Private Sub pvEdwardsAdd(uP As XyztPoint, uQ As XyztPoint)
    Dim gfA             As GF25519Element
    Dim gfB             As GF25519Element
    Dim gfC             As GF25519Element
    Dim gfD             As GF25519Element
    Dim gfE             As GF25519Element
    Dim gfF             As GF25519Element
    Dim gfG             As GF25519Element
    Dim gfH             As GF25519Element
    Dim gfT             As GF25519Element
    
    pvGF25519Sub gfA, uP.gfY, uP.gfX
    pvGF25519Sub gfT, uQ.gfY, uQ.gfX
    pvGF25519Mul gfA, gfA, gfT
    pvGF25519Add gfB, uP.gfX, uP.gfY
    pvGF25519Add gfT, uQ.gfX, uQ.gfY
    pvGF25519Mul gfB, gfB, gfT
    pvGF25519Mul gfC, uP.gfT, uQ.gfT
    pvGF25519Mul gfC, gfC, m_gfD2
    pvGF25519Mul gfD, uP.gfZ, uQ.gfZ
    pvGF25519Add gfD, gfD, gfD
    pvGF25519Sub gfE, gfB, gfA
    pvGF25519Sub gfF, gfD, gfC
    pvGF25519Add gfG, gfD, gfC
    pvGF25519Add gfH, gfB, gfA
    pvGF25519Mul uP.gfX, gfE, gfF
    pvGF25519Mul uP.gfY, gfH, gfG
    pvGF25519Mul uP.gfZ, gfG, gfF
    pvGF25519Mul uP.gfT, gfE, gfH
End Sub

Private Sub pvEdwardsCSwap(uP As XyztPoint, uQ As XyztPoint, ByVal bSwap As Boolean)
    pvGF25519Sel uP.gfX, uQ.gfX, bSwap
    pvGF25519Sel uP.gfY, uQ.gfY, bSwap
    pvGF25519Sel uP.gfZ, uQ.gfZ, bSwap
    pvGF25519Sel uP.gfT, uQ.gfT, bSwap
End Sub

Private Sub pvEdwardsPack(baRetVal() As Byte, ByVal lOutPos As Long, uP As XyztPoint)
    Dim gfTx            As GF25519Element
    Dim gfTy            As GF25519Element
    Dim gfZi            As GF25519Element
    Dim baTemp()        As Byte
    
    pvGF25519Inv gfZi, uP.gfZ
    pvGF25519Mul gfTx, uP.gfX, gfZi
    pvGF25519Mul gfTy, uP.gfY, gfZi
    pvGF25519Pack baTemp, gfTy
    Debug.Assert UBound(baRetVal) + 1 >= lOutPos + LNG_KEYSZ
    Call CopyMemory(baRetVal(lOutPos), baTemp(0), LNG_KEYSZ)
    pvGF25519Pack baTemp, gfTx
    lOutPos = lOutPos + LNG_KEYSZ - 1
    baRetVal(lOutPos) = baRetVal(lOutPos) Xor ((baTemp(0) And 1) * &H80)
End Sub

Private Sub pvEdwardsScalarMult(uP As XyztPoint, uQ As XyztPoint, baKey() As Byte, Optional ByVal lPos As Long)
    Dim lIdx            As Long
    Dim lFlag           As Long
    
    pvInit Extended:=True
    uP.gfX = m_gf0
    uP.gfY = m_gf1
    uP.gfZ = m_gf1
    uP.gfT = m_gf0
    For lIdx = 255 To 0 Step -1
        lFlag = (baKey(lPos + lIdx \ 8) \ LNG_POW2(lIdx And 7)) And 1
        pvEdwardsCSwap uP, uQ, lFlag
        pvEdwardsAdd uQ, uP
        pvEdwardsAdd uP, uP
        pvEdwardsCSwap uP, uQ, lFlag
    Next
End Sub

Private Sub pvEdwardsScalarBase(uP As XyztPoint, baKey() As Byte, Optional ByVal lPos As Long)
    Dim uQ              As XyztPoint
    
    uQ.gfX = m_gfX
    uQ.gfY = m_gfY
    uQ.gfZ = m_gf1
    pvGF25519Mul uQ.gfT, m_gfX, m_gfY
    pvEdwardsScalarMult uP, uQ, baKey, lPos
End Sub

Private Sub pvEdwardsModL(baRetVal() As Byte, ByVal lOutPos As Long, aX As ArrayLong64)
#If HasPtrSafe Then
    Dim lCarry          As LongLong
#Else
    Dim lCarry          As Variant
#End If
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    For lIdx = 63 To 32 Step -1
        lCarry = m_lZero
        For lJdx = lIdx - 32 To lIdx - 13
            aX.Item(lJdx) = aX.Item(lJdx) + lCarry - 16 * aX.Item(lIdx) * m_aL.Item(lJdx - (lIdx - 32))
            lCarry = (aX.Item(lJdx) + 128 And -&H100) \ &H100
            aX.Item(lJdx) = aX.Item(lJdx) - lCarry * &H100
        Next
        aX.Item(lJdx) = aX.Item(lJdx) + lCarry
        aX.Item(lIdx) = 0
    Next
    lCarry = 0
    For lJdx = 0 To 31
        aX.Item(lJdx) = aX.Item(lJdx) + lCarry - ((aX.Item(31) And -&H10) \ &H10) * m_aL.Item(lJdx)
        lCarry = (aX.Item(lJdx) And -&H100) \ &H100
        aX.Item(lJdx) = aX.Item(lJdx) And &HFF
    Next
    For lJdx = 0 To 31
        aX.Item(lJdx) = aX.Item(lJdx) - lCarry * m_aL.Item(lJdx)
    Next
    For lIdx = 0 To 31
        aX.Item(lIdx + 1) = aX.Item(lIdx + 1) + ((aX.Item(lIdx) And -&H100) \ &H100)
        baRetVal(lOutPos + lIdx) = CByte(aX.Item(lIdx) And &HFF)
    Next
End Sub

Private Sub pvEdwardsReduce(baRetVal() As Byte)
    Dim aX              As ArrayLong64
    Dim lIdx            As Long
    
    For lIdx = 0 To 63
        aX.Item(lIdx) = m_lZero + baRetVal(lIdx)
        baRetVal(lIdx) = 0
    Next
    pvEdwardsModL baRetVal, 0, aX
End Sub

Private Function pvEdwardsUnpackNeg(uR As XyztPoint, baKey() As Byte) As Boolean
    Dim gfT             As GF25519Element
    Dim gfChk           As GF25519Element
    Dim gfNum           As GF25519Element
    Dim gfDen           As GF25519Element
    Dim gfDen2          As GF25519Element
    Dim gfDen4          As GF25519Element
    Dim gfDen6          As GF25519Element
    Dim baTemp()        As Byte
    
    uR.gfZ = m_gf1
    pvGF25519Unpack uR.gfY, baKey
    pvGF25519Sqr gfNum, uR.gfY
    pvGF25519Mul gfDen, gfNum, m_gfD
    pvGF25519Sub gfNum, gfNum, m_gf1
    pvGF25519Add gfDen, gfDen, m_gf1
    pvGF25519Sqr gfDen2, gfDen
    pvGF25519Sqr gfDen4, gfDen2
    pvGF25519Mul gfDen6, gfDen4, gfDen2
    pvGF25519Mul gfT, gfDen6, gfNum
    pvGF25519Mul gfT, gfT, gfDen
    pvGF25519Pow2523 gfT, gfT
    pvGF25519Mul gfT, gfT, gfNum
    pvGF25519Mul gfT, gfT, gfDen
    pvGF25519Mul gfT, gfT, gfDen
    pvGF25519Mul uR.gfX, gfT, gfDen
    pvGF25519Sqr gfChk, uR.gfX
    pvGF25519Mul gfChk, gfChk, gfDen
    If pvGF25519Neq(gfChk, gfNum) Then
        pvGF25519Mul uR.gfX, uR.gfX, m_gfI
    End If
    pvGF25519Sqr gfChk, uR.gfX
    pvGF25519Mul gfChk, gfChk, gfDen
    If pvGF25519Neq(gfChk, gfNum) Then
        GoTo QH
    End If
    pvGF25519Pack baTemp, uR.gfX
    If (baTemp(0) And 1) = (baKey(31) \ &H80) Then
        pvGF25519Sub uR.gfX, m_gf0, uR.gfX '-- X = -X
    End If
    pvGF25519Mul uR.gfT, uR.gfX, uR.gfY
    '--- success
    pvEdwardsUnpackNeg = True
QH:
End Function

Private Function pvEdwardsHash(baOutput() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    #If HasSha512 Then
        baOutput = CryptoSha512ByteArray(512, baInput, Pos, Size)
        Debug.Assert UBound(baOutput) + 1 >= LNG_HASHSZ
    #Else
        Err.Raise vbObjectError, , "SHA-512 not compiled (use CRYPT_HAS_SHA512 = 1)"
    #End If
End Function

Public Sub pvEdwardsPublicKey(baRetVal() As Byte, ByVal lOutPos As Long, baPriv() As Byte)
    Dim baD()           As Byte
    Dim uP              As XyztPoint
    
    pvEdwardsHash baD, baPriv
    pvGF25519Clamp baD
    pvEdwardsScalarBase uP, baD
    pvEdwardsPack baRetVal, lOutPos, uP
End Sub

Public Sub CryptoEd25519PrivateKey(baRetVal() As Byte, Optional Seed As Variant)
    If Not IsMissing(Seed) Then
        baRetVal = Seed
        ReDim Preserve baRetVal(0 To LNG_KEYSZ - 1) As Byte
    Else
        ReDim baRetVal(0 To LNG_KEYSZ - 1) As Byte
        Call RtlGenRandom(baRetVal(0), UBound(baRetVal) + 1)
    End If
End Sub

Public Sub CryptoEd25519PublicKey(baRetVal() As Byte, baPriv() As Byte)
    Debug.Assert UBound(baPriv) + 1 >= LNG_KEYSZ
    pvInit Extended:=True
    ReDim baRetVal(0 To LNG_KEYSZ - 1) As Byte
    pvEdwardsPublicKey baRetVal, 0, baPriv
End Sub

Public Sub CryptoEd25519Sign(baRetVal() As Byte, baPriv() As Byte, baMsg() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim baDelta()       As Byte
    Dim baHash()        As Byte
    Dim baR()           As Byte
    Dim uP              As XyztPoint
    Dim aX              As ArrayLong64
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    Debug.Assert UBound(baPriv) + 1 >= LNG_KEYSZ
    pvInit Extended:=True
    pvEdwardsHash baDelta, baPriv
    pvGF25519Clamp baDelta
    If Size < 0 Then
        Size = UBound(baMsg) + 1 - Pos
    End If
    ReDim baRetVal(0 To LNG_HASHSZ + Size - 1) As Byte
    Call CopyMemory(baRetVal(LNG_HALFHASHSZ), baDelta(LNG_HALFHASHSZ), LNG_HALFHASHSZ)
    If Size > 0 Then
        Call CopyMemory(baRetVal(LNG_HASHSZ), baMsg(Pos), Size)
    End If
    pvEdwardsHash baR, baRetVal, Pos:=LNG_HALFHASHSZ
    pvEdwardsReduce baR
    pvEdwardsScalarBase uP, baR
    pvEdwardsPack baRetVal, 0, uP
    pvEdwardsPublicKey baRetVal, LNG_HALFHASHSZ, baPriv
    pvEdwardsHash baHash, baRetVal
    pvEdwardsReduce baHash
    For lIdx = 0 To LNG_HALFHASHSZ - 1
        aX.Item(lIdx) = baR(lIdx)
    Next
    For lIdx = 0 To LNG_HALFHASHSZ - 1
        For lJdx = 0 To LNG_HALFHASHSZ - 1
            aX.Item(lIdx + lJdx) = aX.Item(lIdx + lJdx) + (m_lZero + baHash(lIdx)) * baDelta(lJdx)
        Next
    Next
    pvEdwardsModL baRetVal, LNG_HALFHASHSZ, aX
End Sub

Public Function CryptoEd25519Open(baRetVal() As Byte, baPub() As Byte, baSigMsg() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Dim uP              As XyztPoint
    Dim uQ              As XyztPoint
    Dim baHash()        As Byte
    Dim baTemp(0 To LNG_KEYSZ - 1) As Byte
    Dim lIdx            As Long
    
    Debug.Assert UBound(baPub) + 1 >= LNG_KEYSZ
    pvInit Extended:=True
    If Size < 0 Then
        Size = UBound(baSigMsg) + 1 - Pos
    End If
    If Size < LNG_HASHSZ Then
        GoTo QH
    End If
    If Not pvEdwardsUnpackNeg(uQ, baPub) Then
        GoTo QH
    End If
    ReDim baRetVal(0 To Size - 1) As Byte
    Debug.Assert UBound(baSigMsg) + 1 >= Pos + Size
    Call CopyMemory(baRetVal(0), baSigMsg(Pos), Size)
    Call CopyMemory(baRetVal(LNG_HALFHASHSZ), baPub(0), LNG_HALFHASHSZ)
    pvEdwardsHash baHash, baRetVal
    pvEdwardsReduce baHash
    pvEdwardsScalarMult uP, uQ, baHash
    pvEdwardsScalarBase uQ, baSigMsg, LNG_HALFHASHSZ
    pvEdwardsAdd uP, uQ
    pvEdwardsPack baTemp, 0, uP
    For lIdx = 0 To LNG_HALFHASHSZ - 1
        If baTemp(lIdx) <> baSigMsg(lIdx) Then
            GoTo QH
        End If
    Next
    If UBound(baSigMsg) + 1 > LNG_HASHSZ Then
        ReDim baRetVal(0 To UBound(baSigMsg) - LNG_HASHSZ) As Byte
        Call CopyMemory(baRetVal(0), baSigMsg(LNG_HASHSZ), UBound(baRetVal) + 1)
    Else
        baRetVal = vbNullString
    End If
    '--- success
    CryptoEd25519Open = True
QH:
End Function

Public Sub CryptoEd25519SignDetached(baRetVal() As Byte, baPriv() As Byte, baMsg() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    CryptoEd25519Sign baRetVal, baPriv, baMsg, Pos, Size
    ReDim Preserve baRetVal(0 To LNG_HASHSZ - 1) As Byte
End Sub

Public Function CryptoEd25519VerifyDetached(baSig() As Byte, baPub() As Byte, baMsg() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Dim baSigMsg()          As Byte
    Dim baTemp()            As Byte
    
    If UBound(baSig) + 1 < LNG_HASHSZ Then
        GoTo QH
    End If
    If Size < 0 Then
        Size = UBound(baMsg) + 1 - Pos
    End If
    ReDim baSigMsg(0 To LNG_HASHSZ + UBound(baMsg)) As Byte
    Call CopyMemory(baSigMsg(0), baSig(0), LNG_HASHSZ)
    If UBound(baMsg) >= 0 Then
        Call CopyMemory(baSigMsg(LNG_HASHSZ), baMsg(0), UBound(baMsg) + 1)
    End If
    CryptoEd25519VerifyDetached = CryptoEd25519Open(baTemp, baPub, baSigMsg)
QH:
End Function
