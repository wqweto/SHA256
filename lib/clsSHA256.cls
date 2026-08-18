VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSHA256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic SHA256 Implementation
' Phil Fresle and David Midkiff (mdj2023@hotmail.com)
'
' Standard SHA256 implementation optimised for the Visual Basic environment.
' Conforms to all standards and can be used in digital signature or password
' protection related schemes.

Private m_lOnBits(30)   As Long
Private m_l2Power(30)   As Long
Private K(63)           As Long

Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE
Private Const MODULUS_BITS As Long = 512
Private Const CONGRUENT_BITS As Long = 448


Private Sub Class_Initialize()
    m_lOnBits(0) = 1            ' 00000000000000000000000000000001
    m_lOnBits(1) = 3            ' 00000000000000000000000000000011
    m_lOnBits(2) = 7            ' 00000000000000000000000000000111
    m_lOnBits(3) = 15           ' 00000000000000000000000000001111
    m_lOnBits(4) = 31           ' 00000000000000000000000000011111
    m_lOnBits(5) = 63           ' 00000000000000000000000000111111
    m_lOnBits(6) = 127          ' 00000000000000000000000001111111
    m_lOnBits(7) = 255          ' 00000000000000000000000011111111
    m_lOnBits(8) = 511          ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    m_l2Power(0) = 1            ' 00000000000000000000000000000001
    m_l2Power(1) = 2            ' 00000000000000000000000000000010
    m_l2Power(2) = 4            ' 00000000000000000000000000000100
    m_l2Power(3) = 8            ' 00000000000000000000000000001000
    m_l2Power(4) = 16           ' 00000000000000000000000000010000
    m_l2Power(5) = 32           ' 00000000000000000000000000100000
    m_l2Power(6) = 64           ' 00000000000000000000000001000000
    m_l2Power(7) = 128          ' 00000000000000000000000010000000
    m_l2Power(8) = 256          ' 00000000000000000000000100000000
    m_l2Power(9) = 512          ' 00000000000000000000001000000000
    m_l2Power(10) = 1024        ' 00000000000000000000010000000000
    m_l2Power(11) = 2048        ' 00000000000000000000100000000000
    m_l2Power(12) = 4096        ' 00000000000000000001000000000000
    m_l2Power(13) = 8192        ' 00000000000000000010000000000000
    m_l2Power(14) = 16384       ' 00000000000000000100000000000000
    m_l2Power(15) = 32768       ' 00000000000000001000000000000000
    m_l2Power(16) = 65536       ' 00000000000000010000000000000000
    m_l2Power(17) = 131072      ' 00000000000000100000000000000000
    m_l2Power(18) = 262144      ' 00000000000001000000000000000000
    m_l2Power(19) = 524288      ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
    
    K(0) = &H428A2F98
    K(1) = &H71374491
    K(2) = &HB5C0FBCF
    K(3) = &HE9B5DBA5
    K(4) = &H3956C25B
    K(5) = &H59F111F1
    K(6) = &H923F82A4
    K(7) = &HAB1C5ED5
    K(8) = &HD807AA98
    K(9) = &H12835B01
    K(10) = &H243185BE
    K(11) = &H550C7DC3
    K(12) = &H72BE5D74
    K(13) = &H80DEB1FE
    K(14) = &H9BDC06A7
    K(15) = &HC19BF174
    K(16) = &HE49B69C1
    K(17) = &HEFBE4786
    K(18) = &HFC19DC6
    K(19) = &H240CA1CC
    K(20) = &H2DE92C6F
    K(21) = &H4A7484AA
    K(22) = &H5CB0A9DC
    K(23) = &H76F988DA
    K(24) = &H983E5152
    K(25) = &HA831C66D
    K(26) = &HB00327C8
    K(27) = &HBF597FC7
    K(28) = &HC6E00BF3
    K(29) = &HD5A79147
    K(30) = &H6CA6351
    K(31) = &H14292967
    K(32) = &H27B70A85
    K(33) = &H2E1B2138
    K(34) = &H4D2C6DFC
    K(35) = &H53380D13
    K(36) = &H650A7354
    K(37) = &H766A0ABB
    K(38) = &H81C2C92E
    K(39) = &H92722C85
    K(40) = &HA2BFE8A1
    K(41) = &HA81A664B
    K(42) = &HC24B8B70
    K(43) = &HC76C51A3
    K(44) = &HD192E819
    K(45) = &HD6990624
    K(46) = &HF40E3585
    K(47) = &H106AA070
    K(48) = &H19A4C116
    K(49) = &H1E376C08
    K(50) = &H2748774C
    K(51) = &H34B0BCB5
    K(52) = &H391C0CB3
    K(53) = &H4ED8AA4A
    K(54) = &H5B9CCA4F
    K(55) = &H682E6FF3
    K(56) = &H748F82EE
    K(57) = &H78A5636F
    K(58) = &H84C87814
    K(59) = &H8CC70208
    K(60) = &H90BEFFFA
    K(61) = &HA4506CEB
    K(62) = &HBEF9A3F7
    K(63) = &HC67178F2
End Sub

Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then LShift = &H80000000 Else LShift = 0
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    If (lValue And m_l2Power(31 - iShiftBits)) Then LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000 Else LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
End Function

Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then RShift = 1 Else RShift = 0
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    If (lValue And &H80000000) Then RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End Function

Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4 As Long, lY4 As Long, lX8 As Long, lY8 As Long, lResult As Long
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8 Else lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
    AddUnsigned = lResult
End Function

Private Function Ch(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    Ch = ((X And Y) Xor ((Not X) And z))
End Function

Private Function Maj(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    Maj = ((X And Y) Xor (X And z) Xor (Y And z))
End Function

Private Function S(ByVal X As Long, ByVal n As Long) As Long
    S = (RShift(X, (n And m_lOnBits(4))) Or LShift(X, (32 - (n And m_lOnBits(4)))))
End Function

Private Function R(ByVal X As Long, ByVal n As Long) As Long
    R = RShift(X, CInt(n And m_lOnBits(4)))
End Function

Private Function Sigma0(ByVal X As Long) As Long
    Sigma0 = (S(X, 2) Xor S(X, 13) Xor S(X, 22))
End Function

Private Function Sigma1(ByVal X As Long) As Long
    Sigma1 = (S(X, 6) Xor S(X, 11) Xor S(X, 25))
End Function

Private Function Gamma0(ByVal X As Long) As Long
    Gamma0 = (S(X, 7) Xor S(X, 18) Xor R(X, 3))
End Function

Private Function Gamma1(ByVal X As Long) As Long
    Gamma1 = (S(X, 17) Xor S(X, 19) Xor R(X, 10))
End Function

Private Function ConvertToWordArray(sMessage As String) As Long()
    Dim lMessageLength As Long
    Dim lNumberOfWords As Long
    Dim lWordArray() As Long
    Dim lBytePosition As Long
    Dim lByteCount As Long
    Dim lWordCount As Long
    Dim lByte As Long
    
    lMessageLength = Len(sMessage)
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        lByte = AscB(Mid$(sMessage, lByteCount + 1, 1))
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
    ConvertToWordArray = lWordArray
End Function

Public Function SHA256(sMessage As String) As String
    Dim HASH(7) As Long, M() As Long, W(63) As Long, A As Long, B As Long
    Dim C As Long, D As Long, E As Long, F As Long, G As Long, H As Long
    Dim I As Long, J As Long, T1 As Long, T2 As Long
    
    HASH(0) = &H6A09E667
    HASH(1) = &HBB67AE85
    HASH(2) = &H3C6EF372
    HASH(3) = &HA54FF53A
    HASH(4) = &H510E527F
    HASH(5) = &H9B05688C
    HASH(6) = &H1F83D9AB
    HASH(7) = &H5BE0CD19
    
    M = ConvertToWordArray(sMessage)
    
    For I = 0 To UBound(M) Step 16
        A = HASH(0)
        B = HASH(1)
        C = HASH(2)
        D = HASH(3)
        E = HASH(4)
        F = HASH(5)
        G = HASH(6)
        H = HASH(7)
        
        For J = 0 To 63
            If J < 16 Then W(J) = M(J + I) Else W(J) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(J - 2)), W(J - 7)), Gamma0(W(J - 15))), W(J - 16))
            T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(H, Sigma1(E)), Ch(E, F, G)), K(J)), W(J))
            T2 = AddUnsigned(Sigma0(A), Maj(A, B, C))
            H = G
            G = F
            F = E
            E = AddUnsigned(D, T1)
            D = C
            C = B
            B = A
            A = AddUnsigned(T1, T2)
        Next
        
        HASH(0) = AddUnsigned(A, HASH(0))
        HASH(1) = AddUnsigned(B, HASH(1))
        HASH(2) = AddUnsigned(C, HASH(2))
        HASH(3) = AddUnsigned(D, HASH(3))
        HASH(4) = AddUnsigned(E, HASH(4))
        HASH(5) = AddUnsigned(F, HASH(5))
        HASH(6) = AddUnsigned(G, HASH(6))
        HASH(7) = AddUnsigned(H, HASH(7))
    Next
    SHA256 = LCase$(Right$("00000000" & Hex(HASH(0)), 8) & Right$("00000000" & Hex(HASH(1)), 8) & Right$("00000000" & Hex(HASH(2)), 8) & Right$("00000000" & Hex(HASH(3)), 8) & Right$("00000000" & Hex(HASH(4)), 8) & Right$("00000000" & Hex(HASH(5)), 8) & Right$("00000000" & Hex(HASH(6)), 8) & Right$("00000000" & Hex(HASH(7)), 8))
End Function

