VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSHA256"
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

Private m_l2Power(0 To 31)   As Long

Private Function LShift(ByVal X As Long, ByVal n As Long) As Long
    LShift = (X And (m_l2Power(31 - n) - 1)) * m_l2Power(n) Or -((X And m_l2Power(31 - n)) <> 0) * &H80000000
End Function

Private Function RShift(ByVal X As Long, ByVal n As Long) As Long
    RShift = (X And &H7FFFFFFF) \ m_l2Power(n) Or -(X < 0) * m_l2Power(31 - n)
End Function

Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
    On Error GoTo EH
    AddUnsigned = lX + lY
    Exit Function
EH:
    AddUnsigned = (lX Xor &H80000000) + lY Xor &H80000000
End Function

Private Function Ch(ByVal X As Long, ByVal Y As Long, ByVal Z As Long) As Long
    Ch = ((X And Y) Xor ((Not X) And Z))
End Function

Private Function Maj(ByVal X As Long, ByVal Y As Long, ByVal Z As Long) As Long
    Maj = ((X And Y) Xor (X And Z) Xor (Y And Z))
End Function

Private Function S(ByVal X As Long, ByVal n As Long) As Long
    S = ((X And &H7FFFFFFF) \ m_l2Power(n) - (X < 0) * m_l2Power(31 - n)) ' RShift(X, n)
    S = S Or ((X And (m_l2Power(n - 1) - 1)) * m_l2Power(32 - n) Or -((X And m_l2Power(n - 1)) <> 0) * &H80000000) ' LShift(X, 32 - n)
End Function

Private Function Sigma0(ByVal X As Long) As Long
    Sigma0 = (S(X, 2) Xor S(X, 13) Xor S(X, 22))
End Function

Private Function Sigma1(ByVal X As Long) As Long
    Sigma1 = (S(X, 6) Xor S(X, 11) Xor S(X, 25))
End Function

Private Function Gamma0(ByVal X As Long) As Long
    Gamma0 = (S(X, 7) Xor S(X, 18) Xor RShift(X, 3))
End Function

Private Function Gamma1(ByVal X As Long) As Long
    Gamma1 = (S(X, 17) Xor S(X, 19) Xor RShift(X, 10))
End Function

Private Function ToBigEndian(baMessage() As Byte) As Long()
    Dim aRetVal()       As Long
    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long
    
    lMessageLength = UBound(baMessage) + 1
    lNumberOfWords = ((lMessageLength + 8) \ 64 + 1) * 16
    ReDim aRetVal(lNumberOfWords - 1)
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ 4
        lBytePosition = (3 - (lByteCount Mod 4)) * 8
        lByte = baMessage(lByteCount)
        aRetVal(lWordCount) = aRetVal(lWordCount) Or (lByte * m_l2Power(lBytePosition))  'LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop
    lWordCount = lByteCount \ 4
    lBytePosition = (3 - (lByteCount Mod 4)) * 8
    aRetVal(lWordCount) = aRetVal(lWordCount) Or LShift(&H80, lBytePosition)
    aRetVal(lNumberOfWords - 1) = lMessageLength * 8            ' LShift(lMessageLength, 3)
    aRetVal(lNumberOfWords - 2) = lMessageLength / &H20000000   ' RShift(lMessageLength, 29)
    ToBigEndian = aRetVal
End Function

Private Function FromBigEndian(aHash() As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    Dim lWord           As Long
    
    ReDim baRetVal(0 To UBound(aHash) * 4 + 3) As Byte
    For lIdx = 0 To UBound(aHash)
        lWord = aHash(lIdx)
        baRetVal(4 * lIdx + 0) = RShift(lWord, 24) And &HFF&
        baRetVal(4 * lIdx + 1) = (lWord And &HFF0000) \ &H10000 And &HFF&
        baRetVal(4 * lIdx + 2) = (lWord And &HFF00) \ &H100& And &HFF&
        baRetVal(4 * lIdx + 3) = lWord And &HFF&
    Next
    FromBigEndian = baRetVal
End Function

Public Function SHA256(baMessage() As Byte) As Byte()
    Static K(0 To 63)   As Long
    Dim HASH(0 To 7)    As Long
    Dim M()             As Long
    Dim W(63)           As Long
    Dim A               As Long
    Dim B               As Long
    Dim C               As Long
    Dim D               As Long
    Dim E               As Long
    Dim F               As Long
    Dim G               As Long
    Dim H               As Long
    Dim I               As Long
    Dim J               As Long
    Dim T1              As Long
    Dim T2              As Long
    
    If m_l2Power(0) = 0 Then
        For I = 0 To 30
            m_l2Power(I) = 2& ^ I
        Next
        m_l2Power(31) = &H80000000
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
    End If
    
    HASH(0) = &H6A09E667
    HASH(1) = &HBB67AE85
    HASH(2) = &H3C6EF372
    HASH(3) = &HA54FF53A
    HASH(4) = &H510E527F
    HASH(5) = &H9B05688C
    HASH(6) = &H1F83D9AB
    HASH(7) = &H5BE0CD19
    
    M = ToBigEndian(baMessage)
    
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
            If J < 16 Then
                W(J) = M(J + I)
            Else
                W(J) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(J - 2)), W(J - 7)), Gamma0(W(J - 15))), W(J - 16))
            End If
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
    SHA256 = FromBigEndian(HASH)
End Function

