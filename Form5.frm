VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   9636
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6012
   LinkTopic       =   "Form5"
   ScaleHeight     =   9636
   ScaleWidth      =   6012
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command18 
      Caption         =   "AES256-GCM"
      Height          =   600
      Left            =   2940
      TabIndex        =   19
      Top             =   8652
      Width           =   2028
   End
   Begin VB.CommandButton Command17 
      Caption         =   "AES128-GCM"
      Height          =   600
      Left            =   588
      TabIndex        =   18
      Top             =   8652
      Width           =   2028
   End
   Begin VB.CommandButton Command16 
      Caption         =   "AES256-CTR"
      Height          =   600
      Left            =   2940
      TabIndex        =   17
      Top             =   7896
      Width           =   2028
   End
   Begin VB.CommandButton Command15 
      Caption         =   "AES256-CBC"
      Height          =   600
      Left            =   2940
      TabIndex        =   16
      Top             =   7140
      Width           =   2028
   End
   Begin VB.CommandButton Command14 
      Caption         =   "AES128-CTR"
      Height          =   600
      Left            =   588
      TabIndex        =   15
      Top             =   7896
      Width           =   2028
   End
   Begin VB.CommandButton Command13 
      Caption         =   "AES128-CBC"
      Height          =   600
      Left            =   588
      TabIndex        =   14
      Top             =   7140
      Width           =   2028
   End
   Begin VB.CommandButton Command12 
      Caption         =   "ChaCha20-Poly1305"
      Height          =   600
      Left            =   2940
      TabIndex        =   12
      Top             =   6468
      Width           =   2028
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Ascon-AEAD"
      Height          =   600
      Left            =   2940
      TabIndex        =   11
      Top             =   5712
      Width           =   2028
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Ascon-Hash"
      Height          =   600
      Left            =   588
      TabIndex        =   10
      Top             =   5712
      Width           =   2028
   End
   Begin VB.CommandButton Command9 
      Caption         =   "HalfSiphash13"
      Height          =   600
      Left            =   2940
      TabIndex        =   9
      Top             =   4956
      Width           =   2028
   End
   Begin VB.CommandButton Command8 
      Caption         =   "HalfSiphash24"
      Height          =   600
      Left            =   588
      TabIndex        =   8
      Top             =   4956
      Width           =   2028
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Siphash13"
      Height          =   600
      Left            =   2940
      TabIndex        =   7
      Top             =   4200
      Width           =   2028
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Siphash24"
      Height          =   600
      Left            =   588
      TabIndex        =   6
      Top             =   4200
      Width           =   2028
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Blake2s"
      Height          =   600
      Left            =   588
      TabIndex        =   5
      Top             =   3444
      Width           =   2028
   End
   Begin VB.CommandButton Command3b 
      Caption         =   "SHA-512"
      Height          =   600
      Left            =   2856
      TabIndex        =   4
      Top             =   1932
      Width           =   2028
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SHA-3"
      Height          =   600
      Left            =   588
      TabIndex        =   3
      Top             =   2688
      Width           =   2028
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SHA-2"
      Height          =   600
      Left            =   588
      TabIndex        =   2
      Top             =   1932
      Width           =   2028
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SHA-1"
      Height          =   600
      Left            =   588
      TabIndex        =   1
      Top             =   1176
      Width           =   2028
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MD5"
      Height          =   600
      Left            =   588
      TabIndex        =   0
      Top             =   420
      Width           =   2028
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   192
      Left            =   0
      TabIndex        =   13
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   36
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_baContents() As Byte

Private Sub Form_Load()
    Dim bInIDE          As Boolean: Debug.Assert pvSetTrue(bInIDE)
    
    If bInIDE Then
        m_baContents = ReadBinaryFile("D:\TEMP\curl-7.86.0_2-win64-mingw.zip")
        'm_baContents = ReadBinaryFile("D:\TEMP\Panels Tutorial.zip")
    Else
        m_baContents = ReadBinaryFile("D:\TEMP\VirtualBox-6.1.34-150636-Win.exe")
    End If
    Caption = UBound(m_baContents) + 1 & " bytes ready"
End Sub

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function ReadBinaryFile(sFile As String) As Byte()
    Dim baBuffer()      As Byte
    Dim nFile           As Integer

    On Error GoTo EH
    baBuffer = vbNullString
    nFile = FreeFile
    Open sFile For Binary Access Read Shared As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
    End If
    Close nFile
    ReadBinaryFile = baBuffer
EH:
End Function

Private Sub Command1_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command1.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoMd5ByteArray(m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command1.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command2_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command2.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoSha1ByteArray(m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command2.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command3_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command3.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoSha2ByteArray(256, m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command3.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command3b_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command3b.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoSha2ByteArray(512, m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command3b.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command4_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command4.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoSha3ByteArray(256, m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command4.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command5_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command5.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoBlake2sByteArray(256, m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command5.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub


Private Sub Command6_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    Command6.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoSiphash24ByteArray(FromHex(""), m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command6.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub


Private Sub Command7_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command7.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoSiphash13ByteArray(FromHex(""), m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command7.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command8_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command8.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoHalfSiphash24ByteArray(FromHex(""), m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command8.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command9_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command9.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoHalfSiphash13ByteArray(FromHex(""), m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command9.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command10_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command10.Caption = "Processing"
    dblTimer = TimerEx
    For lIdx = 1 To ITER
        baOutput = CryptoAsconHashByteArray(m_baContents)
    Next
    Label1.Caption = ToHex(baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command10.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command11_Click()
    Const AVARIANT As String = "Ascon-128a"
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baKey() As Byte
    Dim baTag() As Byte
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    Dim bResult As Boolean
    
    Command11.Caption = "Processing"
    dblTimer = TimerEx
    baKey = FromHex("000102030405060708090a0b0c0d0e0f")
    For lIdx = 1 To ITER
        baOutput = m_baContents
        CryptoAsconEncrypt baKey, baTag, baOutput, AsconVariant:=AVARIANT
    Next
    Command11.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    dblTimer = TimerEx
    bResult = CryptoAsconDecrypt(baKey, baTag, baOutput, AsconVariant:=AVARIANT)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec - " & bResult
    Command11.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command12_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baKey() As Byte
    Dim baTag() As Byte
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    Dim bResult As Boolean
    
    Command12.Caption = "Processing"
    dblTimer = TimerEx
    baKey = FromHex("000102030405060708090a0b0c0d0e0f")
    For lIdx = 1 To ITER
        baOutput = m_baContents
        CryptoChaCha20Poly1305Encrypt baKey, baTag, baOutput
    Next
    Command12.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    dblTimer = TimerEx
    bResult = CryptoChaCha20Poly1305Decrypt(baKey, baTag, baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec - " & bResult
    Command12.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command13_Click()
    Const ITER      As Long = 1
    Dim lIdx        As Long
    Dim baKey()     As Byte
    Dim baOutput()  As Byte
    Dim dblTimer    As Double
    Dim bResult     As Boolean
    Dim uCtx        As CryptoAesContext
    
    Command13.Caption = "Processing"
    dblTimer = TimerEx
    baKey = FromHex("000102030405060708090a0b0c0d0e0f")
    For lIdx = 1 To ITER
        baOutput = m_baContents
        CryptoAesInit uCtx, baKey
        CryptoAesCbcEncrypt uCtx, baOutput
    Next
    Command13.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    dblTimer = TimerEx
    CryptoAesInit uCtx, baKey
    bResult = CryptoAesCbcDecrypt(uCtx, baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec - " & bResult
    Command13.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command14_Click()
    Const ITER      As Long = 1
    Dim lIdx        As Long
    Dim baKey()     As Byte
    Dim baOutput()  As Byte
'    Dim baOutput1() As Byte
'    Dim baOutput2() As Byte
    Dim dblTimer    As Double
    Dim uCtx        As CryptoAesContext
    
    Command14.Caption = "Processing"
    dblTimer = TimerEx
    baKey = FromHex("000102030405060708090a0b0c0d0e0f")
    For lIdx = 1 To ITER
        baOutput = m_baContents
        CryptoAesInit uCtx, baKey
        CryptoAesCtrCrypt uCtx, baOutput
'        baOutput1 = baOutput
    Next
    Command14.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    dblTimer = TimerEx
    CryptoAesInit uCtx, baKey
    CryptoAesCtrCrypt uCtx, baOutput
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command14.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    
'    dblTimer = TimerEx
'    AesChunkedInit baKey, 16
'    AesChunkedCryptArray m_baContents, baOutput
''    For lIdx = 0 To UBound(baOutput1)
''        If baOutput1(lIdx) <> baOutput2(lIdx) Then
''            Caption = lIdx
''            Exit For
''        End If
''    Next
'    Command14.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command15_Click()
    Const ITER      As Long = 1
    Dim lIdx        As Long
    Dim baKey()     As Byte
    Dim baOutput()  As Byte
    Dim dblTimer    As Double
    Dim bResult     As Boolean
    Dim uCtx        As CryptoAesContext
    
    Command15.Caption = "Processing"
    dblTimer = TimerEx
    baKey = FromHex("000102030405060708090a0b0c0d0e0f000102030405060708090a0b0c0d0e0f")
    For lIdx = 1 To ITER
        baOutput = m_baContents
        CryptoAesInit uCtx, baKey
        CryptoAesCbcEncrypt uCtx, baOutput
    Next
    Command15.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    dblTimer = TimerEx
    CryptoAesInit uCtx, baKey
    bResult = CryptoAesCbcDecrypt(uCtx, baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec - " & bResult
    Command15.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command16_Click()
    Const ITER      As Long = 1
    Dim lIdx        As Long
    Dim baKey()     As Byte
    Dim baOutput()  As Byte
    Dim dblTimer    As Double
    Dim uCtx        As CryptoAesContext
    
    Command16.Caption = "Processing"
    dblTimer = TimerEx
    baKey = FromHex("000102030405060708090a0b0c0d0e0f000102030405060708090a0b0c0d0e0f")
    For lIdx = 1 To ITER
        baOutput = m_baContents
        CryptoAesInit uCtx, baKey
        CryptoAesCtrCrypt uCtx, baOutput
    Next
    Command16.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    dblTimer = TimerEx
    CryptoAesInit uCtx, baKey
    CryptoAesCtrCrypt uCtx, baOutput
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    Command16.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command17_Click()
    Const ITER      As Long = 1
    Dim lIdx        As Long
    Dim baKey()     As Byte
    Dim baOutput()  As Byte
    Dim dblTimer    As Double
    Dim bResult     As Boolean
    Dim uCtx        As CryptoAesGcmContext
    
    Command17.Caption = "Processing"
    dblTimer = TimerEx
    baKey = FromHex("000102030405060708090a0b0c0d0e0f")
    For lIdx = 1 To ITER
        baOutput = m_baContents
        CryptoAesGcmInit uCtx, baKey, baKey, baKey
        CryptoAesGcmEncrypt uCtx, baOutput
    Next
    Command17.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    dblTimer = TimerEx
    CryptoAesGcmInit uCtx, baKey, baKey, baKey
    bResult = CryptoAesGcmDecrypt(uCtx, baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec - " & bResult
    Command17.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command18_Click()
    Const ITER      As Long = 1
    Dim lIdx        As Long
    Dim baKey()     As Byte
    Dim baOutput()  As Byte
    Dim dblTimer    As Double
    Dim bResult     As Boolean
    Dim uCtx        As CryptoAesGcmContext
    
    Command18.Caption = "Processing"
    dblTimer = TimerEx
    baKey = FromHex("000102030405060708090a0b0c0d0e0f000102030405060708090a0b0c0d0e0f")
    For lIdx = 1 To ITER
        baOutput = m_baContents
        CryptoAesGcmInit uCtx, baKey, baKey, baKey
        CryptoAesGcmEncrypt uCtx, baOutput
    Next
    Command18.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec"
    dblTimer = TimerEx
    CryptoAesGcmInit uCtx, baKey, baKey, baKey
    bResult = CryptoAesGcmDecrypt(uCtx, baOutput)
    Caption = Format$(TimerEx - dblTimer, "0.000") & " sec - " & bResult
    Command18.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (TimerEx - dblTimer), "0.000") & " MB/s"
End Sub
