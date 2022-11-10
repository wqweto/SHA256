VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   4332
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6012
   LinkTopic       =   "Form5"
   ScaleHeight     =   4332
   ScaleWidth      =   6012
   StartUpPosition =   3  'Windows Default
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
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_baContents() As Byte

Private Sub Form_Load()
    m_baContents = ReadBinaryFile("D:\TEMP\VirtualBox-6.1.34-150636-Win.exe")
    Caption = UBound(m_baContents) + 1 & " bytes ready"
End Sub

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
    dblTimer = Timer
    For lIdx = 1 To ITER
        baOutput = CryptoMd5ByteArray(m_baContents)
    Next
    Caption = Format$(Timer - dblTimer, "0.000") & " sec"
    Command1.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (Timer - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command2_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command2.Caption = "Processing"
    dblTimer = Timer
    For lIdx = 1 To ITER
        baOutput = CryptoSha1ByteArray(m_baContents)
    Next
    Caption = Format$(Timer - dblTimer, "0.000") & " sec"
    Command2.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (Timer - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command3_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command3.Caption = "Processing"
    dblTimer = Timer
    For lIdx = 1 To ITER
        baOutput = CryptoSha2ByteArray(256, m_baContents)
    Next
    Caption = Format$(Timer - dblTimer, "0.000") & " sec"
    Command3.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (Timer - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command3b_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command3b.Caption = "Processing"
    dblTimer = Timer
    For lIdx = 1 To ITER
        baOutput = CryptoSha2ByteArray(512, m_baContents)
    Next
    Caption = Format$(Timer - dblTimer, "0.000") & " sec"
    Command3b.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (Timer - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command4_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command4.Caption = "Processing"
    dblTimer = Timer
    For lIdx = 1 To ITER
        baOutput = CryptoSha3ByteArray(256, m_baContents)
    Next
    Caption = Format$(Timer - dblTimer, "0.000") & " sec"
    Command4.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (Timer - dblTimer), "0.000") & " MB/s"
End Sub

Private Sub Command5_Click()
    Const ITER As Long = 1
    Dim lIdx As Long
    Dim baOutput()  As Byte
    Dim dblTimer As Double
    
    Command5.Caption = "Processing"
    dblTimer = Timer
    For lIdx = 1 To ITER
        baOutput = CryptoBlake2sByteArray(256, m_baContents)
    Next
    Caption = Format$(Timer - dblTimer, "0.000") & " sec"
    Command5.Caption = Format$((UBound(m_baContents) + 1) * ITER / 1024# / 1024# / (Timer - dblTimer), "0.000") & " MB/s"
End Sub


