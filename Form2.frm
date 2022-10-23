VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim baInput()       As Byte
    Dim baHash()        As Byte
    
    baInput = StrConv("abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", vbFromUnicode)
    CryptoSha3 224, baHash, baInput
    Debug.Print ToHex(baHash)
    '-> 543e6868e1666c1a643630df77367ae5a62a85070a51c14cbf665cbc
    
    CryptoSha3 256, baHash, baInput
    Debug.Print ToHex(baHash)
    '-> 916f6061fe879741ca6469b43971dfdb28b1a32dc36cb3254e812be27aad1d18
    
    CryptoSha3 384, baHash, baInput
    Debug.Print ToHex(baHash)
    '-> 79407d3b5916b59c3e30b09822974791c313fb9ecc849e406f23592d04f625dc8c709b98b43b3852b337216179aa7fc7
    
    CryptoSha3 512, baHash, baInput
    Debug.Print ToHex(baHash)
    '-> afebb2ef542e6579c50cad06d2e578f9f8dd6881d7dc824d26360feebf18a4fa73e3261122948efcfd492e74e82e2189ed0fb440d187f382270cb455f21dd185
    
    CryptoShake 128, baHash, 32, baInput, Size:=0
    Debug.Print ToHex(baHash)
    '-> 7f9c2ba4e88f827d616045507605853ed73b8093f6efbc88eb1a6eacfa66ef26
    
    CryptoShake 256, baHash, 64, baInput, Size:=0
    Debug.Print ToHex(baHash)
    '-> 46b9dd2b0ba88d13233b3feb743eeb243fcd52ea62b81b82b50c27646ed5762fd75dc4ddd8c0f200cb05019d67b592f6fc821c49479ab48640292eacb3b7c4be
End Sub

Public Function ToHex(baText() As Byte, Optional Delimiter As String) As String
    Dim aText()         As String
    Dim lIdx            As Long
    
    If LenB(CStr(baText)) <> 0 Then
        ReDim aText(0 To UBound(baText)) As String
        For lIdx = 0 To UBound(baText)
            aText(lIdx) = Right$("0" & Hex$(baText(lIdx)), 2)
        Next
        ToHex = LCase$(Join(aText, Delimiter))
    End If
End Function

'Public Function FromHex(sText As String) As Byte()
'    Dim baRetVal()      As Byte
'    Dim lIdx            As Long
'
'    On Error GoTo QH
'    '--- check for hexdump delimiter
'    If sText Like "*[!0-9A-Fa-f]*" Then
'        ReDim baRetVal(0 To Len(sText) \ 3) As Byte
'        For lIdx = 1 To Len(sText) Step 3
'            baRetVal(lIdx \ 3) = "&H" & Mid$(sText, lIdx, 2)
'        Next
'    ElseIf LenB(sText) <> 0 Then
'        ReDim baRetVal(0 To Len(sText) \ 2 - 1) As Byte
'        For lIdx = 1 To Len(sText) Step 2
'            baRetVal(lIdx \ 2) = "&H" & Mid$(sText, lIdx, 2)
'        Next
'    Else
'        baRetVal = vbNullString
'    End If
'    FromHex = baRetVal
'QH:
'End Function
