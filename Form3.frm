VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form3"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim baPrivKey()     As Byte
    Dim baPubKey()      As Byte
    
    baPrivKey = FromHex("081549973bafbba825b31bcc402a3c4ed8e3185c2f3a31c75e55f423e9629aa3")
    If BCryptPublicKeyFromPrivate("secP256k1", baPrivKey, baPubKey) Then
        Debug.Print ToHex(baPubKey)
        '-> 0443B337DEC65A47B3362C9620A6E6FF39A1DDFA908ABAB1666C8A30A3F8A7CCCCFC24A7914950B6405729A9313CEC6AE5BB4A082F92D05AC49DF4B6DD8387BFEB
        Debug.Print ToHex(BCryptCompressPublicKey(baPubKey))
        '-> 0343B337DEC65A47B3362C9620A6E6FF39A1DDFA908ABAB1666C8A30A3F8A7CCCC
    End If
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

Public Function FromHex(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    On Error GoTo QH
    '--- check for hexdump delimiter
    If sText Like "*[!0-9A-Fa-f]*" Then
        ReDim baRetVal(0 To Len(sText) \ 3) As Byte
        For lIdx = 1 To Len(sText) Step 3
            baRetVal(lIdx \ 3) = "&H" & Mid$(sText, lIdx, 2)
        Next
    ElseIf LenB(sText) <> 0 Then
        ReDim baRetVal(0 To Len(sText) \ 2 - 1) As Byte
        For lIdx = 1 To Len(sText) Step 2
            baRetVal(lIdx \ 2) = "&H" & Mid$(sText, lIdx, 2)
        Next
    Else
        baRetVal = vbNullString
    End If
    FromHex = baRetVal
QH:
End Function

