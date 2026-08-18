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

