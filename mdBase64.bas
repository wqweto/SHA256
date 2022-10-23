Attribute VB_Name = "mdBase64"
'--- mdBase64.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As LongPtr, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As LongPtr, pcchString As Long) As Long
Private Declare PtrSafe Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As LongPtr, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As LongPtr, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
#Else
Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, pcchString As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
#End If

Public Function ToBase64Array(baData() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baData) >= 0 Then
        ToBase64Array = String$(2 * UBound(baData) + 6, 0)
        lSize = Len(ToBase64Array) + 1
        Call CryptBinaryToString(VarPtr(baData(0)), UBound(baData) + 1, 1, StrPtr(ToBase64Array), lSize)
        ToBase64Array = Left$(ToBase64Array, lSize)
    End If
End Function

Public Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim baOutput()      As Byte
    
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    If CryptStringToBinary(StrPtr(sText), Len(sText), 1, VarPtr(baOutput(0)), lSize) <> 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        FromBase64Array = baOutput
    End If
End Function
