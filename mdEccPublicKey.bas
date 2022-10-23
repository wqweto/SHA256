Attribute VB_Name = "mdEccPublicKey"
'--- mdEccPublicKey.bas
Option Explicit
DefObj A-Z

Private Const BCRYPT_ECDSA_PRIVATE_GENERIC_MAGIC  As Long = &H56444345  ' ECDV
Private Const BCRYPT_NO_KEY_VALIDATION            As Long = &H8

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef hAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptImportKeyPair Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, ByRef hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptExportKey Lib "bcrypt" (ByVal hKey As Long, ByVal hExportKey As Long, ByVal pszBlobType As Long, pbOutput As Any, ByVal cbOutput As Long, ByRef cbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long

Private Type BCRYPT_ECCKEY_BLOB
    dwMagic             As Long
    cbKey               As Long
    Buffer(0 To 1000)   As Byte
End Type
Private Const sizeof_BCRYPT_ECCKEY_BLOB As Long = 8

Public Function BCryptPublicKeyFromPrivate(sCurveName As String, baPrivKey() As Byte, baPubKey() As Byte) As Boolean
    Dim hAlg            As Long
    Dim hResult         As Long
    Dim hKey            As Long
    Dim lSize           As Long
    Dim uBlob           As BCRYPT_ECCKEY_BLOB
    
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr("ECDSA"), 0, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    hResult = BCryptSetProperty(hAlg, StrPtr("ECCCurveName"), StrPtr(sCurveName), LenB(sCurveName) + 2, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    uBlob.dwMagic = BCRYPT_ECDSA_PRIVATE_GENERIC_MAGIC
    uBlob.cbKey = UBound(baPrivKey) + 1
    '--- uBlob.Buffer is expected to contain first PubKey(0 to 63) then PrivKey(0 to 31) so copy baPrivKey at offset 64 and leave PubKey empty
    Call CopyMemory(uBlob.Buffer(2 * uBlob.cbKey), baPrivKey(0), UBound(baPrivKey) + 1)
    '--- flag BCRYPT_NO_KEY_VALIDATION is important to skip PubKey supplied (as we pass PrivKey only)
    hResult = BCryptImportKeyPair(hAlg, 0, StrPtr("ECCPRIVATEBLOB"), hKey, uBlob, sizeof_BCRYPT_ECCKEY_BLOB + 3 * uBlob.cbKey, BCRYPT_NO_KEY_VALIDATION)
    If hResult < 0 Then
        GoTo QH
    End If
    hResult = BCryptExportKey(hKey, 0&, StrPtr("ECCPUBLICBLOB"), uBlob, LenB(uBlob), lSize, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    ReDim baPubKey(0 To 2 * uBlob.cbKey) As Byte
    baPubKey(0) = 4 '--- 4 - uncompressed public key
    Call CopyMemory(baPubKey(1), uBlob.Buffer(0), 2 * uBlob.cbKey)
    '--- success
    BCryptPublicKeyFromPrivate = True
QH:
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
End Function

Public Function BCryptCompressPublicKey(baPubKey() As Byte) As Byte()
    Dim lSize           As Long
    Dim baRetVal()      As Byte
    
    If baPubKey(0) = 4 Then
        lSize = UBound(baPubKey)
        ReDim baRetVal(0 To lSize \ 2) As Byte
        '--- 2 - Y is even, 3 - Y is odd
        baRetVal(0) = IIf(baPubKey(lSize \ 2 + 1) < 128, 2, 3)
        Call CopyMemory(baRetVal(1), baPubKey(1), lSize \ 2)
        BCryptCompressPublicKey = baRetVal
    Else
        BCryptCompressPublicKey = baPubKey
    End If
End Function

Public Function GetRandomBytes(ByVal lSize As Long) As Byte()
    Dim baRetVal()      As Byte
    
    ReDim baRetVal(0 To lSize - 1) As Byte
    Call RtlGenRandom(baRetVal(0), lSize)
    GetRandomBytes = baRetVal
End Function

