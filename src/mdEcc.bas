Attribute VB_Name = "mdEcc"
'=========================================================================
'
'  Pure VB6 Crypto (Untested)
'  Copyright (c) 2026 wqweto@gmail.com
'
'  This project is licensed under the terms of the MIT license
'  See the LICENSE file in the project root for more information
'
'=========================================================================
'--- mdEcc.bas
Option Explicit
DefObj A-Z

Private Const BCRYPT_ECDH_PUBLIC_GENERIC_MAGIC    As Long = &H504B4345  ' ECKP
Private Const BCRYPT_ECDH_PRIVATE_GENERIC_MAGIC   As Long = &H564B4345  ' ECKV
Private Const BCRYPT_NO_KEY_VALIDATION            As Long = &H8

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long
'--- bcrypt
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef hAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptImportKeyPair Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, ByRef hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptExportKey Lib "bcrypt" (ByVal hKey As Long, ByVal hExportKey As Long, ByVal pszBlobType As Long, pbOutput As Any, ByVal cbOutput As Long, ByRef cbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
Private Declare Function BCryptSecretAgreement Lib "bcrypt" (ByVal hPrivKey As Long, ByVal hPubKey As Long, ByRef phSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroySecret Lib "bcrypt" (ByVal hSecret As Long) As Long
Private Declare Function BCryptDeriveKey Lib "bcrypt" (ByVal hSharedSecret As Long, ByVal pwszKDF As Long, ByVal pParameterList As Long, ByVal pbDerivedKey As Long, ByVal cbDerivedKey As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long

Private Type BCRYPT_ECCKEY_BLOB
    dwMagic             As Long
    cbKey               As Long
    Buffer(0 To 1023)   As Byte
End Type
Private Const sizeof_BCRYPT_ECCKEY_BLOB As Long = 8

Private CURVE_NAME        As String  ' "secp192k1"
Private LNG_KEYSZ         As Long    ' 192 \ 8

Public Sub EccSetCurve(sName As String, ByVal lKeySize As Long)
    CURVE_NAME = sName
    LNG_KEYSZ = lKeySize
End Sub

Public Sub EccPrivateKey(baRetVal() As Byte, Optional Seed As Variant)
    If Not IsMissing(Seed) Then
        baRetVal = Seed
        ReDim Preserve baRetVal(0 To LNG_KEYSZ - 1) As Byte
    Else
        ReDim baRetVal(0 To LNG_KEYSZ - 1) As Byte
        Call RtlGenRandom(baRetVal(0), UBound(baRetVal) + 1)
    End If
    '--- clamp
'    baRetVal(0) = baRetVal(0) And &HF8
'    baRetVal(31) = baRetVal(31) And &H7F Or &H40
End Sub

Public Sub EccPublicKey(baRetVal() As Byte, baPriv() As Byte)
    Dim hAlg            As Long
    Dim hResult         As Long
    Dim hKey            As Long
    Dim lSize           As Long
    Dim uBlob           As BCRYPT_ECCKEY_BLOB
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    '--- setup curve for ECDH
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr("ECDH"), 0, 0)
    If hResult < 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider"
        GoTo QH
    End If
    hResult = BCryptSetProperty(hAlg, StrPtr("ECCCurveName"), StrPtr(CURVE_NAME), LenB(CURVE_NAME) + 2, 0)
    If hResult < 0 Then
        sApiSource = "BCryptSetProperty"
        GoTo QH
    End If
    '--- import private key
    uBlob.dwMagic = BCRYPT_ECDH_PRIVATE_GENERIC_MAGIC
    uBlob.cbKey = UBound(baPriv) + 1
    Call CopyMemory(uBlob.Buffer(2 * uBlob.cbKey), baPriv(0), UBound(baPriv) + 1)
    hResult = BCryptImportKeyPair(hAlg, 0, StrPtr("ECCPRIVATEBLOB"), hKey, uBlob, sizeof_BCRYPT_ECCKEY_BLOB + 3 * uBlob.cbKey, BCRYPT_NO_KEY_VALIDATION)
    If hResult < 0 Then
        sApiSource = "BCryptImportKeyPair"
        GoTo QH
    End If
    '--- export public key
    hResult = BCryptExportKey(hKey, 0&, StrPtr("ECCPUBLICBLOB"), uBlob, LenB(uBlob), lSize, 0)
    If hResult < 0 Then
        sApiSource = "BCryptExportKey"
        GoTo QH
    End If
    ReDim baRetVal(0 To 2 * uBlob.cbKey - 1) As Byte
    Call CopyMemory(baRetVal(0), uBlob.Buffer(0), 2 * uBlob.cbKey)
QH:
    On Error GoTo 0
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise hResult, , "Error &H" & Hex$(hResult) & " [" & sApiSource & "]"
    End If
    If IsArray(vErr) Then
        Err.Raise vErr(0), vErr(1), vErr(2)
    End If
    Exit Sub
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    Resume QH
End Sub

Public Sub EccSharedSecret(baRetVal() As Byte, baPriv() As Byte, baPub() As Byte)
    Dim hAlg            As Long
    Dim hResult         As Long
    Dim hPrivKey        As Long
    Dim hPubKey         As Long
    Dim hAgreedSecret   As Long
    Dim lSize           As Long
    Dim uBlob           As BCRYPT_ECCKEY_BLOB
    Dim lTemp           As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    '--- setup curve for ECDH
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr("ECDH"), 0, 0)
    If hResult < 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider"
        GoTo QH
    End If
    hResult = BCryptSetProperty(hAlg, StrPtr("ECCCurveName"), StrPtr(CURVE_NAME), LenB(CURVE_NAME) + 2, 0)
    If hResult < 0 Then
        sApiSource = "BCryptSetProperty"
        GoTo QH
    End If
    '--- import private key
    uBlob.dwMagic = BCRYPT_ECDH_PRIVATE_GENERIC_MAGIC
    uBlob.cbKey = UBound(baPriv) + 1
    Call CopyMemory(uBlob.Buffer(2 * uBlob.cbKey), baPriv(0), UBound(baPriv) + 1)
    hResult = BCryptImportKeyPair(hAlg, 0, StrPtr("ECCPRIVATEBLOB"), hPrivKey, uBlob, sizeof_BCRYPT_ECCKEY_BLOB + 3 * uBlob.cbKey, BCRYPT_NO_KEY_VALIDATION)
    If hResult < 0 Then
        sApiSource = "BCryptImportKeyPair(ECCPRIVATEBLOB)"
        GoTo QH
    End If
    '--- import public key
    uBlob.dwMagic = BCRYPT_ECDH_PUBLIC_GENERIC_MAGIC
    uBlob.cbKey = (UBound(baPub) + 1) \ 2
    Erase uBlob.Buffer
    Call CopyMemory(uBlob.Buffer(0), baPub(0), UBound(baPub) + 1)
    hResult = BCryptImportKeyPair(hAlg, 0, StrPtr("ECCPUBLICBLOB"), hPubKey, uBlob, sizeof_BCRYPT_ECCKEY_BLOB + 2 * uBlob.cbKey, 0)
    If hResult < 0 Then
        sApiSource = "BCryptImportKeyPair(ECCPUBLICBLOB)"
        GoTo QH
    End If
    '--- derive key agreement
    hResult = BCryptSecretAgreement(hPrivKey, hPubKey, hAgreedSecret, 0)
    If hResult < 0 Then
        sApiSource = "BCryptSecretAgreement"
        GoTo QH
    End If
    ReDim baRetVal(0 To 1023) As Byte
    hResult = BCryptDeriveKey(hAgreedSecret, StrPtr("TRUNCATE"), 0, VarPtr(baRetVal(0)), UBound(baRetVal) + 1, lSize, 0)
    If hResult < 0 Then
        sApiSource = "BCryptDeriveKey"
        GoTo QH
    End If
    ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    '--- reverse result to big endian
    For lSize = 0 To UBound(baRetVal) \ 2
        lTemp = baRetVal(lSize)
        baRetVal(lSize) = baRetVal(UBound(baRetVal) - lSize)
        baRetVal(UBound(baRetVal) - lSize) = lTemp
    Next
QH:
    On Error GoTo 0
    If hAgreedSecret <> 0 Then
        Call BCryptDestroySecret(hAgreedSecret)
    End If
    If hPubKey <> 0 Then
        Call BCryptDestroyKey(hPubKey)
    End If
    If hPrivKey <> 0 Then
        Call BCryptDestroyKey(hPrivKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise hResult, , "Error &H" & Hex$(hResult) & " [" & sApiSource & "]"
    End If
    If IsArray(vErr) Then
        Err.Raise vErr(0), vErr(1), vErr(2)
    End If
    Exit Sub
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    Resume QH
End Sub
