Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    
Public Function CryptoHash(baRetVal() As Byte, ByVal lHashAlgId As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const PROV_RSA_AES          As Long = 24
    Const CRYPT_VERIFYCONTEXT   As Long = &HF0000000
    Const HP_HASHVAL            As Long = 2
    Const LNG_FACILITY_WIN32    As Long = &H80070000
    Static hProv        As Long
    Dim hHash           As Long
    Dim lHashSize       As Long
    Dim lResult         As Long
    Dim sApiSource      As String
    
    If hProv = 0 Then
        If CryptAcquireContext(hProv, 0, 0, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then
            lResult = Err.LastDllResult
            sApiSource = "CryptAcquireContext"
            GoTo QH
        End If
    End If
    If CryptCreateHash(hProv, lHashAlgId, 0, 0, hHash) = 0 Then
        lResult = Err.LastDllResult
        sApiSource = "CryptCreateHash"
        GoTo QH
    End If
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    If Size > 0 Then
        If CryptHashData(hHash, baInput(Pos), Size, 0) = 0 Then
            lResult = Err.LastDllResult
            sApiSource = "CryptHashData"
            GoTo QH
        End If
    End If
    lHashSize = 1024
    ReDim baRetVal(0 To lHashSize - 1) As Byte
    If CryptGetHashParam(hHash, HP_HASHVAL, baRetVal(0), lHashSize, 0) = 0 Then
        lResult = Err.LastDllResult
        sApiSource = "CryptGetHashParam"
        GoTo QH
    End If
    ReDim Preserve baRetVal(0 To lHashSize - 1) As Byte
    '--- success
    CryptoHash = True
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(lResult < 0, lResult, lResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Public Function ToHex(baData() As Byte) As String
    Dim lIdx            As Long
    Dim sByte           As String
    
    ToHex = String$(UBound(baData) * 2 + 2, 48)
    For lIdx = 0 To UBound(baData)
        sByte = LCase$(Hex$(baData(lIdx)))
        Mid$(ToHex, lIdx * 2 + 3 - Len(sByte)) = sByte
    Next
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

Public Function ReadTextFile(sFile As String) As String
    Dim BOM_UTF8        As String: BOM_UTF8 = Chr$(&HEF) & Chr$(&HBB) & Chr$(&HBF)
    Dim BOM_UNICODE     As String: BOM_UNICODE = Chr$(&HFF) & Chr$(&HFE)
    Const ForReading    As Long = 1
    Dim lSize           As Long
    Dim sPrefix         As String
    
    With CreateObject("Scripting.FileSystemObject")
        lSize = .GetFile(sFile).Size
        If lSize = 0 Then
            Exit Function
        End If
        sPrefix = .OpenTextFile(sFile, ForReading).Read(IIf(lSize < 50, lSize, 50))
        If Left$(sPrefix, Len(BOM_UTF8)) <> BOM_UTF8 And Left$(sPrefix, Len(BOM_UNICODE)) <> BOM_UNICODE Then
            '--- special xml encoding test
            If InStr(1, sPrefix, "<?xml", vbTextCompare) > 0 And InStr(1, sPrefix, "utf-8", vbTextCompare) > 0 Then
                sPrefix = BOM_UTF8
            End If
        End If
        If Left$(sPrefix, Len(BOM_UTF8)) <> BOM_UTF8 Then
            On Error GoTo QH
            ReadTextFile = .OpenTextFile(sFile, ForReading, False, Left$(sPrefix, Len(BOM_UNICODE)) = BOM_UNICODE Or IsTextUnicode(ByVal sPrefix, Len(sPrefix), &HFFFF& - 2) <> 0).ReadAll()
        Else
            With CreateObject("ADODB.Stream")
                .Open
                If Left$(sPrefix, Len(BOM_UNICODE)) = BOM_UNICODE Then
                    .Charset = "Unicode"
                ElseIf Left$(sPrefix, Len(BOM_UTF8)) = BOM_UTF8 Then
                    .Charset = "UTF-8"
                Else
                    .Charset = "_autodetect_all"
                End If
                .LoadFromFile sFile
                ReadTextFile = .ReadText
            End With
        End If
    End With
QH:
End Function

Public Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

