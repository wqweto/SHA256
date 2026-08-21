Attribute VB_Name = "mdMain"
'=========================================================================
'
'  Pure VB6 Crypto (Untested)
'  Copyright (c) 2026 wqweto@gmail.com
'
'  This project is licensed under the terms of the MIT license
'  See the LICENSE file in the project root for more information
'
'=========================================================================
'--- mdMain.bas
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)

'=========================================================================
' API
'=========================================================================

#If Win64 Then
    Private Const PTR_SIZE                  As Long = 8
#Else
    Private Const PTR_SIZE                  As Long = 4
#End If
'--- for GetStdHandle
Private Const STD_OUTPUT_HANDLE             As Long = -11&
Private Const INVALID_HANDLE_VALUE          As Long = -1

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As LongPtr
Private Declare PtrSafe Function AllocConsole Lib "kernel32" () As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As LongPtr, pNumArgs As Long) As LongPtr
Private Declare PtrSafe Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As LongPtr
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As LongPtr, pNumArgs As Long) As LongPtr
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const MODULE_NAME           As String = "mdMain"

Private m_hStdOut                   As LongPtr
Private m_bNoConsole                As Boolean

'=========================================================================
' Error management
'=========================================================================

Private Sub PrintError(sFunction As String)
    ConPrintLine "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Const FUNC_NAME     As String = "Main"
    Dim vArgs           As Variant
    Dim sVerb           As String

    On Error GoTo EH
    vArgs = pvParseArgs(Command$)
    If UBound(vArgs) < 0 Then
        pvPrintUsage
        GoTo QH
    End If
    sVerb = LCase$(vArgs(0))
    Select Case sVerb
    Case "speed"
        BenchRun pvGetTail(vArgs)
    Case "test"
        VectorsRun pvGetTail(vArgs)
    Case "list"
        BenchList
    Case "help", "-h", "--help", "/?", "-?"
        pvPrintUsage
    Case Else
        ConPrintLine "Unknown command: " & sVerb
        ConPrintLine
        pvPrintUsage
    End Select
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvPrintUsage()
    ConPrintLine "vbcrypto -- Pure VB6 Crypto benchmark and test vector runner"
    ConPrintLine
    ConPrintLine "Usage:"
    ConPrintLine "  vbcrypto speed [algo ...]   measure throughput, openssl speed style"
    ConPrintLine "                 [-size N]    use these block sizes, e.g. 1M or 16,1K,1M"
    ConPrintLine "  vbcrypto test  [suite ...]  run test vectors and self checks"
    ConPrintLine "                 [-v]         list every failing case, not just the first"
    ConPrintLine "                 [-id N]      run only test case N and dump its inputs"
    ConPrintLine "  vbcrypto list               list known algorithm names"
    ConPrintLine "  vbcrypto help               this text"
    ConPrintLine
    ConPrintLine "Filters match on substring, so ""speed sha"" runs every SHA variant."
    ConPrintLine "With no filter every algorithm or suite runs."
    ConPrintLine
    ConPrintLine "Examples:"
    ConPrintLine "  vbcrypto speed sha256 blake3"
    ConPrintLine "  vbcrypto speed aes-128"
    ConPrintLine "  vbcrypto speed -size 1M"
    ConPrintLine "  vbcrypto test aes_gcm"
    ConPrintLine "  vbcrypto test kat"
    ConPrintLine "  vbcrypto test aes_ccm -v"
    ConPrintLine "  vbcrypto test aes_ccm -id 9"
End Sub

Public Sub ConPrint(sText As String)
    Dim baText()        As Byte
    Dim lWritten        As Long

    If m_hStdOut = 0 And Not m_bNoConsole Then
        m_hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
        '--- no console attached when running in the IDE, ask for one
        If m_hStdOut = 0 Or m_hStdOut = INVALID_HANDLE_VALUE Then
            Call AllocConsole
            m_hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
        End If
        If m_hStdOut = 0 Or m_hStdOut = INVALID_HANDLE_VALUE Then
            m_bNoConsole = True
        End If
    End If
    If LenB(sText) = 0 Then
        Exit Sub
    End If
    '--- fall back to the IDE debug window when there is nowhere else to go
    If m_bNoConsole Then
        Debug.Print sText;
        Exit Sub
    End If
    baText = StrConv(sText, vbFromUnicode)
    Call WriteFile(m_hStdOut, VarPtr(baText(0)), UBound(baText) + 1, lWritten, 0)
End Sub

Public Sub ConPrintLine(Optional sText As String)
    ConPrint sText & vbCrLf
End Sub

Public Function PadRight(ByVal sText As String, ByVal lWidth As Long) As String
    If Len(sText) >= lWidth Then
        PadRight = sText
    Else
        PadRight = sText & Space$(lWidth - Len(sText))
    End If
End Function

Public Function PadLeft(ByVal sText As String, ByVal lWidth As Long) As String
    If Len(sText) >= lWidth Then
        PadLeft = sText
    Else
        PadLeft = Space$(lWidth - Len(sText)) & sText
    End If
End Function

'--- bytes/sec with a magnitude suffix e.g. "65.1M"
Public Function FormatRate(ByVal dblRate As Double) As String
    If dblRate >= 1000000000# Then
        FormatRate = Format$(dblRate / 1000000000#, "0.0") & "G"
    ElseIf dblRate >= 1000000# Then
        FormatRate = Format$(dblRate / 1000000#, "0.0") & "M"
    ElseIf dblRate >= 1000# Then
        FormatRate = Format$(dblRate / 1000#, "0.0") & "k"
    Else
        FormatRate = Format$(dblRate, "0.0")
    End If
End Function

'--- operations/sec, more decimals the slower the operation
Public Function FormatOps(ByVal dblOps As Double) As String
    If dblOps >= 1000# Then
        FormatOps = Format$(dblOps, "0")
    ElseIf dblOps >= 10# Then
        FormatOps = Format$(dblOps, "0.0")
    Else
        FormatOps = Format$(dblOps, "0.000")
    End If
End Function

Public Function HasFilterMatch(ByVal sName As String, vFilter As Variant) As Boolean
    Dim lIdx            As Long

    If UBound(vFilter) < 0 Then
        HasFilterMatch = True
        Exit Function
    End If
    For lIdx = 0 To UBound(vFilter)
        If InStr(1, sName, vFilter(lIdx), vbTextCompare) > 0 Then
            HasFilterMatch = True
            Exit Function
        End If
    Next
End Function

'--- Command$ carries the arguments only, so prepend a program name for
'--- CommandLineToArgv to consume as argv[0] and skip it on the way out
Private Function pvParseArgs(sCmd As String) As Variant
    Dim sLine           As String
    Dim pArgv           As LongPtr
    Dim lCount          As Long
    Dim lIdx            As Long
    Dim pStr            As LongPtr
    Dim aRetVal()       As String

    sLine = "vbcrypto " & sCmd
    pArgv = CommandLineToArgv(StrPtr(sLine), lCount)
    If pArgv = 0 Or lCount <= 1 Then
        pvParseArgs = Split(vbNullString)
        GoTo QH
    End If
    ReDim aRetVal(0 To lCount - 2) As String
    For lIdx = 1 To lCount - 1
        Call CopyMemory(pStr, ByVal pArgv + lIdx * PTR_SIZE, PTR_SIZE)
        aRetVal(lIdx - 1) = pvToString(pStr)
    Next
    pvParseArgs = aRetVal
QH:
    If pArgv <> 0 Then
        Call LocalFree(pArgv)
    End If
End Function

Private Function pvToString(ByVal pStr As LongPtr) As String
    Call SysReAllocString(VarPtr(pvToString), pStr)
End Function

Private Function pvGetTail(vArgs As Variant) As Variant
    Dim aRetVal()       As String
    Dim lIdx            As Long

    If UBound(vArgs) < 1 Then
        pvGetTail = Split(vbNullString)
        Exit Function
    End If
    ReDim aRetVal(0 To UBound(vArgs) - 1) As String
    For lIdx = 1 To UBound(vArgs)
        aRetVal(lIdx - 1) = vArgs(lIdx)
    Next
    pvGetTail = aRetVal
End Function

'--- high resolution clock for the benchmark timing loops
Public Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency

    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Public Function ToUtf8Array(sText As String) As Byte()
    Const CP_UTF8       As Long = 65001
    Dim baRetVal()      As Byte
    Dim lSize           As Long

    ReDim baRetVal(0 To 4 * Len(sText)) As Byte
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), UBound(baRetVal) + 1, 0, 0)
    If lSize > 0 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
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

'--- the vector files are UTF-8, so this skips the encoding sniffing the
'--- general purpose reader in Module1.bas does and just decodes
Public Function ReadTextFile(sFile As String) As String
    Const CP_UTF8       As Long = 65001
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
    Dim lSize           As Long
    Dim lPos            As Long
    Dim lChars          As Long

    nFile = FreeFile
    Open sFile For Binary Access Read Shared As nFile
    lSize = LOF(nFile)
    If lSize > 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Get nFile, , baBuffer
    End If
    Close nFile
    If lSize = 0 Then
        Exit Function
    End If
    '--- step over an UTF-8 BOM when one is present
    If lSize >= 3 Then
        If baBuffer(0) = &HEF And baBuffer(1) = &HBB And baBuffer(2) = &HBF Then
            lPos = 3
        End If
    End If
    lChars = MultiByteToWideChar(CP_UTF8, 0, baBuffer(lPos), lSize - lPos, 0, 0)
    If lChars > 0 Then
        ReadTextFile = String$(lChars, 0)
        Call MultiByteToWideChar(CP_UTF8, 0, baBuffer(lPos), lSize - lPos, StrPtr(ReadTextFile), lChars)
    End If
End Function
