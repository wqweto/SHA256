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
    ConPrintLine "  vbcrypto test  [suite ...]  run test vectors and self checks"
    ConPrintLine "  vbcrypto list               list known algorithm names"
    ConPrintLine "  vbcrypto help               this text"
    ConPrintLine
    ConPrintLine "Filters match on substring, so ""speed sha"" runs every SHA variant."
    ConPrintLine "With no filter every algorithm or suite runs."
    ConPrintLine
    ConPrintLine "Examples:"
    ConPrintLine "  vbcrypto speed sha256 blake3"
    ConPrintLine "  vbcrypto speed aes-128"
    ConPrintLine "  vbcrypto test aes_gcm"
    ConPrintLine "  vbcrypto test kat"
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
