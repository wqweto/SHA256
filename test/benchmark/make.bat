@echo off
setlocal

set "VB6=C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
set "PROJECT=%~dp0Benchmark.vbp"
set "OUTEXE=%~dp0vbcrypto.exe"
set "LOG=%~dp0make.log"

rem --- the VB6 path holds "(x86)", so always quote it when echoing inside
rem --- a parenthesised block or the closing paren ends the block early
if not exist "%VB6%" (
    echo ERROR: VB6.EXE not found at "%VB6%"
    exit /b 1
)
if not exist "%PROJECT%" (
    echo ERROR: "%PROJECT%" not found
    exit /b 1
)

rem --- start clean so the checks below cannot pass on stale output
if exist "%LOG%" del /q "%LOG%"
if exist "%OUTEXE%" del /q "%OUTEXE%"

echo Compiling Benchmark.vbp ...
rem --- VB6.EXE is a GUI app, so cmd will not wait for it on its own. /out
rem --- captures the build outcome and is written whether it worked or not.
start "" /wait "%VB6%" /make "%PROJECT%" /out "%LOG%"

if not exist "%OUTEXE%" goto :failed
if not exist "%LOG%" goto :succeeded
find /i "succeeded" "%LOG%" >nul 2>&1
if errorlevel 1 goto :failed

:succeeded
echo Build OK: %OUTEXE%
if /i "%~1" == "run" (
    echo.
    shift
    call "%OUTEXE%" %1 %2 %3 %4 %5 %6 %7 %8
)
exit /b 0

:failed
echo.
echo BUILD FAILED
if exist "%LOG%" type "%LOG%"
exit /b 1
