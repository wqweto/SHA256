@echo off
setlocal

set "VB6=C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
set "LOG=%~dp0make.log"

rem --- the VB6 path holds "(x86)", so always quote it when echoing inside
rem --- a parenthesised block or the closing paren ends the block early
if not exist "%VB6%" (
    echo ERROR: VB6.EXE not found at "%VB6%"
    exit /b 1
)

rem --- build both projects, or just the one named on the command line
if not "%~1" == "" (
    call :build "%~dp0%~1" || exit /b 1
    goto :ran
)
call :build "%~dp0Benchmark.vbp" || exit /b 1
call :build "%~dp0BenchmarkSliced.vbp" || exit /b 1

:ran
exit /b 0

:build
set "PROJECT=%~1"
if not exist "%PROJECT%" (
    echo ERROR: "%PROJECT%" not found
    exit /b 1
)
rem --- start clean so the checks below cannot pass on stale output
if exist "%LOG%" del /q "%LOG%"

echo Compiling %~nx1 ...
rem --- VB6.EXE is a GUI app, so cmd will not wait for it on its own. /out
rem --- captures the build outcome and is written whether it worked or not.
start "" /wait "%VB6%" /make "%PROJECT%" /out "%LOG%"

if not exist "%LOG%" goto :buildok
find /i "succeeded" "%LOG%" >nul 2>&1
if errorlevel 1 goto :buildfail

:buildok
for /f "tokens=2 delims==" %%E in ('findstr /b "ExeName32=" "%PROJECT%"') do set "OUTEXE=%%~E"
echo Build OK: %~dp1%OUTEXE%
exit /b 0

:buildfail
echo.
echo BUILD FAILED
if exist "%LOG%" type "%LOG%"
exit /b 1
