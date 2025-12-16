@echo off
setlocal
cd /d "%~dp0"

REM Unregister AutoCAD add-in COM types (x64)
REM Usage: double-click to unregister Debug build, or pass CONFIG=Release
REM Example: Unregister-AcadAddin.bat CONFIG=Release

set CONFIG=%CONFIG%
if "%CONFIG%"=="" set CONFIG=Debug

REM Prefer x64 output path, fallback to non-arch path
set DLL=..\bin\x64\%CONFIG%\InteropFromAcadAddin.dll
if not exist "%DLL%" set DLL=..\bin\%CONFIG%\InteropFromAcadAddin.dll

if not exist "%DLL%" (
    echo DLL not found: %DLL%
    echo Build the project for configuration %CONFIG% x64 and try again.
    pause
    exit /b 1
)

set REGASM64=%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
if not exist "%REGASM64%" (
    echo regasm not found at %REGASM64%
    echo Ensure .NET Framework 4.x is installed.
    pause
    exit /b 1
)

"%REGASM64%" /unregister "%DLL%"

if %errorlevel% neq 0 (
    echo Unregistration failed with error %errorlevel%.
    pause
    exit /b %errorlevel%
)

echo Unregistration completed successfully.
pause
