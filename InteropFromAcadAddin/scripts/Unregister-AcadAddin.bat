@echo off
echo ==========================================
echo Unregistering InteropFromAcadAddin (.NET 8)
echo ==========================================

set COMHOST_PATH=%~dp0..\bin\x64\Debug\InteropFromAcadAddin.comhost.dll

if not exist "%COMHOST_PATH%" (
    echo WARNING: COM host DLL not found at %COMHOST_PATH%
    echo Attempting to unregister anyway...
    echo.
)

echo.
echo COM Host DLL: %COMHOST_PATH%
echo.

echo Unregistering with regsvr32...
regsvr32 /u /s "%COMHOST_PATH%"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ==========================================
    echo SUCCESS: COM unregistration completed
    echo ==========================================
) else (
    echo.
    echo ==========================================
    echo ERROR: Unregistration failed with code %ERRORLEVEL%
    echo ==========================================
    echo.
    echo This may be normal if the DLL was never registered.
)

echo.
pause
