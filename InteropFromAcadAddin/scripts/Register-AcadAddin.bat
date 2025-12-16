@echo off
echo ========================================
echo Registering InteropFromAcadAddin (.NET 8)
echo ========================================

set COMHOST_PATH=%~dp0..\bin\x64\Debug\InteropFromAcadAddin.comhost.dll

if not exist "%COMHOST_PATH%" (
    echo ERROR: COM host DLL not found at %COMHOST_PATH%
    echo Please build the project first with: dotnet build -c Debug -p:Platform=x64
    pause
    exit /b 1
)

echo.
echo COM Host DLL: %COMHOST_PATH%
echo.

echo Registering with regsvr32...
regsvr32 /s "%COMHOST_PATH%"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo SUCCESS: COM registration completed
    echo ========================================
    echo.
    echo Next steps:
    echo 1. Launch AutoCAD 2026
    echo 2. Type NETLOAD
    echo 3. Browse to: %~dp0..\bin\x64\Debug\InteropFromAcadAddin.dll
    echo 4. Run the AcadDocEventsTester console app
) else (
    echo.
    echo ========================================
    echo ERROR: Registration failed with code %ERRORLEVEL%
    echo ========================================
    echo.
    echo Troubleshooting:
    echo - Make sure you're running as Administrator
    echo - Verify the DLL was built successfully
)

echo.
pause
