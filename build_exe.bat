@echo off
setlocal

REM ── Build script for RebateFormGenerator.exe ──────────────────────────────
REM Builds to C:\Temp\RFGen\dist to avoid OneDrive rename issues,
REM then copies the result back to dist\ in the project folder.

set PROJECT_DIR=%~dp0
set BUILD_TEMP=C:\Temp\RFGen
set DIST_OUT=%PROJECT_DIR%dist

echo [1/4] Cleaning previous build...
if exist "%BUILD_TEMP%" rmdir /s /q "%BUILD_TEMP%"
if exist "%DIST_OUT%" rmdir /s /q "%DIST_OUT%"

echo [2/4] Installing / verifying dependencies...
call "%PROJECT_DIR%.venv\Scripts\activate.bat"
pip install pyinstaller --quiet

echo [3/4] Running PyInstaller...
pyinstaller "%PROJECT_DIR%RebateFormGenerator.spec" ^
    --distpath "%BUILD_TEMP%\dist" ^
    --workpath "%BUILD_TEMP%\build" ^
    --noconfirm

if errorlevel 1 (
    echo [ERROR] PyInstaller failed.
    exit /b 1
)

echo [4/4] Copying output to dist\...
xcopy /e /i /y "%BUILD_TEMP%\dist\RebateFormGenerator" "%DIST_OUT%\RebateFormGenerator"

echo.
echo ===== Build complete =====
echo Executable: %DIST_OUT%\RebateFormGenerator\RebateFormGenerator.exe
pause
