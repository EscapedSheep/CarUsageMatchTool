@echo off
echo ============================================
echo   DHL Match Tool - Windows Build Script
echo ============================================
echo.
echo   Place this .bat and match_tool.py in the same folder, then run.
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found!
    echo Install Python 3.8+ from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)
python --version
echo.

echo [1/2] Installing dependencies...
pip install openpyxl Pillow pyinstaller -q
if %errorlevel% neq 0 (
    echo [ERROR] pip install failed. Try running as Administrator.
    pause
    exit /b 1
)
echo Dependencies OK.
echo.

echo [2/2] Building .exe (1-2 minutes)...
python -m PyInstaller --onefile --noconsole --name "DHL_Match_Tool" match_tool.py
if %errorlevel% neq 0 (
    echo [ERROR] Build failed.
    echo Try: python -m pip install pyinstaller --upgrade
    pause
    exit /b 1
)

copy /Y "dist\DHL_Match_Tool.exe" "%USERPROFILE%\Desktop\DHL_Match_Tool.exe" >nul 2>&1
if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo   SUCCESS!
    echo   Desktop: DHL_Match_Tool.exe
    echo   Share this file - no Python required.
    echo ============================================
) else (
    echo.
    echo ============================================
    echo   SUCCESS!
    echo   File: dist\DHL_Match_Tool.exe
    echo ============================================
)
pause
