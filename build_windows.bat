@echo off
echo ============================================
echo   DHL Match Tool - Windows Build Script
echo ============================================
echo.
echo   Place this .bat and match_tool.py in the same folder, then run.
echo.

REM Check Python
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

REM Install dependencies
echo [1/3] Installing openpyxl...
pip install openpyxl -q
if %errorlevel% neq 0 (
    echo [ERROR] openpyxl install failed
    pause
    exit /b 1
)

REM Install PyInstaller
echo [2/3] Installing PyInstaller...
pip install pyinstaller -q
if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller install failed
    pause
    exit /b 1
)

REM Build exe
echo [3/3] Building standalone .exe (1-2 minutes)...
pyinstaller --onefile --noconsole --name "DHL_Match_Tool" match_tool.py
if %errorlevel% neq 0 (
    echo [ERROR] Build failed
    pause
    exit /b 1
)

REM Copy to desktop
copy /Y "dist\DHL_Match_Tool.exe" "%USERPROFILE%\Desktop\DHL_Match_Tool.exe" >nul 2>&1
if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo   Build complete!
    echo   Copied to Desktop: DHL_Match_Tool.exe
    echo   Share this .exe with colleagues - no setup needed.
    echo ============================================
) else (
    echo.
    echo ============================================
    echo   Build complete!
    echo   File: dist\DHL_Match_Tool.exe
    echo ============================================
)
pause
