@echo off
chcp 65001 >nul
echo ============================================
echo   DHL 匹配工具 - Windows 打包脚本
echo ============================================
echo.
echo  把本文件 + match_tool.py 放在同一目录下运行
echo.

REM 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python！请先安装 Python 3.8+
    echo 下载: https://www.python.org/downloads/
    echo 安装时务必勾选 "Add Python to PATH"
    pause
    exit /b 1
)
python --version
echo.

REM 安装依赖
echo [1/3] 安装 openpyxl...
pip install openpyxl -q
if %errorlevel% neq 0 (
    echo [错误] openpyxl 安装失败
    pause
    exit /b 1
)

REM 安装 PyInstaller
echo [2/3] 安装 PyInstaller...
pip install pyinstaller -q
if %errorlevel% neq 0 (
    echo [错误] PyInstaller 安装失败
    pause
    exit /b 1
)

REM 打包
echo [3/3] 打包为独立 .exe（约1-2分钟）...
pyinstaller --onefile --noconsole --name "DHL匹配工具" match_tool.py
if %errorlevel% neq 0 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

REM 复制到桌面
copy /Y "dist\DHL匹配工具.exe" "%USERPROFILE%\Desktop\DHL匹配工具.exe" >nul 2>&1
if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo   打包完成！
    echo   已复制到桌面: DHL匹配工具.exe
    echo   把这个 .exe 发给同事即可，无需安装任何东西
    echo ============================================
) else (
    echo.
    echo ============================================
    echo   打包完成！
    echo   文件在: dist\DHL匹配工具.exe
    echo ============================================
)
pause
