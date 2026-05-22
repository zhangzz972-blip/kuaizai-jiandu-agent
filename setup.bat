@echo off
chcp 65001 >nul
title KuaiZai JianDu Setup

echo.
echo ============================================
echo   KuaiZai JianDu agent v1.0 Setup
echo   Xuzhou Library
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Please install Python 3.10+
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python found

:: Install pip packages
echo.
echo [*] Installing Python packages...
pip install flask openai pandas openpyxl pillow python-docx playwright -q
if %errorlevel% neq 0 (
    echo [*] Retry with mirror...
    pip install flask openai pandas openpyxl pillow python-docx playwright -q -i https://pypi.tuna.tsinghua.edu.cn/simple
)
echo [OK] Python packages installed

:: Install Chromium
echo.
echo [*] Installing Chromium for poster (about 150MB)...
playwright install chromium
if %errorlevel% neq 0 (
    echo [!] Chromium install failed, poster feature may not work
)
echo [OK] Chromium installed

:: Check data files
echo.
if not exist "参考咨询书库馆藏清单.xlsx" (
    echo [!] File not found: 参考咨询书库馆藏清单.xlsx
)
if not exist "馆标黑版.png" (
    echo [!] File not found: 馆标黑版.png
)
echo [OK] Data files check done

:: API Key
echo.
echo Enter your Deepseek API Key:
echo (Get from https://platform.deepseek.com)
set /p API_KEY="  API Key: "

if "%API_KEY%"=="" (
    echo [!] API Key not set. You can set DEEPSEEK_API_KEY env var later.
) else (
    setx DEEPSEEK_API_KEY "%API_KEY%" >nul 2>&1
    echo [OK] API Key saved
)

echo.
echo ============================================
echo   Setup complete!
echo   Double-click 启动.bat to start
echo ============================================
echo.
pause

