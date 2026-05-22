@echo off
chcp 65001 >nul
title KuaiZai JianDu agent v1.0

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Please run setup.bat first
    pause
    exit /b 1
)

if "%DEEPSEEK_API_KEY%"=="" (
    echo [!] DEEPSEEK_API_KEY not set. Please run setup.bat first.
    pause
    exit /b 1
)

echo.
echo   KuaiZai JianDu agent v1.0 starting...
echo.
echo   Browser will open: http://127.0.0.1:5000
echo   Press Ctrl+C to stop
echo.

start "" cmd /c "timeout /t 3 /nobreak >nul && start http://127.0.0.1:5000"

python app.py

