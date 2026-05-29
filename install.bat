@echo off
title M365 AI Add-in - Install

echo.
echo ========================================
echo   M365 AI Add-in - One-time Installation
echo ========================================
echo.

cd /d "%~dp0"

echo [1/4] Checking Node.js...
node -v >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Node.js not found. Download: https://nodejs.org
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('node -v') do echo [OK] Node.js: %%i
echo.

echo [2/4] Installing dependencies...
call npm install --registry=https://registry.npmmirror.com
if %errorlevel% neq 0 (
    echo [ERROR] npm install failed.
    pause
    exit /b 1
)
echo [OK] Dependencies installed
echo.

echo [3/4] Setting up HTTPS certificate...
echo You may see a Windows security popup - click Yes.
if not exist "certs" mkdir "certs"

npx office-addin-dev-certs install --days 3650
if %errorlevel% equ 0 (
    copy /Y "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" "certs\localhost.crt" >nul 2>&1
    copy /Y "%USERPROFILE%\.office-addin-dev-certs\localhost.key" "certs\localhost.key" >nul 2>&1
    echo [OK] Certificate installed and copied
) else (
    echo [WARN] Certificate setup failed, trying fallback...
    node scripts\generate-certs.js
)
echo.

echo [4/4] Registering add-ins with Office...
echo Close all Office apps before continuing.
echo.

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "WordAIAssistant" /t REG_SZ /d "%~dp0manifest\manifest-word.xml" /f >nul 2>&1
if %errorlevel% equ 0 (echo [OK] Word  AI Assistant) else (echo [WARN] Word  registration failed)

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ExcelAIAssistant" /t REG_SZ /d "%~dp0manifest\manifest-excel.xml" /f >nul 2>&1
if %errorlevel% equ 0 (echo [OK] Excel AI Assistant) else (echo [WARN] Excel registration failed)

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "PPTAIAssistant" /t REG_SZ /d "%~dp0manifest\manifest-ppt.xml" /f >nul 2>&1
if %errorlevel% equ 0 (echo [OK] PPT   AI Assistant) else (echo [WARN] PPT   registration failed)

echo.
echo ========================================
echo   Done!
echo ========================================
echo.
echo How to use:
echo   1. Run start.bat to start the service
echo   2. Open Word / Excel / PowerPoint
echo   3. Find AI Assistant in Home tab or Developer tab
echo.
echo If add-in does not appear:
echo   - Close and reopen Office
echo   - Or run: npx office-addin-debugging start manifest\manifest-word.xml desktop --app word
echo.
pause
