@echo off
title M365 AI Add-in

echo ========================================
echo   M365 AI Add-in Launcher
echo ========================================
echo.

cd /d "%~dp0"

:: Kill any existing server on port 3000
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":3000 " ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 1 >nul

:: Start http-server in background (for config.json access)
echo Starting local server...
start "M365 AI Server" /MIN cmd /c "node_modules\.bin\http-server . -p 3000 --ssl --cert certs\localhost.crt --key certs\localhost.key -c-1"
timeout /t 2 >nul

echo.
echo Which app do you want to open?
echo   [W] Word
echo   [E] Excel
echo   [P] PowerPoint
echo   [A] All three
echo.
set /p choice="Enter choice (W/E/P/A): "

if /i "%choice%"=="W" goto :word
if /i "%choice%"=="E" goto :excel
if /i "%choice%"=="P" goto :ppt
if /i "%choice%"=="A" goto :all
goto :word

:word
echo Launching Word...
node_modules\.bin\office-addin-debugging start manifest\manifest-word.xml desktop --app word
goto :end

:excel
echo Launching Excel...
node_modules\.bin\office-addin-debugging start manifest\manifest-excel.xml desktop --app excel
goto :end

:ppt
echo Launching PowerPoint...
node_modules\.bin\office-addin-debugging start manifest\manifest-ppt.xml desktop --app powerpoint
goto :end

:all
echo Launching Word...
start "" node_modules\.bin\office-addin-debugging start manifest\manifest-word.xml desktop --app word
timeout /t 3 >nul
echo Launching Excel...
start "" node_modules\.bin\office-addin-debugging start manifest\manifest-excel.xml desktop --app excel
timeout /t 3 >nul
echo Launching PowerPoint...
node_modules\.bin\office-addin-debugging start manifest\manifest-ppt.xml desktop --app powerpoint
goto :end

:end
echo.
echo Server is running. To stop, close the server window or run stop.bat.
echo.
pause
