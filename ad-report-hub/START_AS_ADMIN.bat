@echo off
:: AD Identity Tool - Run as Administrator
:: This script starts the application with elevated privileges

:: Check if running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with Administrator privileges...
    goto :start
) else (
    echo Requesting Administrator privileges...
    goto :elevate
)

:elevate
:: Self-elevate the script
powershell -Command "Start-Process -Verb RunAs -FilePath '%~dpnx0'"
exit /b

:start
echo.
echo ============================================
echo    AD Identity Tool - Administrator Mode
echo ============================================
echo.

cd /d "%~dp0"

:: Check if Node.js is installed
where node >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Node.js is not installed or not in PATH.
    echo Please install Node.js from https://nodejs.org
    pause
    exit /b 1
)

:: Check if node_modules exists
if not exist "node_modules" (
    echo Installing dependencies...
    call npm install
    if %errorLevel% neq 0 (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
)

echo.
echo Starting AD Identity Tool...
echo.
echo Frontend: http://localhost:3000
echo Backend:  http://localhost:5000
echo.
echo Press Ctrl+C to stop the server.
echo.

:: Start the application
call npm start

pause
