@echo off
:: AD Report Hub - Start with Current User Authentication
:: Uses the logged-in Windows user's credentials (no admin required)

echo.
echo ============================================
echo    AD Report Hub - Starting...
echo ============================================
echo.
echo Using current Windows authentication: %USERNAME%@%USERDOMAIN%
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
echo Starting AD Report Hub...
echo.
echo Frontend: http://localhost:3000
echo Backend:  http://localhost:5000
echo.
echo NOTE: Reports will use your current Windows credentials.
echo       Make sure you have appropriate AD read permissions.
echo.
echo Press Ctrl+C to stop the server.
echo.

:: Start the application
call npm start

pause
