@echo off
title AutoWriter - Development Server
echo.
echo ========================================
echo     AutoWriter - Document Generator
echo ========================================
echo.

:: Set Node.js path
set PATH=C:\Program Files\nodejs;%PATH%

:: Change to app directory
cd /d "%~dp0"

echo Starting development server...
echo.
echo After the server starts, Electron will open automatically.
echo Close this window to stop the app.
echo.

:: Start Vite dev server and wait for it, then launch Electron
start /b cmd /c "npm run dev"

:: Wait for Vite to start (5 seconds)
echo Waiting for server to start...
timeout /t 5 /nobreak > nul

:: Start Electron
echo Launching AutoWriter...
npm run electron
