@echo off
chcp 65001 >nul

:: Change to the script's directory first
cd /d "%~dp0"

echo ========================================
echo   AutoWriter - Windows EXE Builder
echo ========================================
echo.

:: Check for admin rights
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] This script requires Administrator privileges!
    echo Please right-click and "Run as administrator"
    pause
    exit /b 1
)

echo Working directory: %CD%
echo.

echo [1/5] Cleaning previous build...
if exist "build_temp" rmdir /s /q "build_temp"
if not exist "release\exe" mkdir "release\exe"
echo √ Clean complete
echo.

echo [2/5] Building web assets with Vite...
call npm run build
if %errorlevel% neq 0 (
    echo [ERROR] Vite build failed!
    pause
    exit /b 1
)
echo √ Vite build complete
echo.

echo [3/5] Generating icons...
call node generate-icons.js
echo √ Icons generated
echo.

echo [4/5] Building Windows EXE installer...
call npx electron-builder --win nsis
if %errorlevel% neq 0 (
    echo [ERROR] Electron build failed!
    pause
    exit /b 1
)
echo √ EXE build complete
echo.

echo [5/5] Moving to release folder...
if exist "release\exe\AutoWriter-1.1.0-Setup.exe" del /q "release\exe\AutoWriter-1.1.0-Setup.exe"
move /y "build_temp\AutoWriter-1.1.0-Setup.exe" "release\exe\"
rmdir /s /q "build_temp"
echo √ Release ready
echo.

echo ========================================
echo   Build Complete!
echo   Output: release\exe\AutoWriter-1.1.0-Setup.exe
echo ========================================
pause
