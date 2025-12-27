@echo off
echo ===================================
echo    AutoWriter Icon Update Script
echo ===================================
echo.

echo [1/3] Generating icons from SVG...
call node generate-icons.js
if %errorlevel% neq 0 (
    echo ERROR: Icon generation failed!
    pause
    exit /b 1
)
echo.

echo [2/3] Building web assets...
call npm run build
if %errorlevel% neq 0 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)
echo.

echo [3/3] Syncing to Android...
call npx cap sync android
if %errorlevel% neq 0 (
    echo ERROR: Sync failed!
    pause
    exit /b 1
)
echo.

echo [*] Removing adaptive icons folder (to enable transparent PNG icons)...
if exist "android\app\src\main\res\mipmap-anydpi-v26" (
    rmdir /S /Q "android\app\src\main\res\mipmap-anydpi-v26"
    echo ✓ Removed mipmap-anydpi-v26
) else (
    echo   Already removed
)
echo.

echo ===================================
echo    Icons updated successfully!
echo ===================================
echo.
echo Ready to build APK.
pause
