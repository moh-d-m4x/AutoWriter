@echo off
chcp 65001 >nul
echo ========================================
echo   كاتب المستندات - Build Script
echo ========================================
echo.

echo [1/4] Building web assets with Vite...
call npm run build
if %errorlevel% neq 0 (
    echo ERROR: Vite build failed!
    pause
    exit /b 1
)
echo ✓ Vite build complete
echo.

echo [2/4] Syncing with Capacitor Android...
call npx cap sync android
if %errorlevel% neq 0 (
    echo ERROR: Capacitor sync failed!
    pause
    exit /b 1
)
echo ✓ Capacitor sync complete
echo.

echo [3/4] Cleaning and Building Android APK...
cd C:\AutoWriter_Build\android
call gradlew.bat clean
call gradlew.bat assembleDebug
if %errorlevel% neq 0 (
    echo ERROR: Gradle build failed!
    pause
    exit /b 1
)
echo ✓ APK build complete
echo.

echo [4/4] Installing APK on device...
"%LOCALAPPDATA%\Android\Sdk\platform-tools\adb.exe" install -r "app\build\outputs\apk\debug\app-debug.apk"
if %errorlevel% neq 0 (
    echo ERROR: APK installation failed!
    echo Make sure your device is connected and USB debugging is enabled.
    pause
    exit /b 1
)
echo ✓ APK installed successfully
echo.

echo ========================================
echo   Build Complete! App installed on device.
echo ========================================
pause
