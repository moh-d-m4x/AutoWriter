@echo off
chcp 65001 >nul
echo ========================================
echo   كاتب المستندات - Build Script
echo ========================================
echo.

echo [1/6] Cleaning previous dist folder...
if exist "dist" rmdir /s /q "dist"
echo ✓ Clean complete
echo.

echo [2/6] Building web assets with Vite...
call npm run build
if %errorlevel% neq 0 (
    echo ERROR: Vite build failed!
    pause
    exit /b 1
)
echo ✓ Vite build complete
echo.

echo [3/6] Cleaning old Android assets...
:: Clean local android assets
if exist "android\app\src\main\assets\public\assets" (
    rmdir /s /q "android\app\src\main\assets\public\assets"
)
:: Clean build folder assets
if exist "C:\AutoWriter_Build\android\app\src\main\assets\public\assets" (
    rmdir /s /q "C:\AutoWriter_Build\android\app\src\main\assets\public\assets"
)
echo ✓ Old assets cleaned
echo.

echo [4/6] Syncing with Capacitor Android...
call npx cap sync android
if %errorlevel% neq 0 (
    echo ERROR: Capacitor sync failed!
    pause
    exit /b 1
)
echo ✓ Capacitor sync complete
echo.

echo [5/6] Cleaning and Building Android APK...
cd android
call gradlew.bat clean
call gradlew.bat assembleDebug
if %errorlevel% neq 0 (
    echo ERROR: Gradle build failed!
    pause
    exit /b 1
)
echo ✓ APK build complete
echo.

echo [6/6] Installing APK on device...
cd ..
"%LOCALAPPDATA%\Android\Sdk\platform-tools\adb.exe" install -r "android\app\build\outputs\apk\debug\app-debug.apk"
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
