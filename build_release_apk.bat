@echo off
chcp 65001 >nul

:: Change to the script's directory first
cd /d "%~dp0"

echo ========================================
echo   AutoWriter - Android APK Builder
echo ========================================
echo.

echo Working directory: %CD%
echo.

echo [1/6] Cleaning previous build...
if exist "dist" rmdir /s /q "dist"
if not exist "release\apk" mkdir "release\apk"
echo √ Clean complete
echo.

echo [2/6] Building web assets with Vite...
call npm run build
if %errorlevel% neq 0 (
    echo [ERROR] Vite build failed!
    pause
    exit /b 1
)
echo √ Vite build complete
echo.

echo [3/6] Cleaning old Android assets...
:: Clean local android assets
if exist "android\app\src\main\assets\public\assets" (
    rmdir /s /q "android\app\src\main\assets\public\assets"
)
echo √ Old assets cleaned
echo.

echo [4/6] Syncing with Capacitor Android...
call npx cap sync android
if %errorlevel% neq 0 (
    echo [ERROR] Capacitor sync failed!
    pause
    exit /b 1
)
echo √ Capacitor sync complete
echo.

echo [5/6] Building Android Debug APK...
cd /d "%~dp0android"
call gradlew.bat clean
call gradlew.bat assembleDebug
if %errorlevel% neq 0 (
    echo [ERROR] Gradle build failed!
    pause
    exit /b 1
)
echo √ APK build complete
echo.

echo [6/6] Copying to release folder...
cd /d "%~dp0"
if exist "release\apk\AutoWriter-1.1.0.apk" del /q "release\apk\AutoWriter-1.1.0.apk"
copy /y "android\app\build\outputs\apk\debug\app-debug.apk" "release\apk\AutoWriter-1.1.0.apk"
echo √ Release ready
echo.

echo ========================================
echo   Build Complete!
echo   Output: release\apk\AutoWriter-1.1.0.apk
echo.
echo   This APK is signed with debug key.
echo   Ready for direct installation.
echo ========================================
pause
