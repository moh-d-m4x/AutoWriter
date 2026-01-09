@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ========================================
echo   Change App Version Tool
echo ========================================
echo.
echo Current version will be updated in:
echo   - package.json
echo   - android/app/build.gradle
echo   - build_release_apk.bat
echo   - build_release_exe.bat
echo.

set /p NEW_VERSION="Enter new version (e.g., 1.2.0): "

if "%NEW_VERSION%"=="" (
    echo ERROR: No version entered!
    pause
    exit /b 1
)

echo.
echo Updating to version %NEW_VERSION%...
echo.

:: Get current version from package.json
for /f "tokens=2 delims=:, " %%a in ('findstr /c:"\"version\":" package.json') do (
    set "OLD_VERSION=%%~a"
    goto :found
)
:found

echo Old version: %OLD_VERSION%
echo New version: %NEW_VERSION%
echo.

:: Update package.json
powershell -Command "(Get-Content 'package.json') -replace '\"version\": \"%OLD_VERSION%\"', '\"version\": \"%NEW_VERSION%\"' | Set-Content 'package.json'"
echo ✓ Updated package.json

:: Update android/app/build.gradle
powershell -Command "(Get-Content 'android\app\build.gradle') -replace 'versionName \"%OLD_VERSION%\"', 'versionName \"%NEW_VERSION%\"' | Set-Content 'android\app\build.gradle'"
echo ✓ Updated android/app/build.gradle

:: Update build_release_apk.bat
powershell -Command "(Get-Content 'build_release_apk.bat') -replace 'AutoWriter-%OLD_VERSION%', 'AutoWriter-%NEW_VERSION%' | Set-Content 'build_release_apk.bat'"
echo ✓ Updated build_release_apk.bat

:: Update build_release_exe.bat  
powershell -Command "(Get-Content 'build_release_exe.bat') -replace 'AutoWriter-%OLD_VERSION%', 'AutoWriter-%NEW_VERSION%' | Set-Content 'build_release_exe.bat'"
echo ✓ Updated build_release_exe.bat

echo.
echo ========================================
echo   Version updated to %NEW_VERSION%!
echo ========================================
pause
