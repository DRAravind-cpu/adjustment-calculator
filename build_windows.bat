@echo off
echo Building Energy Adjustment Calculator for Windows...
echo.

REM Install build dependencies
echo Installing PyInstaller...
pip install pyinstaller>=5.0

REM Clean previous builds
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

REM Build the application
echo Building executable...
pyinstaller app.spec

REM Check if build was successful
if exist "dist\EnergyAdjustmentCalculator.exe" (
    echo.
    echo ========================================
    echo Build completed successfully!
    echo ========================================
    echo Executable location: dist\EnergyAdjustmentCalculator.exe
    echo.
    echo You can now distribute the 'dist' folder to users.
    echo Users just need to run EnergyAdjustmentCalculator.exe
    echo.
) else (
    echo.
    echo ========================================
    echo Build failed!
    echo ========================================
    echo Please check the error messages above.
)

pause
