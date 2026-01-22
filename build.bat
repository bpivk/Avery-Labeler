@echo off
echo ================================================
echo   Build Script
echo ================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7 or higher
    pause
    exit /b 1
)

echo Step 2: Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist Regalne Nalepke.spec del /q Regalne Nalepke.spec

echo.
echo Step 3: Building executable...
echo This may take a few minutes...
pyinstaller --onefile --windowed --name "Regalne Nalepke" labels.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ================================================
echo   Build Complete!
echo ================================================
echo.
echo Your executable is located at:
echo   dist\Regalne Nalepke.exe
echo.
echo You can now distribute this file to users.
echo No Python or dependencies required!
echo.
pause
