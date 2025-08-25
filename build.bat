@echo off
chcp 936 >nul
echo ========================================
echo Screen OCR Tool v2.0 - Build Script
echo ========================================


:: Activate virtual environment
echo Activating virtual environment...
call pack_env\Scripts\activate.bat


:: Clean previous build files
echo Cleaning build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

:: Start optimized packaging
echo Starting optimized packaging...
pyinstaller --onefile --windowed --name "ScreenOCR_v2.0" main.py

if %errorlevel% neq 0 (
    echo Packaging failed!
    pause
    exit /b 1
)

:: Copy necessary files to dist directory
echo Copying configuration files...
if exist ".env.example" copy ".env.example" "dist\"
if exist "README.md" copy "README.md" "dist\"
if exist "requirements.txt" copy "requirements.txt" "dist\"

:: Create logs and debug_images directories
mkdir "dist\logs" 2>nul
mkdir "dist\debug_images" 2>nul

echo ========================================
echo Packaging completed!
echo Output directory: dist\
echo Executable file: dist\ScreenOCR_v2.0.exe
echo ========================================

:: Display file size
for %%f in (dist\ScreenOCR_v2.0.exe) do echo File size: %%~zf bytes

pause
