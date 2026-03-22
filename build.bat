@echo off
echo ========================================
echo  TxtConverter Build Script
echo ========================================
echo.

REM Install dependencies
echo [1/2] Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install packages.
    pause
    exit /b 1
)

echo.
echo [2/2] Building exe...
python -m PyInstaller --onefile --windowed --name TxtConverter txt_converter.py
if errorlevel 1 (
    echo ERROR: Failed to build exe.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Build complete!
echo  Output: dist\TxtConverter.exe
echo ========================================
pause
