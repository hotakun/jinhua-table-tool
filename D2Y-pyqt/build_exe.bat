@echo off
echo ========================================
echo   JinhuaJuhuo v3.1.2 - PySide6 Build
echo ========================================

echo Checking Python...
python --version
if errorlevel 1 (
    echo [ERROR] Python not found.
    pause
    exit /b 1
)

echo [1/4] Installing dependencies...
pip install PySide6 pyinstaller
if errorlevel 1 (
    echo [ERROR] Failed to install.
    pause
    exit /b 1
)

echo [2/4] Cleaning old build...
rmdir /s /q build dist 2>nul

echo [3/4] Building EXE...
python -m PyInstaller --noconsole --onedir --name JinhuaH --icon favicon.ico ^
    --add-data "favicon.ico;." ^
    --add-data "jinhua_engine.dll;." ^
    --add-data "MLSX.gif;." ^
    main.py
if errorlevel 1 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo [4/4] Done.
echo Output: dist\JinhuaH\JinhuaH.exe
pause
