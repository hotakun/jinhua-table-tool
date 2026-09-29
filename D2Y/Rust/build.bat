@echo off
echo Building jinhua_engine (cdylib)...
cargo build --release
if errorlevel 1 (
    echo Build failed!
    pause
    exit /b 1
)
echo Copying DLL to both GUIs (tkinter + PySide6)...
copy /y target\release\jinhua_engine.dll ..\jinhua_engine.dll >nul
copy /y target\release\jinhua_engine.dll ..\..\D2Y-pyqt\jinhua_engine.dll >nul
echo Done. Remember to re-pack (PyInstaller + Inno Setup) if you release.
pause
