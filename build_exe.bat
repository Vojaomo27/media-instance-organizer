@echo off
setlocal
cd /d "%~dp0"
python -m PyInstaller --noconfirm --clean --onefile --windowed --noupx --name MediaInstanceOrganizer --hidden-import=openpyxl --exclude-module=pandas --exclude-module=numpy --exclude-module=pytest --exclude-module=lxml --exclude-module=psutil move.py
if errorlevel 1 (
    echo.
    echo Build failed.
    pause
    exit /b 1
)
echo.
echo Build complete: dist\MediaInstanceOrganizer.exe
pause
