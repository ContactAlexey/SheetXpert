@echo off
rem Builds dist\SheetXpert.exe (single file, no console window) with PyInstaller.
setlocal
cd /d "%~dp0"
set "PY=python"
where py >nul 2>nul && set "PY=py -3"

%PY% -m pip uninstall -y fpdf >nul 2>nul
%PY% -m pip install --user -r requirements.txt pyinstaller
if errorlevel 1 goto fail

%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name SheetXpert --icon assets\sheetxpert.ico --add-data "assets;assets" SheetXpert.py
if errorlevel 1 goto fail

echo.
echo Done: %~dp0dist\SheetXpert.exe
pause
exit /b 0

:fail
echo.
echo Build failed.
pause
exit /b 1
