@echo off
rem SheetXpert - launcher for Windows (runs from source, installs fpdf2 if needed)
setlocal
cd /d "%~dp0"
set "PY=python"
set "PYW=pythonw"
where py >nul 2>nul && set "PY=py -3"
where pyw >nul 2>nul && set "PYW=pyw -3"

%PY% -c "import sys; sys.exit(0 if sys.version_info>=(3,8) else 1)" >nul 2>nul
if errorlevel 1 (
    echo Python 3.8 or newer was not found.
    echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
    pause
    exit /b 1
)

%PY% -c "import fpdf,sys; sys.exit(0 if int(fpdf.__version__.split('.')[0])>=2 else 1)" >nul 2>nul
if errorlevel 1 (
    echo Installing dependencies - internet needed once...
    %PY% -m pip uninstall -y fpdf >nul 2>nul
    %PY% -m pip install --user -r requirements.txt
    if errorlevel 1 (
        echo Could not install the dependencies.
        pause
        exit /b 1
    )
)

start "" %PYW% SheetXpert.py %*
