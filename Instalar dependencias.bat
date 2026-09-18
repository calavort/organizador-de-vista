@echo off
setlocal
cd /d "%~dp0"

py -3 -m pip install --upgrade pip
py -3 -m pip install -r "%~dp0requirements.txt"

if errorlevel 1 (
    echo.
    echo Tentando com python...
    python -m pip install --upgrade pip
    python -m pip install -r "%~dp0requirements.txt"
)

pause
