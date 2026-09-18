@echo off
setlocal
cd /d "%~dp0"

py -3 "%~dp0Organizador de Vista.py"
if errorlevel 1 (
    echo.
    echo Tentando abrir com python...
    python "%~dp0Organizador de Vista.py"
)

pause
