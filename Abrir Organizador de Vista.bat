@echo off
REM Abre o Organizador de Vista sem a janela preta (console).
cd /d "%~dp0"
where pyw >nul 2>nul
if %errorlevel%==0 (
  start "" pyw "%~dp0Organizador de Vista.py"
) else (
  start "" pythonw "%~dp0Organizador de Vista.py"
)
