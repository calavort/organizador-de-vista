@echo off
REM Abre o Organizador de Vista SEM a janela preta (console).
REM Usado como reserva pelo atalho do Menu Iniciar (que abre direto no pyw.exe).
REM Para ver logs/erros, use "Abrir Organizador de Vista.bat".
cd /d "%~dp0"
where pyw >nul 2>nul
if %errorlevel%==0 (
  start "" pyw "%~dp0Organizador de Vista.py"
) else (
  start "" pythonw "%~dp0Organizador de Vista.py"
)
