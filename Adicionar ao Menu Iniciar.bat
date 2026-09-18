@echo off
setlocal
title Organizador de Vista - Adicionar ao Menu Iniciar

set "APP_DIR=%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $ErrorActionPreference = 'Stop'; $appDir = [IO.Path]::GetFullPath($env:APP_DIR); $appScript = Join-Path $appDir 'Organizador de Vista.pyw'; $icon = Join-Path $appDir 'interface\organizador_de_vista.ico'; $pyw = (Get-Command pyw.exe -ErrorAction SilentlyContinue).Source; if (-not $pyw) { $pyw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source }; if (-not $pyw) { throw 'Python launcher pyw.exe/pythonw.exe nao encontrado. Instale o Python ou execute Instalar dependencias.bat.' }; $startMenu = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'; New-Item -ItemType Directory -Force -Path $startMenu | Out-Null; $shortcutPath = Join-Path $startMenu 'Organizador de Vista.lnk'; $shell = New-Object -ComObject WScript.Shell; $shortcut = $shell.CreateShortcut($shortcutPath); $shortcut.TargetPath = $pyw; $shortcut.Arguments = [char]34 + $appScript + [char]34; $shortcut.WorkingDirectory = $appDir; if (Test-Path -LiteralPath $icon) { $shortcut.IconLocation = $icon + ',0' }; $shortcut.Description = 'Organizador de Vista - Tekla'; $shortcut.Save(); Write-Host ('Atalho criado no Menu Iniciar: ' + $shortcutPath) }"

if errorlevel 1 (
    echo.
    echo Nao foi possivel criar o atalho.
    pause
    exit /b 1
)

REM O Windows guarda uma copia dos icones em cache. Sem limpar, o Menu Iniciar
REM continua mostrando a versao antiga (borrada) mesmo com o icone novo.
echo.
echo Atualizando o cache de icones do Windows...
ie4uinit.exe -show >nul 2>nul
if errorlevel 1 ie4uinit.exe -ClearIconCache >nul 2>nul

echo.
echo Organizador de Vista foi adicionado ao Menu Iniciar.
echo.
echo Se o icone ainda aparecer com a aparencia antiga, faca logoff/login
echo ou reinicie o Explorer (Gerenciador de Tarefas ^> Explorer ^> Reiniciar).
pause
