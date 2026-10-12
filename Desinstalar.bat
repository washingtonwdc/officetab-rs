@echo off
setlocal
title Desinstalador Office Tab (Rust)
cd /d "%~dp0"

echo ========================================================
echo        Desinstalador Office Tab Pro (Rust)
echo ========================================================
echo.
echo Removendo registros COM do Office Tab...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Desinstalar-OfficeTab.ps1"

echo.
echo Pressione qualquer tecla para sair...
pause >nul
