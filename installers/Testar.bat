@echo off
setlocal
title Diagnostico Office Tab (Rust)
cd /d "%~dp0"

echo ========================================================
echo        Diagnostico e Verificacao - Office Tab (Rust)
echo ========================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Testar-OfficeTab.ps1"

echo.
echo Pressione qualquer tecla para sair...
pause >nul
