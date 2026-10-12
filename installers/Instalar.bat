@echo off
setlocal
title Instalador Office Tab (Rust)
cd /d "%~dp0"

echo ========================================================
echo        Instalador Office Tab Pro (Rust) - Multi-App
echo ========================================================
echo.
echo Registrando suplemento COM para Word, Excel, PowerPoint,
echo Access, Project, Visio e Publisher...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Instalar-OfficeTab.ps1"

echo.
echo Pressione qualquer tecla para sair...
pause >nul
