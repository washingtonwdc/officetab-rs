@echo off
setlocal
title Configuracoes Office Tab (Rust)
cd /d "%~dp0"

if exist "%~dp0officetab-config.exe" (
    start "" "%~dp0officetab-config.exe"
    exit /b 0
)

if exist "%~dp0target\release\officetab-config.exe" (
    start "" "%~dp0target\release\officetab-config.exe"
    exit /b 0
)

echo [ERRO] officetab-config.exe nao encontrado.
pause
