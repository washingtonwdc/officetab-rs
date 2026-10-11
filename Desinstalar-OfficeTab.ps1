# ==============================================================================
# Desinstalar-OfficeTab.ps1
# Script de desinstalacao e remocao do suplemento COM Office Tab (Rust)
# ==============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = "Continue"

Write-Host "=== Desinstalacao do Office Tab Pro (Rust) ===" -ForegroundColor Cyan

$ProgId = "OfficeTabs.Connect"
$Clsid = "{1c8ba3cd-5248-4611-8a6e-8ccbaf89b651}"
$Aplicativos = @("Word", "Excel", "PowerPoint")

# 1. Remover entradas dos aplicativos Office
foreach ($App in $Aplicativos) {
    $AddinPath = "HKCU:\Software\Microsoft\Office\$App\Addins\$ProgId"
    if (Test-Path $AddinPath) {
        Remove-Item -Path $AddinPath -Recurse -Force
        Write-Host "[OK] Removido de Microsoft $App" -ForegroundColor Green
    }
}

# 2. Remover CLSID
$ClsidPath = "HKCU:\Software\Classes\CLSID\$Clsid"
if (Test-Path $ClsidPath) {
    Remove-Item -Path $ClsidPath -Recurse -Force
    Write-Host "[OK] Chave CLSID $Clsid removida" -ForegroundColor Green
}

# 3. Remover ProgID
$ProgIdPath = "HKCU:\Software\Classes\$ProgId"
if (Test-Path $ProgIdPath) {
    Remove-Item -Path $ProgIdPath -Recurse -Force
    Write-Host "[OK] Chave ProgID $ProgId removida" -ForegroundColor Green
}

Write-Host "`n[SUCESSO] Office Tab Pro (Rust) desinstalado do registro com sucesso!" -ForegroundColor Green
