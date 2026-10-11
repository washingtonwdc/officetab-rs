# ==============================================================================
# Desinstalar-OfficeTab.ps1
# Script de desinstalacao e remocao do suplemento COM Office Tab (Rust)
# ==============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = "Continue"

Write-Host "=== Desinstalacao do Office Tab Pro (Rust) ===" -ForegroundColor Cyan

$ProgIds = @("OfficeTabRs.Connect", "OfficeTabs.Connect")
$Clsids = @("{9a1929ef-1495-413e-993c-326dd15e35df}", "{1c8ba3cd-5248-4611-8a6e-8ccbaf89b651}")
$Aplicativos = @("Word", "Excel", "PowerPoint")

# 1. Remover entradas dos aplicativos Office
foreach ($ProgId in $ProgIds) {
    foreach ($App in $Aplicativos) {
        $AddinPath = "HKCU:\Software\Microsoft\Office\$App\Addins\$ProgId"
        if (Test-Path $AddinPath) {
            Remove-Item -Path $AddinPath -Recurse -Force
            Write-Host "[OK] Removido de Microsoft $App ($ProgId)" -ForegroundColor Green
        }
    }
}

# 2. Remover CLSIDs
foreach ($Clsid in $Clsids) {
    $ClsidPath = "HKCU:\Software\Classes\CLSID\$Clsid"
    if (Test-Path $ClsidPath) {
        Remove-Item -Path $ClsidPath -Recurse -Force
        Write-Host "[OK] Chave CLSID $Clsid removida" -ForegroundColor Green
    }
}

# 3. Remover ProgIDs
foreach ($ProgId in $ProgIds) {
    $ProgIdPath = "HKCU:\Software\Classes\$ProgId"
    if (Test-Path $ProgIdPath) {
        Remove-Item -Path $ProgIdPath -Recurse -Force
        Write-Host "[OK] Chave ProgID $ProgId removida" -ForegroundColor Green
    }
}

Write-Host "`n[SUCESSO] Office Tab Pro (Rust) desinstalado do registro com sucesso!" -ForegroundColor Green
