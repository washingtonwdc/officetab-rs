# ==============================================================================
# Instalar-OfficeTab.ps1
# Script de instalacao e registro do suplemento COM Office Tab (Rust)
# ==============================================================================

[CmdletBinding()]
param(
    [string]$DllPath = ""
)

$ErrorActionPreference = "Stop"

if (-not $DllPath) {
    if ($PSScriptRoot) {
        $DllPath = Join-Path $PSScriptRoot "target\release\officetab_rs.dll"
    } else {
        $DllPath = "C:\Users\washi\officetab-rs\target\release\officetab_rs.dll"
    }
}

if (-not (Test-Path $DllPath)) {
    $DllPath = "C:\Users\washi\officetab-rs\target\release\officetab_rs.dll"
}

Write-Host "=== Instalacao do Office Tab Pro (Rust) ===" -ForegroundColor Cyan

# 1. Validar existencia do binario da DLL
if (-not (Test-Path $DllPath)) {
    Write-Error "Arquivo DLL nao encontrado em '$DllPath'. Compile o projeto antes com 'cargo build --release'."
    exit 1
}

$FullPath = (Resolve-Path $DllPath).Path
Write-Host "[OK] DLL localizada: $FullPath" -ForegroundColor Green

$ProgId = "OfficeTabs.Connect"
$Clsid = "{1c8ba3cd-5248-4611-8a6e-8ccbaf89b651}"
$Descricao = "Office Tab Pro (Rust Reimplementation)"
$Aplicativos = @("Word", "Excel", "PowerPoint")

# 2. Registrar ProgID em HKCU:\Software\Classes
$ProgIdPath = "HKCU:\Software\Classes\$ProgId"
New-Item -Path $ProgIdPath -Force | Out-Null
Set-ItemProperty -Path $ProgIdPath -Name "(Default)" -Value $Descricao

$ProgIdClsidPath = "$ProgIdPath\CLSID"
New-Item -Path $ProgIdClsidPath -Force | Out-Null
Set-ItemProperty -Path $ProgIdClsidPath -Name "(Default)" -Value $Clsid
Write-Host "[OK] ProgID '$ProgId' registrado em HKCU:\Software\Classes" -ForegroundColor Green

# 3. Registrar CLSID e InprocServer32 em HKCU:\Software\Classes\CLSID
$ClsidPath = "HKCU:\Software\Classes\CLSID\$Clsid"
New-Item -Path $ClsidPath -Force | Out-Null
Set-ItemProperty -Path $ClsidPath -Name "(Default)" -Value $Descricao
Set-ItemProperty -Path $ClsidPath -Name "ProgID" -Value $ProgId

$InprocPath = "$ClsidPath\InprocServer32"
New-Item -Path $InprocPath -Force | Out-Null
Set-ItemProperty -Path $InprocPath -Name "(Default)" -Value $FullPath
Set-ItemProperty -Path $InprocPath -Name "ThreadingModel" -Value "Apartment"
Write-Host "[OK] CLSID '$Clsid' apontando para '$FullPath' com ThreadingModel Apartment" -ForegroundColor Green

# 4. Registrar Suplemento nos Aplicativos do Microsoft Office
foreach ($App in $Aplicativos) {
    $AddinPath = "HKCU:\Software\Microsoft\Office\$App\Addins\$ProgId"
    New-Item -Path $AddinPath -Force | Out-Null
    Set-ItemProperty -Path $AddinPath -Name "FriendlyName" -Value $Descricao
    Set-ItemProperty -Path $AddinPath -Name "Description" -Value $Descricao
    Set-ItemProperty -Path $AddinPath -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $AddinPath -Name "CommandLineSafe" -Value 0 -Type DWord
    Write-Host "[OK] Suplemento registrado para Microsoft $App (LoadBehavior=3)" -ForegroundColor Green
}

Write-Host "`n[SUCESSO] Office Tab Pro (Rust) instalado com sucesso no perfil do usuario!" -ForegroundColor Green
Write-Host "Ao abrir Word, Excel ou PowerPoint, o suplemento sera carregado automaticamente pelo Office." -ForegroundColor Yellow
