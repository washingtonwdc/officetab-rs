; ==============================================================================
; Inno Setup 6 Script para Office Tab Pro (Rust)
; ==============================================================================

#define MyAppName "Office Tab Pro (Rust)"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "Washington / Open Source Community"
#define MyAppURL "https://github.com/washingtonwdc/officetab-rs"
#define MyAppExeName "officetab-config.exe"
#define MyAppDllName "officetab_rs.dll"

[Setup]
AppId={{9A1929EF-1495-413E-993C-326DD15E35DF}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\OfficeTab-rs
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=LICENSE
OutputDir=dist
OutputBaseFilename=OfficeTab-rs-Setup-v0.1.0
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
CloseApplications=yes
RestartApplications=no
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "target\release\{#MyAppDllName}"; DestDir: "{app}"; Flags: ignoreversion restartreplace uninsrestartdelete
Source: "target\release\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion restartreplace uninsrestartdelete
Source: "Instalar-OfficeTab.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "Desinstalar-OfficeTab.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "Testar-OfficeTab.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "Instalar.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "Desinstalar.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "Testar.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "Configurar.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Painel de Configurações do Office Tab"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Diagnóstico e Testes do Office Tab"; Filename: "{app}\Testar.bat"
Name: "{group}\README - Documentação"; Filename: "{app}\README.md"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Office Tab Configurações"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Registra o add-in no Office via script PowerShell que ajusta Word, Excel, PowerPoint, Access, Project, Visio, Publisher e Resiliency
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\Instalar-OfficeTab.ps1"" -DllPath ""{app}\{#MyAppDllName}"""; StatusMsg: "Registrando o suplemento no Microsoft Office..."; Flags: runhidden
; Oferece abrir a interface gráfica de configuração ao concluir
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir o Painel de Configurações agora"; Flags: nowait postinstall skipifsilent unchecked

[UninstallRun]
; Remove os registros COM e entradas do Office antes de deletar os arquivos
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\Desinstalar-OfficeTab.ps1"""; StatusMsg: "Removendo registros do suplemento no Microsoft Office..."; RunOnceId: "OfficeTabRsCleanUninstall"; Flags: runhidden
