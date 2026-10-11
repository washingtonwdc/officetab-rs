# Office Tab Rust (`officetab-rs`)

Biblioteca DLL em Rust para suplemento COM in-process do Microsoft Office (Word, Excel e PowerPoint) com suporte a gerenciamento de abas via subclassing nativo do Windows.

## 🏛️ Arquitetura

- **`com_addin_simple`**: Implementação das interfaces COM `IDTExtensibility2` e `IClassFactory`.
- **`implementation`**: Estrutura `OfficeTab` integrando o ciclo de vida do suplemento com o scanner de documentos.
- **`tab_manager`**: Descoberta de janelas de documentos Office (`OpusApp`, `XLMAIN`, `PPTFrameClass`), associação de abas e reparentamento.
- **`subclass`**: Subclassing estável e seguro via Comctl32 / Shell API (`SetWindowSubclass`, `DefSubclassProc`).
- **`tab_control`**: Wrapper do controle nativo Win32 `SysTabControl32`.
- **`registry_simple`**: Mecanismo de registro do servidor COM e suplemento no Registro do Windows.
- **`utils`**: Suporte a DPI awareness por monitor v2, inicialização COM STA e resolução de caminhos de módulos.

## 🔨 Compilação

Requer Rust estável (edição 2024) no Windows:

```powershell
cargo build --release
```

O binário compilado é gerado em:
`target\release\officetab_rs.dll`

## 🧪 Verificação e Testes

```powershell
cargo fmt --check
cargo clippy --all-targets -- -D warnings
cargo test
```

## 🚀 Instalação e Desinstalação

### Instalação no Registro (Word, Excel e PowerPoint):
```powershell
.\Instalar-OfficeTab.ps1
```

O script registra automaticamente:
- ProgID `OfficeTabs.Connect` em `HKCU:\Software\Classes`
- CLSID `{1c8ba3cd-5248-4611-8a6e-8ccbaf89b651}` apontando para `officetab_rs.dll` (`ThreadingModel=Apartment`)
- Chaves do suplemento COM em `HKCU:\Software\Microsoft\Office\<App>\Addins\OfficeTabs.Connect` com `LoadBehavior=3` (inicialização automática)

### Desinstalação:
```powershell
.\Desinstalar-OfficeTab.ps1
```

