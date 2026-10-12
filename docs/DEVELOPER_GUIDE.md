# Guia Completo do Desenvolvedor — OfficeTab-rs 📑🛠️

Bem-vindo ao guia de desenvolvimento do **OfficeTab-rs**! Este documento reúne todas as informações técnicas, de arquitetura, ferramentas e procedimentos práticos para que desenvolvedores possam entender a base de código, configurar seu ambiente e contribuir ativamente para o projeto.

---

## 📌 Sumário

1. [Visão Geral & Arquitetura Técnica](#1-visão-geral--arquitetura-técnica)
2. [Configuração do Ambiente de Desenvolvimento](#2-configuração-do-ambiente-de-desenvolvimento)
3. [Estrutura da Base de Código (`src/`)](#3-estrutura-da-base-de-código-src)
4. [Ciclo de Desenvolvimento & Debugging](#4-ciclo-de-desenvolvimento--debugging)
5. [Como o Office Carrega e Gerencia o Add-in](#5-como-o-office-carrega-e-gerencia-o-add-in)
6. [Resolução de Problemas & Resiliência do Office](#6-resolução-de-problemas--resiliência-do-office)
7. [Guia de Testes Automatizados e com Frida](#7-guia-de-testes-automatizados-e-com-frida)
8. [Roadmap de Contribuições & Áreas Prioritárias](#8-roadmap-de-contribuições--áreas-prioritárias)
9. [Regras de Ouro de Engenharia (Boas Práticas)](#9-regras-de-ouro-de-engenharia)

---

## 1. Visão Geral & Arquitetura Técnica

O OfficeTab-rs é uma implementação *clean-room* em **Rust** que transforma a experiência do Microsoft Office (Word, Excel, PowerPoint, Access, Project, Visio e Publisher), substituindo as múltiplas janelas independentes (SDI - Single Document Interface) por uma interface elegante e moderna orientada a abas (TDI - Tabbed Document Interface).

### 🧩 Mecanismo Central: Como funciona por baixo dos panos?

```
┌────────────────────────────────────────────────────────┐
│                   Microsoft Office                     │
│         (Word, Excel, PowerPoint, etc.)                │
└──────────────────────────┬─────────────────────────────┘
                           │ 1. Inicia processo e lê HKCU Addins
                           ▼
┌────────────────────────────────────────────────────────┐
│            officetab_rs.dll (InprocServer32)           │
│  • DllGetClassObject                                   │
│  • IClassFactory::CreateInstance                       │
│  • IDTExtensibility2::OnConnection                     │
└──────────────────────────┬─────────────────────────────┘
                           │ 2. Inicializa Hooks e Gerenciador
                           ▼
┌────────────────────────────────────────────────────────┐
│              TOTWindowsManager (Rust)                  │
│  • WinEventHook (EVENT_OBJECT_CREATE / SHOW)           │
│  • SetTimer UI (Fallback periódico)                    │
│  • Enumera janelas do processo atual                   │
└──────────────┬───────────────────────────┬─────────────┘
               │                           │
               ▼                           ▼
┌───────────────────────────┐ ┌───────────────────────────┐
│     Window Subclassing    │ │   SysTabControl32 (GDI)   │
│ • SetWindowSubclass       │ │ • Barra de abas nativa    │
│ • WM_WINDOWPOSCHANGED     │ │ • Custom Draw Fluent      │
│ • WM_ACTIVATE / WM_CLOSE  │ │ • Botão de fechar (x)     │
│ • Single Window Mode      │ │ • Atalhos Ctrl+Tab/Ctrl+W │
└───────────────────────────┘ └───────────────────────────┘
```

1. **Injeção COM Nativa:** O processo do Office lê o Registro do Windows em `HKCU\Software\Microsoft\Office\<App>\Addins\OfficeTabRs.Connect` (`LoadBehavior = 3`), carrega a DLL e invoca a interface COM `IDTExtensibility2`.
2. **Descoberta de Janelas:** Ao se conectar, o gerenciador monitora a criação de novas janelas através de `SetWinEventHook` com fallback por timer na thread de interface.
3. **Subclassing com ComCtl32:** A janela principal do documento recebe um gancho via `SetWindowSubclass`, permitindo interceptar mensagens de redimensionamento (`WM_WINDOWPOSCHANGED`), ativação, fechamento e teclado.
4. **Acoplamento da Barra de Abas:** Um controle nativo `SysTabControl32` é criado dinamicamente como filho da janela principal ou ancorado estrategicamente abaixo da Ribbon (e abaixo da Barra de Fórmulas no Excel).
5. **Modo Janela Única (Single Window Mode):** Mantém apenas a janela do documento ativo visível na barra de tarefas do Windows (`SW_SHOW`), enquanto os outros documentos permanecem associados às abas abertas.

---

## 2. Configuração do Ambiente de Desenvolvimento

### Requisitos de Sistema
- **Sistema Operacional:** Windows 10 ou Windows 11 (64-bit).
- **Rust Toolchain:** Rust 1.85+ com o target `x86_64-pc-windows-msvc`.
  ```powershell
  rustup default stable-x86_64-pc-windows-msvc
  ```
- **Build Tools:** Visual Studio C++ Build Tools (inclui o compilador MSVC e o Windows SDK).
- **Inno Setup 6:** (Opcional, para gerar o instalador visual executável).
  - Pode ser instalado via winget: `winget install JRSoftware.InnoSetup`

### Ferramentas Recomendadas para Inspecionar Win32 / COM
- **Inspect.exe / Spy++:** Para inspecionar a hierarquia de `HWND`, nomes de classes de janelas e mensagens Win32.
- **Process Explorer (Sysinternals):** Para visualizar DLLs carregadas dentro de `WINWORD.EXE` ou `EXCEL.EXE`.
- **Frida:** Para tracing dinâmico de chamadas Win32 / COM sem necessidade de recompilar.

---

## 3. Estrutura da Base de Código (`src/`)

| Arquivo / Módulo | Linhas | Responsabilidade Técnica |
| :--- | :--- | :--- |
| [`src/lib.rs`](file:///C:/Users/washi/officetab-rs/src/lib.rs) | ~220 | Ponto de entrada do DLL (`DllGetClassObject`, `DllCanUnloadNow`, `DllRegisterServer`, `DllUnregisterServer`, `DllInstall`). |
| [`src/com_addin_simple.rs`](file:///C:/Users/washi/officetab-rs/src/com_addin_simple.rs) | ~180 | Implementação das interfaces COM `IClassFactory` e `IDTExtensibility2` com vtables puras em Rust. |
| [`src/implementation.rs`](file:///C:/Users/washi/officetab-rs/src/implementation.rs) | ~200 | Orquestrador principal da extensão, gerenciando o ciclo de vida inicialização/término. |
| [`src/tab_manager.rs`](file:///C:/Users/washi/officetab-rs/src/tab_manager.rs) | ~2500 | **Coração do projeto (`TOTWindowsManager`)**: detecção dos 7 apps, cálculo de alturas, WinEventHook, Single Window Mode, atalhos de teclado e menu de contexto. |
| [`src/tab_control.rs`](file:///C:/Users/washi/officetab-rs/src/tab_control.rs) | ~1100 | Envolvimento de `SysTabControl32`, renderização GDI Custom Draw, cantos arredondados, hit-testing de abas e botão de fechar (`×`). |
| [`src/subclass.rs`](file:///C:/Users/washi/officetab-rs/src/subclass.rs) | ~180 | Gerenciador de subclassing de janelas usando `ComCtl32.dll` (`SetWindowSubclass`, `DefSubclassProc`). |
| [`src/theme.rs`](file:///C:/Users/washi/officetab-rs/src/theme.rs) | ~800 | Sistema completo de temas e paletas de cores (Presets Light, Dark, Word Blue, Excel Green, etc.) com conversões GDI `COLORREF`. |
| [`src/persistence.rs`](file:///C:/Users/washi/officetab-rs/src/persistence.rs) | ~700 | Gerenciador de configurações em `HKCU\Software\OfficeTabRs` e exportação/importação de configurações e favoritos via JSON. |
| [`src/registry_simple.rs`](file:///C:/Users/washi/officetab-rs/src/registry_simple.rs) | ~220 | Manipulação das chaves do Registro do Windows para registro COM e ativação nos 7 aplicativos do Office. |
| [`src/license.rs`](file:///C:/Users/washi/officetab-rs/src/license.rs) | ~500 | Validador offline de licenças e trial (100% local, sem telemetria nem conexões externas). |
| [`src/utils.rs`](file:///C:/Users/washi/officetab-rs/src/utils.rs) | ~300 | Utilitários Win32, manipulação de texto Unicode UTF-16, Clipboard e cálculo de DPI. |
| [`src/config_main.rs`](file:///C:/Users/washi/officetab-rs/src/config_main.rs) | ~400 | Interface gráfica nativa do utilitário `officetab-config.exe`. |

---

## 4. Ciclo de Desenvolvimento & Debugging

### Fluxo de Trabalho (Inner Loop)

```powershell
# 1. Certifique-se de que nenhum aplicativo Office esteja aberto
Get-Process winword, excel, powerpnt -ErrorAction SilentlyContinue | Stop-Process -Force

# 2. Formatar e validar código com clippy
cargo fmt
cargo clippy -- -D warnings

# 3. Rodar suíte de testes unitários (51 testes)
cargo test --lib

# 4. Compilar DLL em release
cargo build --release

# 5. Registrar o add-in no Office do seu usuário
.\Instalar.bat

# 6. Abrir o Word ou Excel para testar ao vivo
start winword
```

### Inspecionando Logs em Tempo Real
O OfficeTab-rs grava informações de diagnóstico e eventos no arquivo de log:
- **Local do log:** `C:\Users\<Usuario>\officetab-rs\officetab.log` (ou no diretório configurado via registro).
- Para monitorar o log em tempo real enquanto interage com o Office:
  ```powershell
  Get-Content -Path .\officetab.log -Wait -Tail 30
  ```

---

## 5. Como o Office Carrega e Gerencia o Add-in

O Office utiliza o padrão de registro COM por usuário (`HKCU`):

```
HKCU\Software\Classes\OfficeTabRs.Connect\CLSID
└── (Default) = {9a1929ef-1495-413e-993c-326dd15e35df}

HKCU\Software\Classes\CLSID\{9a1929ef-1495-413e-993c-326dd15e35df}\InprocServer32
├── (Default) = "C:\caminho\para\officetab_rs.dll"
└── ThreadingModel = "Apartment"

HKCU\Software\Microsoft\Office\<App>\Addins\OfficeTabRs.Connect
├── FriendlyName = "Office Tab Rust"
├── Description = "Office Tab Rust"
├── LoadBehavior = 3 (DWORD: carregar automaticamente ao iniciar)
└── CommandLineSafe = 0
```

### Classes de Janelas dos Aplicativos Suportados
Ao monitorar janelas com `GetClassNameW`, o OfficeTab-rs identifica o tipo de host:

| Aplicativo | Classe da Janela Principal | Janela Filha / Documento |
| :--- | :--- | :--- |
| **Word** | `OpusApp` | `_WwG` (Área de edição) |
| **Excel** | `XLMAIN` | `EXCEL7` (Pasta de trabalho) + `EXCEL;` (Formula Bar) |
| **PowerPoint**| `PPTFrameClass` / `PP9FrameClass` | `MDIClient` / `paneClassDC` |
| **Access** | `OMain` / `MSACCESS` | `MDIClient` |
| **Project** | `PB` | MDI / Subjanelas |
| **Visio** | `VISIO` / `VISIOA` | Subjanelas de diagrama |
| **Publisher** | `PUBLISHER` | Subjanelas de publicação |

---

## 6. Resolução de Problemas & Resiliência do Office

### O suplemento parou de carregar? (Resiliency / DisabledItems)
Se o Office fechar de maneira inesperada enquanto uma DLL estiver carregada, ele pode adicionar o suplemento à sua lista de resiliência:
`HKCU\Software\Microsoft\Office\<Versao>\<App>\Resiliency\DisabledItems`

**Solução Automática:**
Execute o script de diagnóstico ou o instalador:
```powershell
.\Testar.bat
# ou:
.\Instalar.bat
```
O script [`Instalar-OfficeTab.ps1`](file:///C:/Users/washi/officetab-rs/Instalar-OfficeTab.ps1) já detecta e limpa automaticamente qualquer entrada relacionada ao OfficeTab dessa chave.

---

## 7. Guia de Testes Automatizados e com Frida

### Testes Unitários Nativos
O projeto conta com **51 testes unitários** cobrindo:
- Verificação de ponteiros e chamadas COM (`DllGetClassObject`, `IClassFactory`, `QueryInterface`, `OnConnection`).
- Detecção e limpeza de nomes de documentos para todos os 7 aplicativos.
- Conversão e renderização de cores GDI, temas e presets.
- Serialização e migração de configurações JSON e favoritos.
- Cálculo de escalonamento por DPI e alinhamento de abas.

Para executar todos os testes:
```powershell
cargo test --lib
```

### Testes com Frida (Dynamic Hooking)
Scripts para interceptação dinâmica estão disponíveis em `default-workspace/frida_runtime/`:
```powershell
cd C:\Users\washi\OneDrive\Documentos\deepseek-harness\default-workspace\frida_runtime
.\EXECUTE_NOW.ps1 -App excel
```

---

## 8. Roadmap de Contribuições & Áreas Prioritárias

Estamos abertos a contribuições em qualquer uma das áreas abaixo:

### 🎯 Prioridade 1: Refinamento dos Módulos dos Aplicativos Secundários
- [ ] **PowerPoint (`PPTFrameClass`):** Refinar o reposicionamento da barra ao alternar entre o modo Edição e Apresentação de Slides.
- [ ] **Access (`OMain`):** Ajustar o suporte quando o usuário alterna entre interface com abas nativa do Access e janelas flutuantes.
- [ ] **Visio & Project:** Validação de DPI e posicionamento em monitores ultra-wide.

### 🎨 Prioridade 2: Interface Visual & Temas
- [ ] Adicionar suporte a **Mica / Acrylic** do Windows 11 no background da barra de abas.
- [ ] Presets adicionais de Dark Mode adaptativo (compatível com os temas Preto e Cinza Escuro do Office 365).
- [ ] Ícones dinâmicos ao lado do título da aba baseados no tipo de arquivo (`.docx`, `.xlsx`, `.pptx`).

### ⚡ Prioridade 3: Performance e Arquitetura
- [ ] Eliminar completamente o polling periódico de timer em favor de `SetWinEventHook` com escuta precisa de `EVENT_SYSTEM_FOREGROUND` e `EVENT_OBJECT_LOCATIONCHANGE`.
- [ ] Suporte a múltiplos monitores com DPIs mistos (Per-Monitor DPI Aware v2).

---

## 9. Regras de Ouro de Engenharia

Ao contribuir com código para o OfficeTab-rs, siga rigorosamente estas diretrizes:

1. **Thread Safety de UI:** Nunca invoque funções que criem, redimensionem ou destruam `HWND`s fora da thread principal de UI do Office. Toda manipulação de janela deve ocorrer no contexto da mensagem tratada ou postada via `PostMessageW`.
2. **Clean Teardown:** Sempre limpe subclasses (`RemoveWindowSubclass`), feche timers (`KillTimer`), libere pincéis GDI (`DeleteObject`) e desmarque hooks no encerramento.
3. **Privacidade Total:** Zero telemetria. O OfficeTab-rs não realiza conexões de rede, não consulta servidores remotos e armazena suas configurações localmente.
4. **Clean-Room Implementation:** Todo código deve ser autoral, original e licenciado sob Apache-2.0. Nunca incorpore código descompilado de softwares proprietários.

---

> Tem dúvidas ou quer discutir uma ideia de arquitetura? Abra uma [Issue](https://github.com/washingtonwdc/officetab-rs/issues) ou envie um Pull Request! 🚀
