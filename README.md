# OfficeTab-rs 📑✨

> **A modern, ultra-fast, native Tabbed Document Interface (TDI) COM Add-in for Microsoft Office, built with Rust and Vibe Coding.**

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Built with Rust](https://img.shields.io/badge/Built_with-Rust-orange.svg)](https://www.rust-lang.org/)
[![Vibe Coding](https://img.shields.io/badge/Methodology-Vibe_Coding-magenta.svg)](#-vibe-coding-story)
[![Platform](https://img.shields.io/badge/Platform-Windows_10%20%7C%2011-lightgrey.svg)](https://www.microsoft.com/windows)
[![Office Support](https://img.shields.io/badge/Office-2016%20%7C%202019%20%7C%202021%20%7C%20365-green.svg)](https://www.office.com/)

[**Português**](#-sobre-o-projeto-português) | [**English**](#-about-the-project-english) | [**Español**](#-acerca-del-proyecto-español) | [**Français**](#-à-propos-du-projet-français) | [**中文**](#-关于本项目-中文)

---

## 🎨 Vibe Coding Story

This project is an open-source, ground-up reinvention of document tab management for Microsoft Office, entirely built through **Vibe Coding** with AI assistance. It brings modern Windows Fluent Design, sub-millisecond tab switching, Single Window Mode, and memory safety (zero crashes, zero C++ pointer bugs) to your favorite productivity suite.

We are opening the code to the global developer and open-source community to make it the most reliable, elegant, and beloved tabbed add-in for Office worldwide!

---

## 🇧🇷 Sobre o Projeto (Português)

**OfficeTab-rs** é um suplemento COM nativo de código aberto desenvolvido em **Rust** que adiciona uma barra de abas moderna, elegante e de altíssimo desempenho ao **Microsoft Office (Word, Excel, PowerPoint, Access, Project, Visio e Publisher)**.

### ✨ Principais Recursos
- **Aparência Fluent & Moderna:** Abas com cantos arredondados, realce suave, gradientes dinâmicos e botão de fechar (`×`) com animação e hit-testing em tempo real.
- **Modo Janela Única (Single Window Mode):** Mantém apenas uma janela ativa visível na barra de tarefas por aplicativo, organizando todos os seus documentos abertos em abas.
- **Módulo Excel Otimizado:** Posicionamento dinâmico que respeita a Barra de Fórmulas (`fx`), o Name Box e atalhos essenciais como `Ctrl+T` (Criar Tabela).
- **Detecção Inteligente do Backstage:** Oculta a barra de abas automaticamente quando você acessa o menu *Arquivo / Abrir / Novo*.
- **Atalhos e Produtividade:**
  - `Ctrl + Tab` e `Ctrl + Shift + Tab` para alternar documentos.
  - `Ctrl + W` para fechar a aba ativa.
  - Clique duplo na barra ou clique no botão `+` para novo documento.
  - Clique com o botão do meio (Scroll) para fechar abas.
  - Menu de contexto moderno com botão direito (Copiar Caminho, Abrir Pasta, Fechar Outras).
- **100% Rust Seguro:** Construído com `windows-rs`, sem dependências pesadas de runtime, consumindo pouquíssima memória.

### 🤝 Como Contribuir
Este projeto foi concebido e acelerado com **Vibe Coding**! Precisamos da ajuda da comunidade global para:
- Testar e refinar o comportamento em diferentes versões do Office (2016, 2019, 2021, Microsoft 365, LTSC).
- Melhorar os módulos de **PowerPoint**, **Access**, **Project**, **Visio** e **Publisher**.
- Adicionar novos temas (Dark Mode total, Mica / Acrylic).
- Traduzir a interface e a documentação.
- Reportar bugs de abas em Excel SDI e melhorar a detecção de modo apresentação.

## 🐛 Status Atual v0.2.2
- ✅ Abas independentes por documento
- ✅ Fundo da barra igual ao menu Ribbon (sem rosa)
- ✅ Oculta abas em modo apresentação FullpageUIHost/Screen
- ✅ Criar nova aba com `+` sem abrir janela dupla
- ⚠️ Ajuda da comunidade: melhorar suporte a Office 2013 e testes em múltiplos idiomas

---

## 🇬🇧 About the Project (English)

**OfficeTab-rs** is an open-source native COM Add-in written in **Rust** that brings high-performance, modern tabs to **Microsoft Office (Word, Excel, PowerPoint, Access, Project, Visio, and Publisher)**.

### ✨ Key Features
- **Modern Fluent Aesthetics:** Sleek rounded tabs, soft accent pill badges on hover, and smooth interactive close (`×`) buttons.
- **Single Window Mode:** Consolidates multiple open documents into a single taskbar button with tabs.
- **Dedicated Excel Engine:** Perfectly anchored below the Formula Bar (`fx`) and Name Box without covering formulas, while preserving native shortcuts like `Ctrl+T` (Create Table).
- **Backstage Awareness:** Automatically hides the tab bar when opening the backstage *File / Home* view.
- **Productivity & Shortcuts:**
  - `Ctrl + Tab` / `Ctrl + Shift + Tab` to cycle through tabs.
  - `Ctrl + W` to close tabs.
  - Double-click or click `+` to open new documents.
  - Middle-click to close tabs.
  - Full-featured Right-Click Context Menu.
- **100% Memory-Safe Rust:** Zero C++ segfaults, zero memory leaks, lightning-fast native Win32 performance.

### 🤝 Community Help Needed!
Built with **Vibe Coding**, we welcome contributions from all developers!
- Testing across multiple Office versions and Windows builds.
- Refinements for PowerPoint, Visio, Access, and Project.
- New themes, dark mode enhancements, and UI polish.

---

## 🇪🇸 Acerca del Proyecto (Español)

**OfficeTab-rs** es un complemento COM nativo y de código abierto desarrollado en **Rust** que agrega pestañas modernas y de alto rendimiento a **Microsoft Office (Word, Excel, PowerPoint, Access, Project, Visio y Publisher)**.

### ✨ Características Principales
- **Diseño Moderno Fluent:** Pestañas redondeadas con resaltado sutil y botón de cierre (`×`) interactivo.
- **Modo de Ventana Única:** Unifica todos los documentos en un solo botón en la barra de tareas de Windows.
- **Motor Excel Optimizado:** Posicionado exactamente debajo de la Barra de Fórmulas (`fx`), respetando atajos nativos como `Ctrl+T`.
- **Detección de Backstage:** Oculta las pestañas cuando estás en la pantalla de *Archivo / Nuevo*.
- **Atajos Intuitivos:** `Ctrl+Tab`, `Ctrl+W`, clic central para cerrar y menú contextual completo.

---

## 🇫🇷 À propos du projet (Français)

**OfficeTab-rs** est un complément COM natif open-source écrit en **Rust** qui apporte des onglets modernes, élégants et ultra-rapides à **Microsoft Office (Word, Excel, PowerPoint, Access, Project, Visio et Publisher)**.

### ✨ Caractéristiques
- **Design Moderne Fluent:** Onglets soignés avec boutons de fermeture interactifs et intégration fluide.
- **Mode Fenêtre Unique:** Regroupe tous vos documents ouverts sous une seule icône de barre des tâches.
- **Intégration Excel Parfaite:** Positionné sous la barre de formule (`fx`) sans chevauchement.
- **Fiabilité Rust:** Zéro fuite de mémoire, sécurité maximale et performances natives Win32.

---

## 🇨🇳 关于本项目 (中文)

**OfficeTab-rs** 是一个使用 **Rust** 编写的开源 Microsoft Office 原生 COM 标签页插件，为 **Word、Excel、PowerPoint、Access、Project、Visio 和 Publisher** 提供极致流畅的现代化多标签页浏览体验。

### ✨ 主要特点
- **现代 Fluent 设计风格：** 优雅的圆角标签、悬停微光动效与实时交互的关闭按钮。
- **单窗口模式（Single Window Mode）：** 任务栏仅保留一个主窗口图标，所有文档统一在标签页中管理。
- **针对 Excel 深度优化：** 动态贴合在编辑公式栏（`fx`）与名称框下方，绝不遮挡公式，并完整保留 `Ctrl+T` 原生创建表格快捷键。
- **100% Rust 内存安全：** 原生 Win32 API 交互，无任何 C++ 内存泄漏风险，秒速响应。
- **Vibe Coding 理念：** 本项目完全通过 AI 协同的 Vibe Coding 理念打造，诚邀全球开发者共同完善！

---

## 🚀 Como Instalar / Installation Options

Você pode instalar o **OfficeTab-rs** de várias maneiras, desde um instalador visual executável até compilação direta do código fonte:

### 📦 Opção 1: Instalador Gráfico Windows (Recomendado para Usuários)
Baixe e execute o instalador oficial gerado com Inno Setup:
- **Download Direto:** [**`installers/OfficeTab-rs-Setup-v0.2.2.exe`**](./installers/OfficeTab-rs-Setup-v0.2.2.exe) (2.85 MB)
- Suporte a Português do Brasil e Inglês
- Instala a DLL COM x64 e o Painel de Configurações (`officetab-config.exe`)
- Registra automaticamente o suplemento nos 7 aplicativos do Microsoft Office (Word, Excel, PowerPoint, Access, Project, Visio e Publisher)
- Cria atalhos no Menu Iniciar e opção de desinstalação no Painel de Controle / Configurações do Windows.

### ⚡ Opção 2: Instalador em 1-Clique (`Instalar.bat`)
Basta dar duplo clique no arquivo:
```cmd
Instalar.bat
```
Ele executa a ativação e registro em segundo plano sem necessidade de abrir terminais. Para remover: `Desinstalar.bat`.

### 🧰 Opção 3: Pacote Portátil (`Portable.zip`)
Ideal para quem não deseja executar instaladores ou para ambientes corporativos:
- **Download Direto:** [**`installers/OfficeTab-rs-v0.2.2-Portable.zip`**](./installers/OfficeTab-rs-v0.2.2-Portable.zip) (1.66 MB)
- Descompacte em qualquer pasta e execute `Instalar.bat` ou `Instalar-OfficeTab.ps1`.

### 💻 Opção 4: Script PowerShell Automatizado
Execute diretamente no terminal:
```powershell
.\Instalar-OfficeTab.ps1
```
Para diagnosticar o estado do add-in e verificar integridade:
```powershell
.\Testar-OfficeTab.ps1
```
Para desinstalar:
```powershell
.\Desinstalar-OfficeTab.ps1
```

### 🛠️ Opção 5: Compilação a partir do Código Fonte (Desenvolvedores)
#### Pré-requisitos
- Windows 10 ou 11 (x64)
- Microsoft Office (2016, 2019, 2021, ou Microsoft 365)
- [Rust toolchain com MSVC](https://rustup.rs/)

```powershell
# Clonar repositório
git clone https://github.com/washingtonwdc/officetab-rs.git
cd officetab-rs

# Executar suíte de testes (51 testes unitários)
cargo test --lib
cargo clippy -- -D warnings

# Compilar binários de produção
cargo build --release

# Instalar e registrar
.\Instalar.bat
```

---

## 🛠️ Architecture & Technology

```
officetab-rs/
├── src/
│   ├── lib.rs                 # COM inproc server exports (DllGetClassObject, DllRegisterServer)
│   ├── com_addin_simple.rs    # Pure COM IDTExtensibility2 and IClassFactory implementation
│   ├── tab_manager.rs         # Win32 lifecycle, window hierarchy, Single Window Mode, Excel fx detection
│   ├── tab_control.rs         # Native SysTabControl32 subclassing, Fluent custom draw, tooltips
│   ├── theme.rs               # Color palettes, Fluent presets, DPI-aware typography
│   ├── persistence.rs         # Registry & JSON configuration management
│   ├── subclass.rs            # Win32 ComCtl32 SetWindowSubclass management
│   └── utils.rs               # DPI utilities, file path resolver, clipboard integration
├── Instalar-OfficeTab.ps1     # Automated installation script
├── Desinstalar-OfficeTab.ps1   # Clean uninstallation script
├── Testar-OfficeTab.ps1       # Automated diagnostics & health check
└── Cargo.toml                 # Dependencies & Rust 2024 edition config
```

---

## 📜 License & Open Source

This project is open-source under the **[Apache License 2.0](LICENSE)**.
It is an entirely independent, clean-room implementation with no proprietary code from commercial tabbed software.

---

## 🌟 Join the Community!

Help us make Office tab management open, beautiful, and blazingly fast.
Give this repository a ⭐, open an Issue, or send a Pull Request!
