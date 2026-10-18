# Contributing to OfficeTab-rs 🤝🚀

We welcome all contributors with open arms! Whether you are an experienced Rust systems engineer, a Win32 / COM enthusiast, a UI designer, or someone who loves testing and reporting edge cases, your help is essential.

## 🚨 Precisamos da sua ajuda agora!
Versão 0.2.2 estável, mas precisamos de testes em Office 2013, melhorias em temas e internacionalização. Reporte bugs e envie PRs!

---

## 📚 Essential Documentation

Before you start, please check out the following guides:
- 📖 [**Developer Guide (Deep Dive)**](docs/DEVELOPER_GUIDE.md): Architecture, COM injection lifecycle, window subclassing, debugging inner loop, and roadmaps.
- 📐 [**Public API Reference**](docs/API.md): Structures, configuration serialization, and themes.
- 📦 [**Installers & Release Packages Guide**](GUIA_INSTALADORES.md): Details on the Inno Setup installer, 1-click batch scripts, and portable distribution.

---

## 🎨 Built with Vibe Coding

This codebase was developed using **Vibe Coding** — high-velocity, intent-driven AI pair programming combined with rigorous systems testing (51 passing unit tests, memory safety, zero clippy warnings).

We believe Vibe Coding empowers anyone to experiment, innovate, and contribute rapidly. Feel free to submit ideas, refactorings, or whole new features!

---

## 🎯 How You Can Help

Here are areas where community help is most needed:

1. **🧪 Real-World Host Testing**:
   - Test across multiple Office versions (Office 2016, 2019, 2021, Office 365, LTSC).
   - Test in different languages and regional Windows setups.

2. **📊 Other Office Applications**:
   - Word and Excel are heavily optimized.
   - We need contributors to test and fine-tune window docking for:
     - **PowerPoint** (`PPTFrameClass` / `PP9FrameClass`)
     - **Access** (`MSACCESS` / `OMain`)
     - **Visio** (`VISIO` / `VISIOA`)
     - **Project** (`PB`)
     - **Publisher** (`PUBLISHER`)

3. **🎨 Visuals & Theming**:
   - Windows 11 Mica / Acrylic backdrop integration.
   - Comprehensive Dark Mode presets.
   - Animations and custom tab icons.

4. **⚡ Event-Driven Performance**:
   - Replacing periodic UI timers with granular `WinEventHook` events (`EVENT_SYSTEM_FOREGROUND`, `EVENT_OBJECT_LOCATIONCHANGE`).
   - Per-Monitor DPI Aware v2 scaling improvements across multi-monitor setups.

5. **🌐 Internationalization (i18n)**:
   - Translating menus, tooltips, configuration utility, and documentation into more languages.

---

## 🛠️ Development & Quality Standards

Before submitting a Pull Request, please ensure all checks pass:

```powershell
# 1. Format code
cargo fmt -- --check

# 2. Run unit tests (51 tests)
cargo test --lib

# 3. Run linter with strict warnings
cargo clippy -- -D warnings

# 4. Build release binary
cargo build --release
```

### Key Engineering Rules:
- **Thread Safety**: All Win32 calls creating, moving, or destroying `HWND`s must run on the Office UI thread.
- **Clean Teardown**: Hooks, timers, subclasses, and GDI objects must be gracefully released upon shutdown or document closure.
- **Privacy First**: No telemetry, no network phoning home, no background data collection. Settings live cleanly in `HKCU\Software\OfficeTabRs`.
- **Open Source Integrity**: Clean-room implementation only. Do not copy or decompile proprietary binaries.

---

## 📬 Pull Request Process

1. **Fork and Branch**: Create a feature branch from `main` (e.g. `feature/powerpoint-fullscreen` or `fix/excel-high-dpi`).
2. **Commit clearly**: Write meaningful commit messages explaining the rationale behind your Win32/Rust changes.
3. **Verify locally**: Run `cargo test --lib` and test against real Office instances before opening your PR.
4. **Submit PR**: Open a Pull Request on GitHub describing your changes, any tests performed, and screenshots if UI changes are involved.

Thank you for helping make OfficeTab-rs amazing! ⭐
