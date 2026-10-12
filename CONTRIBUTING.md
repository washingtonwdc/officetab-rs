# Contributing to OfficeTab-rs 🤝🚀

We welcome all contributors with open arms! Whether you are an experienced Rust systems engineer, a Win32 / COM enthusiast, a UI designer, or someone who loves testing and reporting edge cases, your help is essential.

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
     - **Access** (`MSACCESS`)
     - **Visio** (`VISIO` / `VISIOA`)
     - **Project** (`PB`)
     - **Publisher** (`PUBLISHER`)

3. **🎨 Visuals & Theming**:
   - Windows 11 Mica / Acrylic backdrop integration.
   - Comprehensive Dark Mode presets.
   - Animations and custom tab icons.

4. **🌐 Internationalization (i18n)**:
   - Translating menus, tooltips, configuration utility, and documentation into more languages.

---

## 🛠️ Development & Quality Standards

Before submitting a Pull Request, please ensure all checks pass:

```powershell
# Format code
cargo fmt -- --check

# Run unit tests (51 tests)
cargo test --lib

# Run linter
cargo clippy -- -D warnings

# Build release DLL
cargo build --release
```

### Key Engineering Rules:
- **Thread Safety**: All Win32 calls creating, moving, or destroying `HWND`s must run on the Office UI thread.
- **Clean Teardown**: Hooks, timers, subclasses, and GDI objects must be gracefully released upon shutdown or document closure.
- **Privacy First**: No telemetry, no network phoning home, no background data collection. Settings live cleanly in `HKCU\Software\OfficeTabRs`.
- **Open Source Integrity**: Clean-room implementation only. Do not copy or decompile proprietary binaries.

---

Thank you for helping make OfficeTab-rs amazing! ⭐
