//! Gerenciador central de abas do Office (inspirado no TOTWindowsManager).

use std::collections::HashMap;
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use windows::Win32::Foundation::{BOOL, HWND, LPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetClassNameW, GetWindowTextW, IsWindowVisible, SetParent,
};
use windows::core::Result;

use crate::subclass::SubclassManager;
use crate::tab_control::TabControl;

/// Estrutura para metadados de uma janela de documento do Office
#[derive(Debug, Clone)]
pub struct JanelaDocumento {
    pub hwnd: HWND,
    pub titulo: String,
    pub classe: String,
}

/// Gerenciador central de abas e janelas do Office
pub struct TabManager {
    subclass_manager: SubclassManager,
    janelas_ativas: HashMap<isize, JanelaDocumento>,
    tab_controls: HashMap<isize, TabControl>,
}

impl TabManager {
    pub fn new() -> Self {
        Self {
            subclass_manager: SubclassManager::new(),
            janelas_ativas: HashMap::new(),
            tab_controls: HashMap::new(),
        }
    }

    /// Descobre janelas visíveis das aplicações do Microsoft Office (Word, Excel, PowerPoint).
    pub fn discover_document_windows(&mut self) -> Vec<JanelaDocumento> {
        let mut janelas: Vec<JanelaDocumento> = Vec::new();
        unsafe {
            let _ = EnumWindows(
                Some(enum_windows_proc),
                LPARAM(&mut janelas as *mut _ as isize),
            );
        }

        for doc in &janelas {
            self.janelas_ativas.insert(doc.hwnd.0 as isize, doc.clone());
        }

        janelas
    }

    /// Aplica o subclassing a uma janela do Office e a registra.
    pub fn monitor_window(&mut self, hwnd: HWND) -> Result<()> {
        self.subclass_manager.subclass_window(hwnd)?;
        let titulo = extrair_texto_janela(hwnd);
        let classe = extrair_classe_janela(hwnd);

        self.janelas_ativas.insert(
            hwnd.0 as isize,
            JanelaDocumento {
                hwnd,
                titulo,
                classe,
            },
        );

        Ok(())
    }

    /// Reparenta uma janela de documento para um controle contêiner de abas.
    pub fn reparent_window(&self, hwnd: HWND, parent_hwnd: HWND) -> Result<()> {
        unsafe {
            let _ = SetParent(hwnd, parent_hwnd)?;
        }
        Ok(())
    }

    /// Vincula uma barra de abas à janela hospedeira.
    pub fn attach_tab_bar(&mut self, parent_hwnd: HWND, id: u32) -> Result<()> {
        let tab_control = TabControl::new(parent_hwnd, id)?;
        self.tab_controls
            .insert(parent_hwnd.0 as isize, tab_control);
        Ok(())
    }

    /// Retorna as janelas atualmente ativas e monitoradas.
    pub fn get_active_windows(&self) -> Vec<JanelaDocumento> {
        self.janelas_ativas.values().cloned().collect()
    }

    /// Limpa todos os hooks e referências ativas.
    pub fn cleanup(&mut self) {
        self.subclass_manager.unsubclass_all();
        self.tab_controls.clear();
        self.janelas_ativas.clear();
    }
}

impl Drop for TabManager {
    fn drop(&mut self) {
        self.cleanup();
    }
}

impl Default for TabManager {
    fn default() -> Self {
        Self::new()
    }
}

/// Extrai o nome da classe de uma janela Win32.
fn extrair_classe_janela(hwnd: HWND) -> String {
    let mut buffer = [0u16; 256];
    let len = unsafe { GetClassNameW(hwnd, &mut buffer) };
    if len > 0 {
        OsString::from_wide(&buffer[..len as usize])
            .to_string_lossy()
            .to_string()
    } else {
        String::new()
    }
}

/// Extrai o título textual de uma janela Win32.
fn extrair_texto_janela(hwnd: HWND) -> String {
    let mut buffer = [0u16; 512];
    let len = unsafe { GetWindowTextW(hwnd, &mut buffer) };
    if len > 0 {
        OsString::from_wide(&buffer[..len as usize])
            .to_string_lossy()
            .to_string()
    } else {
        String::new()
    }
}

/// Callback do EnumWindows para localizar janelas do Microsoft Office.
unsafe extern "system" fn enum_windows_proc(hwnd: HWND, lparam: LPARAM) -> BOOL {
    unsafe {
        if !IsWindowVisible(hwnd).as_bool() {
            return BOOL(1);
        }

        let class_name = extrair_classe_janela(hwnd);

        // OpusApp (Word), XLMAIN / EXCEL7 (Excel), PPTFrameClass (PowerPoint)
        let is_office = class_name.contains("OpusApp")
            || class_name.contains("XLMAIN")
            || class_name.contains("EXCEL7")
            || class_name.contains("PPTFrameClass");

        if is_office {
            let windows = &mut *(lparam.0 as *mut Vec<JanelaDocumento>);
            let titulo = extrair_texto_janela(hwnd);
            windows.push(JanelaDocumento {
                hwnd,
                titulo,
                classe: class_name,
            });
        }
    }

    BOOL(1)
}
