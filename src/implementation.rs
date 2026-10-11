//! Implementação central do Office Tab Pro integrando COM, gerenciamento de abas e janelas.

use std::sync::Mutex;
use windows::Win32::Foundation::HWND;
use windows::core::Result;

use crate::tab_manager::{JanelaDocumento, TabManager};

/// Estrutura orquestradora do Office Tab
pub struct OfficeTab {
    tab_manager: Mutex<TabManager>,
}

impl OfficeTab {
    /// Cria uma nova instância do Office Tab
    pub fn new() -> Self {
        Self {
            tab_manager: Mutex::new(TabManager::new()),
        }
    }

    /// Inicializa os subsistemas necessários (COM, DPI awareness)
    pub fn initialize(&self) -> Result<()> {
        let _ = crate::utils::init_com();
        let _ = crate::utils::set_dpi_awareness();
        Ok(())
    }

    /// Varre e anexa o gerenciamento às janelas ativas do Microsoft Office
    pub fn attach_to_active_office(&self) -> Vec<JanelaDocumento> {
        let mut manager = self.tab_manager.lock().unwrap();
        let janelas = manager.discover_document_windows();
        for janela in &janelas {
            let _ = manager.monitor_window(janela.hwnd);
        }
        janelas
    }

    /// Anexa monitoramento a uma janela específica de documento
    pub fn handle_document_window(&self, hwnd: HWND) -> Result<()> {
        let mut manager = self.tab_manager.lock().unwrap();
        manager.monitor_window(hwnd)
    }

    /// Encerra as operações do add-in e remove todos os ganchos do sistema
    pub fn shutdown(&self) {
        let mut manager = self.tab_manager.lock().unwrap();
        manager.cleanup();
    }
}

impl Default for OfficeTab {
    fn default() -> Self {
        Self::new()
    }
}
