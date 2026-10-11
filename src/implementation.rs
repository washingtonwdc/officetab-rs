//! Implementação central do Office Tab Pro integrando COM, gerenciamento de abas e janelas.

use windows::Win32::Foundation::HWND;
use windows::core::Result;

use crate::tab_manager::{JanelaDocumento, TabManager};

/// Estrutura orquestradora do Office Tab
pub struct OfficeTab {
    _tab_manager: TabManager,
}

impl OfficeTab {
    /// Cria uma nova instância do Office Tab
    pub fn new() -> Self {
        Self {
            _tab_manager: TabManager::new(),
        }
    }

    /// Inicializa os subsistemas necessários (COM, DPI awareness)
    pub fn initialize(&self) -> Result<()> {
        let _ = crate::utils::init_com();
        let _ = crate::utils::set_dpi_awareness();
        crate::tab_manager::obter_ou_criar_gerenciador();
        Ok(())
    }

    /// Varre e anexa o gerenciamento às janelas ativas do Microsoft Office
    pub fn attach_to_active_office(&self) -> Vec<JanelaDocumento> {
        crate::tab_manager::varrer_e_anexar_janelas();
        crate::tab_manager::listar_janelas_processo()
            .into_iter()
            .map(|(hwnd, titulo)| JanelaDocumento {
                hwnd,
                titulo,
                classe: "OpusApp".to_string(),
            })
            .collect()
    }

    /// Anexa monitoramento a uma janela específica de documento
    pub fn handle_document_window(&self, hwnd: HWND) -> Result<()> {
        crate::tab_manager::varrer_e_anexar_janelas();
        let mut manager = TabManager::new();
        manager.monitor_window(hwnd)
    }

    /// Encerra as operações do add-in e remove todos os ganchos do sistema
    pub fn shutdown(&self) {
        crate::tab_manager::desativar_gerenciador_global();
    }
}

impl Default for OfficeTab {
    fn default() -> Self {
        Self::new()
    }
}
