// Implementação central do Office Tab Pro integrando COM, gerenciamento de abas e janelas.
//
// Copyright 2024 The OfficeTabRs Authors
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//!
//! Implementação central do Office Tab Pro integrando COM, gerenciamento de abas e janelas.

use windows::Win32::Foundation::HWND;
use windows::core::Result;

use crate::tab_manager::JanelaDocumento;

/// Estrutura orquestradora do Office Tab
pub struct OfficeTab;

impl OfficeTab {
    /// Cria uma nova instância do Office Tab
    pub fn new() -> Self {
        Self
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
            .map(|(hwnd, titulo)| {
                let classe = crate::tab_manager::extrair_classe_janela(hwnd);
                JanelaDocumento {
                    hwnd,
                    titulo,
                    classe,
                }
            })
            .collect()
    }

    /// Anexa monitoramento a uma janela específica de documento
    pub fn handle_document_window(&self, hwnd: HWND) -> Result<()> {
        unsafe {
            let _ = windows::Win32::UI::WindowsAndMessaging::PostMessageW(
                hwnd,
                crate::tab_manager::WM_OFFICETAB_INICIALIZAR,
                windows::Win32::Foundation::WPARAM(0),
                windows::Win32::Foundation::LPARAM(0),
            );
        }
        Ok(())
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
