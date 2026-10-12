// Subclassing de janelas do Office via Windows Comctl32 / Shell API.
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
//! Subclassing de janelas do Office via Windows Comctl32 / Shell API.

use std::collections::HashSet;
use std::sync::Mutex;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::UI::Shell::{DefSubclassProc, RemoveWindowSubclass, SetWindowSubclass};
use windows::core::Result;

const OFFICETAB_SUBCLASS_ID: usize = 0x4F544142; // 'OTAB'

static JANELAS_SUBCLASSED: Mutex<Option<HashSet<isize>>> = Mutex::new(None);

/// Gerenciador de subclassing de janelas.
pub struct SubclassManager;

impl SubclassManager {
    pub fn new() -> Self {
        let mut guard = JANELAS_SUBCLASSED.lock().unwrap();
        if guard.is_none() {
            *guard = Some(HashSet::new());
        }
        Self
    }

    /// Aplica o subclassing à janela especificada.
    pub fn subclass_window(&self, hwnd: HWND) -> Result<()> {
        unsafe {
            let sucesso =
                SetWindowSubclass(hwnd, Some(procedimento_subclass), OFFICETAB_SUBCLASS_ID, 0);

            if !sucesso.as_bool() {
                return Err(windows::core::Error::from_win32());
            }

            if let Ok(mut guard) = JANELAS_SUBCLASSED.lock()
                && let Some(ref mut set) = *guard
            {
                set.insert(hwnd.0 as isize);
            }

            Ok(())
        }
    }

    /// Remove o subclassing da janela especificada.
    pub fn unsubclass_window(&self, hwnd: HWND) -> Result<()> {
        unsafe {
            let _ = RemoveWindowSubclass(hwnd, Some(procedimento_subclass), OFFICETAB_SUBCLASS_ID);

            if let Ok(mut guard) = JANELAS_SUBCLASSED.lock()
                && let Some(ref mut set) = *guard
            {
                set.remove(&(hwnd.0 as isize));
            }

            Ok(())
        }
    }

    /// Remove o subclassing de todas as janelas ativas rastreadas.
    pub fn unsubclass_all(&self) {
        if let Ok(mut guard) = JANELAS_SUBCLASSED.lock()
            && let Some(ref mut set) = *guard
        {
            for &raw_hwnd in set.iter() {
                let hwnd = HWND(raw_hwnd as *mut _);
                unsafe {
                    let _ = RemoveWindowSubclass(
                        hwnd,
                        Some(procedimento_subclass),
                        OFFICETAB_SUBCLASS_ID,
                    );
                }
            }
            set.clear();
        }
    }
}

impl Default for SubclassManager {
    fn default() -> Self {
        Self::new()
    }
}

/// Procedimento de janela subclassed.
/// Todas as mensagens Win32 chegam aqui NA THREAD DE UI DO OFFICE.
/// Blindado com std::panic::catch_unwind para evitar que qualquer pânico cause desativação do suplemento pelo Office.
unsafe extern "system" fn procedimento_subclass(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    uid_subclass: usize,
    ref_data: usize,
) -> LRESULT {
    let resultado = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        procedimento_subclass_interno(hwnd, msg, wparam, lparam, uid_subclass, ref_data)
    }));

    match resultado {
        Ok(res) => res,
        Err(e) => {
            crate::tab_manager::log_debug(&format!(
                "Pânico capturado no procedimento_subclass: {:?}",
                e
            ));
            unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) }
        }
    }
}

fn procedimento_subclass_interno(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    uid_subclass: usize,
    _ref_data: usize,
) -> LRESULT {
    const WM_DESTROY: u32 = 0x0002;
    const WM_SIZE: u32 = 0x0005;
    const WM_NOTIFY: u32 = 0x004E;
    const WM_NCDESTROY: u32 = 0x0082;
    const WM_TIMER: u32 = 0x0113;

    // Mensagens REA — postadas pela thread de background e recebidas aqui na thread de UI
    let wm_inicializar = crate::tab_manager::WM_OFFICETAB_INICIALIZAR;
    let wm_recarregar = crate::tab_manager::WM_OFFICETAB_RECARREGAR;
    let wm_novo_doc = crate::tab_manager::WM_OFFICETAB_NOVO_DOC;

    match msg {
        // WM_TIMER: atualização periódica (disparado na thread de UI)
        WM_TIMER if wparam.0 == crate::tab_manager::OFFICETAB_TIMER_ID => {
            crate::tab_manager::despachar_timer(hwnd);
            return LRESULT(0);
        }

        // WM_SIZE: redimensionamento da janela do Office (thread de UI)
        WM_SIZE => {
            let res = unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) };
            crate::tab_manager::despachar_size(hwnd);
            return res;
        }

        // WM_NOTIFY: notificações do TabControl (clique em aba, etc.)
        WM_NOTIFY => {
            if let Some(result) = crate::tab_manager::despachar_notify(hwnd, lparam) {
                return result;
            }
        }

        // WM_KEYDOWN: atalhos de teclado do OfficeTab (Ctrl+Tab, Ctrl+Shift+Tab, Ctrl+W, Ctrl+T)
        0x0100 => {
            if crate::tab_manager::despachar_keydown(hwnd, wparam.0) {
                return LRESULT(0);
            }
        }

        // WM_OFFICETAB_INICIALIZAR: REA — postado pela thread de background
        // quando detecta nova janela do Office. Executa subclassing e criação
        // do TabControl com segurança pois estamos na thread de UI.
        m if m == wm_inicializar => {
            crate::tab_manager::despachar_inicializar(hwnd);
            return LRESULT(0);
        }

        // WM_OFFICETAB_RECARREGAR: REA — força recarga de abas
        m if m == wm_recarregar => {
            crate::tab_manager::despachar_recarregar(hwnd);
            return LRESULT(0);
        }

        // WM_OFFICETAB_NOVO_DOC: criação assíncrona de novo documento
        m if m == wm_novo_doc => {
            crate::tab_manager::despachar_novo_doc(hwnd);
            return LRESULT(0);
        }

        // WM_DESTROY: sinaliza destruição da janela para limpeza de estados
        WM_DESTROY => {
            crate::tab_manager::despachar_destroy(hwnd);
        }

        // WM_NCDESTROY: última mensagem da janela. Garante obrigatoriamente
        // que RemoveWindowSubclass seja executado caso o Office destrua a janela antes do cleanup.
        WM_NCDESTROY => {
            crate::tab_manager::despachar_destroy(hwnd);
            unsafe {
                let _ = RemoveWindowSubclass(hwnd, Some(procedimento_subclass), uid_subclass);
            }
            if let Ok(mut guard) = JANELAS_SUBCLASSED.lock()
                && let Some(ref mut set) = *guard
            {
                set.remove(&(hwnd.0 as isize));
            }
            return unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) };
        }

        _ => {}
    }

    // Encaminha para o próximo procedimento na cadeia (preservando o comportamento nativo do Office)
    unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) }
}

#[cfg(test)]
mod testes {
    use super::*;

    #[test]
    fn test_subclass_manager_criacao_e_limpeza() {
        let manager = SubclassManager::new();
        // Não deve entrar em pânico ao desinscrever tudo quando vazio
        manager.unsubclass_all();
    }
}
