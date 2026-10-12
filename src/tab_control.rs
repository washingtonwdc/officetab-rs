// Wrapper para controle de abas nativo do Windows (SysTabControl32).
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

use windows::core::{w, PCWSTR, Result};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::Graphics::Gdi::{
    GetStockObject, InvalidateRect, UpdateWindow, DEFAULT_GUI_FONT,
};
use windows::Win32::System::LibraryLoader::{
    GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
    GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
    GetModuleHandleExW, GetModuleHandleW,
};
use windows::Win32::UI::Controls::{
    InitCommonControlsEx, ICC_TAB_CLASSES, INITCOMMONCONTROLSEX, TCIF_TEXT, TCITEMW,
    TCM_DELETEALLITEMS, TCM_DELETEITEM, TCM_GETCURSEL, TCM_GETITEMCOUNT, TCM_INSERTITEMW,
    TCM_SETCURSEL, TCS_FOCUSNEVER, TCS_TABS,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BringWindowToTop, CreateWindowExW, DestroyWindow, MoveWindow, SendMessageW, SetWindowPos,
    ShowWindow, HMENU, HWND_TOP, SWP_SHOWWINDOW, SW_SHOW,
    WINDOW_EX_STYLE, WINDOW_STYLE, WM_SETFONT, WS_CHILD, WS_CLIPCHILDREN, WS_CLIPSIBLINGS, WS_VISIBLE,
};


/// Estrutura que encapsula um controle de abas nativo do Windows.
pub struct TabControl {
    pub hwnd: HWND,
}

/// Obtém o HINSTANCE do módulo que contém a função `obter_hinstance_dll`.
/// Isso garante que o TabControl seja criado com o HINSTANCE correto da nossa DLL
/// e não do processo do Word, evitando problemas de reclassificação de janelas.
fn obter_hinstance_dll() -> windows::Win32::Foundation::HMODULE {
    unsafe {
        let mut hmod = windows::Win32::Foundation::HMODULE::default();
        // Passa o endereço desta própria função para localizar nossa DLL na memória
        let addr = obter_hinstance_dll as *const std::ffi::c_void;
        let ok = GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS
                | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            windows::core::PCWSTR(addr as *const u16),
            &mut hmod,
        );
        if ok.is_ok() && !hmod.is_invalid() {
            hmod
        } else {
            // Fallback: tenta pegar o handle pelo nome do módulo
            GetModuleHandleW(None).unwrap_or_default()
        }
    }
}


impl TabControl {
    /// Cria um novo controle de abas como filho da janela especificada com posição e tamanho definidos.
    /// DEVE ser chamado da thread de UI que possui a janela pai (parent_hwnd).
    pub fn new(
        parent_hwnd: HWND,
        id: u32,
        x: i32,
        y: i32,
        largura: i32,
        altura: i32,
    ) -> Result<Self> {
        unsafe {
            // Inicializa os controles comuns (necessário para SysTabControl32)
            let icex = INITCOMMONCONTROLSEX {
                dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
                dwICC: ICC_TAB_CLASSES,
            };
            let _ = InitCommonControlsEx(&icex);

            let estilo = WINDOW_STYLE(
                WS_CHILD.0
                    | WS_VISIBLE.0
                    | WS_CLIPSIBLINGS.0
                    | WS_CLIPCHILDREN.0
                    | TCS_TABS
                    | TCS_FOCUSNEVER,
            );

            // CRÍTICO: usar o HINSTANCE da nossa DLL, não do processo do Word
            let hinstance = obter_hinstance_dll();

            crate::tab_manager::log_debug(&format!(
                "TabControl::new — pai=0x{:X} x={} y={} w={} h={} hmod=0x{:X}",
                parent_hwnd.0 as usize, x, y, largura, altura, hinstance.0 as usize
            ));

            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("SysTabControl32"),
                PCWSTR::null(),
                estilo,
                x,
                y,
                largura,
                altura,
                parent_hwnd,
                HMENU(id as usize as *mut _),
                hinstance,
                None,
            )?;

            crate::tab_manager::log_debug(&format!(
                "TabControl::new — HWND criado: 0x{:X}",
                hwnd.0 as usize
            ));

            // Aplicar fonte nítida padrão da interface do Windows
            let hfont = GetStockObject(DEFAULT_GUI_FONT);
            if !hfont.is_invalid() {
                let _ = SendMessageW(hwnd, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            }

            // Posicionar no topo da Z-Order, forçar visibilidade e repintura
            let _ = SetWindowPos(
                hwnd,
                HWND_TOP,
                x,
                y,
                largura,
                altura,
                SWP_SHOWWINDOW,
            );
            let _ = BringWindowToTop(hwnd);
            let _ = ShowWindow(hwnd, SW_SHOW);
            let _ = UpdateWindow(hwnd);
            let _ = InvalidateRect(hwnd, None, true);

            Ok(Self { hwnd })
        }
    }

    /// Adiciona uma nova aba ao final do controle.
    pub fn add_tab(&self, index: i32, titulo: &str) -> Result<i32> {
        let mut utf16: Vec<u16> = titulo.encode_utf16().chain(std::iter::once(0)).collect();

        let mut item = TCITEMW {
            mask: TCIF_TEXT,
            pszText: windows::core::PWSTR(utf16.as_mut_ptr()),
            cchTextMax: utf16.len() as i32,
            ..Default::default()
        };

        unsafe {
            let res = SendMessageW(
                self.hwnd,
                TCM_INSERTITEMW,
                WPARAM(index as usize),
                LPARAM(&mut item as *mut _ as isize),
            );

            if res.0 < 0 {
                return Err(windows::core::Error::from_win32());
            }

            Ok(res.0 as i32)
        }
    }

    /// Remove a aba pelo índice.
    pub fn remove_tab(&self, index: i32) -> Result<()> {
        unsafe {
            let res = SendMessageW(self.hwnd, TCM_DELETEITEM, WPARAM(index as usize), LPARAM(0));

            if res.0 == 0 {
                return Err(windows::core::Error::from_win32());
            }

            Ok(())
        }
    }

    /// Remove todas as abas existentes do controle.
    pub fn clear_tabs(&self) {
        unsafe {
            let _ = SendMessageW(self.hwnd, TCM_DELETEALLITEMS, WPARAM(0), LPARAM(0));
        }
    }

    /// Retorna o índice da aba atualmente selecionada.
    pub fn get_selected_tab(&self) -> i32 {
        unsafe {
            let res = SendMessageW(self.hwnd, TCM_GETCURSEL, WPARAM(0), LPARAM(0));
            res.0 as i32
        }
    }

    /// Define a aba ativa.
    pub fn set_selected_tab(&self, index: i32) -> Result<()> {
        unsafe {
            let res = SendMessageW(self.hwnd, TCM_SETCURSEL, WPARAM(index as usize), LPARAM(0));

            if res.0 < 0 {
                return Err(windows::core::Error::from_win32());
            }

            Ok(())
        }
    }

    /// Retorna a quantidade total de abas.
    pub fn get_tab_count(&self) -> i32 {
        unsafe {
            let res = SendMessageW(self.hwnd, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0));
            res.0 as i32
        }
    }

    /// Atualiza as dimensões e o posicionamento da barra de abas.
    pub fn resize(&self, x: i32, y: i32, largura: i32, altura: i32) -> Result<()> {
        unsafe {
            MoveWindow(self.hwnd, x, y, largura, altura, true)?;
            SetWindowPos(
                self.hwnd,
                HWND_TOP,
                x,
                y,
                largura,
                altura,
                SWP_SHOWWINDOW,
            )?;
            let _ = ShowWindow(self.hwnd, SW_SHOW);
            let _ = InvalidateRect(self.hwnd, None, true);
            Ok(())
        }
    }

    /// Força a repintura imediata e traz para a frente da Z-Order.
    pub fn repaint(&self) {
        unsafe {
            let _ = SetWindowPos(
                self.hwnd,
                HWND_TOP,
                0,
                0,
                0,
                0,
                windows::Win32::UI::WindowsAndMessaging::SWP_NOMOVE
                    | windows::Win32::UI::WindowsAndMessaging::SWP_NOSIZE
                    | SWP_SHOWWINDOW,
            );
            let _ = BringWindowToTop(self.hwnd);
            let _ = ShowWindow(self.hwnd, SW_SHOW);
            let _ = InvalidateRect(self.hwnd, None, true);
            let _ = UpdateWindow(self.hwnd);
        }
    }

    /// Destrói o controle de abas.
    pub fn destroy(&mut self) {
        if !self.hwnd.is_invalid() {
            unsafe {
                let _ = DestroyWindow(self.hwnd);
            }
            self.hwnd = HWND::default();
        }
    }
}

impl Drop for TabControl {
    fn drop(&mut self) {
        self.destroy();
    }
}

unsafe impl Send for TabControl {}
unsafe impl Sync for TabControl {}
