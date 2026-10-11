//! Wrapper para o controle nativo SysTabControl32 do Windows.

use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::Controls::{
    TCIF_TEXT, TCITEMW, TCM_DELETEITEM, TCM_GETCURSEL, TCM_GETITEMCOUNT, TCM_INSERTITEMW,
    TCM_SETCURSEL, TCS_FIXEDWIDTH, TCS_TABS,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, HMENU, MoveWindow, SWP_NOACTIVATE, SWP_NOZORDER, SendMessageW,
    SetWindowPos, WINDOW_EX_STYLE, WINDOW_STYLE, WS_CHILD, WS_CLIPSIBLINGS, WS_VISIBLE,
};
use windows::core::{PCWSTR, Result, w};

/// Estrutura que encapsula um controle de abas nativo do Windows.
pub struct TabControl {
    pub hwnd: HWND,
}

impl TabControl {
    /// Cria um novo controle de abas como filho da janela especificada.
    pub fn new(parent_hwnd: HWND, id: u32) -> Result<Self> {
        let estilo =
            WINDOW_STYLE(WS_CHILD.0 | WS_VISIBLE.0 | WS_CLIPSIBLINGS.0 | TCS_TABS | TCS_FIXEDWIDTH);

        unsafe {
            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("SysTabControl32"),
                PCWSTR::null(),
                estilo,
                0,
                0,
                0,
                28,
                parent_hwnd,
                HMENU(id as usize as *mut _),
                None,
                None,
            )?;

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
                HWND::default(),
                x,
                y,
                largura,
                altura,
                SWP_NOACTIVATE | SWP_NOZORDER,
            )?;
            Ok(())
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
