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
/// Todas as mensagens Win32 chegam aqui NA THREAD DE UI DO OFFICE — seguro para Win32.
unsafe extern "system" fn procedimento_subclass(
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
            crate::tab_manager::despachar_notify(hwnd, lparam);
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

        // WM_DESTROY / WM_NCDESTROY: limpeza ao fechar a janela
        WM_NCDESTROY | WM_DESTROY => {
            crate::tab_manager::despachar_destroy(hwnd);
            unsafe {
                let _ = RemoveWindowSubclass(hwnd, Some(procedimento_subclass), uid_subclass);
            }
            if let Ok(mut guard) = JANELAS_SUBCLASSED.lock()
                && let Some(ref mut set) = *guard
            {
                set.remove(&(hwnd.0 as isize));
            }
        }

        _ => {}
    }

    // Encaminha para o próximo procedimento na cadeia (preservando o comportamento nativo do Office)
    unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) }
}
