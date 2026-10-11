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

/// Procedimento de janela subclassed. Repassa mensagens mantendo a integridade do Office.
unsafe extern "system" fn procedimento_subclass(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    uid_subclass: usize,
    _ref_data: usize,
) -> LRESULT {
    const WM_DESTROY: u32 = 0x0002;
    const WM_NCDESTROY: u32 = 0x0082;

    if msg == WM_NCDESTROY || msg == WM_DESTROY {
        unsafe {
            let _ = RemoveWindowSubclass(hwnd, Some(procedimento_subclass), uid_subclass);
        }
        if let Ok(mut guard) = JANELAS_SUBCLASSED.lock()
            && let Some(ref mut set) = *guard
        {
            set.remove(&(hwnd.0 as isize));
        }
    }

    // Encaminha a mensagem para o próximo procedimento na cadeia (preservando o comportamento nativo do Office)
    unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) }
}
