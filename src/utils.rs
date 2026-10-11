//! Utilitários Win32 e suporte ao ambiente do suplemento Office Tab.

use std::path::PathBuf;
use windows::Win32::Foundation::{HMODULE, HWND};
use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx};
use windows::Win32::System::LibraryLoader::{
    GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
    GetModuleFileNameW, GetModuleHandleExW, GetModuleHandleW, GetProcAddress,
};
use windows::core::{PCWSTR, Result, s, w};

/// Inicializa o subsistema COM em modo Single-Threaded Apartment (STA), padrão do Office.
pub fn init_com() -> Result<()> {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
    Ok(())
}

/// Obtém o DPI associado à janela informada.
pub fn get_dpi_for_window(hwnd: HWND) -> u32 {
    unsafe {
        if let Ok(user32) = GetModuleHandleW(w!("user32.dll"))
            && let Some(proc) = GetProcAddress(user32, s!("GetDpiForWindow"))
        {
            type GetDpiFn = unsafe extern "system" fn(HWND) -> u32;
            let func: GetDpiFn = std::mem::transmute(proc);
            return func(hwnd);
        }
        96 // 96 DPI padrão (escala de 100%)
    }
}

/// Define a conscientização de DPI por monitor v2 para o processo.
pub fn set_dpi_awareness() -> Result<()> {
    unsafe {
        if let Ok(user32) = GetModuleHandleW(w!("user32.dll"))
            && let Some(proc) = GetProcAddress(user32, s!("SetProcessDpiAwarenessContext"))
        {
            type SetContextFn =
                unsafe extern "system" fn(isize) -> windows::Win32::Foundation::BOOL;
            let func: SetContextFn = std::mem::transmute(proc);
            let _ = func(-4); // -4 = DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        }
    }
    Ok(())
}

/// Obtém o caminho absoluto do arquivo desta DLL em execução.
pub fn get_current_module_path() -> Result<PathBuf> {
    unsafe {
        let mut hmodule = HMODULE::default();
        let endereco_marcador = get_current_module_path as *const () as *const u16;

        GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            PCWSTR(endereco_marcador),
            &mut hmodule,
        )?;

        let mut buffer = vec![0u16; 1024];
        let len = GetModuleFileNameW(hmodule, &mut buffer);
        if len == 0 {
            return Err(windows::core::Error::from_win32());
        }

        buffer.truncate(len as usize);
        let caminho_str = String::from_utf16_lossy(&buffer);
        Ok(PathBuf::from(caminho_str))
    }
}
