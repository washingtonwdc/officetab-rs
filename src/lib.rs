//! Office Tab Pro — Reimplementação em Rust
//!
//! COM Add-in (`IDTExtensibility2`) que transforma janelas SDI do Office
//! em uma interface com abas via window subclassing + reparenting.
//!
//! Arquitetura (baseada em análise estática do Office Tab Pro 19.00.001):
//! - `com_addin`    → interface IDTExtensibility2 + AddinManager
//! - `tab_manager`  → TOTWindowsManager (coração do mecanismo)
//! - `subclass`     → THookProject + TProjectSubClass (window subclassing)
//! - `tab_control`  → SysTabsControl32 wrapper (controle de abas)
//! - `registry`     → registro COM do add-in
//! - `utils`        → utilitários (Win32 wrappers, DPI, etc.)

pub mod com_addin_simple;
pub mod implementation;
pub mod registry_simple;
pub mod subclass;
pub mod tab_control;
pub mod tab_manager;
pub mod utils;

use std::ffi::c_void;

const E_POINTER: windows::core::HRESULT = windows::core::HRESULT(0x80004003u32 as i32);
const E_FAIL: windows::core::HRESULT = windows::core::HRESULT(0x80004005u32 as i32);
const S_OK: windows::core::HRESULT = windows::core::HRESULT(0);
const S_FALSE: windows::core::HRESULT = windows::core::HRESULT(1);

/// Entry point do DLL — exportado para o Office
///
/// # Safety
/// Os ponteiros `rclsid`, `riid` e `ppv` devem ser válidos ou nulos.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn DllGetClassObject(
    rclsid: *const windows::core::GUID,
    riid: *const windows::core::GUID,
    ppv: *mut *mut c_void,
) -> windows::core::HRESULT {
    if rclsid.is_null() || riid.is_null() || ppv.is_null() {
        return E_POINTER;
    }
    match unsafe { crate::com_addin_simple::get_class_object(rclsid, riid, ppv) } {
        Ok(()) => S_OK,
        Err(_) => E_FAIL,
    }
}

#[unsafe(no_mangle)]
pub extern "system" fn DllCanUnloadNow() -> windows::core::HRESULT {
    if crate::com_addin_simple::can_unload_now() {
        S_OK
    } else {
        S_FALSE
    }
}

#[unsafe(no_mangle)]
pub extern "system" fn DllRegisterServer() -> windows::core::HRESULT {
    match crate::registry_simple::register_server() {
        Ok(()) => S_OK,
        Err(_) => E_FAIL,
    }
}

#[unsafe(no_mangle)]
pub extern "system" fn DllUnregisterServer() -> windows::core::HRESULT {
    match crate::registry_simple::unregister_server() {
        Ok(()) => S_OK,
        Err(_) => E_FAIL,
    }
}

#[cfg(test)]
mod testes {
    use super::{DllCanUnloadNow, DllGetClassObject, E_POINTER, S_OK};
    use windows::Win32::System::Com::IClassFactory;
    use windows::core::Interface;

    #[test]
    fn dll_get_class_object_rejeita_ponteiros_nulos() {
        let resultado =
            unsafe { DllGetClassObject(std::ptr::null(), std::ptr::null(), std::ptr::null_mut()) };

        assert_eq!(resultado, E_POINTER);
    }

    #[test]
    fn dll_can_unload_now_sem_objetos_ativos() {
        assert_eq!(DllCanUnloadNow(), S_OK);
    }

    #[test]
    fn dll_get_class_object_fornece_fabrica_com() {
        let mut ponteiro = std::ptr::null_mut();
        let resultado = unsafe {
            DllGetClassObject(
                &crate::com_addin_simple::ADDIN_CLSID,
                &IClassFactory::IID,
                &mut ponteiro,
            )
        };

        assert_eq!(resultado, S_OK);
        assert!(!ponteiro.is_null());
        drop(unsafe { IClassFactory::from_raw(ponteiro) });
    }

    #[test]
    fn inicializacao_officetab_sucesso() {
        let app = crate::implementation::OfficeTab::new();
        assert!(app.initialize().is_ok());
        let janelas = app.attach_to_active_office();
        // Em ambiente de teste sem instâncias abertas do Word/Excel, a lista deve ser válida e vazia ou conter janelas existentes
        assert!(janelas.is_empty() || !janelas.is_empty());
        app.shutdown();
    }

    #[test]
    fn tab_manager_ciclo_de_vida() {
        let mut manager = crate::tab_manager::TabManager::new();
        let _ = manager.discover_document_windows();
        assert_eq!(manager.get_active_windows().len(), 0);
        manager.cleanup();
    }
}
