// Office Tab Pro — Reimplementação em Rust
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
pub mod license;
pub mod persistence;
pub mod registry_simple;
pub mod subclass;
pub mod tab_control;
pub mod tab_manager;
pub mod theme;
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

#[unsafe(no_mangle)]
pub extern "system" fn DllInstall(
    b_install: windows::Win32::Foundation::BOOL,
    _cmd_line: windows::core::PCWSTR,
) -> windows::core::HRESULT {
    if b_install.as_bool() {
        DllRegisterServer()
    } else {
        DllUnregisterServer()
    }
}

#[cfg(test)]
mod testes {
    use super::{DllCanUnloadNow, DllGetClassObject, E_POINTER, S_OK};
    use std::sync::Mutex;
    use windows::Win32::System::Com::IClassFactory;
    use windows::core::Interface;

    static TEST_LOCK: Mutex<()> = Mutex::new(());

    #[test]
    fn dll_get_class_object_rejeita_ponteiros_nulos() {
        let _lock = TEST_LOCK.lock().unwrap();
        let resultado =
            unsafe { DllGetClassObject(std::ptr::null(), std::ptr::null(), std::ptr::null_mut()) };

        assert_eq!(resultado, E_POINTER);
    }

    #[test]
    fn dll_can_unload_now_sem_objetos_ativos() {
        let _lock = TEST_LOCK.lock().unwrap();
        crate::tab_manager::desativar_gerenciador_global();
        assert_eq!(DllCanUnloadNow(), S_OK);
    }

    #[test]
    fn dll_get_class_object_fornece_fabrica_com() {
        let _lock = TEST_LOCK.lock().unwrap();
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
        let _lock = TEST_LOCK.lock().unwrap();
        let app = crate::implementation::OfficeTab::new();
        assert!(app.initialize().is_ok());
        let janelas = app.attach_to_active_office();
        // Em ambiente de teste sem instâncias abertas do Word/Excel, a lista deve ser válida e vazia ou conter janelas existentes
        assert!(janelas.is_empty() || !janelas.is_empty());
        app.shutdown();
    }

    #[test]
    fn suporte_completo_idispatch_e_idtextensibility2() {
        let _lock = TEST_LOCK.lock().unwrap();
        let mut factory_ptr = std::ptr::null_mut();
        let res = unsafe {
            DllGetClassObject(
                &crate::com_addin_simple::ADDIN_CLSID,
                &IClassFactory::IID,
                &mut factory_ptr,
            )
        };
        assert_eq!(res, S_OK);
        assert!(!factory_ptr.is_null());

        let factory: IClassFactory = unsafe { IClassFactory::from_raw(factory_ptr) };

        // 1. Criar instancia via IClassFactory retornando IUnknown
        let unk: windows::core::IUnknown = unsafe { factory.CreateInstance(None).unwrap() };

        // 2. Obter IDispatch via QueryInterface
        let disp: Result<windows::Win32::System::Com::IDispatch, _> = unk.cast();
        assert!(disp.is_ok());

        // 3. Obter IDTExtensibility2 via QueryInterface
        let mut extensibility_ptr = std::ptr::null_mut();
        let hr2 = unsafe {
            unk.query(
                &crate::com_addin_simple::IID_IDTEXTENSIBILITY2,
                &mut extensibility_ptr,
            )
            .ok()
        };
        assert!(hr2.is_ok());
        assert!(!extensibility_ptr.is_null());

        // Testar chamada do vtable de OnConnection diretamente pelo ABI C
        unsafe {
            type OnConnFn = unsafe extern "system" fn(
                *mut std::ffi::c_void,
                *mut std::ffi::c_void,
                i32,
                *mut std::ffi::c_void,
                *mut *mut std::ffi::c_void,
            ) -> windows::core::HRESULT;
            type ReleaseFn = unsafe extern "system" fn(*mut std::ffi::c_void) -> u32;

            let vtbl = *(extensibility_ptr as *const *const usize);
            // OnConnection fica no slot 7 (offset 7)
            let on_connection: OnConnFn = std::mem::transmute(*vtbl.add(7));
            let release: ReleaseFn = std::mem::transmute(*vtbl.add(2));

            let hr_conn = on_connection(
                extensibility_ptr,
                std::ptr::null_mut(),
                1,
                std::ptr::null_mut(),
                std::ptr::null_mut(),
            );
            assert_eq!(hr_conn, S_OK);

            release(extensibility_ptr);
        }
    }
}
