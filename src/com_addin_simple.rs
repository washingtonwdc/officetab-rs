// Implementacao binaria exata de IDTExtensibility2 e IClassFactory para o Microsoft Office.
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
//! Implementacao binaria exata de IDTExtensibility2 e IClassFactory para o Microsoft Office.
#![allow(non_snake_case)]

use std::ffi::c_void;
use std::sync::Mutex;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::BOOL;
use windows::Win32::System::Com::{
    DISPPARAMS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, IClassFactory, IClassFactory_Impl,
    IDispatch,
};
use windows::Win32::System::Variant::{VT_DISPATCH, VariantClear};
use windows::core::{
    Error, GUID, HRESULT, IUnknown, Interface, PCWSTR, Result, VARIANT, implement, w,
};

/// GUIDs oficiais de classes e interfaces do Microsoft Office
pub const ADDIN_CLSID: GUID = GUID::from_u128(0x9a1929ef_1495_413e_993c_326dd15e35df);
pub const ADDIN_PROG_ID: &str = "OfficeTabRs.Connect";

pub const IID_IDISPATCH: GUID = GUID::from_u128(0x00020400_0000_0000_c000_000000000046);
pub const IID_IDTEXTENSIBILITY2: GUID = GUID::from_u128(0xb65ad801_abaf_11d0_bb8b_00a0c90f2744);

static OFFICE_APP_DISPATCH: Mutex<Option<usize>> = Mutex::new(None);

pub fn salvar_application_dispatch(raw_ptr: *mut c_void) {
    if raw_ptr.is_null() {
        return;
    }
    unsafe {
        let unk = std::mem::ManuallyDrop::new(std::mem::transmute_copy::<*mut c_void, IUnknown>(&raw_ptr));
        let cloned = (*unk).clone();
        let raw = cloned.into_raw();
        if let Ok(mut guard) = OFFICE_APP_DISPATCH.lock() {
            *guard = Some(raw as usize);
        }
        crate::tab_manager::log_debug(&format!(
            "salvar_application_dispatch: salvo com sucesso ptr=0x{:X}",
            raw as usize
        ));
    }
}

pub fn liberar_application_dispatch() {
    if let Ok(mut guard) = OFFICE_APP_DISPATCH.lock()
        && let Some(ptr_val) = guard.take()
    {
        unsafe {
            let _ = IUnknown::from_raw(ptr_val as *mut c_void);
            crate::tab_manager::log_debug("liberar_application_dispatch: liberado com sucesso");
        }
    }
}

pub fn criar_documento_com() -> bool {
    let guard = match OFFICE_APP_DISPATCH.lock() {
        Ok(g) => g,
        Err(_) => return false,
    };
    let raw_ptr = match *guard {
        Some(p) => p as *mut c_void,
        None => {
            crate::tab_manager::log_debug("criar_documento_com: OFFICE_APP_DISPATCH é None");
            return false;
        }
    };
    drop(guard);

    unsafe {
        let app = std::mem::ManuallyDrop::new(std::mem::transmute_copy::<*mut c_void, IDispatch>(&raw_ptr));

        // Nomes das coleções de documentos nos principais apps Office:
        // Word -> "Documents"
        // Excel -> "Workbooks"
        // PowerPoint -> "Presentations"
        // Project -> "Projects"
        let colecoes = [
            w!("Documents"),
            w!("Workbooks"),
            w!("Presentations"),
            w!("Projects"),
        ];

        for nome in colecoes {
            let mut dispid_col = 0i32;
            let hr = app.GetIDsOfNames(&GUID::zeroed(), &nome, 1, 0, &mut dispid_col);
            if hr.is_err() {
                continue;
            }

            let dp = DISPPARAMS::default();
            let mut result_col = VARIANT::default();
            let hr_col = app.Invoke(
                dispid_col,
                &GUID::zeroed(),
                0,
                DISPATCH_PROPERTYGET,
                &dp,
                Some(&mut result_col),
                None,
                None,
            );
            if hr_col.is_err() {
                let _ = VariantClear(&mut result_col);
                continue;
            }

            #[repr(C)]
            struct RawVariant {
                vt: u16,
                r1: u16,
                r2: u16,
                r3: u16,
                pdisp_val: *mut c_void,
                extra: usize,
            }
            let raw = &result_col as *const VARIANT as *const RawVariant;
            if (*raw).vt == VT_DISPATCH.0 && !(*raw).pdisp_val.is_null() {
                let col = std::mem::ManuallyDrop::new(std::mem::transmute_copy::<*mut c_void, IDispatch>(&(*raw).pdisp_val));
                let mut add_id = 0i32;
                let add_name = w!("Add");
                let hr_add_name = col.GetIDsOfNames(&GUID::zeroed(), &add_name, 1, 0, &mut add_id);
                let mut sucesso = false;
                if hr_add_name.is_ok() {
                    let mut result_doc = VARIANT::default();
                    let hr_add = col.Invoke(
                        add_id,
                        &GUID::zeroed(),
                        0,
                        DISPATCH_METHOD,
                        &dp,
                        Some(&mut result_doc),
                        None,
                        None,
                    );
                    let _ = VariantClear(&mut result_doc);

                    if hr_add.is_ok() {
                        crate::tab_manager::log_debug(
                            "criar_documento_com: Novo documento criado com SUCESSO via COM!",
                        );
                        sucesso = true;
                    } else {
                        crate::tab_manager::log_debug(&format!(
                            "criar_documento_com: Invoke(Add) falhou: {:?}",
                            hr_add
                        ));
                    }
                } else {
                    crate::tab_manager::log_debug(&format!(
                        "criar_documento_com: GetIDsOfNames(Add) falhou: {:?}",
                        hr_add_name
                    ));
                }

                // Libera result_col de forma segura uma única vez
                let _ = VariantClear(&mut result_col);
                if sucesso {
                    return true;
                }
            } else {
                let _ = VariantClear(&mut result_col);
            }
        }

        crate::tab_manager::log_debug("criar_documento_com: nenhuma colecao Office respondeu");
        false
    }
}

const S_OK: HRESULT = HRESULT(0);
const E_NOTIMPL: HRESULT = HRESULT(0x80004001u32 as i32);
const E_NOINTERFACE: HRESULT = HRESULT(0x80004002u32 as i32);
const E_POINTER: HRESULT = HRESULT(0x80004003u32 as i32);
const CLASS_E_NOAGGREGATION: HRESULT = HRESULT(0x80040110u32 as i32);
const DISP_E_UNKNOWNNAME: HRESULT = HRESULT(0x80020006u32 as i32);
const DISP_E_MEMBERNOTFOUND: HRESULT = HRESULT(0x80020003u32 as i32);

static OBJETOS_ATIVOS: AtomicU32 = AtomicU32::new(0);
static BLOQUEIOS_SERVIDOR: AtomicU32 = AtomicU32::new(0);

/// Tabela de funções (Vtable) binária do IDTExtensibility2 (herdeira de IDispatch).
/// A ordem e tipos dos 12 métodos correspondem rigorosamente ao ABI COM do Office.
#[repr(C)]
pub struct IDTExtensibility2_Vtbl {
    // IUnknown (slots 0..2)
    pub QueryInterface:
        unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
    pub AddRef: unsafe extern "system" fn(*mut c_void) -> u32,
    pub Release: unsafe extern "system" fn(*mut c_void) -> u32,

    // IDispatch (slots 3..6)
    pub GetTypeInfoCount: unsafe extern "system" fn(*mut c_void, *mut u32) -> HRESULT,
    pub GetTypeInfo: unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> HRESULT,
    pub GetIDsOfNames: unsafe extern "system" fn(
        *mut c_void,
        *const GUID,
        *const PCWSTR,
        u32,
        u32,
        *mut i32,
    ) -> HRESULT,
    pub Invoke: unsafe extern "system" fn(
        *mut c_void,
        i32,
        *const GUID,
        u32,
        u16,
        *const c_void,
        *mut c_void,
        *mut c_void,
        *mut u32,
    ) -> HRESULT,

    // IDTExtensibility2 (slots 7..11)
    pub OnConnection: unsafe extern "system" fn(
        *mut c_void,
        *mut c_void, // Application (IDispatch*)
        i32,         // ConnectMode
        *mut c_void, // AddInInst (IDispatch*)
        *mut *mut c_void,
    ) -> HRESULT,
    pub OnDisconnection: unsafe extern "system" fn(
        *mut c_void,
        i32, // RemoveMode
        *mut *mut c_void,
    ) -> HRESULT,
    pub OnAddInsUpdate: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> HRESULT,
    pub OnStartupComplete: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> HRESULT,
    pub OnBeginShutdown: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> HRESULT,
}

#[repr(C)]
struct AddinInstance {
    lp_vtbl: *const IDTExtensibility2_Vtbl,
    ref_count: AtomicU32,
    instancia: Mutex<Option<crate::implementation::OfficeTab>>,
}

// Implementações C-ABI para a Vtable
unsafe extern "system" fn addin_query_interface(
    this: *mut c_void,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    if ppv.is_null() || riid.is_null() {
        return E_POINTER;
    }
    unsafe { *ppv = std::ptr::null_mut() };

    let iid = unsafe { &*riid };
    if *iid == IUnknown::IID || *iid == IID_IDISPATCH || *iid == IID_IDTEXTENSIBILITY2 {
        unsafe {
            addin_add_ref(this);
            *ppv = this;
        }
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn addin_add_ref(this: *mut c_void) -> u32 {
    let instance = unsafe { &*(this as *const AddinInstance) };
    instance.ref_count.fetch_add(1, Ordering::Relaxed) + 1
}

unsafe extern "system" fn addin_release(this: *mut c_void) -> u32 {
    let instance_ptr = this as *mut AddinInstance;
    let old_count = unsafe { (*instance_ptr).ref_count.fetch_sub(1, Ordering::Release) };
    if old_count == 1 {
        std::sync::atomic::fence(Ordering::Acquire);
        OBJETOS_ATIVOS.fetch_sub(1, Ordering::AcqRel);
        drop(unsafe { Box::from_raw(instance_ptr) });
        0
    } else {
        old_count - 1
    }
}

unsafe extern "system" fn addin_get_type_info_count(
    _this: *mut c_void,
    pctinfo: *mut u32,
) -> HRESULT {
    if !pctinfo.is_null() {
        unsafe { *pctinfo = 0 };
    }
    S_OK
}

unsafe extern "system" fn addin_get_type_info(
    _this: *mut c_void,
    _itinfo: u32,
    _lcid: u32,
    pptinfo: *mut *mut c_void,
) -> HRESULT {
    if !pptinfo.is_null() {
        unsafe { *pptinfo = std::ptr::null_mut() };
    }
    E_NOTIMPL
}

unsafe extern "system" fn addin_get_ids_of_names(
    _this: *mut c_void,
    _riid: *const GUID,
    _rgsznames: *const PCWSTR,
    _cnames: u32,
    _lcid: u32,
    _rgdispid: *mut i32,
) -> HRESULT {
    DISP_E_UNKNOWNNAME
}

unsafe extern "system" fn addin_invoke(
    _this: *mut c_void,
    _dispidmember: i32,
    _riid: *const GUID,
    _lcid: u32,
    _wflags: u16,
    _pdispparams: *const c_void,
    _pvarresult: *mut c_void,
    _pexcepinfo: *mut c_void,
    _puargerr: *mut u32,
) -> HRESULT {
    DISP_E_MEMBERNOTFOUND
}

unsafe extern "system" fn addin_on_connection(
    this: *mut c_void,
    application: *mut c_void,
    connect_mode: i32,
    _add_in_inst: *mut c_void,
    _custom: *mut *mut c_void,
) -> HRESULT {
    crate::tab_manager::log_debug(&format!("addin_on_connection: mode={}", connect_mode));
    salvar_application_dispatch(application);
    let instance = unsafe { &*(this as *const AddinInstance) };
    let app = crate::implementation::OfficeTab::new();
    // initialize() inicia a thread de monitoramento REA — ela fará o PostMessage
    // para a thread de UI, que criará os TabControls com segurança de thread.
    // NÃO chamamos attach_to_active_office aqui pois estaria na thread COM,
    // não na thread de UI do Word.
    let _ = app.initialize();
    if let Ok(mut lock) = instance.instancia.lock() {
        *lock = Some(app);
    }
    S_OK
}

unsafe extern "system" fn addin_on_disconnection(
    this: *mut c_void,
    _remove_mode: i32,
    _custom: *mut *mut c_void,
) -> HRESULT {
    crate::tab_manager::log_debug(&format!(
        "addin_on_disconnection: remove_mode={}",
        _remove_mode
    ));
    liberar_application_dispatch();
    let instance = unsafe { &*(this as *const AddinInstance) };
    if let Ok(mut lock) = instance.instancia.lock()
        && let Some(app) = lock.take()
    {
        app.shutdown();
    }
    S_OK
}

unsafe extern "system" fn addin_on_add_ins_update(
    _this: *mut c_void,
    _custom: *mut *mut c_void,
) -> HRESULT {
    S_OK
}

unsafe extern "system" fn addin_on_startup_complete(
    _this: *mut c_void,
    _custom: *mut *mut c_void,
) -> HRESULT {
    // OnStartupComplete é chamado na thread COM, não na thread de UI do Word.
    // A thread de monitoramento REA já iniciou em OnConnection e fará
    // os PostMessages para a thread de UI automaticamente.
    crate::tab_manager::log_debug("addin_on_startup_complete: OK (monitoramento REA ativo)");
    S_OK
}

unsafe extern "system" fn addin_on_begin_shutdown(
    this: *mut c_void,
    _custom: *mut *mut c_void,
) -> HRESULT {
    crate::tab_manager::log_debug("addin_on_begin_shutdown");
    let instance = unsafe { &*(this as *const AddinInstance) };
    if let Ok(mut lock) = instance.instancia.lock()
        && let Some(app) = lock.take()
    {
        app.shutdown();
    }
    S_OK
}

static ADDIN_VTBL: IDTExtensibility2_Vtbl = IDTExtensibility2_Vtbl {
    QueryInterface: addin_query_interface,
    AddRef: addin_add_ref,
    Release: addin_release,
    GetTypeInfoCount: addin_get_type_info_count,
    GetTypeInfo: addin_get_type_info,
    GetIDsOfNames: addin_get_ids_of_names,
    Invoke: addin_invoke,
    OnConnection: addin_on_connection,
    OnDisconnection: addin_on_disconnection,
    OnAddInsUpdate: addin_on_add_ins_update,
    OnStartupComplete: addin_on_startup_complete,
    OnBeginShutdown: addin_on_begin_shutdown,
};

#[implement(IClassFactory)]
struct FabricaSuplemento;

impl IClassFactory_Impl for FabricaSuplemento_Impl {
    fn CreateInstance(
        &self,
        externo: Option<&IUnknown>,
        riid: *const GUID,
        ppv: *mut *mut c_void,
    ) -> Result<()> {
        if ppv.is_null() || riid.is_null() {
            return Err(Error::from(E_NOINTERFACE));
        }

        unsafe { *ppv = std::ptr::null_mut() };

        if externo.is_some() {
            return Err(Error::from(CLASS_E_NOAGGREGATION));
        }

        OBJETOS_ATIVOS.fetch_add(1, Ordering::AcqRel);
        let instance = Box::new(AddinInstance {
            lp_vtbl: &ADDIN_VTBL,
            ref_count: AtomicU32::new(1),
            instancia: Mutex::new(None),
        });

        let raw = Box::into_raw(instance) as *mut c_void;
        let hr = unsafe { addin_query_interface(raw, riid, ppv) };
        unsafe { addin_release(raw) };

        if hr.0 == 0 {
            Ok(())
        } else {
            Err(Error::from(hr))
        }
    }

    fn LockServer(&self, bloquear: BOOL) -> Result<()> {
        if bloquear.as_bool() {
            BLOQUEIOS_SERVIDOR.fetch_add(1, Ordering::AcqRel);
        } else {
            let _ = BLOQUEIOS_SERVIDOR.try_update(Ordering::AcqRel, Ordering::Acquire, |total| {
                total.checked_sub(1)
            });
        }
        Ok(())
    }
}

fn consultar_interface<T: Interface>(
    objeto: &T,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> Result<()> {
    let desconhecido = objeto.cast::<IUnknown>()?;
    unsafe {
        (Interface::vtable(&desconhecido).QueryInterface)(desconhecido.as_raw(), riid, ppv).ok()
    }
}

/// Retorna a fábrica de classe solicitada pelo carregador COM.
///
/// # Safety
/// Os ponteiros `rclsid`, `riid` e `ppv` devem ser válidos ou nulos.
pub unsafe fn get_class_object(
    rclsid: *const GUID,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> Result<()> {
    if rclsid.is_null() || riid.is_null() || ppv.is_null() {
        return Err(Error::from(E_NOINTERFACE));
    }

    unsafe { *ppv = std::ptr::null_mut() };
    if unsafe { *rclsid } != ADDIN_CLSID {
        return Err(Error::from(E_NOINTERFACE));
    }

    let fabrica: IClassFactory = FabricaSuplemento.into();
    consultar_interface(&fabrica, riid, ppv)
}

/// Indica se o COM pode descarregar a DLL sem invalidar objetos vivos e garantindo cleanup de threads.
pub fn can_unload_now() -> bool {
    OBJETOS_ATIVOS.load(Ordering::Acquire) == 0
        && BLOQUEIOS_SERVIDOR.load(Ordering::Acquire) == 0
        && crate::tab_manager::pode_descarregar()
}
