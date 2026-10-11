//! Implementacao do servidor COM interno do suplemento.
#![allow(non_snake_case)]

use std::ffi::c_void;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::BOOL;
use windows::Win32::System::Com::{IClassFactory, IClassFactory_Impl};
use windows::core::{
    Error, GUID, HRESULT, IUnknown, IUnknown_Vtbl, Interface, Result, implement, interface,
};

/// Identificador estável da classe COM do suplemento.
pub const ADDIN_CLSID: GUID = GUID::from_u128(0x1c8ba3cd_5248_4611_8a6e_8ccbaf89b651);
pub const ADDIN_PROG_ID: &str = "OfficeTabs.Connect";

const E_NOINTERFACE: HRESULT = HRESULT(0x80004002u32 as i32);
const CLASS_E_NOAGGREGATION: HRESULT = HRESULT(0x80040110u32 as i32);

static OBJETOS_ATIVOS: AtomicU32 = AtomicU32::new(0);
static BLOQUEIOS_SERVIDOR: AtomicU32 = AtomicU32::new(0);

/// Interface publicada pelo Microsoft Office para suplementos COM.
#[interface("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
unsafe trait IDTExtensibility2: IUnknown {
    unsafe fn OnConnection(
        &self,
        application: Option<&IUnknown>,
        modo_conexao: i32,
        instancia_suplemento: Option<&IUnknown>,
        personalizado: *const *const c_void,
    ) -> HRESULT;
    fn OnDisconnection(&self, modo_desconexao: i32, personalizado: *const *const c_void)
    -> HRESULT;
    fn OnAddInsUpdate(&self, personalizado: *const *const c_void) -> HRESULT;
    fn OnStartupComplete(&self, personalizado: *const *const c_void) -> HRESULT;
    fn OnBeginShutdown(&self, personalizado: *const *const c_void) -> HRESULT;
}

#[implement(IDTExtensibility2)]
struct ObjetoSuplemento {
    instancia: std::sync::Mutex<Option<crate::implementation::OfficeTab>>,
}

impl ObjetoSuplemento {
    fn novo() -> Self {
        OBJETOS_ATIVOS.fetch_add(1, Ordering::AcqRel);
        Self {
            instancia: std::sync::Mutex::new(None),
        }
    }
}

impl Drop for ObjetoSuplemento {
    fn drop(&mut self) {
        OBJETOS_ATIVOS.fetch_sub(1, Ordering::AcqRel);
    }
}

impl IDTExtensibility2_Impl for ObjetoSuplemento_Impl {
    unsafe fn OnConnection(
        &self,
        _application: Option<&IUnknown>,
        _modo_conexao: i32,
        _instancia_suplemento: Option<&IUnknown>,
        _personalizado: *const *const c_void,
    ) -> HRESULT {
        let app = crate::implementation::OfficeTab::new();
        let _ = app.initialize();
        if let Ok(mut lock) = self.instancia.lock() {
            *lock = Some(app);
        }
        HRESULT(0)
    }

    unsafe fn OnDisconnection(
        &self,
        _modo_desconexao: i32,
        _personalizado: *const *const c_void,
    ) -> HRESULT {
        if let Ok(mut lock) = self.instancia.lock()
            && let Some(app) = lock.take()
        {
            app.shutdown();
        }
        HRESULT(0)
    }

    unsafe fn OnAddInsUpdate(&self, _personalizado: *const *const c_void) -> HRESULT {
        HRESULT(0)
    }

    unsafe fn OnStartupComplete(&self, _personalizado: *const *const c_void) -> HRESULT {
        if let Ok(lock) = self.instancia.lock()
            && let Some(ref app) = *lock
        {
            let _ = app.attach_to_active_office();
        }
        HRESULT(0)
    }

    unsafe fn OnBeginShutdown(&self, _personalizado: *const *const c_void) -> HRESULT {
        if let Ok(mut lock) = self.instancia.lock()
            && let Some(app) = lock.take()
        {
            app.shutdown();
        }
        HRESULT(0)
    }
}

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

        let objeto: IDTExtensibility2 = ObjetoSuplemento::novo().into();
        consultar_interface(&objeto, riid, ppv)
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

/// Indica se o COM pode descarregar a DLL sem invalidar objetos vivos.
pub fn can_unload_now() -> bool {
    OBJETOS_ATIVOS.load(Ordering::Acquire) == 0 && BLOQUEIOS_SERVIDOR.load(Ordering::Acquire) == 0
}
