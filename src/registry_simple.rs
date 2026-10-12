// Registro COM e integração com Microsoft Office no Registro do Windows.
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
//! Registro COM e integração com Microsoft Office no Registro do Windows.

use std::path::PathBuf;
use windows::Win32::Foundation::WIN32_ERROR;
use windows::Win32::System::Registry::{
    HKEY, HKEY_CURRENT_USER, REG_DWORD, REG_OPEN_CREATE_OPTIONS, REG_SAM_FLAGS, REG_SZ,
    RegCloseKey, RegCreateKeyExW, RegDeleteTreeW, RegSetValueExW,
};
use windows::core::{PCWSTR, Result};

const ADDIN_PROG_ID: &str = crate::com_addin_simple::ADDIN_PROG_ID;
const ADDIN_CLSID_STR: &str = "{9a1929ef-1495-413e-993c-326dd15e35df}";
const ADDIN_DESC: &str = "Office Tab Rust";
const LOAD_BEHAVIOR: u32 = 3; // 3 = Carregar na inicialização do Office

const OFFICE_APPS: &[&str] = &["Word", "Excel", "PowerPoint"];

fn criar_chave(hkey_raiz: HKEY, subkey: &str) -> Result<HKEY> {
    unsafe {
        let subkey_utf16: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
        let mut hkey_resultado = HKEY::default();

        let status = RegCreateKeyExW(
            hkey_raiz,
            PCWSTR(subkey_utf16.as_ptr()),
            0,
            None,
            REG_OPEN_CREATE_OPTIONS(0),
            REG_SAM_FLAGS(0x00020006), // KEY_WRITE | KEY_READ
            None,
            &mut hkey_resultado,
            None,
        );

        if status != WIN32_ERROR(0) {
            return Err(windows::core::Error::from_win32());
        }

        Ok(hkey_resultado)
    }
}

fn definir_valor_string(hkey: HKEY, nome: Option<&str>, valor: &str) -> Result<()> {
    unsafe {
        let nome_utf16: Option<Vec<u16>> =
            nome.map(|n| n.encode_utf16().chain(std::iter::once(0)).collect());
        let nome_ptr = nome_utf16
            .as_ref()
            .map(|v| PCWSTR(v.as_ptr()))
            .unwrap_or_else(PCWSTR::null);

        let mut val_utf16: Vec<u16> = valor.encode_utf16().chain(std::iter::once(0)).collect();
        let bytes_len = val_utf16.len() * 2;

        let status = RegSetValueExW(
            hkey,
            nome_ptr,
            0,
            REG_SZ,
            Some(std::slice::from_raw_parts(
                val_utf16.as_mut_ptr() as *const u8,
                bytes_len,
            )),
        );

        if status != WIN32_ERROR(0) {
            return Err(windows::core::Error::from_win32());
        }

        Ok(())
    }
}

fn definir_valor_dword(hkey: HKEY, nome: &str, valor: u32) -> Result<()> {
    unsafe {
        let nome_utf16: Vec<u16> = nome.encode_utf16().chain(std::iter::once(0)).collect();
        let bytes = valor.to_ne_bytes();

        let status = RegSetValueExW(
            hkey,
            PCWSTR(nome_utf16.as_ptr()),
            0,
            REG_DWORD,
            Some(&bytes),
        );

        if status != WIN32_ERROR(0) {
            return Err(windows::core::Error::from_win32());
        }

        Ok(())
    }
}

/// Registra o servidor COM no escopo do usuário atual (HKCU).
pub fn register_server() -> Result<()> {
    let dll_path = crate::utils::get_current_module_path()
        .unwrap_or_else(|_| PathBuf::from("officetab_rs.dll"));
    let dll_path_str = dll_path.to_string_lossy().to_string();

    // 1. Registrar ProgID em HKCU\Software\Classes\OfficeTabs.Connect
    let k_progid = criar_chave(
        HKEY_CURRENT_USER,
        &format!("Software\\Classes\\{}", ADDIN_PROG_ID),
    )?;
    definir_valor_string(k_progid, None, ADDIN_DESC)?;
    unsafe {
        let _ = RegCloseKey(k_progid);
    }

    let k_progid_clsid = criar_chave(
        HKEY_CURRENT_USER,
        &format!("Software\\Classes\\{}\\CLSID", ADDIN_PROG_ID),
    )?;
    definir_valor_string(k_progid_clsid, None, ADDIN_CLSID_STR)?;
    unsafe {
        let _ = RegCloseKey(k_progid_clsid);
    }

    // 2. Registrar CLSID em HKCU\Software\Classes\CLSID\{GUID}
    let k_clsid = criar_chave(
        HKEY_CURRENT_USER,
        &format!("Software\\Classes\\CLSID\\{}", ADDIN_CLSID_STR),
    )?;
    definir_valor_string(k_clsid, None, ADDIN_DESC)?;
    definir_valor_string(k_clsid, Some("ProgID"), ADDIN_PROG_ID)?;
    unsafe {
        let _ = RegCloseKey(k_clsid);
    }

    let k_inproc = criar_chave(
        HKEY_CURRENT_USER,
        &format!(
            "Software\\Classes\\CLSID\\{}\\InprocServer32",
            ADDIN_CLSID_STR
        ),
    )?;
    definir_valor_string(k_inproc, None, &dll_path_str)?;
    definir_valor_string(k_inproc, Some("ThreadingModel"), "Apartment")?;
    unsafe {
        let _ = RegCloseKey(k_inproc);
    }

    // 3. Registrar Add-in para Word, Excel e PowerPoint
    for app in OFFICE_APPS {
        let subkey = format!(
            "Software\\Microsoft\\Office\\{}\\Addins\\{}",
            app, ADDIN_PROG_ID
        );
        let k_addin = criar_chave(HKEY_CURRENT_USER, &subkey)?;
        definir_valor_string(k_addin, Some("FriendlyName"), ADDIN_DESC)?;
        definir_valor_string(k_addin, Some("Description"), ADDIN_DESC)?;
        definir_valor_dword(k_addin, "LoadBehavior", LOAD_BEHAVIOR)?;
        definir_valor_dword(k_addin, "CommandLineSafe", 0)?;
        unsafe {
            let _ = RegCloseKey(k_addin);
        }
    }

    Ok(())
}

/// Remove as entradas do servidor COM e suplemento do Registro do Windows.
pub fn unregister_server() -> Result<()> {
    unsafe {
        // Remover ProgID
        let progid_sub = format!("Software\\Classes\\{}\0", ADDIN_PROG_ID);
        let utf16: Vec<u16> = progid_sub.encode_utf16().collect();
        let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(utf16.as_ptr()));

        // Remover CLSID
        let clsid_sub = format!("Software\\Classes\\CLSID\\{}\0", ADDIN_CLSID_STR);
        let utf16: Vec<u16> = clsid_sub.encode_utf16().collect();
        let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(utf16.as_ptr()));

        // Remover dos aplicativos Office
        for app in OFFICE_APPS {
            let addin_sub = format!(
                "Software\\Microsoft\\Office\\{}\\Addins\\{}\0",
                app, ADDIN_PROG_ID
            );
            let utf16: Vec<u16> = addin_sub.encode_utf16().collect();
            let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(utf16.as_ptr()));
        }
    }

    Ok(())
}
