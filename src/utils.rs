// Utilitários Win32 e suporte ao ambiente do suplemento Office Tab.
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
//! Utilitários Win32 e suporte ao ambiente do suplemento Office Tab.

use std::path::PathBuf;
use windows::Win32::Foundation::{HMODULE, HWND};
use windows::Win32::System::LibraryLoader::{
    GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
    GetModuleFileNameW, GetModuleHandleExW, GetModuleHandleW, GetProcAddress,
};
use windows::core::{PCWSTR, Result, s, w};

/// A inicialização COM pertence ao processo Office que carregou esta DLL.
/// Mantida por compatibilidade com a API pública, sem alterar a contagem COM do host.
pub fn init_com() -> Result<()> {
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
            let dpi = func(hwnd);
            if dpi > 0 {
                return dpi;
            }
        }
        96 // 96 DPI padrão (escala de 100%)
    }
}

/// Define a conscientização de DPI por monitor v2 para o processo.
pub fn set_dpi_awareness() -> Result<()> {
    // Este código roda dentro de WINWORD.EXE/EXCEL.EXE/etc.; alterar a
    // conscientização do processo afetaria o Office e todos os seus add-ins.
    // O host é responsável por definir o contexto DPI. Nossas dimensões são
    // escaladas por janela em get_dpi_for_window().
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

/// Copia uma sequência de texto UTF-16 para a área de transferência do Windows.
pub fn copiar_para_clipboard(texto: &str) -> Result<()> {
    unsafe {
        use windows::Win32::Foundation::GlobalFree;
        use windows::Win32::System::DataExchange::{
            CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData,
        };
        use windows::Win32::System::Memory::{
            GLOBAL_ALLOC_FLAGS, GlobalAlloc, GlobalLock, GlobalUnlock,
        };

        OpenClipboard(HWND::default())?;
        if let Err(error) = EmptyClipboard() {
            let _ = CloseClipboard();
            return Err(error);
        }

        let utf16 = texto_para_clipboard_utf16(texto);
        let bytes_len = utf16.len() * std::mem::size_of::<u16>();

        // GMEM_MOVEABLE = 0x0002
        let hmem = match GlobalAlloc(GLOBAL_ALLOC_FLAGS(0x0002), bytes_len) {
            Ok(h) => h,
            Err(e) => {
                let _ = CloseClipboard();
                return Err(e);
            }
        };

        if hmem.is_invalid() {
            let _ = CloseClipboard();
            return Err(windows::core::Error::from_win32());
        }

        let ptr = GlobalLock(hmem);
        if ptr.is_null() {
            let _ = GlobalFree(hmem);
            let _ = CloseClipboard();
            return Err(windows::core::Error::from_win32());
        }
        std::ptr::copy_nonoverlapping(utf16.as_ptr() as *const u8, ptr as *mut u8, bytes_len);
        let _ = GlobalUnlock(hmem);

        // CF_UNICODETEXT = 13
        let res = SetClipboardData(13, windows::Win32::Foundation::HANDLE(hmem.0));
        let _ = CloseClipboard();

        if let Err(error) = res {
            let _ = GlobalFree(hmem);
            return Err(error);
        }

        Ok(())
    }
}

fn texto_para_clipboard_utf16(texto: &str) -> Vec<u16> {
    texto.encode_utf16().chain(std::iter::once(0)).collect()
}

/// Obtém o diretório da pasta "Documentos" do usuário atual.
pub fn obter_pasta_documentos_usuario() -> PathBuf {
    if let Ok(perfil) = std::env::var("USERPROFILE") {
        let doc = PathBuf::from(perfil).join("Documents");
        if doc.exists() {
            return doc;
        }
    }
    std::env::current_dir().unwrap_or_else(|_| PathBuf::from("C:\\"))
}

/// Tenta resolver o caminho absoluto de um documento com base no título exibido na aba.
pub fn localizar_caminho_documento(titulo: &str) -> Option<PathBuf> {
    let titulo_limpo = titulo.trim();
    if titulo_limpo.is_empty() {
        return None;
    }

    // 1. O próprio título já é um caminho absoluto existente
    let p = std::path::Path::new(titulo_limpo);
    if p.is_absolute() && p.exists() {
        return Some(p.to_path_buf());
    }

    // 2. Se for um arquivo com extensão conhecida, tenta localizar em pastas comuns
    let extensoes = [
        ".docx", ".doc", ".docm", ".dotx", ".xlsx", ".xls", ".xlsm", ".xlsb", ".pptx", ".ppt",
        ".pptm", ".accdb", ".mdb", ".mpp", ".mpt", ".vsdx", ".vsd", ".pub", ".rtf", ".txt",
    ];
    let possui_ext = extensoes
        .iter()
        .any(|ext| titulo_limpo.to_lowercase().ends_with(ext));

    let pastas_busca = [
        obter_pasta_documentos_usuario(),
        std::env::var("USERPROFILE")
            .map(|p| PathBuf::from(p).join("Desktop"))
            .unwrap_or_default(),
        std::env::var("USERPROFILE")
            .map(|p| PathBuf::from(p).join("OneDrive"))
            .unwrap_or_default(),
        std::env::var("USERPROFILE")
            .map(|p| PathBuf::from(p).join("OneDrive").join("Documentos"))
            .unwrap_or_default(),
        std::env::current_dir().unwrap_or_default(),
    ];

    for pasta in pastas_busca {
        if !pasta.as_os_str().is_empty() && pasta.exists() {
            let candidato = pasta.join(titulo_limpo);
            if candidato.is_file() {
                return Some(candidato);
            }

            // Se não possuir extensão explícita no título, testa as extensões padrão do Office
            if !possui_ext {
                for ext in &extensoes {
                    let com_ext = pasta.join(format!("{}{}", titulo_limpo, ext));
                    if com_ext.is_file() {
                        return Some(com_ext);
                    }
                }
            }
        }
    }

    None
}

/// Abre o Windows Explorer exibindo ou selecionando o caminho especificado.
pub fn abrir_pasta_no_explorer(caminho: &std::path::Path) -> Result<()> {
    if caminho.is_file() {
        std::process::Command::new("explorer")
            .arg(format!("/select,\"{}\"", caminho.display()))
            .spawn()?;
    } else if caminho.is_dir() {
        std::process::Command::new("explorer")
            .arg(caminho.as_os_str())
            .spawn()?;
    } else {
        let doc_dir = obter_pasta_documentos_usuario();
        std::process::Command::new("explorer")
            .arg(doc_dir.as_os_str())
            .spawn()?;
    }
    Ok(())
}

#[cfg(test)]
mod testes {
    use super::*;

    #[test]
    fn test_obter_pasta_documentos_usuario_valida() {
        let pasta = obter_pasta_documentos_usuario();
        assert!(pasta.exists());
    }

    #[test]
    fn test_localizar_caminho_documento_inexistente_ou_vazio() {
        assert!(localizar_caminho_documento("").is_none());
        assert!(localizar_caminho_documento("   ").is_none());
        assert!(localizar_caminho_documento("ArquivoNaoExistente_xyz987654.docx").is_none());
    }

    #[test]
    fn test_localizar_caminho_documento_absoluto_existente() {
        let cargo_toml = std::env::current_dir().unwrap().join("Cargo.toml");
        let resultado = localizar_caminho_documento(&cargo_toml.display().to_string());
        assert_eq!(resultado, Some(cargo_toml));
    }

    #[test]
    fn test_texto_clipboard_termina_em_nul_e_preserva_unicode() {
        let data = texto_para_clipboard_utf16("Abação 📄");
        assert_eq!(
            String::from_utf16(&data[..data.len() - 1]).unwrap(),
            "Abação 📄"
        );
        assert_eq!(data.last(), Some(&0));
    }
}
