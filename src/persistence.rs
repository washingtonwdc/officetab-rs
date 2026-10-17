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
//! Persistência de configuração no Registro do Windows e suporte a modo portátil.
//!
//! Inclui validação, migração de esquema de configuração, gerenciamento de
//! documentos favoritos e recentes, grupos de abas/sessões e criptografia leve (XOR/Base64).

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use windows::Win32::Foundation::WIN32_ERROR;
use windows::Win32::System::Registry::{
    HKEY, HKEY_CURRENT_USER, REG_DWORD, REG_OPEN_CREATE_OPTIONS, REG_SAM_FLAGS, REG_SZ,
    RRF_RT_DWORD, RRF_RT_REG_SZ, RegCloseKey, RegCreateKeyExW, RegDeleteTreeW, RegGetValueW,
    RegSetValueExW,
};
use windows::core::{PCWSTR, Result};

use crate::theme::{TabOrientation, TabStyle, Theme};

/// Versão corrente do esquema de configuração do OfficeTab.
pub const CURRENT_CONFIG_VERSION: u32 = 3;

/// Nome padrão do arquivo de configuração para modo portátil.
pub const PORTABLE_CONFIG_FILE: &str = "officetab_config.json";

const REG_ROOT: &str = r"Software\OfficeTabRs";
const REG_GENERAL: &str = "General";
#[allow(dead_code)]
const REG_TABS: &str = "Tabs";
const REG_THEME: &str = "Theme";
const REG_SHORTCUTS: &str = "Shortcuts";
const REG_FAVORITES: &str = "Favorites";
const REG_RECENT: &str = "Recent";
const REG_TAB_GROUPS: &str = "TabGroups";

fn registry_error(status: WIN32_ERROR) -> windows::core::Error {
    windows::core::HRESULT::from_win32(status.0).into()
}

/// Configuração geral do add-in
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct GeneralConfig {
    /// Altura da barra de abas em pixels
    pub tab_bar_height: i32,
    /// Habilitar multi-linha de abas
    pub multiline_tabs: bool,
    /// Posicionamento das abas (0=top, 1=bottom, 2=left, 3=right)
    pub tab_position: u32,
    /// Mostrar botão fechar nas abas
    pub show_close_button: bool,
    /// Habilitar rolagem de abas com mouse wheel
    pub mouse_wheel_scroll: bool,
    /// Habilitar arrastar-e-soltar para reordenar
    pub drag_drop_reorder: bool,
    /// Idioma da interface (código ISO 639-1)
    pub language: String,
    /// Verificar atualizações na inicialização
    pub check_updates: bool,
    /// Iniciar minimizado na bandeja
    pub start_minimized: bool,
    /// Selecionar automaticamente a paleta recomendada para cada produto Office.
    #[serde(default)]
    pub adaptive_theme_by_app: bool,
}

impl Default for GeneralConfig {
    fn default() -> Self {
        Self {
            tab_bar_height: 44,
            multiline_tabs: true,
            tab_position: 0, // Top
            show_close_button: true,
            mouse_wheel_scroll: true,
            drag_drop_reorder: true,
            language: "pt".to_string(),
            check_updates: true,
            start_minimized: false,
            adaptive_theme_by_app: false,
        }
    }
}

/// Configuração de atalhos de teclado
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct ShortcutsConfig {
    pub new_tab: String,
    pub close_tab: String,
    pub next_tab: String,
    pub prev_tab: String,
    pub toggle_tabs: String,
    pub save_all: String,
}

impl Default for ShortcutsConfig {
    fn default() -> Self {
        Self {
            new_tab: "Ctrl+T".to_string(),
            close_tab: "Ctrl+W".to_string(),
            next_tab: "Ctrl+Tab".to_string(),
            prev_tab: "Ctrl+Shift+Tab".to_string(),
            toggle_tabs: "Ctrl+Shift+T".to_string(),
            save_all: "Ctrl+Shift+S".to_string(),
        }
    }
}

/// Item de documento favorito enriquecido.
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct FavoriteItem {
    pub name: String,
    pub path: String,
    pub order: u32,
    #[serde(default)]
    pub icon: Option<String>,
}

/// Registro de documento aberto recentemente.
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct RecentDocument {
    pub title: String,
    pub path: String,
    pub last_opened: u64,
    pub pinned: bool,
}

/// Aba individual pertencente a um grupo/sessão salvo.
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct TabGroupItem {
    pub title: String,
    pub path: String,
    pub is_active: bool,
}

/// Grupo de abas preservando a sessão de trabalho do usuário.
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct TabGroup {
    pub name: String,
    pub items: Vec<TabGroupItem>,
    pub created_at: u64,
    #[serde(default)]
    pub application: Option<String>,
}

/// Configuração completa do OfficeTab
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct OfficeTabConfig {
    pub general: GeneralConfig,
    pub theme: Theme,
    pub shortcuts: ShortcutsConfig,
    /// Favoritos: mapa clássico de nome -> caminho para compatibilidade imediata
    #[serde(default)]
    pub favorites: HashMap<String, String>,
    /// Lista estruturada e ordenada de favoritos
    #[serde(default)]
    pub favorite_items: Vec<FavoriteItem>,
    /// Documentos recentes abertos pelo add-in
    #[serde(default)]
    pub recent_documents: Vec<RecentDocument>,
    /// Conjuntos de sessões/grupos de abas salvos
    #[serde(default)]
    pub tab_groups: HashMap<String, TabGroup>,
    /// Versão do schema de configuração (para migrações futuras)
    pub config_version: u32,
}

impl Default for OfficeTabConfig {
    fn default() -> Self {
        Self {
            general: GeneralConfig::default(),
            theme: Theme::default(),
            shortcuts: ShortcutsConfig::default(),
            favorites: HashMap::new(),
            favorite_items: Vec::new(),
            recent_documents: Vec::new(),
            tab_groups: HashMap::new(),
            config_version: CURRENT_CONFIG_VERSION,
        }
    }
}

impl OfficeTabConfig {
    // -----------------------------------------------------------------------
    // Gerenciamento de Favoritos
    // -----------------------------------------------------------------------

    /// Adiciona ou atualiza um item na lista de favoritos.
    pub fn add_favorite(&mut self, name: &str, path: &str) -> bool {
        if name.trim().is_empty() || path.trim().is_empty() {
            return false;
        }

        self.favorites.insert(name.to_string(), path.to_string());

        if let Some(pos) = self.favorite_items.iter().position(|f| f.name == name) {
            self.favorite_items[pos].path = path.to_string();
        } else {
            let next_order = self
                .favorite_items
                .iter()
                .map(|f| f.order)
                .max()
                .unwrap_or(0)
                + 1;
            self.favorite_items.push(FavoriteItem {
                name: name.to_string(),
                path: path.to_string(),
                order: next_order,
                icon: None,
            });
        }
        true
    }

    /// Remove um favorito pelo nome.
    pub fn remove_favorite(&mut self, name: &str) -> bool {
        let removed_map = self.favorites.remove(name).is_some();
        let initial_len = self.favorite_items.len();
        self.favorite_items.retain(|f| f.name != name);
        removed_map || (self.favorite_items.len() < initial_len)
    }

    /// Retorna todos os favoritos ordenados pelo campo order.
    pub fn list_favorites(&self) -> Vec<FavoriteItem> {
        let mut list = self.favorite_items.clone();
        list.sort_by_key(|f| f.order);
        list
    }

    /// Reordena os favoritos com base na ordem dos nomes fornecida.
    pub fn reorder_favorites(&mut self, ordered_names: &[&str]) {
        for (idx, &name) in ordered_names.iter().enumerate() {
            if let Some(item) = self.favorite_items.iter_mut().find(|f| f.name == name) {
                item.order = (idx + 1) as u32;
            }
        }
        self.favorite_items.sort_by_key(|f| f.order);
    }

    // -----------------------------------------------------------------------
    // Gerenciamento de Documentos Recentes
    // -----------------------------------------------------------------------

    /// Registra um documento recém-aberto, atualizando timestamp e posição.
    pub fn add_recent_document(&mut self, title: &str, path: &str) {
        if path.trim().is_empty() {
            return;
        }
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        // Se já existe, atualiza e move para o início
        if let Some(pos) = self.recent_documents.iter().position(|r| r.path == path) {
            self.recent_documents[pos].last_opened = now;
            self.recent_documents[pos].title = title.to_string();
            let item = self.recent_documents.remove(pos);
            self.recent_documents.insert(0, item);
        } else {
            self.recent_documents.insert(
                0,
                RecentDocument {
                    title: title.to_string(),
                    path: path.to_string(),
                    last_opened: now,
                    pinned: false,
                },
            );
        }

        // Limita o histórico a 50 itens sem remover os fixados
        if self.recent_documents.len() > 50
            && let Some(last_unpinned) = self.recent_documents.iter().rposition(|item| !item.pinned)
        {
            self.recent_documents.remove(last_unpinned);
        }
    }

    /// Remove um documento do histórico recente pelo caminho.
    pub fn remove_recent_document(&mut self, path: &str) -> bool {
        let initial_len = self.recent_documents.len();
        self.recent_documents.retain(|r| r.path != path);
        self.recent_documents.len() < initial_len
    }

    /// Lista os documentos recentes ordenados com fixados no topo.
    pub fn list_recent_documents(&self) -> &[RecentDocument] {
        &self.recent_documents
    }

    /// Alterna a fixação (pin) de um documento recente.
    pub fn toggle_pin_recent(&mut self, path: &str) -> bool {
        if let Some(item) = self.recent_documents.iter_mut().find(|r| r.path == path) {
            item.pinned = !item.pinned;
            return true;
        }
        false
    }

    /// Limpa documentos recentes não fixados.
    pub fn clear_recent_documents(&mut self) {
        self.recent_documents.retain(|r| r.pinned);
    }

    // -----------------------------------------------------------------------
    // Grupos de Abas e Restauração de Sessão
    // -----------------------------------------------------------------------

    /// Salva um grupo de abas atual.
    pub fn save_tab_group(&mut self, name: &str, items: Vec<TabGroupItem>, app: Option<&str>) {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        self.tab_groups.insert(
            name.to_string(),
            TabGroup {
                name: name.to_string(),
                items,
                created_at: now,
                application: app.map(|s| s.to_string()),
            },
        );
    }

    /// Obtém um grupo de abas pelo nome.
    pub fn get_tab_group(&self, name: &str) -> Option<&TabGroup> {
        self.tab_groups.get(name)
    }

    /// Remove um grupo de abas salvo.
    pub fn remove_tab_group(&mut self, name: &str) -> bool {
        self.tab_groups.remove(name).is_some()
    }

    /// Retorna todos os grupos de abas salvos.
    pub fn list_tab_groups(&self) -> Vec<&TabGroup> {
        self.tab_groups.values().collect()
    }

    /// Retorna a lista de itens de uma sessão para restauração.
    pub fn restore_session_tabs(&self, group_name: &str) -> Option<Vec<TabGroupItem>> {
        self.tab_groups.get(group_name).map(|g| g.items.clone())
    }
}

/// Valida a coerência da configuração contra limites aceitáveis.
pub fn validate_config(config: &OfficeTabConfig) -> std::result::Result<(), String> {
    if config.general.tab_bar_height < 16 || config.general.tab_bar_height > 96 {
        return Err(format!(
            "Altura da barra de abas inválida: {} (esperado entre 16 e 96)",
            config.general.tab_bar_height
        ));
    }

    if config.general.tab_position > 3 {
        return Err(format!(
            "Posição das abas inválida: {} (esperado entre 0 e 3)",
            config.general.tab_position
        ));
    }

    if config.theme.font_name.trim().is_empty() {
        return Err("O nome da fonte do tema não pode ser vazio".to_string());
    }

    if config.theme.font_size < 6 || config.theme.font_size > 36 {
        return Err(format!(
            "Tamanho de fonte inválido: {} (esperado entre 6 e 36)",
            config.theme.font_size
        ));
    }

    if config.theme.tab_min_width > config.theme.tab_max_width {
        return Err(format!(
            "tab_min_width ({}) não pode ser maior que tab_max_width ({})",
            config.theme.tab_min_width, config.theme.tab_max_width
        ));
    }

    if config.shortcuts.close_tab.trim().is_empty() || config.shortcuts.new_tab.trim().is_empty() {
        return Err("Atalhos essenciais não podem ser vazios".to_string());
    }

    Ok(())
}

/// Executa migração do esquema de configuração para a versão atual.
pub fn migrate_config(json_str: &str) -> std::result::Result<OfficeTabConfig, String> {
    let mut val: serde_json::Value =
        serde_json::from_str(json_str).map_err(|e| format!("JSON malformado: {}", e))?;

    let version = val
        .get("config_version")
        .and_then(|v| v.as_u64())
        .unwrap_or(0) as u32;

    if version < CURRENT_CONFIG_VERSION {
        // Se a versão for antiga (<= 1), preenche campos ausentes
        if let Some(obj) = val.as_object_mut() {
            obj.insert(
                "config_version".to_string(),
                serde_json::Value::from(CURRENT_CONFIG_VERSION),
            );

            if !obj.contains_key("favorite_items") {
                let mut favorite_items = Vec::new();
                if let Some(fav_map) = obj.get("favorites").and_then(|f| f.as_object()) {
                    let mut order = 1;
                    for (name, path_val) in fav_map {
                        if let Some(path_str) = path_val.as_str() {
                            favorite_items.push(serde_json::json!({
                                "name": name,
                                "path": path_str,
                                "order": order,
                                "icon": null
                            }));
                            order += 1;
                        }
                    }
                }
                obj.insert(
                    "favorite_items".to_string(),
                    serde_json::Value::Array(favorite_items),
                );
            }

            if !obj.contains_key("recent_documents") {
                obj.insert(
                    "recent_documents".to_string(),
                    serde_json::Value::Array(Vec::new()),
                );
            }

            if !obj.contains_key("tab_groups") {
                obj.insert(
                    "tab_groups".to_string(),
                    serde_json::Value::Object(serde_json::Map::new()),
                );
            }
        }
    }

    let config: OfficeTabConfig = serde_json::from_value(val)
        .map_err(|e| format!("Falha ao converter estrutura deserializada: {}", e))?;

    validate_config(&config)?;
    Ok(config)
}

// ---------------------------------------------------------------------------
// Suporte a Modo Portátil
// ---------------------------------------------------------------------------

/// Retorna o caminho onde reside o arquivo de configuração portátil.
pub fn get_portable_config_path() -> PathBuf {
    if let Ok(module_path) = crate::utils::get_current_module_path()
        && let Some(dir) = module_path.parent()
    {
        return dir.join(PORTABLE_CONFIG_FILE);
    }
    std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|d| d.join(PORTABLE_CONFIG_FILE)))
        .unwrap_or_else(|| PathBuf::from(PORTABLE_CONFIG_FILE))
}

/// Indica se o add-in está operando em modo portátil.
pub fn is_portable_mode() -> bool {
    get_portable_config_path().exists()
}

/// Carrega a configuração vigente (priorizando arquivo portátil se presente, senão Registro).
pub fn load_effective_config() -> OfficeTabConfig {
    let port_path = get_portable_config_path();
    if port_path.exists()
        && let Ok(config) = import_config_from_file(&port_path)
    {
        return config;
    }
    load_config()
}

/// Salva a configuração (no arquivo portátil se estiver em modo portátil, senão no Registro).
pub fn save_effective_config(config: &OfficeTabConfig) -> Result<()> {
    validate_config(config).map_err(|error| {
        windows::core::Error::new(windows::core::HRESULT(0x80070057u32 as i32), error)
    })?;
    if is_portable_mode() {
        let port_path = get_portable_config_path();
        export_config_to_file(config, &port_path)
    } else {
        save_config(config)
    }
}

// ---------------------------------------------------------------------------
// Criptografia leve e ofuscação (XOR + Base64)
// ---------------------------------------------------------------------------

const BASE64_ALPHABET: &[u8; 64] =
    b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/// Codifica bytes em string Base64 padrão.
fn base64_encode(data: &[u8]) -> String {
    let mut result = String::with_capacity(data.len().div_ceil(3) * 4);
    let mut i = 0;
    while i < data.len() {
        let b0 = data[i] as u32;
        let b1 = if i + 1 < data.len() {
            data[i + 1] as u32
        } else {
            0
        };
        let b2 = if i + 2 < data.len() {
            data[i + 2] as u32
        } else {
            0
        };

        let triple = (b0 << 16) | (b1 << 8) | b2;

        result.push(BASE64_ALPHABET[((triple >> 18) & 0x3F) as usize] as char);
        result.push(BASE64_ALPHABET[((triple >> 12) & 0x3F) as usize] as char);

        if i + 1 < data.len() {
            result.push(BASE64_ALPHABET[((triple >> 6) & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }

        if i + 2 < data.len() {
            result.push(BASE64_ALPHABET[(triple & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }

        i += 3;
    }
    result
}

/// Decodifica string Base64 para vetor de bytes.
fn base64_decode(encoded: &str) -> std::result::Result<Vec<u8>, String> {
    let clean: Vec<u8> = encoded
        .bytes()
        .filter(|&b| b != b'\r' && b != b'\n' && b != b' ')
        .collect();
    if !clean.len().is_multiple_of(4) {
        return Err("Comprimento inválido para string Base64".to_string());
    }

    let decode_byte = |b: u8| -> std::result::Result<u8, String> {
        match b {
            b'A'..=b'Z' => Ok(b - b'A'),
            b'a'..=b'z' => Ok(b - b'a' + 26),
            b'0'..=b'9' => Ok(b - b'0' + 52),
            b'+' => Ok(62),
            b'/' => Ok(63),
            _ => Err(format!("Caractere Base64 inválido: {}", b as char)),
        }
    };

    let mut result = Vec::with_capacity(clean.len() / 4 * 3);
    for chunk in clean.chunks(4) {
        let b0 = decode_byte(chunk[0])? as u32;
        let b1 = decode_byte(chunk[1])? as u32;

        if chunk[2] == b'=' {
            if chunk[3] != b'=' {
                return Err("Padding inválido em Base64".to_string());
            }
            let triple = (b0 << 18) | (b1 << 12);
            result.push(((triple >> 16) & 0xFF) as u8);
        } else if chunk[3] == b'=' {
            let b2 = decode_byte(chunk[2])? as u32;
            let triple = (b0 << 18) | (b1 << 12) | (b2 << 6);
            result.push(((triple >> 16) & 0xFF) as u8);
            result.push(((triple >> 8) & 0xFF) as u8);
        } else {
            let b2 = decode_byte(chunk[2])? as u32;
            let b3 = decode_byte(chunk[3])? as u32;
            let triple = (b0 << 18) | (b1 << 12) | (b2 << 6) | b3;
            result.push(((triple >> 16) & 0xFF) as u8);
            result.push(((triple >> 8) & 0xFF) as u8);
            result.push((triple & 0xFF) as u8);
        }
    }
    Ok(result)
}

/// Aplica ofuscação de bytes via XOR rotativo e codifica em Base64.
pub fn obfuscate_data(data: &[u8], key: &[u8]) -> String {
    if key.is_empty() {
        return base64_encode(data);
    }
    let mut xored = Vec::with_capacity(data.len());
    for (i, &byte) in data.iter().enumerate() {
        let k = key[i % key.len()];
        let rot = ((i * 7) & 0xFF) as u8;
        xored.push(byte ^ k ^ rot);
    }
    base64_encode(&xored)
}

/// Decodifica e desofusca dados previamente codificados com `obfuscate_data`.
pub fn deobfuscate_data(encoded: &str, key: &[u8]) -> std::result::Result<Vec<u8>, String> {
    let xored = base64_decode(encoded)?;
    if key.is_empty() {
        return Ok(xored);
    }
    let mut original = Vec::with_capacity(xored.len());
    for (i, &byte) in xored.iter().enumerate() {
        let k = key[i % key.len()];
        let rot = ((i * 7) & 0xFF) as u8;
        original.push(byte ^ k ^ rot);
    }
    Ok(original)
}

/// Ofusca uma string de credencial ou chave para armazenamento local seguro.
pub fn obfuscate_string(plain: &str, key: &[u8]) -> String {
    obfuscate_data(plain.as_bytes(), key)
}

/// Restaura uma string a partir de sua representação ofuscada.
pub fn deobfuscate_string(encoded: &str, key: &[u8]) -> std::result::Result<String, String> {
    let bytes = deobfuscate_data(encoded, key)?;
    String::from_utf8(bytes).map_err(|e| format!("Texto UTF-8 inválido: {}", e))
}

// ---------------------------------------------------------------------------
// Operações no Registro do Windows
// ---------------------------------------------------------------------------

/// Chave de registro base
fn get_reg_key(path: &str) -> String {
    format!(r"{}\{}", REG_ROOT, path)
}

/// Cria ou abre uma chave de registro
fn create_key(hkey_root: HKEY, subkey: &str) -> Result<HKEY> {
    unsafe {
        let subkey_utf16: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
        let mut hkey_result = HKEY::default();

        let status = RegCreateKeyExW(
            hkey_root,
            PCWSTR(subkey_utf16.as_ptr()),
            0,
            None,
            REG_OPEN_CREATE_OPTIONS(0),
            REG_SAM_FLAGS(0x0002001F), // KEY_READ | KEY_WRITE
            None,
            &mut hkey_result,
            None,
        );

        if status != WIN32_ERROR(0) {
            return Err(registry_error(status));
        }

        Ok(hkey_result)
    }
}

/// Define um valor string no registro
fn set_string_value(hkey: HKEY, name: Option<&str>, value: &str) -> Result<()> {
    unsafe {
        let name_utf16: Option<Vec<u16>> =
            name.map(|n| n.encode_utf16().chain(std::iter::once(0)).collect());
        let name_ptr = name_utf16
            .as_ref()
            .map(|v| PCWSTR(v.as_ptr()))
            .unwrap_or_else(PCWSTR::null);

        let mut val_utf16: Vec<u16> = value.encode_utf16().chain(std::iter::once(0)).collect();
        let bytes_len = val_utf16.len() * 2;

        let status = RegSetValueExW(
            hkey,
            name_ptr,
            0,
            REG_SZ,
            Some(std::slice::from_raw_parts(
                val_utf16.as_mut_ptr() as *const u8,
                bytes_len,
            )),
        );

        if status != WIN32_ERROR(0) {
            return Err(registry_error(status));
        }

        Ok(())
    }
}

/// Define um valor DWORD no registro
fn set_dword_value(hkey: HKEY, name: &str, value: u32) -> Result<()> {
    unsafe {
        let name_utf16: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        let bytes = value.to_ne_bytes();

        let status = RegSetValueExW(
            hkey,
            PCWSTR(name_utf16.as_ptr()),
            0,
            REG_DWORD,
            Some(&bytes),
        );

        if status != WIN32_ERROR(0) {
            return Err(registry_error(status));
        }

        Ok(())
    }
}

/// Lê um valor string do registro
fn get_string_value(hkey: HKEY, name: &str) -> Result<String> {
    unsafe {
        let name_utf16: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        let mut buffer = vec![0u16; 1024];
        let mut size = (buffer.len() * 2) as u32;

        let status = RegGetValueW(
            hkey,
            PCWSTR::null(),
            PCWSTR(name_utf16.as_ptr()),
            RRF_RT_REG_SZ,
            None,
            Some(buffer.as_mut_ptr() as *mut _),
            Some(&mut size),
        );

        if status != WIN32_ERROR(0) {
            return Err(registry_error(status));
        }

        let len = (size as usize / 2).min(buffer.len());
        buffer.truncate(len);
        Ok(String::from_utf16_lossy(&buffer)
            .trim_end_matches('\0')
            .to_string())
    }
}

/// Lê um valor DWORD do registro
fn get_dword_value(hkey: HKEY, name: &str) -> Result<u32> {
    unsafe {
        let name_utf16: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        let mut value: u32 = 0;
        let mut size = 4u32;

        let status = RegGetValueW(
            hkey,
            PCWSTR::null(),
            PCWSTR(name_utf16.as_ptr()),
            RRF_RT_DWORD,
            None,
            Some(&mut value as *mut _ as *mut _),
            Some(&mut size),
        );

        if status != WIN32_ERROR(0) {
            return Err(registry_error(status));
        }

        Ok(value)
    }
}

/// Salva a configuração completa no registro
pub fn save_config(config: &OfficeTabConfig) -> Result<()> {
    validate_config(config).map_err(|error| {
        windows::core::Error::new(windows::core::HRESULT(0x80070057u32 as i32), error)
    })?;
    // General
    let k_general = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_GENERAL))?;
    let g = &config.general;
    set_dword_value(k_general, "TabBarHeight", g.tab_bar_height as u32)?;
    set_dword_value(k_general, "MultilineTabs", g.multiline_tabs as u32)?;
    set_dword_value(k_general, "TabPosition", g.tab_position)?;
    set_dword_value(k_general, "ShowCloseButton", g.show_close_button as u32)?;
    set_dword_value(k_general, "MouseWheelScroll", g.mouse_wheel_scroll as u32)?;
    set_dword_value(k_general, "DragDropReorder", g.drag_drop_reorder as u32)?;
    set_string_value(k_general, Some("Language"), &g.language)?;
    set_dword_value(k_general, "CheckUpdates", g.check_updates as u32)?;
    set_dword_value(k_general, "StartMinimized", g.start_minimized as u32)?;
    set_dword_value(
        k_general,
        "AdaptiveThemeByApp",
        g.adaptive_theme_by_app as u32,
    )?;
    set_dword_value(k_general, "ConfigVersion", config.config_version)?;
    unsafe {
        let _ = RegCloseKey(k_general);
    }

    // Theme
    let k_theme = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_THEME))?;
    let t = &config.theme;
    set_string_value(k_theme, Some("Name"), &t.name)?;
    set_dword_value(k_theme, "ActiveTabBg", t.active_tab_bg)?;
    set_dword_value(k_theme, "ActiveTabText", t.active_tab_text)?;
    set_dword_value(k_theme, "InactiveTabBg", t.inactive_tab_bg)?;
    set_dword_value(k_theme, "InactiveTabText", t.inactive_tab_text)?;
    set_dword_value(k_theme, "HoverTabBg", t.hover_tab_bg)?;
    set_dword_value(k_theme, "HoverTabText", t.hover_tab_text)?;
    set_dword_value(k_theme, "TabBorder", t.tab_border)?;
    set_dword_value(k_theme, "BarBackground", t.bar_background)?;
    set_string_value(k_theme, Some("FontName"), &t.font_name)?;
    set_dword_value(k_theme, "FontSize", t.font_size)?;
    set_dword_value(k_theme, "TabHeight", t.tab_height as u32)?;
    set_dword_value(k_theme, "TabMinWidth", t.tab_min_width as u32)?;
    set_dword_value(k_theme, "TabMaxWidth", t.tab_max_width as u32)?;
    set_dword_value(k_theme, "TabPadding", t.tab_padding as u32)?;
    set_dword_value(k_theme, "Style", t.style.to_u32())?;
    set_dword_value(k_theme, "Orientation", t.orientation.to_u32())?;
    set_dword_value(k_theme, "Multiline", t.multiline as u32)?;
    set_dword_value(k_theme, "ShowCloseButton", t.show_close_button as u32)?;
    set_dword_value(k_theme, "ModernTheme", t.modern_theme as u32)?;
    unsafe {
        let _ = RegCloseKey(k_theme);
    }

    // Shortcuts
    let k_shortcuts = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_SHORTCUTS))?;
    let s = &config.shortcuts;
    set_string_value(k_shortcuts, Some("NewTab"), &s.new_tab)?;
    set_string_value(k_shortcuts, Some("CloseTab"), &s.close_tab)?;
    set_string_value(k_shortcuts, Some("NextTab"), &s.next_tab)?;
    set_string_value(k_shortcuts, Some("PrevTab"), &s.prev_tab)?;
    set_string_value(k_shortcuts, Some("ToggleTabs"), &s.toggle_tabs)?;
    set_string_value(k_shortcuts, Some("SaveAll"), &s.save_all)?;
    unsafe {
        let _ = RegCloseKey(k_shortcuts);
    }

    // Favorites
    let k_fav = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_FAVORITES))?;
    for (name, path) in &config.favorites {
        set_string_value(k_fav, Some(name), path)?;
    }
    if let Ok(json_fav) = serde_json::to_string(&config.favorite_items) {
        set_string_value(k_fav, Some("__ItemsJson__"), &json_fav)?;
    }
    unsafe {
        let _ = RegCloseKey(k_fav);
    }

    // Recent Documents
    if let Ok(k_rec) = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_RECENT)) {
        if let Ok(json_rec) = serde_json::to_string(&config.recent_documents) {
            let _ = set_string_value(k_rec, Some("RecentList"), &json_rec);
        }
        unsafe {
            let _ = RegCloseKey(k_rec);
        }
    }

    // Tab Groups
    if let Ok(k_groups) = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_TAB_GROUPS)) {
        if let Ok(json_groups) = serde_json::to_string(&config.tab_groups) {
            let _ = set_string_value(k_groups, Some("GroupsList"), &json_groups);
        }
        unsafe {
            let _ = RegCloseKey(k_groups);
        }
    }

    Ok(())
}

/// Carrega a configuração do registro (ou retorna padrão se não existir)
pub fn load_config() -> OfficeTabConfig {
    let mut config = OfficeTabConfig::default();

    // General
    if let Ok(k_general) = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_GENERAL)) {
        let g = &mut config.general;
        if let Ok(v) = get_dword_value(k_general, "TabBarHeight") {
            g.tab_bar_height = v as i32;
        }
        if let Ok(v) = get_dword_value(k_general, "MultilineTabs") {
            g.multiline_tabs = v != 0;
        }
        if let Ok(v) = get_dword_value(k_general, "TabPosition") {
            g.tab_position = v;
        }
        if let Ok(v) = get_dword_value(k_general, "ShowCloseButton") {
            g.show_close_button = v != 0;
        }
        if let Ok(v) = get_dword_value(k_general, "MouseWheelScroll") {
            g.mouse_wheel_scroll = v != 0;
        }
        if let Ok(v) = get_dword_value(k_general, "DragDropReorder") {
            g.drag_drop_reorder = v != 0;
        }
        if let Ok(v) = get_string_value(k_general, "Language") {
            g.language = v;
        }
        if let Ok(v) = get_dword_value(k_general, "CheckUpdates") {
            g.check_updates = v != 0;
        }
        if let Ok(v) = get_dword_value(k_general, "StartMinimized") {
            g.start_minimized = v != 0;
        }
        if let Ok(v) = get_dword_value(k_general, "AdaptiveThemeByApp") {
            g.adaptive_theme_by_app = v != 0;
        }
        if let Ok(v) = get_dword_value(k_general, "ConfigVersion") {
            config.config_version = v;
        }
        unsafe {
            let _ = RegCloseKey(k_general);
        }
    }

    // Theme
    if let Ok(k_theme) = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_THEME)) {
        let t = &mut config.theme;
        if let Ok(v) = get_string_value(k_theme, "Name") {
            t.name = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "ActiveTabBg") {
            t.active_tab_bg = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "ActiveTabText") {
            t.active_tab_text = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "InactiveTabBg") {
            t.inactive_tab_bg = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "InactiveTabText") {
            t.inactive_tab_text = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "HoverTabBg") {
            t.hover_tab_bg = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "HoverTabText") {
            t.hover_tab_text = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "TabBorder") {
            t.tab_border = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "BarBackground") {
            t.bar_background = v;
        }
        if let Ok(v) = get_string_value(k_theme, "FontName") {
            t.font_name = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "FontSize") {
            t.font_size = v;
        }
        if let Ok(v) = get_dword_value(k_theme, "TabHeight") {
            t.tab_height = v as i32;
        }
        if let Ok(v) = get_dword_value(k_theme, "TabMinWidth") {
            t.tab_min_width = v as i32;
        }
        if let Ok(v) = get_dword_value(k_theme, "TabMaxWidth") {
            t.tab_max_width = v as i32;
        }
        if let Ok(v) = get_dword_value(k_theme, "TabPadding") {
            t.tab_padding = v as i32;
        }
        if let Ok(v) = get_dword_value(k_theme, "Style") {
            t.style = TabStyle::from_u32(v);
        }
        if let Ok(v) = get_dword_value(k_theme, "Orientation") {
            t.orientation = TabOrientation::from_u32(v);
        }
        if let Ok(v) = get_dword_value(k_theme, "Multiline") {
            t.multiline = v != 0;
        }
        if let Ok(v) = get_dword_value(k_theme, "ShowCloseButton") {
            t.show_close_button = v != 0;
        }
        if let Ok(v) = get_dword_value(k_theme, "ModernTheme") {
            t.modern_theme = v != 0;
        }
        unsafe {
            let _ = RegCloseKey(k_theme);
        }
    }

    // Shortcuts
    if let Ok(k_shortcuts) = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_SHORTCUTS)) {
        let s = &mut config.shortcuts;
        if let Ok(v) = get_string_value(k_shortcuts, "NewTab") {
            s.new_tab = v;
        }
        if let Ok(v) = get_string_value(k_shortcuts, "CloseTab") {
            s.close_tab = v;
        }
        if let Ok(v) = get_string_value(k_shortcuts, "NextTab") {
            s.next_tab = v;
        }
        if let Ok(v) = get_string_value(k_shortcuts, "PrevTab") {
            s.prev_tab = v;
        }
        if let Ok(v) = get_string_value(k_shortcuts, "ToggleTabs") {
            s.toggle_tabs = v;
        }
        if let Ok(v) = get_string_value(k_shortcuts, "SaveAll") {
            s.save_all = v;
        }
        unsafe {
            let _ = RegCloseKey(k_shortcuts);
        }
    }

    // Favorites
    if let Ok(k_fav) = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_FAVORITES)) {
        if let Ok(json_fav) = get_string_value(k_fav, "__ItemsJson__")
            && let Ok(items) = serde_json::from_str::<Vec<FavoriteItem>>(&json_fav)
        {
            config.favorite_items = items;
        }
        unsafe {
            let _ = RegCloseKey(k_fav);
        }
    }

    // Recent Documents
    if let Ok(k_rec) = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_RECENT)) {
        if let Ok(json_rec) = get_string_value(k_rec, "RecentList")
            && let Ok(recent) = serde_json::from_str::<Vec<RecentDocument>>(&json_rec)
        {
            config.recent_documents = recent;
        }
        unsafe {
            let _ = RegCloseKey(k_rec);
        }
    }

    // Tab Groups
    if let Ok(k_groups) = create_key(HKEY_CURRENT_USER, &get_reg_key(REG_TAB_GROUPS)) {
        if let Ok(json_groups) = get_string_value(k_groups, "GroupsList")
            && let Ok(groups) = serde_json::from_str::<HashMap<String, TabGroup>>(&json_groups)
        {
            config.tab_groups = groups;
        }
        unsafe {
            let _ = RegCloseKey(k_groups);
        }
    }

    config
}

/// Exporta configuração para arquivo JSON
pub fn export_config_to_file(config: &OfficeTabConfig, path: &Path) -> Result<()> {
    validate_config(config).map_err(|error| {
        windows::core::Error::new(windows::core::HRESULT(0x80070057u32 as i32), error)
    })?;
    let json = serde_json::to_string_pretty(config).map_err(|e| {
        windows::core::Error::new(
            windows::core::HRESULT(0x80004005u32 as i32),
            format!("JSON error: {}", e),
        )
    })?;
    std::fs::write(path, json).map_err(|e| {
        windows::core::Error::new(
            windows::core::HRESULT(0x80004005u32 as i32),
            format!("IO error: {}", e),
        )
    })?;
    Ok(())
}

/// Importa configuração de arquivo JSON com validação e migração transparente.
pub fn import_config_from_file(path: &Path) -> Result<OfficeTabConfig> {
    let json = std::fs::read_to_string(path).map_err(|e| {
        windows::core::Error::new(
            windows::core::HRESULT(0x80004005u32 as i32),
            format!("IO error: {}", e),
        )
    })?;
    migrate_config(&json).map_err(|e| {
        windows::core::Error::new(
            windows::core::HRESULT(0x80004005u32 as i32),
            format!("Migration/Validation error: {}", e),
        )
    })
}

/// Exporta os favoritos para um arquivo JSON independente da configuração principal.
pub fn export_favorites_to_file(config: &OfficeTabConfig, path: &Path) -> Result<()> {
    let favorites = if config.favorite_items.is_empty() {
        let mut items: Vec<_> = config
            .favorites
            .iter()
            .map(|(name, path)| FavoriteItem {
                name: name.clone(),
                path: path.clone(),
                order: 0,
                icon: None,
            })
            .collect();
        items.sort_by(|a, b| a.name.cmp(&b.name));
        items
    } else {
        config.list_favorites()
    };
    let json = serde_json::to_string_pretty(&favorites).map_err(|error| {
        windows::core::Error::new(
            windows::core::HRESULT(0x80004005u32 as i32),
            format!("JSON error: {error}"),
        )
    })?;
    std::fs::write(path, json).map_err(|error| {
        windows::core::Error::new(
            windows::core::HRESULT(0x80004005u32 as i32),
            format!("IO error: {error}"),
        )
    })
}

/// Importa e valida um backup de favoritos, substituindo ambas as listas da configuração.
pub fn import_favorites_from_file(config: &mut OfficeTabConfig, path: &Path) -> Result<()> {
    let json = std::fs::read_to_string(path).map_err(|error| {
        windows::core::Error::new(
            windows::core::HRESULT(0x80004005u32 as i32),
            format!("IO error: {error}"),
        )
    })?;
    let items: Vec<FavoriteItem> = serde_json::from_str(&json).map_err(|error| {
        windows::core::Error::new(
            windows::core::HRESULT(0x80004005u32 as i32),
            format!("Favorites JSON error: {error}"),
        )
    })?;

    let mut favorites = HashMap::with_capacity(items.len());
    for item in &items {
        if item.name.trim().is_empty() || item.path.trim().is_empty() {
            return Err(windows::core::Error::new(
                windows::core::HRESULT(0x80070057u32 as i32),
                "Favorite names and paths must not be empty",
            ));
        }
        if favorites
            .insert(item.name.clone(), item.path.clone())
            .is_some()
        {
            return Err(windows::core::Error::new(
                windows::core::HRESULT(0x80070057u32 as i32),
                format!("Duplicate favorite name: {}", item.name),
            ));
        }
    }

    config.favorites = favorites;
    config.favorite_items = items;
    Ok(())
}

/// Remove toda a configuração do registro (desinstalação limpa)
pub fn clear_config() -> Result<()> {
    unsafe {
        let sub = format!("{}\0", REG_ROOT);
        let utf16: Vec<u16> = sub.encode_utf16().collect();
        let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(utf16.as_ptr()));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_config() {
        let cfg = OfficeTabConfig::default();
        assert_eq!(cfg.general.tab_bar_height, 44);
        assert!(cfg.general.multiline_tabs);
        assert_eq!(cfg.general.language, "pt");
        assert!(!cfg.general.adaptive_theme_by_app);
        assert_eq!(cfg.config_version, CURRENT_CONFIG_VERSION);
        assert_eq!(cfg.shortcuts.new_tab, "Ctrl+T");
        assert!(cfg.favorite_items.is_empty());
        assert!(cfg.recent_documents.is_empty());
        assert!(cfg.tab_groups.is_empty());
    }

    #[test]
    fn test_export_and_import_json() {
        let mut original = OfficeTabConfig::default();
        original.general.language = "en".to_string();
        original.general.tab_bar_height = 32;
        original.shortcuts.save_all = "Ctrl+Alt+S".to_string();
        original.add_favorite("ProjPlan", "C:\\Docs\\Proj.docx");

        let temp_dir = std::env::temp_dir();
        let temp_file = temp_dir.join("officetab_test_config.json");

        assert!(export_config_to_file(&original, &temp_file).is_ok());
        let loaded = import_config_from_file(&temp_file).expect("Falha ao importar config");
        assert_eq!(loaded.general.tab_bar_height, 32);
        assert_eq!(loaded.general.language, "en");
        assert_eq!(loaded.favorite_items.len(), 1);
        assert_eq!(loaded.favorite_items[0].name, "ProjPlan");

        let _ = std::fs::remove_file(temp_file);
    }

    #[test]
    fn test_favorites_backup_roundtrip_and_validation() {
        let mut source = OfficeTabConfig::default();
        source.add_favorite("Relatorio", "C:\\Docs\\Relatorio.xlsx");
        source.add_favorite("Projeto", "D:\\Projetos\\Plano.mpp");
        let path =
            std::env::temp_dir().join(format!("officetab_favorites_{}.json", std::process::id()));

        export_favorites_to_file(&source, &path).unwrap();
        let mut restored = OfficeTabConfig::default();
        import_favorites_from_file(&mut restored, &path).unwrap();
        assert_eq!(restored.list_favorites(), source.list_favorites());
        assert_eq!(restored.favorites.len(), 2);

        std::fs::write(&path, r#"[{"name":"","path":"x","order":1}]"#).unwrap();
        assert!(import_favorites_from_file(&mut restored, &path).is_err());
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn test_validation_and_migration() {
        let mut cfg = OfficeTabConfig::default();
        assert!(validate_config(&cfg).is_ok());

        cfg.general.tab_bar_height = 10; // Muito baixo
        assert!(validate_config(&cfg).is_err());

        cfg.general.tab_bar_height = 28;
        cfg.theme.font_size = 50; // Muito alto
        assert!(validate_config(&cfg).is_err());

        // Teste de migração a partir de versão legada (v1)
        let legacy_v1_json = r#"{
            "general": {
                "tab_bar_height": 30,
                "multiline_tabs": false,
                "tab_position": 0,
                "show_close_button": true,
                "mouse_wheel_scroll": true,
                "drag_drop_reorder": true,
                "language": "pt",
                "check_updates": true,
                "start_minimized": false
            },
            "theme": {
                "name": "Light",
                "active_tab_bg": 16777215,
                "active_tab_text": 3355443,
                "inactive_tab_bg": 15790320,
                "inactive_tab_text": 6710886,
                "hover_tab_bg": 15066597,
                "hover_tab_text": 3355443,
                "tab_border": 13684944,
                "bar_background": 16053493,
                "font_name": "Segoe UI",
                "font_size": 9,
                "tab_height": 28,
                "tab_min_width": 70,
                "tab_max_width": 255,
                "tab_padding": 8,
                "style": "Chrome",
                "orientation": "Top",
                "multiline": false,
                "show_close_button": true,
                "modern_theme": true
            },
            "shortcuts": {
                "new_tab": "Ctrl+T",
                "close_tab": "Ctrl+W",
                "next_tab": "Ctrl+Tab",
                "prev_tab": "Ctrl+Shift+Tab",
                "toggle_tabs": "Ctrl+Shift+T",
                "save_all": "Ctrl+Shift+S"
            },
            "favorites": {
                "DocAntigo": "C:\\Docs\\Old.docx"
            },
            "config_version": 1
        }"#;

        let migrated = migrate_config(legacy_v1_json).expect("Falha ao migrar JSON v1");
        assert_eq!(migrated.config_version, CURRENT_CONFIG_VERSION);
        assert_eq!(migrated.favorite_items.len(), 1);
        assert_eq!(migrated.favorite_items[0].name, "DocAntigo");
        assert_eq!(migrated.favorite_items[0].path, "C:\\Docs\\Old.docx");
    }

    #[test]
    fn test_favorites_management() {
        let mut cfg = OfficeTabConfig::default();
        assert!(cfg.add_favorite("Vendas", "C:\\Vendas.xlsx"));
        assert!(cfg.add_favorite("RH", "C:\\RH.docx"));
        assert!(cfg.add_favorite("TI", "C:\\TI.pptx"));

        assert_eq!(cfg.list_favorites().len(), 3);
        assert_eq!(cfg.list_favorites()[0].name, "Vendas");

        // Reordenar
        cfg.reorder_favorites(&["TI", "Vendas", "RH"]);
        let ordered = cfg.list_favorites();
        assert_eq!(ordered[0].name, "TI");
        assert_eq!(ordered[1].name, "Vendas");
        assert_eq!(ordered[2].name, "RH");

        // Remover
        assert!(cfg.remove_favorite("Vendas"));
        assert_eq!(cfg.list_favorites().len(), 2);
    }

    #[test]
    fn test_recent_documents() {
        let mut cfg = OfficeTabConfig::default();
        cfg.add_recent_document("Doc 1", "C:\\Doc1.docx");
        cfg.add_recent_document("Doc 2", "C:\\Doc2.docx");

        assert_eq!(cfg.list_recent_documents().len(), 2);
        assert_eq!(cfg.list_recent_documents()[0].title, "Doc 2");

        // Fixar e limpar
        assert!(cfg.toggle_pin_recent("C:\\Doc1.docx"));
        cfg.clear_recent_documents();
        assert_eq!(cfg.list_recent_documents().len(), 1);
        assert_eq!(cfg.list_recent_documents()[0].title, "Doc 1");

        assert!(cfg.remove_recent_document("C:\\Doc1.docx"));
        assert_eq!(cfg.list_recent_documents().len(), 0);
    }

    #[test]
    fn test_tab_groups_and_session_restore() {
        let mut cfg = OfficeTabConfig::default();
        let items = vec![
            TabGroupItem {
                title: "Orcamento.xlsx".to_string(),
                path: "C:\\Orcamento.xlsx".to_string(),
                is_active: true,
            },
            TabGroupItem {
                title: "Apresentacao.pptx".to_string(),
                path: "C:\\Apresentacao.pptx".to_string(),
                is_active: false,
            },
        ];

        cfg.save_tab_group("ProjetoAlpha", items.clone(), Some("Office"));
        assert!(cfg.get_tab_group("ProjetoAlpha").is_some());
        assert_eq!(cfg.list_tab_groups().len(), 1);

        let restored = cfg
            .restore_session_tabs("ProjetoAlpha")
            .expect("Grupo não encontrado");
        assert_eq!(restored.len(), 2);
        assert_eq!(restored[0].title, "Orcamento.xlsx");

        assert!(cfg.remove_tab_group("ProjetoAlpha"));
        assert!(cfg.get_tab_group("ProjetoAlpha").is_none());
    }

    #[test]
    fn test_obfuscation_and_deobfuscation() {
        let key = b"minha-chave-secreta-officetab";
        let original_text = "CHAVE-LICENCA-1234-ABCD-EFGH";

        let obfuscated = obfuscate_string(original_text, key);
        assert_ne!(obfuscated, original_text);

        let restored = deobfuscate_string(&obfuscated, key).expect("Falha ao desofuscar string");
        assert_eq!(restored, original_text);

        let wrong_key = b"outra-chave-incorreta";
        let failed = deobfuscate_string(&obfuscated, wrong_key);
        // O resultado desofuscado com chave errada não coincidirá
        assert!(failed.is_err() || failed.unwrap() != original_text);
    }

    #[test]
    fn test_save_and_load_registry() {
        let mut custom = OfficeTabConfig::default();
        custom.general.tab_bar_height = 36;
        custom.general.language = "es".to_string();
        custom.general.adaptive_theme_by_app = true;
        custom.theme = Theme::dark();
        custom.add_favorite("Planilha", "C:\\Planilha.xlsx");

        // Salvar no registro
        assert!(save_config(&custom).is_ok());

        // Carregar do registro
        let loaded = load_config();
        assert_eq!(loaded.general.tab_bar_height, 36);
        assert_eq!(loaded.general.language, "es");
        assert!(loaded.general.adaptive_theme_by_app);
        assert_eq!(loaded.theme.name, "Dark");
        assert_eq!(loaded.theme.active_tab_bg, custom.theme.active_tab_bg);
        assert_eq!(loaded.theme.style, custom.theme.style);
        assert_eq!(loaded.theme.orientation, custom.theme.orientation);
        assert_eq!(loaded.favorite_items.len(), 1);
        assert_eq!(loaded.favorite_items[0].name, "Planilha");

        // Limpeza
        assert!(clear_config().is_ok());
    }
}
