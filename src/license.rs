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
//! Validação de licença offline autônoma (100% offline, sem telemetria).
//!
//! Suporta chaves no formato empresarial `XXXX-XXXX-XXXX-XXXX`, gerador de chaves
//! offline para administração, cálculo e persistência de datas OLE Automation Date,
//! e detecção inteligente de retrocesso no relógio do sistema.

use serde::{Deserialize, Serialize};
use std::time::{SystemTime, UNIX_EPOCH};

use windows::Win32::Foundation::WIN32_ERROR;
use windows::Win32::System::Registry::{
    HKEY, HKEY_CURRENT_USER, REG_DWORD, REG_OPEN_CREATE_OPTIONS, REG_SAM_FLAGS, REG_SZ,
    RRF_RT_DWORD, RRF_RT_REG_SZ, RegCloseKey, RegCreateKeyExW, RegGetValueW, RegSetValueExW,
};
use windows::core::{PCWSTR, Result};

/// Chave padrão de registro do OfficeTab para persistência de licença
const REG_LICENSE_PATH: &str = r"Software\OfficeTabRs\License";

fn registry_error(status: WIN32_ERROR) -> windows::core::Error {
    windows::core::HRESULT::from_win32(status.0).into()
}

/// Duração do trial em segundos (30 dias).
pub const TRIAL_DURATION_SECS: u64 = 30 * 24 * 60 * 60;

/// Constante de conversão: dias entre 30/12/1899 e 01/01/1970 (Unix Epoch em OLE Date).
pub const OLE_DATE_UNIX_EPOCH_OFFSET: f64 = 25569.0;

/// Segredo padrão embutido para verificação HMAC offline.
pub const DEFAULT_LICENSE_SECRET: &[u8] = b"OfficeTabRs-Offline-Validation-Secret-Key-2024";

/// Converte timestamp Unix (segundos desde 1970) para formato OLE Automation Date (`f64`).
pub fn unix_to_ole_date(timestamp: u64) -> f64 {
    OLE_DATE_UNIX_EPOCH_OFFSET + (timestamp as f64) / 86400.0
}

/// Converte OLE Automation Date (`f64`) para timestamp Unix (segundos desde 1970).
pub fn ole_date_to_unix(ole_date: f64) -> u64 {
    if ole_date < OLE_DATE_UNIX_EPOCH_OFFSET {
        0
    } else {
        ((ole_date - OLE_DATE_UNIX_EPOCH_OFFSET) * 86400.0).round() as u64
    }
}

/// Estado da licença do add-in.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum LicenseState {
    /// Versão gratuita (sem restrições no clone open source)
    #[default]
    FreeEdition,
    /// Trial ativo (30 dias)
    Trial { days_remaining: u32 },
    /// Licença registrada válida
    Registered { licensee: String, quantity: u32 },
    /// Licença expirada
    Expired { days_past: u32 },
    /// Relógio do sistema adulterado para o passado
    ClockTampered,
}

/// Informações de licença completas com persistência de OLE Date e timestamps.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LicenseInfo {
    pub state: LicenseState,
    pub install_timestamp: u64,
    pub check_timestamp: u64,
    #[serde(default)]
    pub install_ole_date: f64,
    #[serde(default)]
    pub check_ole_date: f64,
    pub license_key: Option<String>,
    pub licensee_name: Option<String>,
    pub quantity: u32,
    #[serde(default)]
    pub clock_tampered: bool,
}

impl Default for LicenseInfo {
    fn default() -> Self {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        let ole = unix_to_ole_date(now);
        Self {
            state: LicenseState::FreeEdition,
            install_timestamp: now,
            check_timestamp: now,
            install_ole_date: ole,
            check_ole_date: ole,
            license_key: None,
            licensee_name: None,
            quantity: 0,
            clock_tampered: false,
        }
    }
}

impl LicenseInfo {
    /// Cria uma nova instância de licença em modo trial de 30 dias.
    pub fn new_trial() -> Self {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        let ole = unix_to_ole_date(now);
        Self {
            state: LicenseState::Trial { days_remaining: 30 },
            install_timestamp: now,
            check_timestamp: now,
            install_ole_date: ole,
            check_ole_date: ole,
            license_key: None,
            licensee_name: None,
            quantity: 1,
            clock_tampered: false,
        }
    }

    /// Verifica e atualiza o estado da licença com base no tempo decorrido e integridade do relógio.
    pub fn refresh_state(&mut self) {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        // Detecção de retrocesso no relógio do sistema (Anti-Tamper)
        if now + 60 < self.check_timestamp {
            self.clock_tampered = true;
            self.state = LicenseState::ClockTampered;
            return;
        }

        self.check_timestamp = now;
        self.check_ole_date = unix_to_ole_date(now);

        match &self.state {
            LicenseState::FreeEdition | LicenseState::Registered { .. } => {
                // Modo ilimitado
            }
            LicenseState::Trial { .. } => {
                let elapsed = now.saturating_sub(self.install_timestamp);
                if elapsed >= TRIAL_DURATION_SECS {
                    let days_past = ((elapsed - TRIAL_DURATION_SECS) / 86400) as u32;
                    self.state = LicenseState::Expired { days_past };
                } else {
                    let remaining = (TRIAL_DURATION_SECS - elapsed) / 86400;
                    self.state = LicenseState::Trial {
                        days_remaining: (remaining as u32) + 1,
                    };
                }
            }
            LicenseState::Expired { .. } => {
                let elapsed = now.saturating_sub(self.install_timestamp);
                let days_past = (elapsed.saturating_sub(TRIAL_DURATION_SECS) / 86400) as u32;
                self.state = LicenseState::Expired { days_past };
            }
            LicenseState::ClockTampered => {
                // Permanece adulterado
            }
        }
    }

    /// Retorna `true` se o add-in deve funcionar (trial ativo, registrado ou free).
    pub fn is_active(&self) -> bool {
        matches!(
            self.state,
            LicenseState::FreeEdition
                | LicenseState::Trial { .. }
                | LicenseState::Registered { .. }
        )
    }

    /// Retorna `true` se deve exibir alerta ou lembrete de aquisição/registro.
    pub fn should_reminder(&self) -> bool {
        match &self.state {
            LicenseState::Trial { days_remaining } => *days_remaining <= 7,
            LicenseState::Expired { .. } | LicenseState::ClockTampered => true,
            _ => false,
        }
    }

    /// Serializa para JSON.
    pub fn to_json(&self) -> std::result::Result<String, serde_json::Error> {
        serde_json::to_string(self)
    }

    /// Deserializa a partir de JSON.
    pub fn from_json(json: &str) -> std::result::Result<Self, serde_json::Error> {
        serde_json::from_str(json)
    }

    /// Valida uma chave no formato unificado (suporta formato clássico e formato `XXXX-XXXX-XXXX-XXXX`).
    pub fn validate_key(key: &str, secret: &[u8]) -> std::result::Result<(String, u32), String> {
        let clean = key.trim();
        if clean.contains('#') {
            Self::validate_legacy_key(clean, secret)
        } else {
            Err(
                "Para chaves no formato XXXX-XXXX-XXXX-XXXX, informe o nome do licenciado usando validate_offline_key"
                    .to_string(),
            )
        }
    }

    /// Valida uma chave legado no formato `<NOME>#<QUANTIDADE>#<TIMESTAMP_HEX>#<HMAC_HEX>`.
    fn validate_legacy_key(key: &str, secret: &[u8]) -> std::result::Result<(String, u32), String> {
        let parts: Vec<&str> = key.split('#').collect();
        if parts.len() != 4 {
            return Err("Formato de chave inválido".to_string());
        }

        let name = parts[0].to_string();
        let qty: u32 = parts[1]
            .parse()
            .map_err(|_| "Quantidade inválida na chave".to_string())?;

        let ts = u64::from_str_radix(parts[2], 16)
            .map_err(|_| "Timestamp inválido na chave".to_string())?;
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        let two_years = 2 * 365 * 24 * 60 * 60;
        if now.saturating_sub(ts) > two_years {
            return Err("Chave expirada".to_string());
        }

        let payload = format!("{}#{}#{}", name, qty, parts[2]);
        let expected_hmac = simple_hmac(payload.as_bytes(), secret);
        let provided_hmac = parts[3].to_lowercase();

        if expected_hmac != provided_hmac {
            return Err("Assinatura da chave inválida".to_string());
        }

        Ok((name, qty))
    }

    /// Registra uma chave no formato clássico com delimitador `#`.
    pub fn register(&mut self, key: &str, secret: &[u8]) -> std::result::Result<(), String> {
        let (name, qty) = Self::validate_key(key, secret)?;
        self.apply_registration(&name, qty, key);
        Ok(())
    }

    /// Registra uma chave no formato empresarial `XXXX-XXXX-XXXX-XXXX` com nome do usuário.
    pub fn register_offline(
        &mut self,
        key: &str,
        licensee: &str,
        secret: &[u8],
    ) -> std::result::Result<(), String> {
        let qty = validate_offline_key(key, licensee, secret)?;
        self.apply_registration(licensee, qty, key);
        Ok(())
    }

    /// Aplica os parâmetros de registro na estrutura.
    fn apply_registration(&mut self, licensee: &str, quantity: u32, key: &str) {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        self.state = LicenseState::Registered {
            licensee: licensee.to_string(),
            quantity,
        };
        self.license_key = Some(key.to_string());
        self.licensee_name = Some(licensee.to_string());
        self.quantity = quantity;
        self.check_timestamp = now;
        self.check_ole_date = unix_to_ole_date(now);
        self.clock_tampered = false;
    }
}

// ---------------------------------------------------------------------------
// Validação e Geração de Chaves Offline Formato XXXX-XXXX-XXXX-XXXX
// ---------------------------------------------------------------------------

const KEY_ALPHABET: &[u8; 32] = b"23456789ABCDEFGHJKLMNPQRSTUVWXYZ"; // Base32 legível (sem 0, 1, I, O)

/// Codifica um valor u16 em 4 caracteres do alfabeto Base32 seguro.
fn encode_u16_to_4chars(val: u16) -> String {
    let mut s = String::with_capacity(4);
    let mut v = val as usize;
    for _ in 0..4 {
        s.push(KEY_ALPHABET[v % KEY_ALPHABET.len()] as char);
        v /= KEY_ALPHABET.len();
    }
    s
}

/// Decodifica 4 caracteres para um u16.
fn decode_4chars_to_u16(chars: &str) -> std::result::Result<u16, String> {
    if chars.len() != 4 {
        return Err("Bloco deve ter exatamente 4 caracteres".to_string());
    }
    let mut multiplier: usize = 1;
    let mut total: usize = 0;
    for b in chars.bytes() {
        let pos = KEY_ALPHABET
            .iter()
            .position(|&c| c == b)
            .ok_or_else(|| format!("Caractere inválido na chave: {}", b as char))?;
        total += pos * multiplier;
        multiplier *= KEY_ALPHABET.len();
    }
    Ok((total & 0xFFFF) as u16)
}

/// Calcula hash de 32 bits (FNV-1a modificado) com segredo.
fn fnv32_hmac(data: &[u8], secret: &[u8]) -> u32 {
    let mut hash: u32 = 0x811c9dc5;
    for (i, &b) in data.iter().enumerate() {
        let s = if !secret.is_empty() {
            secret[i % secret.len()]
        } else {
            0
        };
        hash ^= (b ^ s) as u32;
        hash = hash.wrapping_mul(0x01000193);
    }
    hash
}

/// Gera uma chave offline no formato empresarial `XXXX-XXXX-XXXX-XXXX`.
///
/// O formato é 100% offline e verifica a legitimidade do licenciado,
/// quantidade de assentos e integridade com a chave secreta.
pub fn generate_offline_key(licensee: &str, quantity: u32, secret: &[u8]) -> String {
    let clean_name = licensee.trim().to_uppercase();
    let qty_capped = (quantity.clamp(1, 9999)) as u16;

    // Bloco 1: Quantidade com salt simples
    let b1_val = qty_capped ^ 0xA55A;
    let block1 = encode_u16_to_4chars(b1_val);

    // Bloco 2: Hash do licenciado + quantidade
    let h1 = fnv32_hmac(format!("{}#{}", clean_name, qty_capped).as_bytes(), secret);
    let block2 = encode_u16_to_4chars((h1 & 0xFFFF) as u16);

    // Bloco 3: Hash do segredo + bloco 1 + bloco 2
    let h2 = fnv32_hmac(format!("{}{}", block1, block2).as_bytes(), secret);
    let block3 = encode_u16_to_4chars((h2 & 0xFFFF) as u16);

    // Bloco 4: Checksum global de integridade
    let h3 = fnv32_hmac(
        format!("{}{}{}{}", block1, block2, block3, clean_name).as_bytes(),
        secret,
    );
    let block4 = encode_u16_to_4chars(((h3 >> 16) ^ (h3 & 0xFFFF)) as u16);

    format!("{}-{}-{}-{}", block1, block2, block3, block4)
}

/// Valida uma chave no formato `XXXX-XXXX-XXXX-XXXX` para um determinado licenciado.
/// Retorna a quantidade de assentos autorizados se válida.
pub fn validate_offline_key(
    key: &str,
    licensee: &str,
    secret: &[u8],
) -> std::result::Result<u32, String> {
    let clean_key = key.trim().to_uppercase();
    let parts: Vec<&str> = clean_key.split('-').collect();
    if parts.len() != 4 || parts.iter().any(|p| p.len() != 4) {
        return Err("A chave deve estar no formato XXXX-XXXX-XXXX-XXXX".to_string());
    }

    let clean_name = licensee.trim().to_uppercase();
    if clean_name.is_empty() {
        return Err("O nome do licenciado não pode ser vazio".to_string());
    }

    // Extrair quantidade do Bloco 1
    let b1_raw = decode_4chars_to_u16(parts[0])?;
    let qty = (b1_raw ^ 0xA55A) as u32;
    if qty == 0 || qty > 9999 {
        return Err("Quantidade de licenças inválida na chave".to_string());
    }

    // Regenerar e verificar chave esperada
    let expected_key = generate_offline_key(&clean_name, qty, secret);
    if expected_key != clean_key {
        return Err("Chave de produto inválida para o usuário informado".to_string());
    }

    Ok(qty)
}

// ---------------------------------------------------------------------------
// Persistência no Registro do Windows
// ---------------------------------------------------------------------------

fn create_reg_key(subkey: &str) -> Result<HKEY> {
    unsafe {
        let subkey_utf16: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
        let mut hkey_result = HKEY::default();

        let status = RegCreateKeyExW(
            HKEY_CURRENT_USER,
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

fn set_reg_str(hkey: HKEY, name: &str, value: &str) -> Result<()> {
    unsafe {
        let name_utf16: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        let mut val_utf16: Vec<u16> = value.encode_utf16().chain(std::iter::once(0)).collect();
        let bytes_len = val_utf16.len() * 2;

        let status = RegSetValueExW(
            hkey,
            PCWSTR(name_utf16.as_ptr()),
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

fn set_reg_dword(hkey: HKEY, name: &str, value: u32) -> Result<()> {
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

fn get_reg_str(hkey: HKEY, name: &str) -> Result<String> {
    unsafe {
        let name_utf16: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        let mut buffer = vec![0u16; 512];
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

fn get_reg_dword(hkey: HKEY, name: &str) -> Result<u32> {
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

/// Salva o estado de licença no Registro do Windows.
pub fn save_license_to_registry(info: &LicenseInfo) -> Result<()> {
    let k = create_reg_key(REG_LICENSE_PATH)?;
    set_reg_dword(
        k,
        "InstallTimeLow",
        (info.install_timestamp & 0xFFFFFFFF) as u32,
    )?;
    set_reg_dword(
        k,
        "InstallTimeHigh",
        ((info.install_timestamp >> 32) & 0xFFFFFFFF) as u32,
    )?;
    set_reg_dword(
        k,
        "CheckTimeLow",
        (info.check_timestamp & 0xFFFFFFFF) as u32,
    )?;
    set_reg_dword(
        k,
        "CheckTimeHigh",
        ((info.check_timestamp >> 32) & 0xFFFFFFFF) as u32,
    )?;
    set_reg_dword(k, "Quantity", info.quantity)?;
    set_reg_dword(k, "ClockTampered", info.clock_tampered as u32)?;

    if let Some(key) = &info.license_key {
        let _ = set_reg_str(k, "Key", key);
    }
    if let Some(name) = &info.licensee_name {
        let _ = set_reg_str(k, "Licensee", name);
    }

    if let Ok(json_state) = serde_json::to_string(&info.state) {
        let _ = set_reg_str(k, "StateJson", &json_state);
    }

    unsafe {
        let _ = RegCloseKey(k);
    }
    Ok(())
}

/// Carrega o estado de licença do Registro do Windows (ou inicializa trial se inexistente).
pub fn load_license_from_registry() -> LicenseInfo {
    if let Ok(k) = create_reg_key(REG_LICENSE_PATH) {
        let mut info = LicenseInfo::new_trial();
        let mut has_data = false;

        let install_low = get_reg_dword(k, "InstallTimeLow");
        let install_high = get_reg_dword(k, "InstallTimeHigh");
        if let (Ok(low), Ok(high)) = (install_low, install_high) {
            info.install_timestamp = ((high as u64) << 32) | (low as u64);
            info.install_ole_date = unix_to_ole_date(info.install_timestamp);
            has_data = true;
        }

        let check_low = get_reg_dword(k, "CheckTimeLow");
        let check_high = get_reg_dword(k, "CheckTimeHigh");
        if let (Ok(low), Ok(high)) = (check_low, check_high) {
            info.check_timestamp = ((high as u64) << 32) | (low as u64);
            info.check_ole_date = unix_to_ole_date(info.check_timestamp);
        }

        if let Ok(qty) = get_reg_dword(k, "Quantity") {
            info.quantity = qty;
        }
        if let Ok(tamp) = get_reg_dword(k, "ClockTampered") {
            info.clock_tampered = tamp != 0;
        }
        if let Ok(key) = get_reg_str(k, "Key") {
            info.license_key = Some(key);
        }
        if let Ok(name) = get_reg_str(k, "Licensee") {
            info.licensee_name = Some(name);
        }
        if let Ok(state_json) = get_reg_str(k, "StateJson")
            && let Ok(state) = serde_json::from_str(&state_json)
        {
            info.state = state;
        }

        unsafe {
            let _ = RegCloseKey(k);
        }

        if has_data {
            info.refresh_state();
            return info;
        }
    }

    // Padrão do aplicativo: 100% de código aberto e gratuito (FreeEdition) sem qualquer restrição de trial
    let free_info = LicenseInfo::default();
    let _ = save_license_to_registry(&free_info);
    free_info
}

// ---------------------------------------------------------------------------
// Algoritmo HMAC auxiliar
// ---------------------------------------------------------------------------

/// HMAC simplificado baseado em FNV-1a com XOR cumulativo.
pub fn simple_hmac(data: &[u8], key: &[u8]) -> String {
    if key.is_empty() {
        return String::new();
    }
    let mut hash: u64 = 0xcbf29ce484222325; // FNV offset basis
    for (i, &byte) in data.iter().enumerate() {
        let k = key[i % key.len()];
        hash ^= (byte ^ k) as u64;
        hash = hash.wrapping_mul(0x100000001b3); // FNV prime
    }
    format!("{:016x}", hash)
}

/// Gera uma chave de teste no formato delimitado por `#`.
#[cfg(test)]
pub fn generate_test_key(name: &str, qty: u32, secret: &[u8]) -> String {
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let ts_hex = format!("{:x}", now);
    let payload = format!("{}#{}#{}", name, qty, ts_hex);
    let hmac = simple_hmac(payload.as_bytes(), secret);
    format!("{}#{}#{}#{}", name, qty, ts_hex, hmac)
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEST_SECRET: &[u8] = b"officetab-rs-dev-secret-key";

    #[test]
    fn test_free_edition_is_active() {
        let info = LicenseInfo::default();
        assert!(info.is_active());
        assert!(!info.should_reminder());
    }

    #[test]
    fn test_trial_is_active() {
        let info = LicenseInfo::new_trial();
        assert!(info.is_active());
        assert!(info.install_ole_date > 40000.0);
    }

    #[test]
    fn test_ole_date_conversions() {
        // Teste de ida e volta Unix <-> OLE Automation Date
        let now = 1718000000u64;
        let ole = unix_to_ole_date(now);
        let back_unix = ole_date_to_unix(ole);
        assert_eq!(now, back_unix);

        // Epoch de 1970
        assert_eq!(ole_date_to_unix(25569.0), 0);
    }

    #[test]
    fn test_expired_is_not_active() {
        let info = LicenseInfo {
            state: LicenseState::Expired { days_past: 5 },
            ..Default::default()
        };
        assert!(!info.is_active());
        assert!(info.should_reminder());
    }

    #[test]
    fn test_clock_tamper_detection() {
        let mut info = LicenseInfo::new_trial();
        // Simula última verificação no futuro (usuário voltou o relógio)
        info.check_timestamp += 86400 * 10;
        info.refresh_state();
        assert!(info.clock_tampered);
        assert_eq!(info.state, LicenseState::ClockTampered);
        assert!(!info.is_active());
        assert!(info.should_reminder());
    }

    #[test]
    fn test_offline_key_format_and_validation() {
        let licensee = "Empresa Exemplo Ltda";
        let qty = 15;
        let key = generate_offline_key(licensee, qty, TEST_SECRET);

        // Verificar formato XXXX-XXXX-XXXX-XXXX
        let parts: Vec<&str> = key.split('-').collect();
        assert_eq!(parts.len(), 4);
        for p in parts {
            assert_eq!(p.len(), 4);
        }

        // Validação com parâmetros corretos
        let val_result = validate_offline_key(&key, licensee, TEST_SECRET);
        assert!(val_result.is_ok());
        assert_eq!(val_result.unwrap(), 15);

        // Validação com nome incorreto
        let wrong_user = validate_offline_key(&key, "Outro Usuario", TEST_SECRET);
        assert!(wrong_user.is_err());

        // Validação com secret incorreto
        let wrong_secret = validate_offline_key(&key, licensee, b"outro-secret");
        assert!(wrong_secret.is_err());
    }

    #[test]
    fn test_register_offline_key() {
        let licensee = "Engenharia Corporativa";
        let key = generate_offline_key(licensee, 50, TEST_SECRET);

        let mut info = LicenseInfo::new_trial();
        let res = info.register_offline(&key, licensee, TEST_SECRET);
        assert!(res.is_ok());
        assert!(info.is_active());
        assert_eq!(
            info.state,
            LicenseState::Registered {
                licensee: licensee.to_string(),
                quantity: 50
            }
        );
        assert_eq!(info.license_key, Some(key));
    }

    #[test]
    fn test_license_serialization_roundtrip() {
        let info = LicenseInfo::new_trial();
        let json = info.to_json().unwrap();
        let restored = LicenseInfo::from_json(&json).unwrap();
        assert_eq!(restored.quantity, info.quantity);
        assert!((restored.install_ole_date - info.install_ole_date).abs() < 0.0001);
    }

    #[test]
    fn test_legacy_key_validation() {
        let key = generate_test_key("TestUser", 5, TEST_SECRET);
        let result = LicenseInfo::validate_key(&key, TEST_SECRET);
        assert!(result.is_ok());
        let (name, qty) = result.unwrap();
        assert_eq!(name, "TestUser");
        assert_eq!(qty, 5);

        let mut info = LicenseInfo::new_trial();
        assert!(info.register(&key, TEST_SECRET).is_ok());
        assert!(matches!(info.state, LicenseState::Registered { .. }));
    }

    #[test]
    fn test_registry_save_and_load() {
        let mut info = LicenseInfo::new_trial();
        info.quantity = 42;
        info.licensee_name = Some("Admin Registry Test".to_string());

        assert!(save_license_to_registry(&info).is_ok());
        let loaded = load_license_from_registry();
        assert_eq!(loaded.quantity, 42);
        assert_eq!(loaded.licensee_name.as_deref(), Some("Admin Registry Test"));
    }
}
