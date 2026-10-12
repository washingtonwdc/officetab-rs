//! Logging estruturado com tracing.
//!
//! Copyright 2024 The OfficeTabRs Authors
//!
//! Licensed under the Apache License, Version 2.0 (the "License");
//! you may not use this file except in compliance with the License.
//! You may obtain a copy of the License at
//!
//!     http://www.apache.org/licenses/LICENSE-2.0
//!
//! Unless required by applicable law or agreed to in writing, software
//! distributed under the License is distributed on an "AS IS" BASIS,
//! WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//! See the License for the specific language governing permissions and
//! limitations under the License.

use once_cell::sync::OnceCell;
use tracing_subscriber::{fmt, EnvFilter};

static LOGGER_INIT: OnceCell<()> = OnceCell::new();

/// Inicializa o sistema de logging com tracing.
/// Deve ser chamado uma vez na inicialização do add-in.
pub fn init_logging() {
    LOGGER_INIT.get_or_init(|| {
        let filter = EnvFilter::try_from_default_env()
            .unwrap_or_else(|_| EnvFilter::new("officetab_rs=debug,tracing=info"));
        
        let subscriber = fmt()
            .with_writer(std::io::stderr)
            .with_target(true)
            .with_thread_ids(true)
            .with_file(true)
            .with_line_number(true)
            .with_filter(filter);
        
        // Também loga para arquivo dinâmico
        let caminho_log = crate::utils::get_current_module_path()
            .ok()
            .and_then(|p| p.parent().map(|d| d.join("officetab.log")))
            .unwrap_or_else(|| std::env::temp_dir().join("officetab.log"));

        if let Ok(file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(caminho_log)
        {
            let _ = fmt()
                .with_ansi(false)
                .with_writer(move || file.try_clone().unwrap())
                .with_target(true)
                .with_thread_ids(true)
                .with_file(true)
                .with_line_number(true)
                .with_filter(filter)
                .try_init();
        }
        
        let _ = tracing_subscriber::registry().with(subscriber).try_init();
    });
}

/// Verifica se o logging foi inicializado.
pub fn is_initialized() -> bool {
    LOGGER_INIT.get().is_some()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_logging_initialization() {
        init_logging();
        assert!(is_initialized());
    }
}
