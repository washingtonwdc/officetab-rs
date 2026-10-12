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
//! Sistema de temas para o OfficeTabRs.
//!
//! Fornece estilos visuais completos, temas personalizados por aplicativo Office,
//! gradientes lineares, interpolação de cores e presets consagrados.

use serde::{Deserialize, Serialize};

/// Cor no formato RGB (0x00RRGGBB).
pub type Color = u32;

/// Orientação da barra de abas.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum TabOrientation {
    #[default]
    Top,
    Bottom,
    Left,
    Right,
}

impl TabOrientation {
    pub fn to_u32(self) -> u32 {
        match self {
            Self::Top => 0,
            Self::Bottom => 1,
            Self::Left => 2,
            Self::Right => 3,
        }
    }

    pub fn from_u32(val: u32) -> Self {
        match val {
            1 => Self::Bottom,
            2 => Self::Left,
            3 => Self::Right,
            _ => Self::Top,
        }
    }
}

/// Estilo visual das abas.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum TabStyle {
    /// Abas estilo Chrome (arredondadas no topo)
    #[default]
    Chrome,
    /// Abas quadradas clássicas
    Edge,
    /// Abas com ângulo reto
    RightAngle,
    /// Abas arredondadas à esquerda
    LeftRounded,
    /// Abas arredondadas à direita
    RightRounded,
}

impl TabStyle {
    pub fn to_u32(self) -> u32 {
        match self {
            Self::Chrome => 0,
            Self::Edge => 1,
            Self::RightAngle => 2,
            Self::LeftRounded => 3,
            Self::RightRounded => 4,
        }
    }

    pub fn from_u32(val: u32) -> Self {
        match val {
            1 => Self::Edge,
            2 => Self::RightAngle,
            3 => Self::LeftRounded,
            4 => Self::RightRounded,
            _ => Self::Chrome,
        }
    }
}

/// Aplicativos suportados da suíte Microsoft Office.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum OfficeApp {
    Word,
    Excel,
    PowerPoint,
    Access,
    Project,
    Visio,
    Publisher,
}

impl OfficeApp {
    /// Nome canônico do aplicativo em texto.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Word => "Word",
            Self::Excel => "Excel",
            Self::PowerPoint => "PowerPoint",
            Self::Access => "Access",
            Self::Project => "Project",
            Self::Visio => "Visio",
            Self::Publisher => "Publisher",
        }
    }

    /// Obtém o aplicativo a partir do nome do processo ou executável (case-insensitive).
    pub fn from_name(name: &str) -> Option<Self> {
        let clean = name.trim().to_lowercase();
        if clean.contains("winword") || clean.contains("word") {
            Some(Self::Word)
        } else if clean.contains("excel") {
            Some(Self::Excel)
        } else if clean.contains("powerpnt") || clean.contains("powerpoint") {
            Some(Self::PowerPoint)
        } else if clean.contains("msaccess") || clean.contains("access") {
            Some(Self::Access)
        } else if clean.contains("winproj") || clean.contains("project") {
            Some(Self::Project)
        } else if clean.contains("visio") {
            Some(Self::Visio)
        } else if clean.contains("mspub") || clean.contains("publisher") {
            Some(Self::Publisher)
        } else {
            None
        }
    }
}

/// Parada de cor para gradientes compostos.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GradientStop {
    /// Posição normalizada entre 0.0 e 1.0.
    pub position: f32,
    /// Cor RGB nesta parada.
    pub color: Color,
}

/// Configuração completa de um tema.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Theme {
    pub name: String,

    // Cores das abas
    pub active_tab_bg: Color,
    pub active_tab_text: Color,
    pub inactive_tab_bg: Color,
    pub inactive_tab_text: Color,
    pub hover_tab_bg: Color,
    pub hover_tab_text: Color,
    pub tab_border: Color,

    // Cores da barra de fundo
    pub bar_background: Color,

    // Fonte
    pub font_name: String,
    pub font_size: u32,

    // Dimensões
    pub tab_height: i32,
    pub tab_min_width: i32,
    pub tab_max_width: i32,
    pub tab_padding: i32,

    // Estilo
    pub style: TabStyle,
    pub orientation: TabOrientation,
    pub multiline: bool,
    pub show_close_button: bool,
    pub modern_theme: bool,

    // Suporte a gradientes e realce customizado
    #[serde(default)]
    pub active_tab_gradient_end: Option<Color>,
    #[serde(default)]
    pub bar_gradient_end: Option<Color>,
    #[serde(default)]
    pub custom_accent_color: Option<Color>,
}

impl Default for Theme {
    fn default() -> Self {
        Self::light()
    }
}

impl Theme {
    /// Tema claro padrão (estilo Office 365 moderno).
    pub fn light() -> Self {
        Self {
            name: "Light".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x00333333,
            inactive_tab_bg: 0x00F0F0F0,
            inactive_tab_text: 0x00666666,
            hover_tab_bg: 0x00E5E5E5,
            hover_tab_text: 0x00333333,
            tab_border: 0x00D0D0D0,
            bar_background: 0x00F5F5F5,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: None,
            custom_accent_color: Some(0x000078D7),
        }
    }

    /// Tema escuro (Dark Mode).
    pub fn dark() -> Self {
        Self {
            name: "Dark".to_string(),
            active_tab_bg: 0x002D2D2D,
            active_tab_text: 0x00FFFFFF,
            inactive_tab_bg: 0x001E1E1E,
            inactive_tab_text: 0x00AAAAAA,
            hover_tab_bg: 0x003A3A3A,
            hover_tab_text: 0x00FFFFFF,
            tab_border: 0x00404040,
            bar_background: 0x001E1E1E,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: None,
            custom_accent_color: Some(0x000078D7),
        }
    }

    /// Tema azul (estilo Office clássico).
    pub fn blue() -> Self {
        Self {
            name: "Blue".to_string(),
            active_tab_bg: 0x00D4E4F7,
            active_tab_text: 0x001A3A5C,
            inactive_tab_bg: 0x00B9D1EA,
            inactive_tab_text: 0x00335577,
            hover_tab_bg: 0x00C5DAF0,
            hover_tab_text: 0x001A3A5C,
            tab_border: 0x009BB5D0,
            bar_background: 0x00AEC9E4,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Edge,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: false,
            active_tab_gradient_end: None,
            bar_gradient_end: None,
            custom_accent_color: Some(0x001A3A5C),
        }
    }

    /// Tema verde.
    pub fn green() -> Self {
        Self {
            name: "Green".to_string(),
            active_tab_bg: 0x00D4EDDA,
            active_tab_text: 0x00155724,
            inactive_tab_bg: 0x00C3E6CB,
            inactive_tab_text: 0x002D6A3F,
            hover_tab_bg: 0x00B1DFBB,
            hover_tab_text: 0x00155724,
            tab_border: 0x00A3D1AF,
            bar_background: 0x00C3E6CB,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::LeftRounded,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: false,
            active_tab_gradient_end: None,
            bar_gradient_end: None,
            custom_accent_color: Some(0x00155724),
        }
    }

    /// Dark Modern: Tema escuro refinado com bordas sutis e fundo grafite escuro.
    pub fn dark_modern() -> Self {
        Self {
            name: "Dark Modern".to_string(),
            active_tab_bg: 0x00242424,
            active_tab_text: 0x00F3F3F3,
            inactive_tab_bg: 0x00181818,
            inactive_tab_text: 0x008E8E8E,
            hover_tab_bg: 0x002E2E2E,
            hover_tab_text: 0x00FFFFFF,
            tab_border: 0x00333333,
            bar_background: 0x00121212,
            font_name: "Segoe UI Variable".to_string(),
            font_size: 9,
            tab_height: 30,
            tab_min_width: 75,
            tab_max_width: 260,
            tab_padding: 10,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: Some(0x002B2B2B),
            bar_gradient_end: None,
            custom_accent_color: Some(0x000078D4),
        }
    }

    /// Fluent Blue: Visual moderno inspirado no design Fluent do Windows 11.
    pub fn fluent_blue() -> Self {
        Self {
            name: "Fluent Blue".to_string(),
            active_tab_bg: 0x00EBF3FB,
            active_tab_text: 0x00003E7E,
            inactive_tab_bg: 0x00F8FAFD,
            inactive_tab_text: 0x00506070,
            hover_tab_bg: 0x00DCEAFC,
            hover_tab_text: 0x00002A54,
            tab_border: 0x00C7DBF4,
            bar_background: 0x00F0F5FA,
            font_name: "Segoe UI Variable".to_string(),
            font_size: 9,
            tab_height: 29,
            tab_min_width: 72,
            tab_max_width: 250,
            tab_padding: 9,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: Some(0x00F4F8FD),
            bar_gradient_end: None,
            custom_accent_color: Some(0x000F6CBD),
        }
    }

    /// Excel Emerald: Tons esmeralda vibrantes otimizados para planilhas.
    pub fn excel_emerald() -> Self {
        Self {
            name: "Excel Emerald".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x00107C41,
            inactive_tab_bg: 0x00E8F5E9,
            inactive_tab_text: 0x002E7D32,
            hover_tab_bg: 0x00C8E6C9,
            hover_tab_text: 0x001B5E20,
            tab_border: 0x00A5D6A7,
            bar_background: 0x00107C41,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Edge,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: Some(0x000D5C30),
            custom_accent_color: Some(0x00107C41),
        }
    }

    /// PowerPoint Carmine: Tons carmim/laranja característicos de apresentações.
    pub fn powerpoint_carmine() -> Self {
        Self {
            name: "PowerPoint Carmine".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x00C43E1C,
            inactive_tab_bg: 0x00FBE9E7,
            inactive_tab_text: 0x00BF360C,
            hover_tab_bg: 0x00FFCCBC,
            hover_tab_text: 0x00A72B00,
            tab_border: 0x00FFAB91,
            bar_background: 0x00C43E1C,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::RightAngle,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: Some(0x009C2C0E),
            custom_accent_color: Some(0x00C43E1C),
        }
    }

    /// Access Crimson: Vinho carmesim elegante estilo bancos de dados Access.
    pub fn access_crimson() -> Self {
        Self {
            name: "Access Crimson".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x00A4373A,
            inactive_tab_bg: 0x00FDE8E8,
            inactive_tab_text: 0x00781E21,
            hover_tab_bg: 0x00F8C8C8,
            hover_tab_text: 0x005E1215,
            tab_border: 0x00E0A3A5,
            bar_background: 0x00A4373A,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: Some(0x007A1D20),
            custom_accent_color: Some(0x00A4373A),
        }
    }

    /// OneNote Purple: Tons púrpura e violeta para organização e notas.
    pub fn onenote_purple() -> Self {
        Self {
            name: "OneNote Purple".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x007719AA,
            inactive_tab_bg: 0x00F3E8FD,
            inactive_tab_text: 0x00591081,
            hover_tab_bg: 0x00E5CCFA,
            hover_tab_text: 0x00450A66,
            tab_border: 0x00D0A5F5,
            bar_background: 0x007719AA,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: Some(0x005A0F82),
            custom_accent_color: Some(0x007719AA),
        }
    }

    /// High Contrast: Tema de alto contraste para máxima acessibilidade.
    pub fn high_contrast() -> Self {
        Self {
            name: "High Contrast".to_string(),
            active_tab_bg: 0x00000000,
            active_tab_text: 0x0000FFFF, // Ciano vivo
            inactive_tab_bg: 0x00000000,
            inactive_tab_text: 0x00FFFFFF,
            hover_tab_bg: 0x001A1A1A,
            hover_tab_text: 0x0000FF00, // Verde neon
            tab_border: 0x0000FFFF,
            bar_background: 0x00000000,
            font_name: "Segoe UI".to_string(),
            font_size: 10,
            tab_height: 32,
            tab_min_width: 80,
            tab_max_width: 260,
            tab_padding: 10,
            style: TabStyle::Edge,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: false,
            active_tab_gradient_end: None,
            bar_gradient_end: None,
            custom_accent_color: Some(0x0000FFFF),
        }
    }

    /// Nord: Paleta ártica fria e harmônica.
    pub fn nord() -> Self {
        Self {
            name: "Nord".to_string(),
            active_tab_bg: 0x003B4252,     // nord1
            active_tab_text: 0x00ECEFF4,   // nord6
            inactive_tab_bg: 0x002E3440,   // nord0
            inactive_tab_text: 0x00D8DEE9, // nord4
            hover_tab_bg: 0x00434C5E,      // nord2
            hover_tab_text: 0x0088C0D0,    // nord8
            tab_border: 0x004C566A,        // nord3
            bar_background: 0x002E3440,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: Some(0x00434C5E),
            bar_gradient_end: None,
            custom_accent_color: Some(0x0088C0D0),
        }
    }

    /// Dracula: Tema escuro consagrado com toques roxo e rosa pastel.
    pub fn dracula() -> Self {
        Self {
            name: "Dracula".to_string(),
            active_tab_bg: 0x0044475A,     // Selection / Current Line
            active_tab_text: 0x00F8F8F2,   // Foreground
            inactive_tab_bg: 0x00282A36,   // Background
            inactive_tab_text: 0x006272A4, // Comment
            hover_tab_bg: 0x00565B75,
            hover_tab_text: 0x0050FA7B, // Green
            tab_border: 0x006272A4,
            bar_background: 0x001E1F29,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: None,
            custom_accent_color: Some(0x00BD93F9), // Purple
        }
    }

    /// Solarized: Paleta equilibrada Solarized Dark.
    pub fn solarized() -> Self {
        Self {
            name: "Solarized".to_string(),
            active_tab_bg: 0x00073642,     // base02
            active_tab_text: 0x0093A1A1,   // base1
            inactive_tab_bg: 0x00002B36,   // base03
            inactive_tab_text: 0x00586E75, // base01
            hover_tab_bg: 0x000A404F,
            hover_tab_text: 0x002AA198, // cyan
            tab_border: 0x00073642,
            bar_background: 0x0000212B,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: None,
            custom_accent_color: Some(0x00268BD2), // blue
        }
    }

    /// Retorna tema temático nativo para Microsoft Word.
    pub fn word() -> Self {
        Self {
            name: "Word Classic".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x002B579A,
            inactive_tab_bg: 0x00EAF1FB,
            inactive_tab_text: 0x001B3A68,
            hover_tab_bg: 0x00D6E4F8,
            hover_tab_text: 0x00122949,
            tab_border: 0x00B2CDF2,
            bar_background: 0x002B579A,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: Some(0x001F3E6E),
            custom_accent_color: Some(0x002B579A),
        }
    }

    /// Retorna tema temático nativo para Microsoft Excel.
    pub fn excel() -> Self {
        Self::excel_emerald()
    }

    /// Retorna tema temático nativo para Microsoft PowerPoint.
    pub fn powerpoint() -> Self {
        Self::powerpoint_carmine()
    }

    /// Retorna tema temático nativo para Microsoft Access.
    pub fn access() -> Self {
        Self::access_crimson()
    }

    /// Retorna tema temático nativo para Microsoft Project.
    pub fn project() -> Self {
        Self {
            name: "Project Teal".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x0031752F,
            inactive_tab_bg: 0x00EAF4EA,
            inactive_tab_text: 0x00235422,
            hover_tab_bg: 0x00D4EBD4,
            hover_tab_text: 0x00173916,
            tab_border: 0x00B0DDB0,
            bar_background: 0x0031752F,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: Some(0x00235422),
            custom_accent_color: Some(0x0031752F),
        }
    }

    /// Retorna tema temático nativo para Microsoft Visio.
    pub fn visio() -> Self {
        Self {
            name: "Visio Indigo".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x003955A3,
            inactive_tab_bg: 0x00ECEFF8,
            inactive_tab_text: 0x0026396E,
            hover_tab_bg: 0x00DCE2F3,
            hover_tab_text: 0x00182548,
            tab_border: 0x00B8C5E7,
            bar_background: 0x003955A3,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: Some(0x00283C74),
            custom_accent_color: Some(0x003955A3),
        }
    }

    /// Retorna tema temático nativo para Microsoft Publisher.
    pub fn publisher() -> Self {
        Self {
            name: "Publisher Teal".to_string(),
            active_tab_bg: 0x00FFFFFF,
            active_tab_text: 0x00008272,
            inactive_tab_bg: 0x00E5F5F3,
            inactive_tab_text: 0x00005C51,
            hover_tab_bg: 0x00CCEBE7,
            hover_tab_text: 0x00003D36,
            tab_border: 0x0099D6CE,
            bar_background: 0x00008272,
            font_name: "Segoe UI".to_string(),
            font_size: 9,
            tab_height: 28,
            tab_min_width: 70,
            tab_max_width: 255,
            tab_padding: 8,
            style: TabStyle::Chrome,
            orientation: TabOrientation::Top,
            multiline: false,
            show_close_button: true,
            modern_theme: true,
            active_tab_gradient_end: None,
            bar_gradient_end: Some(0x00005C51),
            custom_accent_color: Some(0x00008272),
        }
    }

    /// Retorna o tema padrão otimizado para um determinado aplicativo Office.
    pub fn for_app(app: OfficeApp) -> Self {
        match app {
            OfficeApp::Word => Self::word(),
            OfficeApp::Excel => Self::excel(),
            OfficeApp::PowerPoint => Self::powerpoint(),
            OfficeApp::Access => Self::access(),
            OfficeApp::Project => Self::project(),
            OfficeApp::Visio => Self::visio(),
            OfficeApp::Publisher => Self::publisher(),
        }
    }

    /// Retorna todos os temas pré-definidos disponíveis.
    pub fn presets() -> Vec<Theme> {
        vec![
            Self::light(),
            Self::dark(),
            Self::blue(),
            Self::green(),
            Self::dark_modern(),
            Self::fluent_blue(),
            Self::excel_emerald(),
            Self::powerpoint_carmine(),
            Self::access_crimson(),
            Self::onenote_purple(),
            Self::high_contrast(),
            Self::nord(),
            Self::dracula(),
            Self::solarized(),
        ]
    }

    /// Busca um preset pelo nome (case-insensitive e tolerante a espaços).
    pub fn find_preset(name: &str) -> Option<Theme> {
        let clean = name.trim().to_lowercase().replace(['-', '_'], " ");
        Self::presets().into_iter().find(|t| {
            let t_clean = t.name.to_lowercase().replace(['-', '_'], " ");
            t_clean == clean
        })
    }

    /// Converte uma cor do formato RGB interno para COLORREF do Win32 (0x00BBGGRR).
    pub fn to_colorref(rgb: Color) -> u32 {
        let r = (rgb >> 16) & 0xFF;
        let g = (rgb >> 8) & 0xFF;
        let b = rgb & 0xFF;
        (b << 16) | (g << 8) | r
    }

    /// Decompõe uma cor RGB (0x00RRGGBB) nas componentes (R, G, B).
    pub fn rgb_to_components(rgb: Color) -> (u8, u8, u8) {
        let r = ((rgb >> 16) & 0xFF) as u8;
        let g = ((rgb >> 8) & 0xFF) as u8;
        let b = (rgb & 0xFF) as u8;
        (r, g, b)
    }

    /// Constrói uma cor RGB (0x00RRGGBB) a partir das componentes (R, G, B).
    pub fn components_to_rgb(r: u8, g: u8, b: u8) -> Color {
        ((r as u32) << 16) | ((g as u32) << 8) | (b as u32)
    }

    /// Interpola linearmente entre duas cores com um fator de 0.0 a 1.0.
    pub fn interpolate_color(c1: Color, c2: Color, factor: f32) -> Color {
        let factor = factor.clamp(0.0, 1.0);
        let (r1, g1, b1) = Self::rgb_to_components(c1);
        let (r2, g2, b2) = Self::rgb_to_components(c2);

        let r = ((r1 as f32) + ((r2 as f32) - (r1 as f32)) * factor).round() as u8;
        let g = ((g1 as f32) + ((g2 as f32) - (g1 as f32)) * factor).round() as u8;
        let b = ((b1 as f32) + ((b2 as f32) - (b1 as f32)) * factor).round() as u8;

        Self::components_to_rgb(r, g, b)
    }

    /// Clareia uma cor RGB em um percentual dado (0.0 a 1.0).
    pub fn lighten(color: Color, amount: f32) -> Color {
        Self::interpolate_color(color, 0x00FFFFFF, amount)
    }

    /// Escurece uma cor RGB em um percentual dado (0.0 a 1.0).
    pub fn darken(color: Color, amount: f32) -> Color {
        Self::interpolate_color(color, 0x00000000, amount)
    }

    /// Converte uma string hexadecimal (ex: "#0078D7" ou "0078D7") para Color RGB.
    pub fn hex_to_color(hex: &str) -> Option<Color> {
        let s = hex.trim().trim_start_matches('#');
        if s.len() == 6 {
            u32::from_str_radix(s, 16).ok()
        } else {
            None
        }
    }

    /// Converte Color RGB para representação hexadecimal "#RRGGBB".
    pub fn color_to_hex(color: Color) -> String {
        format!("#{:06X}", color & 0x00FFFFFF)
    }

    /// Calcula os degraus intermediários de um gradiente linear.
    pub fn calculate_gradient_steps(start: Color, end: Color, steps: usize) -> Vec<Color> {
        if steps == 0 {
            return Vec::new();
        }
        if steps == 1 {
            return vec![start];
        }
        let mut result = Vec::with_capacity(steps);
        for i in 0..steps {
            let factor = (i as f32) / ((steps - 1) as f32);
            result.push(Self::interpolate_color(start, end, factor));
        }
        result
    }

    /// Cria uma HBRUSH a partir de uma cor RGB.
    /// O caller é responsável por chamar DeleteObject na brush retornada.
    #[cfg(windows)]
    pub fn create_brush(rgb: Color) -> windows::Win32::Graphics::Gdi::HBRUSH {
        use windows::Win32::Foundation::COLORREF;
        use windows::Win32::Graphics::Gdi::CreateSolidBrush;
        unsafe { CreateSolidBrush(COLORREF(Self::to_colorref(rgb))) }
    }

    /// Serializa o tema para JSON formatado.
    pub fn to_json(&self) -> Result<String, serde_json::Error> {
        serde_json::to_string_pretty(self)
    }

    /// Deserializa um tema a partir de uma string JSON.
    pub fn from_json(json: &str) -> Result<Self, serde_json::Error> {
        serde_json::from_str(json)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_colorref_conversion() {
        // RGB 0x00FF0000 (vermelho) → COLORREF 0x000000FF
        assert_eq!(Theme::to_colorref(0x00FF0000), 0x000000FF);
        // RGB 0x0000FF00 (verde) → COLORREF 0x0000FF00
        assert_eq!(Theme::to_colorref(0x0000FF00), 0x0000FF00);
        // RGB 0x000000FF (azul) → COLORREF 0x00FF0000
        assert_eq!(Theme::to_colorref(0x000000FF), 0x00FF0000);
    }

    #[test]
    fn test_theme_serialization_roundtrip() {
        let theme = Theme::dark_modern();
        let json = theme.to_json().unwrap();
        let restored = Theme::from_json(&json).unwrap();
        assert_eq!(restored.name, "Dark Modern");
        assert_eq!(restored.active_tab_bg, theme.active_tab_bg);
        assert_eq!(restored.font_size, theme.font_size);
        assert_eq!(
            restored.active_tab_gradient_end,
            theme.active_tab_gradient_end
        );
        assert_eq!(restored, theme);
    }

    #[test]
    fn test_find_preset() {
        assert!(Theme::find_preset("dark").is_some());
        assert!(Theme::find_preset("LIGHT").is_some());
        assert!(Theme::find_preset("Blue").is_some());
        assert!(Theme::find_preset("green").is_some());
        assert!(Theme::find_preset("Dark Modern").is_some());
        assert!(Theme::find_preset("fluent-blue").is_some());
        assert!(Theme::find_preset("excel emerald").is_some());
        assert!(Theme::find_preset("PowerPoint Carmine").is_some());
        assert!(Theme::find_preset("Access Crimson").is_some());
        assert!(Theme::find_preset("OneNote Purple").is_some());
        assert!(Theme::find_preset("High Contrast").is_some());
        assert!(Theme::find_preset("Nord").is_some());
        assert!(Theme::find_preset("Dracula").is_some());
        assert!(Theme::find_preset("Solarized").is_some());
        assert!(Theme::find_preset("nonexistent").is_none());
    }

    #[test]
    fn test_presets_count() {
        // Pelo menos 14 presets (4 originais + 10 adicionais)
        assert!(Theme::presets().len() >= 12);
        assert_eq!(Theme::presets().len(), 14);
    }

    #[test]
    fn test_default_is_light() {
        let theme = Theme::default();
        assert_eq!(theme.name, "Light");
    }

    #[test]
    fn test_tab_orientation_conversion() {
        assert_eq!(TabOrientation::Top.to_u32(), 0);
        assert_eq!(TabOrientation::Bottom.to_u32(), 1);
        assert_eq!(TabOrientation::Left.to_u32(), 2);
        assert_eq!(TabOrientation::Right.to_u32(), 3);

        assert_eq!(TabOrientation::from_u32(0), TabOrientation::Top);
        assert_eq!(TabOrientation::from_u32(1), TabOrientation::Bottom);
        assert_eq!(TabOrientation::from_u32(2), TabOrientation::Left);
        assert_eq!(TabOrientation::from_u32(3), TabOrientation::Right);
        assert_eq!(TabOrientation::from_u32(99), TabOrientation::Top);
    }

    #[test]
    fn test_tab_style_conversion() {
        assert_eq!(TabStyle::Chrome.to_u32(), 0);
        assert_eq!(TabStyle::Edge.to_u32(), 1);
        assert_eq!(TabStyle::RightAngle.to_u32(), 2);
        assert_eq!(TabStyle::LeftRounded.to_u32(), 3);
        assert_eq!(TabStyle::RightRounded.to_u32(), 4);

        assert_eq!(TabStyle::from_u32(0), TabStyle::Chrome);
        assert_eq!(TabStyle::from_u32(1), TabStyle::Edge);
        assert_eq!(TabStyle::from_u32(2), TabStyle::RightAngle);
        assert_eq!(TabStyle::from_u32(3), TabStyle::LeftRounded);
        assert_eq!(TabStyle::from_u32(4), TabStyle::RightRounded);
        assert_eq!(TabStyle::from_u32(99), TabStyle::Chrome);
    }

    #[test]
    fn test_office_app_detection_and_themes() {
        assert_eq!(OfficeApp::from_name("WINWORD.EXE"), Some(OfficeApp::Word));
        assert_eq!(OfficeApp::from_name("excel.exe"), Some(OfficeApp::Excel));
        assert_eq!(
            OfficeApp::from_name("POWERPNT.exe"),
            Some(OfficeApp::PowerPoint)
        );
        assert_eq!(
            OfficeApp::from_name("MSACCESS.exe"),
            Some(OfficeApp::Access)
        );
        assert_eq!(
            OfficeApp::from_name("winproj.exe"),
            Some(OfficeApp::Project)
        );
        assert_eq!(OfficeApp::from_name("VISIO.exe"), Some(OfficeApp::Visio));
        assert_eq!(
            OfficeApp::from_name("mspub.exe"),
            Some(OfficeApp::Publisher)
        );
        assert_eq!(OfficeApp::from_name("notepad.exe"), None);

        let word_th = Theme::for_app(OfficeApp::Word);
        assert_eq!(word_th.name, "Word Classic");

        let excel_th = Theme::for_app(OfficeApp::Excel);
        assert_eq!(excel_th.name, "Excel Emerald");

        let ppt_th = Theme::for_app(OfficeApp::PowerPoint);
        assert_eq!(ppt_th.name, "PowerPoint Carmine");

        let acc_th = Theme::for_app(OfficeApp::Access);
        assert_eq!(acc_th.name, "Access Crimson");

        let proj_th = Theme::for_app(OfficeApp::Project);
        assert_eq!(proj_th.name, "Project Teal");

        let visio_th = Theme::for_app(OfficeApp::Visio);
        assert_eq!(visio_th.name, "Visio Indigo");

        let pub_th = Theme::for_app(OfficeApp::Publisher);
        assert_eq!(pub_th.name, "Publisher Teal");
    }

    #[test]
    fn test_gradient_and_color_manipulation() {
        let c1 = 0x00000000; // Preto
        let c2 = 0x00FFFFFF; // Branco
        let mid = Theme::interpolate_color(c1, c2, 0.5);
        assert_eq!(mid, 0x00808080); // Cinza intermediário

        let lighter = Theme::lighten(0x00000000, 0.5);
        assert_eq!(lighter, 0x00808080);

        let darker = Theme::darken(0x00FFFFFF, 0.5);
        assert_eq!(darker, 0x00808080);

        let hex = Theme::color_to_hex(0x000078D7);
        assert_eq!(hex, "#0078D7");

        let parsed = Theme::hex_to_color("#0078D7").unwrap();
        assert_eq!(parsed, 0x000078D7);

        let steps = Theme::calculate_gradient_steps(0x00000000, 0x000000FF, 3);
        assert_eq!(steps.len(), 3);
        assert_eq!(steps[0], 0x00000000);
        assert_eq!(steps[1], 0x00000080);
        assert_eq!(steps[2], 0x000000FF);
    }

    #[test]
    #[cfg(windows)]
    fn test_create_brush_valid() {
        use windows::Win32::Graphics::Gdi::DeleteObject;
        let brush = Theme::create_brush(0x00AABBCC);
        assert!(!brush.is_invalid());
        unsafe {
            let _ = DeleteObject(brush);
        }
    }
}
