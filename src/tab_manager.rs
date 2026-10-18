// Gerenciador central de abas do Office (baseado no TOTWindowsManager).
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
//! Gerenciador central de abas do Office (baseado no TOTWindowsManager).
//! Arquitetura orientada a eventos: WinEventHook e timer UI somente como fallback.
//! Criação/renderização de janelas Win32 exclusivamente na thread de UI.

use std::collections::HashMap;
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};

use windows::Win32::Foundation::{BOOL, COLORREF, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    CLEARTYPE_QUALITY, CLIP_DEFAULT_PRECIS, ClientToScreen, CreateFontW, CreateSolidBrush,
    DEFAULT_CHARSET, DEFAULT_PITCH, DT_CENTER, DT_END_ELLIPSIS, DT_LEFT, DT_NOPREFIX,
    DT_SINGLELINE, DT_VCENTER, DeleteObject, DrawTextW, FF_DONTCARE, FW_BOLD, FillRect, HDC,
    HGDIOBJ, OUT_DEFAULT_PRECIS, ScreenToClient, SelectObject, SetBkMode, SetTextColor,
    TRANSPARENT, InvalidateRect, UpdateWindow,
};
use windows::Win32::System::Threading::GetCurrentProcessId;
use windows::Win32::UI::Accessibility::{HWINEVENTHOOK, SetWinEventHook, UnhookWinEvent};
use windows::Win32::UI::Controls::{
    CDDS_ITEMPREPAINT, CDDS_PREPAINT, CDIS_HOT, CDIS_SELECTED, CDRF_NOTIFYITEMDRAW,
    CDRF_SKIPDEFAULT, NM_CUSTOMDRAW, NMCUSTOMDRAW, NMCUSTOMDRAW_DRAW_STATE_FLAGS, NMHDR,
    TCN_SELCHANGE,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    KEYEVENTF_KEYUP, SetFocus, VK_CONTROL, keybd_event,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, BringWindowToTop, CreatePopupMenu, DestroyMenu, EVENT_OBJECT_CREATE,
    EVENT_OBJECT_SHOW, EnumChildWindows, EnumWindows, FindWindowExW, GWL_STYLE, GetClassNameW,
    GetClientRect, GetCursorPos, GetDlgItem, GetWindowLongPtrW, GetWindowRect, GetWindowTextW,
    GetWindowThreadProcessId, HWND_TOP, IsIconic, IsWindow, IsWindowVisible, IsZoomed,
    KillTimer, MF_SEPARATOR, MF_STRING, MoveWindow, PostMessageW, SW_HIDE, SW_MAXIMIZE, SW_RESTORE,
    SW_SHOW, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SendMessageW, SetForegroundWindow,
    SetTimer, SetWindowLongPtrW, SetWindowPos, ShowWindow, TPM_RETURNCMD, TPM_RIGHTBUTTON,
    TrackPopupMenu, WINEVENT_OUTOFCONTEXT,
    WM_ACTIVATE, WM_CLOSE, WS_CLIPCHILDREN,
};
use windows::core::{PCWSTR, Result, w};

use crate::persistence::OfficeTabConfig;
use crate::subclass::SubclassManager;
use crate::tab_control::TabControl;

// ---------------------------------------------------------------------------
// Mensagens customizadas REA
// ---------------------------------------------------------------------------

/// Mensagem REA: fallback — postada pela thread de polling quando WinEventHook não disponível.
pub const WM_OFFICETAB_INICIALIZAR: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 0x01;

/// Mensagem REA: recarga forçada dos títulos/abas.
pub const WM_OFFICETAB_RECARREGAR: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 0x02;

/// Mensagem customizada: criação assíncrona de novo documento.
pub const WM_OFFICETAB_NOVO_DOC: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 0x03;

/// Mensagem customizada: alternância DIFERIDA de janela (Modo Janela Única).
/// Postada pelo handler de clique em aba e processada FORA do WM_NOTIFY,
/// na própria janela de destino — evita reentrância e o WM_ACTIVATE forjado
/// com lParam=0 que causava Access Violation dentro do Excel ao alternar abas.
/// wParam = HWND bruto da janela de origem (que será ocultada).
pub const WM_OFFICETAB_ALTERNAR: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 0x04;

// ---------------------------------------------------------------------------
// Constantes e utilitários de log
// ---------------------------------------------------------------------------

pub fn log_debug(msg: &str) {
    use std::io::Write;
    let caminho_log = crate::utils::get_current_module_path()
        .ok()
        .and_then(|p| p.parent().map(|d| d.join("officetab.log")))
        .unwrap_or_else(|| std::env::temp_dir().join("officetab.log"));

    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(caminho_log)
    {
        let _ = writeln!(f, "[LOG] {}", msg);
    }
}

pub const OFFICETAB_CONTROL_ID: u32 = 5001;
pub const OFFICETAB_TIMER_ID: usize = 0x544142; // 'TAB'
pub const ALTURA_BARRA_ABAS: i32 = 44;

// ---------------------------------------------------------------------------
// Estruturas de dados
// ---------------------------------------------------------------------------

/// Metadados de uma janela de documento do Office
#[derive(Debug, Clone)]
pub struct JanelaDocumento {
    pub hwnd: HWND,
    pub titulo: String,
    pub classe: String,
}

/// Enumeração dos 7 aplicativos Microsoft Office suportados e mapeados pela engenharia reversa:
/// Word, Excel, PowerPoint, Access, Project, Visio e Publisher.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TipoAppOffice {
    Word,
    Excel,
    PowerPoint,
    Access,
    Project,
    Visio,
    Publisher,
    Desconhecido,
}

impl TipoAppOffice {
    /// Retorna verdadeiro se o tipo de aplicação pertencer a um aplicativo Office mapeado.
    pub fn e_office(&self) -> bool {
        !matches!(self, TipoAppOffice::Desconhecido)
    }

    /// Nome amigável do aplicativo Office para logs e depuração.
    pub fn nome(&self) -> &'static str {
        match self {
            TipoAppOffice::Word => "Word",
            TipoAppOffice::Excel => "Excel",
            TipoAppOffice::PowerPoint => "PowerPoint",
            TipoAppOffice::Access => "Access",
            TipoAppOffice::Project => "Project",
            TipoAppOffice::Visio => "Visio",
            TipoAppOffice::Publisher => "Publisher",
            TipoAppOffice::Desconhecido => "Desconhecido",
        }
    }
}

/// Identifica o aplicativo Office com base no nome da classe da janela Win32.
/// Reconhece com precisão os 7 aplicativos originais mapeados pelo relatório de engenharia reversa:
/// - Word: OpusApp
/// - Excel: XLMAIN, EXCEL7, EXCEL2
/// - PowerPoint: PPTFrameClass, PP9FrameClass, PP10FrameClass, PP11FrameClass, PP12FrameClass
/// - Access: MSACCESS, OMain, AccessApp
/// - Project: PB, WINPROJ, Project
/// - Visio: VISIO, VISIOA, VISIOM, VISIOU
/// - Publisher: PUBLISHER, MSPUB, PUBDesk
pub fn obter_tipo_app(classe: &str) -> TipoAppOffice {
    if classe.eq_ignore_ascii_case("OpusApp") {
        TipoAppOffice::Word
    } else if ["XLMAIN", "EXCEL7", "EXCEL2"]
        .iter()
        .any(|c| classe.eq_ignore_ascii_case(c))
    {
        TipoAppOffice::Excel
    } else if [
        "PPTFrameClass",
        "PP9FrameClass",
        "PP10FrameClass",
        "PP11FrameClass",
        "PP12FrameClass",
    ]
    .iter()
    .any(|c| classe.eq_ignore_ascii_case(c))
    {
        TipoAppOffice::PowerPoint
    } else if ["MSACCESS", "OMain", "AccessApp"]
        .iter()
        .any(|c| classe.eq_ignore_ascii_case(c))
    {
        TipoAppOffice::Access
    } else if ["PB", "WINPROJ", "Project"]
        .iter()
        .any(|c| classe.eq_ignore_ascii_case(c))
    {
        TipoAppOffice::Project
    } else if ["VISIO", "VISIOA", "VISIOM", "VISIOU"]
        .iter()
        .any(|c| classe.eq_ignore_ascii_case(c))
    {
        TipoAppOffice::Visio
    } else if ["PUBLISHER", "MSPUB", "PUBDesk"]
        .iter()
        .any(|c| classe.eq_ignore_ascii_case(c))
    {
        TipoAppOffice::Publisher
    } else {
        TipoAppOffice::Desconhecido
    }
}

/// Classes de janela reconhecidas para as 7 aplicações Office
pub const OFFICE_WINDOW_CLASSES: &[&str] = &[
    // Microsoft Word
    "OpusApp",
    // Microsoft Excel
    "XLMAIN",
    "EXCEL7",
    "EXCEL2",
    // Microsoft PowerPoint
    "PPTFrameClass",
    "PP9FrameClass",
    "PP10FrameClass",
    "PP11FrameClass",
    "PP12FrameClass",
    // Microsoft Access
    "MSACCESS",
    "OMain",
    "AccessApp",
    // Microsoft Project
    "PB",
    "WINPROJ",
    "Project",
    // Microsoft Visio
    "VISIO",
    "VISIOA",
    "VISIOM",
    "VISIOU",
    // Microsoft Publisher
    "PUBLISHER",
    "MSPUB",
    "PUBDesk",
];

/// Localiza a janela interna de documento (MDI client desk ou área de visualização) para os 7 aplicativos Office.
/// Retorna APENAS janelas reais que estejam atualmente visíveis na tela de edição.
pub fn localizar_area_documento(hwnd_top: HWND) -> Option<HWND> {
    unsafe {
        const CLASSES_DOC: &[PCWSTR] = &[
            w!("_WwF"),          // Word
            w!("XLDESK"),        // Excel MDI desk
            w!("paneClassDC"),   // PowerPoint slide canvas
            w!("PP9FrameClass"), // PowerPoint legacy frame
            w!("MDIClient"),     // Access / Visio / Project / PowerPoint MDI
            w!("MDICLIENT"),     // Access MDI alternativo
            w!("VISIOA"),        // Visio workspace
            w!("VISIOM"),        // Visio main view
            w!("ViewDesk"),      // Project / Visio desktop
            w!("ProjDesk"),      // Project desk
            w!("PBView"),        // Project view
            w!("PUBDesk"),       // Publisher layout desk
            w!("mdiClass"),      // Modern Office MDI container
            w!("_WwG"),          // Word document edit area
            // Excel adicional para versões modernas
            w!("EXCEL7"),        // Excel 7+ document area
            w!("EXCEL2"),        // Excel 2 document area
            w!("Sheet"),         // Excel worksheet
            w!("XLSpreadsheet"), // Excel spreadsheet class
        ];

        for &classe in CLASSES_DOC {
            let child = FindWindowExW(hwnd_top, HWND::default(), classe, PCWSTR::null())
                .unwrap_or_default();
            if !child.is_invalid() && IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                if GetWindowRect(child, &mut rc).is_ok() {
                    let w = rc.right - rc.left;
                    let h = rc.bottom - rc.top;
                    if w > 80 && h > 80 {
                        return Some(child);
                    }
                }
            }
        }

        // Fallback: procurar qualquer janela filha grande e visível
        let mut child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, PCWSTR::null(), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                if GetWindowRect(child, &mut rc).is_ok() {
                    let w = rc.right - rc.left;
                    let h = rc.bottom - rc.top;
                    if w > 200 && h > 200 {
                        let classe = extrair_classe_janela(child);
                        log_debug(&format!("localizar_area_documento: fallback encontrou classe '{}' ({})x({})", classe, w, h));
                        return Some(child);
                    }
                }
            }
        }

        None
    }
}

/// Calcula a altura da barra de status no rodapé da janela do Office (se presente),
/// garantindo que a área do documento não se sobreponha à barra de status nos 7 aplicativos.
pub fn calcular_altura_status_bar(hwnd_top: HWND, altura_cliente: i32) -> i32 {
    unsafe {
        let sb = FindWindowExW(
            hwnd_top,
            HWND::default(),
            w!("msctls_statusbar32"),
            PCWSTR::null(),
        )
        .unwrap_or_default();
        if !sb.is_invalid() && IsWindowVisible(sb).as_bool() {
            let mut rc = RECT::default();
            if GetWindowRect(sb, &mut rc).is_ok() {
                let h = rc.bottom - rc.top;
                if (16..=48).contains(&h) {
                    return h;
                }
            }
        }

        let mut child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, PCWSTR::null(), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                if GetWindowRect(child, &mut rc).is_ok() {
                    let mut pt = POINT {
                        x: rc.left,
                        y: rc.top,
                    };
                    let _ = ScreenToClient(hwnd_top, &mut pt);
                    let h = rc.bottom - rc.top;
                    if pt.y >= altura_cliente - 36
                        && pt.y < altura_cliente
                        && (18..=36).contains(&h)
                    {
                        return h;
                    }
                }
            }
        }

        24 // Altura padrão do rodapé nos aplicativos Office modernos
    }
}

/// Detecta se a janela do Office está na tela inicial ou no modo Backstage (Menu Arquivo / Novo / Abrir).
pub fn backstage_esta_aberto(hwnd_top: HWND) -> bool {
    unsafe {
        let mut rc_client = RECT::default();
        if GetClientRect(hwnd_top, &mut rc_client).is_err() {
            return false;
        }
        let w_client = rc_client.right - rc_client.left;
        let h_client = rc_client.bottom - rc_client.top;
        if w_client <= 100 || h_client <= 100 {
            return false;
        }

        // Se a área de documento não existe ou não está visível, o Office está no Backstage/tela inicial
        if localizar_area_documento(hwnd_top).is_none() {
            return true;
        }

        // Verificar se há janela NetUIHWND de grande porte cobrindo mais de 60% da tela (Backstage ativo)
        let mut child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, w!("NetUIHWND"), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                let _ = GetWindowRect(child, &mut rc);
                let w = rc.right - rc.left;
                let h = rc.bottom - rc.top;
                if w > (w_client * 3 / 4) && h > (h_client * 3 / 5) {
                    return true;
                }
            }
        }

        false
    }
}

/// Estado interno e visual associado a cada janela principal de documento do Office
pub struct EstadoJanela {
    pub hwnd: HWND,
    pub tab_control: Option<TabControl>,
    pub hwnd_doc_view: Option<HWND>,
    pub y_base: i32,
    pub doc_hwnds: Vec<HWND>,
    pub titulos_cache: Vec<String>,
    pub ult_largura: i32,
    pub ult_altura: i32,
    pub ult_y_ribbon: i32,
    pub ult_y_doc: i32,
    pub ult_x_doc: i32,
}

unsafe impl Send for EstadoJanela {}
unsafe impl Sync for EstadoJanela {}

// ---------------------------------------------------------------------------
// Cálculo Dinâmico da Posição do Ribbon
// ---------------------------------------------------------------------------

/// Contexto para coletar a borda inferior da barra de fórmulas do Excel durante EnumChildWindows.
struct ContextoFormulaBar {
    hwnd_top: HWND,
    bottom_cliente: i32,
}

/// Callback de EnumChildWindows: procura classes EXCEL; (Name Box) e EXCEL< (editor fx)
/// em TODOS os descendentes (não apenas filhos diretos) e coleta a maior borda inferior
/// convertida para coordenadas de cliente da janela principal.
unsafe extern "system" fn callback_formula_bar(child: HWND, lparam: LPARAM) -> BOOL {
    unsafe {
        let ctx = &mut *(lparam.0 as *mut ContextoFormulaBar);
        let mut nome_buf = [0u16; 32];
        let n = GetClassNameW(child, &mut nome_buf);
        if n > 0 {
            let nome = String::from_utf16_lossy(&nome_buf[..n as usize]);
            if nome == "EXCEL;" || nome == "EXCEL<" {
                if IsWindowVisible(child).as_bool() {
                    let mut rc = RECT::default();
                    if GetWindowRect(child, &mut rc).is_ok() {
                        let mut pt = POINT {
                            x: rc.left,
                            y: rc.top,
                        };
                        if ScreenToClient(ctx.hwnd_top, &mut pt).as_bool() {
                            let bottom = pt.y + (rc.bottom - rc.top);
                            if bottom > ctx.bottom_cliente {
                                ctx.bottom_cliente = bottom;
                            }
                        }
                    }
                }
            }
        }
        true.into() // continuar a enumeração (BOOL)
    }
}

/// Localiza a borda inferior (em coordenadas de cliente da janela principal) da barra
/// de fórmulas do Excel, buscando recursivamente por EXCEL; / EXCEL< em toda a árvore
/// de janelas. Retorna None se nenhuma barra de fórmulas visível for encontrada.
fn localizar_fim_formula_bar_excel(hwnd_top: HWND) -> Option<i32> {
    let mut ctx = ContextoFormulaBar {
        hwnd_top,
        bottom_cliente: 0,
    };
    unsafe {
        let ok = EnumChildWindows(
            hwnd_top,
            Some(callback_formula_bar),
            LPARAM(&mut ctx as *mut ContextoFormulaBar as isize),
        );
        if ok.as_bool() && ctx.bottom_cliente > 0 {
            Some(ctx.bottom_cliente)
        } else {
            None
        }
    }
}

/// Calcula a coordenada Y logo abaixo do Ribbon para ancorar a barra de abas.
/// Detecta dinamicamente DropShadow, MsoCommandBarDock e NetUIHWND no topo,
/// aplicando verificação completa com DPI awareness e limites estritos da Ribbon.
pub fn calcular_y_ribbon(hwnd_top: HWND) -> i32 {
    let dpi = crate::utils::get_dpi_for_window(hwnd_top);
    let dpi_safe = if dpi == 0 { 96 } else { dpi };
    let escala = |px: i32| -> i32 { ((px * dpi_safe as i32) + 48) / 96 };

    unsafe {
        // No Microsoft Excel, se a Barra de Fórmulas (EXCEL< / EXCEL;) estiver visível,
        // a barra de abas deve ficar posicionada IMEDIATAMENTE ABAIXO da barra de fórmulas,
        // para não obstruir o editor de fórmulas fx nem o Name Box.
        let classe_top = extrair_classe_janela(hwnd_top);
        if obter_tipo_app(&classe_top) == TipoAppOffice::Excel {
            // Busca RECURSIVA (EnumChildWindows) por EXCEL; / EXCEL< em toda a árvore
            // de descendentes — as versões modernas do Excel não as expõem como filhos
            // diretos de XLMAIN, o que fazia a detecção antiga falhar e cair no
            // fallback escala(148) (barra sobreposta à Ribbon e área errada).
            if let Some(bottom) = localizar_fim_formula_bar_excel(hwnd_top) {
                if bottom > escala(20) && bottom < escala(350) {
                    return bottom.clamp(escala(20), escala(350));
                }
            }
        }

        // 1. Procurar DropShadow com pt.y >= escala(20) e dentro dos limites da Ribbon
        let mut child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, w!("DropShadow"), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                if GetWindowRect(child, &mut rc).is_ok() {
                    let mut pt = POINT { x: rc.left, y: rc.top };
                    if ScreenToClient(hwnd_top, &mut pt).as_bool() {
                        let h = rc.bottom - rc.top;
                        if pt.y >= escala(20) && pt.y < escala(250) && pt.x.abs() <= escala(4) {
                            return (pt.y + h).clamp(escala(25), escala(280)); // Borda inferior do DropShadow
                        }
                    }
                }
            }
        }

        // 2. Procurar MsoCommandBarDock no topo com limites da Ribbon
        child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, w!("MsoCommandBarDock"), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                if GetWindowRect(child, &mut rc).is_ok() {
                    let mut pt = POINT { x: rc.left, y: rc.top };
                    if ScreenToClient(hwnd_top, &mut pt).as_bool() {
                        let h = rc.bottom - rc.top;
                        if pt.x.abs() <= escala(4)
                            && pt.y <= escala(25)
                            && h >= escala(20)
                            && h < escala(250)
                        {
                            return (pt.y + h).clamp(escala(25), escala(280));
                        }
                    }
                }
            }
        }

        // 3. Procurar NetUIHWND no topo com limites da Ribbon
        child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, w!("NetUIHWND"), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                if GetWindowRect(child, &mut rc).is_ok() {
                    let mut pt = POINT { x: rc.left, y: rc.top };
                    if ScreenToClient(hwnd_top, &mut pt).as_bool() {
                        let w = rc.right - rc.left;
                        let h = rc.bottom - rc.top;
                        if pt.x.abs() <= escala(4)
                            && pt.y <= escala(25)
                            && w > escala(200)
                            && h >= escala(20)
                            && h < escala(250)
                        {
                            return (pt.y + h).clamp(escala(25), escala(280));
                        }
                    }
                }
            }
        }

        // 4. Fallback pelo topo da área de documento (apenas se a barra de abas ainda não foi criada)
        let tab_ja_existe = !GetDlgItem(hwnd_top, OFFICETAB_CONTROL_ID as i32)
            .unwrap_or_default()
            .is_invalid();
        if !tab_ja_existe
            && let Some(doc_view) = localizar_area_documento(hwnd_top)
        {
            let mut rc = RECT::default();
            if GetWindowRect(doc_view, &mut rc).is_ok() {
                let mut pt = POINT { x: rc.left, y: rc.top };
                if ScreenToClient(hwnd_top, &mut pt).as_bool() {
                    if pt.y >= escala(25) && pt.y < escala(250) {
                        return pt.y.clamp(escala(25), escala(250));
                    }
                }
            }
        }

        // Fallback padrão ajustado para DPI (148px base em 96 DPI do Office moderno com Ribbon expandido)
        escala(148).clamp(20, 350)
    }
}

// ---------------------------------------------------------------------------
// TabManager — opera SOMENTE na thread de UI
// ---------------------------------------------------------------------------

/// Identificadores de comandos do menu de contexto de abas
pub const IDM_FECHAR: usize = 3001;
pub const IDM_FECHAR_OUTRAS: usize = 3002;
pub const IDM_FECHAR_DIREITA: usize = 3003;
pub const IDM_COPIAR_CAMINHO: usize = 3004;
pub const IDM_ABRIR_PASTA: usize = 3005;
pub const IDM_NOVA_ABA: usize = 3006;
pub const IDM_FECHAR_TODAS: usize = 3007;

/// Gerenciador central de abas e janelas do Office.
pub struct TabManager {
    subclass_manager: SubclassManager,
    estados: HashMap<isize, EstadoJanela>,
    pub config: OfficeTabConfig,
    ordem_documentos: Vec<HWND>,
}

unsafe impl Send for TabManager {}
unsafe impl Sync for TabManager {}

impl TabManager {
    pub fn new() -> Self {
        let config = crate::persistence::load_effective_config();
        Self {
            subclass_manager: SubclassManager::new(),
            estados: HashMap::new(),
            config,
            ordem_documentos: Vec::new(),
        }
    }

    /// Retorna a altura da barra de abas configurada escalonada pelo DPI da janela especificada.
    pub fn obter_altura_barra(&self, hwnd: HWND) -> i32 {
        let base = if self.config.general.tab_bar_height > 0 {
            self.config.general.tab_bar_height
        } else if self.config.theme.tab_height > 0 {
            self.config.theme.tab_height
        } else {
            ALTURA_BARRA_ABAS
        };
        let altura = crate::tab_control::calcular_altura_por_dpi(hwnd, base);
        altura.max(20).min(96)
    }

    /// Recarrega as configurações persistidas do registro (temas, posicionamento, etc.).
    pub fn recarregar_configuracao(&mut self) {
        self.config = crate::persistence::load_effective_config();
    }

    /// Descobre janelas visíveis das aplicações do Microsoft Office.
    pub fn discover_document_windows(&mut self) -> Vec<JanelaDocumento> {
        listar_todas_janelas_office()
    }

    /// Aplica o subclassing a uma janela do Office e a registra.
    /// DEVE ser chamado da thread de UI do Office.
    pub fn monitor_window(&mut self, hwnd: HWND) -> Result<()> {
        let raw_hwnd = hwnd.0 as isize;
        if self.estados.contains_key(&raw_hwnd) {
            return Ok(());
        }

        log_debug(&format!(
            "monitor_window (thread UI): HWND=0x{:X}",
            hwnd.0 as usize
        ));

        // Assegurar WS_CLIPCHILDREN na janela moldura para não repintar sobre o TabControl
        unsafe {
            let estilo = GetWindowLongPtrW(hwnd, GWL_STYLE);
            let _ = SetWindowLongPtrW(hwnd, GWL_STYLE, estilo | (WS_CLIPCHILDREN.0 as isize));
        }

        self.subclass_manager.subclass_window(hwnd)?;

        self.estados.insert(
            raw_hwnd,
            EstadoJanela {
                hwnd,
                tab_control: None,
                hwnd_doc_view: None,
                y_base: 0,
                doc_hwnds: Vec::new(),
                titulos_cache: Vec::new(),
                ult_largura: 0,
                ult_altura: 0,
                ult_y_ribbon: 0,
                ult_y_doc: 0,
                ult_x_doc: 0,
            },
        );

        // Ativar timer periódico de atualização (800ms)
        unsafe {
            let _ = SetTimer(hwnd, OFFICETAB_TIMER_ID, 800, None);
        }

        // Sincronização inicial imediata
        let res = self.sincronizar_janela(hwnd);
        log_debug(&format!(
            "monitor_window: sincronizar res={:?}",
            res.is_ok()
        ));

        res
    }

    /// Alterna suavemente para a janela de documento selecionada no Modo de Janela Única.
    /// Exibe e focaliza o documento de destino e oculta a janela anterior (SW_HIDE),
    /// garantindo que apenas uma única janela do Word permaneça visível na tela
    /// e na barra de tarefas do Windows (Single Window Mode / Single Taskbar Button).
    pub fn alternar_para_janela(&self, hwnd_origem: HWND, hwnd_alvo: HWND) {
        unsafe {
            if IsIconic(hwnd_alvo).as_bool() {
                let _ = ShowWindow(hwnd_alvo, SW_RESTORE);
            }

            if IsZoomed(hwnd_origem).as_bool() {
                let _ = ShowWindow(hwnd_alvo, SW_MAXIMIZE);
            } else {
                let mut rc = RECT::default();
                if GetWindowRect(hwnd_origem, &mut rc).is_ok() {
                    let _ = SetWindowPos(
                        hwnd_alvo,
                        HWND::default(),
                        rc.left,
                        rc.top,
                        rc.right - rc.left,
                        rc.bottom - rc.top,
                        SWP_NOZORDER,
                    );
                }
            }
            let _ = ShowWindow(hwnd_alvo, SW_SHOW);
            let _ = BringWindowToTop(hwnd_alvo);
            let _ = SetForegroundWindow(hwnd_alvo);

            // NOTA: NÃO forjamos WM_ACTIVATE via SendMessageW. O SetForegroundWindow
            // acima gera a ativação REAL com wParam/lParam corretos. O WM_ACTIVATE
            // forjado com lParam=0 desreferenciava handle nulo nos internos MDI do
            // Excel e derrubava o processo ao alternar abas (Access Violation).

            // IMPORTANTE (Excel 2013+ SDI / Word MDI): NÃO ocultamos a janela de origem
            // com SW_HIDE. No Excel SDI cada pasta de trabalho é uma janela XLMAIN
            // top-level que CONTÉM a própria Ribbon/menus; ocultá-la faz a interface
            // inteira sumir da tela ("menus somem"). Mantemos todas as janelas visíveis
            // e apenas trazemos o documento de destino para frente/ativo.
        }

        log_debug(&format!(
            "Alternado foco de HWND=0x{:X} para HWND=0x{:X} (Modo Janela Única)",
            hwnd_origem.0 as usize, hwnd_alvo.0 as usize
        ));
    }

    /// Sincroniza o controle visual de abas e redimensionamento.
    /// Suporta os 7 aplicativos do Office (Word, Excel, PowerPoint, Access, Project, Visio, Publisher).
    /// DEVE ser chamado da thread de UI do Office.
    pub fn sincronizar_janela(&mut self, hwnd: HWND) -> Result<()> {
        unsafe {
            if !IsWindow(hwnd).as_bool() {
                log_debug(&format!("sincronizar_janela: HWND inválido 0x{:X}", hwnd.0 as usize));
                self.estados.remove(&(hwnd.0 as isize));
                return Ok(());
            }

            let mut rc_client = RECT::default();
            if GetClientRect(hwnd, &mut rc_client).is_err() {
                log_debug(&format!("sincronizar_janela: GetClientRect falhou para HWND 0x{:X}", hwnd.0 as usize));
                return Ok(());
            }

            let largura_cliente = rc_client.right - rc_client.left;
            let altura_cliente = rc_client.bottom - rc_client.top;

            if largura_cliente <= 100 || altura_cliente <= 100 {
                log_debug(&format!("sincronizar_janela: dimensões muito pequenas {}x{} para HWND 0x{:X}", largura_cliente, altura_cliente, hwnd.0 as usize));
                return Ok(());
            }

            // Verificar se a janela do Office está na tela inicial ou Backstage (Novo/Abrir/Opções)
            // NOTA: Para Excel, a detecção de backstage pode falhar e esconder as abas incorretamente.
            // Usamos uma verificação mais robusta que não depende de localizar_area_documento.
            let classe_top = extrair_classe_janela(hwnd);
            let app_tipo = obter_tipo_app(&classe_top);
            let is_backstage = if app_tipo == TipoAppOffice::Excel {
                // Para Excel, verificar apenas se há NetUIHWND grande cobrindo a tela
                // (Backstage real), sem depender de localizar_area_documento
                let mut rc_client = RECT::default();
                if GetClientRect(hwnd, &mut rc_client).is_err() {
                    false
                } else {
                    let w_client = rc_client.right - rc_client.left;
                    let h_client = rc_client.bottom - rc_client.top;
                    let mut child = HWND::default();
                    let mut found_big_netui = false;
                    loop {
                        child = match FindWindowExW(hwnd, child, w!("NetUIHWND"), PCWSTR::null()) {
                            Ok(h) if !h.is_invalid() => h,
                            _ => break,
                        };
                        if IsWindowVisible(child).as_bool() {
                            let mut rc = RECT::default();
                            if GetWindowRect(child, &mut rc).is_ok() {
                                let w = rc.right - rc.left;
                                let h = rc.bottom - rc.top;
                                if w > (w_client * 3 / 4) && h > (h_client * 3 / 5) {
                                    found_big_netui = true;
                                    break;
                                }
                            }
                        }
                    }
                    found_big_netui
                }
            } else {
                backstage_esta_aberto(hwnd)
            };
            log_debug(&format!("sincronizar_janela: backstage_esta_aberto={} para HWND 0x{:X} ({:?})", is_backstage, hwnd.0 as usize, app_tipo));
            if is_backstage {
                if let Some(estado) = self.estados.get_mut(&(hwnd.0 as isize))
                    && let Some(ref tab) = estado.tab_control
                {
                    let _ = ShowWindow(tab.hwnd, SW_HIDE);
                }
                return Ok(());
            }

            // Documento em edição ativo: garantir que o controle de abas esteja visível
            if let Some(estado) = self.estados.get_mut(&(hwnd.0 as isize))
                && let Some(ref tab) = estado.tab_control
                && !IsWindowVisible(tab.hwnd).as_bool()
            {
                let _ = ShowWindow(tab.hwnd, SW_SHOW);
            }

            // Obter altura da barra escalonada por DPI e configurações de fonte
            let altura_barra = self.obter_altura_barra(hwnd);
            let font_name = self.config.theme.font_name.clone();
            let font_size = self.config.theme.font_size;

            let estado = match self.estados.get_mut(&(hwnd.0 as isize)) {
                Some(e) => e,
                None => return Ok(()),
            };

            // Localizar janela interna de documento nos 7 aplicativos suportados
            let doc_view = localizar_area_documento(hwnd);
            estado.hwnd_doc_view = doc_view;

            // Calcular y_base DINAMICAMENTE para acompanhar ribbon recolhido/expandido com DPI awareness
            let y_ribbon_raw = calcular_y_ribbon(hwnd);
            log_debug(&format!("sincronizar_janela: y_ribbon_raw={} para HWND 0x{:X}", y_ribbon_raw, hwnd.0 as usize));
            let dpi = crate::utils::get_dpi_for_window(hwnd);
            let dpi_safe = if dpi == 0 { 96 } else { dpi };
            let escala = |px: i32| -> i32 { ((px * dpi_safe as i32) + 48) / 96 };
            let y_ribbon = y_ribbon_raw.clamp(escala(20), altura_cliente - 40).max(escala(20));
            log_debug(&format!("sincronizar_janela: y_ribbon_clamped={} (raw={}, altura_cliente={}) para HWND 0x{:X}", y_ribbon, y_ribbon_raw, altura_cliente, hwnd.0 as usize));
            estado.y_base = y_ribbon;

            let layout_mudou = largura_cliente != estado.ult_largura
                || altura_cliente != estado.ult_altura
                || y_ribbon != estado.ult_y_ribbon;

            // Criar controle de abas se ainda não existir
            if estado.tab_control.is_none() {
                match TabControl::new(
                    hwnd,
                    OFFICETAB_CONTROL_ID,
                    0,
                    y_ribbon,
                    largura_cliente,
                    altura_barra,
                ) {
                    Ok(mut tab) => {
                        tab.aplicar_fonte_tema(&font_name, font_size);
                        log_debug(&format!(
                            "TabControl CRIADO (thread UI): HWND=0x{:X}, pai=0x{:X}, y={}, h={}",
                            tab.hwnd.0 as usize, hwnd.0 as usize, y_ribbon, altura_barra
                        ));
                        estado.tab_control = Some(tab);
                        estado.ult_largura = largura_cliente;
                        estado.ult_altura = altura_cliente;
                        estado.ult_y_ribbon = y_ribbon;
                    }
                    Err(e) => {
                        log_debug(&format!("ERRO AO CRIAR TabControl: {:?}", e));
                        return Err(e);
                    }
                }
            } else if layout_mudou && let Some(ref tab) = estado.tab_control {
                let _ = tab.resize(0, y_ribbon, largura_cliente, altura_barra);
                estado.ult_largura = largura_cliente;
                estado.ult_altura = altura_cliente;
                estado.ult_y_ribbon = y_ribbon;
            }

            // Manter o TabControl no topo da Z-order: o FullpageUIHost do Excel cobre
            // toda a área do cliente (incluindo a barra de abas) e precisa ser superado
            // a cada sincronização, senão a barra fica invisível.
            if let Some(ref tab) = estado.tab_control {
                unsafe {
                    let _ = SetWindowPos(
                        tab.hwnd,
                        HWND_TOP,
                        0,
                        0,
                        0,
                        0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                    );
                    // Forçar invalidação da barra logo após trazer ao topo
                    let _ = InvalidateRect(tab.hwnd, None, true);
                    let _ = UpdateWindow(tab.hwnd);
                    // Garantir que a barra fique acima do documento
                    if let Some(dv) = estado.hwnd_doc_view {
                        let _ = SetWindowPos(dv, tab.hwnd, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    }
                }
            }

            // Deslocar área do documento para baixo da barra de abas apenas quando o layout mudar
            if let Some(dv) = estado.hwnd_doc_view
            {
                let y_doc = y_ribbon + altura_barra;
                let layout_mudou_ou_dimensao = layout_mudou
                    || y_doc != estado.ult_y_doc
                    || largura_cliente != estado.ult_x_doc;
                if layout_mudou_ou_dimensao || estado.hwnd_doc_view.is_some() {
                    let visivel = unsafe { IsWindowVisible(dv).as_bool() };
                    if visivel {
                        let h_status = calcular_altura_status_bar(hwnd, altura_cliente);
                        let h_doc = (altura_cliente - y_doc - h_status).max(10);
                        if y_doc >= 0 && h_doc > 0 && largura_cliente > 0 && altura_cliente > 0 {
                            let _ = MoveWindow(dv, 0, y_doc, largura_cliente, h_doc, true);
                        }
                        estado.ult_y_doc = y_doc;
                        estado.ult_x_doc = largura_cliente;
                    }
                }
            }

            // Popular abas com documentos abertos
            // As janelas não ativas do Word ficam ocultas no modo de abas,
            // mas continuam sendo documentos abertos e precisam permanecer
            // representadas na barra.
            let mut docs_processo = listar_janelas_processo();
            if docs_processo.is_empty() {
                let tit = extrair_texto_janela(hwnd);
                let tit_limpo = if tit.is_empty() {
                    "Documento".to_string()
                } else {
                    limpar_titulo_documento(&tit)
                };
                docs_processo.push((hwnd, tit_limpo));
            }

            // Preserva rigorosamente a posição estável das abas (não altera a ordem ao alternar foco):
            let hwnds_atuais: Vec<HWND> = docs_processo.iter().map(|(h, _)| *h).collect();
            self.ordem_documentos.retain(|h| hwnds_atuais.contains(h));

            for &h in &hwnds_atuais {
                if !self.ordem_documentos.contains(&h) {
                    self.ordem_documentos.push(h);
                }
            }

            let mut docs_estaveis: Vec<(HWND, String)> = Vec::new();
            for &h in &self.ordem_documentos {
                if let Some((_, tit)) = docs_processo.iter().find(|(dh, _)| *dh == h) {
                    docs_estaveis.push((h, tit.clone()));
                }
            }

            let novos_titulos: Vec<String> = docs_estaveis.iter().map(|(_, t)| t.clone()).collect();
            let novos_hwnds: Vec<HWND> = docs_estaveis.iter().map(|(h, _)| *h).collect();

            if let Some(ref tab) = estado.tab_control {
                if novos_titulos != estado.titulos_cache || novos_hwnds != estado.doc_hwnds {
                    tab.clear_tabs();

                    for (idx, (doc_hwnd, doc_titulo)) in docs_estaveis.iter().enumerate() {
                        let _ = tab.add_tab(idx as i32, doc_titulo);
                        if *doc_hwnd == hwnd {
                            let _ = tab.set_selected_tab(idx as i32);
                        }
                    }

                    // Aba "+" para novo documento
                    let _ = tab.add_tab(docs_estaveis.len() as i32, "+");

                    estado.doc_hwnds = novos_hwnds;
                    estado.titulos_cache = novos_titulos;
                    tab.repaint();
                } else if let Some(pos) = estado.doc_hwnds.iter().position(|&h| h == hwnd)
                    && tab.get_selected_tab() != pos as i32
                {
                    let _ = tab.set_selected_tab(pos as i32);
                    tab.repaint();
                }
            }

            Ok(())
        }
    }

    /// Processa notificações do TabControl (clique em aba).
    pub fn processar_notify(&mut self, hwnd: HWND, lparam: LPARAM) -> Option<LRESULT> {
        if lparam.0 == 0 {
            return None;
        }

        let pnmhdr = unsafe { &*(lparam.0 as *const NMHDR) };
        if pnmhdr.idFrom != OFFICETAB_CONTROL_ID as usize || pnmhdr.code != TCN_SELCHANGE {
            if pnmhdr.idFrom == OFFICETAB_CONTROL_ID as usize && pnmhdr.code == NM_CUSTOMDRAW {
                let draw = unsafe { &*(lparam.0 as *const NMCUSTOMDRAW) };
                return self
                    .processar_custom_draw(hwnd, draw)
                    .map(|result| LRESULT(result as isize));
            }

            // Notificações de Tooltip (TTN_GETDISPINFOW / TTN_GETDISPINFOA)
            const TTN_GETDISPINFOW: u32 = 0xFFFF_FDEE; // -530
            const TTN_GETDISPINFOA: u32 = 0xFFFF_FDF8; // -520
            if pnmhdr.code == TTN_GETDISPINFOW || pnmhdr.code == TTN_GETDISPINFOA {
                return self.processar_tooltip_notify(hwnd, lparam);
            }

            return None;
        }

        let estado = self.estados.get(&(hwnd.0 as isize))?;
        let tab = estado.tab_control.as_ref()?;
        let (sel, total_docs, doc_hwnds) = (
            tab.get_selected_tab(),
            estado.doc_hwnds.len(),
            estado.doc_hwnds.clone(),
        );

        if sel >= 0 && (sel as usize) < total_docs {
            let alvo = doc_hwnds[sel as usize];
            if alvo != hwnd {
                static ULTIMA_TROCA: Mutex<Option<(isize, std::time::Instant)>> = Mutex::new(None);
                let agora = std::time::Instant::now();
                if let Ok(mut lock) = ULTIMA_TROCA.lock() {
                    if let Some((ult_alvo, ult_tempo)) = *lock
                        && ult_alvo == alvo.0 as isize
                        && agora.duration_since(ult_tempo).as_millis() < 250
                    {
                        return Some(LRESULT(0));
                    }
                    *lock = Some((alvo.0 as isize, agora));
                }

                // ALTERNÂNCIA DIFERIDA: NÃO tocamos nas janelas dentro do WM_NOTIFY.
                // Postamos WM_OFFICETAB_ALTERNAR para a janela de DESTINO; ela fará
                // mostrar/ocultar/ativar no próprio contexto de mensagem, fora do
                // handler aninhado — elimina a reentrância e o WM_ACTIVATE forjado
                // (lParam=0) que derrubava o Excel ao alternar abas.
                unsafe {
                    let _ = PostMessageW(
                        alvo,
                        WM_OFFICETAB_ALTERNAR,
                        WPARAM(hwnd.0 as usize),
                        LPARAM(0),
                    );
                }
            }
            return Some(LRESULT(0));
        } else if (sel as usize) == total_docs {
            // Clicou na aba "+"
            unsafe {
                let _ = PostMessageW(hwnd, WM_OFFICETAB_NOVO_DOC, WPARAM(0), LPARAM(0));
            }
            if let Some(estado) = self.estados.get(&(hwnd.0 as isize))
                && let Some(ref tab) = estado.tab_control
                && let Some(pos) = doc_hwnds.iter().position(|&h| h == hwnd)
            {
                let _ = tab.set_selected_tab(pos as i32);
            }
            return Some(LRESULT(0));
        }
        None
    }

    fn processar_custom_draw(&self, hwnd: HWND, draw: &NMCUSTOMDRAW) -> Option<u32> {
        let estado = self.estados.get(&(hwnd.0 as isize))?;
        let tab = estado.tab_control.as_ref()?;
        if draw.hdr.hwndFrom != tab.hwnd {
            return None;
        }
        let theme = self.tema_da_janela(hwnd);

        // Diagnóstico throttled (primeiras 3 pinturas) para confirmar rota + show_close + rect
        static LOG_COUNT: AtomicU32 = AtomicU32::new(0);
        let _n = LOG_COUNT.fetch_add(1, Ordering::Relaxed);
        if _n < 6 {
            log_debug(&format!(
                "CUSTOM_DRAW stage={} item={} show_close={} rc=({},{},{},{}) w={} h={}",
                draw.dwDrawStage.0,
                draw.dwItemSpec,
                theme.show_close_button,
                draw.rc.left,
                draw.rc.top,
                draw.rc.right,
                draw.rc.bottom,
                draw.rc.right - draw.rc.left,
                draw.rc.bottom - draw.rc.top
            ));
        }

        if draw.dwDrawStage == CDDS_PREPAINT {
            unsafe {
                // Preenche o fundo da barra de abas
                let brush = CreateSolidBrush(COLORREF(crate::theme::Theme::to_colorref(
                    theme.bar_background,
                )));
                let _ = FillRect(draw.hdc, &draw.rc, brush);
                let _ = DeleteObject(brush);

                // Linha de base / divisor inferior contínuo em toda a barra (1px)
                let border_color = COLORREF(crate::theme::Theme::to_colorref(theme.tab_border));
                let border_brush = CreateSolidBrush(border_color);
                let mut line_rc = draw.rc;
                line_rc.top = line_rc.bottom - 1;
                let _ = FillRect(draw.hdc, &line_rc, border_brush);
                let _ = DeleteObject(border_brush);
            }
            return Some(CDRF_NOTIFYITEMDRAW);
        }

        if draw.dwDrawStage != CDDS_ITEMPREPAINT {
            return None;
        }

        let tab_idx = draw.dwItemSpec;
        let total_docs = estado.doc_hwnds.len();
        let is_plus = tab_idx == total_docs;

        let selected = (draw.uItemState.0 & CDIS_SELECTED.0) != 0;
        let hot = (draw.uItemState.0 & CDIS_HOT.0) != 0;
        let rc = draw.rc;

        let accent_color = theme.custom_accent_color.unwrap_or(0x00185ABD);
        let border_cr = COLORREF(crate::theme::Theme::to_colorref(theme.tab_border));

        unsafe {
            if is_plus {
                // -----------------------------------------------------------
                // Botão "+" (Novo Documento) com design moderno flutuante
                // -----------------------------------------------------------
                let btn_rc = RECT {
                    left: rc.left + 3,
                    top: rc.top + 3,
                    right: rc.right - 3,
                    bottom: rc.bottom - 2,
                };

                let bg_color = if hot {
                    theme.hover_tab_bg
                } else {
                    theme.bar_background
                };

                let brush = CreateSolidBrush(COLORREF(crate::theme::Theme::to_colorref(bg_color)));
                let _ = FillRect(draw.hdc, &btn_rc, brush);
                let _ = DeleteObject(brush);

                if hot {
                    // Borda sutil de destaque ao passar o mouse
                    let border_brush = CreateSolidBrush(border_cr);
                    let mut b_top = btn_rc;
                    b_top.bottom = b_top.top + 1;
                    let _ = FillRect(draw.hdc, &b_top, border_brush);

                    let mut b_bottom = btn_rc;
                    b_bottom.top = b_bottom.bottom - 1;
                    let _ = FillRect(draw.hdc, &b_bottom, border_brush);

                    let mut b_left = btn_rc;
                    b_left.right = b_left.left + 1;
                    let _ = FillRect(draw.hdc, &b_left, border_brush);

                    let mut b_right = btn_rc;
                    b_right.left = b_right.right - 1;
                    let _ = FillRect(draw.hdc, &b_right, border_brush);

                    let _ = DeleteObject(border_brush);
                }

                let text_color = if hot {
                    accent_color
                } else {
                    theme.inactive_tab_text
                };

                let _ = SetBkMode(draw.hdc, TRANSPARENT);
                let _ = SetTextColor(
                    draw.hdc,
                    COLORREF(crate::theme::Theme::to_colorref(text_color)),
                );

                let mut plus_str: [u16; 1] = [b'+' as u16];
                let mut text_rc = btn_rc;
                let _ = DrawTextW(
                    draw.hdc,
                    &mut plus_str,
                    &mut text_rc,
                    DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX,
                );
            } else {
                // -----------------------------------------------------------
                // Abas de Documentos (Design Premium Fluent / Modern Office)
                // -----------------------------------------------------------
                let bg_color = if selected {
                    theme.active_tab_bg
                } else if hot {
                    theme.hover_tab_bg
                } else {
                    theme.inactive_tab_bg
                };

                let brush = CreateSolidBrush(COLORREF(crate::theme::Theme::to_colorref(bg_color)));
                let _ = FillRect(draw.hdc, &rc, brush);
                let _ = DeleteObject(brush);

                if selected {
                    // Indicador de destaque premium (Accent Bar de 3px no topo)
                    let accent_brush =
                        CreateSolidBrush(COLORREF(crate::theme::Theme::to_colorref(accent_color)));
                    let accent_rc = RECT {
                        left: rc.left,
                        top: rc.top,
                        right: rc.right,
                        bottom: rc.top + 3,
                    };
                    let _ = FillRect(draw.hdc, &accent_rc, accent_brush);
                    let _ = DeleteObject(accent_brush);

                    // Bordas laterais sutis (esquerda e direita)
                    let border_brush = CreateSolidBrush(border_cr);
                    let border_left = RECT {
                        left: rc.left,
                        top: rc.top + 3,
                        right: rc.left + 1,
                        bottom: rc.bottom,
                    };
                    let _ = FillRect(draw.hdc, &border_left, border_brush);

                    let border_right = RECT {
                        left: rc.right - 1,
                        top: rc.top + 3,
                        right: rc.right,
                        bottom: rc.bottom,
                    };
                    let _ = FillRect(draw.hdc, &border_right, border_brush);
                    let _ = DeleteObject(border_brush);

                    // A base da aba ativa conecta perfeitamente ao documento
                    let active_base_brush = CreateSolidBrush(COLORREF(
                        crate::theme::Theme::to_colorref(theme.active_tab_bg),
                    ));
                    let base_rc = RECT {
                        left: rc.left + 1,
                        top: rc.bottom - 1,
                        right: rc.right - 1,
                        bottom: rc.bottom,
                    };
                    let _ = FillRect(draw.hdc, &base_rc, active_base_brush);
                    let _ = DeleteObject(active_base_brush);
                } else {
                    // Aba inativa: divisor vertical sutil à direita para separar abas
                    let div_brush = CreateSolidBrush(border_cr);
                    let div_rc = RECT {
                        left: rc.right - 1,
                        top: rc.top + 5,
                        right: rc.right,
                        bottom: rc.bottom - 5,
                    };
                    let _ = FillRect(draw.hdc, &div_rc, div_brush);
                    let _ = DeleteObject(div_brush);
                }

                // Título do Documento com ClearType e alinhamento impecável
                let titulo = if tab_idx < estado.titulos_cache.len() {
                    &estado.titulos_cache[tab_idx]
                } else {
                    "Documento"
                };

                let show_close = theme.show_close_button;
                let text_right = if show_close {
                    rc.right - 34
                } else {
                    rc.right - 8
                };
                let mut text_rc = RECT {
                    left: rc.left + 12,
                    top: rc.top,
                    right: text_right.max(rc.left + 12),
                    bottom: rc.bottom,
                };

                let text_color = if selected {
                    theme.active_tab_text
                } else if hot {
                    theme.hover_tab_text
                } else {
                    theme.inactive_tab_text
                };

                let _ = SetBkMode(draw.hdc, TRANSPARENT);
                let _ = SetTextColor(
                    draw.hdc,
                    COLORREF(crate::theme::Theme::to_colorref(text_color)),
                );

                let mut utf16: Vec<u16> = titulo.encode_utf16().collect();
                let _ = DrawTextW(
                    draw.hdc,
                    &mut utf16,
                    &mut text_rc,
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX,
                );

                // Botão de fechar elegante ("×") — SEMPRE visível em todas as abas
                // de documento (selecionadas, com hover ou não), como no Office Tab comercial
                if show_close {
                    let close_size = 20;
                    let close_x = rc.right - close_size - 8;
                    let close_y = rc.top + (rc.bottom - rc.top - close_size) / 2;
                    let mut close_rc = RECT {
                        left: close_x,
                        top: close_y,
                        right: close_x + close_size,
                        bottom: close_y + close_size,
                    };

                    // Hit test no cursor do mouse para detecção de hover em tempo real
                    let mut pt_cursor = POINT::default();
                    let mut x_hover = false;
                    if GetCursorPos(&mut pt_cursor).is_ok() {
                        let _ = ScreenToClient(draw.hdr.hwndFrom, &mut pt_cursor);
                        x_hover = pt_cursor.x >= (close_x - 3)
                            && pt_cursor.x <= (close_x + close_size + 3)
                            && pt_cursor.y >= (close_y - 3)
                            && pt_cursor.y <= (close_y + close_size + 3);
                    }

                    if x_hover {
                        // Fluent Close Button Hover: badge suavemente avermelhado com cantos arredondados
                        let bg_hover_rc = RECT {
                            left: close_x - 2,
                            top: close_y - 2,
                            right: close_x + close_size + 2,
                            bottom: close_y + close_size + 2,
                        };
                        let red_bg_brush = CreateSolidBrush(COLORREF(
                            crate::theme::Theme::to_colorref(0x00FDE7E9),
                        ));
                        let _ = FillRect(draw.hdc, &bg_hover_rc, red_bg_brush);
                        let _ = DeleteObject(red_bg_brush);

                        // Glyph '×' em vermelho carmesim Fluent vivo
                        let _ = SetTextColor(
                            draw.hdc,
                            COLORREF(crate::theme::Theme::to_colorref(0x00C42B1C)),
                        );
                    } else {
                        let close_color = if selected {
                            0x00507A5E // Verde suave elegante harmonizado com o tom esverdeado da aba ativa
                        } else {
                            0x008A8886 // Cinza sutil para aba inativa
                        };
                        let _ = SetTextColor(
                            draw.hdc,
                            COLORREF(crate::theme::Theme::to_colorref(close_color)),
                        );
                    }

                    let mut x_glyph: [u16; 1] = [0x00D7]; // '×' (U+00D7)

                    // Fonte maior e em negrito para o '×' ficar bem visível e clicável
                    let hfont_x = CreateFontW(
                        -20, // altura em pixels (negativo = altura de célula)
                        0,
                        0,
                        0,
                        FW_BOLD.0 as i32,
                        0,
                        0,
                        0,
                        DEFAULT_CHARSET.0 as u32,
                        OUT_DEFAULT_PRECIS.0 as u32,
                        CLIP_DEFAULT_PRECIS.0 as u32,
                        CLEARTYPE_QUALITY.0 as u32,
                        (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                        w!("Segoe UI"),
                    );
                    let old_font_x = if !hfont_x.is_invalid() {
                        SelectObject(draw.hdc, hfont_x)
                    } else {
                        HGDIOBJ::default()
                    };

                    let _ = DrawTextW(
                        draw.hdc,
                        &mut x_glyph,
                        &mut close_rc,
                        DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX,
                    );

                    if !hfont_x.is_invalid() {
                        let _ = SelectObject(draw.hdc, old_font_x);
                        let _ = DeleteObject(hfont_x);
                    }
                }
            }
        }

        Some(CDRF_SKIPDEFAULT)
    }

    /// Pinta manualmente o TabControl reutilizando processar_custom_draw com
    /// NMCUSTOMDRAW sintetizados. Usado pelo WM_PAINT do próprio controle porque
    /// no Excel o NM_CUSTOMDRAW enviado ao pai é engolido (ver despachar_tab_paint).
    fn pintar_tab_control(&self, hwnd_tab: HWND, hdc: HDC, _rc_paint: RECT) {
        let mut hwnd_parent = HWND::default();
        let mut tab_count = 0i32;
        let mut selected_idx = -1i32;
        for (raw, estado) in &self.estados {
            if let Some(ref tab) = estado.tab_control
                && tab.hwnd == hwnd_tab
            {
                hwnd_parent = HWND(*raw as *mut _);
                tab_count = tab.get_tab_count();
                selected_idx = tab.get_selected_tab();
                break;
            }
        }
        if hwnd_parent.is_invalid() {
            return;
        }

        // Rect completo do cliente para o fundo da barra (ps.rcPaint pode vir vazio)
        let mut rc_bar = RECT::default();
        unsafe {
            let _ = GetClientRect(hwnd_tab, &mut rc_bar);
        }

        // 1. Fundo da barra (equivale a CDDS_PREPAINT)
        let nmcd_bg = NMCUSTOMDRAW {
            hdr: NMHDR {
                hwndFrom: hwnd_tab,
                idFrom: OFFICETAB_CONTROL_ID as usize,
                code: NM_CUSTOMDRAW,
            },
            dwDrawStage: CDDS_PREPAINT,
            hdc,
            rc: rc_bar,
            dwItemSpec: 0,
            uItemState: NMCUSTOMDRAW_DRAW_STATE_FLAGS(0),
            lItemlParam: LPARAM(0),
        };
        let _ = self.processar_custom_draw(hwnd_parent, &nmcd_bg);

        // 2. Itens (equivale a CDDS_ITEMPREPAINT)
        let mut pt_cursor = POINT::default();
        unsafe {
            if GetCursorPos(&mut pt_cursor).is_ok() {
                let _ = ScreenToClient(hwnd_tab, &mut pt_cursor);
            }
        }

        for i in 0..tab_count {
            let mut item_rc = RECT::default();
            unsafe {
                let _ = SendMessageW(
                    hwnd_tab,
                    crate::tab_control::TCM_GETITEMRECT,
                    WPARAM(i as usize),
                    LPARAM(&mut item_rc as *mut _ as isize),
                );
            }

            let selected = selected_idx == i;
            let hot = pt_cursor.x >= item_rc.left
                && pt_cursor.x < item_rc.right
                && pt_cursor.y >= item_rc.top
                && pt_cursor.y < item_rc.bottom;

            let mut state = 0u32;
            if selected {
                state |= CDIS_SELECTED.0;
            }
            if hot {
                state |= CDIS_HOT.0;
            }

            let nmcd = NMCUSTOMDRAW {
                hdr: NMHDR {
                    hwndFrom: hwnd_tab,
                    idFrom: OFFICETAB_CONTROL_ID as usize,
                    code: NM_CUSTOMDRAW,
                },
                dwDrawStage: CDDS_ITEMPREPAINT,
                hdc,
                rc: item_rc,
                dwItemSpec: i as usize,
                uItemState: NMCUSTOMDRAW_DRAW_STATE_FLAGS(state),
                lItemlParam: LPARAM(0),
            };
            let _ = self.processar_custom_draw(hwnd_parent, &nmcd);
        }
    }

    /// Processa requisições de Tooltip para abas exibindo o caminho completo do arquivo.
    fn processar_tooltip_notify(&self, hwnd: HWND, lparam: LPARAM) -> Option<LRESULT> {
        let estado = self.estados.get(&(hwnd.0 as isize))?;
        let tab = estado.tab_control.as_ref()?;

        let mut pt = POINT::default();
        unsafe {
            if GetCursorPos(&mut pt).is_ok() {
                let _ = ScreenToClient(tab.hwnd, &mut pt);
            }
        }

        #[repr(C)]
        #[allow(non_snake_case, clippy::upper_case_acronyms)]
        struct NMTTDISPINFOW {
            hdr: NMHDR,
            lpszText: *mut u16,
            szText: [u16; 80],
            hinst: windows::Win32::Foundation::HINSTANCE,
            uFlags: u32,
            lParam: LPARAM,
        }

        if let Some(idx) = tab.hit_test(pt.x, pt.y) {
            let total_docs = estado.doc_hwnds.len();
            if idx >= 0 && (idx as usize) < total_docs {
                let titulo = if (idx as usize) < estado.titulos_cache.len() {
                    &estado.titulos_cache[idx as usize]
                } else {
                    "Documento"
                };

                let tooltip_text =
                    if let Some(caminho) = crate::utils::localizar_caminho_documento(titulo) {
                        format!("{}\n{}", titulo, caminho.display())
                    } else {
                        format!("{}\n[Documento Ativo]", titulo)
                    };

                static TOOLTIP_BUFFER: Mutex<Vec<u16>> = Mutex::new(Vec::new());
                if let Ok(mut buf) = TOOLTIP_BUFFER.lock() {
                    *buf = tooltip_text
                        .encode_utf16()
                        .chain(std::iter::once(0))
                        .collect();
                    unsafe {
                        let p_disp = &mut *(lparam.0 as *mut NMTTDISPINFOW);
                        p_disp.lpszText = buf.as_mut_ptr();
                    }
                }
                return Some(LRESULT(0));
            } else if (idx as usize) == total_docs {
                static TOOLTIP_NOVO: Mutex<Vec<u16>> = Mutex::new(Vec::new());
                if let Ok(mut buf) = TOOLTIP_NOVO.lock() {
                    *buf = "Novo Documento (Ctrl+O / Ctrl+T)\0"
                        .encode_utf16()
                        .collect();
                    unsafe {
                        let p_disp = &mut *(lparam.0 as *mut NMTTDISPINFOW);
                        p_disp.lpszText = buf.as_mut_ptr();
                    }
                }
                return Some(LRESULT(0));
            }
        }
        None
    }

    fn tema_da_janela(&self, hwnd: HWND) -> crate::theme::Theme {
        let app_tipo = obter_tipo_app(&extrair_classe_janela(hwnd));
        let app_opt = crate::theme::OfficeApp::from_name(app_tipo.nome());

        let mut theme = if self.config.general.adaptive_theme_by_app
            && let Some(app) = app_opt
        {
            crate::theme::Theme::for_app(app)
        } else {
            self.config.theme.clone()
        };

        // Aplica o design visual com cantos arredondados, realce dinâmico e paleta de alto contraste
        // personalizada de acordo com o aplicativo host do Office (Word, Excel, PowerPoint, Access, Project, Visio, Publisher):
        match app_tipo {
            TipoAppOffice::Word => {
                theme.active_tab_bg = 0x00EBF1FB; // Azul gelo suave pastel
                theme.active_tab_text = 0x00185ABD; // Azul Word nítido
                theme.hover_tab_bg = 0x00F0F5FD; // Hover azul sutil
                theme.tab_border = 0x00B8D0F0; // Borda harmônica
                theme.custom_accent_color = Some(0x00185ABD); // Accent bar azul Word
            }
            TipoAppOffice::Excel => {
                theme.active_tab_bg = 0x00EBF5EF; // Verde sálvia suave pastel
                theme.active_tab_text = 0x000F5132; // Verde Excel floresta
                theme.hover_tab_bg = 0x00F0FAF3; // Hover verde sutil
                theme.tab_border = 0x00C4DFCE; // Borda verde sálvia
                theme.custom_accent_color = Some(0x00107C41); // Accent bar verde esmeralda
            }
            TipoAppOffice::PowerPoint => {
                theme.active_tab_bg = 0x00FDF0EC; // Pêssego/carmim suave pastel
                theme.active_tab_text = 0x00A02C0E; // Carmim PowerPoint
                theme.hover_tab_bg = 0x00FFF5F2; // Hover pêssego sutil
                theme.tab_border = 0x00F5C8BC; // Borda carmim pastel
                theme.custom_accent_color = Some(0x00C43E1C); // Accent bar laranja/carmim
            }
            TipoAppOffice::Access => {
                theme.active_tab_bg = 0x00FBF0F1; // Vinho/rosa suave pastel
                theme.active_tab_text = 0x00802225; // Vinho Access
                theme.hover_tab_bg = 0x00FDF5F6; // Hover vinho sutil
                theme.tab_border = 0x00E8BFC1; // Borda vinho pastel
                theme.custom_accent_color = Some(0x00A4373A); // Accent bar vinho Access
            }
            TipoAppOffice::Project => {
                theme.active_tab_bg = 0x00EFF6EF; // Verde menta suave pastel
                theme.active_tab_text = 0x0020521E; // Verde Project floresta
                theme.hover_tab_bg = 0x00F5FAF5; // Hover verde menta
                theme.tab_border = 0x00BFDDBE; // Borda verde pastel
                theme.custom_accent_color = Some(0x0031752F); // Accent bar verde Project
            }
            TipoAppOffice::Visio => {
                theme.active_tab_bg = 0x00EFF2FA; // Lavanda/índigo suave pastel
                theme.active_tab_text = 0x00243870; // Índigo Visio
                theme.hover_tab_bg = 0x00F5F7FC; // Hover índigo sutil
                theme.tab_border = 0x00BDCCEB; // Borda índigo pastel
                theme.custom_accent_color = Some(0x003955A3); // Accent bar índigo Visio
            }
            TipoAppOffice::Publisher => {
                theme.active_tab_bg = 0x00EBF6F5; // Turquesa/água suave pastel
                theme.active_tab_text = 0x00005C51; // Turquesa Publisher escuro
                theme.hover_tab_bg = 0x00F2FAF9; // Hover turquesa sutil
                theme.tab_border = 0x00A3DAD4; // Borda turquesa pastel
                theme.custom_accent_color = Some(0x00008272); // Accent bar turquesa Publisher
            }
            TipoAppOffice::Desconhecido => {
                theme.active_tab_bg = 0x00EBF1FB;
                theme.active_tab_text = 0x00185ABD;
                theme.hover_tab_bg = 0x00F0F5FD;
                theme.tab_border = 0x00B8D0F0;
                theme.custom_accent_color = Some(0x00185ABD);
            }
        }

        theme
    }

    /// Trata a destruição de uma janela no Modo de Janela Única.
    pub fn processar_destroy(&mut self, hwnd: HWND) {
        unsafe {
            let _ = KillTimer(hwnd, OFFICETAB_TIMER_ID);
        }
        if let Some(mut estado) = self.estados.remove(&(hwnd.0 as isize))
            && let Some(ref mut tab) = estado.tab_control
        {
            tab.destroy();
        }

        // Remover da lista ordenada de documentos
        self.ordem_documentos.retain(|&h| h != hwnd);

        // MODO DE JANELA ÚNICA:
        // Se a janela destruída era a janela visível ativa, restaurar e exibir
        // imediatamente a próxima janela restante da lista para que o Word
        // permaneça visível na tela e na barra de tarefas!
        let proxima_janela = self
            .ordem_documentos
            .first()
            .copied()
            .or_else(|| self.estados.keys().next().map(|&raw| HWND(raw as *mut _)));

        if let Some(hwnd_restante) = proxima_janela {
            unsafe {
                if IsWindow(hwnd_restante).as_bool() {
                    let _ = ShowWindow(hwnd_restante, SW_SHOW);
                    let _ = BringWindowToTop(hwnd_restante);
                    let _ = SetForegroundWindow(hwnd_restante);
                    let _ =
                        PostMessageW(hwnd_restante, WM_OFFICETAB_RECARREGAR, WPARAM(0), LPARAM(0));
                }
            }
        }

        // Notificar as demais janelas restantes para atualizar as abas imediatamente
        for &raw_restante in self.estados.keys() {
            let hwnd_restante = HWND(raw_restante as *mut _);
            if Some(hwnd_restante) != proxima_janela {
                unsafe {
                    if IsWindow(hwnd_restante).as_bool() {
                        let _ = PostMessageW(
                            hwnd_restante,
                            WM_OFFICETAB_RECARREGAR,
                            WPARAM(0),
                            LPARAM(0),
                        );
                    }
                }
            }
        }
    }

    /// Limpa todos os hooks e referências ativas.
    pub fn cleanup(&mut self) {
        for (&raw_hwnd, estado) in self.estados.iter_mut() {
            let hwnd = HWND(raw_hwnd as *mut _);
            unsafe {
                let _ = KillTimer(hwnd, OFFICETAB_TIMER_ID);
            }
            if let Some(ref mut tab) = estado.tab_control {
                tab.destroy();
            }
        }
        self.estados.clear();
        self.ordem_documentos.clear();
        self.subclass_manager.unsubclass_all();
    }

    /// Retorna true se o HWND já está sendo monitorado.
    pub fn esta_monitorado(&self, hwnd: HWND) -> bool {
        self.estados.contains_key(&(hwnd.0 as isize))
    }

    /// Trata o clique do botão esquerdo do mouse sobre o controle de abas.
    /// Retorna true se o clique ocorreu sobre o botão de fechar ('×') de uma aba,
    /// interceptando a mensagem e fechando o documento diretamente.
    pub fn tratar_lbutton_up(&mut self, hwnd_tab: HWND, x: i32, y: i32) -> bool {
        for estado in self.estados.values_mut() {
            if let Some(ref tab) = estado.tab_control
                && tab.hwnd == hwnd_tab
            {
                let total_docs = estado.doc_hwnds.len();
                if let Some(tab_idx) = tab.hit_test(x, y)
                    && tab_idx >= 0
                    && (tab_idx as usize) < total_docs
                {
                    let mut rc_tab = RECT::default();
                    unsafe {
                        let _ = windows::Win32::UI::WindowsAndMessaging::SendMessageW(
                            hwnd_tab,
                            crate::tab_control::TCM_GETITEMRECT,
                            WPARAM(tab_idx as usize),
                            LPARAM(&mut rc_tab as *mut _ as isize),
                        );
                    }

                    let close_size = 20;
                    let close_x = rc_tab.right - close_size - 8;
                    let close_y = rc_tab.top + (rc_tab.bottom - rc_tab.top - close_size) / 2;

                    if x >= close_x - 4
                        && x <= close_x + close_size + 4
                        && y >= close_y - 4
                        && y <= close_y + close_size + 4
                    {
                        let doc_fechar = estado.doc_hwnds[tab_idx as usize];
                        let doc_origem = estado.hwnd;
                        log_debug(&format!(
                            "tratar_lbutton_up: fechando doc via clique no botão '×' da aba {} HWND=0x{:X}",
                            tab_idx, doc_fechar.0 as usize
                        ));
                        if doc_fechar != doc_origem {
                            // Alternância diferida antes de fechar (evita crash no Excel)
                            unsafe {
                                let _ = PostMessageW(
                                    doc_fechar,
                                    WM_OFFICETAB_ALTERNAR,
                                    WPARAM(doc_origem.0 as usize),
                                    LPARAM(0),
                                );
                            }
                        }
                        // WM_CLOSE dispara a caixa de diálogo "Salvar?" nativa do Office
                        unsafe {
                            let _ = PostMessageW(doc_fechar, WM_CLOSE, WPARAM(0), LPARAM(0));
                        }
                        return true;
                    }
                }
                break;
            }
        }
        false
    }

    /// Trata o clique com o botão do meio (mouse middle click) em uma aba para fechá-la.
    pub fn tratar_mbutton_up(&mut self, hwnd_tab: HWND, x: i32, y: i32) {
        for estado in self.estados.values_mut() {
            if let Some(ref tab) = estado.tab_control
                && tab.hwnd == hwnd_tab
            {
                if let Some(tab_idx) = tab.hit_test(x, y) {
                    let total_docs = estado.doc_hwnds.len();
                    if tab_idx >= 0 && (tab_idx as usize) < total_docs {
                        let doc_fechar = estado.doc_hwnds[tab_idx as usize];
                        let doc_origem = estado.hwnd;
                        log_debug(&format!(
                            "tratar_mbutton_up: fechando doc na aba {} HWND=0x{:X}",
                            tab_idx, doc_fechar.0 as usize
                        ));
                        if doc_fechar != doc_origem {
                            unsafe {
                                let _ = PostMessageW(
                                    doc_fechar,
                                    WM_OFFICETAB_ALTERNAR,
                                    WPARAM(doc_origem.0 as usize),
                                    LPARAM(0),
                                );
                            }
                        }
                        unsafe {
                            let _ = PostMessageW(doc_fechar, WM_CLOSE, WPARAM(0), LPARAM(0));
                        }
                    }
                }
                break;
            }
        }
    }

    /// Trata o clique duplo do mouse (Double Click) sobre o controle de abas.
    /// Clique duplo na aba fecha o documento.
    /// Clique duplo na barra vazia ou no botão '+' cria novo documento.
    pub fn tratar_lbutton_dblclk(&mut self, hwnd_tab: HWND, x: i32, y: i32) {
        for estado in self.estados.values_mut() {
            if let Some(ref tab) = estado.tab_control
                && tab.hwnd == hwnd_tab
            {
                let total_docs = estado.doc_hwnds.len();
                if let Some(tab_idx) = tab.hit_test(x, y) {
                    if tab_idx >= 0 && (tab_idx as usize) < total_docs {
                        let doc_fechar = estado.doc_hwnds[tab_idx as usize];
                        let doc_origem = estado.hwnd;
                        log_debug(&format!(
                            "tratar_lbutton_dblclk: fechando doc via clique duplo na aba {} HWND=0x{:X}",
                            tab_idx, doc_fechar.0 as usize
                        ));
                        if doc_fechar != doc_origem {
                            unsafe {
                                let _ = PostMessageW(
                                    doc_fechar,
                                    WM_OFFICETAB_ALTERNAR,
                                    WPARAM(doc_origem.0 as usize),
                                    LPARAM(0),
                                );
                            }
                        }
                        unsafe {
                            let _ = PostMessageW(doc_fechar, WM_CLOSE, WPARAM(0), LPARAM(0));
                        }
                    } else if (tab_idx as usize) == total_docs {
                        log_debug("tratar_lbutton_dblclk: novo doc via clique duplo no '+'");
                        unsafe {
                            let _ = PostMessageW(
                                estado.hwnd,
                                WM_OFFICETAB_NOVO_DOC,
                                WPARAM(0),
                                LPARAM(0),
                            );
                        }
                    }
                } else {
                    log_debug("tratar_lbutton_dblclk: novo doc via clique duplo na barra vazia");
                    unsafe {
                        let _ =
                            PostMessageW(estado.hwnd, WM_OFFICETAB_NOVO_DOC, WPARAM(0), LPARAM(0));
                    }
                }
                break;
            }
        }
    }

    /// Trata a rolagem da roda do mouse (Mouse Wheel) sobre a barra de abas, alternando para aba anterior ou seguinte.
    pub fn tratar_mousewheel(&mut self, hwnd_tab: HWND, delta: i16) {
        let mut transicao: Option<(HWND, HWND)> = None;
        for estado in self.estados.values() {
            if let Some(ref tab) = estado.tab_control
                && tab.hwnd == hwnd_tab
            {
                let total = estado.doc_hwnds.len();
                if total > 1 {
                    let cur = tab.get_selected_tab();
                    let proximo = if delta > 0 {
                        // Scroll para cima/frente: aba anterior
                        if cur <= 0 {
                            (total - 1) as i32
                        } else {
                            cur - 1
                        }
                    } else {
                        // Scroll para baixo/trás: próxima aba
                        if cur >= (total - 1) as i32 {
                            0
                        } else {
                            cur + 1
                        }
                    };
                    if (proximo as usize) < total {
                        let alvo = estado.doc_hwnds[proximo as usize];
                        if alvo != estado.hwnd {
                            transicao = Some((estado.hwnd, alvo));
                        }
                    }
                }
                break;
            }
        }

        if let Some((origem, alvo)) = transicao {
            // Alternância DIFERIDA via mensagem postada (mesma proteção do clique em aba)
            unsafe {
                let _ = PostMessageW(
                    alvo,
                    WM_OFFICETAB_ALTERNAR,
                    WPARAM(origem.0 as usize),
                    LPARAM(0),
                );
            }
        }
    }

    /// Trata atalhos de teclado do OfficeTab:
    /// - Ctrl+Tab: Próxima aba
    /// - Ctrl+Shift+Tab: Aba anterior
    /// - Ctrl+W / Ctrl+F4: Fechar documento atual
    /// - Ctrl+T: Nova aba / documento
    pub fn tratar_keydown(&mut self, hwnd: HWND, vk: usize) -> bool {
        let ctrl_pressionado = unsafe {
            (windows::Win32::UI::Input::KeyboardAndMouse::GetKeyState(VK_CONTROL.0 as i32) as u16
                & 0x8000)
                != 0
        };
        if !ctrl_pressionado {
            return false;
        }
        let shift_pressionado = unsafe {
            (windows::Win32::UI::Input::KeyboardAndMouse::GetKeyState(
                windows::Win32::UI::Input::KeyboardAndMouse::VK_SHIFT.0 as i32,
            ) as u16
                & 0x8000)
                != 0
        };

        let (total, cur, doc_hwnds) = match self.estados.get(&(hwnd.0 as isize)) {
            Some(e) => (
                e.doc_hwnds.len(),
                e.tab_control
                    .as_ref()
                    .map(|t| t.get_selected_tab())
                    .unwrap_or(0),
                e.doc_hwnds.clone(),
            ),
            None => return false,
        };

        if total == 0 {
            return false;
        }

        // VK_TAB = 0x09
        if vk == 0x09 {
            let proximo = if shift_pressionado {
                if cur <= 0 {
                    (total - 1) as i32
                } else {
                    cur - 1
                }
            } else {
                if cur >= (total - 1) as i32 {
                    0
                } else {
                    cur + 1
                }
            };
            if (proximo as usize) < total {
                let alvo = doc_hwnds[proximo as usize];
                if alvo != hwnd {
                    // Alternância DIFERIDA via mensagem postada (mesma proteção do clique)
                    unsafe {
                        let _ = PostMessageW(
                            alvo,
                            WM_OFFICETAB_ALTERNAR,
                            WPARAM(hwnd.0 as usize),
                            LPARAM(0),
                        );
                    }
                }
                return true;
            }
        } else if vk == b'W' as usize || vk == 0x73
        /* VK_F4 */
        {
            // Ctrl+W ou Ctrl+F4: fechar documento atual
            unsafe {
                let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
            }
            return true;
        } else if vk == b'T' as usize {
            // No Excel, Ctrl+T é o atalho nativo universal para "Criar Tabela" (Create Table).
            // Para não quebrar o fluxo de trabalho do usuário no Excel, não interceptamos Ctrl+T no Excel,
            // permitindo que o usuário use o botão '+' da barra ou Ctrl+O/Ctrl+N nativos.
            let classe = extrair_classe_janela(hwnd);
            if obter_tipo_app(&classe) != TipoAppOffice::Excel {
                // Ctrl+T: nova aba / novo documento nos demais aplicativos Office
                unsafe {
                    let _ = PostMessageW(hwnd, WM_OFFICETAB_NOVO_DOC, WPARAM(0), LPARAM(0));
                }
                return true;
            }
        }

        false
    }

    /// Trata o clique com o botão direito do mouse exibindo menu de contexto.
    pub fn tratar_rbutton_up(&mut self, hwnd_tab: HWND, x: i32, y: i32) {
        let mut pt_screen = POINT { x, y };
        unsafe {
            let _ = ClientToScreen(hwnd_tab, &mut pt_screen);
        }
        self.abrir_menu_contexto(hwnd_tab, pt_screen, Some(POINT { x, y }));
    }

    /// Trata WM_CONTEXTMENU disparado via coordenadas de tela.
    pub fn tratar_context_menu(&mut self, hwnd_tab: HWND, screen_x: i32, screen_y: i32) {
        let pt_screen = POINT {
            x: screen_x,
            y: screen_y,
        };
        let mut pt_client = pt_screen;
        unsafe {
            let _ = ScreenToClient(hwnd_tab, &mut pt_client);
        }
        self.abrir_menu_contexto(hwnd_tab, pt_screen, Some(pt_client));
    }

    /// Trata menu de contexto acionado via teclado (tecla Menu ou Shift+F10).
    pub fn tratar_context_menu_teclado(&mut self, hwnd_tab: HWND) {
        let mut pt_screen = POINT::default();
        unsafe {
            let _ = GetCursorPos(&mut pt_screen);
        }
        self.abrir_menu_contexto(hwnd_tab, pt_screen, None);
    }

    /// Constrói e exibe o menu de contexto moderno para a aba ou barra de abas.
    fn abrir_menu_contexto(&mut self, hwnd_tab: HWND, pt_screen: POINT, pt_client: Option<POINT>) {
        for estado in self.estados.values_mut() {
            if let Some(ref tab) = estado.tab_control
                && tab.hwnd == hwnd_tab
            {
                let total_docs = estado.doc_hwnds.len();
                let tab_idx = match pt_client {
                    Some(pt) => tab.hit_test(pt.x, pt.y),
                    None => {
                        let sel = tab.get_selected_tab();
                        if sel >= 0 { Some(sel) } else { None }
                    }
                };

                unsafe {
                    let hmenu = match CreatePopupMenu() {
                        Ok(m) => m,
                        Err(_) => return,
                    };

                    if let Some(idx) = tab_idx
                        && (idx as usize) < total_docs
                    {
                        // Clicou sobre uma aba de documento válida
                        let _ =
                            AppendMenuW(hmenu, MF_STRING, IDM_FECHAR, w!("&Fechar Aba\tCtrl+W"));
                        let _ = AppendMenuW(
                            hmenu,
                            MF_STRING,
                            IDM_FECHAR_OUTRAS,
                            w!("Fechar &Outras Abas"),
                        );
                        if (idx as usize) + 1 < total_docs {
                            let _ = AppendMenuW(
                                hmenu,
                                MF_STRING,
                                IDM_FECHAR_DIREITA,
                                w!("Fechar Abas à &Direita"),
                            );
                        }
                        let _ = AppendMenuW(hmenu, MF_SEPARATOR, 0, PCWSTR::null());
                        let _ = AppendMenuW(
                            hmenu,
                            MF_STRING,
                            IDM_COPIAR_CAMINHO,
                            w!("&Copiar Caminho do Arquivo"),
                        );
                        let _ = AppendMenuW(
                            hmenu,
                            MF_STRING,
                            IDM_ABRIR_PASTA,
                            w!("&Abrir Pasta do Arquivo"),
                        );
                        let _ = AppendMenuW(hmenu, MF_SEPARATOR, 0, PCWSTR::null());
                        let _ =
                            AppendMenuW(hmenu, MF_STRING, IDM_NOVA_ABA, w!("&Nova Aba\tCtrl+N"));
                    } else {
                        // Clicou na área vazia da barra ou na aba "+"
                        let _ =
                            AppendMenuW(hmenu, MF_STRING, IDM_NOVA_ABA, w!("&Nova Aba\tCtrl+N"));
                        if total_docs > 0 {
                            let _ = AppendMenuW(hmenu, MF_SEPARATOR, 0, PCWSTR::null());
                            let _ = AppendMenuW(
                                hmenu,
                                MF_STRING,
                                IDM_FECHAR_TODAS,
                                w!("Fechar &Todas as Abas"),
                            );
                        }
                    }

                    let cmd = TrackPopupMenu(
                        hmenu,
                        TPM_RETURNCMD | TPM_RIGHTBUTTON,
                        pt_screen.x,
                        pt_screen.y,
                        0,
                        hwnd_tab,
                        None,
                    );
                    let _ = DestroyMenu(hmenu);

                    let cmd_val = cmd.0 as usize;
                    match cmd_val {
                        IDM_FECHAR => {
                            if let Some(idx) = tab_idx
                                && (idx as usize) < total_docs
                            {
                                let doc = estado.doc_hwnds[idx as usize];
                                let doc_origem = estado.hwnd;
                                if doc != doc_origem {
                                    self.alternar_para_janela(doc_origem, doc);
                                    let _ = self.sincronizar_janela(doc);
                                }
                                let _ = PostMessageW(doc, WM_CLOSE, WPARAM(0), LPARAM(0));
                            }
                        }
                        IDM_FECHAR_OUTRAS => {
                            if let Some(idx) = tab_idx
                                && (idx as usize) < total_docs
                            {
                                let manter = estado.doc_hwnds[idx as usize];
                                for &doc in &estado.doc_hwnds {
                                    if doc != manter {
                                        let _ = PostMessageW(doc, WM_CLOSE, WPARAM(0), LPARAM(0));
                                    }
                                }
                            }
                        }
                        IDM_FECHAR_DIREITA => {
                            if let Some(idx) = tab_idx
                                && (idx as usize) < total_docs
                            {
                                for &doc in estado.doc_hwnds.iter().skip((idx as usize) + 1) {
                                    let _ = PostMessageW(doc, WM_CLOSE, WPARAM(0), LPARAM(0));
                                }
                            }
                        }
                        IDM_COPIAR_CAMINHO => {
                            if let Some(idx) = tab_idx
                                && (idx as usize) < estado.titulos_cache.len()
                            {
                                let titulo = &estado.titulos_cache[idx as usize];
                                let texto_copiar = if let Some(caminho) =
                                    crate::utils::localizar_caminho_documento(titulo)
                                {
                                    caminho.display().to_string()
                                } else {
                                    titulo.clone()
                                };
                                let _ = crate::utils::copiar_para_clipboard(&texto_copiar);
                                log_debug(&format!(
                                    "Caminho/título copiado para clipboard: {}",
                                    texto_copiar
                                ));
                            }
                        }
                        IDM_ABRIR_PASTA => {
                            if let Some(idx) = tab_idx
                                && (idx as usize) < estado.titulos_cache.len()
                            {
                                let titulo = &estado.titulos_cache[idx as usize];
                                if let Some(caminho) =
                                    crate::utils::localizar_caminho_documento(titulo)
                                {
                                    let _ = crate::utils::abrir_pasta_no_explorer(&caminho);
                                } else {
                                    let doc_dir = crate::utils::obter_pasta_documentos_usuario();
                                    let _ = crate::utils::abrir_pasta_no_explorer(&doc_dir);
                                }
                            }
                        }
                        IDM_NOVA_ABA => {
                            let _ = PostMessageW(
                                estado.hwnd,
                                WM_OFFICETAB_NOVO_DOC,
                                WPARAM(0),
                                LPARAM(0),
                            );
                        }
                        IDM_FECHAR_TODAS => {
                            for &doc in &estado.doc_hwnds {
                                let _ = PostMessageW(doc, WM_CLOSE, WPARAM(0), LPARAM(0));
                            }
                        }
                        _ => {}
                    }
                }
                break;
            }
        }
    }
}

impl Drop for TabManager {
    fn drop(&mut self) {
        // NÃO chama cleanup() aqui!
        // A destruição de abas e remoção de hooks só deve ocorrer quando desativar_gerenciador_global() for explicitamente invocado.
    }
}

impl Default for TabManager {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// Singleton Global, Timer UI e WinEventHook
// ---------------------------------------------------------------------------

static GLOBAL_TAB_MANAGER: Mutex<Option<TabManager>> = Mutex::new(None);
static MONITORAMENTO_ATIVO: AtomicBool = AtomicBool::new(false);
static WINEVENT_HOOK: Mutex<Option<isize>> = Mutex::new(None);
static TIMER_UI_ID: Mutex<Option<usize>> = Mutex::new(None);

/// Informa se o gerenciador global de abas pode ser descarregado da memória sem risco de crash.
/// Utilizado diretamente por DllCanUnloadNow para garantir cleanup limpo e encerramento das threads.
pub fn pode_descarregar() -> bool {
    if MONITORAMENTO_ATIVO.load(Ordering::Acquire) {
        return false;
    }
    if let Ok(guard) = GLOBAL_TAB_MANAGER.lock()
        && let Some(ref manager) = *guard
    {
        return manager.estados.is_empty();
    }
    true
}

/// Verifica janelas periodicamente no loop da thread Office que inicializou o add-in.
/// O hook é o caminho principal; este timer cobre hooks indisponíveis e eventos perdidos.
unsafe extern "system" fn timer_ui_callback(_hwnd: HWND, _msg: u32, _id: usize, _time: u32) {
    if !MONITORAMENTO_ATIVO.load(Ordering::Acquire) {
        return;
    }
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let janelas = listar_janelas_processo();
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            for (hwnd, _) in janelas {
                if !manager.esta_monitorado(hwnd) {
                    let _ = manager.monitor_window(hwnd);
                }
            }
        }
    }));
}

/// Callback do WinEventHook — executado na thread de UI do processo que gerou o evento.
/// Blindado com catch_unwind para proteger a integridade do processo do Office.
unsafe extern "system" fn winevent_callback(
    _hook: HWINEVENTHOOK,
    _event: u32,
    hwnd: HWND,
    id_object: i32,
    _id_child: i32,
    _id_event_thread: u32,
    _dwms_event_time: u32,
) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if id_object != 0 || hwnd.is_invalid() {
            return;
        }

        unsafe {
            if !IsWindowVisible(hwnd).as_bool() {
                return;
            }
        }

        let classe = extrair_classe_janela(hwnd);
        let tipo_app = obter_tipo_app(&classe);
        if !tipo_app.e_office() {
            return;
        }

        // Posta WM_OFFICETAB_INICIALIZAR de forma assíncrona para que a janela seja anexada
        // no loop de mensagens normal da UI, garantindo ZERO risco de deadlock de reentrância!
        unsafe {
            let _ = PostMessageW(hwnd, WM_OFFICETAB_INICIALIZAR, WPARAM(0), LPARAM(0));
        }
    }));
}

/// Instala WinEventHook para detectar novas janelas Office no processo atual.
fn instalar_winevent_hook() -> bool {
    let mut guard = WINEVENT_HOOK.lock().unwrap();
    if guard.is_some() {
        return true;
    }

    unsafe {
        // Intercepta eventos desde a criação (EVENT_OBJECT_CREATE = 0x8000) até a exibição (EVENT_OBJECT_SHOW = 0x8002)
        // para capturar instantaneamente novas janelas e pastas de trabalho do Office (Word, Excel, etc.).
        let hook = SetWinEventHook(
            EVENT_OBJECT_CREATE,
            EVENT_OBJECT_SHOW,
            windows::Win32::Foundation::HMODULE::default(),
            Some(winevent_callback),
            GetCurrentProcessId(),
            0,
            WINEVENT_OUTOFCONTEXT,
        );

        if !hook.is_invalid() {
            log_debug(&format!(
                "WinEventHook instalado: handle=0x{:x}",
                hook.0 as usize
            ));
            *guard = Some(hook.0 as usize as isize);
            true
        } else {
            log_debug("Falha ao instalar WinEventHook");
            false
        }
    }
}

fn desinstalar_winevent_hook() {
    let mut guard = WINEVENT_HOOK.lock().unwrap();
    if let Some(raw) = guard.take() {
        unsafe {
            let _ = UnhookWinEvent(HWINEVENTHOOK(raw as *mut _));
            log_debug("WinEventHook removido");
        }
    }
}

/// Anexa janelas já abertas diretamente na thread UI que recebeu OnConnection.
fn varrer_janelas_existentes() {
    let janelas = listar_janelas_processo();
    if janelas.is_empty() {
        return;
    }

    log_debug(&format!(
        "Varrendo {} janelas Office já abertas...",
        janelas.len()
    ));

    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock()
        && let Some(ref mut manager) = *guard
    {
        for (hwnd, _) in janelas {
            if !manager.esta_monitorado(hwnd)
                && let Err(error) = manager.monitor_window(hwnd)
            {
                log_debug(&format!(
                    "Falha ao anexar janela inicial {hwnd:?}: {error:?}"
                ));
            }
        }
    }
}

/// Inicia o monitoramento completo do Office Tab
pub fn iniciar_thread_monitoramento() {
    if MONITORAMENTO_ATIVO.swap(true, Ordering::AcqRel) {
        return;
    }
    let _ = instalar_winevent_hook();

    // Mantém o timer de verificação periódico (1000ms) na thread de UI ativa como rede de proteção,
    // garantindo que qualquer janela secundária de pasta de trabalho criada assincronamente pelo Excel
    // seja monitorada e receba a barra de abas mesmo antes de qualquer interação do usuário.
    unsafe {
        let mut guard = TIMER_UI_ID.lock().unwrap();
        if guard.is_none() {
            let tid = SetTimer(HWND::default(), 0, 1000, Some(timer_ui_callback));
            if tid != 0 {
                *guard = Some(tid);
            }
        }
    }

    varrer_janelas_existentes();
}

pub fn obter_ou_criar_gerenciador() {
    let mut guard = GLOBAL_TAB_MANAGER.lock().unwrap();
    if guard.is_none() {
        *guard = Some(TabManager::new());
    }
    drop(guard);
    iniciar_thread_monitoramento();
}

// ---------------------------------------------------------------------------
// Despacho de eventos de janela (thread de UI do Office)
// ---------------------------------------------------------------------------

/// Chamado via WM_TIMER. Blindado com catch_unwind.
pub fn despachar_timer(hwnd: HWND) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            let _ = manager.sincronizar_janela(hwnd);
        }
    }));
}

/// Chamado via WM_SIZE. Blindado com catch_unwind.
pub fn despachar_size(hwnd: HWND) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            let _ = manager.sincronizar_janela(hwnd);
        }
    }));
}

/// Chamado via WM_NOTIFY. Blindado com catch_unwind.
pub fn despachar_notify(hwnd: HWND, lparam: LPARAM) -> Option<LRESULT> {
    std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            return manager.processar_notify(hwnd, lparam);
        }
        None
    }))
    .unwrap_or(None)
}

/// Chamado pelo WM_PAINT do próprio TabControl (SysTabControl32).
/// No Excel (SDI), o WM_NOTIFY/NM_CUSTOMDRAW que o controle envia ao pai (XLMAIN)
/// é engolido pelo Excel antes da nossa subclasse vê-lo, então o desenho
/// customizado (×, título à esquerda, cores) nunca rodava. Aqui pintamos
/// manualmente reaproveitando processar_custom_draw com um NMCUSTOMDRAW sintetizado.
pub fn despachar_tab_paint(hwnd_tab: HWND, hdc: HDC, rc_paint: RECT) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            manager.pintar_tab_control(hwnd_tab, hdc, rc_paint);
        }
    }));
}

/// Chamado via WM_DESTROY/WM_NCDESTROY. Blindado com catch_unwind.
pub fn despachar_destroy(hwnd: HWND) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock()
            && let Some(ref mut manager) = *guard
        {
            manager.processar_destroy(hwnd);
        }
    }));
}

/// Chamado via WM_OFFICETAB_INICIALIZAR. Blindado com catch_unwind.
pub fn despachar_inicializar(hwnd: HWND) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        log_debug(&format!(
            "despachar_inicializar (thread UI): HWND=0x{:X}",
            hwnd.0 as usize
        ));

        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
            if guard.is_none() {
                *guard = Some(TabManager::new());
            }
            if let Some(ref mut manager) = *guard
                && !manager.esta_monitorado(hwnd)
            {
                // MODO DE JANELA ÚNICA: Sincroniza a janela ativa com a última conhecida,
                // mas NUNCA oculta a janela principal do aplicativo (que contém o Ribbon/Menus).
                for &raw_existente in manager.estados.keys() {
                    if raw_existente != (hwnd.0 as isize) {
                        let hwnd_existente = HWND(raw_existente as *mut _);
                        unsafe {
                            if IsWindow(hwnd_existente).as_bool() {
                                // Se a janela existente for a "mestra" (tem ribbon/menus), usamos suas dimensões
                                // para a nova janela, mas não a ocultamos.
                                if IsZoomed(hwnd_existente).as_bool() {
                                    let _ = ShowWindow(hwnd, SW_MAXIMIZE);
                                } else {
                                    let mut rc = RECT::default();
                                    if GetWindowRect(hwnd_existente, &mut rc).is_ok() {
                                        let _ = SetWindowPos(
                                            hwnd,
                                            HWND::default(),
                                            rc.left,
                                            rc.top,
                                            rc.right - rc.left,
                                            rc.bottom - rc.top,
                                            SWP_NOZORDER,
                                        );
                                    }
                                }
                                // CORREÇÃO: Não ocultar janelas que podem ser a principal do Office.
                                // Apenas janelas de documento "filhas" devem ser ocultadas se houver conflito.
                                // Por segurança, mantemos a janela anterior visível se ela for a única outra.
                            }
                        }
                    }
                }

                unsafe {
                    let _ = ShowWindow(hwnd, SW_SHOW);
                    let _ = BringWindowToTop(hwnd);
                    let _ = SetForegroundWindow(hwnd);
                }

                let res = manager.monitor_window(hwnd);
                log_debug(&format!(
                    "despachar_inicializar: monitor_window res={}",
                    res.is_ok()
                ));

                for &raw_existente in manager.estados.keys() {
                    let hwnd_existente = HWND(raw_existente as *mut _);
                    if hwnd_existente != hwnd {
                        unsafe {
                            let _ = PostMessageW(
                                hwnd_existente,
                                WM_OFFICETAB_RECARREGAR,
                                WPARAM(0),
                                LPARAM(0),
                            );
                        }
                    }
                }
            }
        }
    }));
}

/// Chamado via WM_OFFICETAB_RECARREGAR. Blindado com catch_unwind.
pub fn despachar_recarregar(hwnd: HWND) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            if let Some(estado) = manager.estados.get_mut(&(hwnd.0 as isize)) {
                estado.titulos_cache.clear();
            }
            let _ = manager.sincronizar_janela(hwnd);
        }
    }));
}

/// Chamado via WM_OFFICETAB_ALTERNAR na janela de DESTINO. Blindado com catch_unwind.
/// Executa a alternância do Modo Janela Única FORA do WM_NOTIFY de origem:
/// wParam carrega o HWND bruto da janela de origem que deve ser ocultada.
pub fn despachar_alternar(hwnd_alvo: HWND, hwnd_origem: HWND) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        log_debug(&format!(
            "despachar_alternar: origem=0x{:X} → destino=0x{:X}",
            hwnd_origem.0 as usize, hwnd_alvo.0 as usize
        ));
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            // Garantir que a janela de destino está monitorada antes de alternar
            let _ = manager.monitor_window(hwnd_alvo);
            manager.alternar_para_janela(hwnd_origem, hwnd_alvo);
            let _ = manager.sincronizar_janela(hwnd_alvo);
        }
    }));
}

/// Chamado quando o botão esquerdo do mouse é solto sobre o controle de abas (SysTabControl32). Blindado com catch_unwind.
pub fn despachar_tab_lbutton_up(hwnd_tab: HWND, x: i32, y: i32) -> bool {
    std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            return manager.tratar_lbutton_up(hwnd_tab, x, y);
        }
        false
    }))
    .unwrap_or(false)
}

/// Chamado quando o botão do meio do mouse é clicado sobre o controle de abas (SysTabControl32). Blindado com catch_unwind.
pub fn despachar_tab_mbutton_up(hwnd_tab: HWND, x: i32, y: i32) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            manager.tratar_mbutton_up(hwnd_tab, x, y);
        }
    }));
}

/// Chamado quando o botão direito do mouse é clicado sobre o controle de abas (SysTabControl32). Blindado com catch_unwind.
pub fn despachar_tab_rbutton_up(hwnd_tab: HWND, x: i32, y: i32) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            manager.tratar_rbutton_up(hwnd_tab, x, y);
        }
    }));
}

/// Chamado quando o menu de contexto é acionado sobre o controle de abas via coordenadas de tela. Blindado com catch_unwind.
pub fn despachar_tab_context_menu(hwnd_tab: HWND, screen_x: i32, screen_y: i32) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            manager.tratar_context_menu(hwnd_tab, screen_x, screen_y);
        }
    }));
}

/// Chamado quando o menu de contexto é acionado sobre o controle de abas via atalho de teclado. Blindado com catch_unwind.
pub fn despachar_tab_context_menu_teclado(hwnd_tab: HWND) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            manager.tratar_context_menu_teclado(hwnd_tab);
        }
    }));
}

/// Chamado quando ocorre duplo clique sobre o controle de abas (SysTabControl32). Blindado com catch_unwind.
pub fn despachar_tab_lbutton_dblclk(hwnd_tab: HWND, x: i32, y: i32) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            manager.tratar_lbutton_dblclk(hwnd_tab, x, y);
        }
    }));
}

/// Chamado quando a roda do mouse é girada sobre o controle de abas. Blindado com catch_unwind.
pub fn despachar_tab_mousewheel(hwnd_tab: HWND, delta: i16) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            manager.tratar_mousewheel(hwnd_tab, delta);
        }
    }));
}

/// Chamado em WM_KEYDOWN na janela monitorada do Office para interceptar atalhos (Ctrl+Tab, Ctrl+W, Ctrl+T).
pub fn despachar_keydown(hwnd: HWND, vk: usize) -> bool {
    std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if let Ok(mut guard) = GLOBAL_TAB_MANAGER.try_lock()
            && let Some(ref mut manager) = *guard
        {
            return manager.tratar_keydown(hwnd, vk);
        }
        false
    }))
    .unwrap_or(false)
}

/// Varre e anexa o gerenciamento às janelas ativas do Office.
pub fn varrer_e_anexar_janelas() {
    let janelas = listar_janelas_processo();
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if guard.is_none() {
            *guard = Some(TabManager::new());
        }
        if let Some(ref mut manager) = *guard {
            for (hwnd, _) in janelas {
                let _ = manager.monitor_window(hwnd);
            }
        }
    }
}

/// Desativa completamente o gerenciamento global de abas, desconecta hooks e encerra threads.
pub fn desativar_gerenciador_global() {
    MONITORAMENTO_ATIVO.store(false, Ordering::SeqCst);
    desinstalar_winevent_hook();

    unsafe {
        let mut guard = TIMER_UI_ID.lock().unwrap();
        if let Some(tid) = guard.take() {
            let _ = KillTimer(HWND::default(), tid);
            log_debug("Timer de UI cancelado");
        }
    }

    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if let Some(ref mut manager) = *guard {
            manager.cleanup();
        }
        *guard = None;
    }
}

// ---------------------------------------------------------------------------
// Funções Auxiliares Win32
// ---------------------------------------------------------------------------

/// Chamado via WM_OFFICETAB_NOVO_DOC. Blindado com catch_unwind.
pub fn despachar_novo_doc(hwnd: HWND) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        criar_novo_documento(hwnd);
    }));
}

fn criar_novo_documento(hwnd: HWND) {
    log_debug(&format!(
        "criar_novo_documento chamado para HWND=0x{:X}",
        hwnd.0 as usize
    ));

    // 1. Tentar criar via COM IDispatch (método oficial e confiável em qualquer idioma do Office)
    if crate::com_addin_simple::criar_documento_com() {
        log_debug("Documento criado nativamente via COM IDispatch com sucesso");
        return;
    }

    // 2. Fallback via teclado/mensagens Win32 adaptado para cada aplicativo Office
    let classe = extrair_classe_janela(hwnd);
    let tipo_app = obter_tipo_app(&classe);
    log_debug(&format!(
        "COM indisponível; tentando atalho de teclado para {:?}",
        tipo_app
    ));

    unsafe {
        if !hwnd.is_invalid() {
            let _ = SetForegroundWindow(hwnd);
            let _ = BringWindowToTop(hwnd);
            let area_doc = localizar_area_documento(hwnd);
            let foco_alvo = area_doc.unwrap_or(hwnd);
            let _ = SetFocus(foco_alvo);
        }

        match tipo_app {
            TipoAppOffice::Word => {
                // No Word PT-BR: Ctrl+O cria novo documento (Ctrl+N é negrito)
                keybd_event(VK_CONTROL.0 as u8, 0, Default::default(), 0);
                keybd_event(b'O', 0, Default::default(), 0);
                keybd_event(b'O', 0, KEYEVENTF_KEYUP, 0);
                keybd_event(VK_CONTROL.0 as u8, 0, KEYEVENTF_KEYUP, 0);
            }
            _ => {
                // No Excel, PowerPoint, Access, Project, Visio e Publisher:
                // Ctrl+N cria novo arquivo/pasta/apresentação/banco/projeto/diagrama
                keybd_event(VK_CONTROL.0 as u8, 0, Default::default(), 0);
                keybd_event(b'N', 0, Default::default(), 0);
                keybd_event(b'N', 0, KEYEVENTF_KEYUP, 0);
                keybd_event(VK_CONTROL.0 as u8, 0, KEYEVENTF_KEYUP, 0);
            }
        }
    }
}

/// Extrai o nome da classe de uma janela Win32.
pub fn extrair_classe_janela(hwnd: HWND) -> String {
    let mut buffer = [0u16; 256];
    let len = unsafe { GetClassNameW(hwnd, &mut buffer) };
    if len > 0 {
        OsString::from_wide(&buffer[..len as usize])
            .to_string_lossy()
            .to_string()
    } else {
        String::new()
    }
}

/// Extrai o título textual de uma janela Win32.
pub fn extrair_texto_janela(hwnd: HWND) -> String {
    let mut buffer = [0u16; 512];
    let len = unsafe { GetWindowTextW(hwnd, &mut buffer) };
    if len > 0 {
        OsString::from_wide(&buffer[..len as usize])
            .to_string_lossy()
            .to_string()
    } else {
        String::new()
    }
}

/// Limpa sufixos padrão do Office do título da janela para os 7 aplicativos suportados.
pub fn limpar_titulo_documento(titulo: &str) -> String {
    let mut limpo = titulo.trim();
    const SUFIXOS: &[&str] = &[
        " - Word",
        " - Excel",
        " - PowerPoint",
        " - Access",
        " - Project",
        " - Visio",
        " - Publisher",
        " [Modo de Compatibilidade]",
        " (Modo de Compatibilidade)",
        " [Compatibility Mode]",
        " (Compatibility Mode)",
        " [Somente Leitura]",
        " (Somente Leitura)",
        " [Read-Only]",
        " (Read-Only)",
        " [Modo de Exibição Protegida]",
        " (Modo de Exibição Protegida)",
        " [Protected View]",
        " (Protected View)",
        " - Salvo",
        " - Salvando...",
        " - Gravado",
        " - Gravando...",
        " - Saved",
        " - Saving...",
        " - Não Salvo",
        " - Não Gravado",
        " - Unsaved",
        " - Microsoft Word",
        " - Microsoft Excel",
        " - Microsoft PowerPoint",
        " - Microsoft Access",
        " - Microsoft Project",
        " - Microsoft Visio",
        " - Microsoft Publisher",
    ];

    let mut alterado = true;
    while alterado {
        alterado = false;
        for sufixo in SUFIXOS {
            if limpo.ends_with(sufixo) {
                limpo = limpo[..limpo.len() - sufixo.len()].trim();
                alterado = true;
            }
        }
    }

    if limpo.is_empty() {
        "Documento".to_string()
    } else {
        limpo.to_string()
    }
}

/// Lista todas as janelas Office do processo atual para os 7 aplicativos suportados (incluindo janelas ocultas no modo de janela única).
pub fn listar_janelas_processo() -> Vec<(HWND, String)> {
    let mut janelas: Vec<(HWND, String)> = Vec::new();
    unsafe {
        let _ = EnumWindows(
            Some(enum_proc_processo),
            LPARAM(&mut janelas as *mut _ as isize),
        );
    }
    janelas
}

unsafe extern "system" fn enum_proc_processo(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| unsafe {
        if !IsWindow(hwnd).as_bool() {
            return;
        }

        let mut pid = 0u32;
        GetWindowThreadProcessId(hwnd, Some(&mut pid));
        let current_pid = GetCurrentProcessId();

        if pid == current_pid {
            let class_name = extrair_classe_janela(hwnd);
            let tipo_app = obter_tipo_app(&class_name);

            // Filtrar apenas janelas do mesmo aplicativo do Office
            if tipo_app.e_office() {
                // Descartar janelas secundárias/popups (diálogos modais, menus suspensos)
                let owner = windows::Win32::UI::WindowsAndMessaging::GetWindow(
                    hwnd,
                    windows::Win32::UI::WindowsAndMessaging::GW_OWNER,
                );
                if let Ok(owner_hwnd) = owner
                    && !owner_hwnd.is_invalid()
                    && owner_hwnd.0 as usize != 0
                {
                    return;
                }

                let titulo = extrair_texto_janela(hwnd);
                // No Office, janelas principais de documentos possuem título não-vazio
                if titulo.trim().is_empty() {
                    return;
                }

                let mut rc = RECT::default();
                let _ = GetWindowRect(hwnd, &mut rc);
                let w = rc.right - rc.left;
                let h = rc.bottom - rc.top;

                let style = GetWindowLongPtrW(hwnd, GWL_STYLE) as u32;
                let is_document_style = (style
                    & (windows::Win32::UI::WindowsAndMessaging::WS_CAPTION.0
                        | windows::Win32::UI::WindowsAndMessaging::WS_THICKFRAME.0))
                    != 0;

                // Filtrar apenas janelas reais de documentos (com tamanho visível, minimizadas ou ocultadas pelo Modo Janela Única)
                if ((w > 100 && h > 100)
                    || IsIconic(hwnd).as_bool()
                    || !IsWindowVisible(hwnd).as_bool())
                    && is_document_style
                {
                    let titulo_limpo = limpar_titulo_documento(&titulo);
                    let janelas = &mut *(lparam.0 as *mut Vec<(HWND, String)>);
                    janelas.push((hwnd, titulo_limpo));
                }
            }
        }
    }));
    BOOL(1)
}

fn listar_todas_janelas_office() -> Vec<JanelaDocumento> {
    let mut janelas: Vec<JanelaDocumento> = Vec::new();
    unsafe {
        let _ = EnumWindows(
            Some(enum_windows_proc),
            LPARAM(&mut janelas as *mut _ as isize),
        );
    }
    janelas
}

unsafe extern "system" fn enum_windows_proc(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| unsafe {
        if !IsWindowVisible(hwnd).as_bool() {
            return;
        }

        let class_name = extrair_classe_janela(hwnd);
        let tipo_app = obter_tipo_app(&class_name);

        if tipo_app.e_office() {
            let mut rc = RECT::default();
            let _ = GetWindowRect(hwnd, &mut rc);
            let w = rc.right - rc.left;
            let h = rc.bottom - rc.top;

            if w > 150 && h > 150 {
                let windows = &mut *(lparam.0 as *mut Vec<JanelaDocumento>);
                let titulo = extrair_texto_janela(hwnd);
                let titulo_final = if titulo.is_empty() {
                    "Documento".to_string()
                } else {
                    titulo
                };
                windows.push(JanelaDocumento {
                    hwnd,
                    titulo: titulo_final,
                    classe: class_name,
                });
            }
        }
    }));
    BOOL(1)
}

#[cfg(test)]
mod testes {
    use super::*;

    #[test]
    fn test_obter_altura_barra_padrao_e_personalizada() {
        let mut manager = TabManager::new();
        manager.config.general.tab_bar_height = 0;
        manager.config.theme.tab_height = 0;
        // Altura padrão do sistema
        assert_eq!(manager.obter_altura_barra(HWND::default()), 44);

        // Modifica altura geral
        manager.config.general.tab_bar_height = 34;
        assert_eq!(manager.obter_altura_barra(HWND::default()), 34);

        // Modifica com altura de tema
        manager.config.general.tab_bar_height = 0;
        manager.config.theme.tab_height = 30;
        assert_eq!(manager.obter_altura_barra(HWND::default()), 30);
    }

    #[test]
    fn test_obter_tipo_app_7_aplicativos() {
        assert_eq!(obter_tipo_app("OpusApp"), TipoAppOffice::Word);
        assert_eq!(obter_tipo_app("XLMAIN"), TipoAppOffice::Excel);
        assert_eq!(obter_tipo_app("EXCEL7"), TipoAppOffice::Excel);
        assert_eq!(obter_tipo_app("EXCEL2"), TipoAppOffice::Excel);
        assert_eq!(obter_tipo_app("PPTFrameClass"), TipoAppOffice::PowerPoint);
        assert_eq!(obter_tipo_app("PP9FrameClass"), TipoAppOffice::PowerPoint);
        assert_eq!(obter_tipo_app("PP10FrameClass"), TipoAppOffice::PowerPoint);
        assert_eq!(obter_tipo_app("MSACCESS"), TipoAppOffice::Access);
        assert_eq!(obter_tipo_app("OMain"), TipoAppOffice::Access);
        assert_eq!(obter_tipo_app("AccessApp"), TipoAppOffice::Access);
        assert_eq!(obter_tipo_app("PB"), TipoAppOffice::Project);
        assert_eq!(obter_tipo_app("WINPROJ"), TipoAppOffice::Project);
        assert_eq!(obter_tipo_app("VISIO"), TipoAppOffice::Visio);
        assert_eq!(obter_tipo_app("VISIOA"), TipoAppOffice::Visio);
        assert_eq!(obter_tipo_app("PUBLISHER"), TipoAppOffice::Publisher);
        assert_eq!(obter_tipo_app("MSPUB"), TipoAppOffice::Publisher);
        assert_eq!(obter_tipo_app("opusapp"), TipoAppOffice::Word);
        assert_eq!(obter_tipo_app("PBToolbar"), TipoAppOffice::Desconhecido);

        // Nomes amigáveis e flag de suporte
        assert!(TipoAppOffice::Word.e_office());
        assert_eq!(TipoAppOffice::Word.nome(), "Word");
        assert!(TipoAppOffice::Excel.e_office());
        assert_eq!(TipoAppOffice::Excel.nome(), "Excel");
        assert!(TipoAppOffice::PowerPoint.e_office());
        assert_eq!(TipoAppOffice::PowerPoint.nome(), "PowerPoint");
        assert!(TipoAppOffice::Access.e_office());
        assert_eq!(TipoAppOffice::Access.nome(), "Access");
        assert!(TipoAppOffice::Project.e_office());
        assert_eq!(TipoAppOffice::Project.nome(), "Project");
        assert!(TipoAppOffice::Visio.e_office());
        assert_eq!(TipoAppOffice::Visio.nome(), "Visio");
        assert!(TipoAppOffice::Publisher.e_office());
        assert_eq!(TipoAppOffice::Publisher.nome(), "Publisher");

        // Classe desconhecida
        let desc = obter_tipo_app("Notepad");
        assert!(!desc.e_office());
        assert_eq!(desc, TipoAppOffice::Desconhecido);
        assert_eq!(desc.nome(), "Desconhecido");
    }

    #[test]
    fn test_identificadores_menu_contexto_unicos() {
        let ids = [
            IDM_FECHAR,
            IDM_FECHAR_OUTRAS,
            IDM_FECHAR_DIREITA,
            IDM_COPIAR_CAMINHO,
            IDM_ABRIR_PASTA,
            IDM_NOVA_ABA,
            IDM_FECHAR_TODAS,
        ];
        let mut set = std::collections::HashSet::new();
        for &id in &ids {
            assert!(id > 0);
            assert!(set.insert(id), "ID duplicado no menu de contexto: {}", id);
        }
    }

    #[test]
    fn test_limpar_titulo_documento_7_aplicativos() {
        assert_eq!(
            limpar_titulo_documento("Proposta Comercial.docx - Word"),
            "Proposta Comercial.docx"
        );
        assert_eq!(
            limpar_titulo_documento("Planilha Orcamentaria.xlsx - Excel"),
            "Planilha Orcamentaria.xlsx"
        );
        assert_eq!(
            limpar_titulo_documento("Apresentacao.pptx - PowerPoint"),
            "Apresentacao.pptx"
        );
        assert_eq!(
            limpar_titulo_documento("BancoDados.accdb - Access"),
            "BancoDados.accdb"
        );
        assert_eq!(
            limpar_titulo_documento("Cronograma.mpp - Project"),
            "Cronograma.mpp"
        );
        assert_eq!(
            limpar_titulo_documento("Diagrama.vsdx - Visio"),
            "Diagrama.vsdx"
        );
        assert_eq!(
            limpar_titulo_documento("Folheto.pub - Publisher"),
            "Folheto.pub"
        );
        assert_eq!(
            limpar_titulo_documento("Documento Antigo.doc [Modo de Compatibilidade] - Word"),
            "Documento Antigo.doc"
        );
        assert_eq!(
            limpar_titulo_documento("Budget.xlsx [Compatibility Mode] - Excel"),
            "Budget.xlsx"
        );
        assert_eq!(
            limpar_titulo_documento("Pasta1.xlsx - Salvo - Excel"),
            "Pasta1.xlsx"
        );
        assert_eq!(
            limpar_titulo_documento("Relatorio.xlsx [Somente Leitura] - Excel"),
            "Relatorio.xlsx"
        );
        assert_eq!(
            limpar_titulo_documento("Relatorio.docx [Read-Only] - Word"),
            "Relatorio.docx"
        );
        assert_eq!(
            limpar_titulo_documento("Contrato.docx [Modo de Exibição Protegida] - Word"),
            "Contrato.docx"
        );
        assert_eq!(
            limpar_titulo_documento("Rascunho.docx - Não Salvo - Word"),
            "Rascunho.docx"
        );
        assert_eq!(limpar_titulo_documento(""), "Documento");
    }

    #[test]
    fn test_calcular_y_ribbon_com_dpi() {
        let y = calcular_y_ribbon(HWND::default());
        assert!(y >= 100, "Y ribbon deve ser maior que 100: {}", y);
    }

    #[test]
    fn test_calcular_altura_status_bar_default() {
        assert_eq!(calcular_altura_status_bar(HWND::default(), 600), 24);
    }

    #[test]
    fn test_pode_descarregar_inicial() {
        desativar_gerenciador_global();
        assert!(pode_descarregar());
    }

    #[test]
    fn test_despacho_com_hwnd_invalido_sem_panico() {
        // Assegurar thread safety e ausência de pânico em HWNDs nulos/inválidos
        despachar_timer(HWND::default());
        despachar_size(HWND::default());
        despachar_notify(HWND::default(), LPARAM(0));
        despachar_destroy(HWND::default());
        despachar_inicializar(HWND::default());
        despachar_recarregar(HWND::default());
        despachar_tab_mbutton_up(HWND::default(), 0, 0);
        despachar_tab_rbutton_up(HWND::default(), 0, 0);
        despachar_tab_context_menu(HWND::default(), 0, 0);
        despachar_tab_context_menu_teclado(HWND::default());
        despachar_tab_lbutton_dblclk(HWND::default(), 0, 0);
        despachar_tab_mousewheel(HWND::default(), 120);
        assert!(!despachar_keydown(HWND::default(), 0x09));
    }

    #[test]
    fn test_tratar_mousewheel_e_keydown_sem_panico() {
        let mut manager = TabManager::new();
        // Não deve entrar em pânico mesmo com lista de estados vazia
        manager.tratar_mousewheel(HWND::default(), 120);
        manager.tratar_mousewheel(HWND::default(), -120);
        manager.tratar_lbutton_dblclk(HWND::default(), 10, 10);
        assert!(!manager.tratar_keydown(HWND::default(), 0x09));
        assert!(!manager.tratar_keydown(HWND::default(), b'W' as usize));
        assert!(!manager.tratar_keydown(HWND::default(), b'T' as usize));
    }
}
