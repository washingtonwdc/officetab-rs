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
//! Arquitetura REA: detecção via WinEventHook + Timer na thread UI do Word + thread fallback.
//! Criação/renderização de janelas Win32 exclusivamente na thread de UI.

use std::collections::HashMap;
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;

use windows::core::{w, PCWSTR, Result};
use windows::Win32::Foundation::{BOOL, HWND, LPARAM, POINT, RECT};
use windows::Win32::Graphics::Gdi::ScreenToClient;
use windows::Win32::System::Threading::GetCurrentProcessId;
use windows::Win32::UI::Accessibility::{
    SetWinEventHook, UnhookWinEvent, HWINEVENTHOOK,
};
use windows::Win32::UI::Controls::{NMHDR, TCN_SELCHANGE};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    keybd_event, KEYEVENTF_KEYUP, VK_CONTROL,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BringWindowToTop, EnumWindows, FindWindowExW, GetClassNameW, GetClientRect,
    GetWindowLongPtrW, GetWindowRect, GetWindowTextW, GetWindowThreadProcessId,
    IsWindow, IsWindowVisible, KillTimer, MoveWindow, PostMessageW,
    SetForegroundWindow, SetTimer, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    GWL_STYLE, HWND_TOP, SWP_SHOWWINDOW, SW_RESTORE, SW_SHOW, WS_CLIPCHILDREN,
    EVENT_OBJECT_SHOW, WINEVENT_OUTOFCONTEXT,
};

use crate::subclass::SubclassManager;
use crate::tab_control::TabControl;

// ---------------------------------------------------------------------------
// Mensagens customizadas REA
// ---------------------------------------------------------------------------

/// Mensagem REA: fallback — postada pela thread de polling quando WinEventHook não disponível.
pub const WM_OFFICETAB_INICIALIZAR: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 0x01;

/// Mensagem REA: recarga forçada dos títulos/abas.
pub const WM_OFFICETAB_RECARREGAR: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 0x02;

// ---------------------------------------------------------------------------
// Constantes e utilitários de log
// ---------------------------------------------------------------------------

pub fn log_debug(msg: &str) {
    use std::io::Write;
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open("C:\\Users\\washi\\officetab-rs\\officetab.log")
    {
        let _ = writeln!(f, "[LOG] {}", msg);
    }
}

pub const OFFICETAB_CONTROL_ID: u32 = 5001;
pub const OFFICETAB_TIMER_ID: usize = 0x544142; // 'TAB'
pub const ALTURA_BARRA_ABAS: i32 = 28;

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

/// Estado interno e visual associado a cada janela principal de documento do Office
pub struct EstadoJanela {
    pub hwnd: HWND,
    pub tab_control: Option<TabControl>,
    pub hwnd_doc_view: Option<HWND>,
    pub y_base: i32,
    pub doc_hwnds: Vec<HWND>,
    pub titulos_cache: Vec<String>,
}

unsafe impl Send for EstadoJanela {}
unsafe impl Sync for EstadoJanela {}

// ---------------------------------------------------------------------------
// Cálculo Dinâmico da Posição do Ribbon
// ---------------------------------------------------------------------------

/// Calcula a coordenada Y logo abaixo do Ribbon para ancorar a barra de abas.
/// Detecta dinamicamente DropShadow, MsoCommandBarDock e NetUIHWND no topo.
pub fn calcular_y_ribbon(hwnd_top: HWND) -> i32 {
    unsafe {
        // 1. Procurar DropShadow com pt.y > 40
        let mut child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, w!("DropShadow"), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                let _ = GetWindowRect(child, &mut rc);
                let mut pt = POINT { x: rc.left, y: rc.top };
                let _ = ScreenToClient(hwnd_top, &mut pt);
                let h = rc.bottom - rc.top;
                if pt.y > 40 && pt.x == 0 {
                    return pt.y + h; // Borda inferior do DropShadow
                }
            }
        }

        // 2. Procurar MsoCommandBarDock no topo (pt.x == 0, pt.y == 0, h > 40 e h < 500)
        child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, w!("MsoCommandBarDock"), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                let _ = GetWindowRect(child, &mut rc);
                let mut pt = POINT { x: rc.left, y: rc.top };
                let _ = ScreenToClient(hwnd_top, &mut pt);
                let h = rc.bottom - rc.top;
                if pt.x == 0 && pt.y <= 10 && h > 40 && h < 500 {
                    return pt.y + h;
                }
            }
        }

        // 3. Procurar NetUIHWND no topo (pt.x == 0, pt.y == 0, h > 40 e h < 500)
        child = HWND::default();
        loop {
            child = match FindWindowExW(hwnd_top, child, w!("NetUIHWND"), PCWSTR::null()) {
                Ok(h) if !h.is_invalid() => h,
                _ => break,
            };
            if IsWindowVisible(child).as_bool() {
                let mut rc = RECT::default();
                let _ = GetWindowRect(child, &mut rc);
                let mut pt = POINT { x: rc.left, y: rc.top };
                let _ = ScreenToClient(hwnd_top, &mut pt);
                let h = rc.bottom - rc.top;
                if pt.x == 0 && pt.y <= 10 && h > 40 && h < 500 {
                    return pt.y + h;
                }
            }
        }

        // 4. Fallback pelo topo do documento se ainda não deslocado
        let wwf = FindWindowExW(hwnd_top, HWND::default(), w!("_WwF"), PCWSTR::null()).unwrap_or_default();
        let doc_view = if !wwf.is_invalid() {
            wwf
        } else {
            FindWindowExW(hwnd_top, HWND::default(), w!("XLDESK"), PCWSTR::null()).unwrap_or_default()
        };
        if !doc_view.is_invalid() {
            let mut rc = RECT::default();
            let _ = GetWindowRect(doc_view, &mut rc);
            let mut pt = POINT { x: rc.left, y: rc.top };
            let _ = ScreenToClient(hwnd_top, &mut pt);
            if pt.y > 50 {
                return pt.y;
            }
        }

        178 // Padrão da fita moderna do Office (Office 2013-365)
    }
}

// ---------------------------------------------------------------------------
// TabManager — opera SOMENTE na thread de UI
// ---------------------------------------------------------------------------

/// Gerenciador central de abas e janelas do Office.
pub struct TabManager {
    subclass_manager: SubclassManager,
    estados: HashMap<isize, EstadoJanela>,
}

unsafe impl Send for TabManager {}
unsafe impl Sync for TabManager {}

impl TabManager {
    pub fn new() -> Self {
        Self {
            subclass_manager: SubclassManager::new(),
            estados: HashMap::new(),
        }
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

        log_debug(&format!("monitor_window (thread UI): HWND=0x{:X}", hwnd.0 as usize));

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
            },
        );

        // Ativar timer periódico de atualização (400ms)
        unsafe {
            let _ = SetTimer(hwnd, OFFICETAB_TIMER_ID, 400, None);
        }

        // Sincronização inicial imediata
        let res = self.sincronizar_janela(hwnd);
        log_debug(&format!("monitor_window: sincronizar res={:?}", res.is_ok()));

        res
    }

    /// Sincroniza o controle visual de abas e redimensionamento.
    /// DEVE ser chamado da thread de UI do Office.
    pub fn sincronizar_janela(&mut self, hwnd: HWND) -> Result<()> {
        unsafe {
            if !IsWindow(hwnd).as_bool() {
                self.estados.remove(&(hwnd.0 as isize));
                return Ok(());
            }

            let mut rc_client = RECT::default();
            GetClientRect(hwnd, &mut rc_client)?;

            let largura_cliente = rc_client.right - rc_client.left;
            let altura_cliente = rc_client.bottom - rc_client.top;

            if largura_cliente <= 100 || altura_cliente <= 100 {
                return Ok(());
            }

            let estado = match self.estados.get_mut(&(hwnd.0 as isize)) {
                Some(e) => e,
                None => return Ok(()),
            };

            // Localizar janela interna de documento
            let wwf = FindWindowExW(hwnd, HWND::default(), w!("_WwF"), PCWSTR::null())
                .unwrap_or_default();
            let xldesk = if wwf.is_invalid() {
                FindWindowExW(hwnd, HWND::default(), w!("XLDESK"), PCWSTR::null())
                    .unwrap_or_default()
            } else {
                HWND::default()
            };

            let doc_view = if !wwf.is_invalid() {
                wwf
            } else if !xldesk.is_invalid() {
                xldesk
            } else {
                return Ok(());
            };

            estado.hwnd_doc_view = Some(doc_view);

            // Calcular y_base DINAMICAMENTE para acompanhar ribbon recolhido/expandido
            let y_ribbon = calcular_y_ribbon(hwnd);
            estado.y_base = y_ribbon;

            // Criar controle de abas se ainda não existir
            if estado.tab_control.is_none() {
                match TabControl::new(
                    hwnd,
                    OFFICETAB_CONTROL_ID,
                    0,
                    y_ribbon,
                    largura_cliente,
                    ALTURA_BARRA_ABAS,
                ) {
                    Ok(tab) => {
                        log_debug(&format!(
                            "TabControl CRIADO (thread UI): HWND=0x{:X}, pai=0x{:X}, y={}",
                            tab.hwnd.0 as usize,
                            hwnd.0 as usize,
                            y_ribbon
                        ));
                        estado.tab_control = Some(tab);
                    }
                    Err(e) => {
                        log_debug(&format!("ERRO AO CRIAR TabControl: {:?}", e));
                        return Err(e);
                    }
                }
            } else if let Some(ref tab) = estado.tab_control {
                let _ = tab.resize(0, y_ribbon, largura_cliente, ALTURA_BARRA_ABAS);
                let _ = SetWindowPos(
                    tab.hwnd,
                    HWND_TOP,
                    0,
                    y_ribbon,
                    largura_cliente,
                    ALTURA_BARRA_ABAS,
                    SWP_SHOWWINDOW,
                );
                let _ = BringWindowToTop(tab.hwnd);
                let _ = ShowWindow(tab.hwnd, SW_SHOW);
            }

            // Deslocar área do documento para baixo da barra de abas
            if let Some(dv) = estado.hwnd_doc_view {
                let y_doc = y_ribbon + ALTURA_BARRA_ABAS;
                let h_doc = (altura_cliente - y_doc - 24).max(10);
                let _ = MoveWindow(dv, 0, y_doc, largura_cliente, h_doc, true);
            }

            // Popular abas com documentos abertos
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

            let novos_titulos: Vec<String> =
                docs_processo.iter().map(|(_, t)| t.clone()).collect();

            if let Some(ref tab) = estado.tab_control {
                if novos_titulos != estado.titulos_cache {
                    tab.clear_tabs();

                    for (idx, (doc_hwnd, doc_titulo)) in docs_processo.iter().enumerate() {
                        let _ = tab.add_tab(idx as i32, doc_titulo);
                        if *doc_hwnd == hwnd {
                            let _ = tab.set_selected_tab(idx as i32);
                        }
                    }

                    // Aba "+" para novo documento
                    let _ = tab.add_tab(docs_processo.len() as i32, "+");

                    estado.doc_hwnds = docs_processo.iter().map(|(h, _)| *h).collect();
                    estado.titulos_cache = novos_titulos;
                    tab.repaint();
                } else if let Some(pos) = estado.doc_hwnds.iter().position(|&h| h == hwnd) {
                    if tab.get_selected_tab() != pos as i32 {
                        let _ = tab.set_selected_tab(pos as i32);
                    }
                    tab.repaint();
                }
            }

            Ok(())
        }
    }

    /// Processa notificações do TabControl (clique em aba).
    pub fn processar_notify(&mut self, hwnd: HWND, lparam: LPARAM) {
        if lparam.0 == 0 {
            return;
        }

        let pnmhdr = unsafe { &*(lparam.0 as *const NMHDR) };
        if pnmhdr.idFrom != OFFICETAB_CONTROL_ID as usize || pnmhdr.code != TCN_SELCHANGE {
            return;
        }

        let (sel, total_docs, doc_hwnds) =
            if let Some(estado) = self.estados.get(&(hwnd.0 as isize)) {
                if let Some(ref tab) = estado.tab_control {
                    (tab.get_selected_tab(), estado.doc_hwnds.len(), estado.doc_hwnds.clone())
                } else {
                    return;
                }
            } else {
                return;
            };

        if sel >= 0 && (sel as usize) < total_docs {
            let alvo = doc_hwnds[sel as usize];
            if alvo != hwnd {
                unsafe {
                    let _ = ShowWindow(alvo, SW_RESTORE);
                    let _ = SetForegroundWindow(alvo);
                    let _ = BringWindowToTop(alvo);
                }
            }
        } else if (sel as usize) == total_docs {
            // Clicou na aba "+"
            criar_novo_documento();
            if let Some(estado) = self.estados.get(&(hwnd.0 as isize)) {
                if let Some(ref tab) = estado.tab_control {
                    if let Some(pos) = doc_hwnds.iter().position(|&h| h == hwnd) {
                        let _ = tab.set_selected_tab(pos as i32);
                    }
                }
            }
        }
    }

    /// Trata a destruição de uma janela.
    pub fn processar_destroy(&mut self, hwnd: HWND) {
        unsafe {
            let _ = KillTimer(hwnd, OFFICETAB_TIMER_ID);
        }
        if let Some(mut estado) = self.estados.remove(&(hwnd.0 as isize)) {
            if let Some(ref mut tab) = estado.tab_control {
                tab.destroy();
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
        self.subclass_manager.unsubclass_all();
    }

    /// Retorna true se o HWND já está sendo monitorado.
    pub fn esta_monitorado(&self, hwnd: HWND) -> bool {
        self.estados.contains_key(&(hwnd.0 as isize))
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

/// Callback periódico executado DIRETAMENTE NA THREAD DE UI do Office via SetTimer.
/// Garante que novas janelas de documentos visíveis sejam imediatamente subclassed e sincronizadas.
unsafe extern "system" fn timer_ui_callback(
    _hwnd: HWND,
    _msg: u32,
    _id: usize,
    _time: u32,
) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if guard.is_none() {
            *guard = Some(TabManager::new());
        }
        if let Some(ref mut manager) = *guard {
            let janelas = listar_janelas_processo();
            for (h, _) in janelas {
                if !manager.esta_monitorado(h) {
                    log_debug(&format!(
                        "timer_ui_callback: nova janela Office detectada HWND=0x{:X}",
                        h.0 as usize
                    ));
                    let _ = manager.monitor_window(h);
                } else {
                    let _ = manager.sincronizar_janela(h);
                }
            }
        }
    }
}

/// Callback do WinEventHook — executado na thread de UI do processo que gerou o evento.
unsafe extern "system" fn winevent_callback(
    _hook: HWINEVENTHOOK,
    _event: u32,
    hwnd: HWND,
    id_object: i32,
    _id_child: i32,
    _id_event_thread: u32,
    _dwms_event_time: u32,
) {
    if id_object != 0 || hwnd.is_invalid() {
        return;
    }

    unsafe {
        if !IsWindowVisible(hwnd).as_bool() {
            return;
        }
    }

    let classe = extrair_classe_janela(hwnd);
    let e_office = classe == "OpusApp" || classe == "XLMAIN" || classe == "PPTFrameClass";
    if !e_office {
        return;
    }

    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if guard.is_none() {
            *guard = Some(TabManager::new());
        }
        if let Some(ref mut manager) = *guard {
            if !manager.esta_monitorado(hwnd) {
                log_debug(&format!(
                    "WinEventHook: janela Office detectada HWND=0x{:X} classe={}",
                    hwnd.0 as usize,
                    classe
                ));
                let _ = manager.monitor_window(hwnd);
            }
        }
    }
}

/// Instala WinEventHook para detectar novas janelas Office no processo atual.
fn instalar_winevent_hook() -> bool {
    let mut guard = WINEVENT_HOOK.lock().unwrap();
    if guard.is_some() {
        return true;
    }

    unsafe {
        let hook = SetWinEventHook(
            EVENT_OBJECT_SHOW,
            EVENT_OBJECT_SHOW,
            windows::Win32::Foundation::HMODULE::default(),
            Some(winevent_callback),
            GetCurrentProcessId(),
            0,
            WINEVENT_OUTOFCONTEXT,
        );

        if !hook.is_invalid() {
            log_debug(&format!("WinEventHook instalado: handle=0x{:x}", hook.0 as usize));
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

/// Thread de monitoramento fallback para segurança redundante
fn iniciar_thread_monitoramento_fallback() {
    if MONITORAMENTO_ATIVO.swap(true, Ordering::SeqCst) {
        return;
    }

    std::thread::spawn(move || {
        log_debug("Thread de monitoramento REA (fallback) iniciada");
        let mut ja_postados: std::collections::HashSet<isize> = std::collections::HashSet::new();

        while MONITORAMENTO_ATIVO.load(Ordering::Relaxed) {
            let janelas = listar_janelas_processo();
            let hwnds_atuais: std::collections::HashSet<isize> =
                janelas.iter().map(|(h, _)| h.0 as isize).collect();

            ja_postados.retain(|raw| hwnds_atuais.contains(raw));

            for (hwnd, _) in &janelas {
                let raw = hwnd.0 as isize;
                if !ja_postados.contains(&raw) {
                    unsafe {
                        let ok = PostMessageW(
                            *hwnd,
                            WM_OFFICETAB_INICIALIZAR,
                            windows::Win32::Foundation::WPARAM(0),
                            windows::Win32::Foundation::LPARAM(0),
                        );
                        if ok.is_ok() {
                            ja_postados.insert(raw);
                            log_debug(&format!(
                                "REA fallback: PostMessage INICIALIZAR → HWND=0x{:X}",
                                hwnd.0 as usize
                            ));
                        }
                    }
                }
            }

            std::thread::sleep(std::time::Duration::from_millis(500));
        }

        log_debug("Thread de monitoramento REA (fallback) encerrada");
    });
}

/// Varre janelas Office já existentes e inicia monitoramento.
fn varrer_janelas_existentes() {
    let janelas = listar_janelas_processo();
    if janelas.is_empty() {
        return;
    }

    log_debug(&format!("Varrendo {} janelas Office já abertas...", janelas.len()));

    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if guard.is_none() {
            *guard = Some(TabManager::new());
        }
        if let Some(ref mut manager) = *guard {
            for (hwnd, _) in &janelas {
                if !manager.esta_monitorado(*hwnd) {
                    let _ = manager.monitor_window(*hwnd);
                }
            }
        }
    }
}

/// Inicia o monitoramento completo do Office Tab
pub fn iniciar_thread_monitoramento() {
    let _ = instalar_winevent_hook();

    // Instalar timer na thread de UI (mensagens disparadas no loop principal do Word)
    unsafe {
        let mut guard = TIMER_UI_ID.lock().unwrap();
        if guard.is_none() {
            let tid = SetTimer(HWND::default(), 0, 350, Some(timer_ui_callback));
            if tid != 0 {
                *guard = Some(tid);
                log_debug(&format!("Timer de UI instalado: id={}", tid));
            }
        }
    }

    iniciar_thread_monitoramento_fallback();
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

/// Chamado via WM_TIMER.
pub fn despachar_timer(hwnd: HWND) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if let Some(ref mut manager) = *guard {
            let _ = manager.sincronizar_janela(hwnd);
        }
    }
}

/// Chamado via WM_SIZE.
pub fn despachar_size(hwnd: HWND) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if let Some(ref mut manager) = *guard {
            let _ = manager.sincronizar_janela(hwnd);
        }
    }
}

/// Chamado via WM_NOTIFY.
pub fn despachar_notify(hwnd: HWND, lparam: LPARAM) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if let Some(ref mut manager) = *guard {
            manager.processar_notify(hwnd, lparam);
        }
    }
}

/// Chamado via WM_DESTROY/WM_NCDESTROY.
pub fn despachar_destroy(hwnd: HWND) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if let Some(ref mut manager) = *guard {
            manager.processar_destroy(hwnd);
        }
    }
}

/// Chamado via WM_OFFICETAB_INICIALIZAR.
pub fn despachar_inicializar(hwnd: HWND) {
    log_debug(&format!("despachar_inicializar (thread UI): HWND=0x{:X}", hwnd.0 as usize));

    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if guard.is_none() {
            *guard = Some(TabManager::new());
        }
        if let Some(ref mut manager) = *guard {
            if !manager.esta_monitorado(hwnd) {
                let res = manager.monitor_window(hwnd);
                log_debug(&format!(
                    "despachar_inicializar: monitor_window res={}",
                    res.is_ok()
                ));
            }
        }
    }
}

/// Chamado via WM_OFFICETAB_RECARREGAR.
pub fn despachar_recarregar(hwnd: HWND) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if let Some(ref mut manager) = *guard {
            if let Some(estado) = manager.estados.get_mut(&(hwnd.0 as isize)) {
                estado.titulos_cache.clear();
            }
            let _ = manager.sincronizar_janela(hwnd);
        }
    }
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
    }
}

// ---------------------------------------------------------------------------
// Funções Auxiliares Win32
// ---------------------------------------------------------------------------

fn criar_novo_documento() {
    unsafe {
        keybd_event(VK_CONTROL.0 as u8, 0, Default::default(), 0);
        keybd_event(b'N', 0, Default::default(), 0);
        keybd_event(b'N', 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_CONTROL.0 as u8, 0, KEYEVENTF_KEYUP, 0);
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

/// Limpa sufixos padrão do Office do título da janela.
fn limpar_titulo_documento(titulo: &str) -> String {
    let limpo = titulo
        .trim_end_matches(" - Word")
        .trim_end_matches(" - Excel")
        .trim_end_matches(" - PowerPoint")
        .trim_end_matches(" [Modo de Compatibilidade]")
        .trim();

    if limpo.is_empty() {
        "Documento".to_string()
    } else {
        limpo.to_string()
    }
}

/// Lista todas as janelas Office visíveis do processo atual.
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
    unsafe {
        if !IsWindowVisible(hwnd).as_bool() {
            return BOOL(1);
        }

        let mut pid = 0u32;
        GetWindowThreadProcessId(hwnd, Some(&mut pid));
        let current_pid = GetCurrentProcessId();

        if pid == current_pid {
            let class_name = extrair_classe_janela(hwnd);
            let is_office = class_name == "OpusApp"
                || class_name == "XLMAIN"
                || class_name == "PPTFrameClass";

            if is_office {
                let mut rc = RECT::default();
                let _ = GetWindowRect(hwnd, &mut rc);
                let w = rc.right - rc.left;
                let h = rc.bottom - rc.top;

                // Filtrar apenas janelas reais com tamanho visível
                if w > 150 && h > 150 {
                    let titulo = extrair_texto_janela(hwnd);
                    let titulo_limpo = if titulo.is_empty() {
                        "Documento".to_string()
                    } else {
                        limpar_titulo_documento(&titulo)
                    };
                    let lista = &mut *(lparam.0 as *mut Vec<(HWND, String)>);
                    lista.push((hwnd, titulo_limpo));
                }
            }
        }
    }
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
    unsafe {
        if !IsWindowVisible(hwnd).as_bool() {
            return BOOL(1);
        }

        let class_name = extrair_classe_janela(hwnd);
        let is_office = class_name.contains("OpusApp")
            || class_name.contains("XLMAIN")
            || class_name.contains("EXCEL7")
            || class_name.contains("PPTFrameClass");

        if is_office {
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
    }
    BOOL(1)
}
