//! Gerenciador central de abas do Office (baseado no TOTWindowsManager).
//! Arquitetura REA: detecção na thread de background, criação/renderização
//! exclusivamente na thread de UI do Word via PostMessage.

use std::collections::HashMap;
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;

use windows::core::{w, PCWSTR, Result};
use windows::Win32::Foundation::{BOOL, HWND, LPARAM, POINT, RECT};
use windows::Win32::Graphics::Gdi::ScreenToClient;
use windows::Win32::System::Threading::GetCurrentProcessId;
use windows::Win32::UI::Controls::{NMHDR, TCN_SELCHANGE};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    keybd_event, KEYEVENTF_KEYUP, VK_CONTROL,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, FindWindowExW, GetClassNameW, GetClientRect, GetWindowRect,
    GetWindowTextW, GetWindowThreadProcessId, IsWindow, IsWindowVisible, KillTimer, MoveWindow,
    PostMessageW, SetForegroundWindow, SetTimer, ShowWindow, SW_RESTORE,
};

use crate::subclass::SubclassManager;
use crate::tab_control::TabControl;

// ---------------------------------------------------------------------------
// Mensagem customizada REA: PostMessage → thread UI → criação segura de janelas
// ---------------------------------------------------------------------------

/// Mensagem REA: a thread de background posta esta mensagem para o OpusApp
/// quando detecta que a janela ainda não foi subclassed/sincronizada.
/// A thread de UI do Word recebe e executa a criação do TabControl.
pub const WM_OFFICETAB_INICIALIZAR: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 0x01;

/// Mensagem REA: recarga forçada dos títulos/abas (posta pela thread de background
/// quando detecta mudança de documentos abertos).
pub const WM_OFFICETAB_RECARREGAR: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 0x02;

// ---------------------------------------------------------------------------
// Constantes
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
// TabManager — opera SOMENTE na thread de UI
// ---------------------------------------------------------------------------

/// Gerenciador central de abas e janelas do Office.
/// IMPORTANTE: todos os métodos que criam/movem janelas Win32 devem ser
/// chamados exclusivamente da thread de UI do OpusApp (via WM_TIMER, WM_APP, WM_SIZE, WM_NOTIFY).
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

    /// Descobre janelas visíveis das aplicações do Microsoft Office (Word, Excel, PowerPoint).
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

        // Ativar timer periódico de atualização (a cada 500ms — seguro pois
        // WM_TIMER é sempre entregue na thread de UI que criou a janela)
        unsafe {
            let _ = SetTimer(hwnd, OFFICETAB_TIMER_ID, 500, None);
        }

        // Executar sincronização inicial imediata (estamos na thread de UI)
        let res = self.sincronizar_janela(hwnd);
        log_debug(&format!("monitor_window: sincronizar res={:?}", res.is_ok()));

        res
    }

    /// Sincroniza o controle visual de abas e redimensionamento de uma janela específica.
    /// DEVE ser chamado da thread de UI do Office (WM_TIMER, WM_APP, WM_SIZE).
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

            // Se minimizada ou sem dimensões, ignorar temporariamente
            if largura_cliente <= 0 || altura_cliente <= 0 {
                return Ok(());
            }

            let estado = match self.estados.get_mut(&(hwnd.0 as isize)) {
                Some(e) => e,
                None => return Ok(()),
            };

            // Localizar janela interna de documento (Word: _WwF, Excel: XLDESK / EXCEL7)
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
                // Ainda na tela inicial de boas-vindas sem documento ativo
                return Ok(());
            };

            let mut rc_doc = RECT::default();
            let _ = GetWindowRect(doc_view, &mut rc_doc);
            let mut pt = POINT {
                x: rc_doc.left,
                y: rc_doc.top,
            };
            let _ = ScreenToClient(hwnd, &mut pt);

            // Guardar Y base original logo abaixo do Ribbon
            if estado.y_base <= 0 || (pt.y < estado.y_base && pt.y > 40) {
                estado.y_base = if pt.y > 40 { pt.y } else { 178 };
            }
            estado.hwnd_doc_view = Some(doc_view);

            // Criar controle de abas se ainda não existir (estamos na thread de UI — seguro)
            if estado.tab_control.is_none() {
                match TabControl::new(
                    hwnd,
                    OFFICETAB_CONTROL_ID,
                    0,
                    estado.y_base,
                    largura_cliente,
                    ALTURA_BARRA_ABAS,
                ) {
                    Ok(tab) => {
                        log_debug(&format!(
                            "TabControl CRIADO (thread UI): HWND=0x{:X}, pai=0x{:X}",
                            tab.hwnd.0 as usize, hwnd.0 as usize
                        ));
                        estado.tab_control = Some(tab);
                    }
                    Err(e) => {
                        log_debug(&format!("ERRO AO CRIAR TabControl: {:?}", e));
                        return Err(e);
                    }
                }
            } else if let Some(ref tab) = estado.tab_control {
                let _ = tab.resize(0, estado.y_base, largura_cliente, ALTURA_BARRA_ABAS);
            }

            // Deslocar a área do documento para baixo da barra de abas
            if let Some(dv) = estado.hwnd_doc_view {
                let y_doc = estado.y_base + ALTURA_BARRA_ABAS;
                let h_doc = (altura_cliente - y_doc).max(10);
                let _ = MoveWindow(dv, 0, y_doc, largura_cliente, h_doc, true);
            }

            // Coletar documentos abertos para popular as abas
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

                    // Aba de criação "+"
                    let _ = tab.add_tab(docs_processo.len() as i32, "+");

                    estado.doc_hwnds = docs_processo.iter().map(|(h, _)| *h).collect();
                    estado.titulos_cache = novos_titulos;
                } else if let Some(pos) = estado.doc_hwnds.iter().position(|&h| h == hwnd)
                    && tab.get_selected_tab() != pos as i32
                {
                    let _ = tab.set_selected_tab(pos as i32);
                }
            }

            Ok(())
        }
    }

    /// Processa notificações do TabControl (clique em aba / nova aba).
    pub fn processar_notify(&mut self, hwnd: HWND, lparam: LPARAM) {
        if lparam.0 == 0 {
            return;
        }

        let pnmhdr = unsafe { &*(lparam.0 as *const NMHDR) };
        if pnmhdr.idFrom == OFFICETAB_CONTROL_ID as usize
            && pnmhdr.code == TCN_SELCHANGE
            && let Some(estado) = self.estados.get_mut(&(hwnd.0 as isize))
            && let Some(ref tab) = estado.tab_control
        {
            let sel = tab.get_selected_tab();
            let total_docs = estado.doc_hwnds.len();

            if sel >= 0 && (sel as usize) < total_docs {
                let alvo = estado.doc_hwnds[sel as usize];
                if alvo != hwnd {
                    unsafe {
                        let _ = ShowWindow(alvo, SW_RESTORE);
                        let _ = SetForegroundWindow(alvo);
                    }
                }
            } else if (sel as usize) == total_docs {
                // Clicou na aba "+" (Novo Documento)
                criar_novo_documento();

                // Re-seleciona a aba da janela atual até o novo documento abrir
                if let Some(pos) = estado.doc_hwnds.iter().position(|&h| h == hwnd) {
                    let _ = tab.set_selected_tab(pos as i32);
                }
            }
        }
    }

    /// Trata a destruição de uma janela.
    pub fn processar_destroy(&mut self, hwnd: HWND) {
        unsafe {
            let _ = KillTimer(hwnd, OFFICETAB_TIMER_ID);
        }
        if let Some(mut estado) = self.estados.remove(&(hwnd.0 as isize))
            && let Some(ref mut tab) = estado.tab_control
        {
            tab.destroy();
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
        self.cleanup();
    }
}

impl Default for TabManager {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// Singleton Global, Despacho de Eventos Subclass e Thread de Monitoramento
// ---------------------------------------------------------------------------

static GLOBAL_TAB_MANAGER: Mutex<Option<TabManager>> = Mutex::new(None);
static MONITORAMENTO_ATIVO: AtomicBool = AtomicBool::new(false);

/// Inicia uma thread de background para APENAS detectar novas janelas do Office.
/// A thread NUNCA cria ou move janelas Win32 — apenas posta WM_OFFICETAB_INICIALIZAR
/// UMA ÚNICA VEZ por HWND usando um HashSet local, eliminando o flood de mensagens.
pub fn iniciar_thread_monitoramento() {
    if MONITORAMENTO_ATIVO.swap(true, Ordering::SeqCst) {
        return; // Já está rodando
    }

    std::thread::spawn(|| {
        log_debug("Thread de monitoramento REA iniciada");

        // HashSet LOCAL à thread — rastreia HWNDs que já receberam PostMessage.
        // Não depende do Mutex global, portanto nunca bloqueia e nunca re-posta.
        let mut ja_postados: std::collections::HashSet<isize> = std::collections::HashSet::new();

        while MONITORAMENTO_ATIVO.load(Ordering::Relaxed) {
            let janelas = listar_janelas_processo();
            let hwnds_atuais: std::collections::HashSet<isize> =
                janelas.iter().map(|(h, _)| h.0 as isize).collect();

            // Remover HWNDs que não existem mais (janela fechada) para permitir reinicialização
            ja_postados.retain(|raw| hwnds_atuais.contains(raw));

            for (hwnd, _titulo) in &janelas {
                let raw = hwnd.0 as isize;

                // Postar APENAS se ainda não foi postado para este HWND nesta sessão
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
                                "REA: PostMessage INICIALIZAR → HWND=0x{:X}",
                                hwnd.0 as usize
                            ));
                        }
                    }
                }
            }

            std::thread::sleep(std::time::Duration::from_millis(800));
        }

        log_debug("Thread de monitoramento REA encerrada");
    });
}

pub fn obter_ou_criar_gerenciador() {
    let mut guard = GLOBAL_TAB_MANAGER.lock().unwrap();
    if guard.is_none() {
        *guard = Some(TabManager::new());
    }
    drop(guard); // Libera o lock antes de iniciar a thread
    iniciar_thread_monitoramento();
}

// ---------------------------------------------------------------------------
// Despacho de eventos de janela (SOMENTE chamados da thread de UI do Office)
// ---------------------------------------------------------------------------

/// Chamado via WM_TIMER na thread de UI do OpusApp.
/// Atualiza abas e layout — seguro pois WM_TIMER sempre chega na thread correta.
pub fn despachar_timer(hwnd: HWND) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock()
        && let Some(ref mut manager) = *guard
    {
        let _ = manager.sincronizar_janela(hwnd);
    }
}

/// Chamado via WM_SIZE na thread de UI do OpusApp.
pub fn despachar_size(hwnd: HWND) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock()
        && let Some(ref mut manager) = *guard
    {
        let _ = manager.sincronizar_janela(hwnd);
    }
}

/// Chamado via WM_NOTIFY na thread de UI do OpusApp.
pub fn despachar_notify(hwnd: HWND, lparam: LPARAM) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock()
        && let Some(ref mut manager) = *guard
    {
        manager.processar_notify(hwnd, lparam);
    }
}

/// Chamado via WM_DESTROY/WM_NCDESTROY na thread de UI do OpusApp.
pub fn despachar_destroy(hwnd: HWND) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock()
        && let Some(ref mut manager) = *guard
    {
        manager.processar_destroy(hwnd);
    }
}

/// Chamado via WM_OFFICETAB_INICIALIZAR na thread de UI do OpusApp.
/// Realiza o subclassing e a criação do TabControl com segurança de thread.
pub fn despachar_inicializar(hwnd: HWND) {
    log_debug(&format!("despachar_inicializar (thread UI): HWND=0x{:X}", hwnd.0 as usize));

    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock() {
        if guard.is_none() {
            *guard = Some(TabManager::new());
        }
        if let Some(ref mut manager) = *guard {
            if !manager.esta_monitorado(hwnd) {
                let res = manager.monitor_window(hwnd);
                log_debug(&format!("despachar_inicializar: monitor_window res={:?}", res.is_ok()));
            }
        }
    }
}

/// Chamado via WM_OFFICETAB_RECARREGAR na thread de UI do OpusApp.
/// Força a recarga de abas quando detectada mudança de documentos.
pub fn despachar_recarregar(hwnd: HWND) {
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock()
        && let Some(ref mut manager) = *guard
    {
        // Forçar limpeza do cache de títulos para garantir recarga
        if let Some(estado) = manager.estados.get_mut(&(hwnd.0 as isize)) {
            estado.titulos_cache.clear();
        }
        let _ = manager.sincronizar_janela(hwnd);
    }
}

/// Varre e anexa o gerenciamento às janelas ativas do Office.
/// ATENÇÃO: Este método realiza operações de UI — chamar apenas da thread de UI!
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
    if let Ok(mut guard) = GLOBAL_TAB_MANAGER.lock()
        && let Some(ref mut manager) = *guard
    {
        manager.cleanup();
    }
}

// ---------------------------------------------------------------------------
// Funções Auxiliares Win32
// ---------------------------------------------------------------------------

/// Envia o atalho Ctrl+N para abrir um novo documento na aplicação hospedeira.
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

/// Limpa o título da janela retirando sufixos padrão do Office.
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

/// Lista todas as janelas do Office do processo atual.
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
        let mut pid = 0u32;
        GetWindowThreadProcessId(hwnd, Some(&mut pid));
        let current_pid = GetCurrentProcessId();

        if pid == current_pid {
            let class_name = extrair_classe_janela(hwnd);
            let is_office = class_name == "OpusApp"
                || class_name == "XLMAIN"
                || class_name == "PPTFrameClass";

            if is_office {
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

    BOOL(1)
}
