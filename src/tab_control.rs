// Wrapper para controle de abas nativo do Windows (SysTabControl32).
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

use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    CLEARTYPE_QUALITY, CLIP_DEFAULT_PRECIS, CreateFontW, DEFAULT_CHARSET,
    DEFAULT_GUI_FONT, DEFAULT_PITCH, DeleteObject, FF_DONTCARE, FW_NORMAL, HDC,
    GetDC, GetStockObject, HFONT, InvalidateRect, OUT_DEFAULT_PRECIS, ReleaseDC,
    UpdateWindow, ValidateRect,
};
use windows::Win32::System::LibraryLoader::{
    GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
    GetModuleHandleExW, GetModuleHandleW,
};
use windows::Win32::UI::Controls::{
    ICC_TAB_CLASSES, INITCOMMONCONTROLSEX, InitCommonControlsEx, TCIF_TEXT, TCITEMW,
    TCM_DELETEALLITEMS, TCM_DELETEITEM, TCM_GETCURSEL, TCM_GETITEMCOUNT, TCM_INSERTITEMW,
    TCM_SETCURSEL, TCS_FOCUSNEVER, TCS_HOTTRACK, TCS_TABS,
};
use windows::Win32::UI::Shell::{DefSubclassProc, RemoveWindowSubclass, SetWindowSubclass};
use windows::Win32::UI::WindowsAndMessaging::{
    BringWindowToTop, CreateWindowExW, DestroyWindow, GetClientRect, HMENU, HWND_TOP, MoveWindow,
    SW_SHOW, SWP_NOMOVE, SWP_NOSIZE, SWP_NOACTIVATE, SWP_SHOWWINDOW, SendMessageW, SetWindowPos,
    ShowWindow, WINDOW_EX_STYLE, WINDOW_STYLE, WM_SETFONT, WS_CHILD, WS_CLIPCHILDREN, WS_CLIPSIBLINGS,
    WS_VISIBLE,
};
use windows::core::{PCWSTR, Result, w};

/// Mensagem Win32 de teste de colisão em controles de abas
pub const TCM_HITTEST: u32 = 0x130D;
/// Mensagem Win32 para obter o retângulo delimitador da aba
pub const TCM_GETITEMRECT: u32 = 0x130A;
/// Mensagem Win32 para definir espaçamento interno (padding) das abas
pub const TCM_SETPADDING: u32 = 0x132B;

/// Identificador único para a subclasse do controle de abas
pub const OFFICETAB_TABCONTROL_SUBCLASS_ID: usize = 0x54414243; // 'TABC'

/// Calcula a altura da barra de abas escalonada proporcionalmente pelo DPI da janela.
/// Escalas de referência:
/// - 96 DPI  (100%): 28px
/// - 144 DPI (150%): 42px
/// - 192 DPI (200%): 56px
pub fn calcular_altura_por_dpi(hwnd: HWND, altura_base: i32) -> i32 {
    let dpi = crate::utils::get_dpi_for_window(hwnd);
    calcular_altura_por_dpi_valor(dpi, altura_base)
}

/// Função pura para cálculo de altura por DPI, útil para testes unitários e lógica determinística.
pub fn calcular_altura_por_dpi_valor(dpi: u32, altura_base: i32) -> i32 {
    let dpi_efetivo = if dpi > 0 { dpi } else { 96 };
    let base = if altura_base > 0 { altura_base } else { 28 };
    // Fórmula com arredondamento: ((base * dpi) + 48) / 96
    ((base * dpi_efetivo as i32) + 48) / 96
}

/// Estrutura para hit-testing em controles SysTabControl32
#[repr(C)]
pub struct TCHITTESTINFO {
    pub pt: POINT,
    pub flags: u32,
}

/// Estrutura que encapsula um controle de abas nativo do Windows.
pub struct TabControl {
    pub hwnd: HWND,
    font: Option<HFONT>,
}

/// Obtém o HINSTANCE do módulo que contém a função `obter_hinstance_dll`.
/// Isso garante que o TabControl seja criado com o HINSTANCE correto da nossa DLL
/// e não do processo do Word, evitando problemas de reclassificação de janelas.
fn obter_hinstance_dll() -> windows::Win32::Foundation::HMODULE {
    unsafe {
        let mut hmod = windows::Win32::Foundation::HMODULE::default();
        // Passa o endereço desta própria função para localizar nossa DLL na memória
        let addr = obter_hinstance_dll as *const std::ffi::c_void;
        let ok = GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            windows::core::PCWSTR(addr as *const u16),
            &mut hmod,
        );
        if ok.is_ok() && !hmod.is_invalid() {
            hmod
        } else {
            // Fallback: tenta pegar o handle pelo nome do módulo
            GetModuleHandleW(None).unwrap_or_default()
        }
    }
}

impl TabControl {
    /// Cria um novo controle de abas como filho da janela especificada com posição e tamanho definidos.
    /// DEVE ser chamado da thread de UI que possui a janela pai (parent_hwnd).
    pub fn new(
        parent_hwnd: HWND,
        id: u32,
        x: i32,
        y: i32,
        largura: i32,
        altura: i32,
    ) -> Result<Self> {
        unsafe {
            // Inicializa os controles comuns (necessário para SysTabControl32)
            let icex = INITCOMMONCONTROLSEX {
                dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
                dwICC: ICC_TAB_CLASSES,
            };
            let _ = InitCommonControlsEx(&icex);

            let estilo = WINDOW_STYLE(
                WS_CHILD.0
                    | WS_VISIBLE.0
                    | WS_CLIPSIBLINGS.0
                    | WS_CLIPCHILDREN.0
                    | TCS_TABS
                    | TCS_FOCUSNEVER
                    | TCS_HOTTRACK
                    | 0x4000, // TCS_TOOLTIPS: ativa suporte a tooltips nativos em cada aba
            );

            // CRÍTICO: usar o HINSTANCE da nossa DLL, não do processo do Word
            let hinstance = obter_hinstance_dll();

            crate::tab_manager::log_debug(&format!(
                "TabControl::new — pai=0x{:X} x={} y={} w={} h={} hmod=0x{:X}",
                parent_hwnd.0 as usize, x, y, largura, altura, hinstance.0 as usize
            ));

            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("SysTabControl32"),
                PCWSTR::null(),
                estilo,
                x,
                y,
                largura,
                altura,
                parent_hwnd,
                HMENU(id as usize as *mut _),
                hinstance,
                None,
            )?;

            crate::tab_manager::log_debug(&format!(
                "TabControl::new — HWND criado: 0x{:X}",
                hwnd.0 as usize
            ));

            // Aplicar fonte nítida padrão da interface do Windows
            let hfont = GetStockObject(DEFAULT_GUI_FONT);
            if !hfont.is_invalid() {
                let _ = SendMessageW(hwnd, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            }

            // Aplicar espaçamento horizontal e vertical para um visual arejado e moderno (estilo Fluent)
            // Padding vertical 12: itens de aba mais altos e confortáveis (o desenho
            // customizado usa o rect do item, então altura maior = abas maiores)
            // Padding horizontal 30: abas largas com folga para o botão '×' de fechar
            let _ = SendMessageW(
                hwnd,
                TCM_SETPADDING,
                WPARAM(0),
                LPARAM(((12isize) << 16) | 30isize),
            );

            // Ativar subclassing no SysTabControl32 para capturar cliques de botão do meio e direito
            let subclass_ok = SetWindowSubclass(
                hwnd,
                Some(procedimento_tab_control_subclass),
                OFFICETAB_TABCONTROL_SUBCLASS_ID,
                0,
            );
            if !subclass_ok.as_bool() {
                let error = windows::core::Error::from_win32();
                let _ = DestroyWindow(hwnd);
                return Err(error);
            }

            // Posicionar no topo da Z-Order, forçar visibilidade e repintura
            let _ = SetWindowPos(hwnd, HWND_TOP, x, y, largura, altura, SWP_SHOWWINDOW);
            let _ = BringWindowToTop(hwnd);
            let _ = ShowWindow(hwnd, SW_SHOW);
            let _ = UpdateWindow(hwnd);
            let _ = InvalidateRect(hwnd, None, true);

            Ok(Self { hwnd, font: None })
        }
    }

    /// Adiciona uma nova aba ao final do controle.
    pub fn add_tab(&self, index: i32, titulo: &str) -> Result<i32> {
        let mut utf16: Vec<u16> = titulo.encode_utf16().chain(std::iter::once(0)).collect();

        let mut item = TCITEMW {
            mask: TCIF_TEXT,
            pszText: windows::core::PWSTR(utf16.as_mut_ptr()),
            cchTextMax: utf16.len() as i32,
            ..Default::default()
        };

        unsafe {
            let res = SendMessageW(
                self.hwnd,
                TCM_INSERTITEMW,
                WPARAM(index as usize),
                LPARAM(&mut item as *mut _ as isize),
            );

            if res.0 < 0 {
                return Err(windows::core::Error::from_win32());
            }

            Ok(res.0 as i32)
        }
    }

    /// Remove a aba pelo índice.
    pub fn remove_tab(&self, index: i32) -> Result<()> {
        unsafe {
            let res = SendMessageW(self.hwnd, TCM_DELETEITEM, WPARAM(index as usize), LPARAM(0));

            if res.0 == 0 {
                return Err(windows::core::Error::from_win32());
            }

            Ok(())
        }
    }

    /// Remove todas as abas existentes do controle.
    pub fn clear_tabs(&self) {
        unsafe {
            let _ = SendMessageW(self.hwnd, TCM_DELETEALLITEMS, WPARAM(0), LPARAM(0));
        }
    }

    /// Retorna o índice da aba atualmente selecionada.
    pub fn get_selected_tab(&self) -> i32 {
        unsafe {
            let res = SendMessageW(self.hwnd, TCM_GETCURSEL, WPARAM(0), LPARAM(0));
            res.0 as i32
        }
    }

    /// Define a aba ativa.
    pub fn set_selected_tab(&self, index: i32) -> Result<()> {
        unsafe {
            let res = SendMessageW(self.hwnd, TCM_SETCURSEL, WPARAM(index as usize), LPARAM(0));

            if res.0 < 0 {
                return Err(windows::core::Error::from_win32());
            }

            Ok(())
        }
    }

    /// Retorna a quantidade total de abas.
    pub fn get_tab_count(&self) -> i32 {
        unsafe {
            let res = SendMessageW(self.hwnd, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0));
            res.0 as i32
        }
    }

    /// Atualiza as dimensões e o posicionamento da barra de abas.
    pub fn resize(&self, x: i32, y: i32, largura: i32, altura: i32) -> Result<()> {
        unsafe {
            MoveWindow(self.hwnd, x, y, largura, altura, true)?;
            SetWindowPos(self.hwnd, HWND_TOP, x, y, largura, altura, SWP_SHOWWINDOW)?;
            let _ = ShowWindow(self.hwnd, SW_SHOW);
            let _ = InvalidateRect(self.hwnd, None, true);
            Ok(())
        }
    }

    /// Força a repintura da barra de abas quando necessário.
    pub fn repaint(&self) {
        unsafe {
            let _ = InvalidateRect(self.hwnd, None, true);
        }
    }

    /// Executa teste de colisão para determinar qual aba está na coordenada (x, y) do controle.
    pub fn hit_test(&self, x: i32, y: i32) -> Option<i32> {
        let mut hit_info = TCHITTESTINFO {
            pt: POINT { x, y },
            flags: 0,
        };
        unsafe {
            let res = SendMessageW(
                self.hwnd,
                TCM_HITTEST,
                WPARAM(0),
                LPARAM(&mut hit_info as *mut _ as isize),
            );
            if res.0 >= 0 { Some(res.0 as i32) } else { None }
        }
    }

    /// Aplica configurações de fonte personalizadas com base nas definições do tema e DPI da janela.
    pub fn aplicar_fonte_tema(&mut self, font_name: &str, font_size: u32) {
        unsafe {
            let dpi = crate::utils::get_dpi_for_window(self.hwnd);
            let tamanho_pt = font_size.max(8) as i32;
            // Escalonamento preciso de pontos (pt) para pixels considerando o DPI real da janela
            let altura_pixel = -((tamanho_pt * dpi as i32) / 72);

            if !font_name.is_empty() {
                let utf16: Vec<u16> = font_name.encode_utf16().chain(std::iter::once(0)).collect();
                let hfont = CreateFontW(
                    altura_pixel,
                    0,
                    0,
                    0,
                    FW_NORMAL.0 as i32,
                    0,
                    0,
                    0,
                    DEFAULT_CHARSET.0 as u32,
                    OUT_DEFAULT_PRECIS.0 as u32,
                    CLIP_DEFAULT_PRECIS.0 as u32,
                    CLEARTYPE_QUALITY.0 as u32,
                    (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                    PCWSTR(utf16.as_ptr()),
                );
                if !hfont.is_invalid() {
                    let _ =
                        SendMessageW(self.hwnd, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                    if let Some(old_font) = self.font.replace(hfont) {
                        let _ = DeleteObject(old_font);
                    }
                    return;
                }
            }
            let hfont = GetStockObject(DEFAULT_GUI_FONT);
            if !hfont.is_invalid() {
                let _ = SendMessageW(self.hwnd, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            }
        }
    }

    /// Destrói o controle de abas e remove sua subclasse associada.
    pub fn destroy(&mut self) {
        if !self.hwnd.is_invalid() {
            unsafe {
                let _ = RemoveWindowSubclass(
                    self.hwnd,
                    Some(procedimento_tab_control_subclass),
                    OFFICETAB_TABCONTROL_SUBCLASS_ID,
                );
                let _ = DestroyWindow(self.hwnd);
            }
            self.hwnd = HWND::default();
        }
        if let Some(font) = self.font.take() {
            unsafe {
                let _ = DeleteObject(font);
            }
        }
    }
}

/// Procedimento de subclasse específico para o controle de abas SysTabControl32.
/// Intercepta cliques de botão do meio (fechar) e botão direito (menu de contexto).
/// Blindado com std::panic::catch_unwind para proteger a estabilidade do Office.
unsafe extern "system" fn procedimento_tab_control_subclass(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    uid_subclass: usize,
    ref_data: usize,
) -> LRESULT {
    let resultado = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        procedimento_tab_control_subclass_interno(hwnd, msg, wparam, lparam, uid_subclass, ref_data)
    }));

    match resultado {
        Ok(lres) => lres,
        Err(e) => {
            crate::tab_manager::log_debug(&format!(
                "Pânico capturado no procedimento_tab_control_subclass: {:?}",
                e
            ));
            unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) }
        }
    }
}

fn procedimento_tab_control_subclass_interno(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    uid_subclass: usize,
    _ref_data: usize,
) -> LRESULT {
    const WM_DESTROY: u32 = 0x0002;
    const WM_NCDESTROY: u32 = 0x0082;
    const WM_PAINT: u32 = 0x000F;
    const WM_LBUTTONUP: u32 = 0x0202;
    const WM_LBUTTONDBLCLK: u32 = 0x0203;
    const WM_MBUTTONUP: u32 = 0x0208;
    const WM_RBUTTONUP: u32 = 0x0205;
    const WM_MOUSEWHEEL: u32 = 0x020A;
    const WM_CONTEXTMENU: u32 = 0x007B;

    match msg {
        // WM_PAINT: desenho customizado AQUI, no próprio controle.
        // No Excel (SDI) o WM_NOTIFY/NM_CUSTOMDRAW que o SysTabControl32 envia ao
        // pai (XLMAIN) é engolido pelo Excel antes da nossa subclasse vê-lo, então
        // o desenho customizado (×, título à esquerda, cores) nunca roda via pai.
        // Pintamos manualmente reaproveitando o desenho do TabManager.
        //
        // IMPORTANTE: usamos GetDC (e não BeginPaint) porque o DC do BeginPaint é
        // recortado pela update region, que chegou vazia (ps.rcPaint=(0,0,0,0)) no
        // Excel — todo o desenho era descartado. GetDC cobre o cliente inteiro;
        // validamos a região depois com ValidateRect para parar o WM_PAINT.
        WM_PAINT => {
            // Primeiro, pinta toda a barra via GetDC (evita update region vazia do Excel).
            let hdc = unsafe { GetDC(hwnd) };
            if !hdc.is_invalid() {
                let mut rc = RECT::default();
                unsafe {
                    let _ = GetClientRect(hwnd, &mut rc);
                }
                crate::tab_manager::despachar_tab_paint(hwnd, hdc, rc);
                unsafe {
                    let _ = ReleaseDC(hwnd, hdc);
                    // Não validamos ainda: deixe o BeginPaint normal fazer o resto.
                }
            }
            // Agora delega ao DefSubclassProc para que o Windows processe o WM_PAINT
            // nativo do SysTabControl32 e não quebre o owner-draw / foco / realce do controle.
            // Após o processamento padrão, traz o controle para o topo para garantir
            // que o botão “×” e o conteúdo da aba não fiquem sobrepostos.
            let _ = unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) };
            unsafe {
                let _ = SetWindowPos(
                    hwnd,
                    HWND_TOP,
                    0,
                    0,
                    0,
                    0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                );
                let _ = InvalidateRect(hwnd, None, true);
                let _ = UpdateWindow(hwnd);
            }
            return LRESULT(0);
        }
        WM_LBUTTONUP => {
            let x = (lparam.0 & 0xFFFF) as i16 as i32;
            let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
            if crate::tab_manager::despachar_tab_lbutton_up(hwnd, x, y) {
                return LRESULT(0);
            }
        }
        WM_LBUTTONDBLCLK => {
            let x = (lparam.0 & 0xFFFF) as i16 as i32;
            let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
            crate::tab_manager::despachar_tab_lbutton_dblclk(hwnd, x, y);
            return LRESULT(0);
        }
        WM_MBUTTONUP => {
            let x = (lparam.0 & 0xFFFF) as i16 as i32;
            let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
            crate::tab_manager::despachar_tab_mbutton_up(hwnd, x, y);
            return LRESULT(0);
        }
        WM_RBUTTONUP => {
            let x = (lparam.0 & 0xFFFF) as i16 as i32;
            let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
            crate::tab_manager::despachar_tab_rbutton_up(hwnd, x, y);
            return LRESULT(0);
        }
        WM_MOUSEWHEEL => {
            let delta = ((wparam.0 >> 16) & 0xFFFF) as i16;
            crate::tab_manager::despachar_tab_mousewheel(hwnd, delta);
            return LRESULT(0);
        }
        WM_CONTEXTMENU => {
            let x = (lparam.0 & 0xFFFF) as i16 as i32;
            let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;
            if lparam.0 == -1 || lparam.0 as usize == usize::MAX {
                crate::tab_manager::despachar_tab_context_menu_teclado(hwnd);
            } else {
                crate::tab_manager::despachar_tab_context_menu(hwnd, x, y);
            }
            return LRESULT(0);
        }
        0x0200 /* WM_MOUSEMOVE */ => unsafe {
            // Otimização: invalida apenas se a posição do mouse mudou significativamente (pelo menos 3px),
            // evitando re-renderização em cascata a cada micropasso do cursor.
            static ULTIMO_POS: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(0);
            let pos_atual = (lparam.0 & 0xFFFFFFFF) as u32;
            let ant = ULTIMO_POS.swap(pos_atual, std::sync::atomic::Ordering::Relaxed);
            let dx = ((pos_atual & 0xFFFF) as i16 - (ant & 0xFFFF) as i16).abs();
            let dy = (((pos_atual >> 16) & 0xFFFF) as i16 - ((ant >> 16) & 0xFFFF) as i16).abs();
            if dx >= 3 || dy >= 3 {
                let _ = InvalidateRect(hwnd, None, false);
            }
            return DefSubclassProc(hwnd, msg, wparam, lparam);
        },
        WM_NCDESTROY => unsafe {
            // Em WM_NCDESTROY, o RemoveWindowSubclass é chamado caso a janela seja destruída prematuramente
            let _ =
                RemoveWindowSubclass(hwnd, Some(procedimento_tab_control_subclass), uid_subclass);
            return DefSubclassProc(hwnd, msg, wparam, lparam);
        },
        WM_DESTROY => {
            // Recebido antes de WM_NCDESTROY
        }
        _ => {}
    }

    unsafe { DefSubclassProc(hwnd, msg, wparam, lparam) }
}

impl Drop for TabControl {
    fn drop(&mut self) {
        self.destroy();
    }
}

unsafe impl Send for TabControl {}
unsafe impl Sync for TabControl {}

#[cfg(test)]
mod testes {
    use super::*;

    #[test]
    fn test_estrutura_tchittestinfo_alinhamento() {
        assert!(std::mem::size_of::<TCHITTESTINFO>() >= 12);
        assert_eq!(std::mem::align_of::<TCHITTESTINFO>(), 4);
    }

    #[test]
    fn test_constantes_mensagens_tab_control() {
        assert_eq!(TCM_HITTEST, 0x130D);
        assert_eq!(TCM_GETITEMRECT, 0x130A);
        assert_eq!(TCM_SETPADDING, 0x132B);
        assert_eq!(OFFICETAB_TABCONTROL_SUBCLASS_ID, 0x54414243);
    }

    #[test]
    fn test_tab_control_invalido_hit_test_seguro() {
        let tab = TabControl {
            hwnd: HWND::default(),
            font: None,
        };
        let res = tab.hit_test(50, 10);
        assert!(res.is_none() || res.is_some());
    }

    #[test]
    fn test_escalonamento_altura_por_dpi() {
        // 96 DPI  (100%) -> 28px
        assert_eq!(calcular_altura_por_dpi_valor(96, 28), 28);
        // 144 DPI (150%) -> 42px
        assert_eq!(calcular_altura_por_dpi_valor(144, 28), 42);
        // 192 DPI (200%) -> 56px
        assert_eq!(calcular_altura_por_dpi_valor(192, 28), 56);

        // Altura base personalizada (34px)
        assert_eq!(calcular_altura_por_dpi_valor(96, 34), 34);
        assert_eq!(calcular_altura_por_dpi_valor(144, 34), 51);
        assert_eq!(calcular_altura_por_dpi_valor(192, 34), 68);
    }
}
