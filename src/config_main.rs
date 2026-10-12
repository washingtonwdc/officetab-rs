//! Pequeno editor Win32 para as preferências persistidas do OfficeTab Rs.

use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Mutex;

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, GetSaveFileNameW, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_OVERWRITEPROMPT,
    OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, BS_PUSHBUTTON, CB_ADDSTRING, CB_GETCURSEL,
    CB_RESETCONTENT, CB_SETCURSEL, CBS_DROPDOWNLIST, CW_USEDEFAULT, CreateWindowExW,
    DefWindowProcW, DispatchMessageW, ES_NUMBER, GetDlgItem, GetMessageW, GetWindowTextLengthW,
    GetWindowTextW, HMENU, MB_ICONERROR, MB_ICONINFORMATION, MB_OK, MSG, MessageBoxW,
    PostQuitMessage, RegisterClassW, SendMessageW, SetWindowTextW, TranslateMessage, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WNDCLASSW, WS_BORDER, WS_CHILD, WS_OVERLAPPEDWINDOW, WS_TABSTOP,
    WS_VISIBLE,
};
use windows::core::{PCWSTR, w};

use officetab_rs::persistence::{
    OfficeTabConfig, export_config_to_file, import_config_from_file, load_effective_config,
    save_effective_config,
};
use officetab_rs::theme::Theme;

const CLASS_NAME: windows::core::PCWSTR = w!("OfficeTabRsConfigWindow");
const ID_THEME: i32 = 1001;
const ID_HEIGHT: i32 = 1002;
const ID_ADAPTIVE: i32 = 1003;
const ID_SAVE: i32 = 1004;
const ID_IMPORT: i32 = 1005;
const ID_EXPORT: i32 = 1006;

static CONFIG: Mutex<Option<OfficeTabConfig>> = Mutex::new(None);

fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

struct ControlSpec {
    class: PCWSTR,
    text: PCWSTR,
    style: u32,
    id: i32,
    bounds: [i32; 4],
}

unsafe fn add_control(parent: HWND, control: ControlSpec) -> HWND {
    unsafe {
        CreateWindowExW(
            Default::default(),
            control.class,
            control.text,
            windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(control.style),
            control.bounds[0],
            control.bounds[1],
            control.bounds[2],
            control.bounds[3],
            parent,
            HMENU(control.id as usize as *mut _),
            None,
            None,
        )
        .unwrap_or_default()
    }
}

unsafe fn populate_form(hwnd: HWND, config: &OfficeTabConfig) {
    unsafe {
        let combo = GetDlgItem(hwnd, ID_THEME).unwrap_or_default();
        let _ = SendMessageW(combo, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let presets = Theme::presets();
        for preset in &presets {
            let label = wide(&preset.name);
            let _ = SendMessageW(
                combo,
                CB_ADDSTRING,
                WPARAM(0),
                LPARAM(label.as_ptr() as isize),
            );
        }
        let selected = presets
            .iter()
            .position(|theme| theme.name.eq_ignore_ascii_case(&config.theme.name))
            .unwrap_or(0);
        let _ = SendMessageW(combo, CB_SETCURSEL, WPARAM(selected), LPARAM(0));

        let height = wide(&config.general.tab_bar_height.to_string());
        let _ = SetWindowTextW(
            GetDlgItem(hwnd, ID_HEIGHT).unwrap_or_default(),
            PCWSTR(height.as_ptr()),
        );
        let checked = if config.general.adaptive_theme_by_app {
            1
        } else {
            0
        };
        let _ = SendMessageW(
            GetDlgItem(hwnd, ID_ADAPTIVE).unwrap_or_default(),
            BM_SETCHECK,
            WPARAM(checked),
            LPARAM(0),
        );
    }
}

unsafe fn choose_json_file(hwnd: HWND, save: bool) -> Option<std::path::PathBuf> {
    let mut file = vec![0u16; 32768];
    let filter = wide("JSON configuration (*.json)\0*.json\0All files\0*.*\0");
    let mut dialog = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: windows::core::PWSTR(file.as_mut_ptr()),
        nMaxFile: file.len() as u32,
        Flags: if save {
            OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY
        } else {
            OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY
        },
        ..Default::default()
    };
    let accepted = unsafe {
        if save {
            GetSaveFileNameW(&mut dialog).as_bool()
        } else {
            GetOpenFileNameW(&mut dialog).as_bool()
        }
    };
    if !accepted {
        return None;
    }
    let length = file
        .iter()
        .position(|unit| *unit == 0)
        .unwrap_or(file.len());
    Some(std::path::PathBuf::from(String::from_utf16_lossy(
        &file[..length],
    )))
}

unsafe fn show_message(hwnd: HWND, text: &str, error: bool) {
    let title = wide("OfficeTab Rs Config");
    let message = wide(text);
    unsafe {
        let style = if error {
            MB_ICONERROR
        } else {
            MB_ICONINFORMATION
        };
        let _ = MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | style,
        );
    }
}

unsafe fn save_form(hwnd: HWND) -> Result<(), String> {
    let mut config = CONFIG
        .lock()
        .map_err(|_| "A configuração está bloqueada".to_string())?;
    let config = config
        .as_mut()
        .ok_or_else(|| "A configuração ainda não foi carregada".to_string())?;

    let height_hwnd = unsafe { GetDlgItem(hwnd, ID_HEIGHT).unwrap_or_default() };
    let length = unsafe { GetWindowTextLengthW(height_hwnd) }.max(0) as usize;
    let mut value = vec![0u16; length + 1];
    unsafe { GetWindowTextW(height_hwnd, &mut value) };
    let height = String::from_utf16_lossy(&value[..length])
        .parse::<i32>()
        .map_err(|_| "A altura deve ser um número inteiro".to_string())?;
    config.general.tab_bar_height = height;

    let selected = unsafe {
        SendMessageW(
            GetDlgItem(hwnd, ID_THEME).unwrap_or_default(),
            CB_GETCURSEL,
            WPARAM(0),
            LPARAM(0),
        )
    }
    .0;
    let presets = Theme::presets();
    let theme = presets
        .get(selected.max(0) as usize)
        .cloned()
        .ok_or_else(|| "Selecione um tema".to_string())?;
    config.theme = theme;
    config.general.adaptive_theme_by_app = unsafe {
        SendMessageW(
            GetDlgItem(hwnd, ID_ADAPTIVE).unwrap_or_default(),
            BM_GETCHECK,
            WPARAM(0),
            LPARAM(0),
        )
    }
    .0 == 1;

    save_effective_config(config).map_err(|error| error.to_string())
}

unsafe fn window_proc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let config = load_effective_config();
            if let Ok(mut state) = CONFIG.lock() {
                *state = Some(config.clone());
            }
            unsafe {
                let label_theme = wide("Tema");
                let label_height = wide("Altura da barra (px)");
                let label_adaptive = wide("Usar paleta recomendada por aplicativo Office");
                let save = wide("Salvar");
                let import = wide("Importar JSON...");
                let export = wide("Exportar JSON...");
                let _ = add_control(
                    hwnd,
                    ControlSpec {
                        class: w!("STATIC"),
                        text: PCWSTR(label_theme.as_ptr()),
                        style: WS_CHILD.0 | WS_VISIBLE.0,
                        id: 0,
                        bounds: [18, 18, 220, 22],
                    },
                );
                let _ = add_control(
                    hwnd,
                    ControlSpec {
                        class: w!("COMBOBOX"),
                        text: PCWSTR::null(),
                        style: WS_CHILD.0 | WS_VISIBLE.0 | WS_TABSTOP.0 | CBS_DROPDOWNLIST as u32,
                        id: ID_THEME,
                        bounds: [18, 42, 280, 240],
                    },
                );
                let _ = add_control(
                    hwnd,
                    ControlSpec {
                        class: w!("STATIC"),
                        text: PCWSTR(label_height.as_ptr()),
                        style: WS_CHILD.0 | WS_VISIBLE.0,
                        id: 0,
                        bounds: [18, 92, 220, 22],
                    },
                );
                let _ = add_control(
                    hwnd,
                    ControlSpec {
                        class: w!("EDIT"),
                        text: PCWSTR::null(),
                        style: WS_CHILD.0
                            | WS_VISIBLE.0
                            | WS_TABSTOP.0
                            | WS_BORDER.0
                            | ES_NUMBER as u32,
                        id: ID_HEIGHT,
                        bounds: [18, 116, 90, 26],
                    },
                );
                let _ = add_control(
                    hwnd,
                    ControlSpec {
                        class: w!("BUTTON"),
                        text: PCWSTR(label_adaptive.as_ptr()),
                        style: WS_CHILD.0 | WS_VISIBLE.0 | WS_TABSTOP.0 | BS_AUTOCHECKBOX as u32,
                        id: ID_ADAPTIVE,
                        bounds: [18, 154, 330, 26],
                    },
                );
                let _ = add_control(
                    hwnd,
                    ControlSpec {
                        class: w!("BUTTON"),
                        text: PCWSTR(save.as_ptr()),
                        style: WS_CHILD.0 | WS_VISIBLE.0 | WS_TABSTOP.0 | BS_PUSHBUTTON as u32,
                        id: ID_SAVE,
                        bounds: [18, 204, 100, 30],
                    },
                );
                let _ = add_control(
                    hwnd,
                    ControlSpec {
                        class: w!("BUTTON"),
                        text: PCWSTR(import.as_ptr()),
                        style: WS_CHILD.0 | WS_VISIBLE.0 | WS_TABSTOP.0 | BS_PUSHBUTTON as u32,
                        id: ID_IMPORT,
                        bounds: [130, 204, 120, 30],
                    },
                );
                let _ = add_control(
                    hwnd,
                    ControlSpec {
                        class: w!("BUTTON"),
                        text: PCWSTR(export.as_ptr()),
                        style: WS_CHILD.0 | WS_VISIBLE.0 | WS_TABSTOP.0 | BS_PUSHBUTTON as u32,
                        id: ID_EXPORT,
                        bounds: [260, 204, 130, 30],
                    },
                );
                populate_form(hwnd, &config);
            }
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xFFFF) as i32;
            unsafe {
                match id {
                    ID_SAVE => match save_form(hwnd) {
                        Ok(()) => show_message(hwnd, "Configuração salva.", false),
                        Err(error) => show_message(hwnd, &error, true),
                    },
                    ID_EXPORT => {
                        if let Some(path) = choose_json_file(hwnd, true)
                            && let Ok(config) = CONFIG.lock()
                            && let Some(config) = config.as_ref()
                        {
                            match export_config_to_file(config, &path) {
                                Ok(()) => show_message(hwnd, "Configuração exportada.", false),
                                Err(error) => show_message(hwnd, &error.to_string(), true),
                            }
                        }
                    }
                    ID_IMPORT => {
                        if let Some(path) = choose_json_file(hwnd, false) {
                            match import_config_from_file(&path) {
                                Ok(imported) => {
                                    let result = save_effective_config(&imported);
                                    if let Err(error) = result {
                                        show_message(hwnd, &error.to_string(), true);
                                    } else if let Ok(mut state) = CONFIG.lock() {
                                        *state = Some(imported.clone());
                                        populate_form(hwnd, &imported);
                                        show_message(
                                            hwnd,
                                            "Configuração importada e salva.",
                                            false,
                                        );
                                    }
                                }
                                Err(error) => show_message(hwnd, &error.to_string(), true),
                            }
                        }
                    }
                    _ => {}
                }
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            if let Ok(mut state) = CONFIG.lock() {
                *state = None;
            }
            unsafe { PostQuitMessage(0) };
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match catch_unwind(AssertUnwindSafe(|| unsafe {
        window_proc_inner(hwnd, msg, wparam, lparam)
    })) {
        Ok(result) => result,
        Err(_) => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

fn main() {
    unsafe {
        let Ok(module) = GetModuleHandleW(None) else {
            return;
        };
        let instance = HINSTANCE(module.0);
        let class = WNDCLASSW {
            lpfnWndProc: Some(window_proc),
            hInstance: instance,
            lpszClassName: CLASS_NAME,
            ..Default::default()
        };
        if RegisterClassW(&class) == 0 {
            return;
        }

        let Ok(hwnd) = CreateWindowExW(
            Default::default(),
            CLASS_NAME,
            w!("OfficeTab Rs — Configurações"),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            430,
            300,
            None,
            None,
            instance,
            None,
        ) else {
            return;
        };

        let mut message = MSG::default();
        while GetMessageW(&mut message, HWND::default(), 0, 0).as_bool() {
            let _ = TranslateMessage(&message);
            let _ = DispatchMessageW(&message);
        }
        let _ = hwnd;
    }
}
