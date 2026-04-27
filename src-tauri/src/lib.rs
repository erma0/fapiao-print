use tauri::command;

mod pdf_engine;

use pdf_engine::{PdfRequest, PdfResult, PrinterInfo, FileData};

// =====================================================
// Tauri Commands
// =====================================================

/// Read files from given paths (for drag-and-drop and dialog plugin)
#[command]
fn open_invoice_files(paths: Vec<String>) -> Result<Vec<FileData>, String> {
    pdf_engine::read_invoice_files(paths)
}

/// Render pages to PDF only (no print, no open)
#[command]
fn generate_pdf(request: PdfRequest) -> Result<PdfResult, String> {
    let output_path = std::env::temp_dir().join("fapiao_print_output.pdf");
    pdf_engine::generate_pdf_from_pages(&request, &output_path)?;

    Ok(PdfResult {
        success: true,
        message: "PDF生成成功".to_string(),
        pdf_path: Some(output_path.to_string_lossy().to_string()),
    })
}

/// Generate PDF then open system print dialog (or direct print)
#[command]
fn generate_and_print(request: PdfRequest, direct_print: Option<bool>) -> Result<PdfResult, String> {
    let output_path = std::env::temp_dir().join("fapiao_print_output.pdf");
    pdf_engine::generate_pdf_from_pages(&request, &output_path)?;

    let is_direct = direct_print.unwrap_or(false);

    #[cfg(target_os = "windows")]
    {
        // dialog 模式：用 "open" 打开 PDF，用户可在阅读器中预览后手动打印
        // direct 模式：用 "print" 直接发送到打印机
        let verb = if is_direct { "print" } else { "open" };
        shell_execute(verb, &output_path.to_string_lossy())?;
    }

    Ok(PdfResult {
        success: true,
        message: if is_direct { "已发送到打印机".to_string() } else { "已打开PDF预览，请在阅读器中确认打印".to_string() },
        pdf_path: Some(output_path.to_string_lossy().to_string()),
    })
}

/// Save PDF to user-chosen path (or Desktop) and optionally open it
#[command]
fn save_pdf(request: PdfRequest, save_path: Option<String>, auto_open: Option<bool>) -> Result<PdfResult, String> {
    let output_path = if let Some(path) = save_path {
        std::path::PathBuf::from(path)
    } else {
        let desktop = std::env::var("USERPROFILE")
            .map(|p| std::path::PathBuf::from(p).join("Desktop"))
            .unwrap_or_else(|_| std::path::PathBuf::from("."));
        let timestamp = chrono_free_filename();
        desktop.join(format!("发票打印_{}.pdf", timestamp))
    };

    // Ensure parent dir exists
    if let Some(parent) = output_path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }

    pdf_engine::generate_pdf_from_pages(&request, &output_path)?;

    // Open the saved PDF with default viewer only if autoOpen is true
    #[cfg(target_os = "windows")]
    {
        if auto_open.unwrap_or(true) {
            let _ = shell_execute("open", &output_path.to_string_lossy());
        }
    }

    Ok(PdfResult {
        success: true,
        message: "PDF已保存".to_string(),
        pdf_path: Some(output_path.to_string_lossy().to_string()),
    })
}

/// List available printers
#[command]
fn get_printers() -> Result<Vec<PrinterInfo>, String> {
    pdf_engine::list_printers()
}

/// Open a file or folder with default app (no terminal window)
#[command]
fn open_path(path: String) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        shell_execute("open", &path)?;
    }
    Ok(())
}

// =====================================================
// Helpers
// =====================================================

fn chrono_free_filename() -> String {
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    format!("{}", now)
}

/// Call Windows ShellExecuteW — no cmd.exe, no terminal window
#[cfg(target_os = "windows")]
fn shell_execute(verb: &str, file: &str) -> Result<(), String> {
    use windows::core::HSTRING;
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;
    use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};

    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        let v: HSTRING = verb.into();
        let f: HSTRING = file.into();
        let ret = ShellExecuteW(
            None,
            &v,
            &f,
            windows::core::PCWSTR::null(),
            windows::core::PCWSTR::null(),
            SW_SHOWNORMAL,
        );
        if ret.0 as isize <= 32 {
            return Err(format!("ShellExecute 失败，错误码: {}", ret.0 as isize));
        }
    }
    Ok(())
}

// =====================================================
// App Entry
// =====================================================

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_shell::init())
        .setup(|app| {
            if cfg!(debug_assertions) {
                app.handle().plugin(
                    tauri_plugin_log::Builder::default()
                        .level(log::LevelFilter::Info)
                        .build(),
                )?;
            }
            #[cfg(debug_assertions)]
            {
                use tauri::Manager;
                let window = app.get_webview_window("main").unwrap();
                window.open_devtools();
            }
            // Listen for file drag-and-drop events
            {
                use tauri::Manager;
                let window = app.get_webview_window("main").unwrap();
                let win = window.clone();
                window.on_window_event(move |event| {
                    if let tauri::WindowEvent::DragDrop(drop_event) = event {
                        match drop_event {
                            tauri::DragDropEvent::Drop { paths, .. } => {
                                let valid: Vec<String> = paths.iter()
                                    .filter_map(|p| {
                                        let valid_ext = p.extension()
                                            .and_then(|e| e.to_str())
                                            .map(|e| ["pdf", "jpg", "jpeg", "png", "bmp", "webp", "tiff", "tif"].contains(&e.to_lowercase().as_str()))
                                            .unwrap_or(false);
                                        if valid_ext { Some(p.to_string_lossy().to_string()) } else { None }
                                    })
                                    .collect();
                                if !valid.is_empty() {
                                    let json = serde_json::to_string(&valid).unwrap_or_default();
                                    let js = format!("if(window._tauriFileDrop)window._tauriFileDrop({})", json);
                                    let _ = win.eval(&js);
                                }
                            }
                            _ => {}
                        }
                    }
                });
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            open_invoice_files,
            generate_pdf,
            generate_and_print,
            save_pdf,
            get_printers,
            open_path,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
