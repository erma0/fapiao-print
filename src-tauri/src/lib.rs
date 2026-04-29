use tauri::command;

// v1.5.0: upgraded image 0.24→0.25 (webp), printpdf 0.7→0.9 (new API)

mod pdf_engine;
use pdf_engine::{PrinterInfo, FileData, RenderedPage, ComGuard, LayoutRenderRequest};

// =====================================================
// Tauri Commands
// =====================================================

/// Read files from given paths (for drag-and-drop and dialog plugin)
#[command]
fn open_invoice_files(paths: Vec<String>) -> Result<Vec<FileData>, String> {
    pdf_engine::read_invoice_files(paths)
}

/// List available printers
#[command]
fn get_printers() -> Result<Vec<PrinterInfo>, String> {
    pdf_engine::list_printers()
}

/// Render PDF pages to images using Windows native API
#[command]
fn render_pdf_pages(pdf_path: String, dpi: Option<u32>) -> Result<Vec<RenderedPage>, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::render_pdf_pages(&pdf_path, dpi.unwrap_or(pdf_engine::RENDER_DPI))
}

/// Open a file with the default application (for auto-opening saved PDFs)
#[command]
fn open_file(path: String) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        shell_execute("open", &path)?;
    }
    Ok(())
}

/// Open a URL in the default browser
#[command]
fn open_url(url: String) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        shell_execute("open", &url)?;
    }
    Ok(())
}

/// OCR an image from base64 data URL, return recognized text
#[command]
fn ocr_image(data_url: String) -> Result<String, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::ocr_image_from_data(&data_url)
}

/// Get backend configuration (for runtime DPI validation)
#[command]
fn get_config() -> Result<serde_json::Value, String> {
    Ok(serde_json::json!({
        "renderDpi": pdf_engine::RENDER_DPI,
    }))
}

/// Get system temp directory path (for print output)
#[command]
fn get_temp_dir() -> Result<String, String> {
    let temp = std::env::temp_dir();
    // Ensure the temp dir exists
    let _ = std::fs::create_dir_all(&temp);
    Ok(temp.to_string_lossy().to_string())
}

// =====================================================
// New Commands: Trim Image & Layout-based PDF Generation
// =====================================================

/// Trim white edges from an image (base64 data URL → trimmed base64 data URL)
#[command]
fn trim_image(data_url: String) -> Result<String, String> {
    use base64::Engine;
    use std::io::Cursor;

    let img = pdf_engine::decode_base64_image(&data_url)
        .map_err(|e| format!("解码失败: {}", e))?;
    let trimmed = pdf_engine::trim_white_edges(&img, 245);

    // Encode back to PNG base64
    let mut buf = Cursor::new(Vec::new());
    trimmed.write_to(&mut buf, image::ImageFormat::Png)
        .map_err(|e| format!("PNG编码失败: {}", e))?;

    let b64 = base64::engine::general_purpose::STANDARD.encode(buf.into_inner());
    Ok(format!("data:image/png;base64,{}", b64))
}

/// Generate PDF from layout request (files + pages + settings).
/// Replaces JS `renderPageToCanvas` + `generate_pdf_from_pages`.
#[command]
fn generate_pdf_from_layout(
    request: LayoutRenderRequest,
    output_path: String,
    direct_print: Option<bool>,
    printer_name: Option<String>,
) -> Result<pdf_engine::PdfResult, String> {
    use std::sync::atomic::Ordering;

    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    let output = std::path::Path::new(&output_path);
    pdf_engine::generate_pdf_from_layout(&request, output)
        .map_err(|e| format!("PDF生成失败: {}", e))?;

    let is_direct = direct_print.unwrap_or(false);
    let printer = printer_name.as_deref();

    #[cfg(target_os = "windows")]
    {
        if is_direct {
            direct_print_pdf(output, printer)?;
        } else {
            shell_execute("open", &output.to_string_lossy())?;
        }
    }

    let msg = if is_direct {
        if let Some(name) = printer {
            format!("已发送到打印机「{}」", name)
        } else {
            "已发送到默认打印机".to_string()
        }
    } else {
        "已打开PDF预览，请在阅读器中确认打印".to_string()
    };

    Ok(pdf_engine::PdfResult {
        success: true,
        message: msg,
        pdf_path: Some(output.to_string_lossy().to_string()),
    })
}

// =====================================================
// Helpers
// =====================================================

/// Call Windows ShellExecuteW — no cmd.exe, no terminal window
#[cfg(target_os = "windows")]
fn shell_execute(verb: &str, file: &str) -> Result<(), String> {
    use windows::core::HSTRING;
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

    let _com = ComGuard::init();
    unsafe {
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

/// Direct-print a PDF file to a specific printer (or system default) using winprint
/// This bypasses PDF reader software entirely — sends directly to Windows Print Spooler
#[cfg(target_os = "windows")]
fn direct_print_pdf(pdf_path: &std::path::Path, printer_name: Option<&str>) -> Result<(), String> {
    use winprint::printer::{FilePrinter, PrinterDevice, WinPdfPrinter};
    use winprint::ticket::PrintTicket;

    // Find the target printer
    let devices = PrinterDevice::all()
        .map_err(|e| format!("获取打印机列表失败: {}", e))?;

    let device = if let Some(name) = printer_name {
        // User selected a specific printer
        devices.into_iter()
            .find(|d| d.name().eq_ignore_ascii_case(name))
            .ok_or_else(|| format!("找不到打印机「{}」", name))?
    } else {
        // No specific printer selected: find the system default printer
        let default_name = pdf_engine::get_default_printer_name();
        if let Some(ref dn) = default_name {
            devices.into_iter()
                .find(|d| d.name().eq_ignore_ascii_case(dn))
                .ok_or_else(|| format!("找不到默认打印机「{}」", dn))?
        } else {
            // Fallback: use first available printer
            devices.into_iter()
                .next()
                .ok_or_else(|| "系统中没有可用的打印机".to_string())?
        }
    };

    let printer = WinPdfPrinter::new(device);
    let ticket = PrintTicket::default();

    printer.print(pdf_path, ticket)
        .map_err(|e| format!("打印失败: {:?}", e))?;

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
            // Handle window events: close + drag-and-drop
            {
                use tauri::Manager;
                let window = app.get_webview_window("main").unwrap();
                let win = window.clone();

                window.on_window_event(move |event| {
                    match event {
                        // --- Close event: graceful shutdown ---
                        // Only set SHUTTING_DOWN flag to reject new OCR/COM requests.
                        // Let Tauri handle the rest — it will destroy the window, clean up
                        // WebView2 child processes, and exit the process normally.
                        // Previous code used taskkill/TerminateProcess to force-kill the
                        // process tree, which interrupted Tauri's cleanup and left WebView2
                        // child processes as orphans — that was the actual cause of lingering
                        // processes after exit.
                        tauri::WindowEvent::CloseRequested { api, .. } => {
                            use std::sync::atomic::Ordering;
                            // If already shutting down, allow the close to proceed.
                            if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
                                return;
                            }
                            // Prevent immediate close so in-flight operations can abort.
                            api.prevent_close();
                            pdf_engine::SHUTTING_DOWN.store(true, Ordering::SeqCst);
                            // Notify frontend to clear queues and stop accepting new work.
                            let _ = win.eval("if(window._tauriCleanup)window._tauriCleanup();");
                            // Brief delay for OCR/PDF threads to notice the flag,
                            // then explicitly close the window.
                            let win2 = win.clone();
                            std::thread::spawn(move || {
                                std::thread::sleep(std::time::Duration::from_millis(400));
                                let _ = win2.close();
                            });
                        }
                        // --- Drag-and-drop file handling ---
                        tauri::WindowEvent::DragDrop(drop_event) => {
                            if let tauri::DragDropEvent::Drop { paths, .. } = drop_event {
                                let valid: Vec<String> = paths.iter()
                                    .filter_map(|p| {
                                        let valid_ext = p.extension()
                                            .and_then(|e| e.to_str())
                                            .map(|e| ["pdf", "jpg", "jpeg", "png", "bmp", "webp", "tiff", "tif", "ofd"].contains(&e.to_lowercase().as_str()))
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
                        }
                        _ => {}
                    }
                });
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            open_invoice_files,
            get_printers,
            render_pdf_pages,
            open_url,
            open_file,
            ocr_image,
            get_config,
            get_temp_dir,
            trim_image,
            generate_pdf_from_layout,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
