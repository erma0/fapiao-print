use tauri::{command, Emitter};

// v1.7.0: ocr-rs (PP-OCRv5 + MNN) replaces WinRT OCR, coordinate-first extraction

mod pdf_engine;
use pdf_engine::{PrinterInfo, FileData, RenderedPage, RenderedOcrPage, ComGuard, LayoutRenderRequest, OcrResult};

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

/// Render PDF pages AND run OCR in one pass — avoids IPC round-trip.
/// Returns preview images + OCR results together.
#[command]
fn render_and_ocr_pdf(pdf_path: String, dpi: Option<u32>) -> Result<Vec<RenderedOcrPage>, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::render_and_ocr_pdf(&pdf_path, dpi.unwrap_or(pdf_engine::RENDER_DPI))
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

/// OCR an image from base64 data URL or file path, return structured result with text + word coordinates.
/// When `filePath` is provided, Rust reads the image directly from disk — skipping the
/// expensive base64 encode→IPC→decode round-trip (saves ~30% data + CPU for large images).
/// Falls back to `dataUrl` when `filePath` is None or file read fails.
#[command]
fn ocr_image(data_url: String, file_path: Option<String>) -> Result<OcrResult, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::ocr_image(&data_url, file_path.as_deref())
}

/// Render a single PDF page and run OCR on it — zero IPC round-trip.
/// The frontend calls this instead of `render_pdf_pages` + `ocr_image` for PDF pages,
/// avoiding the expensive Rust→base64→IPC→frontend→downsample→base64→IPC→Rust cycle.
#[command]
fn ocr_pdf_page(pdf_path: String, page_index: u32, dpi: Option<u32>) -> Result<OcrResult, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::ocr_pdf_page(&pdf_path, page_index, dpi)
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

/// Show the main window — called by frontend after splash screen renders
#[command]
fn show_window(app: tauri::AppHandle) {
    use tauri::Manager;
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
    }
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
/// Emits `pdf-progress` events to the frontend with { phase, current, total }.
///
/// **Async command**: runs PDF generation on tokio::task::spawn_blocking so
/// the IPC thread is freed immediately. This ensures the frontend JS thread
/// can process pdf-progress events and update the UI (progress bar) while
/// the CPU-heavy work proceeds in the background — no UI freeze.
///
/// - `print_after`: if `Some(true)` (default), print after generating; if `Some(false)`, skip print.
#[command]
async fn generate_pdf_from_layout(
    app: tauri::AppHandle,
    request: LayoutRenderRequest,
    output_path: String,
    direct_print: Option<bool>,
    printer_name: Option<String>,
    print_after: Option<bool>,
) -> Result<pdf_engine::PdfResult, String> {
    use std::sync::atomic::Ordering;

    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    let output = std::path::PathBuf::from(&output_path);
    let app_handle = app.clone();

    let progress_cb: pdf_engine::ProgressFn = Box::new(move |phase, current, total| {
        let _ = app_handle.emit("pdf-progress", serde_json::json!({
            "phase": phase,
            "current": current,
            "total": total,
        }));
    });

    let output_for_print = output.clone();
    // Move the CPU-heavy PDF generation onto a blocking thread.
    // The async fn returns immediately, freeing the IPC thread so
    // frontend can receive progress events and repaint the UI.
    let request = request;
    tauri::async_runtime::spawn_blocking(move || {
        pdf_engine::generate_pdf_from_layout(&request, &output, Some(progress_cb))
    })
    .await
    .map_err(|e| format!("PDF生成任务失败: {}", e))?
    .map_err(|e| format!("PDF生成失败: {}", e))?;

    let should_print = print_after.unwrap_or(true);
    let is_direct = direct_print.unwrap_or(false);

    #[cfg(target_os = "windows")]
    if should_print {
        if is_direct {
            shell_execute_print(&output_for_print, printer_name.as_deref())?;
        } else {
            shell_execute("print", &output_for_print.to_string_lossy())?;
        }
    }

    let msg = if !should_print {
        "PDF已生成".to_string()
    } else if is_direct {
        if let Some(name) = printer_name {
            format!("已发送到打印机「{}」", name)
        } else {
            "已发送到默认打印机".to_string()
        }
    } else {
        "已弹出打印对话框".to_string()
    };

    Ok(pdf_engine::PdfResult {
        success: true,
        message: msg,
        pdf_path: Some(output_for_print.to_string_lossy().to_string()),
    })
}

/// Print or open an existing PDF file (skip PDF generation).
/// Used when the PDF hasn't changed since the last save/print.
#[command]
fn print_pdf_file(
    pdf_path: String,
    direct_print: Option<bool>,
    printer_name: Option<String>,
) -> Result<pdf_engine::PdfResult, String> {
    let output = std::path::Path::new(&pdf_path);
    if !output.exists() {
        return Err("PDF文件不存在".to_string());
    }

    let is_direct = direct_print.unwrap_or(false);

    #[cfg(target_os = "windows")]
    {
        if is_direct {
            shell_execute_print(output, printer_name.as_deref())?;
        } else {
            shell_execute("print", &output.to_string_lossy())?;
        }
    }

    let msg = if is_direct {
        if let Some(name) = printer_name {
            format!("已直接打印 → {}", name)
        } else {
            "已直接打印 → 默认打印机".to_string()
        }
    } else {
        "已弹出打印对话框".to_string()
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

/// Print a PDF file via ShellExecuteW.
/// - Direct mode with specific printer: uses "printto" verb (prints silently)
/// - Direct mode without printer: uses "print" verb with SW_HIDE (prints to default)
/// - Dialog mode: uses "print" verb with SW_SHOWNORMAL (shows print dialog)
#[cfg(target_os = "windows")]
fn shell_execute_print(pdf_path: &std::path::Path, printer_name: Option<&str>) -> Result<(), String> {
    use windows::core::HSTRING;
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::SW_HIDE;

    let _com = ComGuard::init();
    unsafe {
        // "printto" verb: prints to a specific printer without dialog
        // "print" verb with SW_HIDE: prints to default printer without dialog
        let verb: HSTRING = if printer_name.is_some() { "printto" } else { "print" }.into();
        let file: HSTRING = pdf_path.to_string_lossy().to_string().into();

        // Printer name goes into lpParameters for "printto" verb
        let printer_hstring: Option<HSTRING> = printer_name.map(|n| n.into());
        let params = printer_hstring.as_ref()
            .map(|h| windows::core::PCWSTR::from_raw(h.as_ptr()))
            .unwrap_or(windows::core::PCWSTR::null());

        let ret = ShellExecuteW(
            None,
            &verb,
            &file,
            params,
            windows::core::PCWSTR::null(),
            SW_HIDE,
        );
        if ret.0 as isize <= 32 {
            return Err(format!("打印失败，错误码: {}。请确认已安装PDF阅读器且关联了打印功能。", ret.0 as isize));
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
            // Window starts hidden via tauri.conf.json "visible: false" — no white flash.
            // The frontend will call show_window once the UI is fully rendered.
            // Handle window events: close + drag-and-drop
            {
                use tauri::Manager;
                let window = app.get_webview_window("main").unwrap();
                let win = window.clone();

                window.on_window_event(move |event| {
                    match event {
                        // --- Close event: force-exit with short delay ---
                        // ROOT CAUSE of residual processes:
                        // Tauri sync commands run inside tokio's spawn_blocking pool.
                        // WinRT .get() blocks the OS thread (OCR/PDF rendering can take seconds).
                        // When the user closes the window, tokio's Runtime::drop() waits for
                        // ALL spawn_blocking tasks to finish before allowing the process to exit.
                        // If OCR is still running, tokio waits forever → the main process hangs.
                        // WebView2 child processes (msedgewebview2.exe) then become orphans.
                        //
                        // FIX: Set SHUTTING_DOWN so in-flight loops check it on next iteration,
                        // notify frontend to stop queuing new OCR work, then spawn a watchdog
                        // that calls std::process::exit(0) after a short grace period.
                        // std::process::exit(0) calls the Windows ExitProcess API, which
                        // immediately terminates the process and ALL its threads — tokio cannot
                        // block it. WebView2 jobs are cleaned up by the OS within milliseconds.
                        tauri::WindowEvent::CloseRequested { .. } => {
                            use std::sync::atomic::Ordering;
                            if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
                                return; // Already shutting down — let close proceed
                            }
                            // Signal all Rust long-running operations to abort on next checkpoint
                            pdf_engine::SHUTTING_DOWN.store(true, Ordering::SeqCst);
                            // Notify frontend to clear OCR queues and stop new work (best-effort)
                            let _ = win.eval("if(window._tauriCleanup)window._tauriCleanup();");
                            // Unconditional immediate process termination.
                            // Do NOT use a delayed thread — if tokio is blocked on a
                            // spawn_blocking WinRT .get() call, the delay thread may
                            // race with the runtime drop and never get scheduled.
                            std::process::exit(0);
                        }
                        tauri::WindowEvent::Destroyed => {
                            // Fallback: if CloseRequested was somehow bypassed
                            // (e.g. programmatic close, system shutdown), ensure
                            // the process dies immediately when the window is gone.
                            std::process::exit(0);
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
            render_and_ocr_pdf,
            open_url,
            open_file,
            ocr_image,
            ocr_pdf_page,
            get_config,
            get_temp_dir,
            show_window,
            trim_image,
            generate_pdf_from_layout,
            print_pdf_file,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
