use tauri::{command, Emitter};

mod pdf_engine;
use pdf_engine::{PrinterInfo, FileData, RenderedPage, ComGuard, LayoutRenderRequest};
#[cfg(feature = "ocr")]
use pdf_engine::{OcrResult, RenderedOcrPage};

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
#[cfg(feature = "ocr")]
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
#[cfg(feature = "ocr")]
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
#[cfg(feature = "ocr")]
#[command]
fn ocr_pdf_page(pdf_path: String, page_index: u32, dpi: Option<u32>) -> Result<OcrResult, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::ocr_pdf_page(&pdf_path, page_index, dpi)
}

/// Check whether OCR feature is available at runtime.
/// Frontend calls this once at startup to decide whether to show OCR UI.
#[command]
fn check_ocr_available() -> bool {
    pdf_engine::check_ocr_available()
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
            // Direct print: use "printto" verb + SW_HIDE to print silently
            shell_execute_print(&output_for_print, printer_name.as_deref())?;
        } else {
            // Dialog mode: use "print" verb + SW_SHOWNORMAL to show print dialog
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
            // Direct print: use "printto" verb + SW_HIDE to print silently
            shell_execute_print(output, printer_name.as_deref())?;
        } else {
            // Dialog mode: use "print" verb + SW_SHOWNORMAL to show print dialog
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

/// Print a PDF file silently via Windows Print Spooler API.
/// Sends PDF bytes directly to the printer queue — no application window opens.
/// If no printer_name is provided, automatically resolves the system default printer.
/// Falls back to ShellExecuteW "printto" verb if the Spooler approach fails.
#[cfg(target_os = "windows")]
fn shell_execute_print(pdf_path: &std::path::Path, printer_name: Option<&str>) -> Result<(), String> {
    use windows::core::{HSTRING, PCWSTR};

    // Auto-resolve default printer if none specified
    let resolved_printer: Option<String> = match printer_name {
        Some(name) => Some(name.to_string()),
        None => pdf_engine::get_default_printer_name(),
    };
    let printer_str = resolved_printer.as_deref()
        .ok_or("未找到默认打印机，请在系统设置中配置打印机，或在打印设置中手动选择。")?;

    // Read PDF file bytes
    let pdf_bytes = std::fs::read(pdf_path)
        .map_err(|e| format!("读取PDF文件失败: {}", e))?;

    // Try Print Spooler API first (truly silent, no window)
    match spool_print_pdf(&pdf_bytes, printer_str) {
        Ok(()) => return Ok(()),
        Err(spool_err) => {
            log::warn!("Spooler printing failed, falling back to ShellExecute: {}", spool_err);
            // Fall through to ShellExecuteW fallback
        }
    }

    // Fallback: ShellExecuteW "printto" + SW_HIDE
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::SW_HIDE;

    let _com = ComGuard::init();
    unsafe {
        let verb: HSTRING = "printto".into();
        let file: HSTRING = pdf_path.to_string_lossy().to_string().into();
        let printer_hstring: HSTRING = printer_str.into();
        let params = PCWSTR::from_raw(printer_hstring.as_ptr());

        let ret = ShellExecuteW(
            None,
            &verb,
            &file,
            params,
            PCWSTR::null(),
            SW_HIDE,
        );
        if ret.0 as isize <= 32 {
            return Err(format!("打印失败，错误码: {}。请确认已安装PDF阅读器且关联了打印功能。", ret.0 as isize));
        }
    }
    Ok(())
}

/// RAII wrapper for printer handle — ensures ClosePrinter is called on drop.
#[cfg(target_os = "windows")]
struct PrinterHandle(windows::Win32::Foundation::HANDLE);
#[cfg(target_os = "windows")]
impl Drop for PrinterHandle {
    fn drop(&mut self) {
        use windows::Win32::Graphics::Printing::ClosePrinter;
        if !self.0.is_invalid() {
            unsafe { let _ = ClosePrinter(self.0); }
        }
    }
}

/// Send PDF bytes directly to printer via Windows Print Spooler API.
/// No application window opens — the spooler handles everything.
#[cfg(target_os = "windows")]
fn spool_print_pdf(pdf_bytes: &[u8], printer_name: &str) -> Result<(), String> {
    use windows::core::{HSTRING, PCWSTR, PWSTR};
    use windows::Win32::Foundation::{BOOL, HANDLE};
    use windows::Win32::Graphics::Printing::{DOC_INFO_1W, EndDocPrinter, OpenPrinterW, StartDocPrinterW, WritePrinter};

    let printer_hstring: HSTRING = printer_name.into();
    let mut h_printer: HANDLE = HANDLE(std::ptr::null_mut());

    unsafe {
        // Open printer
        OpenPrinterW(
            PCWSTR::from_raw(printer_hstring.as_ptr()),
            &mut h_printer as *mut HANDLE,
            None,
        ).map_err(|e| format!("打开打印机失败: {}", e))?;

        // RAII guard: close printer handle on any exit path
        let _guard = PrinterHandle(h_printer);

        // Prepare document info: use "RAW" data type to pass bytes directly to printer driver
        let doc_name = HSTRING::from("发票打印");
        let data_type = HSTRING::from("RAW");
        let doc_info = DOC_INFO_1W {
            pDocName: PWSTR::from_raw(doc_name.as_ptr() as *mut _),
            pOutputFile: PWSTR::null(),
            pDatatype: PWSTR::from_raw(data_type.as_ptr() as *mut _),
        };

        // Start document
        let doc_id = StartDocPrinterW(h_printer, 1, &doc_info);
        if doc_id == 0 {
            return Err(format!("StartDocPrinter 失败，打印机可能不支持RAW数据类型"));
        }

        // Write PDF bytes in chunks (WritePrinter has a u32 size limit per call)
        let mut offset: usize = 0;
        let total = pdf_bytes.len();
        while offset < total {
            let chunk_size = std::cmp::min(total - offset, 64 * 1024) as u32; // 64KB chunks
            let mut written: u32 = 0;
            let result = WritePrinter(
                h_printer,
                pdf_bytes[offset..].as_ptr() as *const _,
                chunk_size,
                &mut written,
            );
            if result == BOOL(0) {
                let _ = EndDocPrinter(h_printer);
                return Err("WritePrinter 写入数据失败".to_string());
            }
            offset += written as usize;
        }

        // End document
        if EndDocPrinter(h_printer) == BOOL(0) {
            return Err("EndDocPrinter 失败".to_string());
        }
    }

    Ok(())
}

// =====================================================
// App Entry
// =====================================================

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let builder = tauri::Builder::default()
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
            {
                use tauri::Manager;
                let window = app.get_webview_window("main").unwrap();
                let win = window.clone();

                window.on_window_event(move |event| {
                    match event {
                        tauri::WindowEvent::CloseRequested { api, .. } => {
                            use std::sync::atomic::Ordering;
                            // CRITICAL: Prevent Tauri's default close sequence from running
                            // concurrently with our cleanup. Without this, WebView2 starts
                            // its own shutdown while we're still trying to eval() JS and
                            // call process::exit(0), causing deadlocks.
                            api.prevent_close();

                            if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
                                return; // Exit thread already running
                            }
                            pdf_engine::SHUTTING_DOWN.store(true, Ordering::SeqCst);
                            let _ = win.eval("if(window._tauriCleanup)window._tauriCleanup();");

                            // Spawn exit in a separate thread.
                            //
                            // We use TerminateProcess instead of process::exit(0) (ExitProcess)
                            // because ExitProcess has a fatal flaw: it first kills ALL other
                            // threads, then runs DLL_PROCESS_DETACH for each loaded DLL.
                            // If DLL_PROCESS_DETACH hangs (e.g. MNN/OCR engine deadlock),
                            // the process is stuck forever — and the "backup" thread was
                            // already killed, so TerminateProcess never runs.
                            //
                            // TerminateProcess skips DLL_PROCESS_DETACH entirely and terminates
                            // the process immediately — it can never hang.
                            // The 300ms delay gives pending I/O (file writes, OCR cancellation)
                            // time to complete before we pull the plug.
                            std::thread::spawn(move || {
                                std::thread::sleep(std::time::Duration::from_millis(300));
                                #[cfg(target_os = "windows")]
                                unsafe {
                                    use windows::Win32::System::Threading::{GetCurrentProcess, TerminateProcess};
                                    let _ = TerminateProcess(GetCurrentProcess(), 0);
                                }
                                #[cfg(not(target_os = "windows"))]
                                std::process::exit(0);
                            });
                        }
                        tauri::WindowEvent::Destroyed => {
                            // Use TerminateProcess instead of process::exit(0) for the same
                            // reason as CloseRequested: ExitProcess can deadlock in
                            // DLL_PROCESS_DETACH (e.g. MNN/OCR engine holding a lock).
                            #[cfg(target_os = "windows")]
                            unsafe {
                                use windows::Win32::System::Threading::{GetCurrentProcess, TerminateProcess};
                                let _ = TerminateProcess(GetCurrentProcess(), 0);
                            }
                            #[cfg(not(target_os = "windows"))]
                            std::process::exit(0);
                        }
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
        });

    // Register commands — OCR commands are conditionally included
    #[cfg(feature = "ocr")]
    let builder = builder.invoke_handler(tauri::generate_handler![
        open_invoice_files,
        get_printers,
        render_pdf_pages,
        render_and_ocr_pdf,
        open_url,
        open_file,
        ocr_image,
        ocr_pdf_page,
        check_ocr_available,
        get_config,
        get_temp_dir,
        show_window,
        trim_image,
        generate_pdf_from_layout,
        print_pdf_file,
    ]);

    #[cfg(not(feature = "ocr"))]
    let builder = builder.invoke_handler(tauri::generate_handler![
        open_invoice_files,
        get_printers,
        render_pdf_pages,
        open_url,
        open_file,
        check_ocr_available,
        get_config,
        get_temp_dir,
        show_window,
        trim_image,
        generate_pdf_from_layout,
        print_pdf_file,
    ]);

    builder
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
