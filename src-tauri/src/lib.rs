use std::sync::atomic::{AtomicBool, Ordering as AtomicOrdering};
use tauri::{command, Emitter};

mod pdf_engine;
#[cfg(target_os = "windows")]
mod pdfium_print;

static DOWNLOAD_CANCELLED: AtomicBool = AtomicBool::new(false);

#[cfg(target_os = "windows")]
fn check_windows_version() -> Result<(), String> {
    use windows::core::*;
    use windows::Win32::System::Registry::*;
    
    unsafe {
        let mut hkey = HKEY::default();
        let result = RegOpenKeyExW(
            HKEY_LOCAL_MACHINE,
            w!("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"),
            0,
            KEY_READ,
            &mut hkey,
        );
        
        if result.is_ok() {
            let mut build_number = [0u16; 256];
            let mut build_number_size = (build_number.len() * 2) as u32;
            
            let result = RegQueryValueExW(
                hkey,
                w!("CurrentBuildNumber"),
                None,
                None,
                Some(build_number.as_mut_ptr() as *mut u8),
                Some(&mut build_number_size),
            );
            
            let _ = RegCloseKey(hkey);
            
            if result.is_ok() {
                let build_str = String::from_utf16_lossy(&build_number[..(build_number_size as usize / 2)]);
                let build_str = build_str.trim_end_matches('\0');
                
                if let Ok(build) = build_str.parse::<u32>() {
                    if build < 17134 {
                        return Err(format!(
                            "您的系统版本不支持本应用。\n\n当前系统：Windows (Build {})\n\n需要：Windows 10 1803 (Build 17134) 或 Windows 11",
                            build
                        ));
                    }
                    return Ok(());
                }
            }
        }
        
        Ok(())
    }
}

#[cfg(target_os = "windows")]
fn show_error_dialog(message: &str) {
    use windows::core::*;
    use windows::Win32::UI::WindowsAndMessaging::*;

    unsafe {
        let _ = MessageBoxW(
            None,
            &HSTRING::from(message),
            &HSTRING::from("发票酱 - 系统不兼容"),
            MB_ICONERROR | MB_OK,
        );
    }
}
use pdf_engine::{PrinterInfo, FileData, RenderedPage, ComGuard, LayoutRenderRequest, PdfTextResult};
#[cfg(feature = "ocr")]
use pdf_engine::{OcrResult, RenderedOcrPage};

// =====================================================
// Tauri Commands
// =====================================================

/// Read files from given paths (for drag-and-drop and dialog plugin)
#[command]
async fn open_invoice_files(paths: Vec<String>) -> Result<Vec<FileData>, String> {
    tauri::async_runtime::spawn_blocking(move || {
        pdf_engine::read_invoice_files(paths)
    }).await.map_err(|e| format!("文件读取任务失败: {}", e))?
}

/// Parse OFD file: returns SVG vector rendering + structured invoice data from XML.
/// Skips OCR — invoice fields are extracted directly from OFD metadata.
#[command]
fn parse_ofd(ofd_path: String) -> Result<invoice_engine::OfdResult, String> {
    invoice_engine::parse_ofd_file(&ofd_path)
}

/// Parse standalone XML 数电票 file: extracts structured invoice data.
/// XML 数电票 is a pure data format with no layout info — cannot be rendered as a visual page.
/// Used for file list display, summary export, and batch rename.
#[command]
async fn parse_xml_invoice(xml_path: String) -> Result<invoice_engine::XmlInvoiceInfo, String> {
    tauri::async_runtime::spawn_blocking(move || {
        invoice_engine::parse_xml_invoice(&xml_path)
    }).await.map_err(|e| format!("XML 解析任务失败: {}", e))?
}

/// Fallback: extract OFD pages as bitmap images (legacy path).
/// Used when parse_ofd fails (e.g., vector-only OFD with no parseable XML).
#[command]
fn open_ofd_images(ofd_path: String) -> Result<Vec<FileData>, String> {
    let path = std::path::Path::new(&ofd_path);
    let name = path.file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_else(|| "unknown".to_string());
    let size = path.metadata().ok().map(|m| m.len()).unwrap_or(0);

    let images = invoice_engine::extract_ofd_images_raw(&ofd_path)?;
    let mut results = Vec::new();
    for (idx, img) in images.iter().enumerate() {
        let base_name = if name.len() > 4 { &name[..name.len()-4] } else { &name };
        results.push(FileData {
            name: if images.len() > 1 {
                format!("{}_第{}页.ofd", base_name, idx + 1)
            } else {
                name.clone()
            },
            ext: img.ext.clone(),
            size,
            data_url: img.data_url.clone(),
            path: None,
            orig_w: Some(img.width),
            orig_h: Some(img.height),
        });
    }
    Ok(results)
}

/// List available printers
#[command]
fn get_printers() -> Result<Vec<PrinterInfo>, String> {
    pdf_engine::list_printers()
}

/// Render PDF pages to images using Windows native API
/// - `use_jpeg`: if true, encode as JPEG for faster loading (default: true for preview)
#[command]
async fn render_pdf_pages(pdf_path: String, dpi: Option<u32>, use_jpeg: Option<bool>) -> Result<Vec<RenderedPage>, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    let d = dpi.unwrap_or(pdf_engine::RENDER_DPI);
    let j = use_jpeg.unwrap_or(true);
    tauri::async_runtime::spawn_blocking(move || {
        pdf_engine::render_pdf_pages(&pdf_path, d, j)
    }).await.map_err(|e| e.to_string())?
}

#[command]
async fn render_pdf_pages_pdfium(pdf_path: String, dpi: Option<u32>, use_jpeg: Option<bool>) -> Result<Vec<RenderedPage>, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    let d = dpi.unwrap_or(pdf_engine::RENDER_DPI);
    let j = use_jpeg.unwrap_or(true);
    tauri::async_runtime::spawn_blocking(move || {
        pdf_engine::render_pdf_pages_pdfium(&pdf_path, d, j)
    }).await.map_err(|e| e.to_string())?
}

#[command]
fn check_winrt_pdf() -> bool {
    pdf_engine::check_winrt_pdf_available()
}

/// Render PDF pages AND run OCR in one pass — avoids IPC round-trip.
/// Returns preview images + OCR results together.
#[cfg(feature = "ocr")]
#[command]
fn render_and_ocr_pdf(pdf_path: String, dpi: Option<u32>, ocr_precision: Option<String>) -> Result<Vec<RenderedOcrPage>, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::render_and_ocr_pdf(&pdf_path, dpi.unwrap_or(pdf_engine::RENDER_DPI), ocr_precision.as_deref())
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

/// Check whether a path exists and whether it's a directory
#[command]
fn check_path_exists(path: String) -> Result<serde_json::Value, String> {
    let p = std::path::Path::new(&path);
    Ok(serde_json::json!({
        "exists": p.exists(),
        "isDir": p.is_dir(),
        "isFile": p.is_file(),
        "path": path
    }))
}

/// Copy a file from srcPath to destPath.
/// Creates parent directories for the destination if they don't exist.
#[command]
fn copy_file(src_path: String, dest_path: String) -> Result<serde_json::Value, String> {
    let src = std::path::Path::new(&src_path);
    if !src.exists() {
        return Err(format!("Failed to copy file: 源文件不存在: {}", src_path));
    }
    let dest = std::path::Path::new(&dest_path);
    if let Some(parent) = dest.parent() {
        if !parent.exists() {
            std::fs::create_dir_all(parent)
                .map_err(|e| format!("Failed to copy file: 无法创建目标目录 '{}': {}", parent.display(), e))?;
        }
    }
    let size = std::fs::copy(src, dest)
        .map_err(|e| format!("Failed to copy file: {} → {}: {}", src_path, dest_path, e))?;
    Ok(serde_json::json!({
        "success": true,
        "srcPath": src_path,
        "destPath": dest_path,
        "size": size
    }))
}

/// Rename/move a file from srcPath to destPath.
/// Tries atomic fs::rename first; falls back to copy+delete on cross-device errors.
/// Creates parent directories for the destination if they don't exist.
/// Will NOT overwrite an existing destination file.
#[command]
async fn rename_file(src_path: String, dest_path: String) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let src = std::path::Path::new(&src_path);
        if !src.exists() {
            return Err(format!("源文件不存在: {}", src_path));
        }
        let dest = std::path::Path::new(&dest_path);
        if dest.exists() {
            return Err(format!("目标文件已存在: {}", dest_path));
        }
        if let Some(parent) = dest.parent() {
            if !parent.exists() {
                std::fs::create_dir_all(parent)
                    .map_err(|e| format!("无法创建目标目录 '{}': {}", parent.display(), e))?;
            }
        }
        // Try atomic rename first (same device)
        match std::fs::rename(src, dest) {
            Ok(()) => {}
            Err(e) => {
                // On cross-device error, fall back to copy + delete
                if e.raw_os_error() == Some(17) {
                    // ERROR_NOT_SAME_DEVICE on Windows
                    let size = std::fs::copy(src, dest)
                        .map_err(|e2| format!("跨盘符复制失败 {} → {}: {}", src_path, dest_path, e2))?;
                    std::fs::remove_file(src)
                        .map_err(|e2| format!("删除源文件失败 {}: {}", src_path, e2))?;
                    return Ok(serde_json::json!({
                        "success": true,
                        "srcPath": src_path,
                        "destPath": dest_path,
                        "size": size,
                        "fallback": "copy_delete"
                    }));
                }
                return Err(format!("重命名失败 {} → {}: {}", src_path, dest_path, e));
            }
        }
        Ok(serde_json::json!({
            "success": true,
            "srcPath": src_path,
            "destPath": dest_path
        }))
    }).await.map_err(|e| format!("任务执行失败: {}", e))?
}

/// Write UTF-8 text content to a file (creates parent dirs if needed).
/// Used by the CSV summary export feature.
#[command]
async fn write_text_file(path: String, content: String) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let dest = std::path::Path::new(&path);
        if let Some(parent) = dest.parent() {
            if !parent.exists() {
                std::fs::create_dir_all(parent)
                    .map_err(|e| format!("无法创建目录 '{}': {}", parent.display(), e))?;
            }
        }
        std::fs::write(dest, &content)
            .map_err(|e| format!("写入文件失败: {}", e))?;
        Ok(serde_json::json!({ "success": true, "path": path }))
    }).await.map_err(|e| format!("任务执行失败: {}", e))?
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
fn ocr_image(data_url: String, file_path: Option<String>, ocr_precision: Option<String>) -> Result<OcrResult, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::ocr_image(&data_url, file_path.as_deref(), ocr_precision.as_deref())
}

/// Render a single PDF page and run OCR on it — zero IPC round-trip.
/// The frontend calls this instead of `render_pdf_pages` + `ocr_image` for PDF pages,
/// avoiding the expensive Rust→base64→IPC→frontend→downsample→base64→IPC→Rust cycle.
#[cfg(feature = "ocr")]
#[command]
fn ocr_pdf_page(pdf_path: String, page_index: u32, dpi: Option<u32>, ocr_precision: Option<String>) -> Result<OcrResult, String> {
    use std::sync::atomic::Ordering;
    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    pdf_engine::ocr_pdf_page(&pdf_path, page_index, dpi, ocr_precision.as_deref())
}

/// Check whether OCR feature is available at runtime.
/// Frontend calls this once at startup to decide whether to show OCR UI.
#[command]
fn check_ocr_available() -> bool {
    pdf_engine::check_ocr_available()
}

/// Extract text with coordinates from a PDF page's content stream.
/// No OCR needed — parses the PDF's native text layer directly.
/// ~5ms per page vs ~1-3s for OCR. Works in lightweight (non-OCR) builds.
///
/// Uses `spawn_blocking` to run CPU-intensive PDF parsing off the IPC thread,
/// preventing IPC message pump starvation that causes `ERR_CONNECTION_REFUSED`.
#[command]
async fn extract_pdf_text(pdf_path: String, page_idx: u32) -> Result<PdfTextResult, String> {
    tauri::async_runtime::spawn_blocking(move || {
        pdf_engine::extract_pdf_text(&pdf_path, page_idx)
            .map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| format!("PDF文字提取任务失败: {}", e))?
}

/// Extract text from multiple pages in a PDF in parallel, with only one PDF file open.
/// ~5ms per page, with parallelism (rayon) for multi-page PDFs.
/// Uses `spawn_blocking` to run CPU-intensive PDF parsing off the IPC thread.
#[command]
async fn extract_pdf_texts(pdf_path: String, page_indices: Vec<u32>) -> Result<std::collections::HashMap<u32, PdfTextResult>, String> {
    tauri::async_runtime::spawn_blocking(move || {
        pdf_engine::extract_pdf_texts(&pdf_path, &page_indices)
            .map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| format!("PDF文字提取任务失败: {}", e))?
}

/// Get user Downloads directory path
#[command]
fn get_downloads_dir() -> Result<String, String> {
    dirs::download_dir()
        .map(|p| p.to_string_lossy().to_string())
        .ok_or_else(|| "无法获取下载目录".to_string())
}

/// Get app version from Cargo.toml (compiled in at build time)
#[command]
fn get_app_version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

// =====================================================
// Update Check — fetch latest release from GitHub
// =====================================================

#[derive(serde::Serialize)]
struct UpdateAsset {
    name: String,
    download_url: String,
    size: u64,
}

#[derive(serde::Serialize)]
struct UpdateInfo {
    has_update: bool,
    current_version: String,
    latest_version: String,
    release_notes: String,
    release_url: String,
    published_at: String,
    assets: Vec<UpdateAsset>,
}

/// Compare two semver-like version strings (e.g. "2.0.9" vs "2.1.0").
/// Returns -1 if a < b, 0 if equal, 1 if a > b.
fn compare_versions(a: &str, b: &str) -> i32 {
    let pa: Vec<u32> = a.split('.').filter_map(|s| s.parse().ok()).collect();
    let pb: Vec<u32> = b.split('.').filter_map(|s| s.parse().ok()).collect();
    let n = std::cmp::max(pa.len(), pb.len());
    for i in 0..n {
        let va = *pa.get(i).unwrap_or(&0);
        let vb = *pb.get(i).unwrap_or(&0);
        if va < vb { return -1; }
        if va > vb { return 1; }
    }
    0
}

/// Check for updates by querying GitHub Releases API (latest release).
/// Returns update info including latest version, release notes, and download assets.
/// Uses async reqwest to avoid blocking the IPC thread.
///
/// 主备双源: 先尝试直连 api.github.com,失败后 fallback 到 gh-proxy.com 加速代理。
/// 保证大陆网络环境下更新检查可用,网络通畅时直连最快。
#[command]
async fn check_for_updates() -> Result<UpdateInfo, String> {
    let current = env!("CARGO_PKG_VERSION").to_string();

    let client = reqwest::Client::builder()
        .timeout(std::time::Duration::from_secs(15))
        .build()
        .map_err(|e| format!("创建HTTP客户端失败: {}", e))?;

    // 主备双源: 直连 → gh-proxy.com 代理
    let urls = [
        "https://api.github.com/repos/erma0/fapiao-print/releases/latest",
        "https://gh-proxy.com/https://api.github.com/repos/erma0/fapiao-print/releases/latest",
    ];

    let mut body: Option<String> = None;
    let mut last_err = String::new();
    for url in &urls {
        match client
            .get(*url)
            .header("User-Agent", "ticketchan-updater")
            .header("Accept", "application/vnd.github+json")
            .send()
            .await
        {
            Ok(resp) if resp.status().is_success() => {
                match resp.text().await {
                    Ok(text) => { body = Some(text); break; }
                    Err(e) => last_err = format!("读取响应失败: {}", e),
                }
            }
            Ok(resp) => last_err = format!("API返回状态: {}", resp.status()),
            Err(e) => last_err = format!("请求失败: {}", e),
        }
    }
    let body = body.ok_or_else(|| format!("GitHub更新检查失败: {}", last_err))?;

    let json: serde_json::Value = serde_json::from_str(&body)
        .map_err(|e| format!("解析响应失败: {}", e))?;

    let tag = json["tag_name"].as_str().unwrap_or("").to_string();
    let latest = tag.trim_start_matches('v').to_string();
    let release_notes = json["body"].as_str().unwrap_or("").to_string();
    let release_url = json["html_url"].as_str().unwrap_or("").to_string();
    let published_at = json["published_at"].as_str().unwrap_or("").to_string();

    let mut assets = Vec::new();
    if let Some(arr) = json["assets"].as_array() {
        for a in arr {
            assets.push(UpdateAsset {
                name: a["name"].as_str().unwrap_or("").to_string(),
                download_url: a["browser_download_url"].as_str().unwrap_or("").to_string(),
                size: a["size"].as_u64().unwrap_or(0),
            });
        }
    }

    let has_update = compare_versions(&current, &latest) < 0;

    Ok(UpdateInfo {
        has_update,
        current_version: current,
        latest_version: latest,
        release_notes,
        release_url,
        published_at,
        assets,
    })
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

/// Enhance a faint/blurry invoice image (levels stretch + gamma + unsharp mask).
/// Reads the original file at FULL resolution so print clarity is preserved —
/// the preview thumbnail is never used as enhancement source.
/// Returns enhanced JPEG data URL with EXIF orientation baked into pixels.
///
/// **Async command**: CPU-intensive pixel work runs on spawn_blocking,
/// keeping the IPC thread free (same pattern as extract_pdf_text).
#[command]
async fn enhance_image(file_path: String) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || {
        pdf_engine::enhance_image(&file_path)
    })
    .await
    .map_err(|e| format!("图片增强任务失败: {}", e))?
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
    let print_settings = (
        request.settings.copies,
        request.settings.duplex,
        request.settings.color_mode.clone(),
        request.settings.fit_mode.clone(),
        request.settings.paper_w,
        request.settings.paper_h,
    );
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
    let mut shell_fallback_open = false;
    #[cfg(not(target_os = "windows"))]
    let shell_fallback_open = false;
    #[cfg(target_os = "windows")]
    if should_print {
        if is_direct {
            let resolved_printer = match &printer_name {
                Some(name) if !name.is_empty() => name.clone(),
                _ => pdf_engine::get_default_printer_name()
                    .ok_or("未找到默认打印机".to_string())?,
            };
            if let Some(sumatra) = pdf_engine::find_sumatrapdf() {
                let settings_str = pdf_engine::build_sumatra_print_settings(
                    print_settings.0,
                    print_settings.1,
                    &print_settings.2,
                    &print_settings.3,
                    Some(print_settings.4),
                    Some(print_settings.5),
                );
                pdf_engine::print_with_sumatrapdf(
                    &sumatra.path, &output_for_print, &resolved_printer, &settings_str,
                )?;
            } else {
                let printed = shell_execute_print(&output_for_print, printer_name.as_deref())?;
                if !printed {
                    shell_fallback_open = true;
                }
            }
        } else {
            shell_execute("open", &output_for_print.to_string_lossy())?;
        }
    }

    let msg = if !should_print {
        "PDF已生成".to_string()
    } else if shell_fallback_open {
        "PDF阅读器不支持静默打印，已打开PDF，请手动打印".to_string()
    } else if is_direct {
        if let Some(name) = printer_name {
            format!("已发送到打印机「{}」", name)
        } else {
            "已发送到默认打印机".to_string()
        }
    } else {
        "已打开PDF，请在阅读器中打印".to_string()
    };

    Ok(pdf_engine::PdfResult {
        success: true,
        message: msg,
        pdf_path: Some(output_for_print.to_string_lossy().to_string()),
        warnings: None,
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
    let mut shell_fallback_open = false;
    #[cfg(not(target_os = "windows"))]
    let shell_fallback_open = false;
    #[cfg(target_os = "windows")]
    {
        if is_direct {
            let printed = shell_execute_print(output, printer_name.as_deref())?;
            if !printed {
                shell_fallback_open = true;
            }
        } else {
            shell_execute("open", &output.to_string_lossy())?;
        }
    }

    let msg = if shell_fallback_open {
        "PDF阅读器不支持静默打印，已打开PDF，请手动打印".to_string()
    } else if is_direct {
        if let Some(name) = printer_name {
            format!("已直接打印 → {}", name)
        } else {
            "已直接打印 → 默认打印机".to_string()
        }
    } else {
        "已打开PDF，请在阅读器中选择打印".to_string()
    };

    Ok(pdf_engine::PdfResult {
        success: true,
        message: msg,
        pdf_path: Some(output.to_string_lossy().to_string()),
        warnings: None,
    })
}

/// Check if SumatraPDF is available on the system
#[command]
fn check_sumatrapdf_available() -> bool {
    pdf_engine::find_sumatrapdf().is_some()
}

/// Check if pdfium.dll is available for vector printing
#[command]
fn check_pdfium_available() -> bool {
    #[cfg(target_os = "windows")]
    {
        pdfium_print::find_pdfium_dll().is_some()
    }
    #[cfg(not(target_os = "windows"))]
    {
        false
    }
}

/// Cancel any in-progress download (pdfium.dll or SumatraPDF)
#[command]
fn cancel_download() {
    DOWNLOAD_CANCELLED.store(true, AtomicOrdering::SeqCst);
}

/// Print a PDF file using PDFium vector rendering (EMF → printer DC).
/// Generates PDF from layout, then prints via PDFium. Saves PDF path for cache reuse.
#[cfg(target_os = "windows")]
#[command]
async fn pdfium_vector_print(
    app: tauri::AppHandle,
    request: LayoutRenderRequest,
    printer_name: Option<String>,
    copies: u32,
    duplex: bool,
) -> Result<pdf_engine::PdfResult, String> {
    use std::sync::atomic::Ordering;

    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    let resolved_printer = match printer_name {
        Some(name) if !name.is_empty() => name,
        _ => pdf_engine::get_default_printer_name()
            .ok_or("未找到默认打印机，请在系统设置中配置打印机，或在打印设置中手动选择。")?,
    };

    let color_mode = request.settings.color_mode.clone();
    let paper_w = request.settings.paper_w;
    let paper_h = request.settings.paper_h;

    let app_handle = app.clone();
    let progress_cb: pdf_engine::ProgressFn = Box::new(move |phase, current, total| {
        let _ = app_handle.emit("pdf-progress", serde_json::json!({
            "phase": phase,
            "current": current,
            "total": total,
        }));
    });

    let temp_dir = std::env::temp_dir().join("ticketchan_pdfium");
    let _ = std::fs::create_dir_all(&temp_dir);
    let output_path = temp_dir.join("pdfium_cache.pdf");

    let output_path_for_print = output_path.clone();
    let _pdf_result = tauri::async_runtime::spawn_blocking(move || {
        pdf_engine::generate_pdf_from_layout(&request, &output_path, Some(progress_cb))
    })
    .await
    .map_err(|e| format!("PDF生成任务失败: {}", e))?
    .map_err(|e| {
        let _ = std::fs::remove_file(&output_path_for_print);
        format!("PDF生成失败: {}", e)
    })?;

    let pdf_path_str = output_path_for_print.to_string_lossy().to_string();

    let result = pdfium_print_pdf(
        app,
        pdf_path_str,
        Some(resolved_printer),
        copies,
        duplex,
        color_mode,
        paper_w,
        paper_h,
    ).await?;

    Ok(result)
}

/// Print an existing PDF file using PDFium vector rendering.
/// Used when the PDF hasn't changed since last generation (cache reuse).
#[cfg(target_os = "windows")]
#[command]
async fn pdfium_print_pdf(
    app: tauri::AppHandle,
    pdf_path: String,
    printer_name: Option<String>,
    copies: u32,
    duplex: bool,
    color_mode: String,
    paper_w: f32,
    paper_h: f32,
) -> Result<pdf_engine::PdfResult, String> {
    use std::sync::atomic::Ordering;

    if pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    let path = std::path::Path::new(&pdf_path);
    if !path.exists() {
        return Err("PDF文件不存在".to_string());
    }

    let resolved_printer = match printer_name {
        Some(name) if !name.is_empty() => name,
        _ => pdf_engine::get_default_printer_name()
            .ok_or("未找到默认打印机，请在系统设置中配置打印机，或在打印设置中手动选择。")?,
    };

    let pdf_bytes = std::fs::read(path)
        .map_err(|e| format!("读取PDF失败: {}", e))?;

    let app_for_progress = app.clone();
    let printer = resolved_printer.clone();
    let print_progress_cb = move |current: u32, total: u32| {
        let _ = app_for_progress.emit("pdf-progress", serde_json::json!({
            "phase": "print",
            "current": current,
            "total": total,
        }));
    };

    let print_result = tauri::async_runtime::spawn_blocking(move || {
        pdfium_print::pdfium_vector_print(
            &pdf_bytes,
            &printer,
            copies,
            duplex,
            &color_mode,
            paper_w,
            paper_h,
            Some(&print_progress_cb),
        )
    })
    .await
    .map_err(|e| format!("打印任务失败: {}", e))?
    .map_err(|e| format!("PDFium打印失败: {}", e))?;

    let mut result = print_result;
    result.pdf_path = Some(pdf_path);

    Ok(result)
}

/// Download pdfium.dll to the app's tools directory for vector printing
#[cfg(target_os = "windows")]
#[command]
async fn download_pdfium_dll(app: tauri::AppHandle) -> Result<pdf_engine::PdfResult, String> {
    DOWNLOAD_CANCELLED.store(false, AtomicOrdering::SeqCst);

    let exe_dir = std::env::current_exe()
        .map_err(|e| format!("无法获取应用路径: {}", e))?
        .parent()
        .ok_or("无法获取应用目录")?
        .to_path_buf();

    let tools_dir = exe_dir.join("tools");
    std::fs::create_dir_all(&tools_dir)
        .map_err(|e| format!("创建 tools 目录失败: {}", e))?;

    let dest = tools_dir.join("pdfium.dll");
    if dest.exists() {
        return Ok(pdf_engine::PdfResult {
            success: true,
            message: "pdfium.dll 已存在".to_string(),
            pdf_path: Some(dest.to_string_lossy().to_string()),
            warnings: None,
        });
    }

    let dll_url = "https://gh-proxy.com/https://github.com/bblanchon/pdfium-binaries/releases/download/chromium/7834/pdfium-win-x64.tgz";
    let tgz_path = tools_dir.join("pdfium-win-x64.tgz");

    let client = reqwest::Client::builder()
        .timeout(std::time::Duration::from_secs(120))
        .build()
        .map_err(|e| format!("创建下载客户端失败: {}", e))?;

    {
        let mut file = std::fs::File::create(&tgz_path)
            .map_err(|e| format!("创建临时文件失败: {}", e))?;
        let mut stream = client.get(dll_url)
            .send()
            .await
            .map_err(|e| format!("下载失败: {}", e))?;

        if !stream.status().is_success() {
            std::fs::remove_file(&tgz_path).ok();
            return Err(format!("下载失败，HTTP 状态: {}", stream.status()));
        }

        let total_size = stream.content_length().unwrap_or(0);
        let mut downloaded: u64 = 0;

        while let Some(chunk) = stream.chunk().await.map_err(|e| format!("下载出错: {}", e))? {
            if DOWNLOAD_CANCELLED.load(AtomicOrdering::SeqCst) {
                drop(file);
                std::fs::remove_file(&tgz_path).ok();
                return Err("下载已取消".to_string());
            }
            use std::io::Write;
            file.write_all(&chunk)
                .map_err(|e| format!("写入文件失败: {}", e))?;
            downloaded += chunk.len() as u64;
            let _ = app.emit("pdfium-download-progress", serde_json::json!({
                "current": downloaded,
                "total": total_size,
                "percent": if total_size > 0 { (downloaded as f64 / total_size as f64) * 100.0 } else { 0.0 }
            }));
        }
    }

    let tgz_file = std::fs::File::open(&tgz_path)
        .map_err(|e| format!("打开 tgz 失败: {}", e))?;
    let gz_decoder = flate2::read::GzDecoder::new(tgz_file);
    let mut archive = tar::Archive::new(gz_decoder);
    let mut found_dll = false;

    for entry_result in archive.entries().map_err(|e| format!("解析 tgz 失败: {}", e))? {
        let mut entry = entry_result.map_err(|e| format!("读取 tgz 条目失败: {}", e))?;
        let path = entry.path().map_err(|e| format!("获取路径失败: {}", e))?;
        let file_name = path.file_name()
            .map(|f| f.to_string_lossy().to_string())
            .unwrap_or_default();

        if file_name.eq_ignore_ascii_case("pdfium.dll") {
            let mut out_file = std::fs::File::create(&dest)
                .map_err(|e| format!("创建 pdfium.dll 失败: {}", e))?;
            std::io::copy(&mut entry, &mut out_file)
                .map_err(|e| format!("解压失败: {}", e))?;
            found_dll = true;
            break;
        }
    }

    std::fs::remove_file(&tgz_path).ok();

    if !found_dll {
        return Err("tgz 中未找到 pdfium.dll".to_string());
    }

    log::info!("pdfium.dll downloaded to: {}", dest.display());

    Ok(pdf_engine::PdfResult {
        success: true,
        message: format!("pdfium.dll 已下载到: {}", dest.display()),
        pdf_path: Some(dest.to_string_lossy().to_string()),
        warnings: None,
    })
}

#[cfg(not(target_os = "windows"))]
#[command]
async fn pdfium_vector_print(
    _app: tauri::AppHandle,
    _request: LayoutRenderRequest,
    _printer_name: Option<String>,
) -> Result<pdf_engine::PdfResult, String> {
    Err("PDFium打印仅支持Windows系统".to_string())
}

#[cfg(not(target_os = "windows"))]
#[command]
async fn pdfium_print_pdf(
    _app: tauri::AppHandle,
    _pdf_path: String,
    _printer_name: Option<String>,
    _copies: u32,
    _duplex: bool,
    _color_mode: String,
    _paper_w: f32,
    _paper_h: f32,
) -> Result<pdf_engine::PdfResult, String> {
    Err("PDFium打印仅支持Windows系统".to_string())
}

#[cfg(not(target_os = "windows"))]
#[command]
async fn download_pdfium_dll(_app: tauri::AppHandle) -> Result<pdf_engine::PdfResult, String> {
    Err("PDFium下载仅支持Windows系统".to_string())
}

/// Download SumatraPDF portable (ZIP) to the app's tools directory
#[command]
async fn download_sumatrapdf(app: tauri::AppHandle) -> Result<pdf_engine::PdfResult, String> {
    DOWNLOAD_CANCELLED.store(false, AtomicOrdering::SeqCst);

    let tools_dir = std::env::current_exe()
        .map_err(|e| format!("无法获取应用路径: {}", e))?
        .parent()
        .ok_or("无法获取应用目录")?
        .join("tools");

    std::fs::create_dir_all(&tools_dir)
        .map_err(|e| format!("创建目录失败: {}", e))?;

    let dest = tools_dir.join("SumatraPDF.exe");

    if dest.exists() {
        return Ok(pdf_engine::PdfResult {
            success: true,
            message: "SumatraPDF 已存在".to_string(),
            pdf_path: Some(dest.to_string_lossy().to_string()),
            warnings: None,
        });
    }

    let zip_url = "https://www.sumatrapdfreader.org/dl/rel/3.6.1/SumatraPDF-3.6.1-64.zip";
    let zip_path = tools_dir.join("SumatraPDF-3.6.1-64.zip");

    let client = reqwest::Client::builder()
        .timeout(std::time::Duration::from_secs(120))
        .build()
        .map_err(|e| format!("创建下载客户端失败: {}", e))?;

    {
        let mut file = std::fs::File::create(&zip_path)
            .map_err(|e| format!("创建临时文件失败: {}", e))?;
        let mut stream = client.get(zip_url)
            .send()
            .await
            .map_err(|e| format!("下载失败: {}", e))?;

        if !stream.status().is_success() {
            std::fs::remove_file(&zip_path).ok();
            return Err(format!("下载失败，HTTP 状态: {}", stream.status()));
        }

        let total_size = stream.content_length().unwrap_or(0);
        let mut downloaded: u64 = 0;

        while let Some(chunk) = stream.chunk().await.map_err(|e| format!("下载出错: {}", e))? {
            if DOWNLOAD_CANCELLED.load(AtomicOrdering::SeqCst) {
                drop(file);
                std::fs::remove_file(&zip_path).ok();
                return Err("下载已取消".to_string());
            }
            use std::io::Write;
            file.write_all(&chunk)
                .map_err(|e| format!("写入文件失败: {}", e))?;
            downloaded += chunk.len() as u64;
            let _ = app.emit("sumatra-download-progress", serde_json::json!({
                "current": downloaded,
                "total": total_size,
                "percent": if total_size > 0 { (downloaded as f64 / total_size as f64) * 100.0 } else { 0.0 }
            }));
        }
    }

    let zip_file = std::fs::File::open(&zip_path)
        .map_err(|e| format!("打开 ZIP 失败: {}", e))?;
    let mut archive = zip::ZipArchive::new(zip_file)
        .map_err(|e| format!("解析 ZIP 失败: {}", e))?;

    let mut found_exe = false;
    let mut zip_entries: Vec<String> = Vec::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)
            .map_err(|e| format!("读取 ZIP 条目失败: {}", e))?;
        let name = file.name().to_string();
        zip_entries.push(name.clone());

        let file_name = std::path::Path::new(&name)
            .file_name()
            .map(|f| f.to_string_lossy().to_string())
            .unwrap_or_default();
        if !file.is_dir()
            && file_name.to_lowercase().contains("sumatrapdf")
            && file_name.to_lowercase().ends_with(".exe")
        {
            let mut out_file = std::fs::File::create(&dest)
                .map_err(|e| format!("创建 SumatraPDF.exe 失败: {}", e))?;
            std::io::copy(&mut file, &mut out_file)
                .map_err(|e| format!("解压失败: {}", e))?;
            found_exe = true;
            break;
        }
    }

    std::fs::remove_file(&zip_path).ok();

    if !found_exe {
        log::warn!("ZIP entries: {:?}", zip_entries);
        return Err(format!("ZIP 中未找到 SumatraPDF.exe，包含文件: {}", zip_entries.join(", ")));
    }

    log::info!("SumatraPDF downloaded to: {}", dest.display());

    Ok(pdf_engine::PdfResult {
        success: true,
        message: format!("SumatraPDF 已下载到: {}", dest.display()),
        pdf_path: Some(dest.to_string_lossy().to_string()),
        warnings: None,
    })
}

/// Print an existing PDF file using SumatraPDF CLI
#[command]
fn sumatrapdf_print(
    pdf_path: String,
    printer_name: Option<String>,
    copies: Option<u32>,
    duplex: Option<bool>,
    color_mode: Option<String>,
    fit_mode: Option<String>,
    paper_w: Option<f32>,
    paper_h: Option<f32>,
) -> Result<pdf_engine::PdfResult, String> {
    let output = std::path::Path::new(&pdf_path);
    if !output.exists() {
        return Err("PDF文件不存在".to_string());
    }

    let sumatra = pdf_engine::find_sumatrapdf()
        .ok_or("未检测到 SumatraPDF。请安装 SumatraPDF 或切换到「PDF阅读器」模式。".to_string())?;

    let resolved_printer = match printer_name {
        Some(name) if !name.is_empty() => name,
        _ => pdf_engine::get_default_printer_name()
            .ok_or("未找到默认打印机".to_string())?,
    };

    let settings_str = pdf_engine::build_sumatra_print_settings(
        copies.unwrap_or(1),
        duplex.unwrap_or(false),
        &color_mode.unwrap_or_default(),
        &fit_mode.unwrap_or_default(),
        paper_w,
        paper_h,
    );

    pdf_engine::print_with_sumatrapdf(&sumatra.path, output, &resolved_printer, &settings_str)?;

    Ok(pdf_engine::PdfResult {
        success: true,
        message: format!("已发送到打印机「{}」", resolved_printer),
        pdf_path: Some(pdf_path),
        warnings: None,
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

/// Print a PDF file silently. Uses ShellExecute to delegate to PDF reader.
/// Strategy 1: ShellExecuteW "printto" — specify printer, silent print
/// Strategy 2: ShellExecuteW "print" — use default printer, may show dialog
/// Strategy 3: ShellExecuteW "open" — fallback: open PDF for manual printing
/// Returns Ok(true) if printed, Ok(false) if opened as fallback.
#[cfg(target_os = "windows")]
fn shell_execute_print(pdf_path: &std::path::Path, printer_name: Option<&str>) -> Result<bool, String> {
    use windows::core::{HSTRING, PCWSTR};
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::{SW_HIDE, SW_SHOWNORMAL, SW_SHOW};

    let resolved_printer: Option<String> = match printer_name {
        Some(name) => Some(name.to_string()),
        None => pdf_engine::get_default_printer_name(),
    };
    let printer_str = resolved_printer.as_deref()
        .ok_or("未找到默认打印机，请在系统设置中配置打印机，或在打印设置中手动选择。")?;

    let _com = ComGuard::init();
    unsafe {
        let file: HSTRING = pdf_path.to_string_lossy().to_string().into();

        // Strategy 1: ShellExecuteW "printto" — specify printer, silent
        let verb: HSTRING = "printto".into();
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
        if ret.0 as isize > 32 {
            return Ok(true);
        }
        log::warn!("ShellExecute printto failed (code: {}), trying simple print", ret.0 as isize);

        // Strategy 2: ShellExecuteW "print" without specifying printer
        let verb: HSTRING = "print".into();
        let ret = ShellExecuteW(
            None,
            &verb,
            &file,
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOW,
        );
        if ret.0 as isize > 32 {
            return Ok(true);
        }
        log::warn!("ShellExecute print failed (code: {}), falling back to open", ret.0 as isize);

        // Strategy 3: Fallback — open the PDF so user can print manually
        let verb: HSTRING = "open".into();
        let ret = ShellExecuteW(
            None,
            &verb,
            &file,
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        );
        if ret.0 as isize > 32 {
            return Ok(false);
        }

        return Err(format!(
            "打印失败，错误码: {}。请尝试以下解决方法：\n1. 检查打印机是否正常连接\n2. 在打印面板中选择\"打开PDF\"模式手动打印\n3. 安装PDF阅读器（如Adobe Reader）",
            ret.0 as isize
        ));
    }
}


// =====================================================
// App Entry
// =====================================================

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    #[cfg(target_os = "windows")]
    {
        if let Err(err) = check_windows_version() {
            show_error_dialog(&err);
            std::process::exit(1);
        }
    }

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
                        tauri::WindowEvent::DragDrop(drop_event) => match drop_event {
                            // 拖入提示浮层：进入显示 / 离开隐藏 / 松手隐藏后处理文件
                            tauri::DragDropEvent::Enter { .. } => {
                                let _ = win.eval("if(window._tauriDragHover)window._tauriDragHover(true)");
                            }
                            tauri::DragDropEvent::Leave => {
                                let _ = win.eval("if(window._tauriDragHover)window._tauriDragHover(false)");
                            }
                            tauri::DragDropEvent::Drop { paths, .. } => {
                                let _ = win.eval("if(window._tauriDragHover)window._tauriDragHover(false)");
                                let valid: Vec<String> = paths.iter()
                                    .filter_map(|p| {
                                        let valid_ext = p.extension()
                                            .and_then(|e| e.to_str())
                                            .map(|e| ["pdf", "jpg", "jpeg", "png", "bmp", "webp", "tiff", "tif", "ofd", "xml"].contains(&e.to_lowercase().as_str()))
                                            .unwrap_or(false);
                                        if valid_ext { Some(p.to_string_lossy().to_string()) } else { None }
                                    })
                                    .collect();
                                let json = serde_json::to_string(&valid).unwrap_or_default();
                                let js = format!("if(window._tauriFileDrop)window._tauriFileDrop({})", json);
                                let _ = win.eval(&js);
                            }
                            _ => {}
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
        render_pdf_pages_pdfium,
        check_winrt_pdf,
        render_and_ocr_pdf,
        open_url,
        open_file,
        copy_file,
        rename_file,
        write_text_file,
        check_path_exists,
        get_downloads_dir,
        ocr_image,
        ocr_pdf_page,
        check_ocr_available,
        extract_pdf_text,
        extract_pdf_texts,
        get_app_version,
        get_config,
        get_temp_dir,
        show_window,
        trim_image,
        enhance_image,
        generate_pdf_from_layout,
        print_pdf_file,
        parse_ofd,
        parse_xml_invoice,
        open_ofd_images,
        check_sumatrapdf_available,
        check_pdfium_available,
        cancel_download,
        pdfium_vector_print,
        pdfium_print_pdf,
        download_pdfium_dll,
        download_sumatrapdf,
        sumatrapdf_print,
        check_for_updates,
    ]);

    #[cfg(not(feature = "ocr"))]
    let builder = builder.invoke_handler(tauri::generate_handler![
        open_invoice_files,
        get_printers,
        render_pdf_pages,
        render_pdf_pages_pdfium,
        check_winrt_pdf,
        open_url,
        open_file,
        copy_file,
        rename_file,
        write_text_file,
        check_path_exists,
        get_downloads_dir,
        check_ocr_available,
        extract_pdf_text,
        extract_pdf_texts,
        get_app_version,
        get_config,
        get_temp_dir,
        show_window,
        trim_image,
        enhance_image,
        generate_pdf_from_layout,
        print_pdf_file,
        parse_ofd,
        parse_xml_invoice,
        open_ofd_images,
        check_sumatrapdf_available,
        check_pdfium_available,
        cancel_download,
        pdfium_vector_print,
        pdfium_print_pdf,
        download_pdfium_dll,
        download_sumatrapdf,
        sumatrapdf_print,
        check_for_updates,
    ]);

    builder
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
