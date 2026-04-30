use serde::{Deserialize, Serialize};
use std::sync::atomic::{AtomicBool, Ordering};

/// Rendering DPI — must match frontend PDF_RENDER_DPI constant
pub const RENDER_DPI: u32 = 300;

/// Conversion factor: 1 mm = 2.834646 pt (72 pt per inch / 25.4 mm per inch)
const MM_TO_PT: f32 = 72.0 / 25.4;

/// Global shutdown flag — checked by long-running COM operations to abort early.
/// Set to true when the user clicks the close button, before graceful window close.
pub static SHUTTING_DOWN: AtomicBool = AtomicBool::new(false);

// =====================================================
// COM RAII Guard — ensures CoUninitialize is called on drop
// =====================================================

pub(crate) struct ComGuard;

#[cfg(target_os = "windows")]
impl ComGuard {
    pub(crate) fn init() -> Self {
        unsafe {
            let _ = windows::Win32::System::Com::CoInitializeEx(
                None,
                windows::Win32::System::Com::COINIT_APARTMENTTHREADED,
            );
        }
        ComGuard
    }
}

#[cfg(not(target_os = "windows"))]
impl ComGuard {
    pub(crate) fn init() -> Self {
        ComGuard
    }
}

#[cfg(target_os = "windows")]
impl Drop for ComGuard {
    fn drop(&mut self) {
        unsafe { windows::Win32::System::Com::CoUninitialize(); }
    }
}

// =====================================================
// Types
// =====================================================

/// Result of PDF generation / printing
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfResult {
    pub success: bool,
    pub message: String,
    pub pdf_path: Option<String>,
}

/// Printer info
#[derive(Debug, Serialize)]
pub struct PrinterInfo {
    pub name: String,
    pub is_default: bool,
}

/// File data returned to frontend
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FileData {
    pub name: String,
    pub ext: String,
    pub size: u64,
    /// Base64-encoded file content (data URL format)
    pub data_url: String,
    /// Original file path (for WinRT PDF rendering)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
}

/// Rendered PDF page image
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RenderedPage {
    pub index: u32,
    /// Base64-encoded PNG data URL
    pub image_data_url: String,
    pub width: u32,
    pub height: u32,
    /// Actual DPI used for rendering (may differ from requested DPI due to adaptive scaling)
    pub render_dpi: u32,
}

// =====================================================
// Windows PDF Rendering (WinRT)
// =====================================================

// Note: previously used IBufferByteAccess COM interface, but buffer.cast::<IBufferByteAccess>()
// fails with E_NOINTERFACE (0x80004002). Switched to DataReader which works reliably.

/// Render PDF pages to PNG images using Windows.Data.Pdf API
/// This handles PDFs with system font references that PDF.js cannot render
#[cfg(target_os = "windows")]
pub(crate) fn render_pdf_pages(pdf_path: &str, dpi: u32) -> Result<Vec<RenderedPage>, String> {
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    use windows::core::HSTRING;
    use windows::Data::Pdf::{PdfDocument, PdfPageRenderOptions};
    use windows::Storage::StorageFile;
    use windows::Storage::Streams::{DataReader, InMemoryRandomAccessStream};
    use base64::Engine;

    let _com = ComGuard::init();

    let path_h = HSTRING::from(pdf_path);

    // Load file and document
    let file = StorageFile::GetFileFromPathAsync(&path_h)
        .map_err(|e| format!("创建异步操作失败: {}", e))?
        .get()
        .map_err(|e| format!("加载文件失败: {}", e))?;

    let doc = PdfDocument::LoadFromFileAsync(&file)
        .map_err(|e| format!("创建异步操作失败: {}", e))?
        .get()
        .map_err(|e| format!("加载PDF失败: {}（文件可能受密码保护）", e))?;

    let page_count = doc.PageCount().map_err(|e| format!("获取页数失败: {}", e))?;
    log::info!("WinRT PDF rendering: {} pages, dpi={}", page_count, dpi);

    let mut results = Vec::new();

    for i in 0..page_count {
        // Check shutdown flag frequently so we can abort early
        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，渲染已中止".to_string());
        }
        let page = doc.GetPage(i).map_err(|e| format!("获取第{}页失败: {}", i + 1, e))?;

        // Get page size via Size() which returns Foundation::Size { Width, Height }
        // Size is in device-independent pixels (96 DPI base)
        let size = page.Size().map_err(|e| format!("获取第{}页尺寸失败: {}", i + 1, e))?;
        
        // Adaptive DPI: small PDF pages need higher DPI so rendered pixels
        // are sufficient for A4 print at RENDER_DPI (300)
        // Ensure the longest side has at least MIN_RENDER_PX pixels
        let min_render_px: u32 = 3508; // A4 long side at 300 DPI
        let longest_side = size.Width.max(size.Height) as u32;
        let base_pixels = longest_side * dpi / 96; // pixels at requested DPI
        let effective_dpi = if base_pixels >= min_render_px {
            dpi // already enough pixels
        } else {
            let needed = (min_render_px as f32 * 96.0 / longest_side as f32).ceil() as u32;
            dpi.max(needed).min(1200)
        };
        
        let scale = effective_dpi as f32 / 96.0;
        let dest_w = (size.Width * scale) as u32;
        let dest_h = (size.Height * scale) as u32;

        // Set up render options
        let options = PdfPageRenderOptions::new().map_err(|e| format!("创建渲染选项失败: {}", e))?;
        options.SetDestinationWidth(dest_w).map_err(|e| format!("设置宽度失败: {}", e))?;
        options.SetDestinationHeight(dest_h).map_err(|e| format!("设置高度失败: {}", e))?;

        // Render to stream
        let stream = InMemoryRandomAccessStream::new().map_err(|e| format!("创建流失败: {}", e))?;

        // Check shutdown before starting render
        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，渲染已中止".to_string());
        }

        page.RenderWithOptionsToStreamAsync(&stream, &options)
            .map_err(|e| format!("创建渲染操作失败: {}", e))?
            .get()
            .map_err(|e| format!("渲染第{}页失败: {}", i + 1, e))?;

        // Read stream data using DataReader (IBufferByteAccess cast fails with E_NOINTERFACE)
        let stream_size = stream.Size().map_err(|e| format!("获取流大小失败: {}", e))? as u32;
        stream.Seek(0).map_err(|e| format!("Seek失败: {}", e))?;

        let reader = DataReader::CreateDataReader(&stream)
            .map_err(|e| format!("创建DataReader失败: {}", e))?;

        reader.LoadAsync(stream_size)
            .map_err(|e| format!("创建LoadAsync操作失败: {}", e))?
            .get()
            .map_err(|e| format!("加载第{}页数据失败: {}", i + 1, e))?;

        let mut data = vec![0u8; stream_size as usize];
        reader.ReadBytes(&mut data)
            .map_err(|e| format!("读取第{}页字节失败: {}", i + 1, e))?;

        // Explicitly release per-page COM objects
        drop(reader);
        stream.Close().ok();
        drop(stream);
        drop(page);

        let b64 = base64::engine::general_purpose::STANDARD.encode(&data);
        let data_url = format!("data:image/png;base64,{}", b64);

        results.push(RenderedPage {
            index: i,
            image_data_url: data_url,
            width: dest_w,
            height: dest_h,
            render_dpi: effective_dpi,
        });

        log::info!("Rendered page {} ({}x{}) @ {}dpi", i + 1, dest_w, dest_h, effective_dpi);
    }

    // Explicitly release document-level COM objects before ComGuard drops.
    // PdfDocument doesn't implement IClosable, but PdfPage does (already closed in loop).
    // StorageFile doesn't implement IClosable either.
    drop(doc);
    drop(file);
    // ComGuard (_com) drops here last, calling CoUninitialize()

    Ok(results)
}

// =====================================================
// Read files from disk
// =====================================================

pub fn read_invoice_files(paths: Vec<String>) -> Result<Vec<FileData>, String> {
    use rayon::prelude::*;

    // Filter and validate paths first (fast, no I/O)
    let valid_paths: Vec<(String, String, String, u64)> = paths.iter()
        .filter_map(|path_str| {
            let path = std::path::Path::new(path_str);
            if !path.exists() { return None; }

            let name = path.file_name()
                .map(|n| n.to_string_lossy().to_string())
                .unwrap_or_else(|| "unknown".to_string());

            let ext = path.extension()
                .map(|e| e.to_string_lossy().to_lowercase())
                .unwrap_or_default();

            if !["pdf", "jpg", "jpeg", "png", "bmp", "webp", "tiff", "tif", "ofd"].contains(&ext.as_str()) {
                return None;
            }

            let size = path.metadata().ok()?.len();
            Some((path_str.clone(), name, ext, size))
        })
        .collect();

    // Process OFD files first (sequential, they need ZIP extraction)
    let mut results: Vec<FileData> = Vec::new();
    let mut non_ofd_paths: Vec<(String, String, String, u64)> = Vec::new();

    for (path_str, name, ext, size) in valid_paths {
        if ext == "ofd" {
            match extract_ofd_images(&path_str) {
                Ok(images) => {
                    for (idx, (img_data_url, img_ext)) in images.iter().enumerate() {
                        let base_name = if name.to_uppercase().ends_with(".OFD") && name.len() > 4 {
                            &name[..name.len()-4]
                        } else if name.len() > 4 {
                            &name[..name.len()-4]
                        } else {
                            &name
                        };
                        results.push(FileData {
                            name: if images.len() > 1 {
                                format!("{}_第{}页.ofd", base_name, idx + 1)
                            } else {
                                name.clone()
                            },
                            ext: img_ext.to_string(),
                            size,
                            data_url: img_data_url.clone(),
                            path: None,
                        });
                    }
                }
                Err(e) => {
                    log::warn!("OFD extraction failed for {}: {}", name, e);
                }
            }
        } else {
            non_ofd_paths.push((path_str, name, ext, size));
        }
    }

    // Read and encode non-OFD files in parallel using rayon
    let parallel_results: Vec<FileData> = non_ofd_paths
        .par_iter()
        .filter_map(|(path_str, name, ext, size)| {
            // Read file bytes
            let bytes = std::fs::read(path_str).ok()?;

            let mime = match ext.as_str() {
                "pdf" => "application/pdf",
                "jpg" | "jpeg" => "image/jpeg",
                "png" => "image/png",
                "bmp" => "image/bmp",
                "webp" => "image/webp",
                "tiff" | "tif" => "image/tiff",
                _ => "application/octet-stream",
            };

            use base64::Engine;
            let b64 = base64::engine::general_purpose::STANDARD.encode(&bytes);
            let data_url = format!("data:{};base64,{}", mime, b64);

            Some(FileData {
                name: name.clone(),
                ext: ext.clone(),
                size: *size,
                data_url,
                path: Some(path_str.clone()),
            })
        })
        .collect();

    results.extend(parallel_results);
    Ok(results)
}

// =====================================================
// PDF Generation from layout request (only remaining path)
// =====================================================

#[cfg(target_os = "windows")]
pub fn list_printers() -> Result<Vec<PrinterInfo>, String> {
    use winprint::printer::PrinterDevice;

    // Get printer list from winprint (native API, no encoding issues)
    let devices = PrinterDevice::all()
        .map_err(|e| format!("获取打印机列表失败: {}", e))?;

    // Get default printer name via PowerShell (winprint doesn't expose is_default)
    let default_name = get_default_printer_name();

    Ok(devices.into_iter().map(|d| {
        let name = d.name().to_string();
        let is_default = default_name.as_ref().map_or(false, |dn| dn.eq_ignore_ascii_case(&name));
        PrinterInfo { name, is_default }
    }).collect())
}

/// Get the system default printer name via Win32 API (fast, no PowerShell needed)
#[cfg(target_os = "windows")]
pub fn get_default_printer_name() -> Option<String> {
    use windows::Win32::Graphics::Printing::GetDefaultPrinterW;
    use windows::core::PWSTR;

    unsafe {
        // Step 1: query required buffer size (pass null PWSTR)
        let mut size: u32 = 0;
        let _ = GetDefaultPrinterW(PWSTR::null(), &mut size);
        if size == 0 {
            return None;
        }

        // Step 2: allocate buffer and get the name
        let mut buf = vec![0u16; size as usize];
        let result = GetDefaultPrinterW(PWSTR(buf.as_mut_ptr()), &mut size);
        if result.as_bool() && size > 0 {
            let len = buf.iter().position(|&c| c == 0).unwrap_or(size as usize);
            if len > 0 {
                return Some(String::from_utf16_lossy(&buf[..len]));
            }
        }
        None
    }
}

#[cfg(not(target_os = "windows"))]
pub fn get_default_printer_name() -> Option<String> {
    None
}

#[cfg(not(target_os = "windows"))]
pub fn list_printers() -> Result<Vec<PrinterInfo>, String> {
    Ok(vec![])
}

// =====================================================
// Helpers
// =====================================================

pub(crate) fn decode_base64_image(data_url: &str) -> Result<image::DynamicImage, String> {
    use base64::Engine;

    let base64_data = if data_url.contains(',') {
        data_url.split(',').nth(1).unwrap_or("")
    } else {
        data_url
    };

    let bytes = base64::engine::general_purpose::STANDARD
        .decode(base64_data)
        .map_err(|e| format!("Base64解码失败: {}", e))?;

    image::load_from_memory(&bytes).map_err(|e| format!("图片解码失败: {}", e))
}

// =====================================================
// OCR — Windows.Media.Ocr (lightweight, built-in, Chinese support)
// =====================================================

/// A single OCR word with its bounding rectangle
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct OcrWord {
    pub text: String,
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

/// An OCR line containing words
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct OcrLine {
    pub words: Vec<OcrWord>,
}

/// Structured OCR result with coordinates
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OcrResult {
    /// Flat text (backward compatible)
    pub text: String,
    /// Lines with word-level bounding boxes
    pub lines: Vec<OcrLine>,
    /// Image dimensions in pixels (for coordinate normalization)
    pub img_w: u32,
    pub img_h: u32,
}

/// OCR an image from base64 data URL, return structured result with coordinates
#[cfg(target_os = "windows")]
pub fn ocr_image_from_data(data_url: &str) -> Result<String, String> {
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    use windows::Media::Ocr::OcrEngine;
    use windows::Graphics::Imaging::BitmapDecoder;
    use windows::Storage::Streams::InMemoryRandomAccessStream;
    use base64::Engine;

    let _com = ComGuard::init();

    // Decode base64 data
    let base64_data = if data_url.contains(',') {
        data_url.split(',').nth(1).unwrap_or("")
    } else {
        data_url
    };

    let bytes = base64::engine::general_purpose::STANDARD
        .decode(base64_data)
        .map_err(|e| format!("Base64解码失败: {}", e))?;

    if bytes.is_empty() {
        return Err("图片数据为空".to_string());
    }

    // Create stream from bytes
    let stream = InMemoryRandomAccessStream::new()
        .map_err(|e| format!("创建流失败: {}", e))?;

    // Check shutdown before starting OCR
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭，OCR已中止".to_string());
    }

    let writer = windows::Storage::Streams::DataWriter::CreateDataWriter(&stream)
        .map_err(|e| format!("创建DataWriter失败: {}", e))?;

    writer.WriteBytes(&bytes)
        .map_err(|e| format!("写入数据失败: {}", e))?;

    writer.StoreAsync()
        .map_err(|e| format!("创建Store操作失败: {}", e))?
        .get()
        .map_err(|e| format!("存储数据失败: {}", e))?;

    writer.FlushAsync()
        .map_err(|e| format!("创建Flush操作失败: {}", e))?
        .get()
        .map_err(|e| format!("刷新数据失败: {}", e))?;

    writer.Close().ok();

    stream.Seek(0)
        .map_err(|e| format!("Seek失败: {}", e))?;

    // Decode bitmap from stream
    let decoder = BitmapDecoder::CreateAsync(&stream)
        .map_err(|e| format!("创建解码操作失败: {}", e))?
        .get()
        .map_err(|e| format!("解码图片失败: {}", e))?;

    // Get SoftwareBitmap (no-arg version in windows 0.58)
    let bitmap = decoder.GetSoftwareBitmapAsync()
        .map_err(|e| format!("创建位图操作失败: {}", e))?
        .get()
        .map_err(|e| format!("获取位图失败: {}", e))?;

    // Get image pixel dimensions for coordinate normalization
    let img_w = bitmap.PixelWidth().map_err(|e| format!("获取宽度失败: {}", e))?;
    let img_h = bitmap.PixelHeight().map_err(|e| format!("获取高度失败: {}", e))?;

    // Create OCR engine with user profile languages (includes Chinese on Chinese Windows)
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭，OCR已中止".to_string());
    }
    let engine = OcrEngine::TryCreateFromUserProfileLanguages()
        .map_err(|e| format!("创建OCR引擎失败: {}（请确保系统已安装中文语言包）", e))?;

    // Run OCR
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭，OCR已中止".to_string());
    }
    let result = engine.RecognizeAsync(&bitmap)
        .map_err(|e| format!("创建OCR识别操作失败: {}", e))?
        .get()
        .map_err(|e| format!("OCR识别失败: {}", e))?;

    // --- Extract data from OCR result BEFORE releasing COM objects ---
    let flat_text = result.Text()
        .map_err(|e| format!("获取OCR文本失败: {}", e))?
        .to_string();

    let mut ocr_lines: Vec<OcrLine> = Vec::new();
    let lines = result.Lines()
        .map_err(|e| format!("获取OCR行失败: {}", e))?;

    for line in lines {
        let mut ocr_words: Vec<OcrWord> = Vec::new();
        let words = line.Words()
            .map_err(|e| format!("获取OCR词失败: {}", e))?;

        for word in words {
            let rect = word.BoundingRect()
                .map_err(|e| format!("获取词边界失败: {}", e))?;
            let text = word.Text()
                .map_err(|e| format!("获取词文本失败: {}", e))?
                .to_string();
            ocr_words.push(OcrWord {
                text,
                x: rect.X as f64,
                y: rect.Y as f64,
                w: rect.Width as f64,
                h: rect.Height as f64,
            });
        }

        if !ocr_words.is_empty() {
            ocr_lines.push(OcrLine { words: ocr_words });
        }
    }

    // --- Explicitly release ALL WinRT COM objects in reverse creation order ---
    // This ensures COM resources are freed BEFORE CoUninitialize (via ComGuard drop).
    // Without explicit Close(), WinRT objects only get Release() on Drop, which
    // decrements refcount but may not flush I/O or release OS handles.
    // This is critical for clean process exit — leaked COM objects can prevent
    // the thread from terminating, which blocks the process from exiting.
    drop(result);
    drop(engine);
    bitmap.Close().ok();  // SoftwareBitmap implements IClosable
    drop(bitmap);
    drop(decoder);        // BitmapDecoder does NOT implement IClosable — just drop
    stream.Close().ok();  // InMemoryRandomAccessStream implements IClosable
    drop(stream);
    // ComGuard (_com) drops here last, calling CoUninitialize()

    let ocr_result = OcrResult {
        text: flat_text,
        lines: ocr_lines,
        img_w: img_w as u32,
        img_h: img_h as u32,
    };

    log::info!("OCR recognized {} chars, {} lines", ocr_result.text.len(), ocr_result.lines.len());

    // Return as JSON string
    serde_json::to_string(&ocr_result)
        .map_err(|e| format!("OCR结果序列化失败: {}", e))
}

#[cfg(not(target_os = "windows"))]
pub fn ocr_image_from_data(_data_url: &str) -> Result<String, String> {
    Err("OCR仅支持Windows系统".to_string())
}

// =====================================================
// OFD Format Support
// =====================================================

/// Extract embedded images from an OFD file (Chinese electronic invoice format)
/// OFD is a ZIP archive containing XML page descriptions and image resources.
/// For electronic invoices, the content is typically a full-page image.
fn extract_ofd_images(ofd_path: &str) -> Result<Vec<(String, String)>, String> {
    use base64::Engine;
    use std::io::Read;

    let file = std::fs::File::open(ofd_path)
        .map_err(|e| format!("打开OFD文件失败: {}", e))?;

    let mut archive = zip::ZipArchive::new(file)
        .map_err(|e| format!("解析OFD ZIP失败: {}", e))?;

    // Strategy: find all image files in the archive and return them
    // OFD images are typically in paths like:
    //   - Pages/Page_0/Res/xxx.jpg (per-page resources)
    //   - Res/xxx.jpg (document-level resources)
    //   - DocumentRes/xxx.jpg
    // Common image extensions: jpg, jpeg, png
    let mut image_entries: Vec<String> = Vec::new();

    for i in 0..archive.len() {
        let entry = archive.by_index(i).map_err(|e| format!("读取ZIP条目失败: {}", e))?;
        let name = entry.name().to_string();
        let lower = name.to_lowercase();

        // Look for image files (not in signature or annotation paths)
        if (lower.ends_with(".jpg") || lower.ends_with(".jpeg") || lower.ends_with(".png"))
            && !lower.contains("sign_")
            && !lower.contains("seal_")
        {
            image_entries.push(name);
        }
    }

    if image_entries.is_empty() {
        return Err("OFD文件中未找到图片资源".to_string());
    }

    // Sort entries: prioritize page-ordered paths, then alphabetical
    // OFD pages are typically: Pages/Page_0/Res/..., Pages/Page_1/Res/..., etc.
    fn extract_page_index(path: &str) -> u32 {
        // Look for "Page_N" pattern in path
        let lower = path.to_lowercase();
        if let Some(pos) = lower.find("page_") {
            let rest = &path[pos + 5..];
            let num_str: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
            if let Ok(idx) = num_str.parse::<u32>() {
                return idx;
            }
        }
        u32::MAX // no page index found, sort last
    }
    image_entries.sort_by(|a, b| {
        extract_page_index(a).cmp(&extract_page_index(b)).then(a.cmp(b))
    });

    // Read and encode each image
    let mut results = Vec::new();
    for entry_name in &image_entries {
        let mut entry = archive.by_name(entry_name)
            .map_err(|e| format!("读取OFD图片失败: {}", e))?;
        let mut data = Vec::new();
        entry.read_to_end(&mut data)
            .map_err(|e| format!("读取OFD图片数据失败: {}", e))?;

        // Determine MIME type and extension based on actual image format
        let lower = entry_name.to_lowercase();
        let (mime, img_ext) = if lower.ends_with(".png") {
            ("image/png", "png")
        } else {
            ("image/jpeg", "jpg")
        };

        let b64 = base64::engine::general_purpose::STANDARD.encode(&data);
        let data_url = format!("data:{};base64,{}", mime, b64);
        results.push((data_url, img_ext.to_string()));
    }

    log::info!("OFD extracted {} images from {}", results.len(), ofd_path);
    Ok(results)
}

// =====================================================
// White Edge Trimming
// =====================================================

/// Trim white edges from an image.
/// `threshold`: pixels where R, G, B are all >= threshold are considered "white".
/// Returns the cropped image with 5px padding.
pub fn trim_white_edges(img: &image::DynamicImage, threshold: u8) -> image::DynamicImage {
    let rgba = img.to_rgba8();
    let (w, h) = rgba.dimensions();
    if w == 0 || h == 0 {
        return img.clone();
    }

    // Find top
    let mut top = 0u32;
    'outer: for y in 0..h {
        for x in 0..w {
            let p = rgba.get_pixel(x, y);
            if p[0] < threshold || p[1] < threshold || p[2] < threshold {
                top = y;
                break 'outer;
            }
        }
    }

    // Find bottom
    let mut bottom = h - 1;
    'outer2: for y in (0..h).rev() {
        for x in 0..w {
            let p = rgba.get_pixel(x, y);
            if p[0] < threshold || p[1] < threshold || p[2] < threshold {
                bottom = y;
                break 'outer2;
            }
        }
    }

    // Find left
    let mut left = 0u32;
    'outer3: for x in 0..w {
        for y in top..=bottom {
            let p = rgba.get_pixel(x, y);
            if p[0] < threshold || p[1] < threshold || p[2] < threshold {
                left = x;
                break 'outer3;
            }
        }
    }

    // Find right
    let mut right = w - 1;
    'outer4: for x in (0..w).rev() {
        for y in top..=bottom {
            let p = rgba.get_pixel(x, y);
            if p[0] < threshold || p[1] < threshold || p[2] < threshold {
                right = x;
                break 'outer4;
            }
        }
    }

    if top >= bottom || left >= right {
        return img.clone();
    }

    // Add 5px padding, clamp to image bounds
    let p: u32 = 5;
    let top    = top.saturating_sub(p);
    let left   = left.saturating_sub(p);
    let bottom = (bottom + p).min(h - 1);
    let right  = (right + p).min(w - 1);

    let cw = right - left + 1;
    let ch = bottom - top + 1;
    let cropped = image::imageops::crop_imm(&rgba, left, top, cw, ch);
    image::DynamicImage::from(cropped.to_image())
}

// =====================================================
// Layout Rendering (JS canvas → Rust)
// =====================================================

/// Settings for layout rendering — mirrors JS getSettings() output.
/// Fields used only for deserialization from JS (border/number/watermark rendered in preview only)
/// are allowed to be dead code since they're needed for serde but not used in Rust PDF generation.
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
#[allow(dead_code)]
pub struct RenderSettings {
    pub paper_w: f32,
    pub paper_h: f32,
    pub cols: u32,
    pub rows: u32,
    pub margin_top: f32,
    pub margin_bottom: f32,
    pub margin_left: f32,
    pub margin_right: f32,
    pub gap_h: f32,
    pub gap_v: f32,
    pub fit_mode: String,
    pub custom_scale: f32,
    pub global_rotation: String,
    pub color_mode: String,
    pub border: bool,
    pub number: bool,
    pub cutline: bool,
    pub watermark: bool,
    pub watermark_text: Option<String>,
    pub watermark_color: String,
    pub watermark_opacity: f32,
    pub watermark_angle: f32,
    pub border_width: Option<f32>,
    pub border_color: Option<String>,
    pub trim_white: Option<bool>,
}

/// A file image with its metadata — sent from JS.
/// ow/oh/rotation are used by JS for layout decisions but not directly by Rust
/// (Rust gets rotation from SlotSpec and dimensions from decoded image).
#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
#[allow(dead_code)]
pub struct FileSpec {
    pub data_url: String,
    pub ow: u32,
    pub oh: u32,
    pub rotation: i32,
}

/// A slot on a page — which file (if any) goes here, and its rotation.
#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct SlotSpec {
    pub file_index: Option<usize>,
    pub rotation: i32,
}

/// A page = array of slots.
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageSpec {
    pub slots: Vec<SlotSpec>,
}

/// Full request for layout-based PDF generation.
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutRenderRequest {
    pub files: Vec<FileSpec>,
    pub pages: Vec<PageSpec>,
    pub settings: RenderSettings,
}

/// A layout slot in mm coordinates (bottom-left origin, for printpdf).
struct LayoutSlotMm {
    x_mm: f32,
    y_mm: f32,
    w_mm: f32,
    h_mm: f32,
}

/// Calculate layout slot positions in mm (bottom-left origin for printpdf).
fn calculate_layout_mm(settings: &RenderSettings) -> (Vec<LayoutSlotMm>, f32, f32) {
    let pw = settings.paper_w;
    let ph = settings.paper_h;
    let mt = settings.margin_top;
    let mb = settings.margin_bottom;
    let ml = settings.margin_left;
    let mr = settings.margin_right;
    let gh = settings.gap_h;
    let gv = settings.gap_v;
    let cols = settings.cols as f32;
    let rows = settings.rows as f32;

    let sw = (pw - cols * (ml + mr) - (cols - 1.0) * gh) / cols;
    let sh = (ph - rows * (mt + mb) - (rows - 1.0) * gv) / rows;

    let mut slots = Vec::new();
    for r in 0..settings.rows as usize {
        for c in 0..settings.cols as usize {
            // Convert row from JS (top-down) to printpdf (bottom-up)
            let row_from_bottom = settings.rows as usize - 1 - r;
            let x_mm = ml + c as f32 * (sw + ml + mr + gh);
            let y_mm = mb + row_from_bottom as f32 * (sh + mt + mb + gv);
            slots.push(LayoutSlotMm { x_mm, y_mm, w_mm: sw, h_mm: sh });
        }
    }

    (slots, pw, ph)
}

/// Apply grayscale or B&W conversion to an image.
fn apply_color_mode(img: image::DynamicImage, mode: &str) -> image::DynamicImage {
    match mode {
        "grayscale" => {
            let gray = img.to_luma8();
            image::DynamicImage::from(gray)
        }
        "bw" => {
            let gray = img.to_luma8();
            let bw = image::ImageBuffer::from_fn(gray.width(), gray.height(), |x, y| {
                let p = gray.get_pixel(x, y);
                let v = if p[0] > 128 { 255u8 } else { 0u8 };
                image::Luma([v])
            });
            image::DynamicImage::from(bw)
        }
        _ => img,
    }
}

/// Cached XObject info: image dimensions in mm + registered XObjectId.
struct CachedXobj {
    iw_mm: f32,
    ih_mm: f32,
    xobj_id: printpdf::XObjectId,
}

/// Decode all unique images (base64 → DynamicImage), apply trim + color mode.
/// Rotation is NOT applied here — it's per-slot and handled in build_page_ops.
/// Returns decoded images indexed by file_index.
/// Uses rayon for parallel decoding when multiple files are present.
fn decode_images(
    files: &[FileSpec],
    settings: &RenderSettings,
) -> Vec<Option<image::DynamicImage>> {
    use rayon::prelude::*;

    let trim = settings.trim_white.unwrap_or(false);
    let color_mode = settings.color_mode.clone();

    // Parallel decode — each file is independent
    let decoded: Vec<Option<image::DynamicImage>> = files
        .par_iter()
        .map(|file_spec| {
            // Check shutdown flag — abort image decoding if app is closing
            if SHUTTING_DOWN.load(Ordering::SeqCst) {
                return None;
            }
            let mut img = match decode_base64_image(&file_spec.data_url) {
                Ok(i) => i,
                Err(e) => {
                    log::warn!("Image decode failed: {}", e);
                    return None;
                }
            };

            // Apply trim (global setting, not per-slot)
            if trim {
                img = trim_white_edges(&img, 245);
            }

            // Apply color mode (global setting, not per-slot)
            img = apply_color_mode(img, &color_mode);

            Some(img)
        })
        .collect();

    decoded
}

/// Get or create a cached XObject for (file_index, rotation).
/// Decoded images are rotated, converted to RawImage, and registered once per unique combo.
fn get_cached_xobj(
    doc: &mut printpdf::PdfDocument,
    cache: &mut std::collections::HashMap<(usize, i32), CachedXobj>,
    file_idx: usize,
    rotation: i32,
    decoded: &[Option<image::DynamicImage>],
) -> Option<CachedXobj> {
    let key = (file_idx, rotation);

    if let Some(cached) = cache.get(&key) {
        return Some(CachedXobj {
            iw_mm: cached.iw_mm,
            ih_mm: cached.ih_mm,
            xobj_id: cached.xobj_id.clone(),
        });
    }

    let img = decoded[file_idx].as_ref()?;

    // Apply rotation to a clone of the decoded image
    let rotated = match ((rotation % 360) + 360) % 360 {
        90  => img.rotate90(),
        180 => img.rotate180(),
        270 => img.rotate270(),
        _   => img.clone(),
    };

    let (iw, ih) = (rotated.width(), rotated.height());
    let iw_mm = iw as f32 * 25.4 / RENDER_DPI as f32;
    let ih_mm = ih as f32 * 25.4 / RENDER_DPI as f32;

    let raw_image = match printpdf::RawImage::from_dynamic_image(rotated) {
        Ok(ri) => ri,
        Err(e) => {
            log::warn!("RawImage conversion failed for file {} rot {}: {}", file_idx, rotation, e);
            return None;
        }
    };
    let xobj_id = doc.add_image(&raw_image);

    let cached = CachedXobj { iw_mm, ih_mm, xobj_id: xobj_id.clone() };
    cache.insert(key, cached);

    Some(CachedXobj { iw_mm, ih_mm, xobj_id })
}

/// Build page operations for one page using decoded images + XObject cache.
fn build_page_ops(
    doc: &mut printpdf::PdfDocument,
    page_spec: &PageSpec,
    settings: &RenderSettings,
    slot_positions: &[LayoutSlotMm],
    decoded: &[Option<image::DynamicImage>],
    xobj_cache: &mut std::collections::HashMap<(usize, i32), CachedXobj>,
) -> Vec<printpdf::Op> {
    let mut ops = Vec::new();

    for (slot_idx, slot_spec) in page_spec.slots.iter().enumerate() {
        let file_idx = match slot_spec.file_index {
            Some(idx) if idx < decoded.len() && decoded[idx].is_some() => idx,
            _ => continue,
        };

        let rotation = slot_spec.rotation;
        let cached = match get_cached_xobj(doc, xobj_cache, file_idx, rotation, decoded) {
            Some(c) => c,
            None => continue,
        };

        let iw_mm = cached.iw_mm;
        let ih_mm = cached.ih_mm;

        // Compute scale to fit in slot
        let (scale_x, scale_y) = match settings.fit_mode.as_str() {
            "fill" => {
                let sx = slot_positions[slot_idx].w_mm / iw_mm;
                let sy = slot_positions[slot_idx].h_mm / ih_mm;
                (sx, sy)
            }
            "original" => (1.0, 1.0),
            "custom" => {
                let contain_s = (slot_positions[slot_idx].w_mm / iw_mm)
                    .min(slot_positions[slot_idx].h_mm / ih_mm);
                let s = contain_s * settings.custom_scale;
                (s, s)
            }
            _ => {
                // "contain"
                let s = (slot_positions[slot_idx].w_mm / iw_mm)
                    .min(slot_positions[slot_idx].h_mm / ih_mm);
                (s, s)
            }
        };

        // Centered position in slot (bottom-left origin)
        let draw_w_mm = iw_mm * scale_x;
        let draw_h_mm = ih_mm * scale_y;
        let offset_x_mm = slot_positions[slot_idx].x_mm
            + (slot_positions[slot_idx].w_mm - draw_w_mm) / 2.0;
        let offset_y_mm = slot_positions[slot_idx].y_mm
            + (slot_positions[slot_idx].h_mm - draw_h_mm) / 2.0;

        // Convert mm to Pt — XObjectTransform uses Pt
        let offset_x_pt = offset_x_mm * MM_TO_PT;
        let offset_y_pt = offset_y_mm * MM_TO_PT;

        ops.push(printpdf::Op::UseXobject {
            id: cached.xobj_id.clone(),
            transform: printpdf::XObjectTransform {
                translate_x: Some(printpdf::Pt(offset_x_pt)),
                translate_y: Some(printpdf::Pt(offset_y_pt)),
                scale_x: Some(scale_x),
                scale_y: Some(scale_y),
                dpi: Some(RENDER_DPI as f32),
                rotate: None,
            },
        });
    }

    ops
}

/// Generate PDF from layout request (files + pages + settings).
/// This replaces JS `renderPageToCanvas` + `generate_pdf_from_pages`.
pub fn generate_pdf_from_layout(
    request: &LayoutRenderRequest,
    output_path: &std::path::Path,
) -> Result<(), String> {
    if request.pages.is_empty() {
        return Err("没有页面数据".to_string());
    }

    let (slot_positions, pw, ph) = calculate_layout_mm(&request.settings);

    // Create PDF document (new API: no page dimensions at creation time)
    let mut doc = printpdf::PdfDocument::new("发票打印");

    // Step 1: Decode all unique images (base64 → DynamicImage), apply trim + color mode.
    // Rotation is per-slot and deferred to build_page_ops for correct (file, rotation) caching.
    let decoded = decode_images(&request.files, &request.settings);

    // Step 2: Build pages, caching XObjects by (file_index, rotation) to avoid redundant work.
    let mut xobj_cache: std::collections::HashMap<(usize, i32), CachedXobj> = std::collections::HashMap::new();

    for page_spec in &request.pages {
        // Check shutdown flag — abort PDF generation if app is closing
        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，PDF生成已中止".to_string());
        }
        let ops = build_page_ops(
            &mut doc,
            page_spec,
            &request.settings,
            &slot_positions,
            &decoded,
            &mut xobj_cache,
        );

        let page = printpdf::PdfPage::new(
            printpdf::Mm(pw),
            printpdf::Mm(ph),
            ops,
        );
        doc.pages.push(page);
    }

    // Save PDF — custom options for print quality
    // Default PdfSaveOptions limits images to 2MB and uses 0.85 JPEG quality,
    // which downsamples high-res invoice images and makes text blurry.
    // We override: no size limit, higher quality, prefer JPEG for speed.
    let save_opts = printpdf::PdfSaveOptions {
        optimize: true,
        subset_fonts: true,
        secure: true,
        image_optimization: Some(printpdf::ImageOptimizationOptions {
            quality: Some(0.95),           // High quality (default 0.85 too lossy for text)
            max_image_size: None,          // NO size limit (default "2MB" downsamples invoices!)
            auto_optimize: Some(true),     // Remove alpha if opaque, detect greyscale
            convert_to_greyscale: None,    // Don't force greyscale
            dither_greyscale: None,
            format: Some(printpdf::ImageCompression::Auto), // Auto: JPEG for photos, Flate for sharp edges
        }),
    };

    let mut warnings = Vec::new();
    let pdf_bytes = doc.save(&save_opts, &mut warnings);

    if !warnings.is_empty() {
        log::warn!("PDF save warnings: {} items", warnings.len());
    }

    std::fs::write(output_path, &pdf_bytes)
        .map_err(|e| format!("写入文件失败: {}", e))?;

    Ok(())
}
