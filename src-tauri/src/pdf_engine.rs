use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::BufWriter;
use std::sync::atomic::{AtomicBool, Ordering};

/// Rendering DPI — must match frontend PDF_RENDER_DPI constant
pub const RENDER_DPI: u32 = 300;

/// Global shutdown flag — checked by long-running COM operations to abort early.
/// Set to true when the user clicks the close button, before TerminateProcess.
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

/// Request from frontend: each page is a rendered base64 image
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfRequest {
    /// Rendered page images as base64 data URLs
    pub pages: Vec<String>,
    /// Paper width in mm
    pub paper_w: f32,
    /// Paper height in mm
    pub paper_h: f32,
}

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

        reader.Close().ok();

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

    Ok(results)
}

// =====================================================
// Read files from disk
// =====================================================

pub fn read_invoice_files(paths: Vec<String>) -> Result<Vec<FileData>, String> {
    let mut results = Vec::new();
    for path_str in paths {
        let path = std::path::Path::new(&path_str);
        if !path.exists() {
            continue;
        }

        let name = path.file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_else(|| "unknown".to_string());

        let ext = path.extension()
            .map(|e| e.to_string_lossy().to_lowercase())
            .unwrap_or_default();

        // Only accept supported formats (including OFD)
        if !["pdf", "jpg", "jpeg", "png", "bmp", "webp", "tiff", "tif", "ofd"].contains(&ext.as_str()) {
            continue;
        }

        let metadata = path.metadata().map_err(|e| format!("读取文件信息失败: {}", e))?;
        let size = metadata.len();

        // OFD: extract embedded images from the ZIP archive
        if ext == "ofd" {
            match extract_ofd_images(&path_str) {
                Ok(images) => {
                    for (idx, (img_data_url, img_ext)) in images.iter().enumerate() {
                        // Remove .ofd/.OFD extension from name (case-insensitive)
                        let base_name = if name.to_uppercase().ends_with(".OFD") {
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
                            path: None, // no filePath needed, already converted to image
                        });
                    }
                    continue;
                }
                Err(e) => {
                    log::warn!("OFD extraction failed for {}: {}", name, e);
                    continue;
                }
            }
        }

        let bytes = std::fs::read(path).map_err(|e| format!("读取文件失败 {}: {}", name, e))?;

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

        results.push(FileData {
            name,
            ext,
            size,
            data_url,
            path: Some(path_str),
        });
    }
    Ok(results)
}

// =====================================================
// PDF Generation from page images
// =====================================================

pub fn generate_pdf_from_pages(request: &PdfRequest, output_path: &std::path::Path) -> Result<(), String> {
    use printpdf::*;

    let paper_w = Mm(request.paper_w);
    let paper_h = Mm(request.paper_h);

    if request.pages.is_empty() {
        return Err("没有页面数据".to_string());
    }

    // Decode first page image
    let first_img = decode_base64_image(&request.pages[0])?;

    let (doc, page1_idx, layer1_idx) = PdfDocument::new(
        "发票打印",
        paper_w,
        paper_h,
        "Layer 1",
    );

    // Place first image
    add_image_to_layer(&doc.get_page(page1_idx).get_layer(layer1_idx), &first_img, request.paper_w, request.paper_h)?;

    // Add remaining pages
    for (i, page_data) in request.pages.iter().skip(1).enumerate() {
        let img = decode_base64_image(page_data)?;
        let (page_idx, layer_idx) = doc.add_page(paper_w, paper_h, &format!("Layer {}", i + 2));
        add_image_to_layer(&doc.get_page(page_idx).get_layer(layer_idx), &img, request.paper_w, request.paper_h)?;
    }

    // Save
    let mut out = BufWriter::new(File::create(output_path).map_err(|e| format!("创建文件失败: {}", e))?);
    doc.save(&mut out).map_err(|e| format!("保存PDF失败: {}", e))?;

    Ok(())
}

fn add_image_to_layer(
    layer: &printpdf::PdfLayerReference,
    img: &image::DynamicImage,
    paper_w_mm: f32,
    paper_h_mm: f32,
) -> Result<(), String> {
    use printpdf::*;

    let image = Image::from_dynamic_image(img);

    // Scale image to fill the page
    // Frontend renders at RENDER_DPI, convert to mm
    let img_w_mm = img.width() as f32 * 25.4 / RENDER_DPI as f32;
    let img_h_mm = img.height() as f32 * 25.4 / RENDER_DPI as f32;

    let scale_x = paper_w_mm / img_w_mm;
    let scale_y = paper_h_mm / img_h_mm;

    image.add_to_layer(
        layer.clone(),
        ImageTransform {
            translate_x: Some(Mm(0.0)),
            translate_y: Some(Mm(0.0)),
            scale_x: Some(scale_x),
            scale_y: Some(scale_y),
            rotate: None,
            dpi: Some(RENDER_DPI as f32),
        },
    );

    Ok(())
}

// =====================================================
// List Printers (Windows)
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

fn decode_base64_image(data_url: &str) -> Result<image::DynamicImage, String> {
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

/// OCR an image from base64 data URL, return recognized text
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

    // Create OCR engine with user profile languages (includes Chinese on Chinese Windows)
    let engine = OcrEngine::TryCreateFromUserProfileLanguages()
        .map_err(|e| format!("创建OCR引擎失败: {}（请确保系统已安装中文语言包）", e))?;

    // Run OCR
    let result = engine.RecognizeAsync(&bitmap)
        .map_err(|e| format!("创建OCR识别操作失败: {}", e))?
        .get()
        .map_err(|e| format!("OCR识别失败: {}", e))?;

    let text = result.Text()
        .map_err(|e| format!("获取OCR文本失败: {}", e))?
        .to_string();

    log::info!("OCR recognized {} chars", text.len());

    Ok(text)
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
    let cropped: image::ImageBuffer<image::Rgba<u8>, Vec<u8>> =
        rgba.crop_imm(left, top, cw, ch).to_image();
    image::DynamicImage::ImageRgba8(cropped)
}

// =====================================================
// Layout Rendering (JS canvas → Rust)
// =====================================================

use serde::{Deserialize, Serialize};

/// Settings for layout rendering — mirrors JS getSettings() output.
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
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
#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
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

/// Decode a base64 data URL to DynamicImage.
fn decode_file_image(data_url: &str) -> Result<image::DynamicImage, String> {
    decode_base64_image(data_url)
}

/// Apply grayscale or B&W conversion to an image.
fn apply_color_mode(img: image::DynamicImage, mode: &str) -> image::DynamicImage {
    match mode {
        "grayscale" => {
            let gray = img.to_luma8();
            let rgba = image::ImageBuffer::from_fn(gray.width(), gray.height(), |x, y| {
                let p = gray.get_pixel(x, y);
                image::Rgba([p[0], p[0], p[0], 255])
            });
            image::DynamicImage::ImageRgba8(rgba)
        }
        "bw" => {
            let gray = img.to_luma8();
            let rgba = image::ImageBuffer::from_fn(gray.width(), gray.height(), |x, y| {
                let p = gray.get_pixel(x, y);
                let v = if p[0] > 128 { 255 } else { 0 };
                image::Rgba([v, v, v, 255])
            });
            image::DynamicImage::ImageRgba8(rgba)
        }
        _ => img,
    }
}

/// Render a single page: add all slot images to the PDF page.
fn render_one_page(
    doc: &printpdf::PdfDocument,
    page_idx: printpdf::indices::PdfPageIndex,
    layer_idx: printpdf::indices::PdfLayerIndex,
    page_spec: &PageSpec,
    files: &[FileSpec],
    settings: &RenderSettings,
    slot_positions: &[LayoutSlotMm],
) -> Result<(), String> {
    use printpdf::*;

    let layer = doc.get_page(page_idx).get_layer(layer_idx);

    for (slot_idx, slot_spec) in page_spec.slots.iter().enumerate() {
        let file_idx = match slot_spec.file_index {
            Some(idx) if idx < files.len() => idx,
            _ => continue,
        };
        let file_spec = &files[file_idx];

        // Decode image
        let mut img = decode_file_image(&file_spec.data_url)?;
        let rot = slot_spec.rotation;

        // Apply trim
        if settings.trim_white.unwrap_or(false) {
            img = trim_white_edges(&img, 245);
        }

        // Apply rotation (90° multiples)
        img = match ((rot % 360) + 360) % 360 {
            90  => image::DynamicImage::ImageRgba8(image::imageops::rotate90(&img.to_rgba8())),
            180 => image::DynamicImage::ImageRgba8(image::imageops::rotate180(&img.to_rgba8())),
            270 => image::DynamicImage::ImageRgba8(image::imageops::rotate270(&img.to_rgba8())),
            _   => img,
        };

        // Apply color mode
        img = apply_color_mode(img, &settings.color_mode);

        let (iw, ih) = (img.width(), img.height());

        // Compute draw dimensions in mm at RENDER_DPI
        let iw_mm = iw as f32 * 25.4 / RENDER_DPI as f32;
        let ih_mm = ih as f32 * 25.4 / RENDER_DPI as f32;

        // Compute scale to fit in slot
        let (scale_x, scale_y) = match settings.fit_mode.as_str() {
            "fill" => {
                let sx = slot_positions[slot_idx].w_mm / iw_mm;
                let sy = slot_positions[slot_idx].h_mm / ih_mm;
                (sx, sy)
            }
            "original" => (1.0, 1.0),
            "custom" => {
                let ref_dim = slot_positions[slot_idx].w_mm.min(slot_positions[slot_idx].h_mm);
                let s = ref_dim / iw_mm * settings.custom_scale;
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

        // The image's natural size in PDF at RENDER_DPI:
        //   iw_px → iw_mm = iw * 25.4 / RENDER_DPI
        // We want to scale it to draw_w_mm:
        //   scale_for_pdf = draw_w_mm / iw_mm
        // But printpdf applies: natural_size * dpi_scale * transform_scale
        // where natural_size at dpi=D: DPI_inches * 25.4 mm
        // Actually, printpdf's ImageTransform with dpi=Some(d):
        //   rendered_width_mm = img_width_px / d * 25.4 * scale_x
        // So to get draw_w_mm: scale_x = draw_w_mm / (iw as f32 / RENDER_DPI as f32 * 25.4)
        // Which is exactly: scale_x = draw_w_mm / iw_mm
        // And we already computed that!

        let pdf_image = Image::from_dynamic_image(&img);
        pdf_image.add_to_layer(
            layer.clone(),
            ImageTransform {
                translate_x: Some(Mm(offset_x_mm)),
                translate_y: Some(Mm(offset_y_mm)),
                scale_x: Some(scale_x),
                scale_y: Some(scale_y),
                rotate: None,
                dpi: Some(RENDER_DPI as f32),
            },
        );

        // Number badge
        if settings.number {
            let num_str = (slot_idx + 1).to_string();
            // Position at top-right of slot (convert to bottom-left)
            let num_x_mm = slot_positions[slot_idx].x_mm + slot_positions[slot_idx].w_mm - 5.0;
            let num_y_mm = slot_positions[slot_idx].y_mm + slot_positions[slot_idx].h_mm - 3.0;
            layer.use_text(
                num_str,
                11.0,
                Mm(num_x_mm),
                Mm(num_y_mm),
                printpdf::IndirectFontRef::default(),
            );
        }
    }

    // Watermark
    if settings.watermark {
        if let Some(ref text) = settings.watermark_text {
            let center_x = settings.paper_w / 2.0;
            let center_y = settings.paper_h / 2.0;
            // Note: printpdf's use_text doesn't support rotation easily.
            // For rotated text, we'd need to use a transformation matrix or
            // accept un-rotated watermark text.
            // For now, add un-rotated text at center.
            layer.use_text(
                text.clone(),
                48.0,
                Mm(center_x - 40.0),
                Mm(center_y),
                printpdf::IndirectFontRef::default(),
            );
        }
    }

    Ok(())
}

/// Generate PDF from layout request (files + pages + settings).
/// This replaces JS `renderPageToCanvas` + `generate_pdf_from_pages`.
pub fn generate_pdf_from_layout(
    request: &LayoutRenderRequest,
    output_path: &std::path::Path,
) -> Result<(), String> {
    use printpdf::*;

    if request.pages.is_empty() {
        return Err("没有页面数据".to_string());
    }

    let (slot_positions, pw, ph) = calculate_layout_mm(&request.settings);

    // Create PDF document
    let (doc, page1_idx, layer1_idx) = PdfDocument::new(
        "发票打印",
        Mm(pw),
        Mm(ph),
        "Layer 1",
    );

    for (page_idx, page_spec) in request.pages.iter().enumerate() {
        let (p_idx, l_idx) = if page_idx == 0 {
            (page1_idx, layer1_idx)
        } else {
            let (pi, li) = doc.add_page(Mm(pw), Mm(ph), &format!("Layer {}", page_idx + 1));
            (pi, li)
        };

        render_one_page(
            &doc,
            p_idx,
            l_idx,
            page_spec,
            &request.files,
            &request.settings,
            &slot_positions,
        )?;
    }

    // Save PDF
    let mut out = std::io::BufWriter::new(
        std::fs::File::create(output_path)
            .map_err(|e| format!("创建文件失败: {}", e))?,
    );
    doc.save(&mut out)
        .map_err(|e| format!("保存PDF失败: {}", e))?;

    Ok(())
}
