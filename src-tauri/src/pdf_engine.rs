use serde::{Deserialize, Serialize};
use std::sync::atomic::{AtomicBool, Ordering};
use std::io::Read;
use ab_glyph::{Font as AbFont, Glyph, ScaleFont};

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
// JPEG Passthrough Utilities
// =====================================================

/// Check if bytes start with JPEG magic bytes (0xFF 0xD8 0xFF).
fn is_jpeg_bytes(bytes: &[u8]) -> bool {
    bytes.len() >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF
}

/// Parse JPEG dimensions and color component count from SOF marker without full decode.
/// Returns (width, height, num_components) or None if SOF marker not found.
/// Supports SOF0 (baseline), SOF1, SOF2 (progressive), and other SOF variants.
fn parse_jpeg_info(bytes: &[u8]) -> Option<(u32, u32, u8)> {
    let mut i: usize = 0;
    while i + 8 < bytes.len() {
        if bytes[i] != 0xFF { break; }
        let marker = u16::from_be_bytes([bytes[i], bytes[i + 1]]);
        i += 2;

        // SOF markers contain image dimensions and component info
        if (0xFFC0..=0xFFC3).contains(&marker)
            || (0xFFC5..=0xFFC7).contains(&marker)
            || (0xFFC9..=0xFFCB).contains(&marker)
            || (0xFFCD..=0xFFCF).contains(&marker)
        {
            // SOF structure: length(2) + precision(1) + height(2) + width(2) + num_components(1) + ...
            let height = u16::from_be_bytes([bytes[i + 3], bytes[i + 4]]) as u32;
            let width = u16::from_be_bytes([bytes[i + 5], bytes[i + 6]]) as u32;
            let num_components = bytes[i + 7];
            return Some((width, height, num_components));
        }

        // RST markers (0xFFD0-0xFFD7) and SOI (0xFFD8) have no segment length
        if (0xFFD0..=0xFFD9).contains(&marker) {
            continue;
        }

        // SOS marker (0xFFDA): skip entropy-coded data to find next marker
        if marker == 0xFFDA {
            // Read segment length to skip SOS header
            if i + 1 < bytes.len() {
                let seg_len = u16::from_be_bytes([bytes[i], bytes[i + 1]]) as usize;
                i = i.saturating_add(seg_len);
            }
            // Scan for next marker (skip entropy-coded data)
            while i + 1 < bytes.len() {
                if bytes[i] == 0xFF && bytes[i + 1] != 0x00 && !(0xD0..=0xD7).contains(&bytes[i + 1]) {
                    break;
                }
                i += 1;
            }
            continue;
        }

        // All other markers: read segment length and skip
        if i + 1 < bytes.len() {
            let seg_len = u16::from_be_bytes([bytes[i], bytes[i + 1]]) as usize;
            if seg_len < 2 { break; } // malformed
            i = i.saturating_add(seg_len);
        } else {
            break;
        }
    }
    None
}

// =====================================================
// Image Source — tracks how the image was loaded
// =====================================================

/// Image source: tracks whether the image can skip decode-re-encode.
enum ImageSource {
    /// Standard decoded image (current pipeline: decode → RawImage → add_image)
    Decoded(image::DynamicImage),
    /// JPEG passthrough: raw JPEG bytes with known dimensions and color space.
    /// Preserved from read step to avoid re-reading from disk.
    /// At PDF generation time, always decoded → raw pixels → FlateDecode (lossless).
    #[allow(dead_code)]
    JpegPassthrough {
        raw_bytes: Vec<u8>,
        width: u32,
        height: u32,
        /// Number of color components: 1=grayscale, 3=RGB, 4=CMYK
        num_components: u8,
    },
}

// =====================================================
// EXIF Orientation Handling
// =====================================================

/// Read EXIF orientation tag from JPEG bytes.
/// Returns the orientation value (1-8), or 1 (normal) if not found.
/// EXIF orientation values:
///   1 = Normal, 2 = Flipped H, 3 = Rotated 180°, 4 = Flipped V,
///   5 = Transposed, 6 = Rotated 90° CW, 7 = Transverse, 8 = Rotated 90° CCW
fn read_exif_orientation(bytes: &[u8]) -> u32 {
    let mut cursor = std::io::Cursor::new(bytes);
    let reader = exif::Reader::new();
    match reader.read_from_container(&mut cursor) {
        Ok(exif) => {
            if let Some(field) = exif.get_field(exif::Tag::Orientation, exif::In::PRIMARY) {
                if let Some(val) = field.value.get_uint(0) {
                    if val >= 1 && val <= 8 {
                        return val;
                    }
                }
            }
            1 // default: normal
        }
        Err(_) => 1, // no EXIF or parse error — assume normal
    }
}

/// Apply EXIF orientation to an image by rotating/flipping pixels.
/// This bakes the orientation into the pixel data so the image displays
/// correctly without EXIF awareness (e.g., in PDF viewers).
///
/// EXIF orientation values describe where the 0th row/column is in the visual image:
///   1 = Normal (top-left)
///   2 = Flipped horizontally (top-right)
///   3 = Rotated 180° (bottom-right)
///   4 = Flipped vertically (bottom-left)
///   5 = Transposed (left-top)
///   6 = Rotated 90° CW (right-top) — most common for phone photos
///   7 = Transversed (right-bottom)
///   8 = Rotated 90° CCW (left-bottom)
///
/// image crate: rotate90() = 90° CW, rotate270() = 90° CCW
fn apply_exif_orientation(img: image::DynamicImage, orientation: u32) -> image::DynamicImage {
    match orientation {
        1 => img,                            // Normal — no change
        2 => img.fliph(),                    // Flipped horizontally
        3 => img.rotate180(),                // Rotated 180°
        4 => img.flipv(),                    // Flipped vertically
        5 => img.fliph().rotate90(),         // Transposed (flip H + rotate 90° CW)
        6 => img.rotate90(),                 // Rotated 90° CW
        7 => img.fliph().rotate270(),        // Transversed (flip H + rotate 90° CCW)
        8 => img.rotate270(),                // Rotated 90° CCW
        _ => img,
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
    #[serde(skip_serializing_if = "Option::is_none")]
    pub warnings: Option<String>,
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
    /// Base64-encoded preview image (data URL format).
    /// For image files: a JPEG thumbnail (max 600px longest side) for fast IPC.
    /// For PDF files: empty (rendered via render_and_ocr_pdf command).
    /// For OFD files: the extracted page image (no thumbnail — already small).
    pub data_url: String,
    /// Original file path on disk.
    /// Used for: WinRT PDF rendering, OCR via file_path, PDF generation via file_path.
    /// Frontend should store this as fileObj._filePath and pass it to Rust commands
    /// instead of sending the full base64 dataUrl back.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    /// Original image width in pixels (before thumbnail downscaling).
    /// Frontend uses this for layout rotation decisions and PDF generation sizing.
    /// For PDF/OFD files, this is 0 (dimensions come from rendered pages).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub orig_w: Option<u32>,
    /// Original image height in pixels (before thumbnail downscaling).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub orig_h: Option<u32>,
}

/// Rendered PDF page image
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RenderedPage {
    pub index: u32,
    /// Base64-encoded image data URL (PNG or JPEG)
    pub image_data_url: String,
    pub width: u32,
    pub height: u32,
    /// Actual DPI used for rendering (may differ from requested DPI due to adaptive scaling)
    pub render_dpi: u32,
    /// Image format: "png" or "jpeg"
    #[serde(default)]
    pub format: String,
}

/// Rendered PDF page with OCR result — avoids IPC round-trip for OCR.
/// The image is rendered and OCR'd in Rust in a single pass.
#[cfg(feature = "ocr")]
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RenderedOcrPage {
    pub index: u32,
    /// Base64-encoded PNG data URL (for preview)
    pub image_data_url: String,
    pub width: u32,
    pub height: u32,
    pub render_dpi: u32,
    /// OCR result (computed in Rust, no need to send image back for OCR)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ocr_result: Option<OcrResult>,
}

// =====================================================
// Windows PDF Rendering (WinRT)
// =====================================================

// Note: previously used IBufferByteAccess COM interface, but buffer.cast::<IBufferByteAccess>()
// fails with E_NOINTERFACE (0x80004002). Switched to DataReader which works reliably.

/// Render PDF pages to images using Windows.Data.Pdf API
/// This handles PDFs with system font references that PDF.js cannot render
/// - `use_jpeg`: if true, encode as JPEG for smaller size and faster transfer
#[cfg(target_os = "windows")]
pub(crate) fn render_pdf_pages(pdf_path: &str, dpi: u32, use_jpeg: bool) -> Result<Vec<RenderedPage>, String> {
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
    log::info!("WinRT PDF rendering: {} pages, dpi={}, jpeg={}", page_count, dpi, use_jpeg);

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
        
        // For preview, use requested DPI directly without adaptive scaling
        // Adaptive scaling is only needed for print quality output
        let effective_dpi = dpi;
        
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

        let mut png_data = vec![0u8; stream_size as usize];
        reader.ReadBytes(&mut png_data)
            .map_err(|e| format!("读取第{}页字节失败: {}", i + 1, e))?;

        // Explicitly release per-page COM objects
        drop(reader);
        stream.Close().ok();
        drop(stream);
        drop(page);

        // Encode to JPEG if requested
        let (data_url, format) = if use_jpeg {
            let img = image::load_from_memory(&png_data)
                .map_err(|e| format!("解码PNG失败: {}", e))?;
            let mut jpeg_buf = std::io::Cursor::new(Vec::new());
            img.write_to(&mut jpeg_buf, image::ImageFormat::Jpeg)
                .map_err(|e| format!("JPEG编码失败: {}", e))?;
            let b64 = base64::engine::general_purpose::STANDARD.encode(&jpeg_buf.into_inner());
            (format!("data:image/jpeg;base64,{}", b64), "jpeg".to_string())
        } else {
            let b64 = base64::engine::general_purpose::STANDARD.encode(&png_data);
            (format!("data:image/png;base64,{}", b64), "png".to_string())
        };

        log::info!("Rendered page {} ({}x{}) @ {}dpi, format={}", i + 1, dest_w, dest_h, effective_dpi, format);
        
        results.push(RenderedPage {
            index: i,
            image_data_url: data_url,
            width: dest_w,
            height: dest_h,
            render_dpi: effective_dpi,
            format,
        });
    }

    // Explicitly release document-level COM objects before ComGuard drops.
    // PdfDocument doesn't implement IClosable, but PdfPage does (already closed in loop).
    // StorageFile doesn't implement IClosable either.
    drop(doc);
    drop(file);
    // ComGuard (_com) drops here last, calling CoUninitialize()

    Ok(results)
}

pub(crate) fn render_pdf_pages_pdfium(pdf_path: &str, dpi: u32, use_jpeg: bool) -> Result<Vec<RenderedPage>, String> {
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    if crate::pdfium_print::find_pdfium_dll().is_none() {
        return Err("pdfium.dll 不可用，无法使用 PDFium 渲染".to_string());
    }

    let pdf_bytes = std::fs::read(pdf_path)
        .map_err(|e| format!("读取PDF文件失败: {}", e))?;

    let images = crate::pdfium_print::render_pdf_to_images(&pdf_bytes, dpi)?;

    let results: Vec<RenderedPage> = images.into_iter().map(|img| {
        // Convert PNG to JPEG if requested
        if use_jpeg && img.image_data_url.starts_with("data:image/png;base64,") {
            let (data_url, format) = match convert_png_data_url_to_jpeg(&img.image_data_url) {
                Ok((url, fmt)) => (url, fmt),
                Err(e) => {
                    log::warn!("JPEG conversion failed, falling back to PNG: {}", e);
                    (img.image_data_url, "png".to_string())
                }
            };
            RenderedPage {
                index: img.index,
                image_data_url: data_url,
                width: img.width,
                height: img.height,
                render_dpi: img.render_dpi,
                format,
            }
        } else {
            RenderedPage {
                index: img.index,
                image_data_url: img.image_data_url,
                width: img.width,
                height: img.height,
                render_dpi: img.render_dpi,
                format: "png".to_string(),
            }
        }
    }).collect();

    Ok(results)
}

fn convert_png_data_url_to_jpeg(data_url: &str) -> Result<(String, String), String> {
    use base64::Engine;
    if !data_url.starts_with("data:image/png;base64,") {
        return Err("Not a PNG data URL".to_string());
    }
    let base64_data = data_url.strip_prefix("data:image/png;base64,").ok_or("Invalid data URL")?;
    let png_data = base64::engine::general_purpose::STANDARD.decode(base64_data)
        .map_err(|e| format!("Base64 decode failed: {}", e))?;
    
    let img = image::load_from_memory(&png_data)
        .map_err(|e| format!("Image decode failed: {}", e))?;
    let mut jpeg_buf = std::io::Cursor::new(Vec::new());
    img.write_to(&mut jpeg_buf, image::ImageFormat::Jpeg)
        .map_err(|e| format!("JPEG encode failed: {}", e))?;
    
    let b64 = base64::engine::general_purpose::STANDARD.encode(&jpeg_buf.into_inner());
    Ok((format!("data:image/jpeg;base64,{}", b64), "jpeg".to_string()))
}

pub(crate) fn check_winrt_pdf_available() -> bool {
    use windows::core::HSTRING;
    use windows::Storage::StorageFile;

    let _com = ComGuard::init();

    let test_path = std::env::temp_dir().join("_ticketchan_winrt_pdf_test.pdf");
    let test_content = b"%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R>>endobj\nxref\n0 4\n0000000000 65535 f \n0000000009 00000 n \n0000000058 00000 n \n0000000115 00000 n \ntrailer<</Size 4/Root 1 0 R>>\nstartxref\n190\n%%EOF";
    if std::fs::write(&test_path, test_content).is_err() {
        log::warn!("WinRT PDF check: cannot write test file");
        return false;
    }

    let path_h = HSTRING::from(test_path.to_string_lossy().as_ref());
    let result = (|| -> Result<(), String> {
        let file = StorageFile::GetFileFromPathAsync(&path_h)
            .map_err(|e| format!("{}", e))?
            .get()
            .map_err(|e| format!("{}", e))?;
        let doc = windows::Data::Pdf::PdfDocument::LoadFromFileAsync(&file)
            .map_err(|e| format!("{}", e))?
            .get()
            .map_err(|e| format!("{}", e))?;
        let _ = doc.PageCount().map_err(|e| format!("{}", e))?;
        Ok(())
    })();

    let _ = std::fs::remove_file(&test_path);

    match result {
        Ok(()) => {
            log::info!("WinRT PDF component: available");
            true
        }
        Err(e) => {
            log::warn!("WinRT PDF component: NOT available ({})", e);
            false
        }
    }
}

/// Render a single PDF page and run OCR on it — zero IPC round-trip for OCR.
/// The frontend calls this instead of `render_pdf_pages` + `ocr_image` to avoid:
///   Rust render → base64 → IPC → frontend → downsample → base64 → IPC → Rust decode → OCR
/// Instead: Rust render → decode in memory → OCR → return result directly.
/// Returns OcrResult with coordinates in the original (full-DPI) pixel space.
#[cfg(all(target_os = "windows", feature = "ocr"))]
pub(crate) fn ocr_pdf_page(pdf_path: &str, page_index: u32, dpi: Option<u32>, ocr_precision: Option<&str>) -> Result<OcrResult, String> {
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    use windows::core::HSTRING;
    use windows::Data::Pdf::{PdfDocument, PdfPageRenderOptions};
    use windows::Storage::StorageFile;
    use windows::Storage::Streams::{DataReader, InMemoryRandomAccessStream};

    let _com = ComGuard::init();
    let dpi = dpi.unwrap_or(RENDER_DPI);
    let path_h = HSTRING::from(pdf_path);

    let file = StorageFile::GetFileFromPathAsync(&path_h)
        .map_err(|e| format!("创建异步操作失败: {}", e))?
        .get()
        .map_err(|e| format!("加载文件失败: {}", e))?;

    let doc = PdfDocument::LoadFromFileAsync(&file)
        .map_err(|e| format!("创建异步操作失败: {}", e))?
        .get()
        .map_err(|e| format!("加载PDF失败: {}（文件可能受密码保护）", e))?;

    let page_count = doc.PageCount().map_err(|e| format!("获取页数失败: {}", e))?;
    if page_index >= page_count {
        return Err(format!("页码超出范围: 请求第{}页，共{}页", page_index + 1, page_count));
    }

    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    let page = doc.GetPage(page_index).map_err(|e| format!("获取第{}页失败: {}", page_index + 1, e))?;

    let size = page.Size().map_err(|e| format!("获取第{}页尺寸失败: {}", page_index + 1, e))?;

    // Adaptive DPI (same logic as render_pdf_pages)
    let min_render_px: u32 = 3508;
    let longest_side = size.Width.max(size.Height) as u32;
    let base_pixels = longest_side * dpi / 96;
    let effective_dpi = if base_pixels >= min_render_px {
        dpi
    } else {
        let needed = (min_render_px as f32 * 96.0 / longest_side as f32).ceil() as u32;
        dpi.max(needed).min(1200)
    };

    let scale = effective_dpi as f32 / 96.0;
    let dest_w = (size.Width * scale) as u32;
    let dest_h = (size.Height * scale) as u32;

    let options = PdfPageRenderOptions::new().map_err(|e| format!("创建渲染选项失败: {}", e))?;
    options.SetDestinationWidth(dest_w).map_err(|e| format!("设置宽度失败: {}", e))?;
    options.SetDestinationHeight(dest_h).map_err(|e| format!("设置高度失败: {}", e))?;

    let stream = InMemoryRandomAccessStream::new().map_err(|e| format!("创建流失败: {}", e))?;

    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    page.RenderWithOptionsToStreamAsync(&stream, &options)
        .map_err(|e| format!("创建渲染操作失败: {}", e))?
        .get()
        .map_err(|e| format!("渲染第{}页失败: {}", page_index + 1, e))?;

    let stream_size = stream.Size().map_err(|e| format!("获取流大小失败: {}", e))? as u32;
    stream.Seek(0).map_err(|e| format!("Seek失败: {}", e))?;

    let reader = DataReader::CreateDataReader(&stream)
        .map_err(|e| format!("创建DataReader失败: {}", e))?;

    reader.LoadAsync(stream_size)
        .map_err(|e| format!("创建LoadAsync操作失败: {}", e))?
        .get()
        .map_err(|e| format!("加载第{}页数据失败: {}", page_index + 1, e))?;

    let mut data = vec![0u8; stream_size as usize];
    reader.ReadBytes(&mut data)
        .map_err(|e| format!("读取第{}页字节失败: {}", page_index + 1, e))?;

    // Release per-page COM objects
    drop(reader);
    stream.Close().ok();
    drop(stream);
    drop(page);
    drop(doc);
    drop(file);
    // ComGuard (_com) drops at end of scope

    // Decode PNG bytes in memory — no base64 round-trip!
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    let img = image::load_from_memory(&data)
        .map_err(|e| format!("图片解码失败: {}", e))?;

    log::info!("ocr_pdf_page: page {} ({}x{}) decoded, running OCR", page_index + 1, img.width(), img.height());

    let max_dim = ocr_max_dim_for_precision(ocr_precision.unwrap_or("standard"));
    run_ocr_on_image(img, max_dim)
}

/// Render PDF pages and run OCR in one pass — avoids the IPC round-trip
/// where the frontend sends the rendered dataUrl back to Rust for OCR.
/// The image is decoded from PNG bytes ONCE, OCR'd, then base64-encoded for preview.
#[cfg(all(target_os = "windows", feature = "ocr"))]
pub(crate) fn render_and_ocr_pdf(pdf_path: &str, dpi: u32, ocr_precision: Option<&str>) -> Result<Vec<RenderedOcrPage>, String> {
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }
    use windows::core::HSTRING;
    use windows::Data::Pdf::{PdfDocument, PdfPageRenderOptions};
    use windows::Storage::StorageFile;
    use windows::Storage::Streams::{DataReader, InMemoryRandomAccessStream};
    use base64::Engine;
    use std::time::Instant;

    let _com = ComGuard::init();

    let max_dim = ocr_max_dim_for_precision(ocr_precision.unwrap_or("standard"));

    let path_h = HSTRING::from(pdf_path);

    let file = StorageFile::GetFileFromPathAsync(&path_h)
        .map_err(|e| format!("创建异步操作失败: {}", e))?
        .get()
        .map_err(|e| format!("加载文件失败: {}", e))?;

    let doc = PdfDocument::LoadFromFileAsync(&file)
        .map_err(|e| format!("创建异步操作失败: {}", e))?
        .get()
        .map_err(|e| format!("加载PDF失败: {}（文件可能受密码保护）", e))?;

    let page_count = doc.PageCount().map_err(|e| format!("获取页数失败: {}", e))?;
    log::info!("WinRT PDF render+OCR: {} pages, dpi={}", page_count, dpi);

    let mut results = Vec::new();

    for i in 0..page_count {
        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，渲染已中止".to_string());
        }
        let page = doc.GetPage(i).map_err(|e| format!("获取第{}页失败: {}", i + 1, e))?;

        let size = page.Size().map_err(|e| format!("获取第{}页尺寸失败: {}", i + 1, e))?;

        // Adaptive DPI (same logic as render_pdf_pages)
        let min_render_px: u32 = 3508;
        let longest_side = size.Width.max(size.Height) as u32;
        let base_pixels = longest_side * dpi / 96;
        let effective_dpi = if base_pixels >= min_render_px {
            dpi
        } else {
            let needed = (min_render_px as f32 * 96.0 / longest_side as f32).ceil() as u32;
            dpi.max(needed).min(1200)
        };

        let scale = effective_dpi as f32 / 96.0;
        let dest_w = (size.Width * scale) as u32;
        let dest_h = (size.Height * scale) as u32;

        let options = PdfPageRenderOptions::new().map_err(|e| format!("创建渲染选项失败: {}", e))?;
        options.SetDestinationWidth(dest_w).map_err(|e| format!("设置宽度失败: {}", e))?;
        options.SetDestinationHeight(dest_h).map_err(|e| format!("设置高度失败: {}", e))?;

        let stream = InMemoryRandomAccessStream::new().map_err(|e| format!("创建流失败: {}", e))?;

        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，渲染已中止".to_string());
        }

        page.RenderWithOptionsToStreamAsync(&stream, &options)
            .map_err(|e| format!("创建渲染操作失败: {}", e))?
            .get()
            .map_err(|e| format!("渲染第{}页失败: {}", i + 1, e))?;

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

        // Release per-page COM objects
        drop(reader);
        stream.Close().ok();
        drop(stream);
        drop(page);

        // === OCR on raw PNG bytes (no base64 round-trip!) ===
        let t_ocr_start = Instant::now();
        let ocr_result = if !SHUTTING_DOWN.load(Ordering::SeqCst) {
            // Decode image once for OCR
            match image::load_from_memory(&data) {
                Ok(img) => {
                    let orig_w = img.width();
                    let orig_h = img.height();
                    let longest = orig_w.max(orig_h);

                    let ocr_img = if longest > max_dim {
                        let rscale = max_dim as f32 / longest as f32;
                        let nw = (orig_w as f32 * rscale).round() as u32;
                        let nh = (orig_h as f32 * rscale).round() as u32;
                        img.resize_exact(nw, nh, image::imageops::FilterType::Lanczos3)
                    } else {
                        img
                    };

                    // Enhance contrast for better OCR accuracy
                    let ocr_img = enhance_contrast_ocr(ocr_img);

                    let resized_w = ocr_img.width();
                    let resized_h = ocr_img.height();

                    // Run OCR
                    match get_ocr_engine() {
                        Ok(lock) => {
                            let engine = lock.as_ref();
                            match engine {
                                Some(eng) => {
                                    match eng.recognize(&ocr_img) {
                                        Ok(rec_results) => {
                                            let coord_scale_x = if resized_w > 0 { orig_w as f64 / resized_w as f64 } else { 1.0 };
                                            let coord_scale_y = if resized_h > 0 { orig_h as f64 / resized_h as f64 } else { 1.0 };

                                            let mut ocr_lines: Vec<OcrLine> = Vec::new();
                                            let mut flat_text_parts: Vec<String> = Vec::new();

                                            for result in &rec_results {
                                                let line_text = result.text.trim().to_string();
                                                if line_text.is_empty() { continue; }
                                                flat_text_parts.push(line_text.clone());

                                                let bbox = &result.bbox;
                                                let rect = bbox.rect;
                                                let bx = rect.left() as f64 * coord_scale_x;
                                                let by = rect.top() as f64 * coord_scale_y;
                                                let bw = (rect.right() - rect.left()) as f64 * coord_scale_x;
                                                let bh = (rect.bottom() - rect.top()) as f64 * coord_scale_y;

                                                let line_points = bbox.points.as_ref().map(|pts| {
                                                    pts.iter().map(|p| OcrPoint {
                                                        x: p.x as f64 * coord_scale_x,
                                                        y: p.y as f64 * coord_scale_y,
                                                    }).collect()
                                                });

                                                let tokens = split_line_to_words(&line_text);
                                                let line_confidence = result.confidence;

                                                if tokens.is_empty() {
                                                    ocr_lines.push(OcrLine {
                                                        words: vec![OcrWord { text: line_text, x: bx, y: by, w: bw, h: bh }],
                                                        points: line_points,
                                                        confidence: line_confidence,
                                                    });
                                                    continue;
                                                }

                                                let total_weight: f64 = tokens.iter().map(|t| token_width_weight(t)).sum();
                                                let mut words: Vec<OcrWord> = Vec::new();
                                                let mut x_offset = 0.0f64;
                                                for token in &tokens {
                                                    let token_w = if total_weight > 0.0 { bw * token_width_weight(token) / total_weight } else { bw };
                                                    words.push(OcrWord { text: token.clone(), x: bx + x_offset, y: by, w: token_w, h: bh });
                                                    x_offset += token_w;
                                                }
                                                ocr_lines.push(OcrLine { words, points: line_points, confidence: line_confidence });
                                            }

                                            drop(lock); // release engine lock ASAP

                                            let flat_text = flat_text_parts.join("\n");
                                            Some(OcrResult { text: flat_text, lines: ocr_lines, img_w: orig_w, img_h: orig_h })
                                        }
                                        Err(e) => {
                                            log::warn!("PDF页{} OCR识别失败: {:?}", i + 1, e);
                                            None
                                        }
                                    }
                                }
                                None => None,
                            }
                        }
                        Err(e) => {
                            log::warn!("PDF页{} 获取OCR引擎失败: {}", i + 1, e);
                            None
                        }
                    }
                }
                Err(e) => {
                    log::warn!("PDF页{} 图片解码失败: {}", i + 1, e);
                    None
                }
            }
        } else {
            None // shutting down
        };

        let ocr_elapsed = t_ocr_start.elapsed().as_millis();

        // Encode to base64 data URL for preview
        let b64 = base64::engine::general_purpose::STANDARD.encode(&data);
        let data_url = format!("data:image/png;base64,{}", b64);

        let ocr_info = ocr_result.as_ref()
            .map(|r| format!("{} chars, {} lines", r.text.len(), r.lines.len()))
            .unwrap_or_else(|| "skipped".to_string());

        log::info!(
            "Render+OCR page {} ({}x{}) @ {}dpi, OCR: {}ms ({})",
            i + 1, dest_w, dest_h, effective_dpi, ocr_elapsed, ocr_info
        );

        results.push(RenderedOcrPage {
            index: i,
            image_data_url: data_url,
            width: dest_w,
            height: dest_h,
            render_dpi: effective_dpi,
            ocr_result,
        });
    }

    drop(doc);
    drop(file);

    Ok(results)
}

#[cfg(all(not(target_os = "windows"), feature = "ocr"))]
pub(crate) fn render_and_ocr_pdf(_pdf_path: &str, _dpi: u32) -> Result<Vec<RenderedOcrPage>, String> {
    Ok(vec![])
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

            if !["pdf", "jpg", "jpeg", "png", "bmp", "webp", "tiff", "tif", "ofd", "xml"].contains(&ext.as_str()) {
                return None;
            }

            let size = path.metadata().ok()?.len();
            Some((path_str.clone(), name, ext, size))
        })
        .collect();

    // Process OFD/XML files first (sequential, they need separate parsing)
    let mut results: Vec<FileData> = Vec::new();
    let mut non_ofd_paths: Vec<(String, String, String, u64)> = Vec::new();

    for (path_str, name, ext, size) in valid_paths {
        if ext == "ofd" {
            // Return OFD as a single entry with ext="ofd" — frontend will call parse_ofd
            // to get SVG vector rendering + structured invoice data from XML (skipping OCR).
            // Fallback to bitmap path is handled by the frontend via open_ofd_images command.
            results.push(FileData {
                name: name.clone(),
                ext: "ofd".to_string(),
                size,
                data_url: String::new(),
                path: Some(path_str.clone()),
                orig_w: None,
                orig_h: None,
            });
        } else if ext == "xml" {
            // Return XML as a single entry with ext="xml" — frontend will call parse_xml_invoice
            // to extract structured invoice data. XML has no visual layout, no preview image.
            results.push(FileData {
                name: name.clone(),
                ext: "xml".to_string(),
                size,
                data_url: String::new(),
                path: Some(path_str.clone()),
                orig_w: None,
                orig_h: None,
            });
        } else {
            non_ofd_paths.push((path_str, name, ext, size));
        }
    }

    // Process non-OFD files in parallel using rayon.
    // **Optimization**: For image files, generate a small JPEG thumbnail instead of
    // sending the full base64-encoded image. A 300 DPI invoice (~3MB) would become
    // ~4MB in base64 — the thumbnail is only ~30KB, a 100x reduction in IPC data.
    // The original file path is passed so Rust can read the full image for OCR/PDF.
    // For PDF files, data_url is empty — they are rendered via render_and_ocr_pdf.
    const THUMB_MAX_DIM: u32 = 600; // Thumbnail max longest side in pixels

    let parallel_results: Vec<FileData> = non_ofd_paths
        .par_iter()
        .filter_map(|(path_str, name, ext, size)| {
            if ext == "pdf" {
                // PDF files: no data_url needed — rendered on demand by render_and_ocr_pdf
                return Some(FileData {
                    name: name.clone(),
                    ext: ext.clone(),
                    size: *size,
                    data_url: String::new(),
                    path: Some(path_str.clone()),
                    orig_w: None,
                    orig_h: None,
                });
            }

            // Image files: read, decode, generate thumbnail
            let bytes = std::fs::read(path_str).ok()?;

            // Read EXIF orientation (JPEG only, non-JPEG returns 1)
            let exif_orient = if is_jpeg_bytes(&bytes) {
                read_exif_orientation(&bytes)
            } else {
                1
            };

            // Decode image and capture original dimensions
            let (thumbnail_data_url, img_orig_w, img_orig_h) = match image::load_from_memory(&bytes) {
                Ok(img) => {
                    // Apply EXIF orientation so thumbnail + reported dimensions are correct.
                    // Browsers auto-rotate <img> based on EXIF, so frontend preview looks right.
                    // Without this, orig_w/orig_h would be swapped for rotated photos,
                    // causing wrong layout rotation decisions.
                    let img = if exif_orient != 1 {
                        apply_exif_orientation(img, exif_orient)
                    } else {
                        img
                    };
                    let ow = img.width();
                    let oh = img.height();
                    let longest = ow.max(oh);

                    let thumb_img = if longest > THUMB_MAX_DIM {
                        let scale = THUMB_MAX_DIM as f32 / longest as f32;
                        let new_w = (ow as f32 * scale).round() as u32;
                        let new_h = (oh as f32 * scale).round() as u32;
                        img.resize_exact(new_w, new_h, image::imageops::FilterType::Triangle)
                    } else {
                        img
                    };

                    // Encode as JPEG (much smaller than PNG for photos/scanned invoices)
                    let data_url = encode_thumbnail_jpeg(&thumb_img)
                        .or_else(|| encode_thumbnail_png(&thumb_img))
                        .unwrap_or_else(|| encode_raw_base64(&bytes, ext));

                    (data_url, ow, oh)
                }
                Err(_) => {
                    // Image decode failed — fall back to raw base64
                    let data_url = encode_raw_base64(&bytes, ext);
                    (data_url, 0, 0)
                }
            };

            Some(FileData {
                name: name.clone(),
                ext: ext.clone(),
                size: *size,
                data_url: thumbnail_data_url,
                path: Some(path_str.clone()),
                orig_w: if img_orig_w > 0 { Some(img_orig_w) } else { None },
                orig_h: if img_orig_h > 0 { Some(img_orig_h) } else { None },
            })
        })
        .collect();

    results.extend(parallel_results);
    Ok(results)
}

/// Encode an image as JPEG thumbnail, returns data URL on success.
fn encode_thumbnail_jpeg(img: &image::DynamicImage) -> Option<String> {
    use base64::Engine;
    let mut buf = std::io::Cursor::new(Vec::new());
    img.write_to(&mut buf, image::ImageFormat::Jpeg).ok()?;
    let b64 = base64::engine::general_purpose::STANDARD.encode(buf.into_inner());
    Some(format!("data:image/jpeg;base64,{}", b64))
}

/// Encode an image as PNG thumbnail, returns data URL on success.
fn encode_thumbnail_png(img: &image::DynamicImage) -> Option<String> {
    use base64::Engine;
    let mut buf = std::io::Cursor::new(Vec::new());
    img.write_to(&mut buf, image::ImageFormat::Png).ok()?;
    let b64 = base64::engine::general_purpose::STANDARD.encode(buf.into_inner());
    Some(format!("data:image/png;base64,{}", b64))
}

/// Encode raw bytes as base64 data URL (fallback when thumbnail generation fails).
fn encode_raw_base64(bytes: &[u8], ext: &str) -> String {
    use base64::Engine;
    let mime = match ext {
        "jpg" | "jpeg" => "image/jpeg",
        "png" => "image/png",
        "bmp" => "image/bmp",
        "webp" => "image/webp",
        "tiff" | "tif" => "image/tiff",
        _ => "application/octet-stream",
    };
    let b64 = base64::engine::general_purpose::STANDARD.encode(bytes);
    format!("data:{};base64,{}", mime, b64)
}

// =====================================================
// Text Enhancement — 浅色/模糊图片发票增强
// =====================================================

/// Enhance a faint/blurry invoice image: auto levels stretch + gamma + unsharp mask.
/// Reads the ORIGINAL file from disk at full resolution (print clarity preserved),
/// bakes EXIF orientation into pixels, returns enhanced JPEG data URL.
///
/// Algorithm (per-pixel LUT, preserves hue — red seals stay red):
///   1. Luminance histogram → 1%/99% percentile clip points
///   2. Levels stretch [lo, hi] → [0, 255] with gamma 1.4 (darkens midtones,
///      turning faint gray text dark) — same LUT on all RGB channels
///   3. Mild unsharp mask (sigma 1.0, amount 0.6, threshold 8) sharpens blurry
///      edges without amplifying flat-area noise
pub(crate) fn enhance_image(file_path: &str) -> Result<String, String> {
    use base64::Engine;
    use image::ImageEncoder;

    let bytes = std::fs::read(file_path)
        .map_err(|e| format!("读取文件失败: {}", e))?;
    let img = image::load_from_memory(&bytes)
        .map_err(|e| format!("图片解码失败: {}", e))?;

    // Bake EXIF orientation so enhanced output displays identically to source
    // (browsers auto-rotate by EXIF; our JPEG output carries no EXIF tag).
    let orient = if is_jpeg_bytes(&bytes) { read_exif_orientation(&bytes) } else { 1 };
    let img = if orient != 1 { apply_exif_orientation(img, orient) } else { img };

    let mut rgb = img.to_rgb8();
    let (w, h) = rgb.dimensions();
    let total = (w as u64 * h as u64) as usize;
    if total == 0 {
        return Err("图片尺寸无效".to_string());
    }

    // 1. Luminance histogram (Rec.601 integer approximation)
    let mut hist = [0u32; 256];
    for px in rgb.pixels() {
        let lum = (299 * px[0] as u32 + 587 * px[1] as u32 + 114 * px[2] as u32 + 500) / 1000;
        hist[lum as usize] += 1;
    }

    // 2. Percentile clip points
    let percentile = |p: f32| -> u8 {
        let target = (total as f32 * p) as u32;
        let mut acc = 0u32;
        for (i, &c) in hist.iter().enumerate() {
            acc += c;
            if acc > target {
                return i as u8;
            }
        }
        255
    };
    let lo = percentile(0.01);
    let hi = percentile(0.99);

    // 3. Build LUT: levels stretch + gamma. Degenerate flat images (hi ≈ lo)
    //    keep identity to avoid noise amplification.
    let mut lut = [0u8; 256];
    if hi > lo + 10 {
        let gamma = 1.4f32;
        let span = (hi - lo) as f32;
        for (i, v) in lut.iter_mut().enumerate() {
            let t = ((i as f32 - lo as f32) / span).clamp(0.0, 1.0);
            *v = (255.0 * t.powf(gamma)).round() as u8;
        }
    } else {
        for (i, v) in lut.iter_mut().enumerate() {
            *v = i as u8;
        }
    }

    for px in rgb.pixels_mut() {
        px[0] = lut[px[0] as usize];
        px[1] = lut[px[1] as usize];
        px[2] = lut[px[2] as usize];
    }

    // 4. Unsharp mask: out = orig + amount * (orig - blur), gated by threshold
    let blurred = image::imageops::blur(&rgb, 1.0);
    const AMOUNT_PCT: i32 = 60;
    const THRESHOLD: i32 = 8;
    for (px, bp) in rgb.pixels_mut().zip(blurred.pixels()) {
        for c in 0..3 {
            let diff = px[c] as i32 - bp[c] as i32;
            if diff.abs() > THRESHOLD {
                px[c] = (px[c] as i32 + diff * AMOUNT_PCT / 100).clamp(0, 255) as u8;
            }
        }
    }

    // 5. Encode JPEG q92 (no EXIF — orientation already baked into pixels)
    let mut buf: Vec<u8> = Vec::new();
    let encoder = image::codecs::jpeg::JpegEncoder::new_with_quality(&mut buf, 92);
    encoder
        .write_image(rgb.as_raw(), w, h, image::ExtendedColorType::Rgb8)
        .map_err(|e| format!("JPEG编码失败: {}", e))?;
    let b64 = base64::engine::general_purpose::STANDARD.encode(&buf);
    Ok(format!("data:image/jpeg;base64,{}", b64))
}

// =====================================================
// PDF Generation from layout request (only remaining path)
// =====================================================

#[cfg(target_os = "windows")]
pub fn list_printers() -> Result<Vec<PrinterInfo>, String> {
    use windows::Win32::Graphics::Printing::{EnumPrintersW, PRINTER_ENUM_LOCAL, PRINTER_ENUM_CONNECTIONS, PRINTER_INFO_4W};
    use windows::core::PCWSTR;

    let default_name = get_default_printer_name();

    unsafe {
        let mut bytes_needed: u32 = 0;
        let mut count_returned: u32 = 0;
        let flags = PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS;
        let null_name = PCWSTR::null();

        // Step 1: query required buffer size
        let _ = EnumPrintersW(flags, null_name, 4, None, &mut bytes_needed, &mut count_returned);
        if bytes_needed == 0 {
            return Ok(vec![]);
        }

        // Step 2: allocate buffer and enumerate
        let mut buffer: Vec<u8> = vec![0u8; bytes_needed as usize];
        EnumPrintersW(
            flags,
            null_name,
            4,
            Some(&mut buffer),
            &mut bytes_needed,
            &mut count_returned,
        ).map_err(|e| format!("获取打印机列表失败: {}", e))?;

        let ptr = buffer.as_ptr() as *const PRINTER_INFO_4W;
        let mut result = Vec::with_capacity(count_returned as usize);

        for i in 0..count_returned {
            let info = &*ptr.offset(i as isize);
            // pPrinterName is PWSTR — convert from UTF-16 to Rust String
            let name = if info.pPrinterName.is_null() {
                continue;
            } else {
                let ptr = info.pPrinterName.0;
                let len = (0..).take_while(|&j| *ptr.offset(j) != 0).count();
                String::from_utf16_lossy(std::slice::from_raw_parts(ptr, len))
            };

            let is_default = default_name.as_ref().map_or(false, |dn| dn.eq_ignore_ascii_case(&name));
            result.push(PrinterInfo { name, is_default });
        }

        Ok(result)
    }
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
// SumatraPDF CLI Silent Print
// =====================================================

#[derive(Debug, Clone)]
pub struct SumatraPdfInfo {
    pub path: std::path::PathBuf,
}

#[cfg(target_os = "windows")]
pub fn find_sumatrapdf() -> Option<SumatraPdfInfo> {
    if let Ok(exe) = std::env::current_exe() {
        if let Some(parent) = exe.parent() {
            let bundled = parent.join("tools").join("SumatraPDF.exe");
            if bundled.exists() {
                log::info!("Found bundled SumatraPDF: {}", bundled.display());
                return Some(SumatraPdfInfo { path: bundled });
            }
        }
    }

    if let Some(path) = find_sumatrapdf_from_registry() {
        log::info!("Found SumatraPDF from registry: {}", path.display());
        return Some(SumatraPdfInfo { path });
    }

    let common_paths = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ];
    for p in &common_paths {
        let path = std::path::PathBuf::from(p);
        if path.exists() {
            log::info!("Found SumatraPDF at common path: {}", path.display());
            return Some(SumatraPdfInfo { path });
        }
    }

    if let Some(path_var) = std::env::var_os("PATH") {
        for dir in std::env::split_paths(&path_var) {
            let candidate = dir.join("SumatraPDF.exe");
            if candidate.exists() {
                log::info!("Found SumatraPDF in PATH: {}", candidate.display());
                return Some(SumatraPdfInfo { path: candidate });
            }
        }
    }

    None
}

#[cfg(not(target_os = "windows"))]
pub fn find_sumatrapdf() -> Option<SumatraPdfInfo> {
    None
}

#[cfg(target_os = "windows")]
fn find_sumatrapdf_from_registry() -> Option<std::path::PathBuf> {
    use windows::Win32::System::Registry::{
        RegCloseKey, RegOpenKeyExW, RegQueryValueExW, HKEY, HKEY_LOCAL_MACHINE, KEY_READ,
    };
    use windows::core::PCWSTR;

    unsafe {
        let mut hkey = HKEY::default();
        let subkey_w: Vec<u16> = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\SumatraPDF.exe"
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        let result = RegOpenKeyExW(
            HKEY_LOCAL_MACHINE,
            PCWSTR(subkey_w.as_ptr()),
            0,
            KEY_READ,
            &mut hkey,
        );

        if result.is_err() {
            return None;
        }

        let mut buf = [0u16; 260];
        let mut size = (buf.len() * 2) as u32;
        let mut typ: windows::Win32::System::Registry::REG_VALUE_TYPE = windows::Win32::System::Registry::REG_NONE;

        let query_result = RegQueryValueExW(
            hkey,
            PCWSTR::null(),
            None,
            Some(&mut typ as *mut _),
            Some(buf.as_mut_ptr() as *mut u8),
            Some(&mut size),
        );

        let _ = RegCloseKey(hkey);

        if query_result.is_err() || typ != windows::Win32::System::Registry::REG_SZ {
            return None;
        }

        let u16_len = (size as usize / 2).saturating_sub(1).min(buf.len());
        let path_str = String::from_utf16_lossy(&buf[..u16_len]);
        let path = std::path::PathBuf::from(path_str.trim_end_matches('\0'));

        if path.exists() {
            Some(path)
        } else {
            None
        }
    }
}

pub fn build_sumatra_print_settings(
    copies: u32,
    duplex: bool,
    color_mode: &str,
    _fit_mode: &str,
    paper_w: Option<f32>,
    paper_h: Option<f32>,
) -> String {
    let mut parts: Vec<String> = Vec::new();

    // Always use noscale: the PDF is already laid out at the correct size
    // with all scaling/rotation applied during generation. Using shrink/fit
    // would cause SumatraPDF to rasterize vector content before sending to
    // the printer, ballooning data from ~1MB to 60MB+.
    parts.push("noscale".to_string());

    match color_mode {
        "color" => parts.push("color".to_string()),
        "grayscale" | "bw" => parts.push("monochrome".to_string()),
        _ => {}
    }

    if duplex {
        parts.push("duplexlong".to_string());
    }

    if let (Some(w), Some(h)) = (paper_w, paper_h) {
        if let Some(paper) = infer_paper_size(w, h) {
            parts.push(format!("paper={}", paper));
        }
    }

    if copies > 1 {
        parts.push(format!("{}x", copies));
    }

    parts.join(",")
}

fn infer_paper_size(w: f32, h: f32) -> Option<&'static str> {
    let sizes: [(f32, f32, &str); 6] = [
        (210.0, 297.0, "A4"),
        (148.0, 210.0, "A5"),
        (105.0, 148.0, "A6"),
        (297.0, 420.0, "A3"),
        (216.0, 279.0, "letter"),
        (216.0, 356.0, "legal"),
    ];
    for (sw, sh, name) in &sizes {
        if (w - sw).abs() < 2.0 && (h - sh).abs() < 2.0 {
            return Some(name);
        }
        if (w - sh).abs() < 2.0 && (h - sw).abs() < 2.0 {
            return Some(name);
        }
    }
    None
}

#[cfg(target_os = "windows")]
pub fn print_with_sumatrapdf(
    sumatra_path: &std::path::Path,
    pdf_path: &std::path::Path,
    printer_name: &str,
    settings_str: &str,
) -> Result<(), String> {
    use std::os::windows::process::CommandExt;
    use std::process::{Command, Stdio};

    let is_virtual_printer = printer_name.to_lowercase().contains("print to pdf")
        || printer_name.to_lowercase().contains("microsoft print to pdf")
        || printer_name.to_lowercase().contains("onenote")
        || printer_name.to_lowercase().contains("fax");

    let mut cmd = Command::new(sumatra_path);
    cmd.arg("-print-to").arg(printer_name);

    if !is_virtual_printer {
        cmd.arg("-silent");
    }

    if !settings_str.is_empty() {
        cmd.arg("-print-settings").arg(settings_str);
    }

    cmd.arg(pdf_path);

    const CREATE_NO_WINDOW: u32 = 0x08000000;
    cmd.creation_flags(CREATE_NO_WINDOW);
    cmd.stdout(Stdio::null()).stderr(Stdio::null());

    log::info!(
        "SumatraPDF print: {:?} -print-to {} {} {} {}",
        sumatra_path,
        printer_name,
        if is_virtual_printer { "" } else { "-silent" },
        if settings_str.is_empty() { "" } else { "-print-settings" },
        settings_str
    );

    let output = cmd
        .output()
        .map_err(|e| format!("启动 SumatraPDF 失败: {}", e))?;

    if !output.status.success() {
        let code = output.status.code().unwrap_or(-1);
        return Err(format!("SumatraPDF 打印失败，退出码: {}", code));
    }

    Ok(())
}

#[cfg(not(target_os = "windows"))]
pub fn print_with_sumatrapdf(
    _sumatra_path: &std::path::Path,
    _pdf_path: &std::path::Path,
    _printer_name: &str,
    _settings_str: &str,
) -> Result<(), String> {
    Err("SumatraPDF 打印仅支持 Windows".to_string())
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
// OCR — PaddleOCR via ocr-rs (MNN inference, high-accuracy Chinese OCR)
// =====================================================
// PDF Text Layer Extraction (no OCR dependency)
// =====================================================

/// A single text word extracted from PDF content stream, with bounding box.
/// Coordinates are in the frontend pixel space (origin top-left, y-down),
/// converted from PDF pt coordinates by the RENDER_DPI/72 scale factor.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PdfTextWord {
    pub text: String,
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

/// A line of text words from PDF content stream.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PdfTextLine {
    pub words: Vec<PdfTextWord>,
    pub confidence: f32, // Always 1.0 for PDF text layer extraction
}

/// Result of PDF text layer extraction with word-level coordinates.
/// Structured identically to OcrResult so the frontend can reuse
/// `extractByCoordinates()` without modification.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PdfTextResult {
    pub text: String,
    pub lines: Vec<PdfTextLine>,
    pub img_w: u32, // Page width in frontend pixels
    pub img_h: u32, // Page height in frontend pixels
    pub has_text_layer: bool, // true if PDF has text content in content stream
}

/// Extract text with coordinates from a PDF page's content stream.
/// Parses Tm/Td (position) + Tj/TJ (text) operations, decodes text via
/// font Encoding (ToUnicode CMap, UniGB, etc.), and outputs word-level
/// bounding boxes in the same coordinate system as OCR results.
///
/// For scanned PDFs (no text layer), returns an empty PdfTextResult.
/// ~5ms per page (pure data parsing, no AI inference).
pub fn extract_pdf_text(pdf_path: &str, page_idx: u32) -> Result<PdfTextResult, String> {
    let doc = lopdf::Document::load(pdf_path)
        .map_err(|e| format!("PDF加载失败: {}", e))?;
    extract_pdf_text_from_doc(&doc, pdf_path, page_idx)
}

/// Extract text from multiple pages in a single PDF document.
/// Only opens the PDF once and processes pages in parallel using rayon.
/// ~5ms per page, with parallelism for multi-page PDFs.
pub fn extract_pdf_texts(pdf_path: &str, page_indices: &[u32]) -> Result<std::collections::HashMap<u32, PdfTextResult>, String> {
    use rayon::prelude::*;
    
    let doc = lopdf::Document::load(pdf_path)
        .map_err(|e| format!("PDF加载失败: {}", e))?;
    
    // Process pages in parallel
    let results: Vec<(u32, Result<PdfTextResult, String>)> = page_indices
        .par_iter()
        .map(|&page_idx| {
            let result = extract_pdf_text_from_doc(&doc, pdf_path, page_idx);
            (page_idx, result)
        })
        .collect();
    
    // Collect into HashMap, skip failed pages (single page failure should not
    // trigger full-batch fallback to slow single-page mode)
    let mut map = std::collections::HashMap::new();
    for (page_idx, result) in results {
        if let Ok(r) = result {
            map.insert(page_idx, r);
        }
    }

    Ok(map)
}

/// Helper function to extract text from a single page of an already-loaded PDF document.
fn extract_pdf_text_from_doc(doc: &lopdf::Document, pdf_path: &str, page_idx: u32) -> Result<PdfTextResult, String> {
    use lopdf::Object;

    let pages: std::collections::BTreeMap<u32, lopdf::ObjectId> = doc.get_pages();
    let page_id = *pages.get(&(page_idx + 1)) // lopdf pages are 1-indexed
        .ok_or_else(|| format!("PDF页面索引{}不存在", page_idx))?;

    // Get page dimensions (pt units)
    let ((_x1, _y1, _x2, _y2), (page_w_pt, page_h_pt)) = get_page_effective_box(&doc, page_id)?;
    let scale = RENDER_DPI as f64 / 72.0; // pt → px
    let page_w_px = (page_w_pt as f64 * scale) as u32;
    let page_h_px = (page_h_pt as f64 * scale) as u32;

    // Collect ToUnicode CMaps for CID font text decoding.
    // We use our own CMap parser (not lopdf's Encoding) to avoid lifetime issues
    // with borrowed references from the Document.
    let mut tounicode_cmaps: std::collections::BTreeMap<Vec<u8>, CMap> = std::collections::BTreeMap::new();
    // Also collect lopdf Encodings for non-CID fonts (WinAnsi, Standard, etc.)
    // These borrow from the Document, so they must be used within this function scope.
    let mut lopdf_encodings: std::collections::BTreeMap<Vec<u8>, lopdf::Encoding> = std::collections::BTreeMap::new();
    // Collect raw Encoding names for fonts that lopdf can't decode (e.g., GBK-EUC-H).
    // Key = font name bytes, Value = encoding name bytes (e.g., b"GBK-EUC-H")
    let mut font_encoding_names: std::collections::BTreeMap<Vec<u8>, Vec<u8>> = std::collections::BTreeMap::new();

    // Helper: load ToUnicode CMap from a font dictionary, also record encoding name
    fn load_cmap_from_font(
        doc: &lopdf::Document,
        font_dict: &lopdf::Dictionary,
    ) -> Option<CMap> {
        let tounicode = font_dict.get(b"ToUnicode").ok()?;
        let resolved_obj = match tounicode {
            lopdf::Object::Reference(xref) => doc.get_object(*xref).ok()?,
            obj => obj,
        };
        let content_bytes = match &resolved_obj {
            lopdf::Object::Stream(s) => s.content.as_slice(),
            _ => resolved_obj.as_stream().ok()?.content.as_slice(),
        };
        let cmap_bytes = match &resolved_obj {
            lopdf::Object::Stream(s) => s.decompressed_content().unwrap_or_else(|_| content_bytes.to_vec()),
            _ => {
                if content_bytes.starts_with(&[0x78, 0x01]) ||
                   content_bytes.starts_with(&[0x78, 0x9C]) ||
                   content_bytes.starts_with(&[0x78, 0xDA]) {
                    flate2::read::ZlibDecoder::new(content_bytes).bytes().collect::<Result<Vec<_>, _>>()
                        .unwrap_or_else(|_| content_bytes.to_vec())
                } else {
                    content_bytes.to_vec()
                }
            }
        };
        if let Ok(content) = String::from_utf8(cmap_bytes.clone()) {
            if let Some(cmap) = parse_cmap(&content) {
                return Some(cmap);
            }
        }
        let content = String::from_utf8_lossy(&cmap_bytes);
        parse_cmap(&content)
    }

    // Step 1: Load fonts from page-level Resources
    let page_fonts = doc.get_page_fonts(page_id).unwrap_or_default();
    for (name, font_dict) in &page_fonts {
        // Record the raw Encoding name for GBK/CJK decoding fallback
        if let Ok(enc_name) = font_dict.get(b"Encoding").and_then(|e| e.as_name().map(|n| n.to_vec())) {
            font_encoding_names.insert(name.clone(), enc_name);
        }
        // Try lopdf's get_font_encoding first (handles WinAnsi, Identity-H→CMap, etc.)
        if let Ok(enc) = font_dict.get_font_encoding(&doc) {
            lopdf_encodings.insert(name.clone(), enc);
        }
        // Also try our own CMap parser (works for all CID fonts with ToUnicode)
        if let Some(cmap) = load_cmap_from_font(&doc, font_dict) {
            tounicode_cmaps.insert(name.clone(), cmap);
        }
    }

    // Step 2: Look for Form XObjects and load fonts from their embedded Resources.
    // IMPORTANT: We must ALWAYS scan Form XObjects, not just when page-level fonts are empty.
    // Many PDFs (e.g., dzcp-format VAT invoices) have page-level label text (STKaiti fonts for
    // "名称:", "统一社会信用代码/" etc.) AND Form XObjects containing the actual values
    // (company names in SimSun, credit codes in CourierNew, amounts, etc.).
    // If we skip Form XObject font loading when page-level fonts exist, the value text
    // in Form XObjects can't be decoded and is silently lost.
    let mut form_xobject_ids: Vec<lopdf::ObjectId> = Vec::new();
    {
        if let Ok(page_dict) = doc.get_dictionary(page_id) {
            if let Ok(resources) = page_dict.get(b"Resources") {
                let resolved_res = match resources {
                    lopdf::Object::Reference(id) => doc.get_dictionary(*id).ok(),
                    _ => resources.as_dict().ok(),
                };
                if let Some(res_dict) = resolved_res {
                    // Check XObject dictionary for Form XObjects
                    if let Ok(xobj_val) = res_dict.get(b"XObject") {
                        let resolved_xobj = match xobj_val {
                            lopdf::Object::Reference(id) => doc.get_dictionary(*id).ok(),
                            _ => xobj_val.as_dict().ok(),
                        };
                        if let Some(xobj_dict) = resolved_xobj {
                            for (_xobj_name, xobj_ref) in xobj_dict.iter() {
                                let xobj_id = match xobj_ref {
                                    lopdf::Object::Reference(id) => Some(*id),
                                    _ => None,
                                };
                                if let Some(id) = xobj_id {
                                    if let Ok(obj) = doc.get_object(id) {
                                        if let Ok(stream) = obj.as_stream() {
                                            // Check if this is a Form XObject
                                            let subtype = stream.dict.get(b"Subtype")
                                                .and_then(|s| s.as_name()).ok();
                                            if subtype == Some(b"Form") {
                                                form_xobject_ids.push(id);
                                                // Get fonts from Form XObject's Resources
                                                // Only load fonts that aren't already in the page-level set
                                                if let Ok(form_res) = stream.dict.get(b"Resources") {
                                                    let resolved_form_res = match form_res {
                                                        lopdf::Object::Reference(rid) => doc.get_dictionary(*rid).ok(),
                                                        _ => form_res.as_dict().ok(),
                                                    };
                                                    if let Some(form_res_dict) = resolved_form_res {
                                                        if let Ok(form_font_dict) = form_res_dict.get(b"Font") {
                                                            let resolved_fonts = match form_font_dict {
                                                                lopdf::Object::Reference(rid) => doc.get_dictionary(*rid).ok(),
                                                                _ => form_font_dict.as_dict().ok(),
                                                            };
                                                            if let Some(fonts_dict) = resolved_fonts {
                                                                for (fname, fref) in fonts_dict.iter() {
                                                                    // Skip if this font name is already loaded from page-level
                                                                    if lopdf_encodings.contains_key(fname) && tounicode_cmaps.contains_key(fname) {
                                                                        continue;
                                                                    }
                                                                    let font_dict = match fref {
                                                                        lopdf::Object::Reference(fid) => doc.get_dictionary(*fid).ok(),
                                                                        lopdf::Object::Dictionary(d) => Some(d),
                                                                        _ => None,
                                                                    };
                                                                    if let Some(fd) = font_dict {
                                                                        // Record raw Encoding name for GBK/CJK fallback
                                                                        if let Ok(enc_name) = fd.get(b"Encoding").and_then(|e| e.as_name().map(|n| n.to_vec())) {
                                                                            font_encoding_names.insert(fname.clone(), enc_name);
                                                                        }
                                                                        if !lopdf_encodings.contains_key(fname) {
                                                                            if let Ok(enc) = fd.get_font_encoding(&doc) {
                                                                                lopdf_encodings.insert(fname.clone(), enc);
                                                                            }
                                                                        }
                                                                        if !tounicode_cmaps.contains_key(fname) {
                                                                            if let Some(cmap) = load_cmap_from_font(&doc, fd) {
                                                                                tounicode_cmaps.insert(fname.clone(), cmap);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Debug: write log file
    let _debug_log_path = std::env::temp_dir().join("ticketchan_text_extract_debug.txt");
    let mut _debug_log = String::new();
    _debug_log.push_str(&format!("PDF: {}, page: {}\n", pdf_path, page_idx));
    _debug_log.push_str(&format!("LopdfEncodings: {}, CMaps: {}, FormXObjs: {}\n",
        lopdf_encodings.len(), tounicode_cmaps.len(), form_xobject_ids.len()));
    let _ = std::fs::write(&_debug_log_path, &_debug_log);

    log::info!("PDF文本提取: {} lopdf encodings, {} ToUnicode CMaps, {} Form XObjects",
        lopdf_encodings.len(), tounicode_cmaps.len(), form_xobject_ids.len());

    // Get the content stream to parse
    // IMPORTANT: Always expand Form XObject content streams when they exist.
    // Many PDFs (e.g., dzcp-format invoices) have both page-level label text AND
    // Form XObjects containing the actual values (company names, amounts, credit codes).
    // We must append Form XObject text operations AFTER the page-level text operations,
    // so that both labels and values are extracted.
    let page_content = doc.get_and_decode_page_content(page_id)
        .map_err(|e| format!("PDF内容流解码失败: {}", e))?;

    // Build a mapping: XObject name → Form XObject ObjectId
    // This is needed to match "Do" operations in the page content to their Form XObjects.
    let mut xobj_name_to_id: std::collections::HashMap<Vec<u8>, lopdf::ObjectId> = std::collections::HashMap::new();
    {
        if let Ok(page_dict) = doc.get_dictionary(page_id) {
            if let Ok(resources) = page_dict.get(b"Resources") {
                let resolved_res = match resources {
                    lopdf::Object::Reference(id) => doc.get_dictionary(*id).ok(),
                    _ => resources.as_dict().ok(),
                };
                if let Some(res_dict) = resolved_res {
                    if let Ok(xobj_val) = res_dict.get(b"XObject") {
                        let resolved_xobj = match xobj_val {
                            lopdf::Object::Reference(id) => doc.get_dictionary(*id).ok(),
                            _ => xobj_val.as_dict().ok(),
                        };
                        if let Some(xobj_dict) = resolved_xobj {
                            for (xobj_name, xobj_ref) in xobj_dict.iter() {
                                if let lopdf::Object::Reference(id) = xobj_ref {
                                    xobj_name_to_id.insert(xobj_name.clone(), *id);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Scan the page content to find cm operations that precede Do operations for Form XObjects.
    // Pattern: q [cm] /Name Do Q
    // We record the cm parameters for each Form XObject name so we can apply them when
    // processing the Form XObject's content.
    let mut form_cm_params: std::collections::HashMap<Vec<u8>, [f64; 6]> = std::collections::HashMap::new();
    {
        let ops = &page_content.operations;
        let mut i = 0;
        while i < ops.len() {
            // Look for "Do" operations
            if ops[i].operator == "Do" {
                if let Some(name_obj) = ops[i].operands.first() {
                    if let lopdf::Object::Name(name) = name_obj {
                        // Check if this Do references a Form XObject
                        if xobj_name_to_id.contains_key(name) {
                            // Look backwards for a preceding "cm" operation (skip "q" if present)
                            let j = if i > 0 && ops[i - 1].operator == "cm" { i - 1 }
                                        else if i > 1 && ops[i - 1].operator == "q" && ops[i - 2].operator == "cm" { i - 2 }
                                        else { i };
                            if j < i && ops[j].operator == "cm" && ops[j].operands.len() == 6 {
                                let params: Vec<f64> = ops[j].operands.iter().map(|o| match o {
                                    lopdf::Object::Real(r) => *r as f64,
                                    lopdf::Object::Integer(n) => *n as f64,
                                    _ => 0.0,
                                }).collect();
                                form_cm_params.insert(name.clone(), [params[0], params[1], params[2], params[3], params[4], params[5]]);
                                log::info!("PDF文本提取: Form XObject {:?} 的页面级cm: {:?}", String::from_utf8_lossy(name), params);
                            }
                        }
                    }
                }
            }
            i += 1;
        }
    }

    let content = if !form_xobject_ids.is_empty() {
        // Build combined content: page operations (with Do/cm/q/Q for Form XObjects removed)
        // followed by Form XObject content wrapped in q/cm/[content]/Q.
        let mut combined_ops: Vec<lopdf::content::Operation> = Vec::new();
        let ops = &page_content.operations;
        let mut i = 0;
        while i < ops.len() {
            let op = &ops[i];
            // Detect patterns: q? cm? /Name Do Q? — skip them from page content
            // and we'll insert Form XObject content separately after page content.
            if op.operator == "Do" {
                if let Some(name_obj) = op.operands.first() {
                    if let lopdf::Object::Name(name) = name_obj {
                        if xobj_name_to_id.contains_key(name) {
                            // This Do references a Form XObject — skip it.
                            // Also skip preceding q/cm and following Q.
                            // Remove preceding q and cm if they were for this Do.
                            if combined_ops.len() >= 2 {
                                let last2 = &combined_ops[combined_ops.len() - 2..];
                                if last2[0].operator == "cm" && last2[1].operator == "q" {
                                    combined_ops.truncate(combined_ops.len() - 2);
                                } else if last2[1].operator == "cm" {
                                    combined_ops.truncate(combined_ops.len() - 1);
                                } else if combined_ops.last().map_or(false, |o| o.operator == "q") {
                                    combined_ops.truncate(combined_ops.len() - 1);
                                }
                            }
                            // Skip following Q if present
                            if i + 1 < ops.len() && ops[i + 1].operator == "Q" {
                                i += 1;
                            }
                            i += 1;
                            continue;
                        }
                    }
                }
            }
            combined_ops.push(op.clone());
            i += 1;
        }

        // Now append Form XObject content with proper cm transformations.
        // For each Form XObject, insert: q [page_cm] [form_matrix_cm] [content] Q
        for fxobj_id in &form_xobject_ids {
            if let Ok(obj) = doc.get_object(*fxobj_id) {
                if let Ok(stream) = obj.as_stream() {
                    // Find the XObject name for this Form XObject (to look up page-level cm)
                    let fxobj_name = xobj_name_to_id.iter()
                        .find(|(_, id)| *id == fxobj_id)
                        .map(|(name, _)| name.clone());

                    // Insert q (save graphics state)
                    combined_ops.push(lopdf::content::Operation {
                        operator: "q".into(),
                        operands: vec![],
                    });

                    // Insert page-level cm (from the page content, before the Do)
                    if let Some(ref name) = fxobj_name {
                        if let Some(cm) = form_cm_params.get(name) {
                            // Only insert if not identity
                            if !(cm[0] == 1.0 && cm[1] == 0.0 && cm[2] == 0.0 && cm[3] == 1.0 && cm[4] == 0.0 && cm[5] == 0.0) {
                                combined_ops.push(lopdf::content::Operation {
                                    operator: "cm".into(),
                                    operands: cm.iter().map(|&v| lopdf::Object::Real(v as f32)).collect(),
                                });
                            }
                        }
                    }

                    // Insert Form XObject's Matrix as cm (if not identity)
                    let form_matrix: [f64; 6] = stream.dict.get(b"Matrix")
                        .and_then(|m| m.as_array())
                        .map(|arr| {
                            let vals: Vec<f64> = arr.iter().map(|o| match o {
                                lopdf::Object::Real(r) => *r as f64,
                                lopdf::Object::Integer(n) => *n as f64,
                                _ => 0.0,
                            }).collect();
                            if vals.len() == 6 { [vals[0], vals[1], vals[2], vals[3], vals[4], vals[5]] }
                            else { [1.0, 0.0, 0.0, 1.0, 0.0, 0.0] }
                        })
                        .unwrap_or([1.0, 0.0, 0.0, 1.0, 0.0, 0.0]);

                    if !(form_matrix[0] == 1.0 && form_matrix[1] == 0.0 && form_matrix[2] == 0.0 && form_matrix[3] == 1.0 && form_matrix[4] == 0.0 && form_matrix[5] == 0.0) {
                        combined_ops.push(lopdf::content::Operation {
                            operator: "cm".into(),
                            operands: form_matrix.iter().map(|&v| lopdf::Object::Real(v as f32)).collect(),
                        });
                        log::info!("PDF文本提取: Form XObject Matrix: {:?}", form_matrix);
                    }

                    // Append Form XObject content
                    if let Ok(decompressed) = stream.decompressed_content() {
                        let fixed_bytes = escape_backslashes_in_literal_strings(&decompressed);
                        if let Ok(form_content) = lopdf::content::Content::decode(&fixed_bytes) {
                            combined_ops.extend(form_content.operations);
                        }
                    }

                    // Insert Q (restore graphics state)
                    combined_ops.push(lopdf::content::Operation {
                        operator: "Q".into(),
                        operands: vec![],
                    });
                }
            }
        }
        lopdf::content::Content { operations: combined_ops }
    } else {
        page_content
    };

    // Check if content stream has any text operations (BT...ET with Tj/TJ)
    let has_text_ops = content.operations.iter().any(|op| {
        op.operator == "Tj" || op.operator == "TJ"
    });
    if !has_text_ops {
        log::info!("PDF文本提取: 页面无文本操作(扫描件)，需OCR回退");
    }

    // State tracking for text position
    let mut cur_x: f64 = 0.0;
    let mut cur_y: f64 = 0.0;
    let mut line_start_x: f64 = 0.0; // x at start of current line (for Td offset)
    let mut font_size: f64 = 12.0;
    let mut leading: f64 = 0.0; // TL-set leading (0 = use font_size * 1.2)
    let mut current_font: Vec<u8> = Vec::new();
    let mut in_text_block = false;

    // CTM (Current Transformation Matrix) tracking: [a, b, c, d, e, f]
    // In PDF, the CTM transforms coordinates from user space to device space.
    // For text extraction, we apply the CTM to text positions to get correct page coordinates.
    // Default is identity: x' = x, y' = y.
    // General transform: x' = a*x + c*y + e, y' = b*x + d*y + f
    let mut ctm: [f64; 6] = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];

    // Graphics state stack for q/Q
    #[derive(Clone)]
    struct GfxState {
        x: f64,
        y: f64,
        line_start_x: f64,
        font_size: f64,
        leading: f64,
        font_name: Vec<u8>,
        ctm: [f64; 6],
    }
    let mut state_stack: Vec<GfxState> = Vec::new();

    let mut all_words: Vec<PdfTextWord> = Vec::new();
    let mut full_text_parts: Vec<String> = Vec::new();
    let mut need_space_before = false; // Insert space between Tj/TJ text fragments

    for op in &content.operations {
        match op.operator.as_str() {
            "BT" => {
                in_text_block = true;
                cur_x = 0.0;
                cur_y = 0.0;
                line_start_x = 0.0;
                leading = 0.0;
                need_space_before = false;
            }
            "ET" => {
                in_text_block = false;
                full_text_parts.push("\n".to_string());
                need_space_before = false;
            }
            "q" => {
                state_stack.push(GfxState {
                    x: cur_x, y: cur_y, line_start_x,
                    font_size, leading, font_name: current_font.clone(),
                    ctm,
                });
            }
            "Q" => {
                if let Some(state) = state_stack.pop() {
                    cur_x = state.x;
                    cur_y = state.y;
                    line_start_x = state.line_start_x;
                    font_size = state.font_size;
                    leading = state.leading;
                    current_font = state.font_name;
                    ctm = state.ctm;
                }
            }
            "cm" if op.operands.len() == 6 => {
                // Concatenate Matrix: a b c d e f cm
                // Multiplies the current CTM with the new matrix.
                // CTM = CTM × M, where M = [a b c d e f]
                let m: [f64; 6] = op.operands.iter().map(|o| match o {
                    Object::Real(r) => *r as f64, Object::Integer(n) => *n as f64, _ => 0.0
                }).collect::<Vec<_>>().try_into().unwrap_or([1.0, 0.0, 0.0, 1.0, 0.0, 0.0]);
                let [a1, b1, c1, d1, e1, f1] = ctm;
                let [a2, b2, c2, d2, e2, f2] = m;
                ctm = [
                    a1*a2 + b1*c2,  // new_a
                    a1*b2 + b1*d2,  // new_b
                    c1*a2 + d1*c2,  // new_c
                    c1*b2 + d1*d2,  // new_d
                    e1*a2 + f1*c2 + e2, // new_e
                    e1*b2 + f1*d2 + f2, // new_f
                ];
                log::debug!("PDF文本提取: cm操作 [{:.1} {:.1} {:.1} {:.1} {:.1} {:.1}], CTM更新为 [{:.1} {:.1} {:.1} {:.1} {:.1} {:.1}]",
                    m[0], m[1], m[2], m[3], m[4], m[5],
                    ctm[0], ctm[1], ctm[2], ctm[3], ctm[4], ctm[5]);
            }
            "Tf" if op.operands.len() >= 2 => {
                // Font selection: /FontName size
                if let Some(font_obj) = op.operands.first() {
                    match font_obj {
                        Object::Name(name) => current_font = name.clone(),
                        Object::Reference(id) => {
                            // Some PDFs reference font by object ID
                            current_font = format!("{:?},{:?}", id.0, id.1).into_bytes();
                        }
                        _ => {}
                    }
                }
                if let Some(size_obj) = op.operands.get(1) {
                    match size_obj {
                        Object::Real(r) => font_size = *r as f64,
                        Object::Integer(i) => font_size = *i as f64,
                        _ => {}
                    }
                }
            }
            "Tm" if op.operands.len() >= 6 && in_text_block => {
                // Text matrix: a b c d e f Tm
                // e = x position, f = y position (in PDF coordinate space, pt)
                // Font size = vertical scale = d (or sqrt(a²+b²) for rotated text)
                let d = match &op.operands[3] {
                    Object::Real(r) => *r as f64, Object::Integer(i) => *i as f64, _ => 0.0
                };
                cur_x = match &op.operands[4] {
                    Object::Real(r) => *r as f64, Object::Integer(i) => *i as f64, _ => 0.0
                };
                cur_y = match &op.operands[5] {
                    Object::Real(r) => *r as f64, Object::Integer(i) => *i as f64, _ => 0.0
                };
                // Use vertical component d for font size (more reliable than a)
                if d > 1.0 { font_size = d; }
                // Tm sets a new absolute position — this becomes the line start
                line_start_x = cur_x;
            }
            "Td" | "TD" if op.operands.len() >= 2 && in_text_block => {
                // Move to next line: tx ty Td
                // PDF spec: offset from start of current line, not from cur_x
                let tx = match &op.operands[0] {
                    Object::Real(r) => *r as f64, Object::Integer(i) => *i as f64, _ => 0.0
                };
                let ty = match &op.operands[1] {
                    Object::Real(r) => *r as f64, Object::Integer(i) => *i as f64, _ => 0.0
                };
                cur_x = line_start_x + tx;
                line_start_x = cur_x;
                cur_y += ty;
            }
            "TL" if op.operands.len() >= 1 && in_text_block => {
                // Set text leading
                match &op.operands[0] {
                    Object::Real(r) => leading = *r as f64,
                    Object::Integer(i) => leading = *i as f64,
                    _ => {}
                }
            }
            "T*" if in_text_block => {
                // Move to start of next line (leading offset)
                let effective_leading = if leading > 0.0 { leading } else { font_size * 1.2 };
                cur_y -= effective_leading;
                cur_x = line_start_x; // Return to line start
            }
            "Tj" if in_text_block => {
                // Show text string
                if let Some(obj) = op.operands.first() {
                    if let Some(decoded) = decode_text_object(obj, &lopdf_encodings, &tounicode_cmaps, &current_font, &font_encoding_names) {
                        if !decoded.is_empty() {
                            // Apply CTM to get page coordinates
                            let (px, py) = apply_ctm(&ctm, cur_x, cur_y);
                            let word = make_word(&decoded, px, py, font_size, page_h_pt, scale);
                            all_words.push(word);
                            if need_space_before { full_text_parts.push(" ".to_string()); }
                            full_text_parts.push(decoded.clone());
                            need_space_before = true;
                            // Advance x position by approximate text width
                            cur_x += approximate_text_width(&decoded, font_size);
                        }
                    }
                }
            }
            "TJ" if in_text_block => {
                // Show text with individual glyph positioning (array)
                // Kern < -80 (in 1/1000 em) indicates a real word break;
                // smaller kerns (-20..-50) are just spacing micro-adjustments.
                const KERN_WORD_BREAK: f64 = -80.0;
                if let Some(Object::Array(arr)) = op.operands.first() {
                    let mut text_buf = String::new();
                    for item in arr {
                        match item {
                            Object::String(bytes, _format) => {
                                if let Some(decoded) = decode_bytes_with_encoding(bytes, &lopdf_encodings, &tounicode_cmaps, &current_font, &font_encoding_names) {
                                    text_buf.push_str(&decoded);
                                }
                            }
                            Object::Integer(kern) => {
                                // Kern displacement in 1/1000 of a unit of text space
                                let kern_f = *kern as f64;
                                // Only flush on large negative kern (word break)
                                if !text_buf.is_empty() && kern_f < KERN_WORD_BREAK {
                                    let (px, py) = apply_ctm(&ctm, cur_x, cur_y);
                                    let word = make_word(&text_buf, px, py, font_size, page_h_pt, scale);
                                    all_words.push(word);
                                    if need_space_before { full_text_parts.push(" ".to_string()); }
                                    full_text_parts.push(text_buf.clone());
                                    need_space_before = true;
                                    cur_x += approximate_text_width(&text_buf, font_size);
                                    text_buf.clear();
                                }
                                // Apply kern offset
                                cur_x += kern_f / 1000.0 * font_size;
                            }
                            Object::Real(kern) => {
                                let kern_f = *kern as f64;
                                if !text_buf.is_empty() && kern_f < KERN_WORD_BREAK {
                                    let (px, py) = apply_ctm(&ctm, cur_x, cur_y);
                                    let word = make_word(&text_buf, px, py, font_size, page_h_pt, scale);
                                    all_words.push(word);
                                    if need_space_before { full_text_parts.push(" ".to_string()); }
                                    full_text_parts.push(text_buf.clone());
                                    need_space_before = true;
                                    cur_x += approximate_text_width(&text_buf, font_size);
                                    text_buf.clear();
                                }
                                cur_x += kern_f / 1000.0 * font_size;
                            }
                            _ => {}
                        }
                    }
                    // Flush remaining text
                    if !text_buf.is_empty() {
                        let (px, py) = apply_ctm(&ctm, cur_x, cur_y);
                        let word = make_word(&text_buf, px, py, font_size, page_h_pt, scale);
                        all_words.push(word);
                        if need_space_before { full_text_parts.push(" ".to_string()); }
                        full_text_parts.push(text_buf);
                        need_space_before = true;
                    }
                }
            }
            _ => {}
        }
    }

    // Group words into lines by y-coordinate proximity
    let lines = group_words_into_lines(&all_words, page_h_px as f64);

    Ok(PdfTextResult {
        text: full_text_parts.join(""),
        lines,
        img_w: page_w_px,
        img_h: page_h_px,
        has_text_layer: has_text_ops,
    })
}

/// Create a PdfTextWord with coordinate conversion (PDF pt → frontend px).
/// PDF coordinates: origin bottom-left, y-up.
/// Frontend coordinates: origin top-left, y-down.
/// Apply CTM (Current Transformation Matrix) to a point.
/// CTM = [a, b, c, d, e, f] represents:
///   | a  b  0 |
///   | c  d  0 |
///   | e  f  1 |
/// Transform: x' = a*x + c*y + e, y' = b*x + d*y + f
fn apply_ctm(ctm: &[f64; 6], x: f64, y: f64) -> (f64, f64) {
    let [a, b, c, d, e, f] = *ctm;
    (a * x + c * y + e, b * x + d * y + f)
}

fn make_word(text: &str, pdf_x: f64, pdf_y: f64, font_size: f64,
             page_h_pt: f32, scale: f64) -> PdfTextWord {
    let w = approximate_text_width(text, font_size) * scale;
    let h = font_size * scale;
    // Convert: frontend_y = (page_h - pdf_y - font_size) * scale
    // The y in PDF is the baseline; the top of the glyph is approximately at y + font_size
    let fx = pdf_x * scale;
    let fy = (page_h_pt as f64 - pdf_y - font_size) * scale;
    PdfTextWord {
        text: text.to_string(),
        x: fx,
        y: fy.max(0.0),
        w,
        h,
    }
}

/// Approximate text width in pt based on character types.
/// CJK characters ≈ font_size, Latin/digits ≈ font_size * 0.5.
fn approximate_text_width(text: &str, font_size: f64) -> f64 {
    let mut width = 0.0;
    for ch in text.chars() {
        if is_cjk(ch) {
            width += font_size;
        } else {
            width += font_size * 0.5;
        }
    }
    width
}

/// Check if a character is CJK (Chinese, Japanese, Korean)
fn is_cjk(ch: char) -> bool {
    let cp = ch as u32;
    // CJK Unified Ideographs + common CJK ranges
    (0x4E00..=0x9FFF).contains(&cp)   // CJK Unified Ideographs
    || (0x3400..=0x4DBF).contains(&cp) // CJK Extension A
    || (0xF900..=0xFAFF).contains(&cp) // CJK Compatibility Ideographs
    || (0x3000..=0x303F).contains(&cp) // CJK Symbols and Punctuation
    || (0xFF00..=0xFFEF).contains(&cp) // Fullwidth Forms
    || (0x2E80..=0x2EFF).contains(&cp) // CJK Radicals Supplement
    || (0x3040..=0x309F).contains(&cp) // Hiragana
    || (0x30A0..=0x30FF).contains(&cp) // Katakana
}

/// Decode a text object (String or other) using the current font's encoding.
fn decode_text_object(
    obj: &lopdf::Object,
    encodings: &std::collections::BTreeMap<Vec<u8>, lopdf::Encoding>,
    cmaps: &std::collections::BTreeMap<Vec<u8>, CMap>,
    font_name: &[u8],
    encoding_names: &std::collections::BTreeMap<Vec<u8>, Vec<u8>>,
) -> Option<String> {
    match obj {
        lopdf::Object::String(bytes, _format) => {
            decode_bytes_with_encoding(bytes, encodings, cmaps, font_name, encoding_names)
        }
        lopdf::Object::Array(arr) => {
            // Some PDFs use arrays in Tj
            let mut result = String::new();
            for item in arr {
                if let lopdf::Object::String(bytes, _fmt) = item {
                    if let Some(decoded) = decode_bytes_with_encoding(bytes, encodings, cmaps, font_name, encoding_names) {
                        result.push_str(&decoded);
                    }
                }
            }
            if result.is_empty() { None } else { Some(result) }
        }
        _ => None,
    }
}

/// Simple CMap parser for ToUnicode mappings
#[derive(Clone)]
struct CMap {
    mappings: std::collections::BTreeMap<u16, char>,
    ranges: Vec<(u16, u16, u16)>, // (start, end, unicode_start)
}

impl CMap {
    fn new() -> Self {
        CMap {
            mappings: std::collections::BTreeMap::new(),
            ranges: Vec::new(),
        }
    }
    
    fn add_mapping(&mut self, glyph: u16, ch: char) {
        self.mappings.insert(glyph, ch);
    }
    
    fn add_range(&mut self, start: u16, end: u16, unicode_start: u16) {
        self.ranges.push((start, end, unicode_start));
    }
    
    fn lookup(&self, glyph: u16) -> Option<char> {
        // Check direct mappings first
        if let Some(ch) = self.mappings.get(&glyph) {
            return Some(*ch);
        }
        
        // Check ranges
        for &(start, end, unicode_start) in &self.ranges {
            if glyph >= start && glyph <= end {
                let offset = glyph - start;
                return std::char::from_u32((unicode_start + offset) as u32);
            }
        }
        
        None
    }
}

/// Escape raw backslash bytes (0x5C) inside PDF literal strings to prevent
/// lopdf's escape processing from corrupting 2-byte CID alignment.
///
/// Some PDF producers (e.g., dzcp-format invoices) embed raw 0x5C bytes in
/// literal strings without proper PDF escaping. When lopdf parses these strings,
/// it interprets 0x5C as the start of an escape sequence (e.g., \t → TAB, \n → LF),
/// which breaks the 2-byte CID alignment for Identity-H encoded CID-keyed fonts.
///
/// This function scans the raw content stream bytes, finds literal strings
// TODO: 待办 — 非税发票解析 pdf_oxide换库重构
// TODO: 待办 — 个别发票解析适配
//   当前遗留问题:
//   1. SimSun等子集字体ToUnicode CMap不完整(如dzcp发票缺失"司/贸/限"等CID映射),
//      需TrueType cmap表fallback(需换pdf_oxide等库支持); 当前由OCR兜底。
//   2. 0x5C字节转义预处理(escape_backslashes_in_literal_strings)是lopdf的workaround,
//      换库后可能不再需要。
//   3. Form XObject混合架构(页面标签+XO值)的解析在lopdf下需手动展开,
//      pdf_oxide原生支持后可简化。
/// Scans PDF content-stream bytes for literal strings (delimited by balanced
/// parentheses), and doubles any unescaped 0x5C bytes
/// (i.e., replaces \ with \\), so that lopdf's escape processing will produce
/// the original 0x5C byte.
fn escape_backslashes_in_literal_strings(data: &[u8]) -> Vec<u8> {
    let mut result = Vec::with_capacity(data.len() + data.len() / 10); // extra space for escapes
    let mut i = 0;
    let len = data.len();

    while i < len {
        let b = data[i];
        result.push(b);

        if b == b'(' {
            // Enter literal string — track balanced parentheses
            let mut depth = 1i32;
            i += 1;
            while i < len && depth > 0 {
                let c = data[i];
                if c == b'\\' {
                    // Check if this is a real PDF escape or a raw 0x5C that should be escaped
                    // In a well-formed PDF, \\ followed by a recognized escape char (n, r, t, b, f, (, ), \\)
                    // or octal digit is a real escape. Otherwise, it's a raw 0x5C that needs escaping.
                    if i + 1 < len {
                        let next = data[i + 1];
                        match next {
                            b'n' | b'r' | b't' | b'b' | b'f' | b'(' | b')' | b'\\'
                            | b'0'..=b'7' => {
                                // This looks like a legitimate PDF escape — keep as-is
                                result.push(c);
                                i += 1;
                                result.push(data[i]);
                            }
                            _ => {
                                // Raw 0x5C not part of a recognized escape — escape it
                                result.push(b'\\');
                                result.push(b'\\'); // double backslash → lopdf will decode to single 0x5C
                            }
                        }
                    } else {
                        // Trailing backslash at end of string — escape it
                        result.push(b'\\');
                        result.push(b'\\');
                    }
                    i += 1;
                } else if c == b'(' {
                    depth += 1;
                    result.push(c);
                    i += 1;
                } else if c == b')' {
                    depth -= 1;
                    result.push(c);
                    i += 1;
                } else {
                    result.push(c);
                    i += 1;
                }
            }
        } else {
            i += 1;
        }
    }

    result
}

/// Parse a ToUnicode CMap string
fn parse_cmap(content: &str) -> Option<CMap> {
    let mut cmap = CMap::new();
    
    // Split by whitespace (handles all line endings: \n, \r\n, \r)
    let tokens: Vec<&str> = content.split(|c: char| c.is_whitespace()).filter(|s| !s.is_empty()).collect();
    let mut i = 0;
    
    while i < tokens.len() {
        let token = tokens[i];
        
        // Parse beginbfchar / endbfchar
        if token == "beginbfchar" {
            i += 1;
            while i < tokens.len() && tokens[i] != "endbfchar" {
                // Each entry may be "<glyph> <unicode>" (two tokens)
                // or "<glyph><unicode>" (one token with concatenated hex pairs)
                let parts = split_hex_pairs(tokens[i]);
                if parts.len() >= 2 {
                    if let (Some(glyph), Some(ch)) = (parse_hex_pair(parts[0]), parse_hex_pair(parts[1])) {
                        if let Some(unicode_char) = std::char::from_u32(ch as u32) {
                            cmap.add_mapping(glyph, unicode_char);
                        }
                    }
                    i += 1;
                } else if i + 1 < tokens.len() {
                    // Two separate tokens
                    let glyph_str = tokens[i];
                    let unicode_str = tokens[i + 1];
                    if let (Some(glyph), Some(ch)) = (parse_hex_pair(glyph_str), parse_hex_pair(unicode_str)) {
                        if let Some(unicode_char) = std::char::from_u32(ch as u32) {
                            cmap.add_mapping(glyph, unicode_char);
                        }
                    }
                    i += 2;
                } else {
                    i += 1;
                }
            }
            i += 1; // Skip "endbfchar"
        }
        
        // Parse beginbfrange / endbfrange
        else if token == "beginbfrange" {
            i += 1;
            while i < tokens.len() && tokens[i] != "endbfrange" {
                // Each entry may be "<start> <end> <unicode>" (three tokens)
                // or "<start><end><unicode>" (one token with concatenated hex pairs)
                // or "<start> <end><unicode>" etc.
                let mut all_parts = Vec::new();
                // Collect hex pairs from current token and possibly next tokens
                let mut j = i;
                while all_parts.len() < 3 && j < tokens.len() && tokens[j] != "endbfrange" {
                    let parts = split_hex_pairs(tokens[j]);
                    if !parts.is_empty() {
                        all_parts.extend(parts);
                    } else {
                        // Might be a standalone hex pair token
                        if let Some(_) = parse_hex_pair(tokens[j]) {
                            all_parts.push(tokens[j]);
                        }
                    }
                    j += 1;
                    // If we already have 3 parts, stop
                    if all_parts.len() >= 3 { break; }
                }
                if all_parts.len() >= 3 {
                    if let (Some(start), Some(end), Some(unicode_start)) = 
                        (parse_hex_pair(all_parts[0]), parse_hex_pair(all_parts[1]), parse_hex_pair(all_parts[2])) {
                        cmap.add_range(start, end, unicode_start);
                    }
                }
                i = j;
            }
            i += 1; // Skip "endbfrange"
        }
        
        i += 1;
    }
    
    // Debug: log parsing result
    log::debug!("CMap parsed: {} mappings, {} ranges", cmap.mappings.len(), cmap.ranges.len());
    
    // If we found any mappings, return the CMap
    if !cmap.mappings.is_empty() || !cmap.ranges.is_empty() {
        Some(cmap)
    } else {
        None
    }
}

/// Parse a hex string like "<0041>" to u16
fn parse_hex_pair(s: &str) -> Option<u16> {
    let s = s.trim();
    if s.starts_with('<') && s.ends_with('>') {
        let hex_str = &s[1..s.len()-1];
        u16::from_str_radix(hex_str, 16).ok()
    } else {
        None
    }
}

/// Split a string like "<0005><0007><0031>" into individual hex pairs ["<0005>", "<0007>", "<0031>"]
fn split_hex_pairs(s: &str) -> Vec<&str> {
    let mut result = Vec::new();
    let mut start = None;
    for (i, c) in s.char_indices() {
        if c == '<' {
            start = Some(i);
        } else if c == '>' {
            if let Some(s_idx) = start {
                result.push(&s[s_idx..=i]);
                start = None;
            }
        }
    }
    result
}

/// Decode raw bytes using the font's encoding.
fn decode_bytes_with_encoding(
    bytes: &[u8],
    encodings: &std::collections::BTreeMap<Vec<u8>, lopdf::Encoding>,
    cmaps: &std::collections::BTreeMap<Vec<u8>, CMap>,
    font_name: &[u8],
    encoding_names: &std::collections::BTreeMap<Vec<u8>, Vec<u8>>,
) -> Option<String> {
    // Try the font's encoding first
    if let Some(encoding) = encodings.get(font_name) {
        if let Ok(text) = encoding.bytes_to_string(bytes) {
            return Some(text);
        }
    }

    // Try GBK/EUC CJK encoding fallback — lopdf returns Err for these SimpleEncodings.
    // Common CJK CMap names: GBK-EUC-H, GBK-EUC-V, GBKp-EUC-H, GBKp-EUC-V,
    //   ETen-B5-H, ETen-B5-V, etc.
    // Note: UniGB-UCS2-H and UniGB-UTF16-H are UTF-16 encodings handled by lopdf's
    // bytes_to_string (they succeed), so we don't need to handle them here.
    // Only handle EUC-based CJK encodings that lopdf can't decode.
    if let Some(enc_name) = encoding_names.get(font_name) {
        let enc_str = String::from_utf8_lossy(enc_name);
        // GBK-EUC-H/V, GBKp-EUC-H/V — GBK byte encoding
        let is_gbk_euc = enc_str.starts_with("GBK");
        // ETen-B5-H/V, B5pc-H/V — Big5 byte encoding
        let is_big5_euc = enc_str.starts_with("ETen") || enc_str.starts_with("B5pc");
        // 90ms-RKSJ-H/V, 83pv-RKSJ-H/V — Shift_JIS (Japanese)
        let _is_sjis = enc_str.contains("RKSJ") || enc_str.contains("90ms") || enc_str.contains("83pv");
        // KSC-EUC-H/V, KSCms-UHC-H/V — EUC-KR/UHC (Korean)
        let _is_korean = enc_str.starts_with("KSC");
        if is_gbk_euc {
            // Decode as GBK (GB18030 compatible, covers GB2312/GBK)
            let (cow, _encoding_used, _had_errors) = encoding_rs::GBK.decode(bytes);
            let text = cow.into_owned();
            if !text.is_empty() && text.chars().any(|c| is_cjk(c) || c.is_alphanumeric()) {
                log::debug!("GBK decode for font {:?} (encoding {}): {} chars",
                    String::from_utf8_lossy(font_name), enc_str, text.chars().count());
                return Some(text);
            }
        } else if is_big5_euc {
            // Decode as Big5 (Traditional Chinese)
            let (cow, _encoding_used, _had_errors) = encoding_rs::BIG5.decode(bytes);
            let text = cow.into_owned();
            if !text.is_empty() && text.chars().any(|c| is_cjk(c) || c.is_alphanumeric()) {
                log::debug!("BIG5 decode for font {:?} (encoding {}): {} chars",
                    String::from_utf8_lossy(font_name), enc_str, text.chars().count());
                return Some(text);
            }
        }
    }

    // Try ToUnicode CMap lookup for CID-keyed fonts (Identity-H encoding)
    if let Some(cmap) = cmaps.get(font_name) {
        let mut result = String::new();
        let mut i = 0;
        while i < bytes.len() {
            // Identity-H uses 2-byte glyph indices (big-endian)
            if i + 1 < bytes.len() {
                let glyph_idx = u16::from_be_bytes([bytes[i], bytes[i + 1]]);
                if let Some(ch) = cmap.lookup(glyph_idx) {
                    result.push(ch);
                }
                i += 2;
            } else {
                i += 1;
            }
        }
        if !result.is_empty() && result.chars().any(|c| !c.is_control()) {
            return Some(result);
        }
    }

    // Fallback: try UTF-8 decode (some PDFs store Unicode text directly)
    if let Ok(text) = std::str::from_utf8(bytes) {
        if !text.is_empty() && text.chars().any(|c| !c.is_control()) {
            return Some(text.to_string());
        }
    }

    // Fallback: try UTF-16BE decode (some CIDFonts use this encoding)
    if bytes.len() >= 2 && bytes.len() % 2 == 0 {
        let utf16: Vec<u16> = bytes.chunks(2)
            .filter_map(|c| Some(u16::from_be_bytes([*c.get(0)?, *c.get(1)?])))
            .collect();
        if let Ok(text) = String::from_utf16(&utf16) {
            if !text.is_empty() && text.chars().any(|c| !c.is_control()) {
                return Some(text);
            }
        }
    }

    // Fallback: try UTF-16LE decode (some CIDFonts use this encoding, especially Chinese PDFs)
    if bytes.len() >= 2 && bytes.len() % 2 == 0 {
        let utf16: Vec<u16> = bytes.chunks(2)
            .filter_map(|c| Some(u16::from_le_bytes([*c.get(0)?, *c.get(1)?])))
            .collect();
        if let Ok(text) = String::from_utf16(&utf16) {
            if !text.is_empty() && text.chars().any(|c| !c.is_control()) {
                return Some(text);
            }
        }
    }

    // Last resort: lossy Latin-1 decode
    let text = String::from_utf8_lossy(bytes);
    if !text.is_empty() && text.chars().any(|c| c.is_alphanumeric() || is_cjk(c)) {
        return Some(text.to_string());
    }

    None
}

/// Group words into lines based on y-coordinate proximity.
/// Words on the same line (within half a font-size vertical distance) are grouped together.
/// Adjacent words on the same line with very small gaps (<3px) are merged to fix
/// PDF text fragmentation (e.g., "¥" + "4500.00" → "¥4500.00").
fn group_words_into_lines(words: &[PdfTextWord], _page_h: f64) -> Vec<PdfTextLine> {
    if words.is_empty() {
        return Vec::new();
    }

    // Compute a FIXED band size from median word height.
    // The previous approach used (a.h + b.h) * 0.25 which depends on BOTH elements
    // being compared — this violates transitivity and causes Rust's sort_by to panic
    // with "user-provided comparison function does not correctly implement a total order".
    let median_h = {
        let mut heights: Vec<f64> = words.iter().map(|w| w.h).collect();
        heights.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
        let mid = heights.len() / 2;
        if heights.len() > 0 { heights[mid] } else { 10.0 }
    };
    let band = if median_h > 0.5 { median_h * 0.5 } else { 1.0 }; // min band = 1px

    // Sort by y-band first, then by x within same band — using fixed band size
    let mut sorted: Vec<&PdfTextWord> = words.iter().collect();
    sorted.sort_by(|a, b| {
        let band_a = (a.y / band).round();
        let band_b = (b.y / band).round();
        band_a.partial_cmp(&band_b).unwrap_or(std::cmp::Ordering::Equal)
            .then_with(|| a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal))
    });

    let mut lines: Vec<PdfTextLine> = Vec::new();
    let mut current_words: Vec<PdfTextWord> = Vec::new();
    let mut current_y: f64 = sorted[0].y;
    let mut line_height: f64 = sorted[0].h;

    for word in sorted {
        // Same line if y is within half a line height
        if (word.y - current_y).abs() <= line_height * 0.5 {
            current_words.push(word.clone());
        } else {
            // Finish current line
            if !current_words.is_empty() {
                // Sort words within line by x position
                current_words.sort_by(|a, b| a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal));
                // Merge adjacent words with very small gaps (<3px)
                current_words = merge_adjacent_words(&current_words);
                lines.push(PdfTextLine {
                    words: current_words.clone(),
                    confidence: 1.0,
                });
            }
            // Start new line
            current_words.clear();
            current_words.push(word.clone());
            current_y = word.y;
            line_height = word.h;
        }
    }

    // Don't forget the last line
    if !current_words.is_empty() {
        current_words.sort_by(|a, b| a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal));
        current_words = merge_adjacent_words(&current_words);
        lines.push(PdfTextLine {
            words: current_words,
            confidence: 1.0,
        });
    }

    lines
}

/// Merge adjacent words on the same line that have a very small gap (<3px).
/// PDF text extraction often fragments text: "¥" and "4500.00" as separate Tj operations,
/// or single CJK characters that should form one word ("合"+"计").
/// Merging them restores the semantic word boundaries that the original PDF intended.
fn merge_adjacent_words(words: &[PdfTextWord]) -> Vec<PdfTextWord> {
    if words.len() <= 1 {
        return words.to_vec();
    }
    let max_gap = 3.0; // pixels — only merge touching/nearly-touching words
    let mut result: Vec<PdfTextWord> = Vec::new();
    let mut buf = words[0].clone();

    for i in 1..words.len() {
        let next = &words[i];
        let gap = next.x - (buf.x + buf.w); // distance from right edge of buf to left edge of next
        if gap >= 0.0 && gap <= max_gap {
            // Small positive gap — merge: extend buf to include next
            buf.text.push_str(&next.text);
            buf.w = (next.x + next.w) - buf.x; // new width spans both
            if next.h > buf.h { buf.h = next.h; }
        } else if gap < 0.0 && gap >= -max_gap {
            // Slight overlap (from approximate text widths) — also merge
            buf.text.push_str(&next.text);
            buf.w = buf.w.max(next.x + next.w - buf.x); // take the larger extent
            if next.h > buf.h { buf.h = next.h; }
        } else {
            // Gap too large or overlap too large — flush buf
            result.push(buf);
            buf = next.clone();
        }
    }
    result.push(buf);
    result
}

// =====================================================
// OCR Structures & Functions
// =====================================================

#[cfg(feature = "ocr")]
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

/// A 2D point for polygon coordinates
#[cfg(feature = "ocr")]
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct OcrPoint {
    pub x: f64,
    pub y: f64,
}

/// An OCR line containing words, with line-level bounding polygon and confidence
#[cfg(feature = "ocr")]
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct OcrLine {
    pub words: Vec<OcrWord>,
    /// Four corner points of the text line polygon (from detection model).
    /// Top-left, top-right, bottom-right, bottom-left (roughly).
    /// Used for more accurate coordinate analysis in frontend.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub points: Option<Vec<OcrPoint>>,
    /// OCR confidence for this line (0.0 - 1.0)
    pub confidence: f32,
}

/// Structured OCR result with coordinates
#[cfg(feature = "ocr")]
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

/// Lazy-initialized global OCR engine (PaddleOCR + MNN)
/// Initialized on first use, persists for the app lifetime.
#[cfg(feature = "ocr")]
use std::sync::Mutex;
#[cfg(feature = "ocr")]
static OCR_ENGINE: Mutex<Option<ocr_rs::OcrEngine>> = Mutex::new(None);

/// Get or create the OCR engine.
/// Model files are expected alongside the executable:
///   - PP-OCRv6_small_det.mnn   (detection model)
///   - PP-OCRv6_small_rec.mnn   (recognition model)
///   - ppocr_keys_v6_small.txt  (character set)
#[cfg(feature = "ocr")]
fn get_ocr_engine() -> Result<std::sync::MutexGuard<'static, Option<ocr_rs::OcrEngine>>, String> {
    let mut lock = OCR_ENGINE.lock().map_err(|e| format!("OCR引擎锁失败: {}", e))?;

    if lock.is_none() {
        let exe_dir = std::env::current_exe()
            .map_err(|e| format!("获取exe路径失败: {}", e))?
            .parent()
            .ok_or("无法获取exe目录")?
            .to_path_buf();

        // Tauri 2.x bundle.resources preserves directory structure:
        // "models/X.mnn" → <exe_dir>/models/X.mnn
        // Also try <exe_dir>/X.mnn as fallback (green portable deployment)
        let det_path = if exe_dir.join("models").join("PP-OCRv6_small_det.mnn").exists() {
            exe_dir.join("models").join("PP-OCRv6_small_det.mnn")
        } else {
            exe_dir.join("PP-OCRv6_small_det.mnn")
        };
        let rec_path = if exe_dir.join("models").join("PP-OCRv6_small_rec.mnn").exists() {
            exe_dir.join("models").join("PP-OCRv6_small_rec.mnn")
        } else {
            exe_dir.join("PP-OCRv6_small_rec.mnn")
        };
        let keys_path = if exe_dir.join("models").join("ppocr_keys_v6_small.txt").exists() {
            exe_dir.join("models").join("ppocr_keys_v6_small.txt")
        } else {
            exe_dir.join("ppocr_keys_v6_small.txt")
        };

        // Validate model files exist
        if !det_path.exists() {
            return Err(format!(
                "OCR检测模型不存在: {}（请确保模型文件在exe同级目录或models子目录）",
                det_path.display()
            ));
        }
        if !rec_path.exists() {
            return Err(format!(
                "OCR识别模型不存在: {}（请确保模型文件在exe同级目录或models子目录）",
                rec_path.display()
            ));
        }
        if !keys_path.exists() {
            return Err(format!(
                "OCR字符集文件不存在: {}（请确保模型文件在exe同级目录或models子目录）",
                keys_path.display()
            ));
        }

        log::info!(
            "Loading PaddleOCR models from: {}",
            exe_dir.display()
        );

        let config = ocr_rs::OcrEngineConfig::new()
            .with_parallel(false) // CRITICAL: disable rayon — MNN InferenceEngine is not truly
                                  // thread-safe (unsafe impl Sync). Rayon parallelism with a
                                  // single MNN session causes thread contention and actually
                                  // *slows down* recognition. Use batch inference instead,
                                  // which MNN handles internally with its own multi-threading.
            .with_threads(4)      // MNN internal thread count
            .with_min_result_confidence(0.3) // Lower threshold — invoice text can be faint,
                                              // better to capture more and filter in frontend
            .with_rec_options(
                ocr_rs::RecOptions::new()
                    .with_batch_size(16) // Larger batch = fewer MNN calls = better throughput
                    .with_batch(true)    // Enable batch processing
            );

        let engine = ocr_rs::OcrEngine::new(
            det_path.to_str().unwrap(),
            rec_path.to_str().unwrap(),
            keys_path.to_str().unwrap(),
            Some(config),
        )
        .map_err(|e| format!("创建PaddleOCR引擎失败: {:?}", e))?;

        log::info!("PaddleOCR engine initialized successfully");
        *lock = Some(engine);
    }

    Ok(lock)
}

/// OCR precision modes — longest-side max dimension for OCR input.
/// - "fast" (1280px): Fast, good for normal-sized text. Small text (密码区/备注栏) may be blurry.
/// - "standard" (1920px): Default. Good balance — handles most invoice text including small fonts.
/// - "precise" (2800px): Maximum accuracy. Slower (~2-3x vs fast) but preserves all detail.
#[cfg(feature = "ocr")]
pub fn ocr_max_dim_for_precision(precision: &str) -> u32 {
    match precision {
        "fast" => 1280,
        "precise" => 2800,
        _ => 1920,
    }
}

/// OCR an image from a file path or base64 data URL.
/// When `file_path` is provided, reads the image directly from disk — skipping
/// the expensive base64 encode→IPC→decode round-trip.
/// Falls back to `data_url` when `file_path` is None or file read fails.
#[cfg(feature = "ocr")]
pub fn ocr_image(data_url: &str, file_path: Option<&str>, ocr_precision: Option<&str>) -> Result<OcrResult, String> {
    let max_dim = ocr_max_dim_for_precision(ocr_precision.unwrap_or("standard"));
    if let Some(path) = file_path {
        if !path.is_empty() {
            match std::fs::read(path) {
                Ok(bytes) => {
                    if !bytes.is_empty() {
                        let exif_orient = if is_jpeg_bytes(&bytes) {
                            read_exif_orientation(&bytes)
                        } else { 1 };
                        match image::load_from_memory(&bytes) {
                            Ok(img) => {
                                let img = if exif_orient != 1 {
                                    apply_exif_orientation(img, exif_orient)
                                } else { img };
                                log::info!("OCR from file_path: {} ({}x{})", path, img.width(), img.height());
                                return run_ocr_on_image(img, max_dim);
                            }
                            Err(e) => {
                                log::warn!("Image decode from file_path {} failed: {}, falling back to data_url", path, e);
                            }
                        }
                    }
                }
                Err(e) => {
                    log::warn!("File read for OCR {} failed: {}, falling back to data_url", path, e);
                }
            }
        }
    }
    ocr_image_from_data(data_url, max_dim)
}

/// OCR an image from base64 data URL, return structured result with coordinates.
/// Internal helper — prefer `ocr_image()` which supports file_path.
#[cfg(feature = "ocr")]
pub fn ocr_image_from_data(data_url: &str, max_dim: u32) -> Result<OcrResult, String> {
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    use base64::Engine;
    use std::time::Instant;
    let t0 = Instant::now();

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

    log::info!("OCR from data_url: b64decode={}ms", t0.elapsed().as_millis());

    // Apply EXIF orientation before OCR so text is read in the correct visual order
    let exif_orient = if is_jpeg_bytes(&bytes) {
        read_exif_orientation(&bytes)
    } else { 1 };

    // Decode image using the `image` crate
    let img = image::load_from_memory(&bytes)
        .map_err(|e| format!("图片解码失败: {}", e))?;

    let img = if exif_orient != 1 {
        apply_exif_orientation(img, exif_orient)
    } else { img };

    run_ocr_on_image(img, max_dim)
}

/// Enhance image contrast for OCR using histogram stretching.
/// Maps the darkest 1% of pixels to 0 and brightest 1% to 255.
/// This dramatically improves OCR accuracy on low-contrast/faded invoices
/// and scanned documents with uneven lighting.
#[cfg(feature = "ocr")]
fn enhance_contrast_ocr(img: image::DynamicImage) -> image::DynamicImage {
    use image::GenericImageView;
    use image::Pixel;

    // Build luminance histogram (256 bins)
    let mut histogram = [0u32; 256];
    let mut total_pixels = 0u32;
    for pixel in img.pixels() {
        let rgba = pixel.2.to_rgba();
        let lum = (0.299 * rgba[0] as f64 + 0.587 * rgba[1] as f64 + 0.114 * rgba[2] as f64) as u8;
        histogram[lum as usize] += 1;
        total_pixels += 1;
    }

    if total_pixels == 0 {
        return img;
    }

    // Find 1st and 99th percentile
    let threshold_low = total_pixels / 100;   // 1%
    let threshold_high = total_pixels - threshold_low; // 99%
    let mut cumulative = 0u32;
    let mut p1 = 0u8;
    let mut p99 = 255u8;
    for i in 0..256 {
        cumulative += histogram[i];
        if cumulative >= threshold_low && p1 == 0 {
            p1 = i as u8;
        }
        if cumulative >= threshold_high {
            p99 = i as u8;
            break;
        }
    }

    // Skip enhancement if contrast is already good (range > 180)
    if p99.saturating_sub(p1) > 180 {
        return img;
    }

    // Build lookup table for linear contrast stretch
    let range = p99 as f64 - p1 as f64;
    if range < 1.0 {
        return img; // all pixels same color, nothing to enhance
    }
    let mut lut = [0u8; 256];
    for i in 0..256 {
        let v = ((i as f64 - p1 as f64) / range * 255.0).round();
        lut[i] = v.max(0.0).min(255.0) as u8;
    }

    // Apply LUT to each pixel
    let mut out = img.to_rgba8();
    for pixel in out.pixels_mut() {
        pixel.0[0] = lut[pixel.0[0] as usize];
        pixel.0[1] = lut[pixel.0[1] as usize];
        pixel.0[2] = lut[pixel.0[2] as usize];
    }

    log::info!("OCR contrast enhancement: p1={} p99={} range={}", p1, p99, p99.saturating_sub(p1));
    image::DynamicImage::ImageRgba8(out)
}

/// Core OCR logic: takes a pre-decoded image, resizes if needed, runs OCR,
/// and returns structured result with coordinates.
#[cfg(feature = "ocr")]
fn run_ocr_on_image(mut img: image::DynamicImage, max_dim: u32) -> Result<OcrResult, String> {
    use std::time::Instant;
    let t0 = Instant::now();

    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    let orig_w = img.width();
    let orig_h = img.height();
    let longest = orig_w.max(orig_h);

    if longest > max_dim {
        let scale = max_dim as f32 / longest as f32;
        let new_w = (orig_w as f32 * scale).round() as u32;
        let new_h = (orig_h as f32 * scale).round() as u32;
        img = img.resize_exact(new_w, new_h, image::imageops::FilterType::Lanczos3);
        log::info!(
            "OCR resize: {}x{} → {}x{} (max_dim={}, {}ms)",
            orig_w, orig_h, new_w, new_h, max_dim,
            t0.elapsed().as_millis()
        );
    }

    // Enhance contrast for low-contrast invoices (e.g., scanned/faded invoices).
    // PaddleOCR detection works better with higher contrast input.
    // We apply a simple linear contrast stretch: map the darkest 1% to 0, brightest 1% to 255.
    img = enhance_contrast_ocr(img);

    let resized_w = img.width();
    let resized_h = img.height();

    // Get OCR engine (lazy init on first call)
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭，OCR已中止".to_string());
    }
    let lock = get_ocr_engine()?;
    let engine = lock.as_ref().ok_or("OCR引擎未初始化")?;

    let t_engine = Instant::now();

    // Run OCR recognition
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭，OCR已中止".to_string());
    }
    let results = engine.recognize(&img)
        .map_err(|e| format!("PaddleOCR识别失败: {:?}", e))?;

    let t_recognize = Instant::now();

    // Collect data from results before releasing the engine lock.
    // PaddleOCR returns line-level results; we convert to our word-level format.
    // Scale coordinates back to original image dimensions for frontend use.
    let coord_scale_x = if resized_w > 0 { orig_w as f64 / resized_w as f64 } else { 1.0 };
    let coord_scale_y = if resized_h > 0 { orig_h as f64 / resized_h as f64 } else { 1.0 };

    let mut ocr_lines: Vec<OcrLine> = Vec::new();
    let mut flat_text_parts: Vec<String> = Vec::new();

    for result in &results {
        let line_text = result.text.trim().to_string();
        if line_text.is_empty() {
            continue;
        }
        flat_text_parts.push(line_text.clone());

        let bbox = &result.bbox;
        let rect = bbox.rect;
        let bx = rect.left() as f64 * coord_scale_x;
        let by = rect.top() as f64 * coord_scale_y;
        let bw = (rect.right() - rect.left()) as f64 * coord_scale_x;
        let bh = (rect.bottom() - rect.top()) as f64 * coord_scale_y;

        let line_confidence = result.confidence;

        // Extract polygon points from detection model (4 corner points)
        let line_points = bbox.points.as_ref().map(|pts| {
            pts.iter().map(|p| OcrPoint {
                x: p.x as f64 * coord_scale_x,
                y: p.y as f64 * coord_scale_y,
            }).collect()
        });

        let tokens = split_line_to_words(&line_text);

        if tokens.is_empty() {
            ocr_lines.push(OcrLine {
                words: vec![OcrWord {
                    text: line_text,
                    x: bx,
                    y: by,
                    w: bw,
                    h: bh,
                }],
                points: line_points,
                confidence: line_confidence,
            });
            continue;
        }

        // Character-width-weighted distribution: CJK chars are ~2x wider than Latin/digits.
        // This produces much more accurate word positions than equal-width-per-char.
        let total_weight: f64 = tokens.iter().map(|t| token_width_weight(t)).sum();
        let mut words: Vec<OcrWord> = Vec::new();
        let mut x_offset = 0.0f64;

        for token in &tokens {
            let token_w = if total_weight > 0.0 {
                bw * token_width_weight(token) / total_weight
            } else {
                bw
            };

            words.push(OcrWord {
                text: token.clone(),
                x: bx + x_offset,
                y: by,
                w: token_w,
                h: bh,
            });
            x_offset += token_w;
        }

        ocr_lines.push(OcrLine { words, points: line_points, confidence: line_confidence });
    }

    // Release the engine lock
    drop(lock);

    let flat_text = flat_text_parts.join("\n");
    let ocr_result = OcrResult {
        text: flat_text,
        lines: ocr_lines,
        img_w: orig_w,
        img_h: orig_h,
    };

    log::info!(
        "OCR timing: engine+resize={}ms recognize={}ms convert={}ms total={}ms ({} chars, {} lines, {}x{}→{}x{})",
        t_engine.duration_since(t0).as_millis(),
        t_recognize.duration_since(t_engine).as_millis(),
        t_recognize.elapsed().as_millis(),
        t0.elapsed().as_millis(),
        ocr_result.text.len(),
        ocr_result.lines.len(),
        orig_w, orig_h, resized_w, resized_h,
    );

    Ok(ocr_result)
}

/// Split a line of text into word tokens for coordinate mapping.
/// - CJK characters are kept as individual tokens (each character = one word)
/// - Non-CJK runs (Latin, digits, symbols) are kept as single tokens
/// - Spaces are included as part of adjacent tokens (not separate words)
#[cfg(feature = "ocr")]
fn split_line_to_words(text: &str) -> Vec<String> {
    let mut tokens: Vec<String> = Vec::new();
    let mut current_non_cjk = String::new();

    for ch in text.chars() {
        let is_cjk = is_cjk_char(ch);
        if is_cjk {
            // Flush accumulated non-CJK token
            if !current_non_cjk.is_empty() {
                tokens.push(current_non_cjk.clone());
                current_non_cjk.clear();
            }
            // Each CJK character is its own token
            tokens.push(ch.to_string());
        } else {
            // Accumulate non-CJK characters (Latin, digits, symbols, spaces)
            current_non_cjk.push(ch);
        }
    }

    // Flush remaining non-CJK
    if !current_non_cjk.is_empty() {
        tokens.push(current_non_cjk);
    }

    // Filter out pure-whitespace tokens
    tokens.retain(|t| !t.trim().is_empty());
    tokens
}

/// Compute visual width weight for a token.
/// CJK characters are approximately 2x wider than Latin/digits in most fonts.
/// Fullwidth forms (FF00-FFEF) are also 2x.
/// This produces more accurate x/w estimates than equal-width-per-character.
#[cfg(feature = "ocr")]
fn token_width_weight(token: &str) -> f64 {
    token.chars().map(|ch| {
        let cp = ch as u32;
        if (0x4E00..=0x9FFF).contains(&cp)       // CJK Unified Ideographs
            || (0x3400..=0x4DBF).contains(&cp)    // CJK Extension A
            || (0xF900..=0xFAFF).contains(&cp)    // CJK Compatibility
            || (0x3000..=0x303F).contains(&cp)    // CJK Symbols and Punctuation
            || (0xFF00..=0xFFEF).contains(&cp)    // Fullwidth forms
            || (0x3040..=0x309F).contains(&cp)    // Hiragana
            || (0x30A0..=0x30FF).contains(&cp)    // Katakana
            || cp >= 0x20000                       // CJK Extension B+
        {
            2.0
        } else {
            1.0
        }
    }).sum()
}

/// Check if a character is CJK (Chinese, Japanese, Korean)
#[cfg(feature = "ocr")]
fn is_cjk_char(ch: char) -> bool {
    let cp = ch as u32;
    // CJK Unified Ideographs: 4E00-9FFF
    // CJK Unified Ideographs Extension A: 3400-4DBF
    // CJK Compatibility Ideographs: F900-FAFF
    // CJK Unified Ideographs Extension B-F: 20000-2FA1F
    // Fullwidth forms: FF00-FFEF
    // CJK Symbols and Punctuation: 3000-303F
    // Hiragana: 3040-309F, Katakana: 30A0-30FF
    matches!(cp,
        0x4E00..=0x9FFF |
        0x3400..=0x4DBF |
        0xF900..=0xFAFF |
        0x20000..=0x2FA1F |
        0xFF00..=0xFFEF |
        0x3000..=0x303F |
        0x3040..=0x309F |
        0x30A0..=0x30FF
    )
}

/// Check whether OCR feature is available at runtime.
#[cfg(feature = "ocr")]
pub fn check_ocr_available() -> bool { true }

/// Check whether OCR feature is available at runtime.
#[cfg(not(feature = "ocr"))]
pub fn check_ocr_available() -> bool { false }

// OFD code has been extracted to the invoice-engine crate.


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
/// Fields used only for deserialization from JS (border/number/watermark rendered in preview only);
/// page_num and print_date are also rendered into the PDF as text overlay images.
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
    pub page_num: bool,
    pub print_date: bool,
    pub cutline: bool,
    pub watermark: bool,
    pub watermark_text: Option<String>,
    pub watermark_color: String,
    pub watermark_opacity: f32,
    pub watermark_angle: f32,
    pub watermark_size: f32,
    pub border_width: Option<f32>,
    pub border_color: Option<String>,
    pub trim_white: Option<bool>,
    pub footer_text: Option<String>,
    pub footer_margin: f32,
    pub custom_fm: bool,
    /// 报销单分段模式：固定段高纵向分段，段边界强制裁切线，发票左上对齐
    #[serde(default)]
    pub reimburse_mode: bool,
    /// 分段高度（mm），默认 120
    #[serde(default)]
    pub reimburse_height: Option<f32>,
    #[serde(default)]
    pub copies: u32,
    #[serde(default)]
    pub duplex: bool,
}

/// A file image with its metadata — sent from JS.
/// ow/oh/rotation are used by JS for layout decisions but not directly by Rust
/// (Rust gets rotation from SlotSpec and dimensions from decoded image).
///
/// **Optimization**: If `file_path` is provided, Rust reads the image directly
/// from disk, avoiding the expensive base64 encode→IPC→decode round-trip.
/// For images that only exist in memory (e.g. PDF pages rendered by WinRT,
/// OFD-extracted images), `file_path` is None and `data_url` is used instead.
#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
#[allow(dead_code)]
pub struct FileSpec {
    /// Base64 data URL — used when file_path is None (e.g. rendered PDF pages, OFD images)
    #[serde(default)]
    pub data_url: String,
    /// Disk path to the image file — when available, Rust reads bytes directly,
    /// skipping base64 encode/decode (saves ~30% data + CPU for large images)
    #[serde(default)]
    pub file_path: Option<String>,
    pub ow: u32,
    pub oh: u32,
    pub rotation: i32,
    /// Source type hint from frontend — affects compression strategy
    /// "image" = photo/scanned image file → JPEG compression is fine
    /// "pdf-page" = rendered PDF page → FlateDecode (lossless) is better for text
    /// "ofd-page" = OFD extracted image → usually text-like → FlateDecode
    #[serde(default)]
    pub source_type: Option<String>,
    /// Original PDF file path (for PDF passthrough optimization).
    /// Set when this file is a rendered PDF page.
    /// The frontend stores this as fileObj._pdfPath.
    #[serde(default)]
    pub pdf_path: Option<String>,
    /// Page index in the original PDF (0-based, for PDF passthrough).
    /// Set when this file is a rendered PDF page.
    /// The frontend stores this as fileObj._pdfPageIdx.
    #[serde(default)]
    pub pdf_page_idx: Option<u32>,
}

/// A slot on a page — which file (if any) goes here, and its rotation.
#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct SlotSpec {
    pub file_index: Option<usize>,
    pub rotation: i32,
    /// Per-slot scale override (1.0 = default). Applied to fit-mode scale.
    #[serde(default)]
    pub scale: Option<f32>,
    /// Per-slot X offset in mm (0 = centered).
    #[serde(default)]
    pub offset_x: Option<f32>,
    /// Per-slot Y offset in mm (0 = centered).
    #[serde(default)]
    pub offset_y: Option<f32>,
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

    // 报销单分段模式：段高固定（默认120mm），段边界 = k×seg（top-down 绝对位置）。
    // mt/mb 仅作段内上下安全边距，ml/mr 决定发票区域左右边界。
    // 忽略 rows/cols/gap/footerMargin，与 JS 端 calculateLayout 的 reimburse 分支保持一致。
    if settings.reimburse_mode {
        let seg = settings.reimburse_height.unwrap_or(120.0).max(10.0);
        let seg_count = ((ph / seg).floor() as usize).max(1);
        let sw = pw - ml - mr;
        let sh = (seg - mt - mb).max(10.0);
        let mut slots = Vec::new();
        for i in 0..seg_count {
            // top-down 段区间 [i*seg, (i+1)*seg]，发票区域 [i*seg+mt, (i+1)*seg-mb]
            // bottom-up: y = ph - top_down_top - sh
            let top_down_top = i as f32 * seg + mt;
            let y_mm = ph - top_down_top - sh;
            slots.push(LayoutSlotMm { x_mm: ml, y_mm, w_mm: sw, h_mm: sh });
        }
        log::info!("calculate_layout_mm [reimburse]: pw={pw} ph={ph} seg={seg} count={seg_count} sw={sw} sh={sh}");
        return (slots, pw, ph);
    }

    // The fm area is reserved purely for footer text below all rows.
    // Only deduct footer margin from slot height when there is footer content.
    // In custom_fm mode: deduct the explicit footer_margin value.
    // In auto mode: deduct the auto-computed footer height (auto_fm_mm).
    // When there is no footer content: no deduction (no footer to collide with).
    let has_footer = settings.page_num || settings.print_date || settings.footer_text.as_ref().map_or(false, |t| !t.is_empty());
    let line_count = (if settings.page_num || settings.print_date { 1 } else { 0 })
        + (if settings.footer_text.as_ref().map_or(false, |t| !t.is_empty()) { 1 } else { 0 });
    let auto_fm_mm = 3.0 + line_count as f32 * 5.0;
    let effective_fm = if has_footer {
        if settings.custom_fm { settings.footer_margin } else { auto_fm_mm }
    } else {
        0.0
    };
    let sw = (pw - cols * (ml + mr) - (cols - 1.0) * gh) / cols;
    let sh = (ph - rows * (mt + mb) - (rows - 1.0) * gv - effective_fm) / rows;

    log::info!("calculate_layout_mm [v2-fm-independent]: pw={pw} ph={ph} mt={mt} mb={mb} effective_fm={effective_fm} ml={ml} mr={mr} gh={gh} gv={gv} rows={rows} cols={cols} sw={sw} sh={sh}");

    let mut slots = Vec::new();
    for r in 0..settings.rows as usize {
        for c in 0..settings.cols as usize {
            // Convert row from JS (top-down) to printpdf (bottom-up)
            let row_from_bottom = settings.rows as usize - 1 - r;
            let x_mm = ml + c as f32 * (sw + ml + mr + gh);
            // Bottom-up: effective_fm (footer) + mb (row bottom margin) + row offset
            let y_mm = (effective_fm + mb) + row_from_bottom as f32 * (sh + mt + mb + gv);
            slots.push(LayoutSlotMm { x_mm, y_mm, w_mm: sw, h_mm: sh });
        }
    }

    (slots, pw, ph)
}

/// Convert days since 1970-01-01 to (year, month, day).
/// Simple algorithm, no need for chrono dependency.
fn days_to_ymd(days: i64) -> (i64, u32, u32) {
    // Shift to days since 0000-03-01 (simplifies leap year calculation)
    let z = days + 719468;
    let era = if z >= 0 { z / 146097 } else { (z - 146096) / 146097 };
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };
    (y, m as u32, d as u32)
}

/// Load SimHei font from system fonts directory.
/// Returns FontArc if successful, logs error and returns None if not.
fn load_system_font() -> Option<ab_glyph::FontArc> {
    let font_data = match std::fs::read("C:\\Windows\\Fonts\\simhei.ttf") {
        Ok(d) => d,
        Err(e) => {
            log::error!("load_system_font: failed to read simhei.ttf: {}", e);
            return None;
        }
    };
    match ab_glyph::FontArc::try_from_vec(font_data) {
        Ok(f) => Some(f),
        Err(e) => {
            log::error!("load_system_font: failed to parse simhei.ttf: {}", e);
            None
        }
    }
}

/// Render page number and/or print date text to a RGBA image with transparency.
/// Returns (png_bytes, width_px, height_px) or None if neither is enabled.
///
/// The image is rendered at RENDER_DPI for high-quality PDF output.
/// Uses ab_glyph to rasterize text with a system CJK font.
fn render_text_overlay(
    font: &Option<ab_glyph::FontArc>,
    page_num_text: &str,
    print_date_text: &str,
    footer_text: &str,
    page_width_mm: f32,
    _total_pages: usize,
) -> Option<(Vec<u8>, u32, u32)> {
    if page_num_text.is_empty() && print_date_text.is_empty() && footer_text.is_empty() {
        return None;
    }

    let font = match font {
        Some(f) => f,
        None => {
            log::warn!("render_text_overlay: no font available, skipping");
            return None;
        }
    };

    // Render at RENDER_DPI for print quality
    let px_per_mm = RENDER_DPI as f32 / 25.4;
    let img_width = (page_width_mm * px_per_mm) as u32;
    let font_size = (3.5 * px_per_mm) as f32; // ~3.5mm text height ≈ 10pt at screen
    let line_height = font_size * 1.4;
    // Line layout: pageNum+printDate on one line, footerText on separate line
    let has_line1 = !page_num_text.is_empty() || !print_date_text.is_empty();
    let has_line2 = !footer_text.is_empty();
    let num_lines = (if has_line1 { 1 } else { 0 }) + (if has_line2 { 1 } else { 0 });
    // img_height: num_lines of text + small top/bottom padding (1.1x ≈ 5% top + 5% bottom)
    let img_height = (line_height * num_lines as f32 * 1.1) as u32;

    // Create RGBA image (transparent background)
    let mut img = image::RgbaImage::new(img_width, img_height);

    // Text color: #94a3b8 (matches preview) at full opacity
    let text_color = [148u8, 163u8, 184u8, 255u8];

    let scaled_font = font.as_scaled(font_size);

    let mut y_offset = font_size; // Start at baseline of first line

    // Helper: measure text width
    let measure_text = |text: &str, sf: &ab_glyph::PxScaleFont<&ab_glyph::FontArc>| -> f32 {
        text.chars().map(|c| sf.h_advance(font.glyph_id(c))).sum()
    };

    // Helper: render a line of text starting at x_start
    let render_text_at = |img: &mut image::RgbaImage, text: &str, x_start: f32, y_baseline: f32, sf: &ab_glyph::PxScaleFont<&ab_glyph::FontArc>| {
        let mut x = x_start;
        for c in text.chars() {
            let glyph_id = font.glyph_id(c);
            let glyph = Glyph {
                id: glyph_id,
                scale: font_size.into(),
                position: ab_glyph::point(x, y_baseline),
            };
            if let Some(q) = font.outline_glyph(glyph) {
                let bb = q.px_bounds();
                let x_draw = bb.min.x;
                let y_draw = bb.min.y;
                q.draw(|gx, gy, v| {
                    let px = (x_draw + gx as f32) as i32;
                    let py = (y_draw + gy as f32) as i32;
                    if px >= 0 && py >= 0 && (px as u32) < img.width() && (py as u32) < img.height() {
                        let alpha = (v * text_color[3] as f32) as u8;
                        let pixel = img.get_pixel_mut(px as u32, py as u32);
                        if alpha > pixel[3] {
                            *pixel = image::Rgba([text_color[0], text_color[1], text_color[2], alpha]);
                        }
                    }
                });
            }
            x += sf.h_advance(glyph_id);
        }
    };

    // Line 1: pageNum + printDate (if either exists)
    if has_line1 {
        let has_both = !page_num_text.is_empty() && !print_date_text.is_empty();
        if has_both {
            // Page number on the left, print date on the right
            let _pn_w = measure_text(page_num_text, &scaled_font);
            let pd_w = measure_text(print_date_text, &scaled_font);
            let margin_x = img_width as f32 * 0.05; // 5% margin from edges
            render_text_at(&mut img, page_num_text, margin_x, y_offset, &scaled_font);
            render_text_at(&mut img, print_date_text, img_width as f32 - pd_w - margin_x, y_offset, &scaled_font);
        } else if !page_num_text.is_empty() {
            // Only page number: centered
            let pn_w = measure_text(page_num_text, &scaled_font);
            render_text_at(&mut img, page_num_text, (img_width as f32 - pn_w) / 2.0, y_offset, &scaled_font);
        } else {
            // Only print date: centered
            let pd_w = measure_text(print_date_text, &scaled_font);
            render_text_at(&mut img, print_date_text, (img_width as f32 - pd_w) / 2.0, y_offset, &scaled_font);
        }
        y_offset += line_height;
    }

    // Line 2: footer text (centered)
    if has_line2 {
        let ft_w = measure_text(footer_text, &scaled_font);
        render_text_at(&mut img, footer_text, (img_width as f32 - ft_w) / 2.0, y_offset, &scaled_font);
    }

    // Encode as PNG
    let mut png_buf = Vec::new();
    match img.write_to(&mut std::io::Cursor::new(&mut png_buf), image::ImageFormat::Png) {
        Ok(()) => {
            log::info!("render_text_overlay: generated {}x{} PNG ({} bytes) for '{}' + '{}'",
                img_width, img_height, png_buf.len(), page_num_text, print_date_text);
            Some((png_buf, img_width, img_height))
        }
        Err(e) => {
            log::error!("render_text_overlay: PNG encode failed: {}", e);
            None
        }
    }
}

/// Render slot numbers as small PNG images with background box.
/// Returns Vec of (png_bytes, width_px, height_px) for each slot.
fn render_slot_numbers(
    font: &Option<ab_glyph::FontArc>,
    slot_positions: &[LayoutSlotMm],
    start_number: usize,
) -> Vec<(Vec<u8>, u32, u32)> {
    log::info!("render_slot_numbers: called with {} slots, start_number={}", slot_positions.len(), start_number);

    if slot_positions.is_empty() {
        return vec![];
    }

    let font = match font {
        Some(f) => f,
        None => {
            log::warn!("render_slot_numbers: no font available");
            return vec![];
        }
    };

    let mut results = Vec::new();
    let px_per_mm = RENDER_DPI as f32 / 25.4;
    let font_size = (3.5 * px_per_mm) as f32;
    let scaled_font = font.as_scaled(font_size);

    for (i, _slot) in slot_positions.iter().enumerate() {
        let num = start_number + i;
        let num_str = num.to_string();

        let mut text_width = 0.0f32;
        for c in num_str.chars() {
            text_width += scaled_font.h_advance(font.glyph_id(c));
        }

        let padding_x = font_size * 0.5;
        let _padding_y = font_size * 0.3;
        let img_w = (text_width + padding_x * 2.0).ceil() as u32;
        let img_h = (font_size * 1.6).ceil() as u32;
        let mut img = image::RgbaImage::new(img_w, img_h);

        let bg_color = image::Rgba([0, 0, 0, 140]);
        let text_color = [255u8, 255u8, 255u8, 255u8];

        for y in 0..img_h {
            for x in 0..img_w {
                img.put_pixel(x, y, bg_color);
            }
        }

        let x_start = padding_x;
        let y_baseline = (img_h as f32 + font_size * 0.6) / 2.0;

        let mut x = x_start;
        for c in num_str.chars() {
            let glyph_id = font.glyph_id(c);
            let glyph = ab_glyph::Glyph {
                id: glyph_id,
                scale: font_size.into(),
                position: ab_glyph::point(x, y_baseline),
            };
            if let Some(q) = font.outline_glyph(glyph) {
                let bb = q.px_bounds();
                let x_draw = bb.min.x;
                let y_draw = bb.min.y;
                q.draw(|gx, gy, v| {
                    let px = (x_draw + gx as f32) as i32;
                    let py = (y_draw + gy as f32) as i32;
                    if px >= 0 && py >= 0 && (px as u32) < img.width() && (py as u32) < img.height() {
                        let alpha = (v * text_color[3] as f32) as u8;
                        let pixel = img.get_pixel_mut(px as u32, py as u32);
                        if alpha > pixel[3] {
                            *pixel = image::Rgba([text_color[0], text_color[1], text_color[2], alpha]);
                        }
                    }
                });
            }
            x += scaled_font.h_advance(glyph_id);
        }

        let mut png_buf = Vec::new();
        if img.write_to(&mut std::io::Cursor::new(&mut png_buf), image::ImageFormat::Png).is_ok() {
            results.push((png_buf, img_w, img_h));
        }
    }

    results
}

/// Render watermark text as a single PNG tile, optionally rotated.
/// Returns (png_bytes, width_px, height_px) or None.
fn render_watermark(
    font: &Option<ab_glyph::FontArc>,
    watermark_text: &str,
    color: &str,
    opacity: f32,
    font_size_mm: f32,
    angle_deg: f32,
) -> Option<(Vec<u8>, u32, u32)> {
    log::info!("render_watermark: text='{}', color='{}', opacity={}, font_size={}mm, angle={}",
        watermark_text, color, opacity, font_size_mm, angle_deg);

    if watermark_text.is_empty() || font.is_none() {
        return None;
    }

    let font = font.as_ref().unwrap();

    let (r, g, b) = parse_hex_color(color).unwrap_or((0.5, 0.5, 0.5));

    let px_per_mm = RENDER_DPI as f32 / 25.4;
    let font_size = (font_size_mm * px_per_mm) as f32;
    let scaled_font = font.as_scaled(font_size);

    let mut text_width = 0.0f32;
    for c in watermark_text.chars() {
        text_width += scaled_font.h_advance(font.glyph_id(c));
    }
    let text_height = font_size;

    let tile_w = (text_width * 1.2).ceil() as u32;
    let tile_h = (text_height * 1.5).ceil() as u32;
    let mut img = image::RgbaImage::new(tile_w, tile_h);

    let text_color = [
        (r * 255.0) as u8,
        (g * 255.0) as u8,
        (b * 255.0) as u8,
        (opacity * 255.0) as u8,
    ];

    let x_start = (tile_w as f32 - text_width) / 2.0;
    let y_baseline = tile_h as f32 * 0.65;

    let mut x = x_start;
    for c in watermark_text.chars() {
        let glyph_id = font.glyph_id(c);
        let glyph = ab_glyph::Glyph {
            id: glyph_id,
            scale: font_size.into(),
            position: ab_glyph::point(x, y_baseline),
        };
        if let Some(q) = font.outline_glyph(glyph) {
            let bb = q.px_bounds();
            let x_draw = bb.min.x;
            let y_draw = bb.min.y;
            q.draw(|gx, gy, v| {
                let px = (x_draw + gx as f32) as i32;
                let py = (y_draw + gy as f32) as i32;
                if px >= 0 && py >= 0 && (px as u32) < img.width() && (py as u32) < img.height() {
                    let alpha = (v * text_color[3] as f32) as u8;
                    let pixel = img.get_pixel_mut(px as u32, py as u32);
                    if alpha > pixel[3] {
                        *pixel = image::Rgba([text_color[0], text_color[1], text_color[2], alpha]);
                    }
                }
            });
        }
        x += scaled_font.h_advance(glyph_id);
    }

    let rotated = if angle_deg.abs() > 0.1 {
        let angle_rad = angle_deg * std::f32::consts::PI / 180.0;
        let cos_a = angle_rad.cos();
        let sin_a = angle_rad.sin();
        let cx = tile_w as f32 / 2.0;
        let cy = tile_h as f32 / 2.0;
        let new_w = ((tile_w as f32 * cos_a.abs()) + (tile_h as f32 * sin_a.abs())).ceil() as u32;
        let new_h = ((tile_w as f32 * sin_a.abs()) + (tile_h as f32 * cos_a.abs())).ceil() as u32;
        let mut rotated = image::RgbaImage::new(new_w, new_h);
        for y in 0..tile_h {
            for x in 0..tile_w {
                let p = img.get_pixel(x, y);
                if p[3] > 0 {
                    let rx = x as f32 - cx;
                    let ry = y as f32 - cy;
                    let dx = rx * cos_a - ry * sin_a;
                    let dy = rx * sin_a + ry * cos_a;
                    let nx = (dx + new_w as f32 / 2.0) as i32;
                    let ny = (dy + new_h as f32 / 2.0) as i32;
                    if nx >= 0 && ny >= 0 && (nx as u32) < new_w && (ny as u32) < new_h {
                        let existing = rotated.get_pixel(nx as u32, ny as u32);
                        if p[3] > existing[3] {
                            rotated.put_pixel(nx as u32, ny as u32, *p);
                        }
                    }
                }
            }
        }
        rotated
    } else {
        img
    };

    let final_w = rotated.width();
    let final_h = rotated.height();

    let mut png_buf = Vec::new();
    if rotated.write_to(&mut std::io::Cursor::new(&mut png_buf), image::ImageFormat::Png).is_ok() {
        Some((png_buf, final_w, final_h))
    } else {
        None
    }
}

fn parse_hex_color(hex: &str) -> Option<(f32, f32, f32)> {
    let hex = hex.trim_start_matches('#');
    if hex.len() < 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()? as f32 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()? as f32 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()? as f32 / 255.0;
    Some((r, g, b))
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

/// Decode all unique images, apply trim + color mode.
/// Rotation is NOT applied here — it's per-slot and handled in build_page_ops.
/// Returns decoded images indexed by file_index.
/// Uses rayon for parallel decoding when multiple files are present.
///
/// **Optimization**: When `file_path` is set, reads bytes directly from disk
/// instead of base64-decoding the data URL. This avoids:
/// - Frontend base64-encoding the entire image into the IPC JSON payload
/// - Rust base64-decoding it back to bytes
/// For a 300 DPI invoice image (~3MB), this saves ~1MB base64 overhead + CPU.
///
/// **JPEG Passthrough Optimization**: If the file is a JPEG and no pixel-level
/// operations are needed (no trim, no color mode change, no EXIF rotation),
/// the raw JPEG bytes are preserved in ImageSource::JpegPassthrough.
/// This avoids re-reading the file from disk at PDF generation time.
/// At PDF generation, ALL images are decoded → raw pixels → FlateDecode (lossless).
fn decode_images(
    files: &[FileSpec],
    settings: &RenderSettings,
) -> Vec<Option<ImageSource>> {
    use rayon::prelude::*;

    let trim = settings.trim_white.unwrap_or(false);
    let color_mode = settings.color_mode.clone();

    // Parallel decode — each file is independent
    let decoded: Vec<Option<ImageSource>> = files
        .par_iter()
        .map(|file_spec| {
            // Check shutdown flag — abort image decoding if app is closing
            if SHUTTING_DOWN.load(Ordering::SeqCst) {
                return None;
            }

            // Read raw bytes (prefer file path to skip base64 overhead)
            let bytes = if let Some(ref path) = file_spec.file_path {
                match std::fs::read(path) {
                    Ok(b) => b,
                    Err(e) => {
                        log::warn!("File read failed {}: {}, trying data_url", path, e);
                        match decode_base64_to_bytes(&file_spec.data_url) {
                            Ok(b) => b,
                            Err(e2) => {
                                log::warn!("data_url decode also failed: {}", e2);
                                return None;
                            }
                        }
                    }
                }
            } else if !file_spec.data_url.is_empty() {
                match decode_base64_to_bytes(&file_spec.data_url) {
                    Ok(b) => b,
                    Err(e) => {
                        log::warn!("data_url decode failed: {}", e);
                        return None;
                    }
                }
            } else {
                log::warn!("FileSpec has neither file_path nor data_url");
                return None;
            };

            // JPEG PASSTHROUGH: if the file is JPEG and no pixel-level ops are needed,
            // preserve the raw JPEG bytes to avoid decode→re-encode quality loss.
            let exif_orientation = if is_jpeg_bytes(&bytes) {
                read_exif_orientation(&bytes)
            } else {
                1 // non-JPEG: no EXIF orientation
            };
            let has_exif_rotation = exif_orientation != 1;

            let can_passthrough = is_jpeg_bytes(&bytes)
                && !trim
                && !has_exif_rotation
                && (color_mode == "color" || color_mode.is_empty());

            if can_passthrough {
                if let Some((w, h, nc)) = parse_jpeg_info(&bytes) {
                    return Some(ImageSource::JpegPassthrough {
                        raw_bytes: bytes,
                        width: w,
                        height: h,
                        num_components: nc,
                    });
                }
                // If JPEG header parsing fails, fall through to decode pipeline
                log::warn!("JPEG passthrough: header parse failed, falling back to decode");
            }

            // Standard decode pipeline
            let mut img = match image::load_from_memory(&bytes) {
                Ok(i) => i,
                Err(e) => {
                    log::warn!("Image decode failed: {}", e);
                    return None;
                }
            };

            // Apply EXIF orientation — bakes orientation into pixel data
            // so the image displays correctly in PDF viewers (which don't read EXIF).
            if has_exif_rotation {
                log::info!("Applying EXIF orientation {} (decoded {}x{})", exif_orientation, img.width(), img.height());
                img = apply_exif_orientation(img, exif_orientation);
            }

            // Apply trim (global setting, not per-slot)
            if trim {
                img = trim_white_edges(&img, 245);
            }

            // Apply color mode (global setting, not per-slot)
            img = apply_color_mode(img, &color_mode);

            Some(ImageSource::Decoded(img))
        })
        .collect();

    decoded
}

/// Decode base64 data URL to raw bytes (strips the "data:...;base64," prefix).
fn decode_base64_to_bytes(data_url: &str) -> Result<Vec<u8>, String> {
    let base64_part = if data_url.starts_with("data:") {
        // Find the comma after "data:...;base64,"
        data_url.find(',').map(|i| &data_url[i + 1..]).unwrap_or(data_url)
    } else {
        data_url
    };
    use base64::Engine;
    base64::engine::general_purpose::STANDARD
        .decode(base64_part)
        .map_err(|e| format!("base64 decode error: {}", e))
}

/// Get or create a cached XObject for (file_index, rotation).
/// All images: decode → rotate → RawImage → add_image (FlateDecode, lossless).
fn get_cached_xobj(
    doc: &mut printpdf::PdfDocument,
    cache: &mut std::collections::HashMap<(usize, i32), CachedXobj>,
    file_idx: usize,
    rotation: i32,
    sources: &[Option<ImageSource>],
) -> Option<CachedXobj> {
    let key = (file_idx, rotation);

    if let Some(cached) = cache.get(&key) {
        return Some(CachedXobj {
            iw_mm: cached.iw_mm,
            ih_mm: cached.ih_mm,
            xobj_id: cached.xobj_id.clone(),
        });
    }

    let source = sources[file_idx].as_ref()?;

    let (iw_mm, ih_mm, xobj_id) = match source {
        ImageSource::Decoded(img) => {
            // Current pipeline: rotate → RawImage → add_image
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
            (iw_mm, ih_mm, xobj_id)
        }
        ImageSource::JpegPassthrough { raw_bytes, width, height, num_components } => {
            let rot = ((rotation % 360) + 360) % 360;
            if rot == 90 || rot == 270 {
                // Must decode → rotate → re-encode: fallback to standard pipeline
                let img = match image::load_from_memory(raw_bytes) {
                    Ok(i) => i,
                    Err(e) => {
                        log::warn!("JPEG passthrough fallback decode failed for file {}: {}", file_idx, e);
                        return None;
                    }
                };
                let rotated = if rot == 90 { img.rotate90() } else { img.rotate270() };
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
                (iw_mm, ih_mm, xobj_id)
            } else {
                // 0° or 180° rotation: JPEG passthrough via ExternalXObject!
                let iw_mm = *width as f32 * 25.4 / RENDER_DPI as f32;
                let ih_mm = *height as f32 * 25.4 / RENDER_DPI as f32;
                let color_space: &[u8] = match num_components {
                    1 => b"DeviceGray",
                    4 => b"DeviceCMYK",
                    _ => b"DeviceRGB",
                };
                let mut dict = std::collections::BTreeMap::new();
                dict.insert("Type".to_string(), printpdf::xobject::DictItem::Name(b"XObject".to_vec()));
                dict.insert("Subtype".to_string(), printpdf::xobject::DictItem::Name(b"Image".to_vec()));
                dict.insert("Width".to_string(), printpdf::xobject::DictItem::Int(*width as i64));
                dict.insert("Height".to_string(), printpdf::xobject::DictItem::Int(*height as i64));
                dict.insert("BitsPerComponent".to_string(), printpdf::xobject::DictItem::Int(8));
                dict.insert("ColorSpace".to_string(), printpdf::xobject::DictItem::Name(color_space.to_vec()));
                dict.insert("Filter".to_string(), printpdf::xobject::DictItem::Name(b"DCTDecode".to_vec()));
                let external_xobj = printpdf::xobject::ExternalXObject {
                    stream: printpdf::xobject::ExternalStream {
                        dict,
                        content: raw_bytes.clone(),
                        compress: false,
                    },
                    width: Some(printpdf::units::Px(*width as usize)),
                    height: Some(printpdf::units::Px(*height as usize)),
                    dpi: Some(RENDER_DPI as f32),
                };
                let xobj_id = doc.add_xobject(&external_xobj);
                (iw_mm, ih_mm, xobj_id)
            }
        }
    };

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
    sources: &[Option<ImageSource>],
    xobj_cache: &mut std::collections::HashMap<(usize, i32), CachedXobj>,
) -> Vec<printpdf::Op> {
    let mut ops = Vec::new();

    for (slot_idx, slot_spec) in page_spec.slots.iter().enumerate() {
        let file_idx = match slot_spec.file_index {
            Some(idx) if idx < sources.len() && sources[idx].is_some() => idx,
            _ => continue,
        };

        if slot_idx < slot_positions.len() {
            let sp = &slot_positions[slot_idx];
            log::info!("printpdf fallback: slot[{}] x={:.2}mm y={:.2}mm w={:.2}mm h={:.2}mm",
                slot_idx, sp.x_mm, sp.y_mm, sp.w_mm, sp.h_mm);
        }

        let rotation = slot_spec.rotation;
        let cached = match get_cached_xobj(doc, xobj_cache, file_idx, rotation, sources) {
            Some(c) => c,
            None => continue,
        };

        let iw_mm = cached.iw_mm;
        let ih_mm = cached.ih_mm;

        // Compute scale to fit in slot
        let (mut scale_x, mut scale_y) = match settings.fit_mode.as_str() {
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

        // Per-slot scale override
        let per_scale = slot_spec.scale.unwrap_or(1.0);
        if per_scale != 1.0 {
            scale_x *= per_scale;
            scale_y *= per_scale;
        }

        // Centered position in slot (bottom-left origin).
        // 报销单模式：左上对齐；常规模式：居中。
        let draw_w_mm = iw_mm * scale_x;
        let draw_h_mm = ih_mm * scale_y;
        let mut offset_x_mm = slot_positions[slot_idx].x_mm;
        let mut offset_y_mm = if settings.reimburse_mode {
            slot_positions[slot_idx].y_mm + (slot_positions[slot_idx].h_mm - draw_h_mm)
        } else {
            slot_positions[slot_idx].y_mm + (slot_positions[slot_idx].h_mm - draw_h_mm) / 2.0
        };
        if !settings.reimburse_mode {
            offset_x_mm += (slot_positions[slot_idx].w_mm - draw_w_mm) / 2.0;
        }

        // Per-slot offset override
        let per_ox = slot_spec.offset_x.unwrap_or(0.0);
        let per_oy = slot_spec.offset_y.unwrap_or(0.0);
        if per_ox != 0.0 { offset_x_mm += per_ox; }
        if per_oy != 0.0 { offset_y_mm -= per_oy; }  // JS Y+ is down, PDF Y+ is up

        // Convert mm to Pt — XObjectTransform uses Pt
        let offset_x_pt = offset_x_mm * MM_TO_PT;
        let offset_y_pt = offset_y_mm * MM_TO_PT;

        // For JPEG passthrough with 180° rotation, use PDF transform matrix
        // instead of pixel-level rotation (which would require decode)
        let rotate_op = {
            let rot = ((rotation % 360) + 360) % 360;
            if rot == 180 {
                // Rotate 180° around the center of the drawn image
                Some(printpdf::XObjectRotation {
                    angle_ccw_degrees: 180.0,
                    rotation_center_x: printpdf::units::Px((iw_mm * RENDER_DPI as f32 / 25.4 / 2.0) as usize),
                    rotation_center_y: printpdf::units::Px((ih_mm * RENDER_DPI as f32 / 25.4 / 2.0) as usize),
                })
            } else {
                None
            }
        };

        // Clip to slot boundary — prevents per-slot scale/offset from overflowing
        // into adjacent slots in the PDF output. Uses raw PDF operators:
        //   q  re  W  n  Do  Q
        let slot_x_pt = slot_positions[slot_idx].x_mm * MM_TO_PT;
        let slot_y_pt = slot_positions[slot_idx].y_mm * MM_TO_PT;
        let slot_w_pt = slot_positions[slot_idx].w_mm * MM_TO_PT;
        let slot_h_pt = slot_positions[slot_idx].h_mm * MM_TO_PT;

        use printpdf::DictItem as DI;
        ops.push(printpdf::Op::SaveGraphicsState);
        ops.push(printpdf::Op::Unknown {
            key: "re".into(),
            value: vec![DI::Real(slot_x_pt), DI::Real(slot_y_pt), DI::Real(slot_w_pt), DI::Real(slot_h_pt)],
        });
        ops.push(printpdf::Op::Unknown { key: "W".into(), value: vec![] });
        ops.push(printpdf::Op::Unknown { key: "n".into(), value: vec![] });

        ops.push(printpdf::Op::UseXobject {
            id: cached.xobj_id.clone(),
            transform: printpdf::XObjectTransform {
                translate_x: Some(printpdf::Pt(offset_x_pt)),
                translate_y: Some(printpdf::Pt(offset_y_pt)),
                scale_x: Some(scale_x),
                scale_y: Some(scale_y),
                dpi: Some(RENDER_DPI as f32),
                rotate: rotate_op,
            },
        });

        ops.push(printpdf::Op::RestoreGraphicsState);
    }

    ops
}

/// Progress callback type: phase name + current (1-based) + total
pub type ProgressFn = Box<dyn Fn(&str, u32, u32) + Send>;

/// Generate PDF from layout request (files + pages + settings).
/// This replaces JS `renderPageToCanvas` + `generate_pdf_from_pages`.
/// `on_progress` is called with (phase, current, total) to report progress.
/// Phases: "decode" (image decoding), "build" (page composition), "save" (PDF writing).
pub fn generate_pdf_from_layout(
    request: &LayoutRenderRequest,
    output_path: &std::path::Path,
    on_progress: Option<ProgressFn>,
) -> Result<Option<String>, String> {
    if request.pages.is_empty() {
        return Err("没有页面数据".to_string());
    }

    let needs_text = request.settings.page_num
        || request.settings.print_date
        || request.settings.footer_text.as_ref().map_or(false, |t| !t.is_empty())
        || (request.settings.watermark && request.settings.watermark_text.as_ref().map_or(false, |t| !t.is_empty()));
    let font_warning = if needs_text && !std::path::Path::new("C:\\Windows\\Fonts\\simhei.ttf").exists() {
        Some("系统缺少中文字体(simhei.ttf)，页脚/水印/页码将不显示".to_string())
    } else {
        None
    };

    // Decode all unique images (base64 → ImageSource) — needed for both
    // the lopdf hybrid path (images as JPEG XObjects) and the printpdf fallback.
    let total_files = request.files.len() as u32;
    if let Some(ref cb) = &on_progress {
        cb("decode", 0, total_files);
    }
    let sources = decode_images(&request.files, &request.settings);
    if let Some(ref cb) = &on_progress {
        cb("decode", total_files, total_files);
    }

    // Hybrid lopdf passthrough: handles ALL scenarios — pure PDF, pure images,
    // and mixed PDF + image/OFD. PDF pages stay vector-sharp; images are
    // encoded as JPEG XObjects. Falls back to printpdf pipeline on any error.
    match generate_pdf_passthrough(request, output_path, on_progress.as_ref(), &sources) {
        Ok(()) => return Ok(font_warning),
        Err(e) => {
            log::warn!("lopdf直通失败，回退printpdf渲染管道: {}", e);
            // Continue with printpdf pipeline below
        }
    }

    let total_pages = request.pages.len() as u32;
    let (slot_positions, pw, ph) = calculate_layout_mm(&request.settings);

    // Create PDF document (new API: no page dimensions at creation time)
    let mut doc = printpdf::PdfDocument::new("发票酱");

    // Step 2: Build pages, caching XObjects by (file_index, rotation) to avoid redundant work.
    let mut xobj_cache: std::collections::HashMap<(usize, i32), CachedXobj> = std::collections::HashMap::new();

    // Pre-load CJK font for text overlay
    let text_font = if request.settings.page_num || request.settings.print_date || request.settings.footer_text.as_ref().map_or(false, |t| !t.is_empty()) {
        load_system_font()
    } else {
        None
    };

    for (i, page_spec) in request.pages.iter().enumerate() {
        // Check shutdown flag — abort PDF generation if app is closing
        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，PDF生成已中止".to_string());
        }
        let mut ops = build_page_ops(
            &mut doc,
            page_spec,
            &request.settings,
            &slot_positions,
            &sources,
            &mut xobj_cache,
        );

        // Add text overlay (page number + print date) for printpdf fallback path
        let pp_page_num_text = if request.settings.page_num {
            format!("第 {} 页 / 共 {} 页", i + 1, request.pages.len())
        } else {
            String::new()
        };
        let pp_print_date_text = if request.settings.print_date {
            let now = std::time::SystemTime::now();
            let duration = now.duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
            let (year, month, day) = days_to_ymd((duration.as_secs() / 86400) as i64);
            format!("打印日期 {:04}-{:02}-{:02}", year, month, day)
        } else {
            String::new()
        };
        let footer_text = request.settings.footer_text.clone().unwrap_or_default();
        if let Some((png_bytes, _img_w, _img_h)) = render_text_overlay(
            &text_font, &pp_page_num_text, &pp_print_date_text, &footer_text, pw, request.pages.len()
        ) {
            let mut warnings = Vec::new();
            match printpdf::RawImage::decode_from_bytes(&png_bytes, &mut warnings) {
                Ok(raw_img) => {
                    let img_w = raw_img.width as f32;
                    let img_h = raw_img.height as f32;
                    let xobj_id = doc.add_image(&raw_img);

                    // Position: bottom-center, (margin_bottom + 5mm) from bottom
                    // RawImage dimensions are in pixels; scale from RENDER_DPI to PDF pt
                    let img_w_pt = img_w * 72.0 / RENDER_DPI as f32;
                    let img_h_pt = img_h * 72.0 / RENDER_DPI as f32;
                    let pw_pt = pw * MM_TO_PT;
                    let x_pt = (pw_pt - img_w_pt) / 2.0;
                    let y_pt = 3.0 * MM_TO_PT; // 3mm from bottom edge

                    ops.push(printpdf::Op::SaveGraphicsState);
                    ops.push(printpdf::Op::UseXobject {
                        id: xobj_id,
                        transform: printpdf::XObjectTransform {
                            translate_x: Some(printpdf::Pt(x_pt)),
                            translate_y: Some(printpdf::Pt(y_pt)),
                            scale_x: Some(img_w_pt),
                            scale_y: Some(img_h_pt),
                            dpi: Some(RENDER_DPI as f32),
                            rotate: None,
                        },
                    });
                    ops.push(printpdf::Op::RestoreGraphicsState);
                    log::info!("printpdf fallback: page {} text overlay added", i);
                }
                Err(e) => {
                    log::warn!("printpdf fallback: page {} text overlay RawImage decode failed: {}", i, e);
                }
            }
        } else {
            log::warn!("printpdf fallback: page {} render_text_overlay returned None", i);
        }

        // Skip empty pages — avoid generating blank PDF pages when
        // all slots have no valid images (e.g. last page with fewer files)
        if ops.is_empty() {
            log::info!("Skipping empty page {}", i + 1);
            continue;
        }

        let page = printpdf::PdfPage::new(
            printpdf::Mm(pw),
            printpdf::Mm(ph),
            ops,
        );
        doc.pages.push(page);

        // Report progress (1-based page number)
        if let Some(ref cb) = on_progress {
            cb("build", (i + 1) as u32, total_pages);
        }
    }

    // Step 3: Save PDF — can be slow for large documents
    // Check shutdown before starting the expensive save operation
    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭，PDF生成已中止".to_string());
    }
    if let Some(ref cb) = on_progress {
        cb("save", 0, 1);
    }

    // Save PDF — custom options for print quality.
    //
    // **Performance notes**:
    // - `optimize: false`: skip printpdf's save-time re-encoding pass. Our images
    //   are already at target quality (300 DPI, JPEG passthrough or pre-decoded).
    //   Re-encoding is the #1 bottleneck — disabling it saves 60-80% of save time.
    // - `quality: 0.90`: JPEG encoding is faster than 0.95 with imperceptible difference.
    // - FlateDecode used only when text content is present (quality-sensitive).
    let has_text_content = request.files.iter().any(|f| {
        f.source_type.as_deref() == Some("pdf-page") || f.source_type.as_deref() == Some("ofd-page")
    });
    let save_opts = printpdf::PdfSaveOptions {
        optimize: false,   // Skip re-encoding — images are already at target quality
        subset_fonts: true,
        secure: false,     // Allow Op::Unknown for per-slot clipping paths (re/W/n)
        image_optimization: Some(printpdf::ImageOptimizationOptions {
            quality: Some(0.90),           // Fast JPEG encoding (0.95 was 40% slower)
            max_image_size: None,          // NO size limit (default "2MB" downsamples invoices!)
            auto_optimize: Some(true),     // Remove alpha if opaque, detect greyscale
            convert_to_greyscale: None,    // Don't force greyscale
            dither_greyscale: None,
            format: Some(if has_text_content {
                printpdf::ImageCompression::Flate  // Lossless for rendered text pages
            } else {
                printpdf::ImageCompression::Auto   // JPEG for photos, Flate for sharp
            }),
        }),
    };

    let mut warnings = Vec::new();
    let pdf_bytes = doc.save(&save_opts, &mut warnings);

    if !warnings.is_empty() {
        log::warn!("PDF save warnings: {} items", warnings.len());
    }

    std::fs::write(output_path, &pdf_bytes)
        .map_err(|e| format!("写入文件失败: {}", e))?;

    if let Some(ref cb) = on_progress {
        cb("save", 1, 1);
    }

    Ok(font_warning)
}

// =====================================================
// PDF Passthrough — Form XObject based vector-preserving pipeline
// =====================================================

/// Extract the effective visible box from a PDF page, respecting CropBox over MediaBox.
/// Returns ((x1, y1, x2, y2), (width_pt, height_pt)).
/// CropBox takes precedence over MediaBox (PDF spec 7.7.3.3).
/// Walks up the page tree to inherit from parent nodes if not on the page itself.
fn get_page_effective_box(source: &lopdf::Document, page_id: lopdf::ObjectId) -> Result<((f32, f32, f32, f32), (f32, f32)), String> {
    let mut current_id = page_id;
    let mut visited = std::collections::HashSet::new();

    loop {
        if !visited.insert(current_id) {
            return Err("页面box查找遇到循环引用".to_string());
        }

        let dict = match source.get_object(current_id) {
            Ok(lopdf::Object::Dictionary(d)) => d,
            Ok(lopdf::Object::Reference(id)) => {
                match source.get_object(*id) {
                    Ok(lopdf::Object::Dictionary(d)) => d,
                    _ => return Err("页面对象不是字典".to_string()),
                }
            }
            _ => return Err("页面对象不是字典".to_string()),
        };

        // CropBox takes precedence over MediaBox (PDF spec)
        let cropbox = dict.get(b"CropBox")
            .or_else(|_| dict.get(b"cropbox"))
            .ok();
        if let Some(cb) = cropbox {
            if let Ok(box_val) = parse_box_array(cb, source) {
                return Ok(box_val);
            }
        }

        let mediabox = dict.get(b"MediaBox")
            .or_else(|_| dict.get(b"mediabox"))
            .ok();
        if let Some(mb) = mediabox {
            if let Ok(box_val) = parse_box_array(mb, source) {
                return Ok(box_val);
            }
        }

        // Not found — walk up to parent
        match dict.get(b"Parent").and_then(|v| v.as_reference()) {
            Ok(parent_id) => current_id = parent_id,
            Err(_) => return Err("页面及父节点均缺少MediaBox/CropBox".to_string()),
        }
    }
}

/// Parse a box array (MediaBox/CropBox) into ((x1, y1, x2, y2), (width, height)).
fn parse_box_array(box_obj: &lopdf::Object, source: &lopdf::Document) -> Result<((f32, f32, f32, f32), (f32, f32)), String> {
    match box_obj {
        lopdf::Object::Array(arr) => {
            if arr.len() >= 4 {
                let x1 = match &arr[0] {
                    lopdf::Object::Integer(i) => *i as f32,
                    lopdf::Object::Real(r) => *r as f32,
                    _ => return Err("box x1不是数字".to_string()),
                };
                let y1 = match &arr[1] {
                    lopdf::Object::Integer(i) => *i as f32,
                    lopdf::Object::Real(r) => *r as f32,
                    _ => return Err("box y1不是数字".to_string()),
                };
                let x2 = match &arr[2] {
                    lopdf::Object::Integer(i) => *i as f32,
                    lopdf::Object::Real(r) => *r as f32,
                    _ => return Err("box x2不是数字".to_string()),
                };
                let y2 = match &arr[3] {
                    lopdf::Object::Integer(i) => *i as f32,
                    lopdf::Object::Real(r) => *r as f32,
                    _ => return Err("box y2不是数字".to_string()),
                };
                // Normalize: some PDFs have inverted CropBox (e.g. y1 > y2)
                // which would produce negative width/height and flip content.
                let (nx1, nx2) = if x1 <= x2 { (x1, x2) } else { (x2, x1) };
                let (ny1, ny2) = if y1 <= y2 { (y1, y2) } else { (y2, y1) };
                if x1 != nx1 || y1 != ny1 {
                    log::warn!("parse_box_array: inverted box [{:.1} {:.1} {:.1} {:.1}] → normalized [{:.1} {:.1} {:.1} {:.1}]",
                        x1, y1, x2, y2, nx1, ny1, nx2, ny2);
                }
                Ok(((nx1, ny1, nx2, ny2), (nx2 - nx1, ny2 - ny1)))
            } else {
                Err("box数组长度不足".to_string())
            }
        }
        lopdf::Object::Reference(id) => {
            match source.get_object(*id) {
                Ok(obj) => parse_box_array(obj, source),
                Err(e) => Err(format!("box引用解引用失败: {}", e)),
            }
        }
        _ => Err("box不是数组".to_string()),
    }
}

/// Read the /Rotate attribute from a PDF page dictionary.
/// Walks up the page tree (inherits from parent if not set on page itself).
/// Returns the rotation angle in degrees (0, 90, 180, or 270). Defaults to 0.
fn get_page_rotation(source: &lopdf::Document, page_id: lopdf::ObjectId) -> i32 {
    let mut current_id = page_id;
    let mut visited = std::collections::HashSet::new();

    loop {
        if !visited.insert(current_id) {
            return 0; // cycle protection
        }

        let dict = match source.get_object(current_id) {
            Ok(lopdf::Object::Dictionary(d)) => d.clone(),
            Ok(lopdf::Object::Reference(id)) => {
                match source.get_object(*id) {
                    Ok(lopdf::Object::Dictionary(d)) => d.clone(),
                    _ => return 0,
                }
            }
            _ => return 0,
        };

        if let Ok(rotate_val) = dict.get(b"Rotate") {
            let r = match rotate_val {
                lopdf::Object::Integer(i) => *i as i32,
                lopdf::Object::Real(r) => *r as i32,
                _ => 0,
            };
            return ((r % 360) + 360) % 360;
        }

        // Walk up to parent
        match dict.get(b"Parent").and_then(|v| v.as_reference()) {
            Ok(parent_id) => current_id = parent_id,
            Err(_) => return 0, // no Rotate attribute found → default 0
        }
    }
}

/// Extract the MediaBox from a PDF page, returning (width_pt, height_pt).
/// Walks up the page tree to inherit MediaBox from parent nodes if not on the page itself.
#[allow(dead_code)]
fn get_page_mediabox(source: &lopdf::Document, page_id: lopdf::ObjectId) -> Result<(f32, f32), String> {
    let ((_x1, _y1, _x2, _y2), (w, h)) = get_page_effective_box(source, page_id)?;
    Ok((w, h))
}

/// Recursively copy an object from source doc to dest doc, remapping ObjectId references.
fn deep_copy_object(
    source: &lopdf::Document,
    source_id: lopdf::ObjectId,
    dest: &mut lopdf::Document,
    id_map: &mut std::collections::HashMap<lopdf::ObjectId, lopdf::ObjectId>,
) -> lopdf::ObjectId {
    if let Some(&existing) = id_map.get(&source_id) {
        return existing;
    }

    let dest_id = dest.new_object_id();
    id_map.insert(source_id, dest_id);

    let obj = source.objects.get(&source_id).cloned().unwrap_or(lopdf::Object::Null);
    let remapped = remap_references(obj, source, dest, id_map);
    dest.set_object(dest_id, remapped);

    dest_id
}

/// Recursively remap all ObjectId references in a PDF object tree.
fn remap_references(
    obj: lopdf::Object,
    source: &lopdf::Document,
    dest: &mut lopdf::Document,
    id_map: &mut std::collections::HashMap<lopdf::ObjectId, lopdf::ObjectId>,
) -> lopdf::Object {
    use lopdf::Object;
    match obj {
        Object::Reference(id) => {
            let new_id = deep_copy_object(source, id, dest, id_map);
            Object::Reference(new_id)
        }
        Object::Array(arr) => {
            Object::Array(arr.into_iter()
                .map(|o| remap_references(o, source, dest, id_map))
                .collect())
        }
        Object::Dictionary(dict) => {
            Object::Dictionary(dict.into_iter()
                .map(|(k, v)| (k, remap_references(v, source, dest, id_map)))
                .collect())
        }
        Object::Stream(stream) => {
            let dict: lopdf::Dictionary = stream.dict.into_iter()
                .map(|(k, v)| (k, remap_references(v, source, dest, id_map)))
                .collect();
            // If the stream dict already has a Filter entry, the content is already
            // compressed. Set allows_compression = false to prevent lopdf from
            // compressing it AGAIN during save (which would cause double compression
            // and corrupt the stream data, leading to blank pages).
            let already_compressed = dict.get(b"Filter").is_ok();
            Object::Stream(lopdf::Stream::new(dict, stream.content)
                .with_compression(!already_compressed && stream.allows_compression))
        }
        other => other,
    }
}

/// Merge a source resource dictionary into a merged dictionary.
/// Handles both inline Dictionary and Reference entries by dereferencing them.
/// Child entries override parent entries with the same key (correct PDF inheritance semantics).
fn merge_resource_dict(
    merged: &mut lopdf::Dictionary,
    source_dict: &lopdf::Dictionary,
    doc: &lopdf::Document,
) {
    for (key, value) in source_dict.iter() {
        // Dereference if it's a Reference to get the actual dictionary
        let dict_value = match value {
            lopdf::Object::Reference(id) => {
                match doc.get_object(*id) {
                    Ok(obj) => obj.clone(),
                    Err(_) => value.clone(),
                }
            }
            _ => value.clone(),
        };

        // For sub-dictionaries (Font, XObject, ColorSpace, etc.), merge entries
        match dict_value {
            lopdf::Object::Dictionary(sub_dict) => {
                // Check if merged already has this category (lopdf dict.get returns Result)
                let existing_opt = merged.get(key).ok().cloned();
                match existing_opt {
                    Some(existing) => {
                        let existing_dict = match existing {
                            lopdf::Object::Dictionary(d) => d,
                            lopdf::Object::Reference(id) => {
                                match doc.get_object(id) {
                                    Ok(lopdf::Object::Dictionary(d)) => d.clone(),
                                    _ => {
                                        // Can't merge, just override
                                        merged.set(key.clone(), lopdf::Object::Dictionary(sub_dict));
                                        continue;
                                    }
                                }
                            }
                            _ => {
                                merged.set(key.clone(), lopdf::Object::Dictionary(sub_dict));
                                continue;
                            }
                        };

                        // Merge sub-dictionary entries (child overrides parent)
                        let mut combined = existing_dict;
                        for (sub_key, sub_value) in sub_dict.iter() {
                            combined.set(sub_key.clone(), sub_value.clone());
                        }
                        merged.set(key.clone(), lopdf::Object::Dictionary(combined));
                    }
                    None => {
                        merged.set(key.clone(), lopdf::Object::Dictionary(sub_dict));
                    }
                }
            }
            other => {
                // Non-dictionary entries (ProcSet, etc.) — just override
                merged.set(key.clone(), other);
            }
        }
    }
}

/// Extract a source PDF page as a Form XObject and register it in the output document.
/// Returns (form_xobj_id, page_width_pt, page_height_pt).
fn extract_page_as_form_xobject(
    source: &lopdf::Document,
    page_id: lopdf::ObjectId,
    mut output_doc: &mut lopdf::Document,
    id_map: &mut std::collections::HashMap<lopdf::ObjectId, lopdf::ObjectId>,
) -> Result<(lopdf::ObjectId, f32, f32), String> {
    // 1. Get page content stream bytes (decompressed and concatenated)
    let content_bytes = source.get_page_content(page_id)
        .map_err(|e| format!("提取内容流失败: {}", e))?;

    // 2. Get effective visible box — CropBox takes precedence over MediaBox.
    let ((box_x1, box_y1, _box_x2, _box_y2), (page_w_pt, page_h_pt)) =
        get_page_effective_box(source, page_id)?;

    // 3. Read /Rotate attribute from page dictionary (walks up page tree).
    // PDF viewers auto-rotate the content based on this value. We need to bake
    // the rotation into the Form XObject content stream so it renders correctly
    // when the /Rotate key is stripped.
    let page_rotation = get_page_rotation(source, page_id);
    let rot = ((page_rotation % 360) + 360) % 360;

    // For 90°/270° rotation, the effective visual dimensions swap.
    let (effective_w, effective_h) = if rot == 90 || rot == 270 {
        (page_h_pt, page_w_pt)
    } else {
        (page_w_pt, page_h_pt)
    };

    // 4. Build content stream with rotation + cropbox transforms prepended.
    // We use "q ... cm ...content... Q" to apply transforms.
    // Order: first translate for CropBox, then apply rotation.
    let mut prefix = Vec::new();
    prefix.extend_from_slice(b"q\n");

    // Apply rotation transform (before CropBox shift, so rotation is in page coords)
    match rot {
        90 => {
            // Rotate 90° CW: (x,y) → (y, -x) then translate to fit
            prefix.extend_from_slice(
                format!("0 1 -1 0 {:.4} 0 cm\n", page_w_pt).as_bytes()
            );
        }
        180 => {
            // Rotate 180°: (x,y) → (-x, -y) then translate to fit
            prefix.extend_from_slice(
                format!("-1 0 0 -1 {:.4} {:.4} cm\n", page_w_pt, page_h_pt).as_bytes()
            );
        }
        270 => {
            // Rotate 270° CW (= 90° CCW): (x,y) → (-y, x) then translate to fit
            prefix.extend_from_slice(
                format!("0 -1 1 0 0 {:.4} cm\n", page_h_pt).as_bytes()
            );
        }
        _ => {} // 0°: no rotation needed
    }

    // Apply CropBox offset if non-zero origin
    if box_x1.abs() > 0.01 || box_y1.abs() > 0.01 {
        prefix.extend_from_slice(
            format!("1 0 0 1 {:.4} {:.4} cm\n", -box_x1, -box_y1).as_bytes()
        );
    }

    // Close the graphics state after content
    let mut suffix = Vec::new();
    suffix.extend_from_slice(b"\nQ\n");

    // Combine: prefix + content (annotations and suffix appended after /Annots processing)
    let mut final_content = prefix;
    final_content.extend_from_slice(&content_bytes);

    if rot != 0 {
        log::info!("extract_page_as_form_xobject: page rotation={}°, page {:.1}x{:.1}pt → effective {:.1}x{:.1}pt",
            rot, page_w_pt, page_h_pt, effective_w, effective_h);
    }

    // 5. Get page resources — merge ALL resource dictionaries including inherited ones.
    let (resources_opt, ref_ids) = source.get_page_resources(page_id)
        .map_err(|e| format!("提取资源失败: {}", e))?;

    // 6. Merge all resource dictionaries: page's own + all inherited from parents.
    let mut merged = lopdf::Dictionary::new();
    for rid in &ref_ids {
        if let Ok(res_dict) = source.get_dictionary(*rid) {
            merge_resource_dict(&mut merged, res_dict, source);
        }
    }
    if let Some(dict) = resources_opt {
        merge_resource_dict(&mut merged, dict, source);
    }

    let mut remapped_resources = {
        let obj = lopdf::Object::Dictionary(merged);
        remap_references(obj, source, output_doc, id_map)
    };

    // 6.5 Process page annotations (stamps, signatures, etc.)
    // Annotations are NOT part of the page content stream — they are separate objects
    // in the page's /Annots array. PDF viewers render them on top of the page content,
    // but lopdf's get_page_content() only returns the content stream, so we must
    // explicitly extract annotation appearances and append them to the Form XObject.
    let mut annot_draw_cmds = Vec::new();
    let mut annot_xobjects: Vec<(Vec<u8>, lopdf::ObjectId)> = Vec::new();

    if let Ok(page_dict) = source.get_dictionary(page_id) {
        if let Ok(annots_obj) = page_dict.get(b"Annots") {
            let annot_refs: Vec<lopdf::ObjectId> = match annots_obj {
                lopdf::Object::Array(arr) => {
                    arr.iter().filter_map(|o| {
                        if let lopdf::Object::Reference(id) = o { Some(*id) } else { None }
                    }).collect()
                }
                lopdf::Object::Reference(id) => vec![*id],
                _ => vec![],
            };

            for (annot_idx, annot_id) in annot_refs.iter().enumerate() {
                let annot_dict = match source.get_dictionary(*annot_id) {
                    Ok(d) => d,
                    Err(_) => continue,
                };

                // Skip hidden annotations (F bit 2 = Hidden)
                if let Some(lopdf::Object::Integer(f)) = annot_dict.get(b"F").ok() {
                    if *f & 2 != 0 { continue; }
                }

                // Get /AP → /N (normal appearance)
                let normal_ap_obj = match annot_dict.get(b"AP") {
                    Ok(lopdf::Object::Dictionary(ap_dict)) => {
                        match ap_dict.get(b"N") {
                            Ok(obj) => obj.clone(),
                            Err(_) => continue,
                        }
                    }
                    Ok(lopdf::Object::Reference(id)) => {
                        match source.get_dictionary(*id) {
                            Ok(ap_dict) => {
                                match ap_dict.get(b"N") {
                                    Ok(obj) => obj.clone(),
                                    Err(_) => continue,
                                }
                            }
                            Err(_) => continue,
                        }
                    }
                    _ => continue,
                };

                // Get annotation Rect [x1 y1 x2 y2]
                let rect: Vec<f32> = match annot_dict.get(b"Rect") {
                    Ok(lopdf::Object::Array(arr)) => {
                        arr.iter().filter_map(|o| match o {
                            lopdf::Object::Real(f) => Some(*f),
                            lopdf::Object::Integer(i) => Some(*i as f32),
                            lopdf::Object::Reference(id) => {
                                source.get_object(*id).ok().and_then(|obj| match obj {
                                    lopdf::Object::Real(f) => Some(*f),
                                    lopdf::Object::Integer(i) => Some(*i as f32),
                                    _ => None,
                                })
                            }
                            _ => None,
                        }).collect()
                    }
                    _ => continue,
                };
                if rect.len() != 4 { continue; }

                // Deep copy the appearance XObject to output document
                let ap_xobj_id = match normal_ap_obj {
                    lopdf::Object::Reference(id) => {
                        deep_copy_object(source, id, &mut output_doc, id_map)
                    }
                    lopdf::Object::Stream(_) => {
                        let remapped = remap_references(normal_ap_obj, source, &mut output_doc, id_map);
                        output_doc.add_object(remapped)
                    }
                    _ => continue,
                };

                // PDF spec requires annotation appearances to be rendered as isolated
                // transparency groups. When baked into the content stream (instead of
                // rendered by the viewer's annotation engine), we must explicitly add
                // /Group<</S/Transparency/I true>> so blend modes (e.g. /BM/Darken)
                // and SMask work correctly across all PDF readers.
                if let Ok(lopdf::Object::Stream(ref mut s)) = output_doc.get_object_mut(ap_xobj_id) {
                    if s.dict.get(b"Group").is_err() {
                        let mut group = lopdf::Dictionary::new();
                        group.set("S", lopdf::Object::Name(b"Transparency".to_vec()));
                        group.set("I", lopdf::Object::Boolean(true));
                        s.dict.set("Group", lopdf::Object::Dictionary(group));
                    }
                }

                // Get the appearance BBox and Matrix from the deep-copied object
                let (bbox, ap_matrix) = match output_doc.get_object(ap_xobj_id) {
                    Ok(lopdf::Object::Stream(s)) => {
                        let bb = match s.dict.get(b"BBox") {
                            Ok(lopdf::Object::Array(arr)) => {
                                arr.iter().filter_map(|o| match o {
                                    lopdf::Object::Real(f) => Some(*f),
                                    lopdf::Object::Integer(i) => Some(*i as f32),
                                    _ => None,
                                }).collect()
                            }
                            _ => vec![],
                        };
                        // Appearance Matrix [a b c d e f] — identity if absent
                        let mat: Vec<f32> = match s.dict.get(b"Matrix") {
                            Ok(lopdf::Object::Array(arr)) => {
                                arr.iter().filter_map(|o| match o {
                                    lopdf::Object::Real(f) => Some(*f),
                                    lopdf::Object::Integer(i) => Some(*i as f32),
                                    _ => None,
                                }).collect()
                            }
                            _ => vec![],
                        };
                        (bb, mat)
                    }
                    _ => (vec![], vec![]),
                };
                // Default BBox from Rect dimensions if not found
                let bbox = if bbox.len() == 4 { bbox } else {
                    vec![0.0, 0.0, rect[2] - rect[0], rect[3] - rect[1]]
                };

                let (rx1, ry1, rx2, ry2) = (rect[0], rect[1], rect[2], rect[3]);
                let (bx1, by1, bx2, by2) = (bbox[0], bbox[1], bbox[2], bbox[3]);
                let bw = bx2 - bx1;
                let bh = by2 - by1;
                if bw.abs() < 0.01 || bh.abs() < 0.01 { continue; }

                // Build transform: Rect_mapping × Appearance_Matrix
                // Rect_mapping maps BBox → Rect: [sx 0 0 sy tx ty]
                // If appearance has /Matrix [a b c d e f], compose: Rect_mapping × Matrix
                let sx = (rx2 - rx1) / bw;
                let sy = (ry2 - ry1) / bh;
                let tx = rx1 - sx * bx1;
                let ty = ry1 - sy * by1;

                let (ma, mb, mc, md, me, mf) = if ap_matrix.len() == 6 {
                    (ap_matrix[0], ap_matrix[1], ap_matrix[2],
                     ap_matrix[3], ap_matrix[4], ap_matrix[5])
                } else {
                    (1.0, 0.0, 0.0, 1.0, 0.0, 0.0) // identity
                };

                // Compose: [sx 0 0 sy tx ty] × [ma mb mc md me mf]
                // = [sx*ma  sx*mb  sy*mc  sy*md  sx*me+tx  sy*mf+ty]
                let cm_a = sx * ma;
                let cm_b = sx * mb;
                let cm_c = sy * mc;
                let cm_d = sy * md;
                let cm_e = sx * me + tx;
                let cm_f = sy * mf + ty;

                // Use a unique prefix to avoid name collisions with existing XObjects
                let annot_name = format!("__Annot{}", annot_idx);
                annot_xobjects.push((annot_name.clone().into_bytes(), ap_xobj_id));

                // Drawing command: q <composed_matrix> /AnnotN Do Q
                annot_draw_cmds.extend_from_slice(
                    format!("q {:.6} {:.6} {:.6} {:.6} {:.6} {:.6} cm /{} Do Q\n",
                        cm_a, cm_b, cm_c, cm_d, cm_e, cm_f, annot_name).as_bytes()
                );

                log::info!("extract_page_as_form_xobject: annotation[{}] rect=[{:.1},{:.1},{:.1},{:.1}] bbox=[{:.1},{:.1},{:.1},{:.1}]",
                    annot_idx, rx1, ry1, rx2, ry2, bx1, by1, bx2, by2);
            }

            if !annot_xobjects.is_empty() {
                log::info!("extract_page_as_form_xobject: processed {} annotation(s)", annot_xobjects.len());
            }
        }
    }

    // Append closing suffix FIRST, then annotation drawing commands.
    // The suffix (\nQ\n) restores the graphics state, undoing any CTM
    // transformations from the page content (e.g. "2.8346 0 0 2.8346 0 0 cm").
    // Annotation Rect coordinates are in the BBox coordinate system, so they
    // must be drawn AFTER the graphics state is restored — otherwise the CTM
    // scale would push the annotations far outside the BBox bounds.
    final_content.extend_from_slice(&suffix);
    final_content.extend_from_slice(&annot_draw_cmds);

    // Add annotation XObjects to the resources dictionary
    if !annot_xobjects.is_empty() {
        if let lopdf::Object::Dictionary(ref mut res_dict) = remapped_resources {
            let xobject_dict = match res_dict.get(b"XObject") {
                Ok(lopdf::Object::Dictionary(d)) => d.clone(),
                Ok(lopdf::Object::Reference(id)) => {
                    match output_doc.get_object(*id) {
                        Ok(lopdf::Object::Dictionary(d)) => d.clone(),
                        _ => lopdf::Dictionary::new(),
                    }
                }
                _ => lopdf::Dictionary::new(),
            };
            let mut merged_xobject = xobject_dict;
            for (name, id) in annot_xobjects {
                merged_xobject.set(name, lopdf::Object::Reference(id));
            }
            res_dict.set(b"XObject".to_vec(), lopdf::Object::Dictionary(merged_xobject));
        }
    }

    // 7. Build Form XObject stream — BBox uses EFFECTIVE (post-rotation) dimensions.
    let mut dict = lopdf::Dictionary::new();
    dict.set("Type", lopdf::Object::Name(b"XObject".to_vec()));
    dict.set("Subtype", lopdf::Object::Name(b"Form".to_vec()));
    dict.set("FormType", lopdf::Object::Integer(1));
    dict.set("BBox", lopdf::Object::Array(vec![
        lopdf::Object::Real(0.0),
        lopdf::Object::Real(0.0),
        lopdf::Object::Real(effective_w),
        lopdf::Object::Real(effective_h),
    ]));
    dict.set("Resources", remapped_resources);

    // Transparency group — ensures correct rendering of overlapping content
    let mut group_dict = lopdf::Dictionary::new();
    group_dict.set("Type", lopdf::Object::Name(b"Group".to_vec()));
    group_dict.set("S", lopdf::Object::Name(b"Transparency".to_vec()));
    dict.set("Group", lopdf::Object::Dictionary(group_dict));

    let stream = lopdf::Stream::new(dict, final_content).with_compression(true);
    let xobj_id = output_doc.add_object(lopdf::Object::Stream(stream));

    Ok((xobj_id, effective_w, effective_h))
}

/// Per-slot adjustment data for passthrough rendering.
struct SlotAdjustment {
    rotation: i32,
    scale: f32,
    offset_x: f32,
    offset_y: f32,
    /// If true, this XObject is a raw Image (unit square 0→1),
    /// not a Form XObject (BBox coordinate space).
    /// The cm matrix must account for the different coordinate space.
    is_image: bool,
}

/// Build the content stream for one output page using cm + Do operators.
/// Each Form XObject is positioned, scaled, and rotated within its layout slot.
fn build_nup_content_stream(
    form_xobjs: &[(usize, lopdf::ObjectId, f32, f32)],  // (layout_slot_idx, xobj_id, src_w_pt, src_h_pt)
    slot_positions: &[LayoutSlotMm],
    settings: &RenderSettings,
    slot_adjustments: &[SlotAdjustment],  // per-slot rotation/scale/offset
) -> Result<Vec<u8>, String> {
    use lopdf::content::Operation;

    let mut ops = Vec::new();

    for (adj_idx, (layout_slot_idx, _xobj_id, src_w_pt, src_h_pt)) in form_xobjs.iter().enumerate() {
        let slot = &slot_positions[*layout_slot_idx];
        let slot_w_pt = slot.w_mm * MM_TO_PT;
        let slot_h_pt = slot.h_mm * MM_TO_PT;

        log::info!("build_nup: layout_slot[{}] adj[{}] x={:.2}mm y={:.2}mm w={:.2}mm h={:.2}mm",
            layout_slot_idx, adj_idx, slot.x_mm, slot.y_mm, slot.w_mm, slot.h_mm);

        // Handle rotation via transformation matrix
        let adj = if adj_idx < slot_adjustments.len() {
            &slot_adjustments[adj_idx]
        } else {
            &SlotAdjustment { rotation: 0, scale: 1.0, offset_x: 0.0, offset_y: 0.0, is_image: false }
        };
        let rotation = adj.rotation;
        let rot = ((rotation % 360) + 360) % 360;

        // For 90°/270° rotation, the visual dimensions swap (width↔height),
        // so scaling must be computed against the *rotated* dimensions to fit the slot correctly.
        let (vis_w, vis_h) = if rot == 90 || rot == 270 {
            (*src_h_pt, *src_w_pt) // rotated: visual width = original height, etc.
        } else {
            (*src_w_pt, *src_h_pt)
        };

        // Compute scale to fit in slot based on visual (rotated) dimensions
        let (mut scale_x, mut scale_y) = match settings.fit_mode.as_str() {
            "fill" => (slot_w_pt / vis_w, slot_h_pt / vis_h),
            "original" => (1.0, 1.0),
            "custom" => {
                let contain_s = (slot_w_pt / vis_w).min(slot_h_pt / vis_h);
                let s = contain_s * settings.custom_scale;
                (s, s)
            }
            _ => {
                // "contain" (default)
                let s = (slot_w_pt / vis_w).min(slot_h_pt / vis_h);
                (s, s)
            }
        };

        // Per-slot scale override
        if adj.scale != 1.0 {
            scale_x *= adj.scale;
            scale_y *= adj.scale;
        }

        // Centered position in slot (bottom-left origin) based on visual dimensions.
        // 报销单模式：左上对齐（贴段内区域左上角）；常规模式：居中。
        let draw_w = vis_w * scale_x;
        let draw_h = vis_h * scale_y;
        let mut offset_x = slot.x_mm * MM_TO_PT;
        let mut offset_y = if settings.reimburse_mode {
            // bottom-up 坐标：顶部对齐 = slot 顶边 - 图像高度
            slot.y_mm * MM_TO_PT + (slot_h_pt - draw_h)
        } else {
            slot.y_mm * MM_TO_PT + (slot_h_pt - draw_h) / 2.0
        };
        if !settings.reimburse_mode {
            offset_x += (slot_w_pt - draw_w) / 2.0;
        }

        // Per-slot offset override (convert mm to pt)
        if adj.offset_x != 0.0 { offset_x += adj.offset_x * MM_TO_PT; }
        if adj.offset_y != 0.0 { offset_y -= adj.offset_y * MM_TO_PT; }  // JS Y+ is down, PDF Y+ is up

        // PDF transformation matrix: [a b c d e f]
        //
        // For Form XObjects (PDF pages): coordinate space is (0,0)-(src_w_pt,src_h_pt).
        //   sx = draw_w / src_w_pt, sy = draw_h / src_h_pt
        //
        // For Image XObjects (OFD/images): coordinate space is (0,0)-(1,1).
        //   sx = draw_w, sy = draw_h
        //
        // Rotation matrices derived from the desired mapping:
        //   rot=0:   (x,y) → (sx*x+ox, sy*y+oy)
        //   rot=90:  (x,y) → (sx*(src_h-y)+ox, sy*x+oy)
        //   rot=180: (x,y) → (sx*(src_w-x)+ox, sy*(src_h-y)+oy)
        //   rot=270: (x,y) → (sx*y+ox, sy*(src_w-x)+oy)
        let (sx, sy) = if adj.is_image {
            // Image XObject: unit square → direct pixel dimensions
            (draw_w, draw_h)
        } else if rot == 90 || rot == 270 {
            // Form XObject with rotation: visual width = src_h, visual height = src_w
            (draw_w / *src_h_pt, draw_h / *src_w_pt)
        } else {
            // Form XObject no rotation
            (draw_w / *src_w_pt, draw_h / *src_h_pt)
        };

        let matrix: Vec<lopdf::Object> = match rot {
            0 => {
                // [sx 0 0 sy offset_x offset_y]
                vec![
                    lopdf::Object::Real(sx), lopdf::Object::Real(0.0),
                    lopdf::Object::Real(0.0), lopdf::Object::Real(sy),
                    lopdf::Object::Real(offset_x), lopdf::Object::Real(offset_y),
                ]
            }
            90 => {
                // [0 sy -sx 0 offset_x+draw_w offset_y]
                vec![
                    lopdf::Object::Real(0.0), lopdf::Object::Real(sy),
                    lopdf::Object::Real(-sx), lopdf::Object::Real(0.0),
                    lopdf::Object::Real(offset_x + draw_w), lopdf::Object::Real(offset_y),
                ]
            }
            180 => {
                // [-sx 0 0 -sy offset_x+draw_w offset_y+draw_h]
                vec![
                    lopdf::Object::Real(-sx), lopdf::Object::Real(0.0),
                    lopdf::Object::Real(0.0), lopdf::Object::Real(-sy),
                    lopdf::Object::Real(offset_x + draw_w), lopdf::Object::Real(offset_y + draw_h),
                ]
            }
            270 => {
                // [0 -sy sx 0 offset_x offset_y+draw_h]
                vec![
                    lopdf::Object::Real(0.0), lopdf::Object::Real(-sy),
                    lopdf::Object::Real(sx), lopdf::Object::Real(0.0),
                    lopdf::Object::Real(offset_x), lopdf::Object::Real(offset_y + draw_h),
                ]
            }
            _ => {
                vec![
                    lopdf::Object::Real(sx), lopdf::Object::Real(0.0),
                    lopdf::Object::Real(0.0), lopdf::Object::Real(sy),
                    lopdf::Object::Real(offset_x), lopdf::Object::Real(offset_y),
                ]
            }
        };

        // Build the XObject name for this Form XObject
        let xobj_name = lopdf::Object::Name(format!("Fm{}", layout_slot_idx).into_bytes());

        // Clip to slot boundary — prevents per-slot overflow into adjacent slots
        ops.push(Operation { operator: "q".into(), operands: vec![] });
        ops.push(Operation { operator: "re".into(), operands: vec![
            lopdf::Object::Real(slot.x_mm * MM_TO_PT),
            lopdf::Object::Real(slot.y_mm * MM_TO_PT),
            lopdf::Object::Real(slot_w_pt),
            lopdf::Object::Real(slot_h_pt),
        ] });
        ops.push(Operation { operator: "W".into(), operands: vec![] });
        ops.push(Operation { operator: "n".into(), operands: vec![] });
        ops.push(Operation { operator: "cm".into(), operands: matrix });
        ops.push(Operation { operator: "Do".into(), operands: vec![xobj_name] });
        ops.push(Operation { operator: "Q".into(), operands: vec![] });
    }

    let content = lopdf::content::Content { operations: ops };
    content.encode().map_err(|e| format!("内容流编码失败: {}", e))
}

/// Convert an `ImageSource` to a lopdf Image XObject in the output document.
/// Returns `(xobj_id, width_pt, height_pt)`.
///
/// Rotation is ALWAYS baked into pixels here (including 180° for JPEG passthrough).
/// The caller must set SlotAdjustment rotation=0 for the resulting XObject, because
/// `build_nup_content_stream` applies rotation via the PDF cm matrix — if we also
/// baked rotation into pixels, it would be double-rotated.
fn image_to_lopdf_xobject(
    source: &ImageSource,
    rotation: i32,
    output_doc: &mut lopdf::Document,
) -> Result<(lopdf::ObjectId, f32, f32), String> {
    let rot = ((rotation % 360) + 360) % 360;

    match source {
        ImageSource::JpegPassthrough { raw_bytes, width, height, num_components } => {
            if rot == 0 {
                // No rotation: embed raw JPEG bytes directly (zero re-encoding)
                let nc = *num_components;
                let w_pt = *width as f32 * 72.0 / RENDER_DPI as f32;
                let h_pt = *height as f32 * 72.0 / RENDER_DPI as f32;
                let xobj_id = build_lopdf_jpeg_xobject(output_doc, raw_bytes, *width, *height, nc);
                Ok((xobj_id, w_pt, h_pt))
            } else {
                // Any rotation (90°/180°/270°): decode → rotate → re-encode
                let img = image::load_from_memory(raw_bytes)
                    .map_err(|e| format!("JPEG解码失败: {}", e))?;
                let rotated = match rot {
                    90  => img.rotate90(),
                    180 => img.rotate180(),
                    270 => img.rotate270(),
                    _   => img,
                };
                let (w, h) = (rotated.width(), rotated.height());
                let jpeg_bytes = encode_image_to_jpeg_bytes(&rotated)?;
                let w_pt = w as f32 * 72.0 / RENDER_DPI as f32;
                let h_pt = h as f32 * 72.0 / RENDER_DPI as f32;
                let xobj_id = build_lopdf_jpeg_xobject(output_doc, &jpeg_bytes, w, h, 3);
                Ok((xobj_id, w_pt, h_pt))
            }
        }
        ImageSource::Decoded(img) => {
            let rotated = match rot {
                90  => img.rotate90(),
                180 => img.rotate180(),
                270 => img.rotate270(),
                _   => img.clone(),
            };
            let (w, h) = (rotated.width(), rotated.height());
            let jpeg_bytes = encode_image_to_jpeg_bytes(&rotated)?;
            let w_pt = w as f32 * 72.0 / RENDER_DPI as f32;
            let h_pt = h as f32 * 72.0 / RENDER_DPI as f32;
            let xobj_id = build_lopdf_jpeg_xobject(output_doc, &jpeg_bytes, w, h, 3);
            Ok((xobj_id, w_pt, h_pt))
        }
    }
}

/// Encode a DynamicImage to JPEG bytes with high quality (90) for print output.
/// Uses JpegEncoder directly for explicit quality control (default write_to uses 75).
/// For images with alpha (RGBA/La): composites onto white background then strips alpha,
/// since JPEG doesn't support transparency. This prevents transparent areas from becoming black.
#[allow(dead_code)]
fn encode_image_to_jpeg_bytes(img: &image::DynamicImage) -> Result<Vec<u8>, String> {
    use image::codecs::jpeg::JpegEncoder;

    // JPEG doesn't support alpha channels.
    // For RGBA images: composite onto white background (transparent → white, not black).
    let rgb_buf;
    let target: &image::DynamicImage = match img {
        image::DynamicImage::ImageRgba8(buf) => {
            let w = buf.width();
            let h = buf.height();
            let mut rgb = image::RgbImage::new(w, h);
            for y in 0..h {
                for x in 0..w {
                    let px = buf.get_pixel(x, y);
                    let a = px[3] as u32;
                    if a == 255 {
                        rgb.put_pixel(x, y, image::Rgb([px[0], px[1], px[2]]));
                    } else if a == 0 {
                        rgb.put_pixel(x, y, image::Rgb([255, 255, 255])); // white background
                    } else {
                        // Alpha blend with white background
                        let r = ((px[0] as u32 * a + 255 * (255 - a)) / 255) as u8;
                        let g = ((px[1] as u32 * a + 255 * (255 - a)) / 255) as u8;
                        let b = ((px[2] as u32 * a + 255 * (255 - a)) / 255) as u8;
                        rgb.put_pixel(x, y, image::Rgb([r, g, b]));
                    }
                }
            }
            rgb_buf = image::DynamicImage::from(rgb);
            &rgb_buf
        }
        image::DynamicImage::ImageLumaA8(buf) => {
            let w = buf.width();
            let h = buf.height();
            let mut luma = image::GrayImage::new(w, h);
            for y in 0..h {
                for x in 0..w {
                    let px = buf.get_pixel(x, y);
                    let a = px[1] as u32;
                    let v = if a == 0 { 255 }
                            else { ((px[0] as u32 * a + 255 * (255 - a)) / 255) as u8 };
                    luma.put_pixel(x, y, image::Luma([v]));
                }
            }
            rgb_buf = image::DynamicImage::from(luma);
            &rgb_buf
        }
        _ => img,
    };

    let mut buf = std::io::Cursor::new(Vec::new());
    let encoder = JpegEncoder::new_with_quality(&mut buf, 90);
    target.write_with_encoder(encoder)
        .map_err(|e| format!("JPEG编码失败: {}", e))?;
    Ok(buf.into_inner())
}

/// Encode a DynamicImage to raw RGB pixel bytes (for PNG encoding via lopdf FlateDecode).
/// Strips alpha channel if present (composites onto white background).
/// Returns (raw_pixels, num_components).
#[allow(dead_code)]
fn image_to_raw_rgb(img: &image::DynamicImage) -> (Vec<u8>, u8) {
    match img {
        image::DynamicImage::ImageRgba8(buf) => {
            let w = buf.width() as usize;
            let h = buf.height() as usize;
            let mut rgb = Vec::with_capacity(w * h * 3);
            for y in 0..h {
                for x in 0..w {
                    let px = buf.get_pixel(x as u32, y as u32);
                    let a = px[3] as u32;
                    if a == 255 {
                        rgb.extend_from_slice(&[px[0], px[1], px[2]]);
                    } else if a == 0 {
                        rgb.extend_from_slice(&[255, 255, 255]);
                    } else {
                        let r = ((px[0] as u32 * a + 255 * (255 - a)) / 255) as u8;
                        let g = ((px[1] as u32 * a + 255 * (255 - a)) / 255) as u8;
                        let b = ((px[2] as u32 * a + 255 * (255 - a)) / 255) as u8;
                        rgb.extend_from_slice(&[r, g, b]);
                    }
                }
            }
            (rgb, 3)
        }
        image::DynamicImage::ImageRgb8(buf) => {
            (buf.as_raw().to_vec(), 3)
        }
        image::DynamicImage::ImageLuma8(buf) => {
            (buf.as_raw().to_vec(), 1)
        }
        other => {
            let rgb = other.to_rgb8();
            (rgb.as_raw().to_vec(), 3)
        }
    }
}

/// Build a lopdf Image XObject from JPEG bytes using DCTDecode.
fn build_lopdf_jpeg_xobject(
    output_doc: &mut lopdf::Document,
    jpeg_bytes: &[u8],
    width: u32,
    height: u32,
    num_components: u8,
) -> lopdf::ObjectId {
    let color_space: &[u8] = match num_components {
        1 => b"DeviceGray",
        4 => b"DeviceCMYK",
        _ => b"DeviceRGB",
    };

    let mut dict = lopdf::Dictionary::new();
    dict.set("Type", lopdf::Object::Name(b"XObject".to_vec()));
    dict.set("Subtype", lopdf::Object::Name(b"Image".to_vec()));
    dict.set("Width", lopdf::Object::Integer(width as i64));
    dict.set("Height", lopdf::Object::Integer(height as i64));
    dict.set("BitsPerComponent", lopdf::Object::Integer(8));
    dict.set("ColorSpace", lopdf::Object::Name(color_space.to_vec()));
    dict.set("Filter", lopdf::Object::Name(b"DCTDecode".to_vec()));

    let stream = lopdf::Stream::new(dict, jpeg_bytes.to_vec()).with_compression(false);
    output_doc.add_object(lopdf::Object::Stream(stream))
}

/// Generate PDF using Form XObject passthrough — preserves vector content,
/// fonts, and text exactly. Supports all layouts (1×1, 2×1, 3×3, etc.)
/// and rotation via PDF transformation matrices.
///
/// **Hybrid mode**: also handles image/OFD slots by encoding them as JPEG
/// Image XObjects in the same lopdf document. PDF pages stay vector-sharp;
/// image/OFD pages are embedded as high-quality JPEG. No printpdf involved.
fn generate_pdf_passthrough(
    request: &LayoutRenderRequest,
    output_path: &std::path::Path,
    on_progress: Option<&ProgressFn>,
    sources: &[Option<ImageSource>],
) -> Result<(), String> {
    log::info!("generate_pdf_passthrough: settings.footer_margin={}", request.settings.footer_margin);
    let (slot_positions, pw, ph) = calculate_layout_mm(&request.settings);
    let pw_pt = pw * MM_TO_PT;
    let ph_pt = ph * MM_TO_PT;

    let mut output_doc = lopdf::Document::with_version("1.4");

    // Cache loaded source PDFs by path
    let mut source_cache: std::collections::HashMap<String, lopdf::Document> = std::collections::HashMap::new();
    // Global ObjectId remapping: source (doc_path, ObjectId) → output ObjectId
    let mut global_id_maps: std::collections::HashMap<String, std::collections::HashMap<lopdf::ObjectId, lopdf::ObjectId>> =
        std::collections::HashMap::new();

    // Create the Pages tree object
    let pages_id = output_doc.new_object_id();
    let mut all_page_ids: Vec<lopdf::ObjectId> = Vec::new();

    // Pre-load CJK font for text overlay (page_num, print_date, footer, number, watermark)
    let needs_text_font = request.settings.page_num
        || request.settings.print_date
        || request.settings.footer_text.as_ref().map_or(false, |t| !t.is_empty())
        || request.settings.number
        || (request.settings.watermark && request.settings.watermark_text.as_ref().map_or(false, |t| !t.is_empty()));
    let text_font = if needs_text_font {
        load_system_font()
    } else {
        None
    };
    if text_font.is_none() && needs_text_font {
        log::warn!("lopdf hybrid: text overlay enabled but font load failed, overlay will be skipped");
    }

    // Diagnostic: log all file specs
    for (i, f) in request.files.iter().enumerate() {
        log::info!("lopdf hybrid: file[{}] pdf_path={:?} source_type={:?} data_url_len={} file_path={:?}",
            i, f.pdf_path, f.source_type, f.data_url.len(), f.file_path);
    }
    for (i, s) in sources.iter().enumerate() {
        match s {
            Some(ImageSource::Decoded(img)) => log::info!("lopdf hybrid: source[{}] Decoded {}x{}", i, img.width(), img.height()),
            Some(ImageSource::JpegPassthrough { width, height, .. }) => log::info!("lopdf hybrid: source[{}] JpegPassthrough {}x{}", i, width, height),
            None => log::info!("lopdf hybrid: source[{}] None", i),
        }
    }

    for (page_idx, page_spec) in request.pages.iter().enumerate() {
        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，PDF生成已中止".to_string());
        }

        log::info!("lopdf hybrid: page {} has {} slots", page_idx, page_spec.slots.len());

        // Collect Form XObjects for each slot in this page
        // (slot_idx, xobj_id, src_w_pt, src_h_pt) — slot_idx preserves correct layout position
        let mut page_form_xobjs: Vec<(usize, lopdf::ObjectId, f32, f32)> = Vec::new();
        let mut slot_adjustments: Vec<SlotAdjustment> = Vec::new();
        let mut xobj_names: Vec<(std::vec::Vec<u8>, lopdf::ObjectId)> = Vec::new();
        let mut filled_slot_indices: Vec<usize> = Vec::new();

        for (slot_idx, slot) in page_spec.slots.iter().enumerate() {
            let file_idx = match slot.file_index {
                Some(idx) if idx < request.files.len() => idx,
                _ => {
                    log::info!("lopdf hybrid: page {} slot {} — no file_index, skip", page_idx, slot_idx);
                    continue;
                }
            };
            let file = &request.files[file_idx];

            log::info!("lopdf hybrid: page {} slot {} file_idx={} pdf_path={:?} source_type={:?}",
                page_idx, slot_idx, file_idx, file.pdf_path, file.source_type);

            // Two paths: PDF page → Form XObject (vector), image/OFD → Image XObject (JPEG)
            let (xobj_id, src_w_pt, src_h_pt) = if let Some(pdf_path) = &file.pdf_path {
                // PDF passthrough path
                let page_idx_in_pdf = match file.pdf_page_idx {
                    Some(idx) => idx,
                    None => continue,
                };

                // Load source PDF (cached)
                if !source_cache.contains_key(pdf_path) {
                    let source = lopdf::Document::load(pdf_path)
                        .map_err(|e| format!("加载源PDF失败 {}: {}", pdf_path, e))?;
                    source_cache.insert(pdf_path.clone(), source);
                    global_id_maps.insert(pdf_path.clone(), std::collections::HashMap::new());
                }
                let source = source_cache.get_mut(pdf_path).unwrap();
                let id_map = global_id_maps.get_mut(pdf_path).unwrap();

                // Find the source page ObjectId (lopdf uses 1-based page numbers)
                let pages = source.get_pages();
                let source_page_id = pages.get(&(page_idx_in_pdf + 1))
                    .copied()
                    .ok_or_else(|| format!("PDF页面{}不存在 (文件: {})", page_idx_in_pdf + 1, pdf_path))?;

                // Extract as Form XObject (vector quality preserved)
                extract_page_as_form_xobject(
                    source, source_page_id, &mut output_doc, id_map
                )?
            } else {
                // Image/OFD path → encode as FlateDecode (lossless) Image XObject
                // FlateDecode matches the original printpdf behavior — no JPEG quality loss.
                // For a print-focused tool, text sharpness matters more than file size.
                let source = match sources.get(file_idx).and_then(|s| s.as_ref()) {
                    Some(s) => s,
                    None => {
                        log::warn!("lopdf hybrid: image source[{}] is None, skipping slot", file_idx);
                        continue;
                    }
                };
                match image_to_lopdf_xobject(source, slot.rotation, &mut output_doc) {
                    Ok(result) => result,
                    Err(e) => {
                        log::warn!("lopdf hybrid: image slot {} encode failed: {}, skipping", file_idx, e);
                        continue;
                    }
                }
            };

            let xobj_name = format!("Fm{}", slot_idx);
            xobj_names.push((xobj_name.into_bytes(), xobj_id));
            page_form_xobjs.push((slot_idx, xobj_id, src_w_pt, src_h_pt));
            filled_slot_indices.push(slot_idx);
            // For image XObjects: rotation=0 because it's already baked into pixels.
            // For PDF Form XObjects: use the original slot rotation.
            let is_pdf = file.pdf_path.is_some();
            slot_adjustments.push(SlotAdjustment {
                rotation: if is_pdf { slot.rotation } else { 0 },
                scale: slot.scale.unwrap_or(1.0),
                offset_x: slot.offset_x.unwrap_or(0.0),
                offset_y: slot.offset_y.unwrap_or(0.0),
                is_image: !is_pdf,
            });
        }

        if page_form_xobjs.is_empty() {
            continue; // Empty page, skip
        }

        // Build content stream for this output page
        let mut content_bytes = build_nup_content_stream(
            &page_form_xobjs, &slot_positions, &request.settings, &slot_adjustments
        )?;

        // Add text overlay (page number + print date) if enabled
        let page_num_text = if request.settings.page_num {
            format!("第 {} 页 / 共 {} 页", page_idx + 1, request.pages.len())
        } else {
            String::new()
        };
        let print_date_text = if request.settings.print_date {
            let now = std::time::SystemTime::now();
            let duration = now.duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
            let secs = duration.as_secs();
            // Simple date calculation from unix timestamp
            let days = secs / 86400;
            // Calculate year/month/day from days since 1970-01-01
            let (year, month, day) = days_to_ymd(days as i64);
            format!("打印日期 {:04}-{:02}-{:02}", year, month, day)
        } else {
            String::new()
        };

        let footer_text = request.settings.footer_text.clone().unwrap_or_default();

        if let Some((png_bytes, _img_w, _img_h)) = render_text_overlay(
            &text_font, &page_num_text, &print_date_text, &footer_text, pw, request.pages.len()
        ) {
            // Embed as Image XObject (RGBA PNG → decode to raw pixels, then encode as FlateDecode)
            match image::load_from_memory(&png_bytes) {
                Ok(rgba_img) => {
                    let rgba_img = rgba_img.to_rgba8();
                    let (w, h) = rgba_img.dimensions();

                    // Extract alpha channel and create SMask for transparency
                    let alpha_bytes: Vec<u8> = rgba_img.pixels().map(|p| p[3]).collect();
                    let rgb_bytes: Vec<u8> = rgba_img.pixels().flat_map(|p| [p[0], p[1], p[2]]).collect();

                    // SMask (soft mask / alpha channel)
                    use lopdf::Dictionary as LopdfDict;
                    let smask_dict = LopdfDict::from_iter(vec![
                        ("Type", lopdf::Object::Name(b"XObject".to_vec())),
                        ("Subtype", lopdf::Object::Name(b"Image".to_vec())),
                        ("Width", lopdf::Object::Integer(w as i64)),
                        ("Height", lopdf::Object::Integer(h as i64)),
                        ("ColorSpace", lopdf::Object::Name(b"DeviceGray".to_vec())),
                        ("BitsPerComponent", lopdf::Object::Integer(8)),
                    ]);
                    let smask_stream = lopdf::Stream::new(smask_dict, alpha_bytes).with_compression(true);
                    let smask_id = output_doc.add_object(smask_stream);

                    // RGB image with SMask reference
                    let img_dict = LopdfDict::from_iter(vec![
                        ("Type", lopdf::Object::Name(b"XObject".to_vec())),
                        ("Subtype", lopdf::Object::Name(b"Image".to_vec())),
                        ("Width", lopdf::Object::Integer(w as i64)),
                        ("Height", lopdf::Object::Integer(h as i64)),
                        ("ColorSpace", lopdf::Object::Name(b"DeviceRGB".to_vec())),
                        ("BitsPerComponent", lopdf::Object::Integer(8)),
                        ("SMask", lopdf::Object::Reference(smask_id)),
                    ]);
                    let img_stream = lopdf::Stream::new(img_dict, rgb_bytes).with_compression(true);
                    let text_xobj_id = output_doc.add_object(img_stream);

                    // Calculate position: bottom-center of page
                    let img_w_pt = w as f32 * 72.0 / RENDER_DPI as f32;
                    let img_h_pt = h as f32 * 72.0 / RENDER_DPI as f32;
                    let x_pt = (pw_pt - img_w_pt) / 2.0; // centered horizontally
                    let y_pt = 3.0 * MM_TO_PT; // 3mm from bottom edge

                    // Append to content stream: save state, position, draw image, restore state
                    use lopdf::content::Operation;
                    let name = b"TxtOverlay".to_vec();
                    let mut text_ops = Vec::new();
                    text_ops.push(Operation { operator: "q".into(), operands: vec![] });
                    text_ops.push(Operation { operator: "cm".into(), operands: vec![
                        lopdf::Object::Real(img_w_pt),
                        lopdf::Object::Real(0.0),
                        lopdf::Object::Real(0.0),
                        lopdf::Object::Real(img_h_pt),
                        lopdf::Object::Real(x_pt),
                        lopdf::Object::Real(y_pt),
                    ]});
                    text_ops.push(Operation { operator: "Do".into(), operands: vec![
                        lopdf::Object::Name(name.clone()),
                    ]});
                    text_ops.push(Operation { operator: "Q".into(), operands: vec![] });

                    let text_content = lopdf::content::Content { operations: text_ops };
                    if let Ok(text_bytes) = text_content.encode() {
                        // Add separator newline before text overlay ops to avoid
                        // merging with the last operation of the main content stream
                        // (e.g. "Qq" would be an invalid token)
                        if !content_bytes.is_empty() {
                            content_bytes.push(b'\n');
                        }
                        content_bytes.extend_from_slice(&text_bytes);
                        xobj_names.push((name, text_xobj_id));
                        log::info!("lopdf hybrid: page {} text overlay added ({}x{} at x={:.1} y={:.1})",
                            page_idx, w, h, x_pt, y_pt);
                    } else {
                        log::warn!("lopdf hybrid: page {} text overlay content encode failed", page_idx);
                    }
                }
                Err(e) => {
                    log::warn!("lopdf hybrid: page {} text overlay PNG decode failed: {}", page_idx, e);
                }
            }
        } else {
            log::warn!("lopdf hybrid: page {} render_text_overlay returned None", page_idx);
        }

        // Add slot numbers if enabled
        if request.settings.number {
            log::info!("lopdf hybrid: page {} adding slot numbers", page_idx);
            // slot_positions.len() = 实际每页 slot 数（常规 = rows*cols，报销单 = 分段数）
            let start_num = page_idx * slot_positions.len() + 1;
            let num_images = render_slot_numbers(&text_font, &slot_positions, start_num);
            // Only add numbers for filled slots, matched by index
            for slot_idx in filled_slot_indices.iter() {
                let num_idx = *slot_idx;
                if num_idx >= num_images.len() {
                    continue;
                }
                let (png_bytes, _w, _h) = &num_images[num_idx];
                match image::load_from_memory(&png_bytes) {
                    Ok(rgba_img) => {
                        let rgba_img = rgba_img.to_rgba8();
                        let (w, h) = rgba_img.dimensions();
                        let alpha_bytes: Vec<u8> = rgba_img.pixels().map(|p| p[3]).collect();
                        let rgb_bytes: Vec<u8> = rgba_img.pixels().flat_map(|p| [p[0], p[1], p[2]]).collect();

                        use lopdf::Dictionary as LopdfDict;
                        let smask_dict = LopdfDict::from_iter(vec![
                            ("Type", lopdf::Object::Name(b"XObject".to_vec())),
                            ("Subtype", lopdf::Object::Name(b"Image".to_vec())),
                            ("Width", lopdf::Object::Integer(w as i64)),
                            ("Height", lopdf::Object::Integer(h as i64)),
                            ("ColorSpace", lopdf::Object::Name(b"DeviceGray".to_vec())),
                            ("BitsPerComponent", lopdf::Object::Integer(8)),
                        ]);
                        let smask_stream = lopdf::Stream::new(smask_dict, alpha_bytes).with_compression(true);
                        let smask_id = output_doc.add_object(smask_stream);

                        let img_dict = LopdfDict::from_iter(vec![
                            ("Type", lopdf::Object::Name(b"XObject".to_vec())),
                            ("Subtype", lopdf::Object::Name(b"Image".to_vec())),
                            ("Width", lopdf::Object::Integer(w as i64)),
                            ("Height", lopdf::Object::Integer(h as i64)),
                            ("ColorSpace", lopdf::Object::Name(b"DeviceRGB".to_vec())),
                            ("BitsPerComponent", lopdf::Object::Integer(8)),
                            ("SMask", lopdf::Object::Reference(smask_id)),
                        ]);
                        let img_stream = lopdf::Stream::new(img_dict, rgb_bytes).with_compression(true);
                        let num_xobj_id = output_doc.add_object(img_stream);

                        let slot = &slot_positions[*slot_idx];
                        let img_w_pt = w as f32 * 72.0 / RENDER_DPI as f32;
                        let img_h_pt = h as f32 * 72.0 / RENDER_DPI as f32;
                        let padding_pt = 2.0 * MM_TO_PT;
                        let x_pt = (slot.x_mm + slot.w_mm) * MM_TO_PT - img_w_pt - padding_pt;
                        let y_pt = (slot.y_mm + slot.h_mm) * MM_TO_PT - img_h_pt;

                        use lopdf::content::Operation;
                        let name = format!("Num{}", *slot_idx);
                        let mut num_ops = Vec::new();
                        num_ops.push(Operation { operator: "q".into(), operands: vec![] });
                        num_ops.push(Operation { operator: "cm".into(), operands: vec![
                            lopdf::Object::Real(img_w_pt),
                            lopdf::Object::Real(0.0),
                            lopdf::Object::Real(0.0),
                            lopdf::Object::Real(img_h_pt),
                            lopdf::Object::Real(x_pt),
                            lopdf::Object::Real(y_pt),
                        ]});
                        num_ops.push(Operation { operator: "Do".into(), operands: vec![
                            lopdf::Object::Name(name.clone().into_bytes()),
                        ]});
                        num_ops.push(Operation { operator: "Q".into(), operands: vec![] });

                        let num_content = lopdf::content::Content { operations: num_ops };
                        if let Ok(num_bytes) = num_content.encode() {
                            if !content_bytes.is_empty() {
                                content_bytes.push(b'\n');
                            }
                            content_bytes.extend_from_slice(&num_bytes);
                            xobj_names.push((name.into_bytes(), num_xobj_id));
                            log::info!("lopdf hybrid: page {} number added at x={:.1} y={:.1}", page_idx, x_pt, y_pt);
                        }
                    }
                    Err(e) => {
                        log::warn!("lopdf hybrid: page {} number decode failed: {}", page_idx, e);
                    }
                }
            }
        }

        // Add watermark if enabled
        if request.settings.watermark {
            let wm_text = request.settings.watermark_text.clone().unwrap_or_default();
            if !wm_text.is_empty() {
                let wm_color = request.settings.watermark_color.clone();
                let wm_opacity = request.settings.watermark_opacity;
                let wm_size = request.settings.watermark_size;
                let wm_angle = request.settings.watermark_angle;
                if let Some((png_bytes, _w, _h)) = render_watermark(&text_font, &wm_text, &wm_color, wm_opacity, wm_size, wm_angle) {
                    match image::load_from_memory(&png_bytes) {
                        Ok(rgba_img) => {
                            let rgba_img = rgba_img.to_rgba8();
                            let (w, h) = rgba_img.dimensions();
                            let alpha_bytes: Vec<u8> = rgba_img.pixels().map(|p| p[3]).collect();
                            let rgb_bytes: Vec<u8> = rgba_img.pixels().flat_map(|p| [p[0], p[1], p[2]]).collect();

                            use lopdf::Dictionary as LopdfDict;
                            let smask_dict = LopdfDict::from_iter(vec![
                                ("Type", lopdf::Object::Name(b"XObject".to_vec())),
                                ("Subtype", lopdf::Object::Name(b"Image".to_vec())),
                                ("Width", lopdf::Object::Integer(w as i64)),
                                ("Height", lopdf::Object::Integer(h as i64)),
                                ("ColorSpace", lopdf::Object::Name(b"DeviceGray".to_vec())),
                                ("BitsPerComponent", lopdf::Object::Integer(8)),
                            ]);
                            let smask_stream = lopdf::Stream::new(smask_dict, alpha_bytes).with_compression(true);
                            let smask_id = output_doc.add_object(smask_stream);

                            let img_dict = LopdfDict::from_iter(vec![
                                ("Type", lopdf::Object::Name(b"XObject".to_vec())),
                                ("Subtype", lopdf::Object::Name(b"Image".to_vec())),
                                ("Width", lopdf::Object::Integer(w as i64)),
                                ("Height", lopdf::Object::Integer(h as i64)),
                                ("ColorSpace", lopdf::Object::Name(b"DeviceRGB".to_vec())),
                                ("BitsPerComponent", lopdf::Object::Integer(8)),
                                ("SMask", lopdf::Object::Reference(smask_id)),
                            ]);
                            let img_stream = lopdf::Stream::new(img_dict, rgb_bytes).with_compression(true);
                            let wm_xobj_id = output_doc.add_object(img_stream);

                            for slot_idx in filled_slot_indices.iter() {
                                let slot_idx = *slot_idx;
                                if slot_idx >= slot_positions.len() {
                                    continue;
                                }
                                let slot = &slot_positions[slot_idx];
                                let img_w_pt = w as f32 * 72.0 / RENDER_DPI as f32;
                                let img_h_pt = h as f32 * 72.0 / RENDER_DPI as f32;
                                let mut x_pt = slot.x_mm * MM_TO_PT + (slot.w_mm * MM_TO_PT - img_w_pt) / 2.0;
                                let mut y_pt = slot.y_mm * MM_TO_PT + (slot.h_mm * MM_TO_PT - img_h_pt) / 2.0;

                                // 报销单模式：发票左上对齐，水印须跟随图像中心而非 slot 中心，
                                // 定位公式与 build_nup_content_stream 的 reimburse 分支一致
                                if request.settings.reimburse_mode {
                                    if let Some(fx_pos) = page_form_xobjs.iter().position(|(idx, _, _, _)| *idx == slot_idx) {
                                        let (_, _, src_w_pt, src_h_pt) = &page_form_xobjs[fx_pos];
                                        let adj = slot_adjustments.get(fx_pos);
                                        let rot = adj.map(|a| ((a.rotation % 360) + 360) % 360).unwrap_or(0);
                                        let (vis_w, vis_h) = if rot == 90 || rot == 270 { (*src_h_pt, *src_w_pt) } else { (*src_w_pt, *src_h_pt) };
                                        let slot_w_pt = slot.w_mm * MM_TO_PT;
                                        let slot_h_pt = slot.h_mm * MM_TO_PT;
                                        let (mut scx, mut scy) = match request.settings.fit_mode.as_str() {
                                            "fill" => (slot_w_pt / vis_w, slot_h_pt / vis_h),
                                            "original" => (1.0, 1.0),
                                            "custom" => {
                                                let c = (slot_w_pt / vis_w).min(slot_h_pt / vis_h) * request.settings.custom_scale;
                                                (c, c)
                                            }
                                            _ => {
                                                let c = (slot_w_pt / vis_w).min(slot_h_pt / vis_h);
                                                (c, c)
                                            }
                                        };
                                        if let Some(a) = adj {
                                            if a.scale != 1.0 { scx *= a.scale; scy *= a.scale; }
                                        }
                                        let draw_w = vis_w * scx;
                                        let draw_h = vis_h * scy;
                                        let mut ox = slot.x_mm * MM_TO_PT;
                                        let mut oy = slot.y_mm * MM_TO_PT + (slot_h_pt - draw_h);
                                        if let Some(a) = adj {
                                            if a.offset_x != 0.0 { ox += a.offset_x * MM_TO_PT; }
                                            if a.offset_y != 0.0 { oy -= a.offset_y * MM_TO_PT; }
                                        }
                                        x_pt = ox + (draw_w - img_w_pt) / 2.0;
                                        y_pt = oy + (draw_h - img_h_pt) / 2.0;
                                    }
                                }

                                use lopdf::content::Operation;
                                let name = format!("Wm{}", page_idx * 100 + slot_idx);
                                let mut wm_ops = Vec::new();
                                wm_ops.push(Operation { operator: "q".into(), operands: vec![] });
                                wm_ops.push(Operation { operator: "cm".into(), operands: vec![
                                    lopdf::Object::Real(img_w_pt),
                                    lopdf::Object::Real(0.0),
                                    lopdf::Object::Real(0.0),
                                    lopdf::Object::Real(img_h_pt),
                                    lopdf::Object::Real(x_pt),
                                    lopdf::Object::Real(y_pt),
                                ]});
                                wm_ops.push(Operation { operator: "Do".into(), operands: vec![
                                    lopdf::Object::Name(name.clone().into_bytes()),
                                ]});
                                wm_ops.push(Operation { operator: "Q".into(), operands: vec![] });

                                let wm_content = lopdf::content::Content { operations: wm_ops };
                                if let Ok(wm_bytes) = wm_content.encode() {
                                    if !content_bytes.is_empty() {
                                        content_bytes.push(b'\n');
                                    }
                                    content_bytes.extend_from_slice(&wm_bytes);
                                    xobj_names.push((name.into_bytes(), wm_xobj_id));
                                }
                            }
                            log::info!("lopdf hybrid: page {} watermark added", page_idx);
                        }
                        Err(e) => {
                            log::warn!("lopdf hybrid: page {} watermark decode failed: {}", page_idx, e);
                        }
                    }
                }
            }
        }

        // Add cut lines if enabled
        // 报销单模式：段边界绝对位置（k×段高）强制裁切线，不受 cutline 开关与边距影响
        let cutline_content = if request.settings.reimburse_mode {
            build_reimburse_cutline_ops_lopdf(pw_pt, ph_pt, request.settings.reimburse_height.unwrap_or(120.0))
        } else if request.settings.cutline {
            // Calculate footer cut line position
            let has_footer = request.settings.page_num || request.settings.print_date || request.settings.footer_text.as_ref().map_or(false, |t| !t.is_empty());
            let footer_cutline_y_pt = if has_footer {
                if request.settings.custom_fm && request.settings.footer_margin > 0.0 {
                    // 自定义下边距模式：分割线在 fm 位置
                    Some(request.settings.footer_margin * MM_TO_PT)
                } else {
                    // 默认模式：分割线在页脚文本顶部 + 2mm 间隙，避免贴文字
                    // 文本布局（从底部起）：3mm底部间距 + 行数×5mm行高
                    let line_count = (if request.settings.page_num || request.settings.print_date { 1 } else { 0 })
                        + (if request.settings.footer_text.as_ref().map_or(false, |t| !t.is_empty()) { 1 } else { 0 });
                    let footer_text_top_mm = 3.0 + line_count as f32 * 5.0 + 2.0;
                    Some(footer_text_top_mm * MM_TO_PT)
                }
            } else {
                None
            };
            build_cutline_ops_lopdf(&slot_positions, pw_pt, ph_pt, footer_cutline_y_pt)
        } else {
            None
        };
        if let Some(cutline_ops) = cutline_content {
            if !content_bytes.is_empty() {
                content_bytes.push(b'\n');
            }
            if let Ok(cutline_bytes) = cutline_ops.encode() {
                content_bytes.extend_from_slice(&cutline_bytes);
                log::info!("lopdf hybrid: page {} cut lines added", page_idx);
            } else {
                log::warn!("lopdf hybrid: page {} cut line encode failed", page_idx);
            }
        }

        // Add slot borders if enabled
        if request.settings.border {
            if let Some(border_ops) = build_border_ops_lopdf(&slot_positions, &slot_adjustments, &page_form_xobjs, &request.settings) {
                if !content_bytes.is_empty() {
                    content_bytes.push(b'\n');
                }
                if let Ok(border_bytes) = border_ops.encode() {
                    content_bytes.extend_from_slice(&border_bytes);
                    log::info!("lopdf hybrid: page {} borders added", page_idx);
                } else {
                    log::warn!("lopdf hybrid: page {} border encode failed", page_idx);
                }
            }
        }

        // Create the content stream object
        let content_id = output_doc.add_object(lopdf::Stream::new(
            lopdf::Dictionary::new(), content_bytes
        ).with_compression(true));

        // Build Resources dictionary with Form XObjects
        let mut xobjects_dict = lopdf::Dictionary::new();
        for (name, id) in &xobj_names {
            xobjects_dict.set(name.clone(), lopdf::Object::Reference(*id));
        }

        let mut resources_dict = lopdf::Dictionary::new();
        resources_dict.set(b"XObject".to_vec(), lopdf::Object::Dictionary(xobjects_dict));

        // Build the page object
        let mut page_dict = lopdf::Dictionary::new();
        page_dict.set("Type", lopdf::Object::Name(b"Page".to_vec()));
        page_dict.set("Parent", lopdf::Object::Reference(pages_id));
        page_dict.set("MediaBox", lopdf::Object::Array(vec![
            lopdf::Object::Real(0.0),
            lopdf::Object::Real(0.0),
            lopdf::Object::Real(pw_pt),
            lopdf::Object::Real(ph_pt),
        ]));
        page_dict.set("Contents", lopdf::Object::Reference(content_id));
        page_dict.set("Resources", lopdf::Object::Dictionary(resources_dict));

        let page_id = output_doc.add_object(lopdf::Object::Dictionary(page_dict));
        all_page_ids.push(page_id);

        // Report progress
        if let Some(ref cb) = &on_progress {
            cb("build", (page_idx + 1) as u32, request.pages.len() as u32);
        }
    }

    if all_page_ids.is_empty() {
        return Err("没有有效页面".to_string());
    }

    // Build the Pages tree
    let pages_dict = lopdf::Dictionary::from_iter(vec![
        ("Type", lopdf::Object::Name(b"Pages".to_vec())),
        ("Count", lopdf::Object::Integer(all_page_ids.len() as i64)),
        ("Kids", lopdf::Object::Array(
            all_page_ids.iter().map(|&id| lopdf::Object::Reference(id)).collect()
        )),
    ]);
    output_doc.set_object(pages_id, lopdf::Object::Dictionary(pages_dict));

    // Build the Catalog
    let catalog_id = output_doc.add_object(lopdf::Dictionary::from_iter(vec![
        ("Type", lopdf::Object::Name(b"Catalog".to_vec())),
        ("Pages", lopdf::Object::Reference(pages_id)),
    ]));
    output_doc.trailer.set("Root", lopdf::Object::Reference(catalog_id));

    // Save
    if let Some(ref cb) = &on_progress {
        cb("save", 0, 1);
    }
    let mut pdf_buf = Vec::new();
    output_doc.save_to(&mut pdf_buf)
        .map_err(|e| format!("PDF保存失败: {}", e))?;
    // Ensure parent directory exists (e.g. temp dir may not be created yet)
    if let Some(parent) = output_path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    std::fs::write(output_path, &pdf_buf)
        .map_err(|e| format!("写入文件失败: {}", e))?;
    if let Some(ref cb) = &on_progress {
        cb("save", 1, 1);
    }

    Ok(())
}

/// Build PDF operations for reimburse-mode cut lines: horizontal dashed lines
/// at fixed segment boundaries (k × seg_height from page top, absolute position).
/// Coordinate system: PDF content stream is bottom-up, so line k is at ph - k*seg.
/// Returns None when fewer than 2 segments fit on the page (no interior boundary).
fn build_reimburse_cutline_ops_lopdf(
    page_w_pt: f32,
    page_h_pt: f32,
    seg_height_mm: f32,
) -> Option<lopdf::content::Content> {
    use lopdf::content::Operation;

    let seg = seg_height_mm.max(10.0);
    let ph_mm = page_h_pt / MM_TO_PT;
    let seg_count = ((ph_mm / seg).floor() as usize).max(1);

    let mut ops = Vec::new();
    ops.push(Operation { operator: "q".into(), operands: vec![] });
    ops.push(Operation { operator: "w".into(), operands: vec![lopdf::Object::Real(0.5)] });
    ops.push(Operation { operator: "d".into(), operands: vec![
        lopdf::Object::Array(vec![
            lopdf::Object::Integer(3),
            lopdf::Object::Integer(3),
        ]),
        lopdf::Object::Integer(0),
    ]});
    ops.push(Operation { operator: "G".into(), operands: vec![lopdf::Object::Real(0.0)] });

    // 段底裁切线：k = 1..=seg_count，每段底部一条（含最后一段），top-down k*seg → bottom-up ph - k*seg
    for k in 1..=seg_count {
        // top-down k*seg → bottom-up ph - k*seg
        let y = page_h_pt - k as f32 * seg * MM_TO_PT;
        ops.push(Operation { operator: "m".into(), operands: vec![
            lopdf::Object::Real(0.0),
            lopdf::Object::Real(y),
        ]});
        ops.push(Operation { operator: "l".into(), operands: vec![
            lopdf::Object::Real(page_w_pt),
            lopdf::Object::Real(y),
        ]});
        ops.push(Operation { operator: "S".into(), operands: vec![] });
    }

    ops.push(Operation { operator: "Q".into(), operands: vec![] });
    Some(lopdf::content::Content { operations: ops })
}

/// Build PDF operations to draw cut lines (dashed lines at slot boundaries).
/// Only draws lines between slots (not at page edges).
/// Returns None if only 1 slot (no cut lines needed) and no footer cut line.
///
/// Coordinate system: lopdf content stream uses PDF standard bottom-left origin,
/// same as slot_positions (bottom-up y_mm). No conversion needed.
///
/// `footer_cutline_y_pt`: If Some(y), draw a horizontal cut line at that y position (bottom-up pt).
/// The caller decides the position — either at footer_margin_mm or at the top of footer text.
#[allow(dead_code)]
fn build_cutline_ops_lopdf(
    slot_positions: &[LayoutSlotMm],
    page_w_pt: f32,
    page_h_pt: f32,
    footer_cutline_y_pt: Option<f32>,  // bottom-up pt position for footer cut line
) -> Option<lopdf::content::Content> {
    use lopdf::content::Operation;
    use std::collections::BTreeSet;

    let need_row_lines = slot_positions.len() > 1;
    let need_footer_line = footer_cutline_y_pt.is_some();

    if !need_row_lines && !need_footer_line {
        return None;
    }

    // Infer grid dimensions from unique x/y positions
    let mut x_positions = BTreeSet::new();
    let mut y_positions = BTreeSet::new();

    for slot in slot_positions {
        x_positions.insert((slot.x_mm * 100.0).round() as i32);
        y_positions.insert((slot.y_mm * 100.0).round() as i32);
    }

    let cols = x_positions.len();
    let rows = y_positions.len();

    let mut ops = Vec::new();

    // Save graphics state
    ops.push(Operation { operator: "q".into(), operands: vec![] });

    // Set line width (0.5 pt)
    ops.push(Operation { operator: "w".into(), operands: vec![lopdf::Object::Real(0.5)] });

    // Set dash pattern: dashed line (3 pt dash, 3 pt gap)
    ops.push(Operation { operator: "d".into(), operands: vec![
        lopdf::Object::Array(vec![
            lopdf::Object::Integer(3),
            lopdf::Object::Integer(3),
        ]),
        lopdf::Object::Integer(0),
    ]});

    // Set stroke color to black
    ops.push(Operation { operator: "G".into(), operands: vec![lopdf::Object::Real(0.0)] });

    // Draw vertical cut lines (between columns)
    // Both slot_positions and PDF content stream use bottom-left origin, x increases right.
    // Stop at footer area if present (draw only above footer, not into it).
    if cols > 1 {
        // In bottom-up coords: 0.0 is page bottom, page_h_pt is page top.
        // If footer exists: draw from footer_cutline_y_pt (top of footer) UP to page top.
        // If no footer: draw entire page from bottom to top.
        let (y_start, y_end) = if let Some(fy) = footer_cutline_y_pt {
            (fy, page_h_pt)
        } else {
            (0.0, page_h_pt)
        };
        for c in 1..cols {
            // slot[c-1] right edge and slot[c] left edge
            let left_slot = &slot_positions[(c - 1) as usize]; // row 0, col c-1
            let right_slot = &slot_positions[c as usize];       // row 0, col c
            let right_edge_pt = (left_slot.x_mm + left_slot.w_mm) * MM_TO_PT;
            let left_edge_pt = right_slot.x_mm * MM_TO_PT;
            let x = (right_edge_pt + left_edge_pt) / 2.0;
            ops.push(Operation { operator: "m".into(), operands: vec![
                lopdf::Object::Real(x),
                lopdf::Object::Real(y_start),
            ]});
            ops.push(Operation { operator: "l".into(), operands: vec![
                lopdf::Object::Real(x),
                lopdf::Object::Real(y_end),
            ]});
            ops.push(Operation { operator: "S".into(), operands: vec![] });
        }
    }

    // Draw horizontal cut lines (between rows)
    // slot_positions y_mm is bottom-up, PDF content stream is also bottom-up.
    // Between row r-1 (bottom) and row r (top): gap center in bottom-up pt.
    if rows > 1 {
        for r in 1..rows {
            // bottom_slot is the row closer to the page bottom (row_from_bottom = r-1)
            // top_slot is the row above it (row_from_bottom = r)
            let bottom_slot = &slot_positions[(r - 1) as usize * cols]; // row r-1 from bottom
            let top_slot = &slot_positions[r as usize * cols];           // row r from bottom
            // bottom_slot top edge (bottom-up) = bottom_slot.y_mm + bottom_slot.h_mm
            // top_slot bottom edge (bottom-up) = top_slot.y_mm
            // Gap center in bottom-up mm = average, then convert to pt
            let gap_center_mm = ((bottom_slot.y_mm + bottom_slot.h_mm) + top_slot.y_mm) / 2.0;
            let y = gap_center_mm * MM_TO_PT;
            ops.push(Operation { operator: "m".into(), operands: vec![
                lopdf::Object::Real(0.0),
                lopdf::Object::Real(y),
            ]});
            ops.push(Operation { operator: "l".into(), operands: vec![
                lopdf::Object::Real(page_w_pt),
                lopdf::Object::Real(y),
            ]});
            ops.push(Operation { operator: "S".into(), operands: vec![] });
        }
    }

    // Draw footer cut line at caller-specified position (bottom-up pt)
    if let Some(y) = footer_cutline_y_pt {
        ops.push(Operation { operator: "m".into(), operands: vec![
            lopdf::Object::Real(0.0),
            lopdf::Object::Real(y),
        ]});
        ops.push(Operation { operator: "l".into(), operands: vec![
            lopdf::Object::Real(page_w_pt),
            lopdf::Object::Real(y),
        ]});
        ops.push(Operation { operator: "S".into(), operands: vec![] });
    }

    // Restore graphics state
    ops.push(Operation { operator: "Q".into(), operands: vec![] });

    Some(lopdf::content::Content { operations: ops })
}

/// Draw borders around each invoice's visual boundary (follows per-slot adjustments).
/// Borders are drawn at the actual image position, not the slot boundary.
fn build_border_ops_lopdf(
    slot_positions: &[LayoutSlotMm],
    slot_adjustments: &[SlotAdjustment],
    form_xobjs: &[(usize, lopdf::ObjectId, f32, f32)],  // (layout_slot_idx, _, src_w_pt, src_h_pt)
    settings: &RenderSettings,
) -> Option<lopdf::content::Content> {
    use lopdf::content::Operation;

    if form_xobjs.is_empty() {
        return None;
    }

    let mut ops = Vec::new();

    ops.push(Operation { operator: "q".into(), operands: vec![] });

    ops.push(Operation { operator: "w".into(), operands: vec![lopdf::Object::Real(1.0)] });

    ops.push(Operation { operator: "d".into(), operands: vec![
        lopdf::Object::Array(vec![
            lopdf::Object::Integer(1),
            lopdf::Object::Integer(0),
        ]),
        lopdf::Object::Integer(0),
    ]});

    ops.push(Operation { operator: "G".into(), operands: vec![lopdf::Object::Real(0.0)] });

    for (adj_idx, (layout_slot_idx, _xobj_id, src_w_pt, src_h_pt)) in form_xobjs.iter().enumerate() {
        let slot = &slot_positions[*layout_slot_idx];
        let slot_w_pt = slot.w_mm * MM_TO_PT;
        let slot_h_pt = slot.h_mm * MM_TO_PT;

        let adj = if adj_idx < slot_adjustments.len() {
            &slot_adjustments[adj_idx]
        } else {
            continue;
        };
        let rotation = adj.rotation;
        let rot = ((rotation % 360) + 360) % 360;

        // For 90°/270° rotation, visual dimensions swap (width↔height)
        let (vis_w, vis_h) = if rot == 90 || rot == 270 {
            (*src_h_pt, *src_w_pt)
        } else {
            (*src_w_pt, *src_h_pt)
        };

        // Compute scale to fit in slot based on visual dimensions
        let (mut scale_x, mut scale_y) = match settings.fit_mode.as_str() {
            "fill" => (slot_w_pt / vis_w, slot_h_pt / vis_h),
            "original" => (1.0, 1.0),
            "custom" => {
                let contain_s = (slot_w_pt / vis_w).min(slot_h_pt / vis_h);
                let s = contain_s * settings.custom_scale;
                (s, s)
            }
            _ => {
                let s = (slot_w_pt / vis_w).min(slot_h_pt / vis_h);
                (s, s)
            }
        };

        // Per-slot scale override
        if adj.scale != 1.0 {
            scale_x *= adj.scale;
            scale_y *= adj.scale;
        }

        // Calculate visual boundary (bottom-up coordinates, PDF standard).
        // 报销单模式：左上对齐；常规模式：居中（与 build_nup_content_stream 保持一致）。
        let draw_w = vis_w * scale_x;
        let draw_h = vis_h * scale_y;
        let mut offset_x = slot.x_mm * MM_TO_PT;
        let mut offset_y = if settings.reimburse_mode {
            slot.y_mm * MM_TO_PT + (slot_h_pt - draw_h)
        } else {
            slot.y_mm * MM_TO_PT + (slot_h_pt - draw_h) / 2.0
        };
        if !settings.reimburse_mode {
            offset_x += (slot_w_pt - draw_w) / 2.0;
        }

        // Per-slot offset override
        if adj.offset_x != 0.0 { offset_x += adj.offset_x * MM_TO_PT; }
        if adj.offset_y != 0.0 { offset_y -= adj.offset_y * MM_TO_PT; }  // JS Y+ down, PDF Y+ up

        // Draw rectangle at invoice visual boundary
        let x1 = offset_x;
        let y1 = offset_y;
        let x2 = offset_x + draw_w;
        let y2 = offset_y + draw_h;

        ops.push(Operation { operator: "m".into(), operands: vec![
            lopdf::Object::Real(x1),
            lopdf::Object::Real(y1),
        ]});
        ops.push(Operation { operator: "l".into(), operands: vec![
            lopdf::Object::Real(x2),
            lopdf::Object::Real(y1),
        ]});
        ops.push(Operation { operator: "l".into(), operands: vec![
            lopdf::Object::Real(x2),
            lopdf::Object::Real(y2),
        ]});
        ops.push(Operation { operator: "l".into(), operands: vec![
            lopdf::Object::Real(x1),
            lopdf::Object::Real(y2),
        ]});
        ops.push(Operation { operator: "h".into(), operands: vec![] });
        ops.push(Operation { operator: "S".into(), operands: vec![] });
    }

    ops.push(Operation { operator: "Q".into(), operands: vec![] });

    Some(lopdf::content::Content { operations: ops })
}
