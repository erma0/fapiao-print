use image::GenericImageView;
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
    /// Can be embedded as DCTDecode stream directly via ExternalXObject,
    /// no decode/re-encode needed — zero quality loss, smaller file size.
    JpegPassthrough {
        raw_bytes: Vec<u8>,
        width: u32,
        height: u32,
        /// Number of color components: 1=grayscale, 3=RGB, 4=CMYK
        num_components: u8,
    },
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
    /// Base64-encoded PNG data URL
    pub image_data_url: String,
    pub width: u32,
    pub height: u32,
    /// Actual DPI used for rendering (may differ from requested DPI due to adaptive scaling)
    pub render_dpi: u32,
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

/// Render a single PDF page and run OCR on it — zero IPC round-trip for OCR.
/// The frontend calls this instead of `render_pdf_pages` + `ocr_image` to avoid:
///   Rust render → base64 → IPC → frontend → downsample → base64 → IPC → Rust decode → OCR
/// Instead: Rust render → decode in memory → OCR → return result directly.
/// Returns OcrResult with coordinates in the original (full-DPI) pixel space.
#[cfg(all(target_os = "windows", feature = "ocr"))]
pub(crate) fn ocr_pdf_page(pdf_path: &str, page_index: u32, dpi: Option<u32>) -> Result<OcrResult, String> {
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

    // Run OCR directly on the decoded image
    run_ocr_on_image(img)
}

/// Render PDF pages and run OCR in one pass — avoids the IPC round-trip
/// where the frontend sends the rendered dataUrl back to Rust for OCR.
/// The image is decoded from PNG bytes ONCE, OCR'd, then base64-encoded for preview.
#[cfg(all(target_os = "windows", feature = "ocr"))]
pub(crate) fn render_and_ocr_pdf(pdf_path: &str, dpi: u32) -> Result<Vec<RenderedOcrPage>, String> {
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

                    // Resize for OCR if needed (same logic as run_ocr_on_image)
                    let ocr_img = if longest > OCR_MAX_DIM {
                        let rscale = OCR_MAX_DIM as f32 / longest as f32;
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

            // Decode image and capture original dimensions
            let (thumbnail_data_url, img_orig_w, img_orig_h) = match image::load_from_memory(&bytes) {
                Ok(img) => {
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
///   - PP-OCRv5_mobile_det.mnn  (detection model)
///   - PP-OCRv5_mobile_rec.mnn  (recognition model)
///   - ppocr_keys_v5.txt        (character set, 18383 chars)
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
        let det_path = if exe_dir.join("models").join("PP-OCRv5_mobile_det.mnn").exists() {
            exe_dir.join("models").join("PP-OCRv5_mobile_det.mnn")
        } else {
            exe_dir.join("PP-OCRv5_mobile_det.mnn")
        };
        let rec_path = if exe_dir.join("models").join("PP-OCRv5_mobile_rec.mnn").exists() {
            exe_dir.join("models").join("PP-OCRv5_mobile_rec.mnn")
        } else {
            exe_dir.join("PP-OCRv5_mobile_rec.mnn")
        };
        let keys_path = if exe_dir.join("models").join("ppocr_keys_v5.txt").exists() {
            exe_dir.join("models").join("ppocr_keys_v5.txt")
        } else {
            exe_dir.join("ppocr_keys_v5.txt")
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

/// Maximum longest-side dimension for OCR input.
/// 1280px: balances accuracy and speed. v1.6.7 used full resolution (2480×3508 for 300DPI A4)
/// which was more accurate but slower. 960 was too aggressive — small text (密码区/备注栏/明细行)
/// got blurred. 1280 preserves detail while keeping detection model in its optimal range.
#[cfg(feature = "ocr")]
const OCR_MAX_DIM: u32 = 1280;

/// OCR an image from a file path or base64 data URL.
/// When `file_path` is provided, reads the image directly from disk — skipping
/// the expensive base64 encode→IPC→decode round-trip.
/// Falls back to `data_url` when `file_path` is None or file read fails.
#[cfg(feature = "ocr")]
pub fn ocr_image(data_url: &str, file_path: Option<&str>) -> Result<OcrResult, String> {
    // Try file_path first (skip base64 entirely)
    if let Some(path) = file_path {
        if !path.is_empty() {
            match std::fs::read(path) {
                Ok(bytes) => {
                    if !bytes.is_empty() {
                        match image::load_from_memory(&bytes) {
                            Ok(img) => {
                                log::info!("OCR from file_path: {} ({}x{})", path, img.width(), img.height());
                                return run_ocr_on_image(img);
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
    // Fallback to data_url
    ocr_image_from_data(data_url)
}

/// OCR an image from base64 data URL, return structured result with coordinates.
/// Internal helper — prefer `ocr_image()` which supports file_path.
#[cfg(feature = "ocr")]
pub fn ocr_image_from_data(data_url: &str) -> Result<OcrResult, String> {
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

    // Decode image using the `image` crate
    let img = image::load_from_memory(&bytes)
        .map_err(|e| format!("图片解码失败: {}", e))?;

    run_ocr_on_image(img)
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
fn run_ocr_on_image(mut img: image::DynamicImage) -> Result<OcrResult, String> {
    use std::time::Instant;
    let t0 = Instant::now();

    if SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    // Resize for OCR if image is larger than OCR_MAX_DIM on the longest side.
    // We keep the original dimensions for coordinate reporting so the frontend
    // can normalize correctly.
    let orig_w = img.width();
    let orig_h = img.height();
    let longest = orig_w.max(orig_h);

    if longest > OCR_MAX_DIM {
        let scale = OCR_MAX_DIM as f32 / longest as f32;
        let new_w = (orig_w as f32 * scale).round() as u32;
        let new_h = (orig_h as f32 * scale).round() as u32;
        // Lanczos3 produces sharper text edges than Triangle — critical for OCR accuracy
        img = img.resize_exact(new_w, new_h, image::imageops::FilterType::Lanczos3);
        log::info!(
            "OCR resize: {}x{} → {}x{} ({}ms)",
            orig_w, orig_h, new_w, new_h,
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

// =====================================================
// OFD Format Support
// =====================================================

/// Extract embedded images from an OFD file (Chinese electronic invoice format)
/// OFD is a ZIP archive containing XML page descriptions and image resources.
/// For electronic invoices, the content is typically a full-page image.
///
/// Filtering strategy:
/// 1. Path-based: exclude Seals/, Signs/ directories (stamp/signature images)
/// 2. Dimension-based: prefer images where the longest side >= 500px
///    (QR codes ~100-200px, seal stamps ~300-400px; full invoice pages > 800px)
///    If large images exist, small ones are filtered out.
///    If NO large images exist (vector-based OFD), fall back to including all path-filtered images.
/// 3. Per-page dedup: keep only the largest image per page index
fn extract_ofd_images(ofd_path: &str) -> Result<Vec<(String, String, u32, u32)>, String> {
    use base64::Engine;
    use std::io::Read;

    let file = std::fs::File::open(ofd_path)
        .map_err(|e| format!("打开OFD文件失败: {}", e))?;

    let mut archive = zip::ZipArchive::new(file)
        .map_err(|e| format!("解析OFD ZIP失败: {}", e))?;

    // Collect candidate image entries with path-based filtering
    // OFD structure:
    //   Doc_0/Pages/Page_0/Res/xxx.jpg   — per-page resources (invoice image, QR code)
    //   Doc_0/Res/xxx.jpg                 — document-level resources
    //   Doc_0/Seals/xxx.jpg               — seal/stamp images (EXCLUDE)
    //   Doc_0/Signs/xxx.jpg               — signature images (EXCLUDE)
    let mut image_entries: Vec<String> = Vec::new();

    for i in 0..archive.len() {
        let entry = archive.by_index(i).map_err(|e| format!("读取ZIP条目失败: {}", e))?;
        let name = entry.name().to_string();
        let lower = name.to_lowercase();

        // Path-based exclusion: skip Seals/, Signs/ directories and sign_/seal_ filenames
        let path_has_seal_or_sign = lower.contains("/seals/")
            || lower.contains("/signs/")
            || lower.contains("\\seals\\")
            || lower.contains("\\signs\\")
            || lower.contains("sign_")
            || lower.contains("seal_");

        if (lower.ends_with(".jpg") || lower.ends_with(".jpeg") || lower.ends_with(".png"))
            && !path_has_seal_or_sign
        {
            image_entries.push(name);
        }
    }

    if image_entries.is_empty() {
        return Err("OFD文件中未找到图片资源".to_string());
    }

    // Extract page index from path for grouping
    fn extract_page_index(path: &str) -> u32 {
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

    // Read and decode all candidate images, collect (data_url, ext, w, h, page_idx)
    const MIN_LONGEST_SIDE: u32 = 500; // Full invoice pages are always > 500px; QR codes/seals are smaller
    let mut all_decoded: Vec<(String, String, u32, u32, u32)> = Vec::new(); // (data_url, ext, w, h, page_idx)

    for entry_name in &image_entries {
        let mut entry = archive.by_name(entry_name)
            .map_err(|e| format!("读取OFD图片失败: {}", e))?;
        let mut data = Vec::new();
        entry.read_to_end(&mut data)
            .map_err(|e| format!("读取OFD图片数据失败: {}", e))?;

        // Decode image to get dimensions
        let (w, h) = match image::load_from_memory(&data) {
            Ok(img) => img.dimensions(),
            Err(_) => {
                log::warn!("OFD: 无法解码图片 {}, 跳过", entry_name);
                continue;
            }
        };

        // Determine MIME type and extension
        let lower = entry_name.to_lowercase();
        let (mime, img_ext) = if lower.ends_with(".png") {
            ("image/png", "png")
        } else {
            ("image/jpeg", "jpg")
        };

        let b64 = base64::engine::general_purpose::STANDARD.encode(&data);
        let data_url = format!("data:{};base64,{}", mime, b64);
        let page_idx = extract_page_index(entry_name);
        let longest_side = w.max(h);

        log::info!("OFD: 图片 {} ({}x{}, longest={}, page_idx={})",
            entry_name, w, h, longest_side, page_idx);
        all_decoded.push((data_url, img_ext.to_string(), w, h, page_idx));
    }

    if all_decoded.is_empty() {
        return Err("OFD文件中未找到可解码的图片资源".to_string());
    }

    // Two-pass strategy:
    // Pass 1: Try to find large images (>= MIN_LONGEST_SIDE) — these are likely full invoice pages
    // Pass 2: If no large images found (vector-based OFD), fall back to all decoded images
    let large_images: Vec<_> = all_decoded.iter()
        .filter(|c| c.2.max(c.3) >= MIN_LONGEST_SIDE)
        .cloned()
        .collect();

    let candidates = if !large_images.is_empty() {
        log::info!("OFD: 找到{}张大图(>={}px)，过滤小图片", large_images.len(), MIN_LONGEST_SIDE);
        large_images
    } else {
        log::warn!("OFD: 未找到大图(>={}px)，可能是矢量版式OFD，回退到包含所有图片", MIN_LONGEST_SIDE);
        all_decoded
    };

    // Per-page dedup: keep only the largest image (by pixel count) per page index
    let mut sorted = candidates;
    sorted.sort_by(|a, b| {
        a.4.cmp(&b.4) // sort by page_idx first
            .then((b.2 * b.3).cmp(&(a.2 * a.3))) // then by pixel count descending
    });

    let mut seen_pages = std::collections::HashSet::new();
    let mut results = Vec::new();
    for (data_url, img_ext, w, h, page_idx) in sorted {
        if seen_pages.insert(page_idx) {
            results.push((data_url, img_ext, w, h));
        } else {
            log::info!("OFD: 页面{}已保留最大图片，跳过重复", page_idx);
        }
    }

    if results.is_empty() {
        return Err("OFD文件中未找到有效的发票页面图片（可能为矢量版式OFD，建议转换为PDF后使用）".to_string());
    }

    log::info!("OFD extracted {} page images from {}", results.len(), ofd_path);
    Ok(results)
}

/// Public wrapper: extract OFD images and return as Vec<FileData> for frontend fallback.
/// Called by the `open_ofd_images` Tauri command when `parse_ofd` fails.
pub fn extract_ofd_images_as_filedata(ofd_path: &str) -> Result<Vec<FileData>, String> {
    let path = std::path::Path::new(ofd_path);
    let name = path.file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_else(|| "unknown".to_string());
    let size = path.metadata().ok().map(|m| m.len()).unwrap_or(0);

    let images = extract_ofd_images(ofd_path)?;
    let mut results = Vec::new();
    for (idx, (img_data_url, img_ext, img_w, img_h)) in images.iter().enumerate() {
        let base_name = if name.len() > 4 { &name[..name.len()-4] } else { &name };
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
            orig_w: Some(*img_w),
            orig_h: Some(*img_h),
        });
    }
    Ok(results)
}

// =====================================================
// OFD Vector Parsing & SVG Rendering
// =====================================================

/// Invoice data extracted from OFD XML
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct OfdInvoiceInfo {
    pub invoice_no: Option<String>,
    pub invoice_date: Option<String>,
    pub buyer_name: Option<String>,
    pub buyer_tax_id: Option<String>,
    pub seller_name: Option<String>,
    pub seller_tax_id: Option<String>,
    pub amount_no_tax: Option<f64>,
    pub tax_amount: Option<f64>,
    pub amount_tax: Option<f64>,
    pub invoice_type: Option<String>,
}

/// Result returned by parse_ofd command
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfdResult {
    pub svg: String,
    pub invoice_info: OfdInvoiceInfo,
    pub page_width: f64,
    pub page_height: f64,
}

// ----- Internal OFD structures -----

#[derive(Debug, Default)]
struct OfdFont {
    id: u32,
    font_name: String,
    family_name: String,
}

/// DrawParam — inherited styling for paths/text (from PublicRes.xml)
#[derive(Debug, Default, Clone)]
struct OfdDrawParam {
    id: u32,
    relative: Option<u32>,
    line_width: f64,
    stroke_color: Option<(u8, u8, u8)>,
    fill_color: Option<(u8, u8, u8)>,
}

#[derive(Debug, Default)]
#[allow(dead_code)]
struct OfdImage {
    id: u32,
    file_name: String,
    base64: String,
}

#[derive(Debug)]
struct OfdTextObject {
    id: u32,
    boundary: (f64, f64, f64, f64), // x, y, w, h
    font_id: u32,
    size: f64,
    ctm: Option<(f64, f64, f64, f64, f64, f64)>,
    text: String,
    delta_x: Vec<f64>,
    text_x: f64,
    text_y: f64,
    fill_color: Option<(u8, u8, u8)>,
    stroke_color: Option<(u8, u8, u8)>,
    alpha: Option<u8>,
    blend_mode: Option<String>,
    weight: u32, // OFD font weight: 400=normal, 700=bold
}

impl Default for OfdTextObject {
    fn default() -> Self {
        Self {
            id: 0,
            boundary: (0.0, 0.0, 0.0, 0.0),
            font_id: 0,
            size: 3.175,
            ctm: None,
            text: String::new(),
            delta_x: Vec::new(),
            text_x: 0.0,
            text_y: 0.0,
            fill_color: None,
            stroke_color: None,
            alpha: None,
            blend_mode: None,
            weight: 400, // Normal weight by default
        }
    }
}

#[derive(Debug, Default)]
struct OfdPathObject {
    id: u32,
    boundary: (f64, f64, f64, f64),
    line_width: f64,
    stroke_color: Option<(u8, u8, u8)>,
    fill_color: Option<(u8, u8, u8)>,
    fill: bool,
    abbreviated_data: String,
    alpha: Option<u8>,
}

#[derive(Debug, Default)]
struct OfdImageObject {
    id: u32,
    boundary: (f64, f64, f64, f64),
    resource_id: u32,
    ctm: Option<(f64, f64, f64, f64, f64, f64)>,
    blend_mode: Option<String>,
    alpha: Option<u8>,
}

/// Read a file from ZIP archive as string
fn zip_read_str(archive: &mut zip::ZipArchive<std::fs::File>, name: &str) -> Option<String> {
    use std::io::Read;
    let mut entry = archive.by_name(name).ok()?;
    let mut buf = String::new();
    entry.read_to_string(&mut buf).ok()?;
    Some(buf)
}

/// Read a file from ZIP archive as bytes
fn zip_read_bytes(archive: &mut zip::ZipArchive<std::fs::File>, name: &str) -> Option<Vec<u8>> {
    use std::io::Read;
    let mut entry = archive.by_name(name).ok()?;
    let mut buf = Vec::new();
    entry.read_to_end(&mut buf).ok()?;
    Some(buf)
}

/// Parse 2 floats from "x y" string
#[allow(dead_code)]
fn parse_f2(s: &str) -> Option<(f64, f64)> {
    let parts: Vec<&str> = s.split_whitespace().collect();
    if parts.len() >= 2 {
        Some((parts[0].parse().ok()?, parts[1].parse().ok()?))
    } else {
        None
    }
}

fn parse_f4(s: &str) -> Option<(f64, f64, f64, f64)> {
    let parts: Vec<&str> = s.split_whitespace().collect();
    if parts.len() >= 4 {
        Some((
            parts[0].parse().ok()?,
            parts[1].parse().ok()?,
            parts[2].parse().ok()?,
            parts[3].parse().ok()?,
        ))
    } else {
        None
    }
}

fn parse_f6(s: &str) -> Option<(f64, f64, f64, f64, f64, f64)> {
    let parts: Vec<&str> = s.split_whitespace().collect();
    if parts.len() >= 6 {
        Some((
            parts[0].parse().ok()?,
            parts[1].parse().ok()?,
            parts[2].parse().ok()?,
            parts[3].parse().ok()?,
            parts[4].parse().ok()?,
            parts[5].parse().ok()?,
        ))
    } else {
        None
    }
}

/// Parse OFD color value "R G B" → (r, g, b)
fn parse_color(s: &str) -> Option<(u8, u8, u8)> {
    let parts: Vec<&str> = s.split_whitespace().collect();
    if parts.len() >= 3 {
        Some((
            parts[0].parse().ok()?,
            parts[1].parse().ok()?,
            parts[2].parse().ok()?,
        ))
    } else {
        None
    }
}

/// Get attribute value by local name (ignoring namespace prefix)
fn attr_val(e: &quick_xml::events::BytesStart, local_name: &str) -> Option<String> {
    for a in e.attributes().flatten() {
        let key = a.key;
        let local = if let Some(pos) = key.0.iter().position(|&b| b == b':') {
            &key.as_ref()[pos + 1..]
        } else {
            key.as_ref()
        };
        if local == local_name.as_bytes() {
            return std::str::from_utf8(&a.value).ok().map(|s| s.to_string());
        }
    }
    None
}

/// Get element text content from a quick-xml reader (reads until End tag)
fn read_element_text(reader: &mut quick_xml::Reader<&[u8]>) -> String {
    use quick_xml::events::Event;
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(t)) => {
                if let Ok(s) = t.unescape() {
                    text.push_str(&s);
                }
            }
            Ok(Event::End(_)) | Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    text
}

/// Parse DeltaX attribute string into individual character offsets.
/// DeltaX formats:
///   - "3.175 3.175 3.175" — simple space-separated values
///   - "g 19 1.5875" — group: repeat next spacing 19 times at 1.5875
///   - "g 4 1.5875 3.175 g 2 1.5875 3.175" — mixed
fn parse_delta_x(s: &str) -> Vec<f64> {
    let mut result = Vec::new();
    let tokens: Vec<&str> = s.split_whitespace().collect();
    let mut i = 0;
    while i < tokens.len() {
        if tokens[i] == "g" && i + 2 < tokens.len() {
            // Group format: g count value [extra_value...]
            if let (Ok(count), Ok(val)) = (tokens[i + 1].parse::<usize>(), tokens[i + 2].parse::<f64>()) {
                for _ in 0..count {
                    result.push(val);
                }
                i += 3;
                // Check if there's an extra value after the group
                if i < tokens.len() && tokens[i] != "g" {
                    if let Ok(v) = tokens[i].parse::<f64>() {
                        result.push(v);
                        i += 1;
                    }
                }
            } else {
                i += 1;
            }
        } else if let Ok(v) = tokens[i].parse::<f64>() {
            result.push(v);
            i += 1;
        } else {
            i += 1;
        }
    }
    result
}

/// Build SVG text element from an OFD TextObject
fn build_svg_text(
    text_obj: &OfdTextObject,
    font_map: &std::collections::HashMap<u32, OfdFont>,
    _color_spaces: &std::collections::HashMap<u32, String>,
    scale_x: f64,
    scale_y: f64,
) -> String {
    if text_obj.text.is_empty() {
        return String::new();
    }

    let font = font_map.get(&text_obj.font_id);
    let font_family_raw = font.map(|f| {
        if !f.family_name.is_empty() { f.family_name.clone() } else { f.font_name.clone() }
    }).unwrap_or_else(|| "SimSun".to_string());

    // Font fallback: add generic CJK/serif/sans-serif fallbacks for cross-platform rendering.
    // SVG font-family is CSS: names with spaces need single quotes (attr value is in double quotes).
    let font_family = match font_family_raw.as_str() {
        "楷体" | "KaiTi" | "STKaiti" => "楷体, KaiTi, STKaiti, serif".to_string(),
        "宋体" | "SimSun" | "STSong" => "宋体, SimSun, STSong, serif".to_string(),
        "黑体" | "SimHei" | "STHeiti" => "黑体, SimHei, STHeiti, sans-serif".to_string(),
        "仿宋" | "FangSong" | "STFangsong" => "仿宋, FangSong, STFangsong, serif".to_string(),
        "Courier New" => "'Courier New', Courier, monospace".to_string(),
        "Times New Roman" => "'Times New Roman', Times, serif".to_string(),
        other => other.to_string(),
    };

    let font_size = text_obj.size;
    // Use OFD Weight attribute for bold detection (>= 700 = bold)
    let bold = if text_obj.weight >= 700 {
        " font-weight=\"bold\""
    } else {
        ""
    };

    // Build text content using absolute x positions (tspan x).
    // OFD DeltaX = absolute advance from char origin to next char origin (includes char width).
    // SVG tspan dx = ADDITIONAL offset on top of natural char advance — would double the spacing.
    // Solution: use tspan x with absolute positions in the text element's coordinate system.
    // base_x = the x position of the first character (set on <text> element).
    // Subsequent chars: tspan x = base_x + accumulated DeltaX.
    let chars: Vec<char> = text_obj.text.chars().collect();
    let has_delta = !text_obj.delta_x.is_empty() && chars.len() > 1;
    // We'll build the tspans later, after we know the base_x coordinate.
    // For now, just store the char data.

    // CTM transform: translate to boundary origin, apply matrix, then text at local coords
    if let Some(ctm) = text_obj.ctm {
        // CTM text: x is in local coords (text_x * scale)
        let base_x = text_obj.text_x * scale_x;
        let base_y = text_obj.text_y * scale_y;
        let content = if has_delta {
            let mut s = format!("<tspan x=\"{:.4}\">{}</tspan>", base_x, esc_xml(&chars[0].to_string()));
            let mut x_pos = base_x;
            for (i, ch) in chars.iter().enumerate().skip(1) {
                let dx = if i - 1 < text_obj.delta_x.len() {
                    text_obj.delta_x[i - 1]
                } else {
                    *text_obj.delta_x.last().unwrap_or(&font_size)
                };
                x_pos += dx * scale_x;
                s.push_str(&format!("<tspan x=\"{:.4}\">{}</tspan>", x_pos, esc_xml(&ch.to_string())));
            }
            s
        } else {
            esc_xml(&text_obj.text)
        };
        return format!(
            "<text transform=\"translate({bx},{by}) matrix({a},{b},{c},{d},{e},{f})\" x=\"{tx}\" y=\"{ty}\" font-family=\"{ff}\" font-size=\"{fs}\"{fc}{bw}>{ct}</text>",
            bx = text_obj.boundary.0 * scale_x,
            by = text_obj.boundary.1 * scale_y,
            a = ctm.0, b = ctm.1, c = ctm.2, d = ctm.3,
            e = ctm.4 * scale_x, f = ctm.5 * scale_y,
            tx = base_x,
            ty = base_y,
            ff = esc_xml_attr(&font_family),
            fs = font_size * scale_x,
            fc = fill_attr(text_obj.fill_color, text_obj.alpha),
            bw = bold,
            ct = content
        );
    }

    // Normal: position = Boundary + TextCode offset (absolute SVG coords)
    let base_x = (text_obj.boundary.0 + text_obj.text_x) * scale_x;
    let base_y = (text_obj.boundary.1 + text_obj.text_y) * scale_y;
    let content = if has_delta {
        let mut s = format!("<tspan x=\"{:.4}\">{}</tspan>", base_x, esc_xml(&chars[0].to_string()));
        let mut x_pos = base_x;
        for (i, ch) in chars.iter().enumerate().skip(1) {
            let dx = if i - 1 < text_obj.delta_x.len() {
                text_obj.delta_x[i - 1]
            } else {
                *text_obj.delta_x.last().unwrap_or(&font_size)
            };
            x_pos += dx * scale_x;
            s.push_str(&format!("<tspan x=\"{:.4}\">{}</tspan>", x_pos, esc_xml(&ch.to_string())));
        }
        s
    } else {
        esc_xml(&text_obj.text)
    };
    format!(
        "<text x=\"{x}\" y=\"{y}\" font-family=\"{ff}\" font-size=\"{fs}\"{fc}{bw}>{ct}</text>",
        x = base_x,
        y = base_y,
        ff = esc_xml_attr(&font_family),
        fs = font_size * scale_x,
        fc = fill_attr(text_obj.fill_color, text_obj.alpha),
        bw = bold,
        ct = content
    )
}

fn fill_attr(color: Option<(u8, u8, u8)>, alpha: Option<u8>) -> String {
    match (color, alpha) {
        (Some((r, g, b)), Some(a)) => format!(" fill=\"rgba({},{},{},{:.2})\"", r, g, b, a as f64 / 255.0),
        (Some((r, g, b)), None) => format!(" fill=\"rgb({},{},{})\"", r, g, b),
        (None, Some(a)) => format!(" fill=\"rgba(0,0,0,{:.2})\"", a as f64 / 255.0),
        (None, None) => String::new(),
    }
}

fn stroke_attr(color: Option<(u8, u8, u8)>, alpha: Option<u8>) -> String {
    match (color, alpha) {
        (Some((r, g, b)), Some(a)) => format!(" stroke=\"rgba({},{},{},{:.2})\"", r, g, b, a as f64 / 255.0),
        (Some((r, g, b)), None) => format!(" stroke=\"rgb({},{},{})\"", r, g, b),
        (None, Some(a)) => format!(" stroke=\"rgba(0,0,0,{:.2})\"", a as f64 / 255.0),
        (None, None) => String::new(),
    }
}

fn esc_xml(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

fn esc_xml_attr(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;").replace('\'', "&apos;")
}

/// Convert OFD AbbreviatedData to SVG path data.
/// OFD commands: M(moveto), L(lineto), C(cubic bezier), Q(quadratic), A(arc), B(cubic bezier alias), Z(close)
fn ofd_path_to_svg(data: &str) -> String {
    let mut svg = String::new();
    let tokens: Vec<&str> = data.split_whitespace().collect();
    let mut i = 0;
    while i < tokens.len() {
        match tokens[i] {
            "M" => {
                if i + 2 < tokens.len() {
                    svg.push_str(&format!("M {} {} ", tokens[i+1], tokens[i+2]));
                    i += 3;
                } else { i += 1; }
            }
            "L" => {
                if i + 2 < tokens.len() {
                    svg.push_str(&format!("L {} {} ", tokens[i+1], tokens[i+2]));
                    i += 3;
                } else { i += 1; }
            }
            "C" => {
                if i + 6 < tokens.len() {
                    svg.push_str(&format!("C {} {} {} {} {} {} ",
                        tokens[i+1], tokens[i+2], tokens[i+3], tokens[i+4], tokens[i+5], tokens[i+6]));
                    i += 7;
                } else { i += 1; }
            }
            "B" => {
                // OFD B is also cubic bezier (same as C)
                if i + 6 < tokens.len() {
                    svg.push_str(&format!("C {} {} {} {} {} {} ",
                        tokens[i+1], tokens[i+2], tokens[i+3], tokens[i+4], tokens[i+5], tokens[i+6]));
                    i += 7;
                } else { i += 1; }
            }
            "Q" => {
                if i + 4 < tokens.len() {
                    svg.push_str(&format!("Q {} {} {} {} ",
                        tokens[i+1], tokens[i+2], tokens[i+3], tokens[i+4]));
                    i += 5;
                } else { i += 1; }
            }
            "A" => {
                if i + 7 < tokens.len() {
                    svg.push_str(&format!("A {} {} {} {} {} {} {} ",
                        tokens[i+1], tokens[i+2], tokens[i+3], tokens[i+4], tokens[i+5], tokens[i+6], tokens[i+7]));
                    i += 8;
                } else { i += 1; }
            }
            "S" => {
                // Smooth cubic bezier
                if i + 4 < tokens.len() {
                    svg.push_str(&format!("S {} {} {} {} ",
                        tokens[i+1], tokens[i+2], tokens[i+3], tokens[i+4]));
                    i += 5;
                } else { i += 1; }
            }
            "Z" | "z" => {
                svg.push('Z');
                i += 1;
            }
            _ => { i += 1; }
        }
    }
    svg
}

/// Extract Layer DrawParam IDs from content XML.
/// OFD Layer has DrawParam="4" attribute pointing to a DrawParam in PublicRes.xml.
/// Returns all DrawParam IDs found on Layer elements.
fn extract_layer_draw_param_ids(xml: &str) -> Vec<u32> {
    use quick_xml::events::Event;
    use quick_xml::Reader;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut ids = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let tag = local_tag_name(&e.name());
                if tag == "Layer" {
                    if let Some(v) = attr_val(&e, "DrawParam") {
                        if let Ok(id) = v.parse::<u32>() {
                            ids.push(id);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    ids
}

/// Apply DrawParam defaults to paths and texts that have no explicit stroke/fill color.
fn apply_draw_param_defaults(
    paths: &mut [OfdPathObject],
    texts: &mut [OfdTextObject],
    draw_params: &std::collections::HashMap<u32, OfdDrawParam>,
    layer_dp_ids: &[u32],
) {
    // Resolve defaults from the first Layer DrawParam
    let (default_lw, default_stroke, default_fill) = if let Some(&dp_id) = layer_dp_ids.first() {
        resolve_draw_param(draw_params, dp_id)
    } else {
        return; // no DrawParam to inherit
    };

    for p in paths.iter_mut() {
        if p.stroke_color.is_none() {
            p.stroke_color = default_stroke;
        }
        if p.fill_color.is_none() {
            p.fill_color = default_fill;
        }
        if p.line_width == 0.0 {
            p.line_width = default_lw;
        }
    }

    for t in texts.iter_mut() {
        if t.fill_color.is_none() {
            t.fill_color = default_fill;
        }
        if t.stroke_color.is_none() {
            t.stroke_color = default_stroke;
        }
    }
}

/// Find the root DrawParam ID — the one that is referenced by others via Relative
/// but itself has no Relative. This serves as the global document default.
fn find_root_draw_param(draw_params: &std::collections::HashMap<u32, OfdDrawParam>) -> Option<u32> {
    // Find all IDs that are referenced as Relative targets
    let referenced: std::collections::HashSet<u32> = draw_params.values()
        .filter_map(|dp| dp.relative)
        .collect();
    // Root = referenced but has no Relative of its own
    for (&id, dp) in draw_params {
        if referenced.contains(&id) && dp.relative.is_none() {
            return Some(id);
        }
    }
    // Fallback: first DrawParam with no Relative
    for (&id, dp) in draw_params {
        if dp.relative.is_none() {
            return Some(id);
        }
    }
    None
}

/// Parse OFD content XML (Page or Template) and extract render objects.
/// Returns (text_objects, path_objects, image_objects)
fn parse_ofd_content(xml: &str) -> (Vec<OfdTextObject>, Vec<OfdPathObject>, Vec<OfdImageObject>) {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut text_objs = Vec::new();
    let mut path_objs = Vec::new();
    let mut img_objs = Vec::new();

    // We need to track context: which element we're in
    // TextObject, PathObject, ImageObject are direct children of Layer
    // TextCode is a child of TextObject
    // AbbreviatedData is a child of PathObject

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut current_text: Option<OfdTextObject> = None;
    let mut current_path: Option<OfdPathObject> = None;
    let mut current_img: Option<OfdImageObject> = None;
    let mut in_text_code = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag_local = local_tag_name(&e.name());
                match tag_local.as_str() {
                    "TextObject" => {
                        let mut t = OfdTextObject::default();
                        if let Some(v) = attr_val(&e, "ID") { t.id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "Boundary") {
                            if let Some(f4) = parse_f4(&v) { t.boundary = f4; }
                        }
                        if let Some(v) = attr_val(&e, "Font") { t.font_id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "Size") { t.size = v.parse().unwrap_or(3.175); }
                        if let Some(v) = attr_val(&e, "CTM") { t.ctm = parse_f6(&v); }
                        if let Some(v) = attr_val(&e, "Alpha") { t.alpha = v.parse().ok(); }
                        if let Some(v) = attr_val(&e, "BlendMode") { t.blend_mode = Some(v); }
                        if let Some(v) = attr_val(&e, "Weight") { t.weight = v.parse().unwrap_or(400); }
                        current_text = Some(t);
                    }
                    "PathObject" => {
                        let mut p = OfdPathObject::default();
                        if let Some(v) = attr_val(&e, "ID") { p.id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "Boundary") {
                            if let Some(f4) = parse_f4(&v) { p.boundary = f4; }
                        }
                        if let Some(v) = attr_val(&e, "LineWidth") { p.line_width = v.parse().unwrap_or(0.25); }
                        if let Some(v) = attr_val(&e, "Fill") { p.fill = v == "true"; }
                        if let Some(v) = attr_val(&e, "Alpha") { p.alpha = v.parse().ok(); }
                        current_path = Some(p);
                    }
                    "ImageObject" => {
                        let mut img = OfdImageObject::default();
                        if let Some(v) = attr_val(&e, "ID") { img.id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "Boundary") {
                            if let Some(f4) = parse_f4(&v) { img.boundary = f4; }
                        }
                        if let Some(v) = attr_val(&e, "ResourceID") { img.resource_id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "CTM") { img.ctm = parse_f6(&v); }
                        if let Some(v) = attr_val(&e, "BlendMode") { img.blend_mode = Some(v); }
                        if let Some(v) = attr_val(&e, "Alpha") { img.alpha = v.parse().ok(); }
                        current_img = Some(img);
                    }
                    "TextCode" => {
                        in_text_code = true;
                        if let Some(ref mut t) = current_text {
                            if let Some(v) = attr_val(&e, "X") { t.text_x = v.parse().unwrap_or(0.0); }
                            if let Some(v) = attr_val(&e, "Y") { t.text_y = v.parse().unwrap_or(0.0); }
                            if let Some(v) = attr_val(&e, "DeltaX") {
                                t.delta_x = parse_delta_x(&v);
                            }
                        }
                    }
                    "AbbreviatedData" => {
                        let text = read_element_text(&mut reader);
                        if let Some(ref mut p) = current_path {
                            p.abbreviated_data = text;
                        }
                        continue;
                    }
                    "StrokeColor" => {
                        if let Some(v) = attr_val(&e, "Value") {
                            if let Some(c) = parse_color(&v) {
                                if let Some(ref mut p) = current_path { p.stroke_color = Some(c); }
                                if let Some(ref mut t) = current_text { t.stroke_color = Some(c); }
                            }
                        }
                    }
                    "FillColor" => {
                        if let Some(v) = attr_val(&e, "Value") {
                            if let Some(c) = parse_color(&v) {
                                if let Some(ref mut p) = current_path { p.fill_color = Some(c); }
                                if let Some(ref mut t) = current_text { t.fill_color = Some(c); }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                // Self-closing elements like <ImageObject ... /> or <TextObject ... />
                let tag_local = local_tag_name(&e.name());
                match tag_local.as_str() {
                    "TextObject" => {
                        let mut t = OfdTextObject::default();
                        if let Some(v) = attr_val(&e, "ID") { t.id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "Boundary") {
                            if let Some(f4) = parse_f4(&v) { t.boundary = f4; }
                        }
                        if let Some(v) = attr_val(&e, "Font") { t.font_id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "Size") { t.size = v.parse().unwrap_or(3.175); }
                        if let Some(v) = attr_val(&e, "CTM") { t.ctm = parse_f6(&v); }
                        if let Some(v) = attr_val(&e, "Alpha") { t.alpha = v.parse().ok(); }
                        if let Some(v) = attr_val(&e, "Weight") { t.weight = v.parse().unwrap_or(400); }
                        text_objs.push(t);
                    }
                    "PathObject" => {
                        let mut p = OfdPathObject::default();
                        if let Some(v) = attr_val(&e, "ID") { p.id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "Boundary") {
                            if let Some(f4) = parse_f4(&v) { p.boundary = f4; }
                        }
                        if let Some(v) = attr_val(&e, "LineWidth") { p.line_width = v.parse().unwrap_or(0.25); }
                        if let Some(v) = attr_val(&e, "Fill") { p.fill = v == "true"; }
                        if let Some(v) = attr_val(&e, "Alpha") { p.alpha = v.parse().ok(); }
                        path_objs.push(p);
                    }
                    "ImageObject" => {
                        let mut img = OfdImageObject::default();
                        if let Some(v) = attr_val(&e, "ID") { img.id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "Boundary") {
                            if let Some(f4) = parse_f4(&v) { img.boundary = f4; }
                        }
                        if let Some(v) = attr_val(&e, "ResourceID") { img.resource_id = v.parse().unwrap_or(0); }
                        if let Some(v) = attr_val(&e, "CTM") { img.ctm = parse_f6(&v); }
                        if let Some(v) = attr_val(&e, "Alpha") { img.alpha = v.parse().ok(); }
                        img_objs.push(img);
                    }
                    "StrokeColor" => {
                        if let Some(v) = attr_val(&e, "Value") {
                            if let Some(c) = parse_color(&v) {
                                if let Some(ref mut p) = current_path { p.stroke_color = Some(c); }
                                if let Some(ref mut t) = current_text { t.stroke_color = Some(c); }
                            }
                        }
                    }
                    "FillColor" => {
                        if let Some(v) = attr_val(&e, "Value") {
                            if let Some(c) = parse_color(&v) {
                                if let Some(ref mut p) = current_path { p.fill_color = Some(c); }
                                if let Some(ref mut t) = current_text { t.fill_color = Some(c); }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(t)) => {
                if in_text_code {
                    if let Ok(s) = t.unescape() {
                        if let Some(ref mut text_obj) = current_text {
                            text_obj.text.push_str(&s);
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let tag_local = local_tag_name(&e.name());
                match tag_local.as_str() {
                    "TextObject" => {
                        if let Some(t) = current_text.take() {
                            text_objs.push(t);
                        }
                    }
                    "PathObject" => {
                        if let Some(p) = current_path.take() {
                            path_objs.push(p);
                        }
                    }
                    "ImageObject" => {
                        if let Some(img) = current_img.take() {
                            img_objs.push(img);
                        }
                    }
                    "TextCode" => {
                        in_text_code = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }

    (text_objs, path_objs, img_objs)
}

/// Get local tag name (strip namespace prefix)
fn local_tag_name(name: &quick_xml::name::QName) -> String {
    let bytes = name.as_ref();
    if let Some(pos) = bytes.iter().position(|&b| b == b':') {
        String::from_utf8_lossy(&bytes[pos + 1..]).to_string()
    } else {
        String::from_utf8_lossy(bytes).to_string()
    }
}

/// Parse OFD.xml CustomData entries for quick invoice data extraction
fn parse_ofd_custom_data(xml: &str) -> std::collections::HashMap<String, String> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut map = std::collections::HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                let tag = local_tag_name(&e.name());
                if tag == "CustomData" {
                    if let Some(name) = attr_val(&e, "Name") {
                        let value = read_element_text(&mut reader);
                        map.insert(name, value);
                        continue;
                    }
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

/// Parse Tags/CustomTag.xml — maps semantic field names to TextObject IDs
fn parse_custom_tag(xml: &str) -> std::collections::HashMap<String, Vec<u32>> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut map: std::collections::HashMap<String, Vec<u32>> = std::collections::HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_field = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag = local_tag_name(&e.name());
                match tag.as_str() {
                    "InvoiceNo" | "IssueDate" | "BuyerName" | "BuyerTaxID" |
                    "SellerName" | "SellerTaxID" | "TaxExclusiveTotalAmount" |
                    "TaxTotalAmount" | "TaxInclusiveTotalAmount" | "Amount" |
                    "TaxAmount" | "InvoiceClerk" | "Item" | "Price" | "Quantity" |
                    "Note" | "TaxScheme" | "MeasurementDimension" => {
                        current_field = tag;
                    }
                    "ObjectRef" => {
                        if !current_field.is_empty() {
                            // Read text content (the object ID)
                            let text = read_element_text(&mut reader);
                            if let Ok(id) = text.trim().parse::<u32>() {
                                map.entry(current_field.clone()).or_default().push(id);
                            }
                            continue;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let tag = local_tag_name(&e.name());
                match tag.as_str() {
                    "InvoiceNo" | "IssueDate" | "BuyerName" | "BuyerTaxID" |
                    "SellerName" | "SellerTaxID" | "TaxExclusiveTotalAmount" |
                    "TaxTotalAmount" | "TaxInclusiveTotalAmount" | "Amount" |
                    "TaxAmount" | "InvoiceClerk" | "Item" | "Price" | "Quantity" |
                    "Note" | "TaxScheme" | "MeasurementDimension" | "Buyer" | "Seller" => {
                        current_field.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

/// Parse PublicRes.xml for font definitions
fn parse_fonts(xml: &str) -> (std::collections::HashMap<u32, OfdFont>, std::collections::HashMap<u32, String>, std::collections::HashMap<u32, OfdDrawParam>) {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut fonts = std::collections::HashMap::new();
    let mut color_spaces = std::collections::HashMap::new();
    let mut draw_params = std::collections::HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_dp_id: Option<u32> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag = local_tag_name(&e.name());
                if tag == "Font" {
                    let mut font = OfdFont::default();
                    if let Some(v) = attr_val(&e, "ID") { font.id = v.parse().unwrap_or(0); }
                    if let Some(v) = attr_val(&e, "FontName") { font.font_name = v; }
                    if let Some(v) = attr_val(&e, "FamilyName") { font.family_name = v; }
                    fonts.insert(font.id, font);
                } else if tag == "ColorSpace" {
                    if let (Some(id_v), Some(type_v)) = (attr_val(&e, "ID"), attr_val(&e, "Type")) {
                        if let Ok(id) = id_v.parse::<u32>() {
                            color_spaces.insert(id, type_v);
                        }
                    }
                } else if tag == "DrawParam" {
                    let mut dp = OfdDrawParam::default();
                    if let Some(v) = attr_val(&e, "ID") { dp.id = v.parse().unwrap_or(0); }
                    if let Some(v) = attr_val(&e, "Relative") { dp.relative = v.parse().ok(); }
                    if let Some(v) = attr_val(&e, "LineWidth") { dp.line_width = v.parse().unwrap_or(0.25); }
                    current_dp_id = Some(dp.id);
                    draw_params.insert(dp.id, dp);
                } else if tag == "StrokeColor" {
                    if let Some(v) = attr_val(&e, "Value") {
                        if let Some(c) = parse_color(&v) {
                            if let Some(id) = current_dp_id {
                                if let Some(dp) = draw_params.get_mut(&id) {
                                    dp.stroke_color = Some(c);
                                }
                            }
                        }
                    }
                } else if tag == "FillColor" {
                    if let Some(v) = attr_val(&e, "Value") {
                        if let Some(c) = parse_color(&v) {
                            if let Some(id) = current_dp_id {
                                if let Some(dp) = draw_params.get_mut(&id) {
                                    dp.fill_color = Some(c);
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let tag = local_tag_name(&e.name());
                if tag == "DrawParam" { current_dp_id = None; }
            }
            Ok(Event::Empty(e)) => {
                let tag = local_tag_name(&e.name());
                if tag == "Font" {
                    let mut font = OfdFont::default();
                    if let Some(v) = attr_val(&e, "ID") { font.id = v.parse().unwrap_or(0); }
                    if let Some(v) = attr_val(&e, "FontName") { font.font_name = v; }
                    if let Some(v) = attr_val(&e, "FamilyName") { font.family_name = v; }
                    fonts.insert(font.id, font);
                } else if tag == "ColorSpace" {
                    if let (Some(id_v), Some(type_v)) = (attr_val(&e, "ID"), attr_val(&e, "Type")) {
                        if let Ok(id) = id_v.parse::<u32>() {
                            color_spaces.insert(id, type_v);
                        }
                    }
                } else if tag == "StrokeColor" {
                    // Self-closing: <ofd:StrokeColor Value="128 0 0" ColorSpace="2"/>
                    if let Some(v) = attr_val(&e, "Value") {
                        if let Some(c) = parse_color(&v) {
                            if let Some(id) = current_dp_id {
                                if let Some(dp) = draw_params.get_mut(&id) {
                                    dp.stroke_color = Some(c);
                                }
                            }
                        }
                    }
                } else if tag == "FillColor" {
                    if let Some(v) = attr_val(&e, "Value") {
                        if let Some(c) = parse_color(&v) {
                            if let Some(id) = current_dp_id {
                                if let Some(dp) = draw_params.get_mut(&id) {
                                    dp.fill_color = Some(c);
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    (fonts, color_spaces, draw_params)
}

/// Resolve DrawParam inheritance chain: returns fully resolved (line_width, stroke_color, fill_color)
fn resolve_draw_param(draw_params: &std::collections::HashMap<u32, OfdDrawParam>, param_id: u32) -> (f64, Option<(u8, u8, u8)>, Option<(u8, u8, u8)>) {
    let mut lw = 0.25f64;
    let mut stroke: Option<(u8, u8, u8)> = None;
    let mut fill: Option<(u8, u8, u8)> = None;
    let mut visited = std::collections::HashSet::new();
    let mut current_id = param_id;
    // Walk the Relative chain: 4 → 3 → None
    loop {
        if !visited.insert(current_id) { break; } // prevent cycles
        if let Some(dp) = draw_params.get(&current_id) {
            if dp.line_width > 0.0 { lw = dp.line_width; }
            if stroke.is_none() && dp.stroke_color.is_some() { stroke = dp.stroke_color; }
            if fill.is_none() && dp.fill_color.is_some() { fill = dp.fill_color; }
            if let Some(rel) = dp.relative {
                current_id = rel;
            } else {
                break;
            }
        } else {
            break;
        }
    }
    (lw, stroke, fill)
}

/// Parse DocumentRes.xml for image resources
fn parse_image_resources(xml: &str) -> std::collections::HashMap<u32, String> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut images = std::collections::HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_id: Option<u32> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let tag = local_tag_name(&e.name());
                if tag == "MultiMedia" {
                    if let Some(v) = attr_val(&e, "ID") {
                        current_id = v.parse().ok();
                    }
                } else if tag == "MediaFile" {
                    let text = read_element_text(&mut reader);
                    if let Some(id) = current_id.take() {
                        images.insert(id, text.trim().to_string());
                    }
                    continue;
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    images
}

/// Parse Annotations XML for watermark layer.
/// Each Annot contains an Appearance with a global Boundary.
/// Inner TextObject/ImageObject boundaries are relative to the Appearance.
/// This function adds the Appearance offset to convert to page-global coordinates.
fn parse_annotations(xml: &str) -> (Vec<OfdTextObject>, Vec<OfdImageObject>) {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut all_texts = Vec::new();
    let mut all_imgs = Vec::new();

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    // Track current Appearance offset (x, y) to apply to inner objects
    let mut appearance_offset: Option<(f64, f64)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag = local_tag_name(&e.name());
                match tag.as_str() {
                    "Appearance" => {
                        if let Some(v) = attr_val(&e, "Boundary") {
                            if let Some((x, y, _w, _h)) = parse_f4(&v) {
                                appearance_offset = Some((x, y));
                            }
                        }
                    }
                    "TextObject" | "ImageObject" => {
                        // We're inside an Appearance — parse the inner XML fragment
                        // by collecting until the matching End tag, then feed to parse_ofd_content
                        // Simpler approach: reconstruct a minimal Content XML with the object
                        let mut depth = 1u32;
                        let mut frag = format!("<ofd:Content><ofd:Layer>");
                        frag.push_str(&format!("<{} ", tag));
                        // Re-add attributes from the start element
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = std::str::from_utf8(&attr.value).unwrap_or("");
                            frag.push_str(&format!("{}=\"{}\" ", key, esc_xml_attr(val)));
                        }
                        frag.push('>');
                        // Read until matching End tag
                        loop {
                            let mut inner_buf = Vec::new();
                            match reader.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(inner_e)) => {
                                    depth += 1;
                                    let inner_tag = local_tag_name(&inner_e.name());
                                    frag.push_str(&format!("<{} ", inner_tag));
                                    for attr in inner_e.attributes().flatten() {
                                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                        let val = std::str::from_utf8(&attr.value).unwrap_or("");
                                        frag.push_str(&format!("{}=\"{}\" ", key, esc_xml_attr(val)));
                                    }
                                    frag.push('>');
                                }
                                Ok(Event::Empty(inner_e)) => {
                                    let inner_tag = local_tag_name(&inner_e.name());
                                    frag.push_str(&format!("<{} ", inner_tag));
                                    for attr in inner_e.attributes().flatten() {
                                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                        let val = std::str::from_utf8(&attr.value).unwrap_or("");
                                        frag.push_str(&format!("{}=\"{}\" ", key, esc_xml_attr(val)));
                                    }
                                    frag.push_str("/>");
                                }
                                Ok(Event::Text(t)) => {
                                    if let Ok(s) = t.unescape() {
                                        frag.push_str(&esc_xml(&s));
                                    }
                                }
                                Ok(Event::End(_inner_e)) => {
                                    depth -= 1;
                                    let inner_tag = local_tag_name(&_inner_e.name());
                                    frag.push_str(&format!("</{}>", inner_tag));
                                    if depth == 0 { break; }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        frag.push_str("</ofd:Layer></ofd:Content>");

                        let (mut texts, _, mut imgs) = parse_ofd_content(&frag);
                        // Apply Appearance offset to convert local → global coordinates
                        if let Some((ox, oy)) = appearance_offset {
                            for t in &mut texts {
                                t.boundary.0 += ox;
                                t.boundary.1 += oy;
                            }
                            for i in &mut imgs {
                                i.boundary.0 += ox;
                                i.boundary.1 += oy;
                            }
                        }
                        all_texts.extend(texts);
                        all_imgs.extend(imgs);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let tag = local_tag_name(&e.name());
                if tag == "Appearance" {
                    appearance_offset = None;
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }

    (all_texts, all_imgs)
}

/// Main OFD parser: opens ZIP, parses all XML, builds SVG, extracts invoice data
pub fn parse_ofd_file(ofd_path: &str) -> Result<OfdResult, String> {
    use base64::Engine;

    let file = std::fs::File::open(ofd_path)
        .map_err(|e| format!("打开OFD文件失败: {}", e))?;
    let mut archive = zip::ZipArchive::new(file)
        .map_err(|e| format!("解析OFD ZIP失败: {}", e))?;

    // 1. Read OFD.xml — root metadata + CustomData
    let ofd_xml = zip_read_str(&mut archive, "OFD.xml")
        .ok_or("OFD.xml 不存在")?;

    // Find DocRoot path (usually Doc_0/Document.xml)
    let doc_root = {
        use quick_xml::events::Event;
        use quick_xml::Reader;
        let mut rdr = Reader::from_str(&ofd_xml);
        rdr.config_mut().trim_text(true);
        let mut b = Vec::new();
        let mut root = String::from("Doc_0/Document.xml");
        loop {
            match rdr.read_event_into(&mut b) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if local_tag_name(&e.name()) == "DocRoot" {
                        let t = read_element_text(&mut rdr);
                        root = t.trim().to_string();
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            b.clear();
        }
        root
    };

    // Determine base directory from doc_root (e.g., "Doc_0/Document.xml" → "Doc_0")
    let base_dir = if let Some(pos) = doc_root.rfind('/') {
        doc_root[..pos].to_string()
    } else {
        String::from("Doc_0")
    };

    // 2. Parse CustomData from OFD.xml
    let custom_data = parse_ofd_custom_data(&ofd_xml);

    // 3. Read Document.xml to find template and page content paths
    let doc_xml = zip_read_str(&mut archive, &doc_root)
        .ok_or_else(|| format!("{} 不存在", doc_root))?;

    // Parse Document.xml to get template and page content paths
    let (template_path, page_paths) = {
        use quick_xml::events::Event;
        use quick_xml::Reader;
        let mut rdr = Reader::from_str(&doc_xml);
        rdr.config_mut().trim_text(true);
        let mut b = Vec::new();
        let mut tpl = String::new();
        let mut pages = Vec::new();
        loop {
            match rdr.read_event_into(&mut b) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let tag = local_tag_name(&e.name());
                    if tag == "TemplatePage" {
                        if let Some(v) = attr_val(&e, "BaseLoc") {
                            tpl = format!("{}/{}", base_dir, v);
                        }
                    } else if tag == "Page" {
                        if let Some(v) = attr_val(&e, "BaseLoc") {
                            pages.push(format!("{}/{}", base_dir, v));
                        }
                    }
                }
                Ok(Event::End(_)) => {}
                Ok(Event::Eof) => break,
                _ => {}
            }
            b.clear();
        }
        (tpl, pages)
    };

    // 4. Parse PublicRes.xml for fonts + DrawParam
    let public_res_path = format!("{}/PublicRes.xml", base_dir);
    let (font_map, color_spaces, draw_params) = if let Some(xml) = zip_read_str(&mut archive, &public_res_path) {
        parse_fonts(&xml)
    } else {
        (std::collections::HashMap::new(), std::collections::HashMap::new(), std::collections::HashMap::new())
    };

    // 5. Parse DocumentRes.xml for image resources
    let doc_res_path = format!("{}/DocumentRes.xml", base_dir);
    let image_map = if let Some(xml) = zip_read_str(&mut archive, &doc_res_path) {
        parse_image_resources(&xml)
    } else {
        std::collections::HashMap::new()
    };

    // Load actual image data from ZIP
    let mut image_data: std::collections::HashMap<u32, String> = std::collections::HashMap::new();
    for (res_id, file_name) in &image_map {
        let img_path = format!("{}/Res/{}", base_dir, file_name);
        if let Some(bytes) = zip_read_bytes(&mut archive, &img_path) {
            let b64 = base64::engine::general_purpose::STANDARD.encode(&bytes);
            let mime = if file_name.to_lowercase().ends_with(".png") { "image/png" } else { "image/jpeg" };
            image_data.insert(*res_id, format!("data:{};base64,{}", mime, b64));
        }
    }

    // Find the root DrawParam for global defaults (used when a layer has no DrawParam attribute)
    let root_dp_ids: Vec<u32> = find_root_draw_param(&draw_params).into_iter().collect();

    // 6. Parse template content (background layer)
    let (tpl_texts, tpl_paths, tpl_imgs) = if !template_path.is_empty() {
        if let Some(xml) = zip_read_str(&mut archive, &template_path) {
            let layer_dp_ids = extract_layer_draw_param_ids(&xml);
            let (mut t, mut p, i) = parse_ofd_content(&xml);
            apply_draw_param_defaults(&mut p, &mut t, &draw_params, &layer_dp_ids);
            (t, p, i)
        } else {
            (Vec::new(), Vec::new(), Vec::new())
        }
    } else {
        (Vec::new(), Vec::new(), Vec::new())
    };

    // 7. Parse page content (data layer)
    // Note: avoid shadowing `page_paths` (Vec<String> from Document.xml parsing)
    let (mut page_texts, mut page_obj_paths, page_imgs) = if let Some(page_path) = page_paths.first() {
        if let Some(xml) = zip_read_str(&mut archive, page_path) {
            let layer_dp_ids = extract_layer_draw_param_ids(&xml);
            let (mut t, mut p, i) = parse_ofd_content(&xml);
            apply_draw_param_defaults(&mut p, &mut t, &draw_params, &layer_dp_ids);
            (t, p, i)
        } else {
            (Vec::new(), Vec::new(), Vec::new())
        }
    } else {
        (Vec::new(), Vec::new(), Vec::new())
    };
    // Page content Layer may not have DrawParam attribute — apply root DrawParam as global fallback
    apply_draw_param_defaults(&mut page_obj_paths, &mut page_texts, &draw_params, &root_dp_ids);

    // 8. Parse annotations (watermark layer) — uses parse_annotations to handle Appearance offsets
    let annots_path = format!("{}/Annots/Page_0/Annotation.xml", base_dir);
    let (annot_texts, annot_imgs) = if let Some(xml) = zip_read_str(&mut archive, &annots_path) {
        parse_annotations(&xml)
    } else {
        (Vec::new(), Vec::new())
    };

    // 9. Get page dimensions
    let (page_w, page_h) = if let Some(page_path) = page_paths.first() {
        if let Some(xml) = zip_read_str(&mut archive, page_path) {
            // Parse PhysicalBox from the page XML
            use quick_xml::events::Event;
            use quick_xml::Reader;
            let mut rdr = Reader::from_str(&xml);
            rdr.config_mut().trim_text(true);
            let mut b = Vec::new();
            let mut dims = (210.0f64, 140.0f64);
            loop {
                match rdr.read_event_into(&mut b) {
                    Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                        if local_tag_name(&e.name()) == "PhysicalBox" {
                            let text = read_element_text(&mut rdr);
                            if let Some((_, _, w, h)) = parse_f4(text.trim()) {
                                dims = (w, h);
                            }
                            break;
                        }
                    }
                    Ok(Event::Eof) => break,
                    _ => {}
                }
                b.clear();
            }
            dims
        } else {
            (210.0, 140.0)
        }
    } else {
        (210.0, 140.0)
    };

    // 10. Parse CustomTag.xml for semantic field mapping
    let custom_tag_path = format!("{}/Tags/CustomTag.xml", base_dir);
    let tag_map = if let Some(xml) = zip_read_str(&mut archive, &custom_tag_path) {
        parse_custom_tag(&xml)
    } else {
        std::collections::HashMap::new()
    };

    // 11. Extract invoice info from structured data
    let mut invoice_info = OfdInvoiceInfo::default();

    // From OFD.xml CustomData
    invoice_info.invoice_no = custom_data.get("发票号码").cloned();
    invoice_info.invoice_date = custom_data.get("开票日期").cloned();
    invoice_info.buyer_tax_id = custom_data.get("购买方纳税人识别号").cloned();
    invoice_info.seller_tax_id = custom_data.get("销售方纳税人识别号").cloned();
    invoice_info.amount_no_tax = custom_data.get("合计金额").and_then(|s| s.parse().ok());
    invoice_info.tax_amount = custom_data.get("合计税额").and_then(|s| s.parse().ok());

    // Compute total = no_tax + tax (both already in yuan, e.g. 17699.12 + 2300.88 = 20000.00)
    if let (Some(no_tax), Some(tax)) = (invoice_info.amount_no_tax, invoice_info.tax_amount) {
        invoice_info.amount_tax = Some(((no_tax + tax) * 100.0).round() / 100.0);
    }

    // From CustomTag.xml + Content.xml — get buyer/seller names
    // Build a text lookup: TextObject ID → text content
    let mut text_lookup: std::collections::HashMap<u32, &str> = std::collections::HashMap::new();
    for t in &page_texts {
        text_lookup.insert(t.id, &t.text);
    }

    // Map tag fields to text content
    let get_tag_text = |field: &str| -> Option<String> {
        tag_map.get(field).and_then(|ids| {
            ids.iter().filter_map(|id| text_lookup.get(id)).map(|s| s.to_string()).collect::<Vec<_>>().into_iter().next()
        })
    };

    if invoice_info.invoice_no.is_none() {
        invoice_info.invoice_no = get_tag_text("InvoiceNo");
    }
    if invoice_info.invoice_date.is_none() {
        invoice_info.invoice_date = get_tag_text("IssueDate");
    }
    if invoice_info.buyer_name.is_none() {
        invoice_info.buyer_name = get_tag_text("BuyerName");
    }
    if invoice_info.seller_name.is_none() {
        invoice_info.seller_name = get_tag_text("SellerName");
    }

    // Detect invoice type from template title
    for t in &tpl_texts {
        if t.text.contains("增值税专用") {
            invoice_info.invoice_type = Some("增值税专用发票".to_string());
            break;
        } else if t.text.contains("增值税普通") || t.text.contains("增值税电子普通") {
            invoice_info.invoice_type = Some("增值税普通发票".to_string());
            break;
        } else if t.text.contains("电子发票") {
            invoice_info.invoice_type = Some("电子发票".to_string());
            break;
        }
    }

    // 12. Build SVG
    let svg = build_ofd_svg(
        page_w, page_h,
        &tpl_texts, &tpl_paths, &tpl_imgs,
        &page_texts, &page_obj_paths, &page_imgs,
        &annot_texts, &annot_imgs,
        &font_map, &color_spaces, &image_data,
    );

    log::info!("OFD parsed: {}x{}mm, {} template texts, {} page texts, {} paths",
        page_w, page_h, tpl_texts.len(), page_texts.len(), tpl_paths.len() + page_obj_paths.len());

    Ok(OfdResult {
        svg,
        invoice_info,
        page_width: page_w,
        page_height: page_h,
    })
}

/// Build complete SVG from parsed OFD layers
fn build_ofd_svg(
    page_w: f64,
    page_h: f64,
    tpl_texts: &[OfdTextObject],
    tpl_paths: &[OfdPathObject],
    tpl_imgs: &[OfdImageObject],
    page_texts: &[OfdTextObject],
    page_paths: &[OfdPathObject],
    page_imgs: &[OfdImageObject],
    annot_texts: &[OfdTextObject],
    annot_imgs: &[OfdImageObject],
    font_map: &std::collections::HashMap<u32, OfdFont>,
    color_spaces: &std::collections::HashMap<u32, String>,
    image_data: &std::collections::HashMap<u32, String>,
) -> String {
    let scale = 3.5; // Scale factor: 1mm → 3.5 SVG units for good resolution
    let vw = page_w * scale;
    let vh = page_h * scale;

    let mut svg = format!(
        "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" viewBox=\"0 0 {:.1} {:.1}\" width=\"{:.1}\" height=\"{:.1}\" style=\"background:white\">",
        vw, vh, vw, vh
    );

    // Layer 1: Template (background) — grid lines and static labels
    svg.push_str("<g id=\"template\">");
    for p in tpl_paths {
        svg.push_str(&build_svg_path(p, scale));
    }
    for t in tpl_texts {
        svg.push_str(&build_svg_text(t, font_map, color_spaces, scale, scale));
    }
    for img in tpl_imgs {
        svg.push_str(&build_svg_image(img, image_data, scale));
    }
    svg.push_str("</g>");

    // Layer 2: Content (data)
    svg.push_str("<g id=\"content\">");
    for p in page_paths {
        svg.push_str(&build_svg_path(p, scale));
    }
    for t in page_texts {
        svg.push_str(&build_svg_text(t, font_map, color_spaces, scale, scale));
    }
    for img in page_imgs {
        svg.push_str(&build_svg_image(img, image_data, scale));
    }
    svg.push_str("</g>");

    // Layer 3: Annotations (watermarks)
    svg.push_str("<g id=\"annotations\">");
    for t in annot_texts {
        svg.push_str(&build_svg_text(t, font_map, color_spaces, scale, scale));
    }
    for img in annot_imgs {
        svg.push_str(&build_svg_image(img, image_data, scale));
    }
    svg.push_str("</g>");

    svg.push_str("</svg>");
    svg
}

/// Build SVG path from OFD PathObject
fn build_svg_path(p: &OfdPathObject, scale: f64) -> String {
    if p.abbreviated_data.is_empty() {
        return String::new();
    }

    let svg_d = ofd_path_to_svg(&p.abbreviated_data);
    if svg_d.is_empty() {
        return String::new();
    }

    // Boundary = (x, y, w, h) in mm. Path data is in local coords within Boundary.
    // Apply translate to Boundary position, then scale everything.
    let tx = p.boundary.0 * scale;
    let ty = p.boundary.1 * scale;

    let mut attrs = String::new();
    attrs.push_str(&format!(" transform=\"translate({:.4},{:.4}) scale({:.4})\"", tx, ty, scale));
    attrs.push_str(&format!(" stroke-width=\"{:.4}\"", p.line_width));
    if p.fill {
        attrs.push_str(" fill-rule=\"nonzero\"");
    }
    attrs.push_str(&stroke_attr(p.stroke_color, p.alpha));
    if p.fill {
        if let Some(fc) = p.fill_color {
            attrs.push_str(&fill_attr(Some(fc), p.alpha));
        } else {
            // fill=true but no explicit fill_color: don't fall back to stroke_color
            // (that would fill the shape solid and hide internal strokes like the ¥ cross)
            attrs.push_str(" fill=\"none\"");
        }
    } else {
        attrs.push_str(" fill=\"none\"");
    }

    format!("<g{}><path d=\"{}\"/></g>", attrs, svg_d)
}

/// Build SVG image from OFD ImageObject
fn build_svg_image(img: &OfdImageObject, image_data: &std::collections::HashMap<u32, String>, scale: f64) -> String {
    let data_url = match image_data.get(&img.resource_id) {
        Some(url) => url,
        None => return String::new(),
    };

    // Boundary = (x, y, w, h) in mm — already defines where and how big the image should be.
    // Do NOT apply CTM for images: in OFD, CTM often describes the pixel-to-mm mapping
    // (e.g. QR 300px image with CTM [20 0 0 20 ...] means 300px → 20mm),
    // but the Boundary already encodes the target display size.
    // Applying CTM as SVG transform would incorrectly scale the image again.
    let x = img.boundary.0 * scale;
    let y = img.boundary.1 * scale;
    let w = img.boundary.2 * scale;
    let h = img.boundary.3 * scale;

    let opacity = img.alpha.map(|a| format!(" opacity=\"{:.2}\"", a as f64 / 255.0)).unwrap_or_default();

    format!(
        "<image href=\"{}\" x=\"{:.4}\" y=\"{:.4}\" width=\"{:.4}\" height=\"{:.4}\"{}/>",
        data_url, x, y, w, h, opacity
    )
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
/// operations are needed (no trim, no color mode change, rotation 0°/180° can
/// use PDF matrix), the raw JPEG bytes are preserved in ImageSource::JpegPassthrough.
/// This skips the decode→re-encode pipeline entirely, giving zero quality loss
/// and smaller file sizes.
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
            let can_passthrough = is_jpeg_bytes(&bytes)
                && !trim
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
/// For Decoded images: rotates, converts to RawImage, registers via add_image.
/// For JpegPassthrough images with 0°/180° rotation: embeds raw JPEG bytes
/// as DCTDecode stream via ExternalXObject (zero quality loss, no re-encode).
/// For JpegPassthrough with 90°/270°: falls back to decode-rotate-reencode.
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
                // Dimensions: for 0° use (w, h), for 180° the image dims stay the same
                // (rotation handled by PDF transform matrix, not pixel manipulation)
                let iw_mm = *width as f32 * 25.4 / RENDER_DPI as f32;
                let ih_mm = *height as f32 * 25.4 / RENDER_DPI as f32;

                // Determine ColorSpace from JPEG component count
                let color_space: &[u8] = match num_components {
                    1 => b"DeviceGray",
                    4 => b"DeviceCMYK",
                    _ => b"DeviceRGB", // 3 components (default for most JPEGs)
                };

                // Build ExternalXObject with DCTDecode filter — raw JPEG bytes embedded directly
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
                        compress: false, // JPEG is already compressed — no zlib on top!
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

        let rotation = slot_spec.rotation;
        let cached = match get_cached_xobj(doc, xobj_cache, file_idx, rotation, sources) {
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
) -> Result<(), String> {
    if request.pages.is_empty() {
        return Err("没有页面数据".to_string());
    }

    // PDF passthrough: if all files are PDF pages with no pixel-level modifications,
    // use Form XObject approach to preserve vector quality (zero quality loss,
    // ~95% smaller files, ~10x faster). Falls back to current pipeline on any error.
    if can_passthrough_pdf(request) {
        match generate_pdf_passthrough(request, output_path, on_progress.as_ref()) {
            Ok(()) => return Ok(()),
            Err(e) => {
                log::warn!("PDF直通失败，回退渲染管道: {}", e);
                // Continue with current pipeline below
            }
        }
    }

    let total_pages = request.pages.len() as u32;
    let (slot_positions, pw, ph) = calculate_layout_mm(&request.settings);

    // Create PDF document (new API: no page dimensions at creation time)
    let mut doc = printpdf::PdfDocument::new("发票打印");

    // Step 1: Decode all unique images (base64 → DynamicImage), apply trim + color mode.
    // Rotation is per-slot and deferred to build_page_ops for correct (file, rotation) caching.
    let total_files = request.files.len() as u32;
    if let Some(ref cb) = &on_progress {
        cb("decode", 0, total_files);
    }
    let sources = decode_images(&request.files, &request.settings);
    if let Some(ref cb) = &on_progress {
        cb("decode", total_files, total_files);
    }

    // Step 2: Build pages, caching XObjects by (file_index, rotation) to avoid redundant work.
    let mut xobj_cache: std::collections::HashMap<(usize, i32), CachedXobj> = std::collections::HashMap::new();

    for (i, page_spec) in request.pages.iter().enumerate() {
        // Check shutdown flag — abort PDF generation if app is closing
        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，PDF生成已中止".to_string());
        }
        let ops = build_page_ops(
            &mut doc,
            page_spec,
            &request.settings,
            &slot_positions,
            &sources,
            &mut xobj_cache,
        );

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

    // Save PDF — custom options for print quality
    // Default PdfSaveOptions limits images to 2MB and uses 0.85 JPEG quality,
    // which downsamples high-res invoice images and makes text blurry.
    // We override: no size limit, higher quality, prefer JPEG for speed.
    //
    // **Optimization**: When content includes PDF/OFD pages (rendered text),
    // use FlateDecode (lossless) instead of JPEG (lossy) for sharper text.
    // Photo images still benefit from JPEG compression.
    let has_text_content = request.files.iter().any(|f| {
        f.source_type.as_deref() == Some("pdf-page") || f.source_type.as_deref() == Some("ofd-page")
    });
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

    Ok(())
}

// =====================================================
// PDF Passthrough — Form XObject based vector-preserving pipeline
// =====================================================

/// Check if the layout request can use PDF page passthrough.
/// Conditions: all files are PDF pages, no trim, no color mode change.
/// Rotation IS supported (handled via PDF transformation matrix).
fn can_passthrough_pdf(request: &LayoutRenderRequest) -> bool {
    let s = &request.settings;
    if s.trim_white.unwrap_or(false) { return false; }
    if s.color_mode != "color" && !s.color_mode.is_empty() { return false; }

    // All slots must reference files with pdf_path
    for page in &request.pages {
        for slot in &page.slots {
            if let Some(idx) = slot.file_index {
                if idx >= request.files.len() { return false; }
                let file = &request.files[idx];
                if file.pdf_path.is_none() { return false; }
                if file.pdf_page_idx.is_none() { return false; }
            }
        }
    }
    true
}

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
                Ok(((x1, y1, x2, y2), (x2 - x1, y2 - y1)))
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
    output_doc: &mut lopdf::Document,
    id_map: &mut std::collections::HashMap<lopdf::ObjectId, lopdf::ObjectId>,
) -> Result<(lopdf::ObjectId, f32, f32), String> {
    // 1. Get page content stream bytes (decompressed and concatenated)
    let content_bytes = source.get_page_content(page_id)
        .map_err(|e| format!("提取内容流失败: {}", e))?;

    // 2. Get effective visible box — CropBox takes precedence over MediaBox.
    // This fixes the bug where PDFs with CropBox (e.g., only showing the bottom
    // half of the page) would appear shrunk, because we were using MediaBox
    // dimensions for scaling but the visible content was only a portion.
    let ((box_x1, box_y1, _box_x2, _box_y2), (page_w_pt, page_h_pt)) =
        get_page_effective_box(source, page_id)?;

    // If CropBox has a non-zero origin, we need to translate the content stream
    // so that the visible area aligns with the BBox origin (0, 0).
    // Prepend: "1 0 0 1 -box_x1 -box_y1 cm" to shift content.
    let final_content = if box_x1.abs() > 0.01 || box_y1.abs() > 0.01 {
        let translate = format!("1 0 0 1 {:.4} {:.4} cm\n", -box_x1, -box_y1);
        let mut combined = translate.into_bytes();
        combined.extend_from_slice(&content_bytes);
        combined
    } else {
        content_bytes
    };

    // 3. Get page resources — merge ALL resource dictionaries including inherited ones.
    //
    // CRITICAL: lopdf's get_page_resources returns:
    //   - resources_opt: the page's OWN Resources (only if inline Dictionary, NOT Reference!)
    //   - ref_ids: ObjectIds of ALL Resources dicts found by walking up the page tree
    //
    // Most PDFs store Resources as indirect References (e.g., "Resources 5 0 R"),
    // so resources_opt is often None. We MUST also process ref_ids to include
    // inherited resources (fonts, colorspaces, etc. from parent Pages nodes).
    // Without this, Form XObjects have no resources → blank pages.
    let (resources_opt, ref_ids) = source.get_page_resources(page_id)
        .map_err(|e| format!("提取资源失败: {}", e))?;

    // 4. Merge all resource dictionaries: page's own + all inherited from parents.
    // We build a merged dictionary, then deep-copy + remap the whole thing.
    let mut merged = lopdf::Dictionary::new();

    // Add inherited resources FIRST (parent-level, so page-level overrides take precedence)
    for rid in &ref_ids {
        if let Ok(res_dict) = source.get_dictionary(*rid) {
            merge_resource_dict(&mut merged, res_dict, source);
        }
    }

    // Add page's own resources LAST (overrides inherited ones with same keys)
    if let Some(dict) = resources_opt {
        merge_resource_dict(&mut merged, dict, source);
    }

    let remapped_resources = {
        let obj = lopdf::Object::Dictionary(merged);
        remap_references(obj, source, output_doc, id_map)
    };

    // 5. Build Form XObject stream
    // BBox uses the effective visible dimensions (CropBox or MediaBox).
    // For CropBox=[0, 433.75, 595, 842], this becomes [0, 0, 595, 408.25].
    let mut dict = lopdf::Dictionary::new();
    dict.set("Type", lopdf::Object::Name(b"XObject".to_vec()));
    dict.set("Subtype", lopdf::Object::Name(b"Form".to_vec()));
    dict.set("FormType", lopdf::Object::Integer(1));
    dict.set("BBox", lopdf::Object::Array(vec![
        lopdf::Object::Real(0.0),
        lopdf::Object::Real(0.0),
        lopdf::Object::Real(page_w_pt),
        lopdf::Object::Real(page_h_pt),
    ]));
    dict.set("Resources", remapped_resources);

    // Transparency group — ensures correct rendering of overlapping content
    let mut group_dict = lopdf::Dictionary::new();
    group_dict.set("Type", lopdf::Object::Name(b"Group".to_vec()));
    group_dict.set("S", lopdf::Object::Name(b"Transparency".to_vec()));
    dict.set("Group", lopdf::Object::Dictionary(group_dict));

    let stream = lopdf::Stream::new(dict, final_content).with_compression(true);
    let xobj_id = output_doc.add_object(lopdf::Object::Stream(stream));

    Ok((xobj_id, page_w_pt, page_h_pt))
}

/// Build the content stream for one output page using cm + Do operators.
/// Each Form XObject is positioned, scaled, and rotated within its layout slot.
fn build_nup_content_stream(
    form_xobjs: &[(lopdf::ObjectId, f32, f32)],  // (form_xobj_id, src_w_pt, src_h_pt)
    slot_positions: &[LayoutSlotMm],
    settings: &RenderSettings,
    slot_rotations: &[i32],  // per-slot rotation degrees
) -> Result<Vec<u8>, String> {
    use lopdf::content::Operation;

    let mut ops = Vec::new();

    for (slot_idx, (_xobj_id, src_w_pt, src_h_pt)) in form_xobjs.iter().enumerate() {
        if slot_idx >= slot_positions.len() { break; }
        let slot = &slot_positions[slot_idx];
        let slot_w_pt = slot.w_mm * MM_TO_PT;
        let slot_h_pt = slot.h_mm * MM_TO_PT;

        // Handle rotation via transformation matrix
        let rotation = if slot_idx < slot_rotations.len() { slot_rotations[slot_idx] } else { 0 };
        let rot = ((rotation % 360) + 360) % 360;

        // For 90°/270° rotation, the visual dimensions swap (width↔height),
        // so scaling must be computed against the *rotated* dimensions to fit the slot correctly.
        let (vis_w, vis_h) = if rot == 90 || rot == 270 {
            (*src_h_pt, *src_w_pt) // rotated: visual width = original height, etc.
        } else {
            (*src_w_pt, *src_h_pt)
        };

        // Compute scale to fit in slot based on visual (rotated) dimensions
        let (scale_x, scale_y) = match settings.fit_mode.as_str() {
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

        // Centered position in slot (bottom-left origin) based on visual dimensions
        let draw_w = vis_w * scale_x;
        let draw_h = vis_h * scale_y;
        let offset_x = slot.x_mm * MM_TO_PT + (slot_w_pt - draw_w) / 2.0;
        let offset_y = slot.y_mm * MM_TO_PT + (slot_h_pt - draw_h) / 2.0;

        // PDF transformation matrix: [a b c d e f]
        // Maps Form XObject coordinate space (0,0)-(src_w,src_h) to page area.
        // For rotation, we derive the matrix from the desired mapping:
        //   rot=0:   (x,y) → (sx*x+ox, sy*y+oy)
        //   rot=90:  (x,y) → (sx*(src_h-y)+ox, sy*x+oy)  [CCW in PDF = CW visually]
        //   rot=180: (x,y) → (sx*(src_w-x)+ox, sy*(src_h-y)+oy)
        //   rot=270: (x,y) → (sx*y+ox, sy*(src_w-x)+oy)   [CW in PDF = CCW visually]
        // Where sx/sy scale from source to visual dimensions:
        let (sx, sy) = if rot == 90 || rot == 270 {
            // After rotation: visual width = src_h, visual height = src_w
            (draw_w / *src_h_pt, draw_h / *src_w_pt)
        } else {
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
        let xobj_name = lopdf::Object::Name(format!("Fm{}", slot_idx).into_bytes());

        ops.push(Operation { operator: "q".into(), operands: vec![] });
        ops.push(Operation { operator: "cm".into(), operands: matrix });
        ops.push(Operation { operator: "Do".into(), operands: vec![xobj_name] });
        ops.push(Operation { operator: "Q".into(), operands: vec![] });
    }

    let content = lopdf::content::Content { operations: ops };
    content.encode().map_err(|e| format!("内容流编码失败: {}", e))
}

/// Generate PDF using Form XObject passthrough — preserves vector content,
/// fonts, and text exactly. Supports all layouts (1×1, 2×1, 3×3, etc.)
/// and rotation via PDF transformation matrices.
fn generate_pdf_passthrough(
    request: &LayoutRenderRequest,
    output_path: &std::path::Path,
    on_progress: Option<&ProgressFn>,
) -> Result<(), String> {
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

    for (page_idx, page_spec) in request.pages.iter().enumerate() {
        if SHUTTING_DOWN.load(Ordering::SeqCst) {
            return Err("应用正在关闭，PDF生成已中止".to_string());
        }

        // Collect Form XObjects for each slot in this page
        let mut page_form_xobjs: Vec<(lopdf::ObjectId, f32, f32)> = Vec::new();
        let mut slot_rotations: Vec<i32> = Vec::new();
        let mut xobj_names: Vec<(std::vec::Vec<u8>, lopdf::ObjectId)> = Vec::new();

        for (slot_idx, slot) in page_spec.slots.iter().enumerate() {
            let file_idx = match slot.file_index {
                Some(idx) if idx < request.files.len() => idx,
                _ => continue,
            };
            let file = &request.files[file_idx];
            let pdf_path = match file.pdf_path.as_ref() {
                Some(p) => p.clone(),
                None => continue,
            };
            let page_idx_in_pdf = match file.pdf_page_idx {
                Some(idx) => idx,
                None => continue,
            };

            // Load source PDF (cached)
            if !source_cache.contains_key(&pdf_path) {
                let source = lopdf::Document::load(&pdf_path)
                    .map_err(|e| format!("加载源PDF失败 {}: {}", pdf_path, e))?;
                source_cache.insert(pdf_path.clone(), source);
                global_id_maps.insert(pdf_path.clone(), std::collections::HashMap::new());
            }
            let source = source_cache.get_mut(&pdf_path).unwrap();
            let id_map = global_id_maps.get_mut(&pdf_path).unwrap();

            // Find the source page ObjectId
            // lopdf get_pages() returns BTreeMap<u32, ObjectId> where u32 is 1-based page number
            let pages = source.get_pages();
            let source_page_id = pages.get(&(page_idx_in_pdf + 1))  // lopdf uses 1-based page numbers
                .copied()
                .ok_or_else(|| format!("PDF页面{}不存在 (文件: {})", page_idx_in_pdf + 1, pdf_path))?;

            // Extract as Form XObject
            let (xobj_id, src_w_pt, src_h_pt) = extract_page_as_form_xobject(
                source, source_page_id, &mut output_doc, id_map
            )?;

            let xobj_name = format!("Fm{}", slot_idx);
            xobj_names.push((xobj_name.into_bytes(), xobj_id));
            page_form_xobjs.push((xobj_id, src_w_pt, src_h_pt));
            slot_rotations.push(slot.rotation);
        }

        if page_form_xobjs.is_empty() {
            continue; // Empty page, skip
        }

        // Build content stream for this output page
        let content_bytes = build_nup_content_stream(
            &page_form_xobjs, &slot_positions, &request.settings, &slot_rotations
        )?;

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
    std::fs::write(output_path, &pdf_buf)
        .map_err(|e| format!("写入文件失败: {}", e))?;
    if let Some(ref cb) = &on_progress {
        cb("save", 1, 1);
    }

    Ok(())
}
