use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::BufWriter;

/// Rendering DPI — must match frontend PDF_RENDER_DPI constant
pub const RENDER_DPI: u32 = 300;

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
    use windows::core::HSTRING;
    use windows::Data::Pdf::{PdfDocument, PdfPageRenderOptions};
    use windows::Storage::StorageFile;
    use windows::Storage::Streams::{DataReader, InMemoryRandomAccessStream};
    use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
    use base64::Engine;

    // Initialize COM for this thread
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

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
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x08000000;

    let output = std::process::Command::new("powershell")
        .args(["-NoProfile", "-NonInteractive", "-Command",
            "Get-Printer | Select-Object Name, Default | ConvertTo-Json"])
        .creation_flags(CREATE_NO_WINDOW)
        .output()
        .map_err(|e| format!("获取打印机列表失败: {}", e))?;

    let stdout = String::from_utf8_lossy(&output.stdout);
    if stdout.trim().is_empty() {
        return Ok(vec![]);
    }

    #[derive(Deserialize)]
    #[allow(non_snake_case)]
    struct PsPrinter {
        Name: String,
        #[serde(default)]
        Default: Option<bool>,
    }

    let printers: Vec<PsPrinter> = if stdout.trim().starts_with('[') {
        serde_json::from_str(&stdout).unwrap_or_default()
    } else {
        serde_json::from_str::<PsPrinter>(&stdout).map(|p| vec![p]).unwrap_or_default()
    };

    Ok(printers.into_iter().map(|p| PrinterInfo {
        name: p.Name,
        is_default: p.Default.unwrap_or(false),
    }).collect())
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
    use windows::Media::Ocr::OcrEngine;
    use windows::Graphics::Imaging::BitmapDecoder;
    use windows::Storage::Streams::InMemoryRandomAccessStream;
    use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
    use base64::Engine;

    unsafe { let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED); }

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

        // Determine MIME type and extension
        // MIME uses actual image type (needed for Image loading), ext shows "ofd" for display
        let lower = entry_name.to_lowercase();
        let mime = if lower.ends_with(".png") {
            "image/png"
        } else {
            "image/jpeg"
        };
        let img_ext = "ofd";

        let b64 = base64::engine::general_purpose::STANDARD.encode(&data);
        let data_url = format!("data:{};base64,{}", mime, b64);
        results.push((data_url, img_ext.to_string()));
    }

    log::info!("OFD extracted {} images from {}", results.len(), ofd_path);
    Ok(results)
}
