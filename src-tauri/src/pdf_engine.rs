use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::BufWriter;

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
        let size = page.Size().map_err(|e| format!("获取第{}页尺寸失败: {}", i + 1, e))?;
        let scale = dpi as f32 / 96.0;
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
        });

        log::info!("Rendered page {} ({}x{})", i + 1, dest_w, dest_h);
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
                    for (idx, img_data_url) in images.iter().enumerate() {
                        results.push(FileData {
                            name: if images.len() > 1 {
                                format!("{}_第{}页.ofd", name.trim_end_matches(".ofd").trim_end_matches(".OFD"), idx + 1)
                            } else {
                                name.clone()
                            },
                            ext: "png".to_string(), // extracted as PNG
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
    // Frontend renders at 300 DPI, convert to mm
    let img_w_mm = img.width() as f32 * 25.4 / 300.0;
    let img_h_mm = img.height() as f32 * 25.4 / 300.0;

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
            dpi: Some(300.0),
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
// OFD Format Support
// =====================================================

/// Extract embedded images from an OFD file (Chinese electronic invoice format)
/// OFD is a ZIP archive containing XML page descriptions and image resources.
/// For electronic invoices, the content is typically a full-page image.
fn extract_ofd_images(ofd_path: &str) -> Result<Vec<String>, String> {
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

    // Sort entries to get consistent ordering (page 0, page 1, etc.)
    image_entries.sort();

    // Read and encode each image
    let mut results = Vec::new();
    for entry_name in &image_entries {
        let mut entry = archive.by_name(entry_name)
            .map_err(|e| format!("读取OFD图片失败: {}", e))?;
        let mut data = Vec::new();
        entry.read_to_end(&mut data)
            .map_err(|e| format!("读取OFD图片数据失败: {}", e))?;

        // Determine MIME type
        let lower = entry_name.to_lowercase();
        let mime = if lower.ends_with(".png") {
            "image/png"
        } else {
            "image/jpeg"
        };

        let b64 = base64::engine::general_purpose::STANDARD.encode(&data);
        let data_url = format!("data:{};base64,{}", mime, b64);
        results.push(data_url);
    }

    log::info!("OFD extracted {} images from {}", results.len(), ofd_path);
    Ok(results)
}
