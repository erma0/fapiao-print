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

        // Only accept supported formats
        if !["pdf", "jpg", "jpeg", "png", "bmp", "webp", "tiff", "tif"].contains(&ext.as_str()) {
            continue;
        }

        let metadata = path.metadata().map_err(|e| format!("读取文件信息失败: {}", e))?;
        let size = metadata.len();

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
    // Assume 96 DPI for screen images, convert to mm
    let img_w_mm = img.width() as f32 * 25.4 / 96.0;
    let img_h_mm = img.height() as f32 * 25.4 / 96.0;

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
            dpi: Some(96.0),
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
