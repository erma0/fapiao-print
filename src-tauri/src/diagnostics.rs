//! 诊断导出（issue #47 配套）——用户一键导出「脱敏」的文件结构与识别摘要，
//! 经 GitHub issue 回传后即可在本地精确复现渲染/识别问题，无需再靠截图猜测。
//!
//! 硬约束：报告会公开传递，任何用户原文（金额/名称/号码/文件名/路径）都不得出现——
//! 文本一律经 `invoice_engine::sanitize_text` 脱敏（汉字→汉 / 数字→9 / 字母→A / 空格→·），
//! 文件名同样脱敏且只保留文件名（不含目录），系统元数据（版本/尺寸/坐标）原样输出。

use serde::Deserialize;

/// 前端传入的诊断项。`summary` 仅允许包含布尔/枚举/计数等固定词汇
/// （由前端 `buildDiagSummary` 保证，不含用户内容），因此不参与脱敏——
/// 否则「专票 / OFD / 有 / 无」这类枚举词会被误替换。
#[derive(Deserialize)]
pub struct DiagItem {
    pub path: String,
    #[serde(default)]
    pub summary: String,
}

#[tauri::command]
pub async fn export_diagnostics(files: Vec<DiagItem>, exported_at: String) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || build_report(&files, &exported_at))
        .await
        .map_err(|e| format!("任务执行失败: {}", e))
}

/// 文件名脱敏：只保留文件名（不含目录），主体文本脱敏，扩展名原样保留。
fn sanitize_file_name(path: &str) -> String {
    let p = std::path::Path::new(path);
    let stem = p.file_stem().map(|s| s.to_string_lossy().to_string()).unwrap_or_default();
    let ext = p.extension().map(|s| s.to_string_lossy().to_string()).unwrap_or_default();
    let stem_s = invoice_engine::sanitize_text(&stem);
    if ext.is_empty() { stem_s } else { format!("{}.{}", stem_s, ext) }
}

fn build_report(files: &[DiagItem], exported_at: &str) -> String {
    let mut r = String::new();
    r.push_str("发票酱 诊断报告\n");
    r.push_str("================\n");
    r.push_str(&format!(
        "版本: {} ({})\n",
        env!("CARGO_PKG_VERSION"),
        if cfg!(feature = "ocr") { "OCR版" } else { "轻量版" }
    ));
    r.push_str(&format!("系统: {}\n", std::env::consts::OS));
    r.push_str(&format!("导出时间: {}\n", exported_at));
    r.push_str("\n说明: 报告已脱敏（汉字→汉 / 数字→9 / 字母→A / 空格→·），不含任何原文内容；\n");
    r.push_str("      数字、标点与坐标结构完整保留，用于定位渲染/识别问题。\n");

    for (i, f) in files.iter().enumerate() {
        r.push_str(&format!("\n【文件 #{}】{}\n", i + 1, sanitize_file_name(&f.path)));
        match std::fs::metadata(&f.path) {
            Ok(m) => r.push_str(&format!("大小: {:.1} KB\n", m.len() as f64 / 1024.0)),
            Err(e) => {
                r.push_str(&format!("（无法读取文件: {}）\n", e));
                continue;
            }
        }
        if !f.summary.is_empty() {
            r.push_str(&format!("识别摘要: {}\n", f.summary));
        }
        let ext = std::path::Path::new(&f.path)
            .extension()
            .map(|e| e.to_string_lossy().to_lowercase())
            .unwrap_or_default();
        let part = match ext.as_str() {
            "ofd" => invoice_engine::dump_ofd_structure(&f.path),
            "pdf" => dump_pdf(&f.path),
            "xml" => dump_xml(&f.path),
            "jpg" | "jpeg" | "png" | "bmp" => dump_image(&f.path),
            _ => Ok("（未知类型，仅输出摘要）\n".to_string()),
        };
        match part {
            Ok(t) => r.push_str(&t),
            Err(e) => r.push_str(&format!("（结构提取失败: {}）\n", e)),
        }
    }

    // 体积兜底：报告要能贴进 issue，限制 512 KB（按 char 边界截断，避免切坏多字节字符）
    const MAX: usize = 512 * 1024;
    if r.len() > MAX {
        let mut cut = MAX;
        while cut > 0 && !r.is_char_boundary(cut) {
            cut -= 1;
        }
        r.truncate(cut);
        r.push_str("\n…（报告超过 512 KB 已截断）\n");
    }
    r
}

/// PDF：页数 + 文字层提取结果（前 3 页、每页最多 300 个词的脱敏文本与坐标）
/// 词样本按行顺序输出，300 词可覆盖标题、购销方信息块与备注/页脚区
fn dump_pdf(path: &str) -> Result<String, String> {
    let doc = lopdf::Document::load(path).map_err(|e| format!("PDF 加载失败: {}", e))?;
    let total = doc.get_pages().len() as u32;
    let idxs: Vec<u32> = (0..total.min(3)).collect();
    let mut r = format!("总页数: {}\n", total);
    match crate::pdf_engine::extract_pdf_texts(path, &idxs) {
        Ok(map) => {
            for idx in &idxs {
                if let Some(res) = map.get(idx) {
                    r.push_str(&format!(
                        "[第{}页] 文字层={} 文本{}字 行{}\n",
                        idx,
                        if res.has_text_layer { "有" } else { "无" },
                        res.text.chars().count(),
                        res.lines.len()
                    ));
                    let mut n = 0usize;
                    'outer: for line in &res.lines {
                        for w in &line.words {
                            if n >= 300 {
                                break 'outer;
                            }
                            r.push_str(&format!(
                                "  word \"{}\" x={:.1} y={:.1} w={:.1} h={:.1}\n",
                                invoice_engine::sanitize_text(&w.text),
                                w.x, w.y, w.w, w.h
                            ));
                            n += 1;
                        }
                    }
                    if n >= 300 {
                        r.push_str("  …样本截断（每页最多 300 词）\n");
                    }
                }
            }
            if total > 3 {
                r.push_str(&format!("（仅分析前 3 页，共 {} 页）\n", total));
            }
        }
        Err(e) => r.push_str(&format!("文字层提取失败: {}\n", e)),
    }
    Ok(r)
}

/// XML 数电票：解析出的字段清单（值全部脱敏——只看「解析到没有」与值形态）
fn dump_xml(path: &str) -> Result<String, String> {
    let info = invoice_engine::parse_xml_invoice(path)?;
    let s = |o: &Option<String>| {
        o.as_deref()
            .map(invoice_engine::sanitize_text)
            .unwrap_or_else(|| "（无）".to_string())
    };
    let n = |o: &Option<f64>| {
        o.map(|v| invoice_engine::sanitize_text(&format!("{:.2}", v)))
            .unwrap_or_else(|| "（无）".to_string())
    };
    let mut r = String::from("解析字段:\n");
    r.push_str(&format!("  发票号码: {}\n", s(&info.invoice_no)));
    r.push_str(&format!("  开票日期: {}\n", s(&info.invoice_date)));
    r.push_str(&format!("  票种: {}\n", s(&info.invoice_type)));
    r.push_str(&format!("  销方名称: {}\n", s(&info.seller_name)));
    r.push_str(&format!("  销方税号: {}\n", s(&info.seller_tax_id)));
    r.push_str(&format!("  购方名称: {}\n", s(&info.buyer_name)));
    r.push_str(&format!("  购方税号: {}\n", s(&info.buyer_tax_id)));
    r.push_str(&format!("  不含税金额: {}\n", n(&info.amount_no_tax)));
    r.push_str(&format!("  税额: {}\n", n(&info.tax_amount)));
    r.push_str(&format!("  含税金额: {}\n", n(&info.amount_tax)));
    r.push_str(&format!(
        "  通行费标记: {}\n",
        info.is_toll.map(|b| b.to_string()).unwrap_or_else(|| "（无）".to_string())
    ));
    r.push_str(&format!("  项目名称: {}\n", s(&info.item_name)));
    r.push_str(&format!("  发票备注: {}\n", s(&info.remark)));
    Ok(r)
}

/// 图片：像素尺寸（EXIF/DPI 等信息量有限，先不采集）
fn dump_image(path: &str) -> Result<String, String> {
    match image::image_dimensions(path) {
        Ok((w, h)) => Ok(format!("图像尺寸: {}x{}\n", w, h)),
        Err(e) => Ok(format!("（图像读取失败: {}）\n", e)),
    }
}