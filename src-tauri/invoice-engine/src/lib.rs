//! OFD (Open Fixed-layout Document) parser and SVG renderer.
//!
//! Supports Chinese electronic invoices (发票): extracts structured invoice data
//! from OFD XML metadata (CustomData + CustomTag) and renders pages as SVG.
//!
//! Also supports standalone XML 数电票 (fully digitalized e-invoice) parsing:
//! extracts structured invoice fields from `<EInvoice>` XML files.
//!
//! The OFD format is a ZIP archive containing XML page descriptions and image resources,
//! defined by Chinese national standard GB/T 33190-2016.

use image::GenericImageView;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

// =====================================================
// Public Types
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

/// Result returned by `parse_ofd_file`: SVG rendering + structured invoice data.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfdResult {
    pub svg: String,
    pub invoice_info: OfdInvoiceInfo,
    pub page_width: f64,
    pub page_height: f64,
}

/// An image extracted from an OFD file (for bitmap fallback path).
#[derive(Debug, Clone)]
pub struct OfdExtractedImage {
    pub data_url: String,
    pub ext: String,
    pub width: u32,
    pub height: u32,
}

/// Invoice data extracted from standalone XML 数电票 file.
/// XML 数电票 is a structured data format (no layout info), used for archiving and data exchange.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct XmlInvoiceInfo {
    pub invoice_no: Option<String>,
    pub invoice_date: Option<String>,
    pub seller_name: Option<String>,
    pub seller_tax_id: Option<String>,
    pub buyer_name: Option<String>,
    pub buyer_tax_id: Option<String>,
    pub amount_no_tax: Option<f64>,
    pub tax_amount: Option<f64>,
    pub amount_tax: Option<f64>,
    /// Invoice type label (e.g. "增值税专用发票", "电子发票(普通发票)")
    pub invoice_type: Option<String>,
}

// =====================================================
// Internal OFD Structures
// =====================================================

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
    layer_draw_param: Option<u32>, // DrawParam ID from the Layer this object belongs to
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
            layer_draw_param: None,
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
    layer_draw_param: Option<u32>, // DrawParam ID from the Layer this object belongs to
}

#[derive(Debug, Default)]
struct OfdImageObject {
    id: u32,
    boundary: (f64, f64, f64, f64),
    resource_id: u32,
    ctm: Option<(f64, f64, f64, f64, f64, f64)>,
    blend_mode: Option<String>,
    alpha: Option<u8>,
    image_mask: Option<u32>, // ResourceID of mask image (OFD ImageMask attribute)
}

// =====================================================
// ZIP Helpers
// =====================================================

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

// =====================================================
// Parsing Helpers
// =====================================================

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

// =====================================================
// SVG Generation Helpers
// =====================================================

/// Normalize a font name that may contain a subset prefix.
/// Subset font names follow the pattern: `PREFIX+BaseFontName-PREFIX+BaseFontName-Suffix`
/// (e.g., `AEWMEC+KaiTi-AEWMEC+KaiTi-0` → base name `KaiTi`)
/// Also handles PostScript font names like `CourierNewPSMT` → `Courier New`.
fn normalize_font_name(raw: &str) -> String {
    // Step 1: Extract base font name from subset prefix pattern
    // Subset prefix format: UPPERCASE_LETTERS+BaseName
    let base = if let Some(plus_pos) = raw.find('+') {
        // Found subset prefix — extract text after '+' up to next '-' or end
        let after_plus = &raw[plus_pos + 1..];
        let end = after_plus.find('-').unwrap_or(after_plus.len());
        &after_plus[..end]
    } else {
        raw
    };

    // Step 2: Map PostScript font names to standard CSS font-family names
    match base {
        "CourierNewPSMT" => "Courier New",
        "TimesNewRomanPSMT" => "Times New Roman",
        "ArialMT" => "Arial",
        "Arial-BoldMT" => "Arial",
        "SimSun" | "STSong" => "宋体",
        "KaiTi" | "STKaiti" => "楷体",
        "SimHei" | "STHeiti" => "黑体",
        "FangSong" | "STFangsong" => "仿宋",
        other => other,
    }.to_string()
}

/// Build SVG text element from an OFD TextObject
fn build_svg_text(
    text_obj: &OfdTextObject,
    font_map: &HashMap<u32, OfdFont>,
    _color_spaces: &HashMap<u32, String>,
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

    // Normalize subset font names (e.g., "AEWMEC+KaiTi-AEWMEC+KaiTi-0" → "KaiTi")
    let font_base = normalize_font_name(&font_family_raw);

    // Font fallback: add generic CJK/serif/sans-serif fallbacks for cross-platform rendering.
    // SVG font-family is CSS: names with spaces need single quotes (attr value is in double quotes).
    let font_family = match font_base.as_str() {
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

// =====================================================
// OFD Content Parsing
// =====================================================

/// Apply DrawParam defaults to paths and texts that have no explicit stroke/fill color.
/// Each object carries its own `layer_draw_param` from the Layer it belongs to,
/// so we apply per-object DrawParam inheritance rather than a single global default.
fn apply_draw_param_defaults(
    paths: &mut [OfdPathObject],
    texts: &mut [OfdTextObject],
    draw_params: &HashMap<u32, OfdDrawParam>,
) {
    // Cache resolved DrawParam results to avoid re-resolving the same ID
    let mut dp_cache: HashMap<u32, (f64, Option<(u8, u8, u8)>, Option<(u8, u8, u8)>)> = HashMap::new();

    for p in paths.iter_mut() {
        if let Some(dp_id) = p.layer_draw_param {
            let (lw, stroke, fill) = *dp_cache.entry(dp_id).or_insert_with(|| resolve_draw_param(draw_params, dp_id));
            if p.stroke_color.is_none() {
                p.stroke_color = stroke;
            }
            if p.fill_color.is_none() {
                p.fill_color = fill;
            }
            if p.line_width == 0.0 {
                p.line_width = lw;
            }
        }
        // Objects without layer_draw_param use OFD default (black stroke, no fill) — no inheritance
    }

    for t in texts.iter_mut() {
        if let Some(dp_id) = t.layer_draw_param {
            let (_lw, stroke, fill) = *dp_cache.entry(dp_id).or_insert_with(|| resolve_draw_param(draw_params, dp_id));
            if t.fill_color.is_none() {
                t.fill_color = fill;
            }
            if t.stroke_color.is_none() {
                t.stroke_color = stroke;
            }
        }
    }
}

/// Parse OFD content XML (Page or Template) and extract render objects.
/// Returns (text_objects, path_objects, image_objects)
/// Each object records its Layer's DrawParam ID in `layer_draw_param` for per-Layer inheritance.
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
    let mut current_layer_dp: Option<u32> = None; // DrawParam of the current Layer

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag_local = local_tag_name(&e.name());
                match tag_local.as_str() {
                    "Layer" => {
                        // Track this Layer's DrawParam
                        if let Some(v) = attr_val(&e, "DrawParam") {
                            current_layer_dp = v.parse().ok();
                        } else {
                            current_layer_dp = None;
                        }
                    }
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
                        t.layer_draw_param = current_layer_dp;
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
                        p.layer_draw_param = current_layer_dp;
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
                        if let Some(v) = attr_val(&e, "ImageMask") { img.image_mask = v.parse().ok(); }
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
                    "Layer" => {
                        // Self-closing Layer: just update current_layer_dp
                        if let Some(v) = attr_val(&e, "DrawParam") {
                            current_layer_dp = v.parse().ok();
                        } else {
                            current_layer_dp = None;
                        }
                    }
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
                        t.layer_draw_param = current_layer_dp;
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
                        p.layer_draw_param = current_layer_dp;
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
                        if let Some(v) = attr_val(&e, "ImageMask") { img.image_mask = v.parse().ok(); }
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
                    "Layer" => {
                        // Exiting Layer: reset to no DrawParam
                        current_layer_dp = None;
                    }
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

// =====================================================
// OFD Metadata Parsing
// =====================================================

/// Parse OFD.xml CustomData entries for quick invoice data extraction
fn parse_ofd_custom_data(xml: &str) -> HashMap<String, String> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut map = HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                // Self-closing tag: <CustomData Name="xxx"/> — value is empty, do NOT call read_element_text
                let tag = local_tag_name(&e.name());
                if tag == "CustomData" {
                    if let Some(name) = attr_val(&e, "Name") {
                        map.insert(name, String::new());
                    }
                }
            }
            Ok(Event::Start(e)) => {
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
fn parse_custom_tag(xml: &str) -> HashMap<String, Vec<u32>> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut map: HashMap<String, Vec<u32>> = HashMap::new();
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
fn parse_fonts(xml: &str) -> (HashMap<u32, OfdFont>, HashMap<u32, String>, HashMap<u32, OfdDrawParam>) {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut fonts = HashMap::new();
    let mut color_spaces = HashMap::new();
    let mut draw_params = HashMap::new();
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
fn resolve_draw_param(draw_params: &HashMap<u32, OfdDrawParam>, param_id: u32) -> (f64, Option<(u8, u8, u8)>, Option<(u8, u8, u8)>) {
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
fn parse_image_resources(xml: &str) -> HashMap<u32, String> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut images = HashMap::new();
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

// =====================================================
// SVG Assembly
// =====================================================

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
    font_map: &HashMap<u32, OfdFont>,
    color_spaces: &HashMap<u32, String>,
    image_data: &HashMap<u32, String>,
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
    // Per OFD spec, default stroke color is black (0,0,0) when not specified.
    // This ensures PathObjects without explicit StrokeColor (and no DrawParam inheritance)
    // are still visible — e.g. the ⊗ symbol (circled-X) in the uppercase amount area.
    attrs.push_str(&stroke_attr(p.stroke_color.or(Some((0, 0, 0))), p.alpha));
    if p.fill {
        if let Some(fc) = p.fill_color {
            attrs.push_str(&fill_attr(Some(fc), p.alpha));
        } else {
            // fill=true but no explicit fill_color: per OFD spec default is black,
            // but filling solid would hide internal strokes (e.g. the ⊗ cross).
            // Use fill="none" so the circle outline + X cross are both visible via stroke.
            attrs.push_str(" fill=\"none\"");
        }
    } else {
        attrs.push_str(" fill=\"none\"");
    }

    format!("<g{}><path d=\"{}\"/></g>", attrs, svg_d)
}

/// Build SVG image from OFD ImageObject
fn build_svg_image(img: &OfdImageObject, image_data: &HashMap<u32, String>, scale: f64) -> String {
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
// Bitmap Fallback: Extract Images from OFD ZIP
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

// =====================================================
// Text-based Invoice Extraction (Fallback)
// =====================================================

/// Extract invoice data from text content when no CustomData or CustomTag is available.
/// This handles OFD files from non-standard producers that embed subset fonts
/// but don't include structured XML metadata.
///
/// Strategy: scan text objects in order, detect label patterns, and extract
/// values from the same text (after "：") or the next text object.
fn extract_invoice_from_text(texts: &[&OfdTextObject]) -> OfdInvoiceInfo {
    let mut info = OfdInvoiceInfo::default();
    let mut section = ""; // "buyer" or "seller"
    let mut name_count = 0; // 1st "名称" = buyer, 2nd = seller
    let mut taxid_count = 0; // 1st "纳税人识别号" = buyer, 2nd = seller
    let mut found_jiashui_label = false; // "价税合计" marker (may be separate from "小写")
    let mut found_xiaoxie_label = false; // "（小写）" after "价税合计"
    let mut found_heji_label = false; // "合计" (non-价税合计) marker

    // Pre-concatenate adjacent single-character texts (some OFDs split CJK labels into
    // individual characters, e.g., "购""买""方""信""息" instead of "购买方信息").
    // We accumulate into a buffer and flush when we see a multi-char text or a label.
    let mut char_buf = String::new();
    let flush_buf = |buf: &mut String| -> Option<String> {
        if buf.len() >= 2 {
            let s = buf.trim().to_string();
            buf.clear();
            Some(s)
        } else {
            buf.clear();
            None
        }
    };

    // We process texts in a two-pass approach:
    // Pass 1: Concatenate single-char sequences into composite labels
    // Pass 2: Apply pattern matching on the composite sequence

    let mut composite_texts: Vec<String> = Vec::new();
    for t in texts {
        let text = t.text.trim();
        if text.is_empty() { continue; }
        let chars: Vec<char> = text.chars().collect();
        if chars.len() == 1 {
            char_buf.push(chars[0]);
        } else {
            // Flush accumulated single chars first
            if let Some(composite) = flush_buf(&mut char_buf) {
                composite_texts.push(composite);
            }
            composite_texts.push(text.to_string());
        }
    }
    // Flush remaining
    if let Some(composite) = flush_buf(&mut char_buf) {
        composite_texts.push(composite);
    }

    // Pass 2: Pattern matching on composite text sequence
    for (i, text) in composite_texts.iter().enumerate() {
        let t = text.as_str();
        if t.is_empty() { continue; }

        // Remove spaces for flexible matching (e.g., "合        计" → "合计")
        let t_nospace: String = t.chars().filter(|c| !c.is_whitespace()).collect();
        let t_nospace_ref = t_nospace.as_str();

        // Detect buyer/seller section boundaries
        if t_nospace_ref.contains("购买方") || t_nospace_ref.contains("买方") {
            section = "buyer";
        }
        if t_nospace_ref.contains("销售方") || t_nospace_ref.contains("卖方") {
            section = "seller";
        }

        // Invoice number
        if (t.contains("发票号码") || t_nospace_ref.contains("发票号码")) && info.invoice_no.is_none() {
            info.invoice_no = extract_composite_value(t, &composite_texts, i, "taxid");
        }

        // Invoice date
        if (t.contains("开票日期") || t_nospace_ref.contains("开票日期")) && info.invoice_date.is_none() {
            info.invoice_date = extract_composite_value(t, &composite_texts, i, "any");
        }

        // Name label — 1st occurrence = buyer, 2nd = seller
        // (Some OFDs have "名称：" as a standalone label before each section's value)
        if (t.contains("名称") || t_nospace_ref.contains("名称"))
            && !t.contains("货物") && !t.contains("劳务") && !t.contains("项目") {
            name_count += 1;
            let value = extract_composite_value(t, &composite_texts, i, "name");
            // If section is still unknown, use occurrence count
            let effective_section = if section.is_empty() {
                if name_count == 1 { "buyer" } else { "seller" }
            } else { section };
            match effective_section {
                "buyer" if info.buyer_name.is_none() => info.buyer_name = value,
                "seller" if info.seller_name.is_none() => info.seller_name = value,
                _ => {}
            }
        }

        // Tax ID — 1st occurrence = buyer, 2nd = seller
        if (t.contains("纳税人识别号") || t.contains("统一社会信用代码"))
            && !t.contains("货物") {
            taxid_count += 1;
            let value = extract_composite_value(t, &composite_texts, i, "taxid");
            let effective_section = if section.is_empty() {
                if taxid_count == 1 { "buyer" } else { "seller" }
            } else { section };
            match effective_section {
                "buyer" if info.buyer_tax_id.is_none() => info.buyer_tax_id = value,
                "seller" if info.seller_tax_id.is_none() => info.seller_tax_id = value,
                _ => {}
            }
        }

        // Amount detection
        // "价税合计" label — may be followed by separate "（小写）" label
        if t_nospace_ref.contains("价税合计") {
            if t_nospace_ref.contains("小写") {
                found_xiaoxie_label = true;
            } else {
                found_jiashui_label = true;
            }
        }
        // "（小写）" or "小写" after "价税合计"
        if (t.contains("小写") || t_nospace_ref.contains("小写")) && found_jiashui_label {
            found_xiaoxie_label = true;
        }
        // "合计" label (not "价税合计") — handle spaced variants like "合        计"
        if (t_nospace_ref.contains("合计") || t_nospace_ref == "合计")
            && !t_nospace_ref.contains("价税") {
            found_heji_label = true;
        }

        // ¥ amount values
        if t.starts_with("¥") || t.starts_with("￥") {
            let amt_str = t.trim_start_matches('¥').trim_start_matches('￥').trim();
            if let Ok(amt) = amt_str.parse::<f64>() {
                if found_xiaoxie_label {
                    // This ¥ is after "价税合计（小写）" → total amount with tax
                    if info.amount_tax.is_none() {
                        info.amount_tax = Some(amt);
                    }
                    found_xiaoxie_label = false;
                    found_jiashui_label = false;
                } else if found_heji_label {
                    // This ¥ is after "合计" → subtotal (no tax or with tax)
                    if info.amount_no_tax.is_none() {
                        info.amount_no_tax = Some(amt);
                    }
                    found_heji_label = false;
                }
            }
        }

        // Invoice type detection
        if info.invoice_type.is_none() {
            if t.contains("增值税专用") {
                info.invoice_type = Some("增值税专用发票".to_string());
            } else if t.contains("增值税普通") || t.contains("增值税电子普通") {
                info.invoice_type = Some("增值税普通发票".to_string());
            } else if t.contains("电子发票") {
                info.invoice_type = Some("电子发票".to_string());
            }
        }
    }

    // Compute missing amount fields
    // If we have amount_tax but no breakdown, assume no_tax = amount_tax and tax = 0
    if info.amount_tax.is_some() && info.amount_no_tax.is_none() {
        info.amount_no_tax = info.amount_tax;
        info.tax_amount = Some(0.0);
    }
    // If we have amount_no_tax but no amount_tax, try to compute
    if info.amount_no_tax.is_some() && info.amount_tax.is_none() {
        if let Some(tax) = info.tax_amount {
            info.amount_tax = Some(((info.amount_no_tax.unwrap() + tax) * 100.0).round() / 100.0);
        } else {
            // No tax info → assume amount_no_tax IS the total (tax exempt)
            info.amount_tax = info.amount_no_tax;
            info.tax_amount = Some(0.0);
        }
    }

    info
}

/// Extract a value from a label text or the next text in the composite sequence.
/// First tries to get value after "：" or ":" in the same text.
/// If not found, looks at the next 1-3 texts for a non-label value.
/// `value_kind` hints at the expected format: "name" (CJK chars), "taxid" (alphanumeric), or "any".
fn extract_composite_value(label_text: &str, texts: &[String], label_idx: usize, value_kind: &str) -> Option<String> {
    // Try extracting from same text after colon
    for sep in &["：", ":"] {
        if let Some(pos) = label_text.find(sep) {
            let after = label_text[pos + sep.len()..].trim();
            if !after.is_empty() && value_matches(after, value_kind) {
                return Some(after.to_string());
            }
        }
    }

    // Look at next texts for the value
    for j in (label_idx + 1)..std::cmp::min(label_idx + 5, texts.len()) {
        let next_text = texts[j].trim();
        if !next_text.is_empty() && !is_common_label(next_text) && value_matches(next_text, value_kind) {
            return Some(next_text.to_string());
        }
    }

    None
}

/// Check if a candidate value matches the expected format kind.
fn value_matches(text: &str, kind: &str) -> bool {
    match kind {
        "taxid" => {
            // Tax IDs are alphanumeric (digits + possible X/x suffix), no CJK characters
            text.chars().all(|c| c.is_ascii_alphanumeric())
        }
        "name" => {
            // Names should contain at least one CJK character or be a known format
            text.chars().any(|c| c > '\u{2E80}') // CJK and other East Asian chars
        }
        _ => true,
    }
}

/// Check if a text looks like a common invoice label (not a value)
fn is_common_label(text: &str) -> bool {
    let labels = [
        "发票号码", "开票日期", "名称", "纳税人识别号", "统一社会信用代码",
        "地址", "电话", "开户行", "账号", "购买方", "销售方",
        "价税合计", "合计", "备注", "开票人", "收款人", "复核人",
        "货物", "劳务", "规格型号", "单位", "数量", "单价", "金额",
        "税率", "税额", "项目名称", "小写", "大写",
    ];
    labels.iter().any(|l| text.contains(l))
        || text.ends_with("：") || text.ends_with(":")
}

// =====================================================
// Public API
// =====================================================

/// Parse OFD file: returns SVG vector rendering + structured invoice data from XML.
/// Skips OCR — invoice fields are extracted directly from OFD metadata.
///
/// This is the primary entry point for OFD processing. It:
/// 1. Opens the OFD as a ZIP archive
/// 2. Parses OFD.xml for CustomData (quick invoice fields)
/// 3. Parses Document.xml for page/template structure
/// 4. Parses PublicRes.xml for fonts and DrawParam inheritance
/// 5. Parses DocumentRes.xml for image resources
/// 6. Parses page content (TextObject/PathObject/ImageObject)
/// 7. Parses annotations (watermark layer with Appearance offset handling)
/// 8. Maps CustomTag.xml fields to TextObject IDs for buyer/seller names
/// 9. Generates SVG with 3 layers: template + content + annotations
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
                        root = t.trim().trim_start_matches('/').to_string();
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
        (HashMap::new(), HashMap::new(), HashMap::new())
    };

    // 5. Parse DocumentRes.xml for image resources
    let doc_res_path = format!("{}/DocumentRes.xml", base_dir);
    let image_map = if let Some(xml) = zip_read_str(&mut archive, &doc_res_path) {
        parse_image_resources(&xml)
    } else {
        HashMap::new()
    };

    // Load actual image raw bytes from ZIP (data URL generation deferred until after content parsing
    // so we can apply ImageMask from parsed ImageObjects)
    let mut image_raw_bytes: HashMap<u32, Vec<u8>> = HashMap::new();
    let mut image_file_names: HashMap<u32, String> = HashMap::new();
    for (res_id, file_name) in &image_map {
        let img_path = format!("{}/Res/{}", base_dir, file_name);
        if let Some(bytes) = zip_read_bytes(&mut archive, &img_path) {
            image_raw_bytes.insert(*res_id, bytes);
            image_file_names.insert(*res_id, file_name.clone());
        }
    }

    // 6. Parse template content (background layer)
    let (tpl_texts, tpl_paths, tpl_imgs) = if !template_path.is_empty() {
        if let Some(xml) = zip_read_str(&mut archive, &template_path) {
            let (mut t, mut p, i) = parse_ofd_content(&xml);
            apply_draw_param_defaults(&mut p, &mut t, &draw_params);
            (t, p, i)
        } else {
            (Vec::new(), Vec::new(), Vec::new())
        }
    } else {
        (Vec::new(), Vec::new(), Vec::new())
    };

    // 7. Parse page content (data layer)
    // Note: avoid shadowing `page_paths` (Vec<String> from Document.xml parsing)
    // Page content Layer has no DrawParam → OFD default: black (0,0,0). Do NOT apply
    // any DrawParam inheritance — invoice data text and ¥ symbol are naturally black.
    let (page_texts, page_obj_paths, page_imgs) = if let Some(page_path) = page_paths.first() {
        if let Some(xml) = zip_read_str(&mut archive, page_path) {
            parse_ofd_content(&xml)
        } else {
            (Vec::new(), Vec::new(), Vec::new())
        }
    } else {
        (Vec::new(), Vec::new(), Vec::new())
    };

    // 8. Parse annotations (watermark layer) — uses parse_annotations to handle Appearance offsets
    let annots_path = format!("{}/Annots/Page_0/Annotation.xml", base_dir);
    let (annot_texts, annot_imgs) = if let Some(xml) = zip_read_str(&mut archive, &annots_path) {
        parse_annotations(&xml)
    } else {
        (Vec::new(), Vec::new())
    };

    // 5b. Generate image data URLs, applying ImageMask compositing where needed
    // Collect all ImageObjects with ImageMask from template + page + annotations
    let mut mask_map: HashMap<u32, u32> = HashMap::new(); // resource_id → mask_resource_id
    {
        let collect_masks = |imgs: &[OfdImageObject], map: &mut HashMap<u32, u32>| {
            for img in imgs {
                if let Some(mask_id) = img.image_mask {
                    map.insert(img.resource_id, mask_id);
                }
            }
        };
        collect_masks(&tpl_imgs, &mut mask_map);
        collect_masks(&page_imgs, &mut mask_map);
        collect_masks(&annot_imgs, &mut mask_map);
    }

    let mut image_data: HashMap<u32, String> = HashMap::new();
    for (res_id, bytes) in &image_raw_bytes {
        if let Some(&mask_res_id) = mask_map.get(res_id) {
            // Composite: decode main image + mask, merge alpha channel, encode as RGBA PNG
            if let Some(mask_bytes) = image_raw_bytes.get(&mask_res_id) {
                if let Ok(main_img) = image::load_from_memory(bytes) {
                    if let Ok(mask_img) = image::load_from_memory(mask_bytes) {
                        let main_rgba = main_img.to_rgba8();
                        let mask_rgba = mask_img.to_rgba8();
                        // Both images must match dimensions
                        if main_rgba.width() == mask_rgba.width() && main_rgba.height() == mask_rgba.height() {
                            let mut composited = main_rgba.clone();
                            for (pixel, mask_pixel) in composited.pixels_mut().zip(mask_rgba.pixels()) {
                                // Mask: white (255) = opaque, black (0) = transparent
                                // Use the red channel of the mask as alpha
                                pixel[3] = mask_pixel[0];
                            }
                            let mut png_buf = Vec::new();
                            use std::io::Cursor;
                            if composited.write_to(&mut Cursor::new(&mut png_buf), image::ImageFormat::Png).is_ok() {
                                let b64 = base64::engine::general_purpose::STANDARD.encode(&png_buf);
                                image_data.insert(*res_id, format!("data:image/png;base64,{}", b64));
                                log::info!("ImageMask applied: resource {} masked by {}", res_id, mask_res_id);
                                continue;
                            }
                        } else {
                            log::warn!("ImageMask dimension mismatch: main={}x{}, mask={}x{}, skipping mask",
                                main_rgba.width(), main_rgba.height(), mask_rgba.width(), mask_rgba.height());
                        }
                    }
                }
            }
            // Fallback: if mask compositing failed, use the main image as-is
            log::warn!("ImageMask compositing failed for resource {}, using unmasked image", res_id);
        }
        // Default: encode as-is
        let file_name = image_file_names.get(res_id).map(|s| s.as_str()).unwrap_or("");
        let b64 = base64::engine::general_purpose::STANDARD.encode(bytes);
        let mime = if file_name.to_lowercase().ends_with(".png") || file_name.to_lowercase().ends_with(".bmp") {
            "image/png" // BMP decoded → re-encode as PNG for browser compatibility
        } else {
            "image/jpeg"
        };
        image_data.insert(*res_id, format!("data:{};base64,{}", mime, b64));
    }

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
        HashMap::new()
    };

    // 11. Extract invoice info from structured data
    let mut invoice_info = OfdInvoiceInfo::default();

    // From OFD.xml CustomData — skip empty strings (empty self-closing tags)
    let get_custom = |key: &str| -> Option<String> {
        custom_data.get(key).and_then(|s| if s.trim().is_empty() { None } else { Some(s.clone()) })
    };
    invoice_info.invoice_no = get_custom("发票号码");
    invoice_info.invoice_date = get_custom("开票日期");
    invoice_info.buyer_tax_id = get_custom("购买方纳税人识别号");
    invoice_info.seller_tax_id = get_custom("销售方纳税人识别号");
    invoice_info.amount_no_tax = custom_data.get("合计金额").and_then(|s| s.parse().ok());
    invoice_info.tax_amount = custom_data.get("合计税额").and_then(|s| s.parse().ok());

    // Compute total = no_tax + tax (both already in yuan, e.g. 17699.12 + 2300.88 = 20000.00)
    if let (Some(no_tax), Some(tax)) = (invoice_info.amount_no_tax, invoice_info.tax_amount) {
        invoice_info.amount_tax = Some(((no_tax + tax) * 100.0).round() / 100.0);
    }

    // From CustomTag.xml + Content.xml — get buyer/seller names
    // Build a text lookup: TextObject ID → text content
    let mut text_lookup: HashMap<u32, &str> = HashMap::new();
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
    if invoice_info.buyer_tax_id.is_none() {
        invoice_info.buyer_tax_id = get_tag_text("BuyerTaxID");
    }
    if invoice_info.seller_tax_id.is_none() {
        invoice_info.seller_tax_id = get_tag_text("SellerTaxID");
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
    // Also detect from page texts (when there's no template layer)
    if invoice_info.invoice_type.is_none() {
        for t in &page_texts {
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
    }

    // 11b. Text-based fallback extraction when no CustomData or CustomTag
    // This handles OFD files from non-tax producers (e.g., dzcp) that embed fonts
    // but don't include structured metadata.
    if invoice_info.invoice_no.is_none() && invoice_info.invoice_date.is_none()
        && invoice_info.buyer_name.is_none() && invoice_info.seller_name.is_none() {
        // Combine template + page texts (preserving order by ID)
        let mut all_texts: Vec<&OfdTextObject> = Vec::new();
        all_texts.extend(&tpl_texts);
        all_texts.extend(&page_texts);
        all_texts.sort_by_key(|t| t.id);

        let extracted = extract_invoice_from_text(&all_texts);

        // Only fill fields that are still None
        if invoice_info.invoice_no.is_none() { invoice_info.invoice_no = extracted.invoice_no; }
        if invoice_info.invoice_date.is_none() { invoice_info.invoice_date = extracted.invoice_date; }
        if invoice_info.buyer_name.is_none() { invoice_info.buyer_name = extracted.buyer_name; }
        if invoice_info.buyer_tax_id.is_none() { invoice_info.buyer_tax_id = extracted.buyer_tax_id; }
        if invoice_info.seller_name.is_none() { invoice_info.seller_name = extracted.seller_name; }
        if invoice_info.seller_tax_id.is_none() { invoice_info.seller_tax_id = extracted.seller_tax_id; }
        if invoice_info.amount_no_tax.is_none() { invoice_info.amount_no_tax = extracted.amount_no_tax; }
        if invoice_info.tax_amount.is_none() { invoice_info.tax_amount = extracted.tax_amount; }
        if invoice_info.amount_tax.is_none() { invoice_info.amount_tax = extracted.amount_tax; }
        if invoice_info.invoice_type.is_none() { invoice_info.invoice_type = extracted.invoice_type; }
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

/// Extract OFD page images as structured data (for bitmap fallback).
/// Returns `OfdExtractedImage` with base64 data URLs, dimensions, and file extension.
/// The caller can convert these to whatever type they need (e.g., FileData).
pub fn extract_ofd_images_raw(ofd_path: &str) -> Result<Vec<OfdExtractedImage>, String> {
    let images = extract_ofd_images(ofd_path)?;
    Ok(images.into_iter().map(|(data_url, ext, w, h)| OfdExtractedImage {
        data_url,
        ext,
        width: w,
        height: h,
    }).collect())
}

// =====================================================
// XML 数电票 Parsing (standalone .xml files)
// =====================================================

/// Parse a standalone XML 数电票 file and extract structured invoice data.
///
/// The XML format follows the 国家税务总局《电子凭证会计数据标准》specification,
/// with root element `<EInvoice>`. This is a pure data format with no layout info —
/// it cannot be rendered as a visual invoice page.
///
/// Returns `XmlInvoiceInfo` with key fields for file list display, summary export, etc.
pub fn parse_xml_invoice(xml_path: &str) -> Result<XmlInvoiceInfo, String> {
    let content = std::fs::read_to_string(xml_path)
        .map_err(|e| format!("读取 XML 文件失败: {}", e))?;

    // Quick check: must contain <EInvoice> root element
    if !content.contains("<EInvoice") {
        return Err("不是有效的数电票 XML 文件（缺少 EInvoice 根元素）".to_string());
    }

    let info = parse_xml_invoice_content(&content)?;
    Ok(info)
}

/// Parse XML 数电票 content string and extract structured invoice data.
fn parse_xml_invoice_content(content: &str) -> Result<XmlInvoiceInfo, String> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(content);
    reader.config_mut().trim_text(true);

    let mut info = XmlInvoiceInfo::default();
    let mut buf = Vec::new();

    // Track element path for context-aware parsing
    let mut path: Vec<String> = Vec::new();
    // Track LabelName values from EInvoiceType and GeneralOrSpecialVAT
    let mut einvoice_type_label: Option<String> = None;
    let mut general_or_special_label: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name_str = String::from_utf8_lossy(local.as_ref()).to_string();
                path.push(name_str);
            }
            Ok(Event::End(ref _e)) => {
                path.pop();
            }
            Ok(Event::Empty(ref _e)) => {
                // Self-closing tags like <SpecificInformation/> — nothing to extract
            }
            Ok(Event::Text(ref e)) => {
                if let Ok(text) = e.unescape() {
                    let text = text.trim();
                    if text.is_empty() { continue; }

                    let current_tag = path.last().map(|s| s.as_str()).unwrap_or("");
                    // Check parent context for LabelName disambiguation
                    let parent_tag = if path.len() >= 2 {
                        path.get(path.len() - 2).map(|s| s.as_str()).unwrap_or("")
                    } else {
                        ""
                    };

                    match current_tag {
                        // TaxSupervisionInfo
                        "InvoiceNumber" => info.invoice_no = Some(text.to_string()),
                        "IssueTime" => {
                            info.invoice_date = Some(text.split('T').next().unwrap_or(text).to_string());
                        }
                        // Seller — skip empty values (e.g. personal invoices)
                        "SellerName" => info.seller_name = Some(text.to_string()),
                        "SellerIdNum" if !text.is_empty() => info.seller_tax_id = Some(text.to_string()),
                        // Buyer — skip empty values (e.g. personal invoices where BuyerIdNum is empty)
                        "BuyerName" => info.buyer_name = Some(text.to_string()),
                        "BuyerIdNum" if !text.is_empty() => info.buyer_tax_id = Some(text.to_string()),
                        // BasicInformation amounts
                        "TotalAmWithoutTax" => info.amount_no_tax = text.parse().ok(),
                        "TotalTaxAm" => info.tax_amount = text.parse().ok(),
                        // TotalTax-includedAmount: tag name contains hyphen, quick-xml preserves it
                        "TotalTax-includedAmount" => info.amount_tax = text.parse().ok(),
                        // Invoice type: collect LabelName from different parent contexts
                        "LabelName" => match parent_tag {
                            "EInvoiceType" if einvoice_type_label.is_none() => {
                                einvoice_type_label = Some(text.to_string());
                            }
                            "GeneralOrSpecialVAT" if general_or_special_label.is_none() => {
                                general_or_special_label = Some(text.to_string());
                            }
                            _ => {}
                        },
                        _ => {}
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML 解析错误: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    // Compose invoice_type from EInvoiceType + GeneralOrSpecialVAT labels
    // e.g. "电子发票" + "普通发票" → "电子发票(普通发票)"
    // e.g. "电子发票" + "增值税专用发票" → "电子发票(增值税专用发票)"
    if let Some(special_label) = &general_or_special_label {
        let prefix = einvoice_type_label.as_deref().unwrap_or("电子发票");
        info.invoice_type = Some(format!("{}({})", prefix, special_label));
    } else if let Some(type_label) = &einvoice_type_label {
        info.invoice_type = Some(type_label.clone());
    }

    // Fallback: if amount_tax still empty, try alternate tag name
    if info.amount_tax.is_none() {
        // Some XML variants may use TotalTaxIncludedAmount instead of TotalTax-includedAmount
        let alt = content.find("<TotalTaxIncludedAmount>")
            .and_then(|start| {
                let text_start = start + "<TotalTaxIncludedAmount>".len();
                content[text_start..].find("</TotalTaxIncludedAmount>")
                    .map(|end| content[text_start..text_start + end].trim())
            });
        if let Some(v) = alt {
            info.amount_tax = v.parse().ok();
        }
    }

    Ok(info)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_xml_general_invoice_personal() {
        // 普通发票 - 个人购买方 (BuyerIdNum 为空)
        let xml = r#"<?xml version="1.0" encoding="utf-8"?>
<EInvoice>
  <Header>
    <InherentLabel>
      <EInvoiceType><LabelCode>01</LabelCode><LabelName>电子发票</LabelName></EInvoiceType>
      <GeneralOrSpecialVAT><LabelCode>02</LabelCode><LabelName>普通发票</LabelName></GeneralOrSpecialVAT>
    </InherentLabel>
  </Header>
  <EInvoiceData>
    <SellerInformation>
      <SellerIdNum>913416007050059877</SellerIdNum>
      <SellerName>中国联合网络通信有限公司亳州市分公司</SellerName>
    </SellerInformation>
    <BuyerInformation>
      <BuyerIdNum></BuyerIdNum>
      <BuyerName>高宗林（个人）</BuyerName>
    </BuyerInformation>
    <BasicInformation>
      <TotalAmWithoutTax>19.00</TotalAmWithoutTax>
      <TotalTaxAm>0.00</TotalTaxAm>
      <TotalTax-includedAmount>19.00</TotalTax-includedAmount>
    </BasicInformation>
  </EInvoiceData>
  <TaxSupervisionInfo>
    <InvoiceNumber>26347000000117553300</InvoiceNumber>
    <IssueTime>2026-05-05</IssueTime>
  </TaxSupervisionInfo>
</EInvoice>"#;

        let info = parse_xml_invoice_content(xml).unwrap();
        assert_eq!(info.invoice_no.as_deref(), Some("26347000000117553300"));
        assert_eq!(info.invoice_date.as_deref(), Some("2026-05-05"));
        assert_eq!(info.seller_name.as_deref(), Some("中国联合网络通信有限公司亳州市分公司"));
        assert_eq!(info.seller_tax_id.as_deref(), Some("913416007050059877"));
        assert_eq!(info.buyer_name.as_deref(), Some("高宗林（个人）"));
        assert_eq!(info.buyer_tax_id, None, "个人发票 BuyerIdNum 为空应为 None");
        assert_eq!(info.amount_no_tax, Some(19.00));
        assert_eq!(info.tax_amount, Some(0.00));
        assert_eq!(info.amount_tax, Some(19.00));
        assert_eq!(info.invoice_type.as_deref(), Some("电子发票(普通发票)"));
    }

    #[test]
    fn test_parse_xml_special_vat_invoice() {
        // 增值税专用发票
        let xml = r#"<?xml version="1.0" encoding="utf-8"?>
<EInvoice>
  <Header>
    <InherentLabel>
      <EInvoiceType><LabelCode>01</LabelCode><LabelName>电子发票</LabelName></EInvoiceType>
      <GeneralOrSpecialVAT><LabelCode>01</LabelCode><LabelName>增值税专用发票</LabelName></GeneralOrSpecialVAT>
    </InherentLabel>
  </Header>
  <EInvoiceData>
    <SellerInformation>
      <SellerIdNum>91320106751253359F</SellerIdNum>
      <SellerName>安元科技股份有限公司</SellerName>
    </SellerInformation>
    <BuyerInformation>
      <BuyerIdNum>9132020013590404XW</BuyerIdNum>
      <BuyerName>江苏苏豪天鹏农产品集团有限公司</BuyerName>
    </BuyerInformation>
    <BasicInformation>
      <TotalAmWithoutTax>18876.44</TotalAmWithoutTax>
      <TotalTaxAm>1793.56</TotalTaxAm>
      <TotalTax-includedAmount>20670.00</TotalTax-includedAmount>
    </BasicInformation>
  </EInvoiceData>
  <TaxSupervisionInfo>
    <InvoiceNumber>26322000004478296111</InvoiceNumber>
    <IssueTime>2026-06-04</IssueTime>
  </TaxSupervisionInfo>
</EInvoice>"#;

        let info = parse_xml_invoice_content(xml).unwrap();
        assert_eq!(info.invoice_no.as_deref(), Some("26322000004478296111"));
        assert_eq!(info.seller_name.as_deref(), Some("安元科技股份有限公司"));
        assert_eq!(info.buyer_tax_id.as_deref(), Some("9132020013590404XW"));
        assert_eq!(info.amount_no_tax, Some(18876.44));
        assert_eq!(info.tax_amount, Some(1793.56));
        assert_eq!(info.amount_tax, Some(20670.00));
        assert_eq!(info.invoice_type.as_deref(), Some("电子发票(增值税专用发票)"));
    }

    #[test]
    fn test_parse_xml_general_invoice_company() {
        // 普通发票 - 企业购买方
        let xml = r#"<?xml version="1.0" encoding="utf-8"?>
<EInvoice>
  <Header>
    <InherentLabel>
      <EInvoiceType><LabelCode>01</LabelCode><LabelName>电子发票</LabelName></EInvoiceType>
      <GeneralOrSpecialVAT><LabelCode>02</LabelCode><LabelName>普通发票</LabelName></GeneralOrSpecialVAT>
    </InherentLabel>
  </Header>
  <EInvoiceData>
    <SellerInformation>
      <SellerIdNum>52320200509244470T</SellerIdNum>
      <SellerName>无锡市安协安全培训中心</SellerName>
    </SellerInformation>
    <BuyerInformation>
      <BuyerIdNum>9132020013590404XW</BuyerIdNum>
      <BuyerName>江苏苏豪天鹏农产品集团有限公司</BuyerName>
    </BuyerInformation>
    <BasicInformation>
      <TotalAmWithoutTax>235.85</TotalAmWithoutTax>
      <TotalTaxAm>14.15</TotalTaxAm>
      <TotalTax-includedAmount>250.00</TotalTax-includedAmount>
    </BasicInformation>
  </EInvoiceData>
  <TaxSupervisionInfo>
    <InvoiceNumber>25322000000365404822</InvoiceNumber>
    <IssueTime>2025-08-08</IssueTime>
  </TaxSupervisionInfo>
</EInvoice>"#;

        let info = parse_xml_invoice_content(xml).unwrap();
        assert_eq!(info.invoice_no.as_deref(), Some("25322000000365404822"));
        assert_eq!(info.buyer_name.as_deref(), Some("江苏苏豪天鹏农产品集团有限公司"));
        assert_eq!(info.buyer_tax_id.as_deref(), Some("9132020013590404XW"));
        assert_eq!(info.amount_tax, Some(250.00));
        assert_eq!(info.invoice_type.as_deref(), Some("电子发票(普通发票)"));
    }

    #[test]
    fn test_parse_xml_not_einvoice() {
        let result = parse_xml_invoice_content("<root>not an invoice</root>");
        // parse_xml_invoice_content doesn't validate root element; that's done by parse_xml_invoice
        assert!(result.is_ok());
        let info = result.unwrap();
        assert_eq!(info.invoice_no, None);
    }

    #[test]
    fn test_parse_xml_issue_time_with_t() {
        // IssueTime may include time portion with T separator
        let xml = r#"<?xml version="1.0" encoding="utf-8"?>
<EInvoice>
  <EInvoiceData>
    <BasicInformation>
      <TotalAmWithoutTax>100</TotalAmWithoutTax>
      <TotalTaxAm>6</TotalTaxAm>
      <TotalTax-includedAmount>106</TotalTax-includedAmount>
    </BasicInformation>
  </EInvoiceData>
  <TaxSupervisionInfo>
    <InvoiceNumber>12345</InvoiceNumber>
    <IssueTime>2026-01-15T10:30:00</IssueTime>
  </TaxSupervisionInfo>
</EInvoice>"#;

        let info = parse_xml_invoice_content(xml).unwrap();
        assert_eq!(info.invoice_date.as_deref(), Some("2026-01-15"));
    }
}
