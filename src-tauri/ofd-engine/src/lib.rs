//! OFD (Open Fixed-layout Document) parser and SVG renderer.
//!
//! Supports Chinese electronic invoices (发票): extracts structured invoice data
//! from OFD XML metadata (CustomData + CustomTag) and renders pages as SVG.
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

// =====================================================
// OFD Content Parsing
// =====================================================

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
    draw_params: &HashMap<u32, OfdDrawParam>,
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
        (HashMap::new(), HashMap::new(), HashMap::new())
    };

    // 5. Parse DocumentRes.xml for image resources
    let doc_res_path = format!("{}/DocumentRes.xml", base_dir);
    let image_map = if let Some(xml) = zip_read_str(&mut archive, &doc_res_path) {
        parse_image_resources(&xml)
    } else {
        HashMap::new()
    };

    // Load actual image data from ZIP
    let mut image_data: HashMap<u32, String> = HashMap::new();
    for (res_id, file_name) in &image_map {
        let img_path = format!("{}/Res/{}", base_dir, file_name);
        if let Some(bytes) = zip_read_bytes(&mut archive, &img_path) {
            let b64 = base64::engine::general_purpose::STANDARD.encode(&bytes);
            let mime = if file_name.to_lowercase().ends_with(".png") { "image/png" } else { "image/jpeg" };
            image_data.insert(*res_id, format!("data:{};base64,{}", mime, b64));
        }
    }

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
