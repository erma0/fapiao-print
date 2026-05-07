use super::{PdfTextWord, PdfTextLine, PdfTextResult, RENDER_DPI};

pub fn extract_pdf_text_with_pdf_oxide(
    pdf_path: &str,
    page_idx: u32,
) -> Result<PdfTextResult, String> {
    let doc = pdf_oxide::PdfDocument::open(pdf_path)
        .map_err(|e| format!("pdf_oxide 加载PDF失败: {}", e))?;

    let page_count = doc.page_count()
        .map_err(|e| format!("pdf_oxide 获取页数失败: {}", e))?;
    if (page_idx as usize) >= page_count {
        return Err(format!("PDF页面索引{}不存在 (共{}页)", page_idx, page_count));
    }

    let page_text = doc.extract_page_text(page_idx as usize)
        .map_err(|e| format!("pdf_oxide 文本提取失败: {}", e))?;

    let page_w_pt = page_text.page_width as f64;
    let page_h_pt = page_text.page_height as f64;
    let scale = RENDER_DPI as f64 / 72.0;
    let page_w_px = (page_w_pt * scale) as u32;
    let page_h_px = (page_h_pt * scale) as u32;

    let garbled_ratio = calc_garbled_ratio(&page_text);
    let needs_gbk_decode = garbled_ratio > 0.4;
    let gbk_decode_ok = if needs_gbk_decode {
        test_gbk_decode(&page_text)
    } else {
        false
    };
    let has_text_layer = !page_text.spans.is_empty()
        && (!needs_gbk_decode || gbk_decode_ok);

    let page_center_x = page_w_pt / 2.0;

    let mut all_words: Vec<PdfTextWord> = Vec::new();

    for span in &page_text.spans {
        if span.text.is_empty() {
            continue;
        }
        let span_x = span.bbox.x as f64;
        let span_y = span.bbox.y as f64;
        let span_w = span.bbox.width as f64;
        let span_h = span.bbox.height as f64;

        let fx = span_x * scale;
        let fy = (page_h_pt - span_y - span_h) * scale;

        let decoded_text = if needs_gbk_decode && gbk_decode_ok {
            decode_gbk_span(&span.text)
        } else {
            span.text.clone()
        };

        let span_right = span_x + span_w;
        if span_w > page_w_pt * 0.35 && span_x < page_center_x && span_right > page_center_x {
            let split_words = try_split_fused_span(
                &decoded_text, fx, fy, span_w * scale, span_h * scale,
                page_center_x * scale, page_w_pt * scale,
            );
            all_words.extend(split_words);
        } else {
            all_words.push(PdfTextWord {
                text: decoded_text,
                x: fx,
                y: fy.max(0.0),
                w: span_w * scale,
                h: span_h * scale,
            });
        }
    }

    let lines = group_words_into_lines(&all_words, page_h_px as f64);

    let text = if needs_gbk_decode && gbk_decode_ok {
        match extract_gbk_text_lopdf(pdf_path, page_idx) {
            Ok(lopdf_text) if !lopdf_text.is_empty() => lopdf_text,
            _ => lines.iter().map(|line| {
                line.words.iter().map(|w| w.text.as_str()).collect::<Vec<_>>().join("")
            }).collect::<Vec<_>>().join("\n"),
        }
    } else {
        lines.iter().map(|line| {
            line.words.iter().map(|w| w.text.as_str()).collect::<Vec<_>>().join("")
        }).collect::<Vec<_>>().join("\n")
    };

    Ok(PdfTextResult {
        text,
        lines,
        img_w: page_w_px,
        img_h: page_h_px,
        has_text_layer,
    })
}

fn extract_gbk_text_lopdf(pdf_path: &str, page_idx: u32) -> Result<String, String> {
    let doc = lopdf::Document::load(pdf_path)
        .map_err(|e| format!("lopdf 加载失败: {}", e))?;

    let pages = doc.get_pages();
    let page_num = page_idx + 1;
    let page_id = pages.get(&page_num)
        .ok_or_else(|| format!("页面不存在: {}", page_idx))?;

    let content_data = get_page_content_data(&doc, *page_id)?;
    let content = lopdf::content::Content::decode(&content_data)
        .map_err(|e| format!("内容流解析失败: {:?}", e))?;

    let mut text_parts: Vec<String> = Vec::new();
    let mut in_text_block = false;

    for op in &content.operations {
        match op.operator.as_str() {
            "BT" => { in_text_block = true; }
            "ET" => {
                in_text_block = false;
                text_parts.push("\n".to_string());
            }
            "Tj" | "'" => {
                if in_text_block && !op.operands.is_empty() {
                    if let Some(bytes) = operand_to_bytes(&op.operands[0]) {
                        let decoded = decode_gbk_bytes(&bytes);
                        if !decoded.is_empty() {
                            text_parts.push(decoded);
                        }
                    }
                }
            }
            "\"" => {
                if in_text_block && op.operands.len() >= 3 {
                    if let Some(bytes) = operand_to_bytes(&op.operands[2]) {
                        let decoded = decode_gbk_bytes(&bytes);
                        if !decoded.is_empty() {
                            text_parts.push(decoded);
                        }
                    }
                }
            }
            "TJ" => {
                if in_text_block && !op.operands.is_empty() {
                    if let Ok(arr) = op.operands[0].as_array() {
                        for item in arr {
                            if let Some(bytes) = operand_to_bytes(item) {
                                let decoded = decode_gbk_bytes(&bytes);
                                if !decoded.is_empty() {
                                    text_parts.push(decoded);
                                }
                            }
                        }
                    }
                }
            }
            "Td" | "TD" => {
                if in_text_block {
                    text_parts.push(" ".to_string());
                }
            }
            "T*" => {
                text_parts.push("\n".to_string());
            }
            _ => {}
        }
    }

    let raw = text_parts.join("");
    let cleaned = clean_lopdf_text(&raw);
    Ok(cleaned)
}

fn get_page_content_data(doc: &lopdf::Document, page_id: lopdf::ObjectId) -> Result<Vec<u8>, String> {
    let page_obj = doc.get_object(page_id)
        .map_err(|e| format!("获取页面对象失败: {:?}", e))?;
    let page_dict = page_obj.as_dict()
        .map_err(|e| format!("页面不是字典: {:?}", e))?;

    let contents = page_dict.get(b"Contents")
        .map_err(|_| "页面无Contents".to_string())?;

    match contents {
        lopdf::Object::Reference(r) => {
            let stream = doc.get_object(*r)
                .map_err(|e| format!("获取内容流失败: {:?}", e))?;
            let s = stream.as_stream()
                .map_err(|e| format!("内容不是流: {:?}", e))?;
            s.decompressed_content()
                .map_err(|e| format!("解压内容流失败: {:?}", e))
        }
        lopdf::Object::Array(arr) => {
            let mut data = Vec::new();
            for item in arr {
                if let Ok(r) = item.as_reference() {
                    if let Ok(stream_obj) = doc.get_object(r) {
                        if let Ok(s) = stream_obj.as_stream() {
                            if let Ok(decompressed) = s.decompressed_content() {
                                data.extend_from_slice(&decompressed);
                            } else {
                                data.extend_from_slice(&s.content);
                            }
                        }
                    }
                }
            }
            Ok(data)
        }
        _ => Err("Contents格式不支持".to_string()),
    }
}

fn operand_to_bytes(obj: &lopdf::Object) -> Option<Vec<u8>> {
    match obj {
        lopdf::Object::String(bytes, _) => Some(bytes.clone()),
        _ => None,
    }
}

fn decode_gbk_bytes(bytes: &[u8]) -> String {
    let (decoded, _, _) = encoding_rs::GBK.decode(bytes);
    decoded.to_string()
}

fn clean_lopdf_text(text: &str) -> String {
    let mut lines: Vec<String> = Vec::new();
    for line in text.split('\n') {
        let trimmed = line.trim();
        if !trimmed.is_empty() {
            lines.push(trimmed.to_string());
        }
    }
    lines.join("\n")
}

fn calc_garbled_ratio(page_text: &pdf_oxide::PageText) -> f64 {
    if page_text.spans.is_empty() {
        return 0.0;
    }
    let mut garbled_count: usize = 0;
    let mut cjk_count: usize = 0;
    let mut ascii_count: usize = 0;
    for span in &page_text.spans {
        for ch in span.text.chars() {
            if ch.is_ascii() {
                ascii_count += 1;
            } else if is_cjk(ch) {
                cjk_count += 1;
            } else if is_gbk_byte_pair_char(ch) {
                garbled_count += 1;
            }
        }
    }
    let meaningful = cjk_count + ascii_count;
    if garbled_count + meaningful == 0 {
        return 0.0;
    }
    garbled_count as f64 / (garbled_count + meaningful) as f64
}

fn test_gbk_decode(page_text: &pdf_oxide::PageText) -> bool {
    let mut total_gbk_chars: usize = 0;
    let mut valid_gbk_chars: usize = 0;
    let mut total_cjk_result: usize = 0;
    for span in page_text.spans.iter().take(10) {
        for ch in span.text.chars() {
            if is_gbk_byte_pair_char(ch) {
                total_gbk_chars += 1;
                let cp = ch as u32;
                let hi = ((cp >> 8) & 0xFF) as u8;
                let lo = (cp & 0xFF) as u8;
                if hi >= 0x81 && lo >= 0x40 && lo != 0x7F {
                    valid_gbk_chars += 1;
                    let buf = [hi, lo];
                    let (decoded, _, _) = encoding_rs::GBK.decode(&buf);
                    let s = decoded.trim();
                    if s.chars().any(|c| is_cjk(c) || c.is_ascii_alphanumeric()) {
                        total_cjk_result += 1;
                    }
                }
            }
        }
    }
    if total_gbk_chars == 0 {
        return false;
    }
    let valid_ratio = valid_gbk_chars as f64 / total_gbk_chars as f64;
    let cjk_ratio = total_cjk_result as f64 / total_gbk_chars as f64;
    valid_ratio > 0.7 && cjk_ratio > 0.5
}

fn decode_gbk_span(text: &str) -> String {
    let mut gbk_bytes = Vec::new();
    let mut result = String::new();
    for ch in text.chars() {
        let cp = ch as u32;
        if is_gbk_byte_pair_char(ch) && cp <= 0xFFFF {
            let hi = ((cp >> 8) & 0xFF) as u8;
            let lo = (cp & 0xFF) as u8;
            gbk_bytes.push(hi);
            gbk_bytes.push(lo);
        } else {
            if !gbk_bytes.is_empty() {
                let (decoded, _, _) = encoding_rs::GBK.decode(&gbk_bytes);
                result.push_str(&decoded);
                gbk_bytes.clear();
            }
            result.push(ch);
        }
    }
    if !gbk_bytes.is_empty() {
        let (decoded, _, _) = encoding_rs::GBK.decode(&gbk_bytes);
        result.push_str(&decoded);
    }
    result
}

fn group_words_into_lines(words: &[PdfTextWord], _page_h: f64) -> Vec<PdfTextLine> {
    if words.is_empty() {
        return Vec::new();
    }

    let median_h = {
        let mut heights: Vec<f64> = words.iter().map(|w| w.h).collect();
        heights.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
        heights[heights.len() / 2]
    };
    let band = if median_h > 0.5 { median_h * 0.5 } else { 1.0 };

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
        if (word.y - current_y).abs() <= line_height * 0.5 {
            current_words.push(word.clone());
        } else {
            if !current_words.is_empty() {
                current_words.sort_by(|a, b| a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal));
                current_words = merge_adjacent_words(&current_words);
                lines.push(PdfTextLine {
                    words: current_words.clone(),
                    confidence: 1.0,
                });
            }
            current_words.clear();
            current_words.push(word.clone());
            current_y = word.y;
            line_height = word.h;
        }
    }
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

fn merge_adjacent_words(words: &[PdfTextWord]) -> Vec<PdfTextWord> {
    if words.len() <= 1 {
        return words.to_vec();
    }
    let max_gap = 3.0;
    let mut result: Vec<PdfTextWord> = Vec::new();
    let mut buf = words[0].clone();

    for i in 1..words.len() {
        let next = &words[i];
        let gap = next.x - (buf.x + buf.w);
        if gap >= 0.0 && gap <= max_gap {
            buf.text.push_str(&next.text);
            buf.w = (next.x + next.w) - buf.x;
            if next.h > buf.h { buf.h = next.h; }
        } else if gap < 0.0 && gap >= -max_gap {
            buf.text.push_str(&next.text);
            buf.w = buf.w.max(next.x + next.w - buf.x);
            if next.h > buf.h { buf.h = next.h; }
        } else {
            result.push(buf);
            buf = next.clone();
        }
    }
    result.push(buf);
    result
}

fn is_cjk(ch: char) -> bool {
    matches!(ch,
        '\u{4E00}'..='\u{9FFF}' |
        '\u{3400}'..='\u{4DBF}' |
        '\u{F900}'..='\u{FAFF}' |
        '\u{2E80}'..='\u{2EFF}' |
        '\u{3000}'..='\u{303F}' |
        '\u{FF00}'..='\u{FFEF}'
    )
}

fn is_gbk_byte_pair_char(ch: char) -> bool {
    let cp = ch as u32;
    matches!(cp,
        0xAC00..=0xD7AF |
        0x1100..=0x11FF |
        0x3130..=0x318F |
        0xA500..=0xA82F |
        0xA840..=0xA87F |
        0xA880..=0xA8DF |
        0xA900..=0xA95F |
        0xA960..=0xA97F |
        0xD7B0..=0xD7FF |
        0x0E00..=0x0E7F |
        0x0E80..=0x0EFF |
        0x1000..=0x109F |
        0x0F00..=0x0FFF |
        0x0C00..=0x0C7F |
        0x0B80..=0x0BFF |
        0x0980..=0x09FF |
        0x0A80..=0x0AFF |
        0x0B00..=0x0B7F
    )
}

fn try_split_fused_span(
    text: &str,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    center_x: f64,
    _page_w: f64,
) -> Vec<PdfTextWord> {
    let split_pos = find_fused_split_point(text);
    if let Some(pos) = split_pos {
        let left_text = &text[..pos];
        let right_text = &text[pos..];
        if left_text.is_empty() || right_text.is_empty() {
            return vec![PdfTextWord { text: text.to_string(), x, y, w, h }];
        }
        let ratio = left_text.chars().count() as f64 / text.chars().count() as f64;
        let left_w = w * ratio;
        let right_w = w - left_w;
        vec![
            PdfTextWord {
                text: left_text.to_string(),
                x,
                y,
                w: left_w,
                h,
            },
            PdfTextWord {
                text: right_text.to_string(),
                x: x + left_w,
                y,
                w: right_w,
                h,
            },
        ]
    } else {
        let ratio = ((center_x - x) / w).max(0.2).min(0.8);
        let left_w = w * ratio;
        let right_w = w - left_w;
        let left_text = text.to_string();
        vec![
            PdfTextWord { text: left_text, x, y, w: left_w, h },
            PdfTextWord { text: String::new(), x: x + left_w, y, w: right_w, h },
        ]
    }
}

fn find_fused_split_point(text: &str) -> Option<usize> {
    let company_suffixes = [
        "有限公司", "股份有限公司", "有限责任公司", "集团",
        "工作室", "商行", "商店", "经营部", "服务部",
        "事务所", "合作社", "合伙企业", "厂", "店", "部", "院", "所", "中心",
    ];
    for suffix in &company_suffixes {
        if let Some(pos) = text.find(suffix) {
            let end = pos + suffix.len();
            if end < text.len() {
                let after = &text[end..];
                if after.chars().next().map_or(false, |c| c.is_ascii_alphanumeric() || is_cjk(c)) {
                    return Some(end);
                }
            }
        }
    }
    let bytes = text.as_bytes();
    let mut digit_start: Option<usize> = None;
    let mut digit_len: usize = 0;
    for (i, &b) in bytes.iter().enumerate() {
        if b.is_ascii_digit() {
            if digit_start.is_none() {
                digit_start = Some(i);
            }
            digit_len += 1;
        } else if digit_len >= 18 {
            let start = digit_start.unwrap();
            if start + 18 < text.len() {
                let after_char = text.as_bytes()[start + 18];
                if after_char.is_ascii_alphanumeric() {
                    return Some(start + 18);
                }
            }
            digit_start = None;
            digit_len = 0;
        } else {
            digit_start = None;
            digit_len = 0;
        }
    }
    if digit_len >= 36 {
        let start = digit_start.unwrap();
        return Some(start + 18);
    }
    None
}
