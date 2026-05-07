use std::time::Instant;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let pdf_path = if args.len() > 1 { &args[1] } else {
        eprintln!("Usage: bench_extract <pdf_path>");
        std::process::exit(1);
    };

    println!("PDF: {}", pdf_path);
    println!();

    // Benchmark pdf_oxide
    let iterations = 5;
    let mut pdf_oxide_times = Vec::new();
    for i in 0..iterations {
        let start = Instant::now();
        match pdf_oxide::PdfDocument::open(pdf_path) {
            Ok(doc) => {
                let open_time = start.elapsed();
                let page_count = doc.page_count().unwrap_or(0);
                for page_idx in 0..page_count {
                    let _ = doc.extract_page_text(page_idx);
                }
                let total_time = start.elapsed();
                pdf_oxide_times.push(total_time);
                println!("[pdf_oxide] iter {}: open={:.1}ms total={:.1}ms pages={}",
                    i, open_time.as_secs_f64() * 1000.0, total_time.as_secs_f64() * 1000.0, page_count);
            }
            Err(e) => println!("[pdf_oxide] FAILED: {}", e),
        }
    }

    // Benchmark lopdf
    let mut lopdf_times = Vec::new();
    for i in 0..iterations {
        let start = Instant::now();
        match lopdf::Document::load(pdf_path) {
            Ok(doc) => {
                let open_time = start.elapsed();
                let pages = doc.get_pages();
                let page_count = pages.len();
                for (&page_num, &page_id) in &pages {
                    let _ = extract_text_lopdf(&doc, page_id, page_num);
                }
                let total_time = start.elapsed();
                lopdf_times.push(total_time);
                println!("[lopdf]    iter {}: open={:.1}ms total={:.1}ms pages={}",
                    i, open_time.as_secs_f64() * 1000.0, total_time.as_secs_f64() * 1000.0, page_count);
            }
            Err(e) => println!("[lopdf] FAILED: {}", e),
        }
    }

    // Benchmark pdf_oxide + lopdf (current approach for GBK PDFs)
    let mut combined_times = Vec::new();
    for i in 0..iterations {
        let start = Instant::now();
        // Step 1: pdf_oxide
        if let Ok(doc) = pdf_oxide::PdfDocument::open(pdf_path) {
            let page_count = doc.page_count().unwrap_or(0);
            for page_idx in 0..page_count {
                let page_text = doc.extract_page_text(page_idx).unwrap();
                let _garbled = calc_garbled_ratio_bench(&page_text);
                // Simulate GBK decode for each span
                for span in &page_text.spans {
                    let _ = decode_gbk_span_bench(&span.text);
                }
            }
        }
        // Step 2: lopdf (GBK fallback)
        if let Ok(doc) = lopdf::Document::load(pdf_path) {
            let pages = doc.get_pages();
            for (&page_num, &page_id) in &pages {
                let _ = extract_text_lopdf(&doc, page_id, page_num);
            }
        }
        let total_time = start.elapsed();
        combined_times.push(total_time);
        println!("[combined] iter {}: total={:.1}ms", i, total_time.as_secs_f64() * 1000.0);
    }

    println!("\n=== Summary ===");
    if !pdf_oxide_times.is_empty() {
        let avg: f64 = pdf_oxide_times.iter().map(|t| t.as_secs_f64() * 1000.0).sum::<f64>() / pdf_oxide_times.len() as f64;
        println!("pdf_oxide avg: {:.1}ms", avg);
    }
    if !lopdf_times.is_empty() {
        let avg: f64 = lopdf_times.iter().map(|t| t.as_secs_f64() * 1000.0).sum::<f64>() / lopdf_times.len() as f64;
        println!("lopdf avg:    {:.1}ms", avg);
    }
    if !combined_times.is_empty() {
        let avg: f64 = combined_times.iter().map(|t| t.as_secs_f64() * 1000.0).sum::<f64>() / combined_times.len() as f64;
        println!("combined avg: {:.1}ms", avg);
    }
}

fn extract_text_lopdf(doc: &lopdf::Document, page_id: lopdf::ObjectId, _page_num: u32) -> String {
    let mut result = String::new();
    let page_obj = match doc.get_object(page_id) { Ok(o) => o, Err(_) => return result };
    let page_dict = match page_obj.as_dict() { Ok(d) => d, Err(_) => return result };
    let contents = match page_dict.get(b"Contents") { Ok(c) => c, Err(_) => return result };

    let content_data = match contents {
        lopdf::Object::Reference(r) => {
            match doc.get_object(*r) {
                Ok(stream_obj) => match stream_obj.as_stream() {
                    Ok(s) => s.decompressed_content().unwrap_or_else(|_| s.content.clone()),
                    Err(_) => return result,
                },
                Err(_) => return result,
            }
        }
        lopdf::Object::Array(arr) => {
            let mut data = Vec::new();
            for item in arr {
                if let Ok(r) = item.as_reference() {
                    if let Ok(stream_obj) = doc.get_object(r) {
                        if let Ok(s) = stream_obj.as_stream() {
                            match s.decompressed_content() {
                                Ok(d) => data.extend_from_slice(&d),
                                Err(_) => data.extend_from_slice(&s.content),
                            }
                        }
                    }
                }
            }
            data
        }
        _ => return result,
    };

    let content = match lopdf::content::Content::decode(&content_data) { Ok(c) => c, Err(_) => return result };

    let mut in_text = false;
    for op in &content.operations {
        match op.operator.as_str() {
            "BT" => { in_text = true; }
            "ET" => { in_text = false; }
            "Tj" | "'" => {
                if in_text && !op.operands.is_empty() {
                    if let lopdf::Object::String(bytes, _) = &op.operands[0] {
                        let (decoded, _, _) = encoding_rs::GBK.decode(bytes);
                        result.push_str(&decoded);
                    }
                }
            }
            "TJ" => {
                if in_text && !op.operands.is_empty() {
                    if let Ok(arr) = op.operands[0].as_array() {
                        for item in arr {
                            if let lopdf::Object::String(bytes, _) = item {
                                let (decoded, _, _) = encoding_rs::GBK.decode(bytes);
                                result.push_str(&decoded);
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }
    result
}

fn calc_garbled_ratio_bench(page_text: &pdf_oxide::PageText) -> f64 {
    let mut garbled = 0usize;
    let mut meaningful = 0usize;
    for span in &page_text.spans {
        for ch in span.text.chars() {
            if ch.is_ascii() || is_cjk_bench(ch) { meaningful += 1; }
            else if is_gbk_byte_pair_bench(ch) { garbled += 1; }
        }
    }
    if garbled + meaningful == 0 { return 0.0; }
    garbled as f64 / (garbled + meaningful) as f64
}

fn decode_gbk_span_bench(text: &str) -> String {
    let mut gbk_bytes = Vec::new();
    let mut result = String::new();
    for ch in text.chars() {
        let cp = ch as u32;
        if is_gbk_byte_pair_bench(ch) && cp <= 0xFFFF {
            gbk_bytes.push(((cp >> 8) & 0xFF) as u8);
            gbk_bytes.push((cp & 0xFF) as u8);
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

fn is_cjk_bench(ch: char) -> bool {
    matches!(ch, '\u{4E00}'..='\u{9FFF}' | '\u{3400}'..='\u{4DBF}' | '\u{F900}'..='\u{FAFF}' | '\u{FF00}'..='\u{FFEF}')
}

fn is_gbk_byte_pair_bench(ch: char) -> bool {
    let cp = ch as u32;
    matches!(cp, 0xAC00..=0xD7AF | 0x1100..=0x11FF | 0x3130..=0x318F | 0xA500..=0xA82F | 0xA840..=0xA87F | 0xA880..=0xA8DF | 0xA900..=0xA95F | 0xA960..=0xA97F | 0xD7B0..=0xD7FF | 0x0E00..=0x0E7F | 0x0E80..=0x0EFF | 0x1000..=0x109F | 0x0F00..=0x0FFF | 0x0C00..=0x0C7F | 0x0B80..=0x0BFF | 0x0980..=0x09FF | 0x0A80..=0x0AFF | 0x0B00..=0x0B7F)
}
