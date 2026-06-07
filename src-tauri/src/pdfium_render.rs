use std::ffi::c_void;
use std::ptr;
use std::sync::atomic::Ordering;

use crate::pdfium_bindings::{with_pdfium, pdfium_err_desc, FPDF_ANNOT};

pub struct RenderedImage {
    pub index: u32,
    pub image_data_url: String,
    pub width: u32,
    pub height: u32,
    pub render_dpi: u32,
    pub format: String,
}

pub fn render_pdf_to_images(
    pdf_bytes: &[u8],
    dpi: u32,
    use_jpeg: bool,
) -> Result<Vec<RenderedImage>, String> {
    if crate::pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    let pdf_len = pdf_bytes.len();
    if pdf_len > i32::MAX as usize {
        return Err(format!("PDF 文件过大 ({} bytes)", pdf_len));
    }

    let (doc, page_count) = with_pdfium(|funcs| {
        let doc = unsafe {
            (funcs.load_mem_document)(
                pdf_bytes.as_ptr() as *const c_void,
                pdf_len as i32,
                ptr::null(),
            )
        };
        if doc.is_null() {
            let err = unsafe { (funcs.get_last_error)() };
            return Err(format!("PDFium 无法加载 PDF 文档 (错误: {})", pdfium_err_desc(err)));
        }
        let pc = unsafe { (funcs.get_page_count)(doc) };
        if pc <= 0 {
            let err = unsafe { (funcs.get_last_error)() };
            unsafe { (funcs.close_document)(doc) };
            return Err(format!("PDF 文档没有页面 (错误: {})", pdfium_err_desc(err)));
        }
        log::info!("PDFium render: loaded PDF, {} pages, {} bytes", pc, pdf_len);
        Ok((doc, pc))
    })?;

    let mut results = Vec::new();

    for page_idx in 0..page_count {
        if crate::pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
            with_pdfium(|funcs| { unsafe { (funcs.close_document)(doc); } Ok(()) })?;
            return Err("应用正在关闭".to_string());
        }

        let rendered = with_pdfium(|funcs| {
            let page = unsafe { (funcs.load_page)(doc, page_idx) };
            if page.is_null() {
                let err = unsafe { (funcs.get_last_error)() };
                return Err(format!("无法加载第 {} 页 (错误: {})", page_idx + 1, pdfium_err_desc(err)));
            }

            let page_w = unsafe { (funcs.get_page_width_f)(page) };
            let page_h = unsafe { (funcs.get_page_height_f)(page) };
            if page_w <= 0.0 || page_h <= 0.0 {
                unsafe { (funcs.close_page)(page) };
                return Err(format!("第 {} 页尺寸无效 ({:.1}x{:.1})", page_idx + 1, page_w, page_h));
            }

            let scale = dpi as f32 / 72.0;
            let bmp_w = (page_w * scale).round() as i32;
            let bmp_h = (page_h * scale).round() as i32;
            if bmp_w <= 0 || bmp_h <= 0 {
                unsafe { (funcs.close_page)(page) };
                return Err(format!("第 {} 页渲染尺寸无效 ({}x{})", page_idx + 1, bmp_w, bmp_h));
            }

            let bitmap = unsafe { (funcs.bitmap_create)(bmp_w, bmp_h, 0) };
            if bitmap.is_null() {
                unsafe { (funcs.close_page)(page) };
                return Err(format!("创建位图失败 (第 {} 页, {}x{})", page_idx + 1, bmp_w, bmp_h));
            }

            unsafe { (funcs.bitmap_fill_rect)(bitmap, 0, 0, bmp_w, bmp_h, 0xFFFFFFFF) };

            unsafe {
                (funcs.render_page_bitmap)(
                    bitmap, page,
                    0, 0, bmp_w, bmp_h,
                    0,
                    FPDF_ANNOT,
                );
            }

            let stride = unsafe { (funcs.bitmap_get_stride)(bitmap) };
            let buffer = unsafe { (funcs.bitmap_get_buffer)(bitmap) };

            let mut encoded_data = Vec::new();
            if !buffer.is_null() && stride > 0 {
                let row_len = (bmp_w as usize * 4).min(stride as usize);
                let img = unsafe {
                    let mut rgba = Vec::with_capacity(bmp_w as usize * bmp_h as usize * 4);
                    for y in 0..bmp_h {
                        let row_start = buffer.add(y as usize * stride as usize);
                        let row = std::slice::from_raw_parts(row_start as *const u8, row_len);
                        for x in 0..bmp_w as usize {
                            let b = row[x * 4];
                            let g = row[x * 4 + 1];
                            let r = row[x * 4 + 2];
                            rgba.push(r);
                            rgba.push(g);
                            rgba.push(b);
                            rgba.push(255);
                        }
                    }
                    image::RgbaImage::from_raw(bmp_w as u32, bmp_h as u32, rgba)
                        .unwrap_or_else(|| image::RgbaImage::new(bmp_w as u32, bmp_h as u32))
                };

                let mut cursor = std::io::Cursor::new(&mut encoded_data);

                if use_jpeg {
                    let jpeg_quality: u8 = 80;
                    let mut encoder = image::codecs::jpeg::JpegEncoder::new_with_quality(
                        &mut cursor, jpeg_quality,
                    );
                    encoder.encode(
                        img.as_raw(), bmp_w as u32, bmp_h as u32,
                        image::ExtendedColorType::Rgba8,
                    ).map_err(|e| format!("JPEG 编码失败 (第 {} 页): {}", page_idx + 1, e))?;
                } else {
                    img.write_to(&mut cursor, image::ImageFormat::Png)
                        .map_err(|e| format!("PNG 编码失败 (第 {} 页): {}", page_idx + 1, e))?;
                }
            }

            unsafe { (funcs.bitmap_destroy)(bitmap) };
            unsafe { (funcs.close_page)(page) };

            use base64::Engine;
            let b64 = base64::engine::general_purpose::STANDARD.encode(&encoded_data);
            let (mime, format) = if use_jpeg { ("image/jpeg", "jpeg") } else { ("image/png", "png") };
            let data_url = format!("data:{};base64,{}", mime, b64);

            log::info!("PDFium rendered page {} ({}x{}) @ {}dpi [{}]", page_idx + 1, bmp_w, bmp_h, dpi, format);

            Ok(RenderedImage {
                index: page_idx as u32,
                image_data_url: data_url,
                width: bmp_w as u32,
                height: bmp_h as u32,
                render_dpi: dpi,
                format: format.to_string(),
            })
        });

        match rendered {
            Ok(img) => results.push(img),
            Err(e) => {
                log::warn!("PDFium render page {} failed: {}", page_idx + 1, e);
                let _ = with_pdfium(|funcs| { unsafe { (funcs.close_document)(doc); } Ok(()) });
                return Err(e);
            }
        }
    }

    with_pdfium(|funcs| { unsafe { (funcs.close_document)(doc); } Ok(()) })?;

    Ok(results)
}
