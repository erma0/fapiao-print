//! Windows-only PDFium GDI vector printing.
//! Bitmap rendering has been moved to pdfium_render.rs (cross-platform).
//! FFI bindings have been moved to pdfium_bindings.rs (cross-platform).

#![allow(non_camel_case_types)]

use std::ffi::c_void;
use std::ptr;
use std::sync::atomic::Ordering;

use windows::core::PCWSTR;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::Storage::Xps::*;

use crate::pdfium_bindings::{with_pdfium, pdfium_err_desc, FPDF_ANNOT, FPDF_PRINTING};

type FPDF_PAGE = *mut c_void;

type FnRenderPage = unsafe extern "C" fn(*mut c_void, FPDF_PAGE, i32, i32, i32, i32, i32, i32);

extern "C" {
    fn SafeCallRenderPage(
        func: FnRenderPage,
        dc: *mut c_void,
        page: FPDF_PAGE,
        x: i32,
        y: i32,
        w: i32,
        h: i32,
        rotate: i32,
        flags: i32,
    ) -> u32;
}

pub fn pdfium_vector_print(
    pdf_bytes: &[u8],
    printer_name: &str,
    copies: u32,
    duplex: bool,
    color_mode: &str,
    paper_w_mm: f32,
    paper_h_mm: f32,
    progress_cb: Option<&dyn Fn(u32, u32)>,
) -> Result<crate::pdf_engine::PdfResult, String> {
    if crate::pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
        return Err("应用正在关闭".to_string());
    }

    let pdf_len = pdf_bytes.len();
    if pdf_len > i32::MAX as usize {
        return Err(format!("PDF 文件过大 ({} bytes)，PDFium 不支持超过 2GB", pdf_len));
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
        log::info!("PDFium loaded PDF: {} pages, {} bytes", pc, pdf_len);
        Ok((doc, pc))
    })?;

    let printer_name_w: Vec<u16> = printer_name.encode_utf16().chain(std::iter::once(0)).collect();

    let mut base_devmode = match get_printer_default_devmode(printer_name) {
        Ok(dm) => {
            log::info!("Using printer default DEVMODE as base");
            Some(dm)
        }
        Err(e) => {
            log::warn!("Failed to get printer default DEVMODE ({}), using blank", e);
            None
        }
    };

    let dev_mode_buf = build_dev_mode(base_devmode.as_deref_mut(), copies, duplex, color_mode, paper_w_mm, paper_h_mm)?;
    let dev_mode_ptr = dev_mode_buf.as_ptr() as *const DEVMODEW;

    let hdc = unsafe {
        CreateDCW(
            None,
            PCWSTR(printer_name_w.as_ptr()),
            None,
            Some(dev_mode_ptr),
        )
    };

    if hdc.is_invalid() {
        with_pdfium(|funcs| { unsafe { (funcs.close_document)(doc); } Ok(()) })?;
        return Err(format!("无法打开打印机 DC: {}", printer_name));
    }

    let print_dc = hdc;

    unsafe { SetGraphicsMode(print_dc, GM_ADVANCED); }

    let printer_w = unsafe { GetDeviceCaps(print_dc, HORZRES) };
    let printer_h = unsafe { GetDeviceCaps(print_dc, VERTRES) };
    let printer_dpi = unsafe { GetDeviceCaps(print_dc, LOGPIXELSX) };
    log::info!("Printer DC: {}x{} px, {} DPI", printer_w, printer_h, printer_dpi);

    if printer_w <= 0 || printer_h <= 0 || printer_dpi <= 0 {
        unsafe { let _ = DeleteDC(print_dc); }
        with_pdfium(|funcs| { unsafe { (funcs.close_document)(doc); } Ok(()) })?;
        return Err(format!(
            "打印机DC返回无效尺寸 ({}x{} px, {} DPI)，请检查打印机设置和纸张配置",
            printer_w, printer_h, printer_dpi
        ));
    }

    let doc_name_w: Vec<u16> = "发票酱".encode_utf16().chain(std::iter::once(0)).collect();
    let doc_info = DOCINFOW {
        cbSize: std::mem::size_of::<DOCINFOW>() as i32,
        lpszDocName: PCWSTR(doc_name_w.as_ptr()),
        lpszOutput: PCWSTR::null(),
        lpszDatatype: PCWSTR::null(),
        fwType: 0,
    };

    let job_id = unsafe { StartDocW(print_dc, &doc_info) };
    if job_id <= 0 {
        unsafe { let _ = DeleteDC(print_dc); }
        with_pdfium(|funcs| { unsafe { (funcs.close_document)(doc); } Ok(()) })?;
        return Err("StartDoc 失败".to_string());
    }

    let mut pages_printed = 0u32;
    let mut last_error = String::new();
    let mut aborted = false;

    for page_idx in 0..page_count {
        if crate::pdf_engine::SHUTTING_DOWN.load(Ordering::SeqCst) {
            aborted = true;
            last_error = "打印被中止".to_string();
            break;
        }

        if let Some(ref cb) = progress_cb {
            cb(page_idx as u32, page_count as u32);
        }

        let start_page_result = unsafe { StartPage(print_dc) };
        if start_page_result <= 0 {
            last_error = format!("StartPage 失败 (page {})", page_idx + 1);
            continue;
        }

        let render_result = with_pdfium(|funcs| {
            let page = unsafe { (funcs.load_page)(doc, page_idx) };
            if page.is_null() {
                let err = unsafe { (funcs.get_last_error)() };
                return Err(format!("无法加载第 {} 页 (错误: {})", page_idx + 1, pdfium_err_desc(err)));
            }

            let render_flags = FPDF_ANNOT | FPDF_PRINTING;

            let seh_result = unsafe {
                SafeCallRenderPage(
                    funcs.render_page,
                    print_dc.0 as *mut c_void,
                    page,
                    0, 0, printer_w, printer_h,
                    0,
                    render_flags,
                )
            };

            if seh_result == 0 {
                unsafe { (funcs.close_page)(page); }
                return Ok(true);
            }

            log::warn!(
                "FPDF_RenderPage crashed (SEH code: {}), falling back to bitmap for page {}",
                seh_result, page_idx + 1
            );

            let page_w = unsafe { (funcs.get_page_width_f)(page) };
            let page_h = unsafe { (funcs.get_page_height_f)(page) };
            if page_w <= 0.0 || page_h <= 0.0 {
                unsafe { (funcs.close_page)(page); };
                return Err(format!("第 {} 页尺寸无效", page_idx + 1));
            }

            let scale = printer_dpi as f32 / 72.0;
            let bmp_w = (page_w * scale).round() as i32;
            let bmp_h = (page_h * scale).round() as i32;
            if bmp_w <= 0 || bmp_h <= 0 {
                unsafe { (funcs.close_page)(page); };
                return Err(format!("第 {} 页渲染尺寸无效", page_idx + 1));
            }

            let bitmap = unsafe { (funcs.bitmap_create)(bmp_w, bmp_h, 0) };
            if bitmap.is_null() {
                unsafe { (funcs.close_page)(page); };
                return Err(format!("创建位图失败 (第 {} 页)", page_idx + 1));
            }

            unsafe { (funcs.bitmap_fill_rect)(bitmap, 0, 0, bmp_w, bmp_h, 0xFFFFFFFF) };

            unsafe {
                (funcs.render_page_bitmap)(
                    bitmap, page,
                    0, 0, bmp_w, bmp_h,
                    0,
                    render_flags,
                );
            }

            let stride = unsafe { (funcs.bitmap_get_stride)(bitmap) };
            let buffer = unsafe { (funcs.bitmap_get_buffer)(bitmap) };

            if !buffer.is_null() && stride > 0 {
                let bmi = BITMAPINFO {
                    bmiHeader: BITMAPINFOHEADER {
                        biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                        biWidth: bmp_w,
                        biHeight: -bmp_h,
                        biPlanes: 1,
                        biBitCount: 32,
                        biCompression: BI_RGB.0,
                        biSizeImage: 0,
                        biXPelsPerMeter: 0,
                        biYPelsPerMeter: 0,
                        biClrUsed: 0,
                        biClrImportant: 0,
                    },
                    bmiColors: [RGBQUAD { rgbBlue: 0, rgbGreen: 0, rgbRed: 0, rgbReserved: 0 }],
                };

                unsafe {
                    StretchDIBits(
                        print_dc,
                        0, 0, printer_w, printer_h,
                        0, 0, bmp_w, bmp_h,
                        Some(buffer as *const c_void),
                        &bmi,
                        DIB_RGB_COLORS,
                        SRCCOPY,
                    );
                }
            }

            unsafe { (funcs.bitmap_destroy)(bitmap); }
            unsafe { (funcs.close_page)(page); }
            Ok(false)
        });

        match render_result {
            Ok(true) => pages_printed += 1,
            Ok(false) => {
                pages_printed += 1;
                if last_error.is_empty() {
                    last_error = "矢量渲染异常，已自动回退位图渲染".to_string();
                }
            }
            Err(e) => last_error = e,
        }

        unsafe { let _ = EndPage(print_dc); }
    }

    if aborted {
        unsafe { let _ = AbortDoc(print_dc); }
    } else {
        unsafe { let _ = EndDoc(print_dc); }
    }
    unsafe { let _ = DeleteDC(print_dc); }

    with_pdfium(|funcs| { unsafe { (funcs.close_document)(doc); } Ok(()) })?;

    if let Some(ref cb) = progress_cb {
        cb(page_count as u32, page_count as u32);
    }

    if pages_printed == 0 && !last_error.is_empty() {
        return Err(format!("打印失败: {}", last_error));
    }

    let mut msg = format!(
        "PDFium打印完成: {} 页 → {}",
        pages_printed, printer_name
    );
    if !last_error.is_empty() {
        msg.push_str(&format!(" (部分警告: {})", last_error));
    }

    Ok(crate::pdf_engine::PdfResult {
        success: true,
        message: msg,
        pdf_path: None,
        warnings: if last_error.is_empty() { None } else { Some(last_error) },
    })
}

fn get_printer_default_devmode(printer_name: &str) -> Result<Vec<u8>, String> {
    use windows::Win32::Graphics::Printing::{OpenPrinterW, ClosePrinter, DocumentPropertiesW, PRINTER_DEFAULTSW};
    use windows::Win32::Foundation::{HANDLE, HWND};
    use windows::core::PWSTR;

    let printer_name_w: Vec<u16> = printer_name.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        let defaults = PRINTER_DEFAULTSW {
            pDatatype: PWSTR::null(),
            pDevMode: std::ptr::null_mut(),
            DesiredAccess: windows::Win32::Graphics::Printing::PRINTER_ACCESS_USE,
        };
        let mut hprinter = HANDLE::default();

        OpenPrinterW(
            PCWSTR(printer_name_w.as_ptr()),
            &mut hprinter,
            Some(&defaults),
        )
        .map_err(|e| format!("无法打开打印机: {}", e))?;

        let null_hwnd = HWND::default();
        let dm_size = DocumentPropertiesW(
            null_hwnd,
            hprinter,
            PCWSTR(printer_name_w.as_ptr()),
            None,
            None,
            0,
        );
        if dm_size < 0 {
            let _ = ClosePrinter(hprinter);
            return Err(format!("DocumentPropertiesW 查询大小失败: {}", dm_size));
        }

        let dm_size = dm_size as usize;
        if dm_size < std::mem::size_of::<DEVMODEW>() {
            let _ = ClosePrinter(hprinter);
            return Err(format!("DEVMODE 大小异常: {} bytes", dm_size));
        }

        let mut dm_buf: Vec<u8> = vec![0u8; dm_size];
        let dm_ptr = dm_buf.as_mut_ptr() as *mut DEVMODEW;

        let result = DocumentPropertiesW(
            null_hwnd,
            hprinter,
            PCWSTR(printer_name_w.as_ptr()),
            Some(dm_ptr),
            None,
            DM_OUT_BUFFER.0 as u32,
        );
        let _ = ClosePrinter(hprinter);

        if result != 1 {
            return Err(format!("DocumentPropertiesW 获取默认设置失败: {}", result));
        }

        let dm_header = dm_buf.as_ptr() as *const DEVMODEW;
        let driver_extra = (*dm_header).dmDriverExtra as usize;
        log::info!(
            "Got default DEVMODE for printer: {} (total={} bytes, sizeof(DEVMODEW)={}, driverExtra={})",
            printer_name, dm_size, std::mem::size_of::<DEVMODEW>(), driver_extra
        );

        Ok(dm_buf)
    }
}

fn build_dev_mode(
    base: Option<&mut [u8]>,
    copies: u32,
    duplex: bool,
    color_mode: &str,
    paper_w_mm: f32,
    paper_h_mm: f32,
) -> Result<Vec<u8>, String> {
    let mut buf = match base {
        Some(b) => b.to_vec(),
        None => {
            let mut buf = vec![0u8; std::mem::size_of::<DEVMODEW>()];
            let dm = buf.as_mut_ptr() as *mut DEVMODEW;
            unsafe {
                (*dm).dmSize = std::mem::size_of::<DEVMODEW>() as u16;
                (*dm).dmDriverExtra = 0;
            }
            buf
        }
    };

    let dm = buf.as_mut_ptr() as *mut DEVMODEW;
    unsafe {
        if paper_w_mm > paper_h_mm {
            (*dm).Anonymous1.Anonymous1.dmOrientation = DMORIENT_LANDSCAPE as i16;
        } else {
            (*dm).Anonymous1.Anonymous1.dmOrientation = DMORIENT_PORTRAIT as i16;
        }
        (*dm).dmFields |= DM_ORIENTATION;

        if copies > 1 {
            (*dm).Anonymous1.Anonymous1.dmCopies = copies as i16;
            (*dm).dmFields |= DM_COPIES;
        }

        if duplex {
            (*dm).dmDuplex = DEVMODE_DUPLEX(DMDUP_VERTICAL.0);
            (*dm).dmFields |= DM_DUPLEX;
        }

        match color_mode {
            "grayscale" | "monochrome" | "bw" => {
                (*dm).dmColor = DEVMODE_COLOR(DMCOLOR_MONOCHROME.0);
                (*dm).dmFields |= DM_COLOR;
            }
            _ => {
                (*dm).dmColor = DEVMODE_COLOR(DMCOLOR_COLOR.0);
                (*dm).dmFields |= DM_COLOR;
            }
        }

        if let Some(paper) = infer_paper_size(paper_w_mm, paper_h_mm) {
            (*dm).Anonymous1.Anonymous1.dmPaperSize = paper as i16;
            (*dm).dmFields |= DM_PAPERSIZE;
        } else {
            (*dm).Anonymous1.Anonymous1.dmPaperSize = DMPAPER_USER as i16;
            (*dm).Anonymous1.Anonymous1.dmPaperWidth = (paper_w_mm * 10.0) as i16;
            (*dm).Anonymous1.Anonymous1.dmPaperLength = (paper_h_mm * 10.0) as i16;
            (*dm).dmFields |= DM_PAPERSIZE | DM_PAPERWIDTH | DM_PAPERLENGTH;
        }
    }

    Ok(buf)
}

fn infer_paper_size(w: f32, h: f32) -> Option<u32> {
    let sizes: [(f32, f32, u32); 6] = [
        (210.0, 297.0, DMPAPER_A4),
        (148.0, 210.0, DMPAPER_A5),
        (105.0, 148.0, DMPAPER_A6),
        (297.0, 420.0, DMPAPER_A3),
        (216.0, 279.0, DMPAPER_LETTER),
        (216.0, 356.0, DMPAPER_LEGAL),
    ];
    for (sw, sh, paper) in &sizes {
        if (w - sw).abs() < 2.0 && (h - sh).abs() < 2.0 {
            return Some(*paper);
        }
        if (w - sh).abs() < 2.0 && (h - sw).abs() < 2.0 {
            return Some(*paper);
        }
    }
    None
}
