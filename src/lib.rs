// 发票酱 v2.1 — Web 版
// 纯 Axum 服务器，无 Tauri 依赖

use std::sync::atomic::AtomicBool;

pub mod pdf_engine;
pub mod pdfium_bindings;
pub mod pdfium_render;
pub mod platform;
pub mod session;
pub mod server;
#[cfg(target_os = "windows")]
pub mod pdfium_print;

pub static DOWNLOAD_CANCELLED: AtomicBool = AtomicBool::new(false);

/// Windows: ShellExecute print — try printto (silent) → print → open fallback.
/// Returns true if printed silently, false if opened in viewer.
#[cfg(target_os = "windows")]
pub fn shell_execute_print(pdf_path: &std::path::Path, printer_name: Option<&str>) -> Result<bool, String> {
    use windows::core::{HSTRING, PCWSTR};
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::{SW_HIDE, SW_SHOWNORMAL, SW_SHOW};

    let resolved_printer: Option<String> = match printer_name {
        Some(name) => Some(name.to_string()),
        None => platform::get_default_printer_name(),
    };
    let printer_str = resolved_printer.as_deref()
        .ok_or("未找到默认打印机，请在系统设置中配置打印机，或在打印设置中手动选择。")?;

    let _com = pdf_engine::ComGuard::init();
    unsafe {
        let file: HSTRING = pdf_path.to_string_lossy().to_string().into();

        // Strategy 1: ShellExecuteW "printto" — specify printer, silent
        let verb: HSTRING = "printto".into();
        let printer_hstring: HSTRING = printer_str.into();
        let params = PCWSTR::from_raw(printer_hstring.as_ptr());
        let ret = ShellExecuteW(None, &verb, &file, params, PCWSTR::null(), SW_HIDE);
        if ret.0 as isize > 32 { return Ok(true); }
        log::warn!("ShellExecute printto failed (code: {}), trying simple print", ret.0 as isize);

        // Strategy 2: ShellExecuteW "print" without specifying printer
        let verb: HSTRING = "print".into();
        let ret = ShellExecuteW(None, &verb, &file, PCWSTR::null(), PCWSTR::null(), SW_SHOW);
        if ret.0 as isize > 32 { return Ok(true); }
        log::warn!("ShellExecute print failed (code: {}), falling back to open", ret.0 as isize);

        // Strategy 3: Fallback — open the PDF so user can print manually
        let verb: HSTRING = "open".into();
        let ret = ShellExecuteW(None, &verb, &file, PCWSTR::null(), PCWSTR::null(), SW_SHOWNORMAL);
        if ret.0 as isize > 32 { return Ok(false); }

        Err(format!(
            "打印失败，错误码: {}。请检查打印机连接或安装PDF阅读器",
            ret.0 as isize
        ))
    }
}

#[cfg(not(target_os = "windows"))]
pub fn shell_execute_print(pdf_path: &std::path::Path, _printer_name: Option<&str>) -> Result<bool, String> {
    open::that(pdf_path).map(|_| true).map_err(|e| format!("无法打开 PDF: {}", e))
}
