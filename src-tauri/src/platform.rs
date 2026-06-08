/// Platform abstraction layer — the only file that should contain #[cfg(target_os)]
/// All platform-specific code extracted from pdf_engine.rs / lib.rs is consolidated here.

use std::path::PathBuf;

// === PDFium library loading ===

#[cfg(target_os = "windows")]
pub fn load_pdfium_lib() -> Result<libloading::Library, String> {
    if let Ok(exe) = std::env::current_exe() {
        if let Some(parent) = exe.parent() {
            let dll_path = parent.join("tools").join("pdfium.dll");
            if dll_path.exists() {
                log::info!("Loading pdfium.dll from: {}", dll_path.display());
                return unsafe {
                    libloading::Library::new(&dll_path)
                        .map_err(|e| format!("加载 pdfium.dll 失败: {}", e))
                };
            }
        }
    }
    unsafe {
        libloading::Library::new("pdfium.dll")
            .map_err(|e| format!("未找到 pdfium.dll，请将 pdfium.dll 放到 tools 目录下: {}", e))
    }
}

#[cfg(target_os = "linux")]
pub fn load_pdfium_lib() -> Result<libloading::Library, String> {
    // 1. Executable's tools/ directory
    if let Ok(exe) = std::env::current_exe() {
        if let Some(parent) = exe.parent() {
            let so_path = parent.join("tools").join("libpdfium.so");
            if so_path.exists() {
                log::info!("Loading libpdfium.so from: {}", so_path.display());
                return unsafe {
                    libloading::Library::new(&so_path)
                        .map_err(|e| format!("加载 libpdfium.so 失败: {}", e))
                };
            }
        }
    }
    // 2. System ldconfig paths
    let names = ["libpdfium.so", "libpdfium.so.1"];
    for name in &names {
        if let Ok(lib) = unsafe { libloading::Library::new(name) } {
            log::info!("Loading {} from system path", name);
            return Ok(lib);
        }
    }
    Err("未找到 libpdfium.so，可点击「下载 PDFium」自动安装，或手动运行: apt install pdfium-binaries".into())
}

#[cfg(target_os = "macos")]
pub fn load_pdfium_lib() -> Result<libloading::Library, String> {
    // 1. Executable's tools/ directory
    if let Ok(exe) = std::env::current_exe() {
        if let Some(parent) = exe.parent() {
            let dylib_path = parent.join("tools").join("libpdfium.dylib");
            if dylib_path.exists() {
                log::info!("Loading libpdfium.dylib from: {}", dylib_path.display());
                return unsafe {
                    libloading::Library::new(&dylib_path)
                        .map_err(|e| format!("加载 libpdfium.dylib 失败: {}", e))
                };
            }
        }
    }
    // 2. Common library paths
    let paths = ["/usr/local/lib/libpdfium.dylib", "/opt/homebrew/lib/libpdfium.dylib"];
    for path in &paths {
        if std::path::Path::new(path).exists() {
            log::info!("Loading libpdfium.dylib from: {}", path);
            return unsafe {
                libloading::Library::new(*path)
                    .map_err(|e| format!("加载 libpdfium.dylib 失败: {}", e))
            };
        }
    }
    Err("未找到 libpdfium.dylib，可点击「下载 PDFium」自动安装".into())
}

// === PDFium library discovery ===

#[cfg(target_os = "windows")]
pub fn find_pdfium_lib() -> Option<PathBuf> {
    if let Ok(exe) = std::env::current_exe() {
        if let Some(parent) = exe.parent() {
            let dll_path = parent.join("tools").join("pdfium.dll");
            if dll_path.exists() {
                return Some(dll_path);
            }
        }
    }
    None
}

#[cfg(target_os = "linux")]
pub fn find_pdfium_lib() -> Option<PathBuf> {
    if let Ok(exe) = std::env::current_exe() {
        if let Some(parent) = exe.parent() {
            let so_path = parent.join("tools").join("libpdfium.so");
            if so_path.exists() {
                return Some(so_path);
            }
        }
    }
    // Check system paths
    for name in &["libpdfium.so", "libpdfium.so.1"] {
        if let Ok(lib) = unsafe { libloading::Library::new(*name) } {
            drop(lib);
            return Some(PathBuf::from(*name));
        }
    }
    None
}

#[cfg(target_os = "macos")]
pub fn find_pdfium_lib() -> Option<PathBuf> {
    if let Ok(exe) = std::env::current_exe() {
        if let Some(parent) = exe.parent() {
            let dylib_path = parent.join("tools").join("libpdfium.dylib");
            if dylib_path.exists() {
                return Some(dylib_path);
            }
        }
    }
    for path in &["/usr/local/lib/libpdfium.dylib", "/opt/homebrew/lib/libpdfium.dylib"] {
        if std::path::Path::new(path).exists() {
            return Some(PathBuf::from(*path));
        }
    }
    None
}

// === PDFium download helpers ===

#[cfg(target_os = "windows")]
pub fn get_pdfium_download_url() -> &'static str {
    "https://github.com/bblanchon/pdfium-binaries/releases/download/chromium%2F6963/pdfium-win-x64.tgz"
}

#[cfg(target_os = "linux")]
pub fn get_pdfium_download_url() -> &'static str {
    "https://github.com/bblanchon/pdfium-binaries/releases/download/chromium%2F6963/pdfium-linux-x64.tgz"
}

#[cfg(target_os = "macos")]
pub fn get_pdfium_download_url() -> &'static str {
    "https://github.com/bblanchon/pdfium-binaries/releases/download/chromium%2F6963/pdfium-mac-x64.tgz"
}

#[cfg(target_os = "windows")]
pub fn get_pdfium_lib_name() -> &'static str {
    "pdfium.dll"
}

#[cfg(target_os = "linux")]
pub fn get_pdfium_lib_name() -> &'static str {
    "libpdfium.so"
}

#[cfg(target_os = "macos")]
pub fn get_pdfium_lib_name() -> &'static str {
    "libpdfium.dylib"
}

pub fn get_pdfium_lib_dir() -> Option<PathBuf> {
    std::env::current_exe().ok().and_then(|exe| exe.parent().map(|p| p.join("tools")))
}

// === System detection ===

#[cfg(target_os = "windows")]
pub fn check_windows_version() -> Result<(), String> {
    use windows::core::*;
    use windows::Win32::System::Registry::*;

    unsafe {
        let mut hkey = HKEY::default();
        let result = RegOpenKeyExW(
            HKEY_LOCAL_MACHINE,
            w!("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"),
            0,
            KEY_READ,
            &mut hkey,
        );

        if result.is_ok() {
            let mut build_number = [0u16; 256];
            let mut build_number_size = (build_number.len() * 2) as u32;

            let result = RegQueryValueExW(
                hkey,
                w!("CurrentBuildNumber"),
                None,
                None,
                Some(build_number.as_mut_ptr() as *mut u8),
                Some(&mut build_number_size),
            );

            let _ = RegCloseKey(hkey);

            if result.is_ok() {
                let build_str = String::from_utf16_lossy(&build_number[..(build_number_size as usize / 2)]);
                let build_str = build_str.trim_end_matches('\0');

                if let Ok(build) = build_str.parse::<u32>() {
                    if build < 17134 {
                        return Err(format!(
                            "您的系统版本不支持本应用。\n\n当前系统：Windows (Build {})\n\n需要：Windows 10 1803 (Build 17134) 或 Windows 11",
                            build
                        ));
                    }
                    return Ok(());
                }
            }
        }

        Ok(())
    }
}

#[cfg(not(target_os = "windows"))]
pub fn check_windows_version() -> Result<(), String> {
    Ok(())
}

// === WinRT PDF detection ===

#[cfg(target_os = "windows")]
pub fn has_winrt_pdf() -> bool {
    crate::pdf_engine::check_winrt_pdf_available()
}

#[cfg(not(target_os = "windows"))]
pub fn has_winrt_pdf() -> bool {
    false
}

// === Printer listing ===

#[derive(Debug, Clone, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PrinterInfo {
    pub name: String,
    pub is_default: bool,
}

#[cfg(target_os = "windows")]
pub fn list_printers() -> Vec<PrinterInfo> {
    use windows::Win32::Graphics::Printing::{EnumPrintersW, PRINTER_ENUM_LOCAL, PRINTER_ENUM_CONNECTIONS, PRINTER_INFO_4W};
    use windows::core::PCWSTR;

    let default_name = get_default_printer_name();

    unsafe {
        let mut bytes_needed: u32 = 0;
        let mut count_returned: u32 = 0;
        let flags = PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS;
        let null_name = PCWSTR::null();

        let _ = EnumPrintersW(flags, null_name, 4, None, &mut bytes_needed, &mut count_returned);
        if bytes_needed == 0 {
            return vec![];
        }

        let mut buffer: Vec<u8> = vec![0u8; bytes_needed as usize];
        if EnumPrintersW(
            flags, null_name, 4,
            Some(&mut buffer),
            &mut bytes_needed, &mut count_returned,
        ).is_err() {
            return vec![];
        }

        let ptr = buffer.as_ptr() as *const PRINTER_INFO_4W;
        let mut result = Vec::with_capacity(count_returned as usize);

        for i in 0..count_returned {
            let info = &*ptr.offset(i as isize);
            if info.pPrinterName.is_null() {
                continue;
            }
            let ptr = info.pPrinterName.0;
            let len = (0..).take_while(|&j| *ptr.offset(j) != 0).count();
            let name = String::from_utf16_lossy(std::slice::from_raw_parts(ptr, len));
            let is_default = default_name.as_ref().map_or(false, |dn| dn.eq_ignore_ascii_case(&name));
            result.push(PrinterInfo { name, is_default });
        }

        result
    }
}

#[cfg(target_os = "windows")]
pub fn get_default_printer_name() -> Option<String> {
    use windows::Win32::Graphics::Printing::GetDefaultPrinterW;
    use windows::core::PWSTR;

    unsafe {
        let mut size: u32 = 0;
        let _ = GetDefaultPrinterW(PWSTR::null(), &mut size);
        if size == 0 {
            return None;
        }
        let mut buf = vec![0u16; size as usize];
        let result = GetDefaultPrinterW(PWSTR(buf.as_mut_ptr()), &mut size);
        if result.as_bool() && size > 0 {
            let len = buf.iter().position(|&c| c == 0).unwrap_or(size as usize);
            if len > 0 {
                return Some(String::from_utf16_lossy(&buf[..len]));
            }
        }
        None
    }
}

#[cfg(not(target_os = "windows"))]
pub fn list_printers() -> Vec<PrinterInfo> {
    // Use CUPS lpstat to enumerate printers on Linux/macOS
    let output = std::process::Command::new("lpstat")
        .arg("-v")
        .output();
    match output {
        Ok(out) if out.status.success() => {
            let text = String::from_utf8_lossy(&out.stdout);
            text.lines()
                .filter_map(|line| {
                    // Format: "device for PRINTER_NAME: uri"
                    let prefix = "device for ";
                    if let Some(pos) = line.find(prefix) {
                        let rest = &line[pos + prefix.len()..];
                        if let Some(colon) = rest.find(':') {
                            let name = rest[..colon].trim().to_string();
                            return Some(PrinterInfo {
                                name,
                                is_default: false,
                            });
                        }
                    }
                    None
                })
                .collect()
        }
        _ => Vec::new(),
    }
}

#[cfg(not(target_os = "windows"))]
pub fn get_default_printer_name() -> Option<String> {
    // Use lpstat -d to get the default printer on Linux/macOS
    let output = std::process::Command::new("lpstat")
        .arg("-d")
        .output();
    match output {
        Ok(out) if out.status.success() => {
            let text = String::from_utf8_lossy(&out.stdout);
            // Format: "system default destination: PRINTER_NAME"
            let prefix = "system default destination: ";
            if let Some(pos) = text.find(prefix) {
                let name = text[pos + prefix.len()..].trim().to_string();
                if !name.is_empty() {
                    return Some(name);
                }
            }
            None
        }
        _ => None,
    }
}
