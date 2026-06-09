//! PDFium automatic installer — downloads the platform-specific PDFium library
//! from GitHub releases on first run if it is missing from the `tools/` directory.
//!
//! Strategy:
//! 1. On startup, `ensure_pdfium(lib_dir)` checks if `pdfium.{dll|so|dylib}` exists.
//! 2. If missing, downloads the appropriate tarball from pdfium-binaries GitHub releases.
//! 3. Extracts the library to `<lib_dir>/`.
//!
//! The download is fire-and-forget at startup (logged but not blocking) so the
//! server can boot even when offline. The frontend may also call
//! `/api/v1/install_pdfium` to trigger an explicit installation with progress
//! reporting.

use std::fs::File;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU32, Ordering};

use crate::platform::{get_pdfium_download_url, get_pdfium_lib_name};

static INSTALL_PROGRESS: AtomicU32 = AtomicU32::new(0);

pub fn get_install_progress() -> u32 {
    INSTALL_PROGRESS.load(Ordering::SeqCst)
}

/// Spawn a background task to install PDFium if missing.
/// Logs progress; does not block the caller.
pub fn ensure_pdfium_background(lib_dir: PathBuf) {
    std::thread::spawn(move || {
        if let Err(e) = ensure_pdfium_blocking(&lib_dir) {
            log::warn!("PDFium 自动安装失败: {}。可手动从 {} 下载并解压到 {:?}",
                e, get_pdfium_download_url(), lib_dir);
        }
    });
}

/// Synchronous PDFium install — used by the explicit API endpoint.
pub fn ensure_pdfium_blocking(lib_dir: &Path) -> Result<PathBuf, String> {
    INSTALL_PROGRESS.store(0, Ordering::SeqCst);

    if !lib_dir.exists() {
        std::fs::create_dir_all(lib_dir)
            .map_err(|e| format!("创建 tools 目录失败: {}", e))?;
    }

    let lib_name = get_pdfium_lib_name();
    let target = lib_dir.join(lib_name);

    if target.exists() {
        INSTALL_PROGRESS.store(100, Ordering::SeqCst);
        return Ok(target);
    }

    log::info!("PDFium 缺失，开始下载: {} -> {:?}", get_pdfium_download_url(), target);

    let url = get_pdfium_download_url().to_string();
    let tmp_tar = lib_dir.join("_pdfium_download.tgz");

    download_with_progress(&url, &tmp_tar, |pct| {
        INSTALL_PROGRESS.store(pct / 2, Ordering::SeqCst);
    })?;

    INSTALL_PROGRESS.store(50, Ordering::SeqCst);

    extract_lib_from_tar(&tmp_tar, lib_name, &target, |pct| {
        INSTALL_PROGRESS.store(50 + pct / 2, Ordering::SeqCst);
    })?;

    let _ = std::fs::remove_file(&tmp_tar);

    INSTALL_PROGRESS.store(100, Ordering::SeqCst);
    log::info!("PDFium 安装完成: {:?}", target);

    Ok(target)
}

fn download_with_progress(url: &str, dest: &Path, mut on_progress: impl FnMut(u32)) -> Result<(), String> {
    let resp = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(300))
        .build()
        .map_err(|e| format!("创建 HTTP 客户端失败: {}", e))?
        .get(url)
        .send()
        .map_err(|e| format!("下载 PDFium 失败: {}", e))?;

    if !resp.status().is_success() {
        return Err(format!("HTTP 状态码: {}", resp.status()));
    }

    let total = resp.content_length().unwrap_or(0);
    let mut file = File::create(dest).map_err(|e| format!("创建临时文件失败: {}", e))?;
    let mut downloaded: u64 = 0;
    let mut last_pct: u32 = 0;
    let mut stream = resp;

    let mut buf = [0u8; 64 * 1024];
    loop {
        let n = stream.read(&mut buf).map_err(|e| format!("读取下载流失败: {}", e))?;
        if n == 0 { break; }
        file.write_all(&buf[..n]).map_err(|e| format!("写入下载文件失败: {}", e))?;
        downloaded += n as u64;
        if total > 0 {
            let pct = ((downloaded * 100) / total) as u32;
            if pct != last_pct {
                on_progress(pct);
                last_pct = pct;
            }
        }
    }
    if total == 0 {
        on_progress(100);
    }

    Ok(())
}

fn extract_lib_from_tar(tar_path: &Path, lib_name: &str, dest: &Path, mut on_progress: impl FnMut(u32)) -> Result<(), String> {
    let file = File::open(tar_path).map_err(|e| format!("打开下载文件失败: {}", e))?;
    let gz = flate2::read::GzDecoder::new(file);
    let mut archive = tar::Archive::new(gz);

    let total_entries = {
        let probe = tar::Archive::new(File::open(tar_path).map_err(|e| format!("打开下载文件失败: {}", e))?);
        let gz_probe = flate2::read::GzDecoder::new(probe.into_inner());
        let mut probe = tar::Archive::new(gz_probe);
        let mut count = 0u32;
        for _ in probe.entries().map_err(|e| format!("读取 tar 失败: {}", e))? {
            count += 1;
        }
        count.max(1)
    };

    let mut processed: u32 = 0;
    for entry in archive.entries().map_err(|e| format!("解压 tar 失败: {}", e))? {
        let mut entry = entry.map_err(|e| format!("读取 tar 条目失败: {}", e))?;
        let file_name = entry.path().map_err(|e| format!("读取 tar 路径失败: {}", e))?
            .file_name()
            .and_then(|n| n.to_str())
            .map(|s| s.to_string())
            .unwrap_or_default();

        if file_name == lib_name {
            let dest_clone = dest.to_path_buf();
            entry.unpack(dest_clone.as_path()).map_err(|e| format!("解压 PDFium 库失败: {}", e))?;
            log::info!("已解压 PDFium 库: {} -> {:?}", file_name, dest);
        }

        processed += 1;
        let pct = ((processed * 100) / total_entries) as u32;
        on_progress(pct);
    }

    if !dest.exists() {
        return Err(format!("解压完成但未找到 {}，请检查归档结构", lib_name));
    }

    Ok(())
}
