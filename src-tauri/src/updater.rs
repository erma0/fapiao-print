//! 应用内自更新（issue #58）：下载 → 校验 → 替换 exe → 自动重启。
//!
//! 只替换主程序 exe，不运行 NSIS 安装器：
//! - 轻量版 → 下载 `*_x64_portable.exe`（单文件，直接作为新 exe）
//! - OCR 版 → 下载 `*_x64_ocr-portable.zip`，取出内层 exe
//!
//! OCR 模型一两年内不会变化，因此不动 `<exe_dir>/models`；
//! 安装版（NSIS currentUser）与绿色版都落在用户可写目录，换文件无需 UAC。
//!
//! 安全约束：
//! - 下载地址必须命中 GitHub Release / gh-proxy 白名单；
//! - sha256（GitHub API 资产自带 digest）与 size 双重校验；
//! - 替换前先探测目录可写，任一步失败回滚旧 exe；
//! - 退出统一走 TerminateProcess（原因见 lib.rs 关闭逻辑的注释）。

use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};

use tauri::Emitter;

/// 独立于 PDFium 下载的取消标志，互不影响
pub static UPDATE_DOWNLOAD_CANCELLED: AtomicBool = AtomicBool::new(false);

/// 允许的下载来源：GitHub Release 直链、其 CDN 重定向、gh-proxy 加速镜像
const URL_ALLOWLIST: [&str; 3] = [
    "https://github.com/erma0/fapiao-print/releases/download/",
    "https://objects.githubusercontent.com/",
    "https://gh-proxy.com/https://github.com/erma0/fapiao-print/",
];

const PART_NAME: &str = "ticketchan.part";
const UPDATE_NAME: &str = "ticketchan.update";
const OLD_NAME: &str = "ticketchan.old.exe";
const TMP_EXE_NAME: &str = "ticketchan.update.tmp";
/// 主程序 exe 体积下限（正常 14~20MB，防止把错误内容当程序写入）
const MIN_EXE_SIZE: u64 = 1024 * 1024;

fn exe_path() -> Result<PathBuf, String> {
    std::env::current_exe().map_err(|e| format!("无法获取程序路径: {}", e))
}

fn exe_dir() -> Result<PathBuf, String> {
    exe_path()?
        .parent()
        .map(|p| p.to_path_buf())
        .ok_or_else(|| "无法获取程序目录".to_string())
}

/// 目录可写性探测：安装到 Program Files 等只读位置时直接给出可读错误，
/// 让前端回退到「前往 GitHub 手动下载」
fn ensure_writable(dir: &Path) -> Result<(), String> {
    let probe = dir.join("ticketchan.writetest");
    std::fs::write(&probe, b"1")
        .map_err(|e| format!("程序目录不可写（{}），请手动下载更新: {}", dir.display(), e))?;
    let _ = std::fs::remove_file(&probe);
    Ok(())
}

fn is_allowed_url(url: &str) -> bool {
    URL_ALLOWLIST.iter().any(|prefix| url.starts_with(prefix))
}

/// 归一化并校验 sha256：接受 `sha256:xxxx` 或裸 64 位十六进制
fn normalize_sha256(raw: &str) -> Option<String> {
    let v = raw.trim();
    let v = v.strip_prefix("sha256:").unwrap_or(v).to_ascii_lowercase();
    if v.len() == 64 && v.chars().all(|c| c.is_ascii_hexdigit()) {
        Some(v)
    } else {
        None
    }
}

/// 校验是有效的 Windows 可执行文件（PE 头 + 体积下限）
fn verify_exe(path: &Path) -> Result<(), String> {
    let len = std::fs::metadata(path)
        .map_err(|e| format!("读取更新文件失败: {}", e))?
        .len();
    if len < MIN_EXE_SIZE {
        return Err(format!("更新文件体积异常（{} 字节）", len));
    }
    let mut file = std::fs::File::open(path).map_err(|e| format!("打开更新文件失败: {}", e))?;
    let mut magic = [0u8; 2];
    file.read_exact(&mut magic)
        .map_err(|e| format!("读取更新文件失败: {}", e))?;
    if magic != *b"MZ" {
        return Err("更新文件不是有效的 Windows 程序".to_string());
    }
    Ok(())
}

/// 从 OCR 绿色版 zip 中取出内层 exe，写到指定路径
fn extract_exe_from_zip(zip_path: &Path, dest: &Path) -> Result<(), String> {
    let file = std::fs::File::open(zip_path).map_err(|e| format!("打开更新包失败: {}", e))?;
    let mut archive =
        zip::ZipArchive::new(file).map_err(|e| format!("解析更新包失败: {}", e))?;

    for i in 0..archive.len() {
        let mut entry = archive
            .by_index(i)
            .map_err(|e| format!("读取更新包条目失败: {}", e))?;
        if entry.is_dir() {
            continue;
        }
        let lower = entry.name().to_ascii_lowercase();
        if !lower.ends_with(".exe") {
            continue;
        }
        let mut out = std::fs::File::create(dest)
            .map_err(|e| format!("写出新程序失败: {}", e))?;
        std::io::copy(&mut entry, &mut out)
            .map_err(|e| format!("解压新程序失败: {}", e))?;
        return Ok(());
    }
    Err("更新包中未找到主程序 exe".to_string())
}

/// 下载更新包并校验，成功后落地为 `<exe_dir>/ticketchan.update`。
///
/// 前端按 GitHub API 的 `digest` 传 sha256；下载过程中每 64KB 发一次进度事件。
#[tauri::command]
pub async fn download_update(
    app: tauri::AppHandle,
    url: String,
    sha256: String,
    size: u64,
    is_zip: bool,
) -> Result<String, String> {
    if !is_allowed_url(&url) {
        return Err("更新地址不在受信任范围内".to_string());
    }
    let expected = normalize_sha256(&sha256).ok_or_else(|| "更新摘要无效".to_string())?;

    UPDATE_DOWNLOAD_CANCELLED.store(false, Ordering::SeqCst);
    let dir = exe_dir()?;
    ensure_writable(&dir)?;

    tauri::async_runtime::spawn_blocking(move || -> Result<String, String> {
        use sha2::{Digest, Sha256};

        let part = dir.join(PART_NAME);
        let update = dir.join(UPDATE_NAME);
        let tmp_exe = dir.join(TMP_EXE_NAME);
        let _ = std::fs::remove_file(&part);
        let _ = std::fs::remove_file(&update);
        let _ = std::fs::remove_file(&tmp_exe);

        let agent = crate::build_http_agent(std::time::Duration::from_secs(300))?;
        let resp = match agent.get(&url).call() {
            Ok(r) => r,
            Err(ureq::Error::Status(code, _)) => {
                return Err(format!("下载失败，HTTP 状态: {}", code));
            }
            Err(e) => return Err(format!("下载失败: {}", e)),
        };

        let total = if size > 0 {
            size
        } else {
            resp.header("Content-Length")
                .and_then(|v| v.parse::<u64>().ok())
                .unwrap_or(0)
        };

        let mut file =
            std::fs::File::create(&part).map_err(|e| format!("创建临时文件失败: {}", e))?;
        let mut reader = resp.into_reader();
        let mut hasher = Sha256::new();
        let mut buf = [0u8; 64 * 1024];
        let mut downloaded: u64 = 0;

        loop {
            let n = reader
                .read(&mut buf)
                .map_err(|e| format!("下载出错: {}", e))?;
            if n == 0 {
                break;
            }
            if UPDATE_DOWNLOAD_CANCELLED.load(Ordering::SeqCst) {
                drop(file);
                drop(reader);
                std::fs::remove_file(&part).ok();
                return Err("下载已取消".to_string());
            }
            file.write_all(&buf[..n])
                .map_err(|e| format!("写入文件失败: {}", e))?;
            hasher.update(&buf[..n]);
            downloaded += n as u64;
            let _ = app.emit(
                "update-download-progress",
                serde_json::json!({
                    "current": downloaded,
                    "total": total,
                    "percent": if total > 0 { (downloaded as f64 / total as f64) * 100.0 } else { 0.0 }
                }),
            );
        }
        drop(file);
        drop(reader);

        if downloaded == 0 {
            std::fs::remove_file(&part).ok();
            return Err("下载内容为空".to_string());
        }
        if size > 0 && downloaded != size {
            std::fs::remove_file(&part).ok();
            return Err(format!(
                "文件大小与发布信息不一致（{}/{} 字节）",
                downloaded, size
            ));
        }
        let actual = format!("{:x}", hasher.finalize());
        if actual != expected {
            std::fs::remove_file(&part).ok();
            return Err("文件校验失败（sha256 不匹配），请重新下载".to_string());
        }

        if is_zip {
            extract_exe_from_zip(&part, &tmp_exe)?;
            let _ = std::fs::remove_file(&part);
            verify_exe(&tmp_exe)?;
            std::fs::rename(&tmp_exe, &update)
                .map_err(|e| format!("保存更新文件失败: {}", e))?;
        } else {
            verify_exe(&part)?;
            std::fs::rename(&part, &update)
                .map_err(|e| format!("保存更新文件失败: {}", e))?;
        }

        log::info!("update downloaded: {}", update.display());
        Ok(update.to_string_lossy().to_string())
    })
    .await
    .map_err(|e| format!("更新下载任务失败: {}", e))?
}

/// 用已下载的 `ticketchan.update` 替换当前 exe 并启动新版本。
///
/// 成功后本进程在 300ms 内被 TerminateProcess 结束；失败时回滚旧 exe，
/// `ticketchan.update` 保留，前端提示「下次启动自动完成」。
#[tauri::command]
pub fn apply_update_and_restart(app: tauri::AppHandle) -> Result<(), String> {
    let exe = exe_path()?;
    let dir = exe_dir()?;
    let update = dir.join(UPDATE_NAME);
    if !update.exists() {
        return Err("未找到已下载的更新文件，请重新下载".to_string());
    }
    ensure_writable(&dir)?;

    // 先隐藏窗口，避免用户看到文件替换过程
    {
        use tauri::Manager;
        if let Some(window) = app.get_webview_window("main") {
            let _ = window.hide();
        }
    }

    swap_and_launch(&exe, &update)?;
    schedule_exit();
    Ok(())
}

/// 启动时补偿：清理上次残留，若有待安装更新则立即应用（失败则丢弃，绝不重启循环）
#[tauri::command]
pub fn check_pending_update() -> Result<String, String> {
    let dir = exe_dir()?;
    let _ = std::fs::remove_file(dir.join(PART_NAME));
    let _ = std::fs::remove_file(dir.join(TMP_EXE_NAME));
    // 上一次成功升级后留下的旧版本备份，此时已不再被占用
    let _ = std::fs::remove_file(dir.join(OLD_NAME));

    let update = dir.join(UPDATE_NAME);
    if !update.exists() {
        return Ok("none".to_string());
    }

    let exe = exe_path()?;
    match swap_and_launch(&exe, &update) {
        Ok(()) => {
            schedule_exit();
            Ok("applied".to_string())
        }
        Err(e) => {
            let _ = std::fs::remove_file(&update);
            Err(format!("待安装的更新无法应用，已忽略: {}", e))
        }
    }
}

/// 取消进行中的更新下载
#[tauri::command]
pub fn cancel_update_download() {
    UPDATE_DOWNLOAD_CANCELLED.store(true, Ordering::SeqCst);
}

/// 替换 exe 并启动新进程；任一步失败都回滚到旧 exe。
///
/// Windows 允许重命名正在运行的 exe（不允许删除），所以流程是：
/// 当前 exe → `.old`，`.update` → 当前 exe 原名，再启动。
fn swap_and_launch(exe: &Path, update: &Path) -> Result<(), String> {
    let dir = exe.parent().ok_or_else(|| "无法获取程序目录".to_string())?;

    let mut old = dir.join(OLD_NAME);
    if old.exists() && std::fs::remove_file(&old).is_err() {
        // 旧备份被占用时换一个名字，避免阻断升级
        old = dir.join(format!("ticketchan.old.{}.exe", std::process::id()));
        let _ = std::fs::remove_file(&old);
    }

    std::fs::rename(exe, &old)
        .map_err(|e| format!("无法替换程序文件（可能被占用）: {}", e))?;

    if let Err(e) = std::fs::rename(update, exe) {
        let _ = std::fs::rename(&old, exe);
        return Err(format!("写入新程序失败: {}", e));
    }

    if let Err(e) = std::process::Command::new(exe).current_dir(dir).spawn() {
        // 回滚：新程序退回 .update，旧程序复位
        let _ = std::fs::rename(exe, update);
        let _ = std::fs::rename(&old, exe);
        return Err(format!("启动新版本失败: {}", e));
    }

    Ok(())
}

/// 300ms 后结束当前进程，把运行权交给刚启动的新版本。
///
/// 与关闭流程一致使用 TerminateProcess：ExitProcess 会在
/// DLL_PROCESS_DETACH 阶段卡死（MNN/OCR 引擎持锁）。
fn schedule_exit() {
    std::thread::spawn(|| {
        std::thread::sleep(std::time::Duration::from_millis(300));
        #[cfg(target_os = "windows")]
        unsafe {
            use windows::Win32::System::Threading::{GetCurrentProcess, TerminateProcess};
            let _ = TerminateProcess(GetCurrentProcess(), 0);
        }
        #[cfg(not(target_os = "windows"))]
        std::process::exit(0);
    });
}
