use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

use dashmap::DashMap;
use serde::Serialize;

/// A user session for the Web version.
/// Desktop mode uses a single implicit session with absolute file paths.
pub struct Session {
    pub id: String,
    pub dir: PathBuf,
    pub files: DashMap<String, SessionFile>,
    pub created_at: Instant,
    pub last_accessed: AtomicU64,
    pub total_size: AtomicU64,
}

impl Clone for Session {
    fn clone(&self) -> Self {
        Self {
            id: self.id.clone(),
            dir: self.dir.clone(),
            files: self.files.clone(),
            created_at: self.created_at,
            last_accessed: AtomicU64::new(self.last_accessed.load(Ordering::Relaxed)),
            total_size: AtomicU64::new(self.total_size.load(Ordering::Relaxed)),
        }
    }
}

#[derive(Clone)]
pub struct SessionFile {
    pub original_name: String,
    pub disk_path: PathBuf,
    pub file_type: String,
    pub size: u64,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FileInfo {
    pub name: String,
    pub path: String,
    pub file_type: String,
    pub size: u64,
}

impl Session {
    pub fn new(id: String, dir: PathBuf) -> Self {
        Self {
            id,
            dir,
            files: DashMap::new(),
            created_at: Instant::now(),
            last_accessed: AtomicU64::new(0),
            total_size: AtomicU64::new(0),
        }
    }

    /// Resolve a virtual path to an actual disk path.
    /// - Desktop mode (is_desktop=true): absolute paths pass through directly.
    /// - Web mode (is_desktop=false): only session-relative paths are allowed,
    ///   absolute paths are rejected to prevent path traversal attacks.
    pub fn resolve_path(&self, virtual_path: &str, is_desktop: bool) -> Result<PathBuf, String> {
        let path = Path::new(virtual_path);

        if is_desktop {
            // Desktop mode: absolute paths pass through (Tauri file dialog provides real paths)
            if path.is_absolute() {
                if path.exists() {
                    return Ok(path.to_path_buf());
                }
                return Err(format!("文件不存在: {}", virtual_path));
            }
            // Desktop with relative path: join with session dir (unusual but handle gracefully)
        } else {
            // Web mode: reject absolute paths to prevent path traversal
            if path.is_absolute() {
                return Err("Web 模式不允许使用绝对路径".to_string());
            }
        }

        // 1. Look up in the files map (takes priority over directory scan)
        if let Some(f) = self.files.get(virtual_path) {
            return Ok(f.disk_path.clone());
        }

        // 2. Join with session directory + canonicalize check (path traversal protection)
        let real = self.dir.join(virtual_path);
        if !real.exists() {
            return Err(format!("文件不存在: {}", virtual_path));
        }
        let canonical = real.canonicalize().map_err(|e| format!("路径无效: {}", e))?;
        let dir_canonical = self.dir.canonicalize().unwrap_or_else(|_| self.dir.clone());
        if !canonical.starts_with(&dir_canonical) {
            return Err("路径穿越攻击".into());
        }
        Ok(canonical)
    }

    /// Register an uploaded file to the session
    pub fn register_file(&self, original_name: &str, disk_path: PathBuf, size: u64, file_type: &str) {
        self.total_size.fetch_add(size, Ordering::Relaxed);
        self.files.insert(original_name.to_string(), SessionFile {
            original_name: original_name.to_string(),
            disk_path,
            file_type: file_type.to_string(),
            size,
        });
    }

    /// Touch the session (update last accessed timestamp)
    pub fn touch(&self) {
        self.last_accessed.store(
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .as_secs(),
            Ordering::Relaxed,
        );
    }
}

/// Determine file type from extension
pub fn detect_file_type(filename: &str) -> &'static str {
    let lower = filename.to_lowercase();
    if lower.ends_with(".pdf") {
        "pdf"
    } else if lower.ends_with(".ofd") {
        "ofd"
    } else if lower.ends_with(".xml") {
        "xml"
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        "image"
    } else if lower.ends_with(".png") {
        "image"
    } else if lower.ends_with(".bmp") {
        "image"
    } else if lower.ends_with(".tiff") || lower.ends_with(".tif") {
        "image"
    } else if lower.ends_with(".webp") {
        "image"
    } else {
        "unknown"
    }
}

/// Allowed file extensions for upload
pub fn is_allowed_extension(filename: &str) -> bool {
    let lower = filename.to_lowercase();
    const ALLOWED: &[&str] = &[
        ".pdf", ".ofd", ".xml",
        ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp",
    ];
    ALLOWED.iter().any(|ext| lower.ends_with(ext))
}
