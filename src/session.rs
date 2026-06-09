use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

use dashmap::DashMap;
use serde::Serialize;

/// A user session for the Web version.
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

    /// Resolve a virtual path to actual disk path.
    /// Rejects absolute paths for security (path traversal protection).
    pub fn resolve_path(&self, virtual_path: &str) -> Result<PathBuf, String> {
        let path = Path::new(virtual_path);

        if path.is_absolute() {
            return Err("不允许使用绝对路径".to_string());
        }

        // 1. Look up in the files map (takes priority over directory scan)
        if let Some(f) = self.files.get(virtual_path) {
            return Ok(f.disk_path.clone());
        }

        // 2. Join with session directory + canonicalize check
        let real = self.dir.join(virtual_path);
        let canonical = real.canonicalize()
            .map_err(|e| format!("文件路径无效: {} ({})", virtual_path, e))?;

        if !canonical.starts_with(&self.dir) {
            return Err("路径越权访问".to_string());
        }

        Ok(canonical)
    }

    pub fn register_file(&self, original_name: &str, disk_path: PathBuf, size: u64, file_type: &str) {
        self.total_size.fetch_add(size, Ordering::Relaxed);
        self.files.insert(original_name.to_string(), SessionFile {
            original_name: original_name.to_string(),
            disk_path,
            file_type: file_type.to_string(),
            size,
        });
    }

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

pub fn detect_file_type(filename: &str) -> &'static str {
    let lower = filename.to_lowercase();
    if lower.ends_with(".pdf") { "pdf" }
    else if lower.ends_with(".ofd") { "ofd" }
    else if lower.ends_with(".xml") { "xml" }
    else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") { "jpg" }
    else if lower.ends_with(".png") { "png" }
    else if lower.ends_with(".bmp") { "bmp" }
    else if lower.ends_with(".webp") { "webp" }
    else if lower.ends_with(".tiff") || lower.ends_with(".tif") { "tiff" }
    else { "other" }
}

pub fn is_allowed_extension(filename: &str) -> bool {
    matches!(detect_file_type(filename), "pdf" | "ofd" | "xml" | "jpg" | "png" | "bmp" | "webp" | "tiff")
}
