use std::path::PathBuf;
use std::pin::Pin;
use std::sync::atomic::Ordering;
use std::sync::Arc;

use axum::{
    extract::{Multipart, Path as AxumPath, State},
    http::StatusCode,
    response::{IntoResponse, Response, Sse},
    routing::{get, post},
    Json, Router,
};
use dashmap::DashMap;
use serde::{Deserialize, Serialize};
use serde_json::json;
use tokio::sync::broadcast;
use tokio_stream::wrappers::BroadcastStream;
use tokio_stream::StreamExt;
use tower_http::cors::{Any, CorsLayer};
use tower_http::compression::CompressionLayer;
use tower_http::catch_panic::CatchPanicLayer;

use crate::session::{Session, FileInfo, detect_file_type, is_allowed_extension};

// === Unified API Response ===

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ApiResponse {
    ok: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    data: Option<serde_json::Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    code: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    recoverable: Option<bool>,
}

impl ApiResponse {
    fn ok(data: serde_json::Value) -> Self {
        Self { ok: true, data: Some(data), error: None, code: None, recoverable: None }
    }
    fn err(error: &str, code: &str, recoverable: bool) -> Self {
        Self { ok: false, data: None, error: Some(error.to_string()), code: Some(code.to_string()), recoverable: Some(recoverable) }
    }
}

// === AppError ===

#[derive(Debug)]
pub enum AppError {
    BadRequest(String),
    UnsupportedFormat(String),
    FileTooLarge(String),
    FileCorrupted(String),
    PasswordRequired(String),
    SessionExpired(String),
    RateLimited(String),
    Unauthorized(String),
    Internal(String),
    ServiceUnavailable(String),
}

impl IntoResponse for AppError {
    fn into_response(self) -> Response {
        let (status, code, msg, recoverable) = match &self {
            AppError::BadRequest(m) => (StatusCode::BAD_REQUEST, "BAD_REQUEST", m.as_str(), true),
            AppError::UnsupportedFormat(m) => (StatusCode::BAD_REQUEST, "UNSUPPORTED_FORMAT", m.as_str(), true),
            AppError::FileTooLarge(m) => (StatusCode::PAYLOAD_TOO_LARGE, "FILE_TOO_LARGE", m.as_str(), true),
            AppError::FileCorrupted(m) => (StatusCode::UNPROCESSABLE_ENTITY, "FILE_CORRUPTED", m.as_str(), false),
            AppError::PasswordRequired(m) => (StatusCode::UNPROCESSABLE_ENTITY, "OCR_PASSWORD_REQUIRED", m.as_str(), false),
            AppError::SessionExpired(m) => (StatusCode::GONE, "SESSION_EXPIRED", m.as_str(), true),
            AppError::RateLimited(m) => (StatusCode::TOO_MANY_REQUESTS, "RATE_LIMITED", m.as_str(), true),
            AppError::Unauthorized(m) => (StatusCode::UNAUTHORIZED, "UNAUTHORIZED", m.as_str(), false),
            AppError::Internal(m) => (StatusCode::INTERNAL_SERVER_ERROR, "INTERNAL", m.as_str(), false),
            AppError::ServiceUnavailable(m) => (StatusCode::SERVICE_UNAVAILABLE, "SERVICE_UNAVAILABLE", m.as_str(), true),
        };
        let body = ApiResponse::err(msg, code, recoverable);
        (status, Json(body)).into_response()
    }
}

// === Progress Tracker (SSE) ===

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ProgressEvent {
    pub task_id: String,
    pub phase: String,
    pub current: u32,
    pub total: u32,
}

pub struct ProgressTracker {
    channels: Arc<DashMap<String, broadcast::Sender<ProgressEvent>>>,
}

impl ProgressTracker {
    pub fn new() -> Self {
        Self { channels: Arc::new(DashMap::new()) }
    }

    pub fn create_task(&self) -> (String, broadcast::Sender<ProgressEvent>) {
        let task_id = uuid::Uuid::new_v4().to_string();
        let (tx, _) = broadcast::channel(16);
        self.channels.insert(task_id.clone(), tx.clone());
        (task_id, tx)
    }

    pub fn remove_task(&self, task_id: &str) {
        self.channels.remove(task_id);
    }

    pub fn get_sender(&self, task_id: &str) -> Option<broadcast::Sender<ProgressEvent>> {
        self.channels.get(task_id).map(|ch| ch.clone())
    }
}

// === AppState ===

#[derive(Clone)]
pub struct AppState {
    pub session_dir: PathBuf,
    pub sessions: Arc<DashMap<String, Session>>,
    pub ocr_limit: Arc<tokio::sync::Semaphore>,
    pub render_limit: Arc<tokio::sync::Semaphore>,
    pub auth_token: Option<String>,
    pub progress: Arc<ProgressTracker>,
    pub frontend_dir: PathBuf,
    pub is_desktop: bool,
}

// === Build Router ===

pub fn build_router(state: AppState) -> Router {
    let api_routes = Router::new()
        .route("/health", get(health))
        .route("/upload", post(upload_files))
        .route("/render_pdf", post(render_pdf))
        .route("/extract_pdf_text", post(extract_pdf_text))
        .route("/extract_pdf_texts", post(extract_pdf_texts))
        .route("/generate_pdf", post(generate_pdf))
        .route("/download/{session_id}/{filename}", get(download_file))
        .route("/progress/{task_id}", get(progress_sse))
        .route("/printers", get(list_printers))
        .route("/print", post(do_print))
        .route("/ocr_image", post(ocr_image))
        .route("/ocr_pdf_page", post(ocr_pdf_page))
        .route("/parse_ofd", post(parse_ofd))
        .route("/parse_xml_invoice", post(parse_xml_invoice))
        .route("/open_ofd_images", post(open_ofd_images))
        .route("/check_path_exists", post(check_path_exists))
        .route("/get_config", post(get_config))
        .route("/get_app_version", post(get_app_version))
        .route("/cancel_download", post(cancel_download))
        .route("/trim_image", post(trim_image))
        .route("/copy_file", post(copy_file))
        .route("/rename_file", post(rename_file))
        .route("/write_text_file", post(write_text_file))
        .route("/get_temp_dir", post(get_temp_dir))
        .route("/get_downloads_dir", post(get_downloads_dir))
        ;

    let cors = if state.is_desktop {
        CorsLayer::new()
            .allow_origin(Any)
            .allow_methods([axum::http::Method::GET, axum::http::Method::POST])
    } else {
        CorsLayer::permissive()
    };

    let mut app = Router::new()
        .nest("/api/v1", api_routes)
        .with_state(state.clone())
        .layer(CorsLayer::permissive())
        .layer(cors)
        .layer(CompressionLayer::new())
        .layer(CatchPanicLayer::new());

    if !state.is_desktop {
        app = app.fallback_service(
            tower_http::services::ServeDir::new(&state.frontend_dir)
                .append_index_html_on_directories(true)
        );
    }

    app
}

// === Handlers ===

async fn health() -> Json<ApiResponse> {
    Json(ApiResponse::ok(json!({
        "status": "ok",
        "pdfium": crate::pdfium_bindings::find_pdfium_lib().is_some(),
        "version": env!("CARGO_PKG_VERSION"),
    })))
}

async fn upload_files(
    State(state): State<AppState>,
    mut multipart: Multipart,
) -> Result<Json<ApiResponse>, AppError> {
    let session_id = uuid::Uuid::new_v4().to_string();
    let session_dir = state.session_dir.join(&session_id);
    std::fs::create_dir_all(&session_dir)
        .map_err(|e| AppError::Internal(format!("创建 session 目录失败: {}", e)))?;

    let session = Session::new(session_id.clone(), session_dir.clone());
    let mut file_infos = Vec::new();

    while let Some(field) = multipart.next_field().await
        .map_err(|e| AppError::BadRequest(format!("multipart 解析失败: {}", e)))?
    {
        let filename = field.file_name()
            .map(|n| n.to_string())
            .unwrap_or_else(|| "unknown".to_string());

        if !is_allowed_extension(&filename) {
            continue;
        }

        let data = field.bytes().await
            .map_err(|e| AppError::BadRequest(format!("读取上传数据失败: {}", e)))?;

        if data.is_empty() {
            continue;
        }

        let file_type = detect_file_type(&filename);
        let disk_name = format!("{}_{}", uuid::Uuid::new_v4(), filename);
        let disk_path = session_dir.join(&disk_name);

        std::fs::write(&disk_path, &data)
            .map_err(|e| AppError::Internal(format!("写入文件失败: {}", e)))?;

        let size = data.len() as u64;
        session.register_file(&filename, disk_path.clone(), size, file_type);

        file_infos.push(FileInfo {
            name: filename.clone(),
            path: filename, // Virtual path = original name
            file_type: file_type.to_string(),
            size,
        });
    }

    session.touch();
    state.sessions.insert(session_id.clone(), session);

    Ok(Json(ApiResponse::ok(json!({
        "sessionId": session_id,
        "files": file_infos,
    }))))
}

async fn render_pdf(
    State(state): State<AppState>,
    Json(req): Json<RenderPdfRequest>,
) -> Result<Json<ApiResponse>, AppError> {
    let session = get_session(&state, &req.session_id)?;
    let pdf_path = session.resolve_path(&req.pdf_path)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let use_jpeg = req.use_jpeg.unwrap_or(true);
    let dpi = req.dpi.unwrap_or(150);

    let _permit = state.render_limit.acquire().await
        .map_err(|e| AppError::Internal(format!("渲染信号量错误: {}", e)))?;

    let result = tokio::task::spawn_blocking(move || {
        crate::pdf_engine::render_pdf_pages_pdfium(
            &pdf_path.to_string_lossy(), dpi, use_jpeg,
        )
    }).await
        .map_err(|e| AppError::Internal(format!("渲染任务失败: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    Ok(Json(ApiResponse::ok(serde_json::to_value(result).unwrap_or_default())))
}

async fn extract_pdf_text(
    State(state): State<AppState>,
    Json(req): Json<ExtractPdfTextRequest>,
) -> Result<Json<ApiResponse>, AppError> {
    let session = get_session(&state, &req.session_id)?;
    let pdf_path = session.resolve_path(&req.pdf_path)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let result = tokio::task::spawn_blocking(move || {
        crate::pdf_engine::extract_pdf_text(&pdf_path.to_string_lossy(), req.page_index)
    }).await
        .map_err(|e| AppError::Internal(format!("提取任务失败: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    Ok(Json(ApiResponse::ok(serde_json::to_value(result).unwrap_or_default())))
}

async fn extract_pdf_texts(
    State(state): State<AppState>,
    Json(req): Json<ExtractPdfTextsRequest>,
) -> Result<Json<ApiResponse>, AppError> {
    let session = get_session(&state, &req.session_id)?;
    let pdf_path = session.resolve_path(&req.pdf_path)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let page_indices = req.page_indices.clone();
    let result = tokio::task::spawn_blocking(move || {
        crate::pdf_engine::extract_pdf_texts(&pdf_path.to_string_lossy(), &page_indices)
    }).await
        .map_err(|e| AppError::Internal(format!("批量提取任务失败: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    Ok(Json(ApiResponse::ok(serde_json::to_value(result).unwrap_or_default())))
}

async fn generate_pdf(
    State(_state): State<AppState>,
    Json(_req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    // TODO: Implement generate_pdf handler with SSE progress
    Err(AppError::Internal("generate_pdf 尚未实现".into()))
}

async fn download_file(
    State(state): State<AppState>,
    AxumPath((session_id, filename)): AxumPath<(String, String)>,
) -> Result<impl IntoResponse, AppError> {
    let session = get_session(&state, &session_id)?;
    let file_path = session.resolve_path(&filename)
        .map_err(|e| AppError::BadRequest(e))?;

    let data = std::fs::read(&file_path)
        .map_err(|e| AppError::Internal(format!("读取文件失败: {}", e)))?;

    let content_type = if filename.to_lowercase().ends_with(".pdf") {
        "application/pdf"
    } else if filename.to_lowercase().ends_with(".csv") {
        "text/csv"
    } else {
        "application/octet-stream"
    };

    Ok((
        StatusCode::OK,
        [
            ("content-type", content_type.to_string()),
            ("content-disposition", format!("attachment; filename=\"{}\"", filename)),
        ],
        data,
    ))
}

async fn progress_sse(
    AxumPath(task_id): AxumPath<String>,
    State(state): State<AppState>,
) -> Result<Sse<Pin<Box<dyn tokio_stream::Stream<Item = Result<axum::response::sse::Event, std::convert::Infallible>> + Send>>>, AppError> {
    let rx = state.progress.get_sender(&task_id)
        .map(|ch| ch.subscribe());

    let stream: Pin<Box<dyn tokio_stream::Stream<Item = Result<axum::response::sse::Event, std::convert::Infallible>> + Send>> = match rx {
        Some(rx) => {
            Box::pin(
                BroadcastStream::new(rx)
                    .filter_map(|msg| {
                        match msg {
                            Ok(event) => Some(Ok(axum::response::sse::Event::default()
                                .data(serde_json::to_string(&event).unwrap_or_default()))),
                            Err(_) => None,
                        }
                    })
            )
        }
        None => {
            Box::pin(tokio_stream::empty())
        }
    };

    Ok(Sse::new(stream))
}

async fn list_printers(
    State(_state): State<AppState>,
) -> Json<ApiResponse> {
    let printers = crate::platform::list_printers();
    Json(ApiResponse::ok(serde_json::to_value(printers).unwrap_or_default()))
}

async fn do_print(
    State(_state): State<AppState>,
    Json(_req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    // TODO: Implement print handler
    Err(AppError::Internal("打印功能尚未实现".into()))
}

async fn ocr_image(
    State(state): State<AppState>,
    Json(_req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let _permit = state.ocr_limit.acquire().await
        .map_err(|e| AppError::Internal(format!("OCR 信号量错误: {}", e)))?;

    #[cfg(feature = "ocr")]
    {
        let result = tokio::task::spawn_blocking(move || {
            // TODO: Call ocr-rs
            Err::<(), String>("OCR handler not yet implemented".into())
        }).await;
        match result {
            Ok(Ok(_)) => Ok(Json(ApiResponse::ok(json!({})))),
            Ok(Err(e)) => Err(AppError::Internal(e)),
            Err(e) => Err(AppError::Internal(format!("OCR 任务失败: {}", e))),
        }
    }
    #[cfg(not(feature = "ocr"))]
    {
        Err(AppError::ServiceUnavailable("OCR 功能未启用".into()))
    }
}

async fn ocr_pdf_page(
    State(state): State<AppState>,
    Json(_req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let _permit = state.ocr_limit.acquire().await
        .map_err(|e| AppError::Internal(format!("OCR 信号量错误: {}", e)))?;

    #[cfg(feature = "ocr")]
    {
        Err(AppError::Internal("OCR PDF 页面处理尚未实现".into()))
    }
    #[cfg(not(feature = "ocr"))]
    {
        Err(AppError::ServiceUnavailable("OCR 功能未启用".into()))
    }
}

async fn parse_ofd(
    Json(_req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    // TODO: Implement OFD parsing
    Err(AppError::Internal("OFD 解析尚未实现".into()))
}

async fn parse_xml_invoice(
    Json(_req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    // TODO: Implement XML invoice parsing
    Err(AppError::Internal("XML 数电票解析尚未实现".into()))
}

async fn open_ofd_images(
    Json(_req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    // TODO: Implement OFD image opening
    Err(AppError::Internal("OFD 图片打开尚未实现".into()))
}

async fn check_path_exists(
    Json(req): Json<serde_json::Value>,
) -> Json<ApiResponse> {
    let path = req.get("path").and_then(|v| v.as_str()).unwrap_or("");
    let exists = std::path::Path::new(path).exists();
    Json(ApiResponse::ok(json!({ "exists": exists })))
}

async fn get_config() -> Json<ApiResponse> {
    Json(ApiResponse::ok(json!({
        "version": env!("CARGO_PKG_VERSION"),
        "platform": std::env::consts::OS,
    })))
}

async fn get_app_version() -> Json<ApiResponse> {
    Json(ApiResponse::ok(json!({
        "version": env!("CARGO_PKG_VERSION"),
    })))
}

async fn cancel_download() -> Json<ApiResponse> {
    // Web version: PDFium is built-in, no download needed
    Json(ApiResponse::ok(json!({})))
}

async fn trim_image(
    Json(_req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    // TODO: Implement image trimming
    Err(AppError::Internal("图片裁剪尚未实现".into()))
}

async fn copy_file(
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let src = req.get("src").and_then(|v| v.as_str()).unwrap_or("");
    let dest = req.get("dest").and_then(|v| v.as_str()).unwrap_or("");
    std::fs::copy(src, dest)
        .map_err(|e| AppError::Internal(format!("复制文件失败: {}", e)))?;
    Ok(Json(ApiResponse::ok(json!({}))))
}

async fn rename_file(
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let src = req.get("src").and_then(|v| v.as_str()).unwrap_or("");
    let dest = req.get("dest").and_then(|v| v.as_str()).unwrap_or("");

    let result = if std::path::Path::new(src).parent() == std::path::Path::new(dest).parent() {
        std::fs::rename(src, dest)
    } else {
        std::fs::copy(src, dest).and_then(|_| std::fs::remove_file(src))
    };
    result.map_err(|e| AppError::Internal(format!("重命名文件失败: {}", e)))?;
    Ok(Json(ApiResponse::ok(json!({}))))
}

async fn write_text_file(
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let path = req.get("path").and_then(|v| v.as_str()).unwrap_or("");
    let content = req.get("content").and_then(|v| v.as_str()).unwrap_or("");
    std::fs::write(path, content)
        .map_err(|e| AppError::Internal(format!("写入文件失败: {}", e)))?;
    Ok(Json(ApiResponse::ok(json!({}))))
}

async fn get_temp_dir() -> Json<ApiResponse> {
    let temp = std::env::temp_dir().to_string_lossy().to_string();
    Json(ApiResponse::ok(json!({ "path": temp })))
}

async fn get_downloads_dir() -> Json<ApiResponse> {
    let downloads = dirs::download_dir()
        .map(|p| p.to_string_lossy().to_string())
        .unwrap_or_else(|| "/tmp".to_string());
    Json(ApiResponse::ok(json!({ "path": downloads })))
}

// === Helper ===

fn get_session(state: &AppState, session_id: &str) -> Result<Session, AppError> {
    // Desktop mode: create a default session if none exists
    if session_id.is_empty() {
        let default_id = "_desktop";
        if let Some(s) = state.sessions.get(default_id) {
            return Ok(s.clone());
        }
        let session = Session::new(default_id.to_string(), state.session_dir.clone());
        state.sessions.insert(default_id.to_string(), session.clone());
        return Ok(session);
    }
    state.sessions.get(session_id)
        .map(|r| r.clone())
        .ok_or_else(|| AppError::SessionExpired("会话已过期，请重新上传文件".into()))
}

// === Request types ===

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct RenderPdfRequest {
    session_id: String,
    pdf_path: String,
    dpi: Option<u32>,
    use_jpeg: Option<bool>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ExtractPdfTextRequest {
    session_id: String,
    pdf_path: String,
    page_index: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ExtractPdfTextsRequest {
    session_id: String,
    pdf_path: String,
    page_indices: Vec<u32>,
}

// === Session cleanup ===

pub async fn cleanup_sessions(state: &AppState, ttl_secs: u64) {
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();

    let mut to_remove = Vec::new();
    for entry in state.sessions.iter() {
        let last = entry.value().last_accessed.load(Ordering::Relaxed);
        if last > 0 && now - last > ttl_secs {
            to_remove.push(entry.key().clone());
        }
    }

    for id in to_remove {
        if let Some((_, session)) = state.sessions.remove(&id) {
            let _ = std::fs::remove_dir_all(&session.dir);
            log::info!("Cleaned up expired session: {}", id);
        }
    }
}

pub async fn cleanup_all_sessions(session_dir: &PathBuf) {
    if session_dir.exists() {
        if let Ok(entries) = std::fs::read_dir(session_dir) {
            for entry in entries.flatten() {
                if entry.file_type().map(|t| t.is_dir()).unwrap_or(false) {
                    let _ = std::fs::remove_dir_all(entry.path());
                }
            }
        }
    }
}
