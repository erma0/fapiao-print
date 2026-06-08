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
        .layer(cors)
        .layer(tower_http::limit::RequestBodyLimitLayer::new(100 * 1024 * 1024)) // 100MB
        .layer(CompressionLayer::new())
        .layer(CatchPanicLayer::new());

    // Optional auth token protection for Web (non-desktop) deployments.
    // When TICKETCHAN_AUTH_TOKEN is set, all API requests must include
    // the Authorization: Bearer <token> header. Health endpoint is exempt.
    // Desktop mode (is_desktop=true) skips auth entirely.
    if !state.is_desktop && state.auth_token.is_some() {
        let expected_token = state.auth_token.clone().unwrap();
        app = app.layer(axum::middleware::from_fn(move |req: axum::http::Request<axum::body::Body>, next: axum::middleware::Next| {
            let token = expected_token.clone();
            async move {
                if req.uri().path() == "/api/v1/health" {
                    return next.run(req).await;
                }
                let auth_header = req.headers()
                    .get(axum::http::header::AUTHORIZATION)
                    .and_then(|v| v.to_str().ok())
                    .and_then(|v| v.strip_prefix("Bearer "));
                if auth_header != Some(&token) {
                    return axum::response::Response::builder()
                        .status(StatusCode::UNAUTHORIZED)
                        .header("content-type", "application/json")
                        .body(axum::body::Body::from(
                            serde_json::to_string(&ApiResponse::err(
                                "未授权访问", "UNAUTHORIZED", false,
                            )).unwrap_or_default()
                        ))
                        .unwrap();
                }
                next.run(req).await
            }
        }));
    }

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
    let pdf_path = session.resolve_path(&req.pdf_path, state.is_desktop)
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

    Ok(Json(ApiResponse::ok(serde_json::to_value(&result).unwrap_or_else(|e| {
        log::warn!("序列化渲染结果失败: {}", e);
        json!({})
    }))))
}

async fn extract_pdf_text(
    State(state): State<AppState>,
    Json(req): Json<ExtractPdfTextRequest>,
) -> Result<Json<ApiResponse>, AppError> {
    let session = get_session(&state, &req.session_id)?;
    let pdf_path = session.resolve_path(&req.pdf_path, state.is_desktop)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let result = tokio::task::spawn_blocking(move || {
        crate::pdf_engine::extract_pdf_text(&pdf_path.to_string_lossy(), req.page_index)
    }).await
        .map_err(|e| AppError::Internal(format!("提取任务失败: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    Ok(Json(ApiResponse::ok(serde_json::to_value(&result).unwrap_or_else(|e| {
        log::warn!("序列化文本提取结果失败: {}", e);
        json!({})
    }))))
}

async fn extract_pdf_texts(
    State(state): State<AppState>,
    Json(req): Json<ExtractPdfTextsRequest>,
) -> Result<Json<ApiResponse>, AppError> {
    let session = get_session(&state, &req.session_id)?;
    let pdf_path = session.resolve_path(&req.pdf_path, state.is_desktop)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let page_indices = req.page_indices.clone();
    let result = tokio::task::spawn_blocking(move || {
        crate::pdf_engine::extract_pdf_texts(&pdf_path.to_string_lossy(), &page_indices)
    }).await
        .map_err(|e| AppError::Internal(format!("批量提取任务失败: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    Ok(Json(ApiResponse::ok(serde_json::to_value(&result).unwrap_or_else(|e| {
        log::warn!("序列化批量提取结果失败: {}", e);
        json!({})
    }))))
}

async fn generate_pdf(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let session_id = req.get("sessionId").and_then(|v| v.as_str()).unwrap_or("").to_string();
    let session = get_session(&state, &session_id)?;
    session.touch();

    let (task_id, tx) = state.progress.create_task();
    let tid_progress = task_id.clone();
    let tid_file = task_id.clone();
    let tx_clone = tx.clone();
    let progress_cb: crate::pdf_engine::ProgressFn = Box::new(move |phase, current, total| {
        let _ = tx_clone.send(crate::server::ProgressEvent {
            task_id: tid_progress.clone(),
            phase: phase.to_string(),
            current,
            total,
        });
    });

    let req_value = req.clone();
    let result = tokio::task::spawn_blocking(move || {
        let output_path = std::env::temp_dir().join(format!("ticketchan-gen-{}.pdf", tid_file));
        crate::pdf_engine::generate_pdf_from_layout(
            &serde_json::from_value(req_value).map_err(|e| format!("请求解析失败: {}", e))?,
            &output_path,
            Some(progress_cb),
        )
    }).await
        .map_err(|e| AppError::Internal(format!("生成任务崩溃: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    // generate_pdf_from_layout returns Option<String> (None = nothing to generate)
    let output_file = result.ok_or_else(|| AppError::BadRequest("无有效页面数据，PDF 未生成".into()))?;
    let output_path = std::path::PathBuf::from(&output_file);

    // Register generated file into session
    let file_size = std::fs::metadata(&output_path).map(|m| m.len()).unwrap_or(0);
    let file_name = format!("{}.pdf", task_id);
    session.register_file(&file_name, output_path.clone(), file_size, "pdf");

    state.progress.remove_task(&task_id);

    Ok(Json(ApiResponse::ok(json!({
        "taskId": task_id,
        "path": output_path.to_string_lossy(),
        "fileName": file_name,
        "sessionId": session.id,
    }))))
}

async fn download_file(
    State(state): State<AppState>,
    AxumPath((session_id, filename)): AxumPath<(String, String)>,
) -> Result<impl IntoResponse, AppError> {
    let session = get_session(&state, &session_id)?;
    let file_path = session.resolve_path(&filename, state.is_desktop)
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

    // Note: keep_alive is handled at the nginx/reverse-proxy layer (proxy_read_timeout).
    // Adding keep_alive here changes the Sse type signature and would require boxing.
    Ok(Sse::new(stream))
}

async fn list_printers(
    State(_state): State<AppState>,
) -> Json<ApiResponse> {
    let printers = crate::platform::list_printers();
    Json(ApiResponse::ok(serde_json::to_value(printers).unwrap_or_default()))
}

async fn do_print(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let session_id = req.get("sessionId").and_then(|v| v.as_str()).unwrap_or("").to_string();
    let session = get_session(&state, &session_id)?;
    session.touch();

    let pdf_path = req.get("pdfPath").and_then(|v| v.as_str())
        .or_else(|| req.get("path").and_then(|v| v.as_str()))
        .unwrap_or("");
    let printer_name = req.get("printer").and_then(|v| v.as_str());
    let resolved = session.resolve_path(pdf_path, state.is_desktop)
        .map_err(|e| AppError::BadRequest(e))?;

    let path_str = resolved.to_string_lossy().to_string();
    let printer_opt = printer_name.map(|s| s.to_string());

    #[cfg(target_os = "windows")]
    {
        let result = tokio::task::spawn_blocking(move || {
            crate::shell_execute_print(
                std::path::Path::new(&path_str),
                printer_opt.as_deref(),
            )
        }).await
            .map_err(|e| AppError::Internal(format!("打印任务崩溃: {}", e)))?
            .map_err(|e| AppError::Internal(e))?;
        Ok(Json(ApiResponse::ok(json!({ "success": result }))))
    }
    #[cfg(not(target_os = "windows"))]
    {
        // On non-Windows, open the PDF with the system default viewer
        let _ = tokio::task::spawn_blocking(move || {
            open::that(&path_str)
        }).await;
        Ok(Json(ApiResponse::ok(json!({ "success": true, "message": "已使用系统默认程序打开 PDF" }))))
    }
}

async fn ocr_image(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    #[allow(unused_variables)]
    {
    let session_id = req.get("sessionId").and_then(|v| v.as_str()).unwrap_or("").to_string();
    let session = get_session(&state, &session_id)?;
    session.touch();

    let data_url = req.get("dataUrl").and_then(|v| v.as_str()).unwrap_or("");
    let file_path = req.get("filePath").and_then(|v| v.as_str());
    let ocr_precision = req.get("ocrPrecision").and_then(|v| v.as_str());

    if data_url.is_empty() {
        return Err(AppError::BadRequest("缺少图片数据".into()));
    }

    let _permit = state.ocr_limit.acquire().await
        .map_err(|e| AppError::Internal(format!("OCR 信号量错误: {}", e)))?;

    #[cfg(feature = "ocr")]
    {
        let data = data_url.to_string();
        let fp = file_path.map(|s| s.to_string());
        let precision = ocr_precision.map(|s| s.to_string());
        let result = tokio::task::spawn_blocking(move || {
            crate::pdf_engine::ocr_image(&data, fp.as_deref(), precision.as_deref())
                .map_err(|e| format!("OCR 识别失败: {}", e))
        }).await
            .map_err(|e| AppError::Internal(format!("OCR 任务崩溃: {}", e)))?
            .map_err(|e| AppError::Internal(e))?;

        Ok(Json(ApiResponse::ok(serde_json::to_value(&result).unwrap_or_else(|e| {
            log::warn!("序列化 OCR 结果失败: {}", e);
            json!({})
        }))))
    }
    #[cfg(not(feature = "ocr"))]
    {
        Err(AppError::ServiceUnavailable("OCR 功能未启用".into()))
    }
    } // end #[allow(unused_variables)]
}

async fn ocr_pdf_page(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    #[allow(unused_variables)]
    {
    let session_id = req.get("sessionId").and_then(|v| v.as_str()).unwrap_or("").to_string();
    let session = get_session(&state, &session_id)?;
    let pdf_path = req.get("pdfPath").and_then(|v| v.as_str()).unwrap_or("");
    let page_index = req.get("pageIndex").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
    let dpi = req.get("dpi").and_then(|v| v.as_u64()).map(|v| v as u32);
    let ocr_precision = req.get("ocrPrecision").and_then(|v| v.as_str());

    let resolved = session.resolve_path(pdf_path, state.is_desktop)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let _permit = state.ocr_limit.acquire().await
        .map_err(|e| AppError::Internal(format!("OCR 信号量错误: {}", e)))?;

    #[cfg(feature = "ocr")]
    {
        let path = resolved.to_string_lossy().to_string();
        let precision = ocr_precision.map(|s| s.to_string());
        let result = tokio::task::spawn_blocking(move || {
            crate::pdf_engine::ocr_pdf_page(&path, page_index, dpi, precision.as_deref())
                .map_err(|e| format!("OCR PDF 页面识别失败: {}", e))
        }).await
            .map_err(|e| AppError::Internal(format!("OCR 任务崩溃: {}", e)))?
            .map_err(|e| AppError::Internal(e))?;

        Ok(Json(ApiResponse::ok(serde_json::to_value(&result).unwrap_or_else(|e| {
            log::warn!("序列化 OCR PDF 结果失败: {}", e);
            json!({})
        }))))
    }
    #[cfg(not(feature = "ocr"))]
    {
        Err(AppError::ServiceUnavailable("OCR 功能未启用".into()))
    }
    } // end #[allow(unused_variables)]
}

async fn parse_ofd(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let session_id = req.get("sessionId").and_then(|v| v.as_str()).unwrap_or("").to_string();
    let session = get_session(&state, &session_id)?;
    let path = req.get("path").and_then(|v| v.as_str())
        .or_else(|| req.get("ofdPath").and_then(|v| v.as_str()))
        .unwrap_or("");
    let resolved = session.resolve_path(path, state.is_desktop)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let result = tokio::task::spawn_blocking(move || {
        invoice_engine::parse_ofd_file(&resolved.to_string_lossy())
            .map_err(|e| format!("OFD 解析失败: {}", e))
    }).await
        .map_err(|e| AppError::Internal(format!("OFD 解析任务崩溃: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    Ok(Json(ApiResponse::ok(serde_json::to_value(&result).unwrap_or_else(|e| {
        log::warn!("序列化 OFD 结果失败: {}", e);
        json!({})
    }))))
}

async fn parse_xml_invoice(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let session_id = req.get("sessionId").and_then(|v| v.as_str()).unwrap_or("").to_string();
    let session = get_session(&state, &session_id)?;
    let path = req.get("path").and_then(|v| v.as_str())
        .or_else(|| req.get("xmlPath").and_then(|v| v.as_str()))
        .unwrap_or("");
    let resolved = session.resolve_path(path, state.is_desktop)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let result = tokio::task::spawn_blocking(move || {
        invoice_engine::parse_xml_invoice(&resolved.to_string_lossy())
            .map_err(|e| format!("XML 解析失败: {}", e))
    }).await
        .map_err(|e| AppError::Internal(format!("XML 解析任务崩溃: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    Ok(Json(ApiResponse::ok(serde_json::to_value(&result).unwrap_or_else(|e| {
        log::warn!("序列化 XML 结果失败: {}", e);
        json!({})
    }))))
}

async fn open_ofd_images(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let session_id = req.get("sessionId").and_then(|v| v.as_str()).unwrap_or("").to_string();
    let session = get_session(&state, &session_id)?;
    let path = req.get("path").and_then(|v| v.as_str())
        .or_else(|| req.get("ofdPath").and_then(|v| v.as_str()))
        .unwrap_or("");
    let resolved = session.resolve_path(path, state.is_desktop)
        .map_err(|e| AppError::BadRequest(e))?;
    session.touch();

    let result = tokio::task::spawn_blocking(move || {
        let path_str = resolved.to_string_lossy();
        let name = std::path::Path::new(path_str.as_ref())
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_else(|| "unknown".to_string());
        let size = std::path::Path::new(path_str.as_ref())
            .metadata().ok().map(|m| m.len()).unwrap_or(0);
        let images = invoice_engine::extract_ofd_images_raw(&path_str)
            .map_err(|e| format!("OFD 图片提取失败: {}", e))?;
        let mut results = Vec::new();
        for (idx, _img) in images.iter().enumerate() {
            let base_name = if name.len() > 4 { &name[..name.len()-4] } else { &name };
            results.push(serde_json::json!({
                "name": if images.len() > 1 {
                    format!("{}_第{}页.ofd", base_name, idx + 1)
                } else { name.clone() },
                "path": resolved.to_string_lossy(),
                "size": size,
                "type": "ofd",
            }));
        }
        Ok(results)
    }).await
        .map_err(|e| AppError::Internal(format!("OFD 图片任务崩溃: {}", e)))?
        .map_err(|e: String| AppError::Internal(format!("{}", e)))?;

    Ok(Json(ApiResponse::ok(serde_json::to_value(&result).unwrap_or_else(|e| {
        log::warn!("序列化 OFD 图片结果失败: {}", e);
        json!([])
    }))))
}

async fn check_path_exists(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Json<ApiResponse> {
    let path = req.get("path").and_then(|v| v.as_str()).unwrap_or("");
    // Web mode: restrict to session directory
    if !state.is_desktop {
        let p = std::path::Path::new(path);
        if !p.is_absolute() || !p.starts_with(&state.session_dir) {
            return Json(ApiResponse::ok(json!({ "exists": false })));
        }
    }
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
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let data_url = req.get("dataUrl").and_then(|v| v.as_str()).unwrap_or("").to_string();
    if data_url.is_empty() {
        return Err(AppError::BadRequest("缺少图片数据".into()));
    }

    let result = tokio::task::spawn_blocking(move || {
        use base64::Engine;
        let img = crate::pdf_engine::decode_base64_image(&data_url)
            .map_err(|e| format!("图片解码失败: {}", e))?;
        let trimmed = crate::pdf_engine::trim_white_edges(&img, 245);

        let mut buf = std::io::Cursor::new(Vec::new());
        trimmed.write_to(&mut buf, image::ImageFormat::Png)
            .map_err(|e| format!("图片编码失败: {}", e))?;
        let b64 = base64::engine::general_purpose::STANDARD.encode(buf.into_inner());
        Ok::<String, String>(format!("data:image/png;base64,{}", b64))
    }).await
        .map_err(|e| AppError::Internal(format!("裁剪任务崩溃: {}", e)))?
        .map_err(|e| AppError::Internal(e))?;

    Ok(Json(ApiResponse::ok(json!({ "dataUrl": result }))))
}

async fn copy_file(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let src = req.get("srcPath").and_then(|v| v.as_str())
        .or_else(|| req.get("src").and_then(|v| v.as_str()))
        .unwrap_or("");
    let dest = req.get("destPath").and_then(|v| v.as_str())
        .or_else(|| req.get("dest").and_then(|v| v.as_str()))
        .unwrap_or("");
    if src.is_empty() || dest.is_empty() {
        return Err(AppError::BadRequest("源路径和目标路径不能为空".into()));
    }
    // Web mode: restrict to session directory
    if !state.is_desktop {
        let src_path = std::path::Path::new(src);
        if !src_path.is_absolute() || !src_path.starts_with(&state.session_dir) {
            return Err(AppError::Unauthorized("Web 模式仅允许操作会话目录内文件".into()));
        }
        let dest_path = std::path::Path::new(dest);
        if !dest_path.is_absolute() || !dest_path.starts_with(&state.session_dir) {
            return Err(AppError::Unauthorized("Web 模式仅允许操作会话目录内文件".into()));
        }
    }
    std::fs::copy(src, dest)
        .map_err(|e| AppError::Internal(format!("复制文件失败: {}", e)))?;
    Ok(Json(ApiResponse::ok(json!({}))))
}

async fn rename_file(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let src = req.get("srcPath").and_then(|v| v.as_str())
        .or_else(|| req.get("src").and_then(|v| v.as_str()))
        .unwrap_or("");
    let dest = req.get("destPath").and_then(|v| v.as_str())
        .or_else(|| req.get("dest").and_then(|v| v.as_str()))
        .unwrap_or("");
    if src.is_empty() || dest.is_empty() {
        return Err(AppError::BadRequest("源路径和目标路径不能为空".into()));
    }
    // Web mode: restrict to session directory
    if !state.is_desktop {
        let src_path = std::path::Path::new(src);
        if !src_path.is_absolute() || !src_path.starts_with(&state.session_dir) {
            return Err(AppError::Unauthorized("Web 模式仅允许操作会话目录内文件".into()));
        }
        let dest_path = std::path::Path::new(dest);
        if !dest_path.is_absolute() || !dest_path.starts_with(&state.session_dir) {
            return Err(AppError::Unauthorized("Web 模式仅允许操作会话目录内文件".into()));
        }
    }

    let result = if std::path::Path::new(src).parent() == std::path::Path::new(dest).parent() {
        std::fs::rename(src, dest)
    } else {
        std::fs::copy(src, dest).and_then(|_| std::fs::remove_file(src))
    };
    result.map_err(|e| AppError::Internal(format!("重命名文件失败: {}", e)))?;
    Ok(Json(ApiResponse::ok(json!({}))))
}

async fn write_text_file(
    State(state): State<AppState>,
    Json(req): Json<serde_json::Value>,
) -> Result<Json<ApiResponse>, AppError> {
    let path = req.get("path").and_then(|v| v.as_str()).unwrap_or("");
    let content = req.get("content").and_then(|v| v.as_str()).unwrap_or("");
    if path.is_empty() {
        return Err(AppError::BadRequest("文件路径不能为空".into()));
    }
    // Web mode: restrict to session directory or temp directory
    if !state.is_desktop {
        let p = std::path::Path::new(path);
        if !p.is_absolute() || !p.starts_with(&state.session_dir) {
            return Err(AppError::Unauthorized("Web 模式仅允许写入会话目录内文件".into()));
        }
    }
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

pub fn cleanup_all_sessions(session_dir: &PathBuf) {
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
