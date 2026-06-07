use std::net::SocketAddr;
use std::path::PathBuf;
use std::sync::Arc;

use app_lib::server::{self, AppState, ProgressTracker};

#[tokio::main]
async fn main() {
    // Simple logger init without env_logger dependency
    env_logger::init();

    let session_dir = PathBuf::from(
        std::env::var("TICKETCHAN_SESSION_DIR")
            .unwrap_or_else(|_| "/tmp/ticketchan".to_string())
    );
    let frontend_dir = PathBuf::from(
        std::env::var("TICKETCHAN_FRONTEND_DIR")
            .unwrap_or_else(|_| ".".to_string())
    );
    let auth_token = std::env::var("TICKETCHAN_AUTH_TOKEN").ok();
    let port: u16 = std::env::var("TICKETCHAN_SERVER_PORT")
        .ok()
        .and_then(|p| p.parse().ok())
        .unwrap_or(3000);

    // Clean up stale sessions from previous runs
    server::cleanup_all_sessions(&session_dir).await;

    // Create session directory
    std::fs::create_dir_all(&session_dir)
        .expect("Failed to create session directory");

    let state = AppState {
        session_dir: session_dir.clone(),
        sessions: Arc::new(dashmap::DashMap::new()),
        ocr_limit: Arc::new(tokio::sync::Semaphore::new(2)),
        render_limit: Arc::new(tokio::sync::Semaphore::new(4)),
        auth_token,
        progress: Arc::new(ProgressTracker::new()),
        frontend_dir,
        is_desktop: false,
    };

    // Periodic session cleanup (every 30 minutes)
    let cleanup_state = state.clone();
    tokio::spawn(async move {
        let mut interval = tokio::time::interval(std::time::Duration::from_secs(1800));
        loop {
            interval.tick().await;
            server::cleanup_sessions(&cleanup_state, 86400).await; // 24h TTL
        }
    });

    let app = server::build_router(state);
    let addr = SocketAddr::from(([0, 0, 0, 0], port));

    log::info!("发票酱 Web 服务启动于 http://{}", addr);
    let listener = tokio::net::TcpListener::bind(addr).await
        .expect("Failed to bind server port");
    axum::serve(listener, app).await
        .expect("Server error");
}
