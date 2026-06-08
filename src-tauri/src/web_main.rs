use std::net::SocketAddr;
use std::path::PathBuf;
use std::sync::Arc;

use app_lib::server::{self, AppState, ProgressTracker};

#[tokio::main]
async fn main() {
    env_logger::init();

    let session_dir = PathBuf::from(
        std::env::var("TICKETCHAN_SESSION_DIR")
            .unwrap_or_else(|_| {
                std::env::temp_dir().join("ticketchan").to_string_lossy().to_string()
            })
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
    // Default: bind loopback only. Set TICKETCHAN_BIND_ADDR=0.0.0.0 for public access
    // (only if auth_token is also configured or an external reverse proxy is used).
    let bind_ip: std::net::IpAddr = std::env::var("TICKETCHAN_BIND_ADDR")
        .ok()
        .and_then(|a| a.parse().ok())
        .unwrap_or(std::net::IpAddr::V4(std::net::Ipv4Addr::new(127, 0, 0, 1)));

    // Clean up stale sessions from previous runs
    server::cleanup_all_sessions(&session_dir);

    // Create session directory
    std::fs::create_dir_all(&session_dir)
        .expect("Failed to create session directory");

    let state = AppState {
        session_dir: session_dir.clone(),
        sessions: Arc::new(dashmap::DashMap::new()),
        ocr_limit: Arc::new(tokio::sync::Semaphore::new(server::optimal_ocr_concurrency())),
        render_limit: Arc::new(tokio::sync::Semaphore::new(server::optimal_render_concurrency())),
        auth_token,
        progress: Arc::new(ProgressTracker::new()),
        frontend_dir,
        is_desktop: false,
    };

    // Periodic session cleanup (every 30 minutes), cancellable via CancellationToken
    let cleanup_token = tokio_util::sync::CancellationToken::new();
    let cleanup_state = state.clone();
    let ct = cleanup_token.clone();
    tokio::spawn(async move {
        let mut interval = tokio::time::interval(std::time::Duration::from_secs(1800));
        loop {
            tokio::select! {
                _ = ct.cancelled() => break,
                _ = interval.tick() => {
                    server::cleanup_sessions(&cleanup_state, 86400).await; // 24h TTL
                }
            }
        }
    });

    let app = server::build_router(state);
    let addr = SocketAddr::from((bind_ip, port));

    log::info!("发票酱 Web 服务启动于 http://{} (bind: {})", addr, bind_ip);
    let listener = tokio::net::TcpListener::bind(addr).await
        .expect("Failed to bind server port");

    // Graceful shutdown on SIGTERM / SIGINT (Ctrl+C)
    let shutdown_signal = async move {
        tokio::signal::ctrl_c().await
            .expect("Failed to install Ctrl+C handler");
        log::info!("收到关闭信号，正在优雅退出...");
        cleanup_token.cancel();
    };

    axum::serve(listener, app)
        .with_graceful_shutdown(shutdown_signal)
        .await
        .expect("Server error");
}
