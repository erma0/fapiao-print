use std::net::SocketAddr;
use std::path::PathBuf;
use std::sync::Arc;

use crate::server::{self, AppState, ProgressTracker};

pub struct EmbeddedServer {
    shutdown_tx: tokio::sync::oneshot::Sender<()>,
    port: u16,
}

impl EmbeddedServer {
    pub async fn start(session_dir: PathBuf, auth_token: Option<String>) -> Result<Self, String> {
        // Clean up stale sessions
        server::cleanup_all_sessions(&session_dir).await;
        std::fs::create_dir_all(&session_dir)
            .map_err(|e| format!("创建 session 目录失败: {}", e))?;

        let frontend_dir = PathBuf::from("."); // Desktop: Tauri serves frontend

        let state = AppState {
            session_dir: session_dir.clone(),
            sessions: Arc::new(dashmap::DashMap::new()),
            ocr_limit: Arc::new(tokio::sync::Semaphore::new(2)),
            render_limit: Arc::new(tokio::sync::Semaphore::new(4)),
            auth_token,
            progress: Arc::new(ProgressTracker::new()),
            frontend_dir,
            is_desktop: true,
        };

        // Periodic session cleanup
        let cleanup_state = state.clone();
        tokio::spawn(async move {
            let mut interval = tokio::time::interval(std::time::Duration::from_secs(1800));
            loop {
                interval.tick().await;
                server::cleanup_sessions(&cleanup_state, 86400).await;
            }
        });

        let app = server::build_router(state);

        // Try ports 3000-3010
        let mut bound_port = None;
        let mut listener = None;
        for port in 3000..=3010 {
            let addr = SocketAddr::from(([127, 0, 0, 1], port));
            match tokio::net::TcpListener::bind(addr).await {
                Ok(l) => {
                    bound_port = Some(port);
                    listener = Some(l);
                    break;
                }
                Err(_) => continue,
            }
        }

        let listener = listener.ok_or("无法绑定端口 3000-3010，请检查端口占用")?;
        let port = bound_port.unwrap();

        let (shutdown_tx, shutdown_rx) = tokio::sync::oneshot::channel::<()>();

        tokio::spawn(async move {
            axum::serve(listener, app)
                .with_graceful_shutdown(async {
                    let _ = shutdown_rx.await;
                })
                .await
                .ok();
        });

        log::info!("Embedded server started on 127.0.0.1:{}", port);

        Ok(Self { shutdown_tx, port })
    }

    pub fn port(&self) -> u16 {
        self.port
    }

    pub async fn shutdown(self) {
        let _ = self.shutdown_tx.send(());
        log::info!("Embedded server shut down");
    }
}
