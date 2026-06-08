# Build stage
FROM rust:1.82-bookworm AS builder

WORKDIR /app

# Copy manifests first for dependency caching
COPY src-tauri/Cargo.toml src-tauri/Cargo.lock ./
COPY src-tauri/invoice-engine ./invoice-engine/

# Create dummy source files to cache dependencies
RUN mkdir src && echo 'fn main() {}' > src/main.rs && echo '' > src/lib.rs

# Build dependencies only (cached layer)
RUN cargo build --release --bin ticketchan-server 2>/dev/null || true

# Copy actual source code
COPY src-tauri/src ./src
COPY src-tauri/build.rs ./build.rs

# Touch source files to force rebuild
RUN find src -name "*.rs" -exec touch {} +

# Build the actual binary
RUN cargo build --release --bin ticketchan-server

# Runtime stage
FROM debian:bookworm-slim

RUN apt-get update && apt-get install -y \
    ca-certificates \
    libglib2.0-0 \
    libpng16-16 \
    libjpeg62-turbo \
    libfontconfig1 \
    libfreetype6 \
    && rm -rf /var/lib/apt/lists/*

# Install PDFium
RUN mkdir -p /app/tools && \
    curl -sL "https://github.com/bblanchon/pdfium-binaries/releases/download/chromium%2F6963/pdfium-linux-x64.tgz" | \
    tar xz -C /app/tools --strip-components=1 --wildcards '*/lib/libpdfium.so' --wildcards '*/lib/libpdfium.so.1' 2>/dev/null || \
    echo "PDFium download failed - install manually"

WORKDIR /app

# Copy binary
COPY --from=builder /app/target/release/ticketchan-server /app/ticketchan-server

# Copy frontend
COPY src/ /app/frontend/

# Environment
ENV TICKETCHAN_SESSION_DIR=/tmp/ticketchan
ENV TICKETCHAN_FRONTEND_DIR=/app/frontend
ENV TICKETCHAN_SERVER_PORT=3000
ENV RUST_LOG=info

# Session directory
RUN mkdir -p /tmp/ticketchan

EXPOSE 3000

CMD ["/app/ticketchan-server"]
