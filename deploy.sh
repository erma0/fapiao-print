#!/bin/bash
set -euo pipefail

VERSION=${1:-latest}
IMAGE_NAME="ticketchan-web"
REGISTRY=${REGISTRY:-}

echo "=== 发票酱 Web 部署 ==="
echo "Version: $VERSION"

# Build
echo "[1/3] Building Docker image..."
docker build -t ${IMAGE_NAME}:${VERSION} -t ${IMAGE_NAME}:latest .

# Push (if registry is set)
if [ -n "$REGISTRY" ]; then
    echo "[2/3] Pushing to registry..."
    docker tag ${IMAGE_NAME}:${VERSION} ${REGISTRY}/${IMAGE_NAME}:${VERSION}
    docker push ${REGISTRY}/${IMAGE_NAME}:${VERSION}
    docker push ${REGISTRY}/${IMAGE_NAME}:latest
else
    echo "[2/3] Skip push (no REGISTRY set)"
fi

# Deploy
echo "[3/3] Starting services..."
docker compose up -d

echo "=== 部署完成 ==="
echo "访问: http://localhost:3000"
