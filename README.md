# 发票酱 (TicketChan)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform: Pure Web](https://img.shields.io/badge/Platform-Pure%20Web-blue.svg)]()
[![PDF.js](https://img.shields.io/badge/PDF.js-4.10-orange.svg)]()
[![Version](https://img.shields.io/badge/Version-2.2.0-blue.svg)]()

**纯前端电子发票批量排版打印工具**。单页应用,无后端,任何静态托管都能跑。

> 💡 v2.2.0 重构: 移除 Rust 后端、PDFium、OCR、OFD/XML 解析。客户端用 [PDF.js 4.10](https://github.com/mozilla/pdf.js) 渲染发票、[pdf-lib 1.17](https://github.com/Hopding/pdf-lib) 拼版 PDF、浏览器原生 `<iframe>.print()` 打印。文件存 IndexedDB,不外发。

## ✨ 功能

- **多格式**: PDF / JPG / PNG / BMP / WebP / GIF
- **PDF 文字层提取**: 客户端解析,自动识别销售方 / 金额 / 税号 / 发票号 / 日期
- **多张拼版**: 自定义 1-10×1-10 网格,A4/A5/B5/Letter/Legal/自定义纸张
- **完整排版控制**: 边距、间距、缩放、旋转、自动/手动对齐、切割线、边框、水印
- **单票独立调整**: 拖拽移动 + 滚轮缩放 (5%/步) + 九宫格对齐
- **汇总表**: 14 字段可勾选 / 双击编辑 / CSV 导出 (UTF-8 BOM)
- **批量重命名**: 4 快捷模板 + 自定义字段
- **打印**:
  - 客户端 pdf-lib 拼版 → 浏览器内置 PDF 引擎打印 (矢量保真)
  - PDF 下载 (`<a download>` 触发)
- **离线**: 数字发票场景 100% 离线
- **文件记忆**: 可选,启动时自动恢复

## 🚀 部署

`frontend/` 目录就是全部产物。直接上传到任何静态托管即可:

- **Vercel** / **Netlify** / **Cloudflare Pages** / **GitHub Pages** — 拖入 `frontend/` 目录即部署
- **对象存储** (OSS / S3 / COS) — 开启静态网站托管, 根目录指向 `frontend/`
- **公司内网** — nginx / Apache / IIS 指向 `frontend/` 目录

### 本地预览

```bash
python3 -m http.server 8080 --directory frontend
# 访问 http://localhost:8080/
```

### 部署要求

- 静态文件服务, 无后端进程
- 需正确返回 `.mjs` 文件的 `application/javascript` MIME (大多数托管默认支持)
- 启用 SPA 模式: 404 fallback 到 `index.html`(本项目无路由, 但建议配置)
- 启用 gzip/brotli 压缩 (vendor 文件 gzip 后约 800KB)

## 🛠️ 开发

无构建步骤,直接编辑 `frontend/` 下的文件,刷新浏览器即可。

## 📁 目录结构

```
fapiao/
├── frontend/                # 全部产物 (2.8MB) — 唯一需要部署的目录
│   ├── index.html
│   ├── css/styles.css
│   └── js/
│       ├── idb-store.js     # IndexedDB 文件存储 (ESM)
│       ├── pdf-client.js    # PDF.js 4.10 封装 (ESM)
│       ├── layout.js        # 拼版算法 (拼版坐标系 0 改动)
│       ├── print.js         # pdf-lib 拼版 + iframe.print()
│       ├── app.js           # 主应用
│       └── vendor/          # 第三方库 (~2.7MB)
│           ├── pdf.min.mjs
│           ├── pdf.worker.min.mjs
│           ├── pdf-lib.min.js
│           └── cmaps/       # 34 个 CJK bcmap (简繁中)
├── sample/                  # 测试发票
├── screenshots/
├── AGENTS.md
├── CHANGELOG.md
└── LICENSE
```

## 🔧 技术栈

| 用途 | 库 | 许可 |
|---|---|---|
| PDF 渲染 | PDF.js 4.10.38 | Apache 2.0 |
| PDF 生成/拼版 | pdf-lib 1.17.1 | Apache 2.0 |
| 文件存储 | IndexedDB (浏览器原生) | — |
| CJK 字体 | 34 个 cMap bcmap | Apache 2.0 |
| 打印 | 浏览器原生 `<iframe>.print()` | — |

总 vendor 大小: **~2.7MB** (gzip 后约 800KB)

## ❌ 暂不支持

- **OFD 发票** — 需要后端 OFD 解析,本次重构不包含
- **XML 数电票** — 同上
- **扫描件 OCR** — 用户主动放弃
- **服务端打印** — Web 架构无意义

## 📝 版本

- **v2.2.0**: 纯前端重构,移除 Rust + PDFium + OCR
- **v2.1.0**: 纯 Web (Axum + PDFium)
- **v2.0.7**: 最后桌面版 (Tauri)
