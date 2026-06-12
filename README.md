# 发票酱 (TicketChan)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform: Pure Web](https://img.shields.io/badge/Platform-Pure%20Web-blue.svg)]()
[![PDF.js](https://img.shields.io/badge/PDF.js-4.10-orange.svg)]()
[![Version](https://img.shields.io/badge/Version-2.2.0-blue.svg)]()

**纯前端电子发票批量排版打印工具**。单页应用,零后端,零构建,浏览器即用。

## ✨ 功能

- **多格式**: PDF / OFD / XML 数电票 / JPG / PNG / BMP / WebP / GIF
- **字段提取**: PDF 文字层 / OFD XML / XML 数电票自动提取发票信息
- **多张拼版**: 自定义 1-10×1-10 网格,A4/A5/B5/Letter/Legal/自定义纸张
- **完整排版控制**: 边距、间距、缩放、旋转、切割线、边框、水印
- **单票独立调整**: 拖拽移动 + 滚轮缩放 (5%/步) + 九宫格对齐
- **汇总表**: 14 字段可勾选 / 双击编辑 / 3 种金额合计 / CSV 导出 (UTF-8 BOM)
- **批量重命名**: 4 快捷模板 + 自定义字段
- **PDF 矢量嵌入**: pdf-lib embedPage 直接嵌入原始 PDF 页面,保留矢量
- **打印**: 浏览器原生 `<iframe>.print()` 打印,PDF 矢量保真
- **离线**: 无需网络,全浏览器运行
- **文件记忆**: 可选,启动时自动恢复文件列表

## 🚀 部署

上传到任何静态托管即可:

```bash
python3 -m http.server 8080
# 访问 http://localhost:8080/
```

- **Vercel** / **Netlify** / **Cloudflare Pages** / **GitHub Pages** — 拖入根目录即部署
- **对象存储** (OSS / S3 / COS) — 开启静态网站托管
- **公司内网** — nginx / Apache / IIS

需正确返回 `.mjs` 的 `application/javascript` MIME。建议启用 gzip/brotli (vendor gzip 后约 800KB)。

## 📁 目录结构

```
fapiao/
├── index.html
├── css/styles.css
├── js/
│   ├── app.js              # 主应用 (文件加载/状态/汇总/CSV)
│   ├── print.js            # pdf-lib 拼版 + iframe.print()
│   ├── layout.js           # 排版计算 + 预览渲染
│   ├── pdf-client.js       # PDF.js 封装 (ESM)
│   ├── ofd-client.js       # OFD 前端解析
│   ├── xml-client.js       # XML 数电票解析
│   ├── xml-utils.js        # 共享 XML 工具
│   └── idb-store.js        # IndexedDB 存储 (ESM)
├── vendor/                 # 第三方库 (~2.7MB)
│   ├── pdf.min.mjs / pdf.worker.min.mjs   # PDF.js 4.10
│   ├── pdf-lib.min.js                     # pdf-lib 1.17
│   ├── fontkit.umd.min.js                 # 字体解析
│   ├── jszip.min.js                       # JSZip
│   └── cmaps/                             # CJK 字符映射
├── screenshots/
├── CHANGELOG.md / README.md / LICENSE
└── .gitignore
```

## 🔧 技术栈

| 用途 | 库 | 许可 |
|---|---|---|
| PDF 渲染 | PDF.js 4.10 | Apache 2.0 |
| PDF 拼版 | pdf-lib 1.17 | Apache 2.0 |
| ZIP 解压 | JSZip | MIT |
| 文件存储 | IndexedDB (浏览器原生) | — |
| 打印 | 浏览器原生 `<iframe>.print()` | — |

总 vendor: **~3MB** (gzip 后约 1MB)

## 📝 版本

- **v2.2.0**: 纯前端,零后端零依赖 (当前)
- **v2.1.0**: Axum Web Server + PDFium (已废弃)
- **v2.0.7**: Tauri 桌面版 (master 分支)
