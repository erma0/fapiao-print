![fapiao-print](https://socialify.git.ci/erma0/fapiao-print/image?description=1&font=Source+Code+Pro&forks=1&issues=1&language=1&name=1&owner=1&pattern=Circuit+Board&stargazers=1&theme=Auto)

# 📄 发票酱

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform: Cross-platform](https://img.shields.io/badge/Platform-Win%20%7C%20Linux%20%7C%20macOS%20%7C%20Web-blue.svg)]()
[![Tauri 2.x](https://img.shields.io/badge/Tauri-2.x-orange.svg)]()
[![Version](https://img.shields.io/badge/Version-2.1.0-blue.svg)]()

轻量桌面应用，专为批量打印电子发票设计。支持 PDF、OFD、图片等多格式导入，智能排版，一键打印或导出。

提供 **轻量版** 和 **OCR 版**（含 PP-OCRv5 智能识别），支持 **Windows / Linux / macOS 桌面** 及 **Docker / 裸金属 Web 部署**。

## ✨ 功能特性

### 🏆 OFD 完整支持

OFD（开放版式文档）是国家标准电子发票格式，本工具提供原生完整支持 — 矢量渲染、发票信息直提、印章保真，拖入即用，无需 OCR。

> ⚠️ 不同厂商/转换工具生成的 OFD 发票格式存在差异（如税务原版 OFD、iloveofd 转换、dzcp 公共服务平台等），如遇解析渲染问题请及时反馈，我们会持续适配。

### 📥 文件管理

- **多格式支持**：PDF、OFD、XML 数电票、JPG、PNG、BMP、WebP、TIFF
- **XML 数电票**（v2.0.7）：解析 `<EInvoice>` 格式，提取发票号码/日期/金额/买卖方信息，汇总表、CSV 导出、批量重命名全兼容；纯数据格式不参与排版打印
- **文件列表记忆**（v2.0.7）：可选开关，启动时自动恢复上次打开的文件列表，仅记忆文件路径
- **打印状态追踪**（v2.0.7）：三种过滤（全部/未打印/已打印），打印后自动标记绿色 ✓，状态持久化
- **PDF 渲染双引擎**（v1.10.0+）：首选 WinRT 原生渲染（`Windows.Data.Pdf`），自动 fallback PDFium（Chromium 内核），兼容企业精简版/LTSC 系统
- **PDF 文字层提取**（轻量版也可用）：解析 PDF 内容流 Tm+Tj/TJ 指令直接提取文字坐标，~5ms/页，无需 OCR 即可识别发票信息
- **PP-OCRv5 智能识别**（OCR 版，适用于图片型 PDF 和图片）：文本优先 + 坐标回退双重架构，含税价 / 不含税价 / 税额数学验证配对，发票号码 / 日期 / 买卖方信息自动提取
- **金额校验可视化**：OCR / PDF 提取金额求和校验失败时，发票卡片金额徽章 ⚠ 警告标识，hover 可查看含税/不含税/税额验证详情
- **EXIF 方向自动修正**：导入图片/车票时自动读取 EXIF Orientation 旋转像素，PDF /Rotate 属性 + CropBox 坐标归一化保障页面方向正确
- **发票查验**：一键跳转国家税务总局查验平台
- **骨架屏渐进加载**：批量导入时骨架屏秒出 + 逐文件渐进渲染 + 持久进度 toast，大文件不卡 UI
- **↑↓ 排序**：↑↓ 按钮排序（替换 Tauri webview 拖拽卡顿），hover 浮动显示不占空间
- **批量重命名**（v2.0.5）：汇总表内嵌面板，预设模板（金额+销售方+号码等）或自定义字段勾选，一键批量重命名发票磁盘文件，重名自动序号
- **设置自动记忆**：关闭后自动记住布局、纸张、打印模式等全部设置，打开即恢复

### 📐 排版设置

- **纸张**：A4 / A5 / B5 / Letter / Legal / 自定义
- **布局**：6 预设（1×1 / 2×1 / 3×2 / 1×2 / 2×2 / 3×3）+ 自定义行列（1-10 × 1-10），自动横纵方向
- **边距 / 间距**：独立可调，预设快捷按钮
- **缩放**：自适应 / 拉伸填充 / 原始大小 / 自定义百分比
- **旋转**：全局 0° / 90° / 180° / 270° / 自动 + 单张旋转
- **单票独立调整**（v1.9.0+，v2.0.1-v2.0.2 增强）：每张发票预览拖拽移动 + 角落 handle 缩放，九宫格快速对齐，滚轮微调 + 滚轮单票缩放（5%/步），±150mm 偏移范围，拖拽约束动态化，放大上限 3x，编辑态溢出预览，双击重置，调整参数可选持久化记忆，侧边栏「单票调整」面板或发票弹窗参数编辑，PDF 按参数裁剪输出

### ✂️ 辅助功能

- 裁切线、编号标记、边框显示、裁剪白边、自定义水印
- 金额统计、车票票种标签、发票类型自动检测
- **页脚**：打印页码（第 X 页 / 共 Y 页）、打印日期、自定义页脚文本，独立下边距控制

### 📊 数据导出

- **发票汇总表**（v2.0.3）：报销必备，一键导出所有发票明细，字段可勾选（14 项），金额/名称等可直接编辑修正，合计行自动汇总含税/不含税/税额，CSV 格式 Excel 直接打开，列选择和备注持久化记忆

### 🖨️ 打印与导出

- **打印模式**：四种模式可选
  - **PDF 阅读器**（默认）：生成 PDF 后由系统默认程序处理，保持矢量质量，数据量最小
  - **弹窗确认**：预览后确认打印，可选 PDFium 或 SumatraPDF 引擎
  - **静默打印（PDFium）**：Chromium PDFium 引擎直打打印机 DC，打印清晰（需下载 pdfium.dll）
  - **静默打印（SumatraPDF）**：通过 SumatraPDF 直接发送到打印机（需安装 SumatraPDF）
- **PDF 统一直通**（v1.9.0+）：lopdf Form XObject + JPEG DCTDecode 直通，PDF 页面以原始质量嵌入合成 PDF，无二次压缩
- **印章烘焙**（v2.0.4）：生成 PDF 时自动将原票印章/签章标注烘焙到输出，印章位置/大小与原票一致
- **份数控制**：全局 + 单张份数，逐份 / 逐页打印，双面打印，彩色 / 灰度 / 黑白
- **PDF 导出**：自动打开或自定义保存目录
- **确认弹窗**：打印前显示发票数量 / 版面 / 纸张 / 打印机 / 引擎 / 份数，防止误操作

### 🎨 界面

- 深色 / 浅色模式、实时预览（缩放 + 翻页）
- **快捷键**：`Ctrl+O` 添加 · `Ctrl+P` 打印 · `Ctrl++/-` 缩放 · `Ctrl+0` 自适应 · `←→` 翻页

## 📸 界面预览

<table>
  <tr>
    <td align="center">☀️ 浅色模式</td>
    <td align="center">🌙 深色模式</td>
  </tr>
  <tr>
    <td><img src="screenshots/light.png" alt="浅色模式" width="480"/></td>
    <td><img src="screenshots/dark.png" alt="深色模式" width="480"/></td>
  </tr>
</table>

## 📦 下载

从 [Releases](../../releases) 下载最新版本：

### 桌面版

| 文件 | 平台 | 说明 |
|------|------|------|
| `发票酱_x64-setup.exe` | Windows | 轻量版安装包 |
| `发票酱_x64_绿色版.exe` | Windows | 轻量版便携（单文件 exe） |
| `发票酱_x64_OCR版-setup.exe` | Windows | OCR 版安装包（含 PP-OCRv5） |
| `发票酱_x64_OCR绿色版.zip` | Windows | OCR 版便携（exe + models/） |
| `ticketchan_x64.deb` | Linux | Debian/Ubuntu 安装包 |
| `ticketchan_x86_64.AppImage` | Linux | 通用 Linux 便携版 |
| `ticketchan_x64.dmg` | macOS | macOS 安装镜像 |

### Web 版

```bash
# Docker
docker run -d -p 3000:3000 -v ticketchan-sessions:/tmp/ticketchan ghcr.io/erma0/ticketchan-server

# 裸金属（Linux）
sudo apt install libpdfium-dev
cargo install ticketchan-server
TICKETCHAN_FRONTEND_DIR=/opt/ticketchan/frontend ticketchan-server
```

> 💡 文字型 PDF / OFD 发票选轻量版即可自动提取金额和销售方信息；图片型 PDF 和图片需 OCR 版。

**桌面系统要求**：Windows 10 1803+ / Windows 11 / Linux (glibc 2.28+) / macOS 12+。

**⚠️ 不支持 Windows 7/8**：依赖的 WebView2 和系统 PDF 组件已停止支持。

## 📋 使用说明

1. **添加发票**：点击「➕ 添加」或拖放文件（支持 PDF / OFD / 图片混选）
2. **排版设置**：左侧「⚙ 排版」面板调整纸张、布局、边距
3. **预览检查**：主区域实时预览，支持缩放翻页；文字型 PDF / OFD 自动提取金额信息，图片型 PDF 和图片需 OCR 版
4. **打印**：点击「🖨 打印」，选择弹出预览或直接打印
5. **保存 PDF**：点击「📥 PDF」导出合成 PDF

## 🛠 技术栈

| 层级 | 技术 | 说明 |
|------|------|------|
| 前端 | 原生 HTML/CSS/JS | 模块化（app / ocr / layout / print / api），零依赖框架 |
| 通信 | Axum HTTP Server（模式 B） | Tauri 启动时 spawn 本地 server，桌面+Web 统一 fetch() |
| 桌面框架 | Tauri 2.x (Rust) | 跨平台窗口壳，Rust 条件编译管理功能开关 |
| PDF 渲染 | WinRT + PDFium 双引擎 | WinRT 原生渲染优先，自动 fallback PDFium（Chromium 内核） |
| PDF 生成 | printpdf 0.9 + lopdf 0.39 | JPEG 直通零质量损失、PDF 页面 Form XObject 全布局直通 |
| OFD/XML 解析 | Rust 独立 crate (`invoice-engine/`) | 矢量 SVG 渲染 + 发票 XML/数电票字段直提 |
| OCR | ocr-rs 2.2 (PP-OCRv5 + MNN) | 文本优先 + 坐标回退，对比度增强（OCR 版可选） |
| 打印 | Print Spooler API + PDFium + open crate | 静默打印 / 弹窗确认 / PDF 阅读器（Windows）/ 系统打开（Linux/macOS） |
| Web 部署 | Axum 0.8 + Tokio + Nginx + Docker | REST API + SSE 进度 + Session 管理 + Token 认证 |

## 📁 项目结构

```
ticketchan/
├── src/                            # 前端
│   ├── index.html / styles.css
│   ├── api.js                      # 统一通信适配层（桌面+Web）
│   ├── app.js                      # 主入口、状态、文件加载
│   ├── ocr.js                      # OCR 提取（文本优先 + 坐标回退）
│   ├── layout.js                   # calculateLayout() + 预览渲染
│   └── print.js                    # 打印 / 导出 PDF
├── src-tauri/                      # Tauri / Rust 后端
│   ├── src/
│   │   ├── main.rs                 # 桌面入口
│   │   ├── lib.rs                  # Tauri 命令、拖放、进程管理
│   │   ├── pdf_engine.rs           # PDF 生成/渲染/文字提取/OCR
│   │   ├── pdfium_bindings.rs      # PDFium FFI 绑定（跨平台）
│   │   ├── pdfium_render.rs        # PDFium 位图渲染（跨平台）
│   │   ├── pdfium_print.rs         # PDFium GDI 矢量打印（Windows）
│   │   ├── platform.rs             # 平台抽象层（Win/Linux/macOS）
│   │   ├── server.rs               # Axum HTTP Server（REST API + SSE）
│   │   ├── session.rs              # Web Session 管理
│   │   ├── tauri_server.rs         # Tauri 嵌入 server 启动
│   │   └── web_main.rs             # Web 版入口
│   ├── invoice-engine/              # 发票引擎（OFD SVG + XML 解析）
│   ├── models/                     # PP-OCRv5 MNN 模型
│   ├── Cargo.toml                  # ocr feature flag
│   ├── tauri.conf.json             # 轻量版配置
│   └── tauri.ocr.conf.json         # OCR 版配置
├── Dockerfile                      # Docker 镜像构建
├── docker-compose.yml              # Docker 编排
├── nginx.conf                      # 反向代理配置
├── ticketchan.service              # systemd 服务文件
├── deploy.sh                       # 一键部署脚本
├── scripts/
│   ├── build-all.js                # 一键全量构建
│   └── bump-version.js             # 版本号同步
└── package.json
```

## 🚀 开发

**环境要求**：Node.js 18+、Rust 1.77+、Windows 10/11 / Linux / macOS

```bash
npm install

# 桌面版开发
npm run dev          # 轻量版
npm run dev:ocr      # OCR 版

# Web 版开发
cd src-tauri
cargo run --bin ticketchan-server
# → http://localhost:3000

# Docker 构建
docker build -t ticketchan-server .
docker-compose up -d

# 构建
npm run build        # 轻量版
npm run build:ocr    # OCR 版
npm run build:all    # 一键全量构建

# 版本号
npm run bump 2.1.0   # 同步 package.json → Cargo.toml → tauri.conf.json
```

## 🗺 路线图

- [x] OFD 完整支持
- [x] PDF 全布局直通
- [x] 单票独立调整
- [x] Print Spooler API 静默打印
- [x] PDFium 矢量静默打印
- [x] OCR Feature Flag 双版本构建
- [x] PDF 文字层提取
- [x] 设置持久化 / 金额校验可视化
- [x] 发票汇总表导出
- [x] 批量重命名发票文件
- [x] XML 数电票支持
- [x] 文件列表记忆 / 打印状态追踪（v2.0.7）
- [x] **跨平台桌面 + Web 部署**（v2.1.0）
- [ ] Docker OCR 版（MNN 模型）
- [ ] LDAP/OIDC 用户认证
- [ ] PWA 离线支持
- [ ] 水平扩展（对象存储 + Redis Session）

## 🤖 关于此项目

本项目由 AI 辅助生成，历经 180+ 轮迭代。主要攻克：Tauri 2.x 对话框死锁、WebView2 拖放失效、WinRT COM 接口适配、ocr-rs 条件编译集成、OFD 矢量渲染（DrawParam 继承链 / 印章偏移 / ImageMask 遮罩合成）、PDF 引擎 JPEG 直通与 lopdf Form XObject 全布局直通、PDFium 矢量打印（DLL 生命周期管理 / SEH 原生崩溃保护）、PDF 文字层坐标提取（rayon 并行）、跨平台改造（Axum HTTP Server 嵌入 / 平台抽象层 / Docker Web 部署 / SSE 进度推送）。

## 📄 许可证

[MIT License](LICENSE)

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=erma0/fapiao-print&type=Date)](https://star-history.com/#erma0/fapiao-print&Date)
