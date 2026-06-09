# 发票酱 — Agent 指南

> 本文档反映当前分支 `feat/web-cross-platform`（v2.1.0）的真实状态。
> v2.1.0 不是「加 Web 模式」,而是**彻底移除 Tauri 桌面壳**,重构为纯 Axum Web 应用 + 浏览器前端。
> 桌面打包 (NSIS / 便携版 / OCR 版双配置) 在 v2.1.0 已**完全废弃**。

## 项目概览

- **版本**: v2.1.0
- **技术栈**: Axum 0.8 (Rust) + 原生 HTML/CSS/JS (无框架) + 浏览器/WebView
- **本质**: 单二进制 `ticketchan-server` HTTP 服务器,前端通过 `fetch()` 调用 REST API
- **后端**: `src/{main.rs, lib.rs, pdf_engine.rs, pdfium_bindings.rs, pdfium_render.rs, pdfium_print.rs, platform.rs, server.rs, session.rs}` (8250 行)
- **前端**: `frontend/{index.html, css/styles.css, js/{api, app, layout, ocr, print}.js}`
- **OFD/XML 解析**: `invoice-engine/` — 独立 crate,纯 Rust 解析
- **部署形态**: 本地 HTTP server / Docker 容器 / 裸金属 + systemd
- **客户端**: 任意现代浏览器,无需安装,可选 Tauri WebView shell 嵌入(已不在本仓库)

## 目录结构

```
fapiao/
├── Cargo.toml              # 单一 crate,只构建 ticketchan-server
├── Cargo.lock
├── build.rs                # 编译 Windows SEH 包装器 (seh_wrapper.c)
├── Dockerfile              # 多阶段构建,只产 ticketchan-server
├── docker-compose.yml      # 一键启动
├── deploy.sh               # 构建+推送脚本
├── nginx.conf              # 反向代理 + CORS + gzip + SSE 配置
├── ticketchan.service      # systemd unit
├── src/                    # Rust 后端 (已从 src-tauri/src/ 移到根)
│   ├── main.rs             # ticketchan-server 入口,加载环境变量 + AppState
│   ├── lib.rs              # app_lib,导出各模块 + 共享工具
│   ├── server.rs           # Axum 路由 + 25 个 handler
│   ├── session.rs          # Web session 管理 (24h TTL)
│   ├── pdf_engine.rs       # 5819 行,核心 PDF 排版/生成/打印
│   ├── pdfium_bindings.rs  # FFI 绑定
│   ├── pdfium_render.rs    # 跨平台位图渲染
│   ├── pdfium_print.rs     # Windows GDI 矢量打印 (cfg windows)
│   ├── platform.rs         # 平台抽象层 (PDFium/打印机/文件打开)
│   └── seh_wrapper.c       # Windows SEH 保护
├── invoice-engine/         # OFD + XML 数电票解析 (独立 crate)
├── frontend/               # 浏览器前端 (从 src-tauri 根 src/ 移到根 frontend/)
│   ├── index.html
│   ├── css/styles.css
│   └── js/{api,app,layout,ocr,print}.js
├── models/                 # OCR 模型 (PP-OCRv5, MNN)
├── sample/                 # 测试样例
├── screenshots/            # README 截图
├── docs/                   # 专题文档
└── target/                 # 编译产物 (请勿提交)
```

## 常用命令

### Web Server (唯一部署形态)

```bash
# 开发模式 — 绑定 127.0.0.1:3000
cargo run --release --bin ticketchan-server

# 自定义端口
TICKETCHAN_SERVER_PORT=8080 cargo run --bin ticketchan-server

# 公网监听 (必须配合 AUTH_TOKEN 或外层反向代理)
TICKETCHAN_BIND_ADDR=0.0.0.0 TICKETCHAN_AUTH_TOKEN=mysecret cargo run --bin ticketchan-server

# 自定义前端目录 (默认 ./frontend)
TICKETCHAN_FRONTEND_DIR=/var/www/ticketchan cargo run --bin ticketchan-server

# 自定义 session 目录 (默认系统临时目录/ticketchan)
TICKETCHAN_SESSION_DIR=/data/ticketchan/sessions cargo run --bin ticketchan-server
```

### Docker

```bash
docker build -t ticketchan-web .
docker compose up -d
# 或
./deploy.sh latest
```

### 编译检查

```bash
cargo check                              # 轻量版 (无 OCR)
cargo check --features ocr               # OCR 版
cargo check --bin ticketchan-server      # 同 cargo check,显式指定二进制
```

### 清理旧桌面产物

```bash
# target/ 下 v2.1.0 之前的旧二进制需手动清理
rm -f target/release/{fapiao-print.exe,ticketchan.exe,发票打印工具*.exe}
# Cargo.lock 在 v2.1.0 重新生成,版本号已统一
```

## 环境变量

| 变量 | 默认 | 说明 |
|------|------|------|
| `TICKETCHAN_SERVER_PORT` | `3000` | HTTP 监听端口 |
| `TICKETCHAN_BIND_ADDR` | `127.0.0.1` | 绑定 IP,公网用 `0.0.0.0` (需配 AUTH_TOKEN) |
| `TICKETCHAN_AUTH_TOKEN` | (无) | Bearer Token 认证,设置后所有 API 必带 `Authorization: Bearer ...` |
| `TICKETCHAN_FRONTEND_DIR` | `./frontend` | 前端静态文件目录 |
| `TICKETCHAN_SESSION_DIR` | `<temp>/ticketchan` | Web session 临时文件目录 |

## 通信架构 (v2.1.0 纯 Web)

**没有 Tauri,没有 IPC,没有 WebView shell**。所有客户端走 HTTP:

- 浏览器打开 `http://<host>:<port>/` → 服务端从 `TICKETCHAN_FRONTEND_DIR` 返回 `index.html`
- 前端通过 `fetch('/api/v1/<cmd>', { method, body })` 调用 25 个 REST 端点
- 长任务进度通过 SSE (`/api/v1/progress/<taskId>`) 推送
- 文件上传走 `multipart/form-data`,下载走 `<a download>` 或 `Content-Disposition`

### 前端 `api.js` 适配层

`frontend/js/api.js` — 统一 HTTP 通信层:

- `__api.call(cmd, params)` — 构造 `/api/v1/<cmd>` POST 请求,处理 JSON 序列化/反序列化
- `__api.listen(taskId, cb)` — SSE 进度监听 (`EventSource`)
- `__api.downloadFile(sessionId, filename)` — 浏览器 `<a download>` 触发下载
- `__api.openUrl(url)` — `window.open()`
- `__api.init()` — 启动时 `fetch('/api/v1/health')` 等待 server 就绪

### Web Server 层

`src/server.rs` (1002 行) — Axum 0.8 路由 + 25 个 handler:

- `AppState` — 共享状态 (`sessions: DashMap` / `ocr_limit & render_limit: Semaphore` / `progress` / `auth_token` / `frontend_dir` / `session_dir`)
- `AppError` — 11 种错误码统一 `IntoResponse`
- `ProgressTracker` — `DashMap<String, broadcast::Sender>` SSE 进度架构
- 安全: Web 模式路径隔离、100MB 上传限制 (`tower-http` RequestBodyLimitLayer)、可选 Bearer Token 认证

### 会话管理

`src/session.rs` (122 行) — Web 文件路径管理:

- `resolve_path()` — `Path::canonicalize()` 防路径穿越,Web 模式拒绝绝对路径外的访问
- `register_file()` — DashMap 原子计数,UUID 重命名防冲突
- 24h TTL,30min 定时清理 + 启动清理
- Session 目录是**唯一可信的 Web 文件存储区**,所有上传文件落入此目录

### 平台抽象层

`src/platform.rs` (361 行) — 所有 `#[cfg(target_os)]` 集中管理:

| 函数 | Windows | Linux/macOS |
|------|---------|-------------|
| `load_pdfium_lib()` | `tools/pdfium.dll` | `tools/libpdfium.so` / ldconfig |
| `list_printers()` | `EnumPrintersW` | `lpstat -v` (CUPS) |
| `get_default_printer_name()` | Win32 默认打印机 | CUPS 默认 |
| `open_file()` | `ShellExecuteW` | `xdg-open` / `open` crate |
| `has_winrt_pdf()` | WinRT API 可用 | 返回 false |

### Web 部署

- **Docker**: 多阶段构建,`rust:1.82-bookworm` 编译 → `debian:bookworm-slim` 运行,内置 PDFium 下载
- **裸金属**: `ticketchan-server` 二进制 + `ticketchan.service` (systemd) — 自动重启 + 日志
- **Nginx**: 反向代理 + CORS + gzip + SSE `proxy_buffering off` (否则进度推送卡死)

## IPC 异步化 (async + spawn_blocking)

所有 CPU 密集型后端 handler 必须用 `async fn` + `tokio::task::spawn_blocking` 包装,防止 tokio runtime 线程饥饿。

- `render_pdf_pages` / `render_pdf_pages_pdfium` / `extract_pdf_text` / `extract_pdf_texts` 均已异步化
- `spawn_blocking` 将计算移到 blocking thread pool,主 reactor 线程可继续处理请求
- 同步 handler 在大量并发时会导致 60s 无响应

## 架构要点

### PDF 生成双管道

首选 **lopdf 直通管道** (矢量无损) → 失败时自动回退 **printpdf 渲染管道**

- `generate_pdf_from_layout()` 入口
- lopdf 直通: `can_passthrough_pdf()` 判断 → `extract_page_as_form_xobject()` → JPEG DCTDecode 嵌入
- 打印三模式: PDF阅读器(默认) / 弹窗确认 / 静默打印PDFium — **SumatraPDF 已在 v2.1.0 完全移除**
- **PDF阅读器模式已知限制**: 通过 `ShellExecuteW` 委托系统默认 PDF 阅读器打印,`printto` 动词能否指定打印机取决于阅读器实现 (Edge/Chrome 内置查看器不支持),多数情况下 fallback 到 `print` 动词使用默认打印机,**无法可靠控制打印机选择**

### PDF 渲染双引擎 (v1.9.10+)

首选 **WinRT PDF** (Windows 系统组件) → 失败时自动回退 **PDFium 渲染**

- 启动检测: `check_winrt_pdf_available()` 创建临时 PDF 测试 WinRT `PdfDocument` API
- WinRT 渲染: `render_pdf_pages()` — `windows::Data::Pdf::PdfDocument` + `StorageFile`
- PDFium 渲染: `render_pdf_pages_pdfium()` — `FPDF_LoadMemDocument` + `FPDF_RenderPageBitmap` → PNG
- 前端 fallback 链: `_winrtPdfAvailable` 标志 → WinRT 失败自动切换 PDFium
- PDFium 位图渲染: `pdfium_render::render_pdf_to_images()` — BGRA→RGBA 转换 + PNG 编码

### 预览与打印 DPI 分离 (v1.10.5)

预览和打印使用不同的 DPI 和图片格式,兼顾速度与质量:

- **预览 DPI**: `PDF_PREVIEW_DPI = 150` (屏幕显示足够清晰,是打印 DPI 的一半)
- **打印/保存 DPI**: `PDF_RENDER_DPI = 300` (高质量输出,不变)
- **预览格式**: JPEG (quality 80%),文件体积比 PNG 小 60-80%
- **打印格式**: PDF 直通管道输出矢量 PDF,不受预览分辨率影响
- `RenderedPage.format` 字段: 前端据此判断图片格式 (`"png"` 或 `"jpeg"`)

### 发票字段提取

**路径优先级**: PDF文字层 > OFD XML > XML 数电票 > OCR

- **发票类型检测**: `_detectInvoiceType()` — nontax (优先级最高) > vat > ticket > ride > unknown
- **金额三阶段**: 含税价 → 数学验证配对 → 区域解析
- **中文大写兜底**: `parseChineseNumeral()` — 阿拉伯金额因字体/编码丢失时的 fallback
- **OCR 跳过条件**: `_pdfTextExtracted && sellerName && amountTax > 0`

### XML 数电票解析 (v2.0.7+)

`invoice-engine::parse_xml_invoice()` 解析独立 XML 数电票文件,提取结构化发票数据。

- **格式**: 纯结构化数据 (`<EInvoice>` 根元素),**无版式/排版信息**,不可渲染票面
- **用途**: 文件列表展示、金额统计、汇总表导出、批量重命名
- **不参与排版打印**: `getActiveFiles()` 过滤 `_xmlInvoice` 标记,`getFileIndex()` 返回 null
- **字段提取**: `parse_xml_invoice_fields()` — 字符串匹配提取标签内容
- **前端标记**: `fileObj._xmlInvoice = true`,无 `previewUrl`/`ow`/`oh`
- **文件列表**: 显示 XML 占位符 + 发票尾号,而非图片缩略图

### 文件列表记忆 (v2.0.7)

可选功能,启动时自动恢复上次打开的文件列表。

- **开关**: `S.feat.fileListMemory`,设置面板「记忆发票列表」,默认关闭
- **恢复机制**: `restoreFiles()` 启动时批量恢复文件路径 → `check_path_exists` 校验
- **标志保护**: `_isRestoringFiles` 标志阻止恢复期间触发 OCR 自动识别
- **轻量设计**: 仅记忆文件路径 (不保存金额/OCR 数据),与设置持久化分离

### 打印状态追踪 (v2.0.7)

追踪发票是否已打印,支持过滤和持久化。

- **三种过滤**: 侧边栏顶部「全部/未打印/已打印」`.print-filter-bar` 过滤按钮组
- **自动标记**: 三种打印模式成功后自动 `markFilesAsPrinted()` (SumatraPDF 已移除) → 绿色 ✓ 标识
- **持久化**: `_printedMap` 始终保存到 localStorage,不受功能开关影响
- **迁移**: `clearAll()` / `executeRename()` / `resetSettings()` 均正确迁移打印状态 key

### PDF 文字层提取 (v1.9.4+ / 批量 v1.10.5)

Rust `extract_pdf_text()` 解析 lopdf content stream,前端 `applyPdfTextResult()` 复用 `extractByCoordinates()`。

**批量提取 (v1.10.5)**:
- `extract_pdf_texts(pdf_path, page_indices)` — 一次打开 PDF,rayon 并行提取多页文字
- 前端 `applyPdfTextToResults(results, pdfPath)` — 按 PDF 路径分组,多 PDF 文件独立批量调用
- 批量失败时自动回退到单页 `extract_pdf_text()`

**关键坑**:
- Form XObject 内嵌字体需展开 (`/Subtype /Form`)
- GBK-EUC-H 编码需 `encoding_rs::GBK.decode()` 兜底
- `Content::encode()` 最后无换行,追加字节前必须加 `\n`
- 内容流顺序 ≠ 视觉顺序,金额取**最大** ¥ 金额

### 页脚与分割线

- **页脚边距模型**: footerMargin 是纸张底部额外独立空间,不影响 slot 边距
- **分割线**: JS 端 top-down 坐标,Rust 端 bottom-up 坐标 (PDF 标准),⚠️ 不要做坐标转换

### PDFium 模块拆分 (v2.1.0)

- `pdfium_bindings.rs` (149 行) — FFI 绑定,17 个函数指针
- `pdfium_render.rs` (182 行) — 跨平台位图渲染,JPEG quality 80% + RGBA→RGB 颜色转换
- `pdfium_print.rs` (465 行) — `#[cfg(target_os = "windows")]` 仅 GDI 矢量打印
- **PDFium 线程安全**: 2.1.0 早期尝试 `RwLock` 允许多线程并发渲染,实际遇到 `ACCESS_VIOLATION`,**已回退 `Mutex`** (commit 556180c),串行化访问
- `Drop` 实现调用 `FPDF_DestroyLibrary`

### PDFium 下载统一 (v2.1.0)

`download_pdfium_dll` 单一实现跨三平台:

- 通过 `platform::get_pdfium_lib_name()` 获取 `.dll` / `.so` / `.dylib`
- 通过 `platform::get_pdfium_download_url()` 获取 GitHub 下载 URL
- 同一个下载→解压→提取逻辑,消除 ~100 行重复代码

### PDFium 打印 SEH 保护 (v1.10.3)

部分打印机驱动的 GDI 实现有 bug,`FPDF_RenderPage` 直打 DC 时可能触发原生访问违例 (ACCESS_VIOLATION),Rust 无法捕获导致直接闪退。

- **SEH 包装器**:`src/seh_wrapper.c` C 文件,用 `__try/__except` 捕获原生崩溃
- **矢量优先 + 位图 fallback**: 始终先尝试矢量直打 DC (零质量损失),仅在 SEH 捕获异常时自动 fallback 到 `FPDF_RenderPageBitmap` + `StretchDIBits` 位图渲染
- **编译**: `cc` build-dependency 将 C 文件编译为静态库链接

### DEVMODE 完整缓冲区 (v1.10.3)

`get_printer_default_devmode()` 必须保留驱动私有数据,否则 `CreateDCW` 访问违例。

- 原先用 `std::ptr::read` 只复制 `sizeof(DEVMODEW)` 字节,丢弃 `dmDriverExtra` 字节
- 现改为返回完整 `Vec<u8>` 缓冲区,保留全部驱动配置 (纸盒选择、纸张来源等)

### 打印流程解耦 (v1.10.4)

各打印模式独立调用对应命令,不再经 `generate_pdf_from_layout` 隐式降级。

- PDFium / PDF 阅读器模式直接调用各自的打印命令
- SumatraPDF 模式已删除

### 设置持久化 (v1.10.1)

关闭浏览器后下次打开自动恢复设置 (localStorage):

- **统一入口**: `saveSettings()` / `loadSettings()` — `ticketchan-settings` JSON 存储
- **覆盖范围**: 排版布局、纸张、边距、缩放、旋转、份数、颜色、打印模式、辅助开关、水印、页脚、下边距
- **防抖保存**: `updatePreview()` 500ms 防抖自动触发 `saveSettings()`
- **恢复默认**: 清除所有持久化数据

### 金额校验可视化 (v1.10.4)

OCR 和 PDF 文字提取金额求和校验失败时可视化提示。

- 发票卡片金额徽章显示 ⚠ 警告标识
- hover 警告徽章可查看含税/不含税/税额/验证计算详情
- 汇总栏新增校验异常发票计数提示

### 排版份数批量设置 (v1.10.4)

文件列表新增 ② 按钮,支持批量设置选中发票排版份数 (×1/×2/×3)。

- **区分概念**: 「排版份数」= 每张发票在版面中重复几次 / 「打印份数」= 整版打印几份
- 模态框和设置面板分别标注,避免混淆

### 单票独立调整增强 (v2.0.1+v2.0.2)

每张发票可独立缩放/偏移,CSS transform 预览 + Rust `SlotSpec` 参数 PDF 裁剪输出。

**v2.0.1 — UI 完善**:
- **快速对齐九宫格**: 一键贴边/居中,9 种对齐方向
- **鼠标滚轮增减**: 所有数字输入框和滑块支持滚轮微调
- **拖动修复**: CSS transform 应用到 wrapper div (与渲染一致),消除拖动错位
- **偏移范围扩展**: ±50→±150mm,覆盖 A3/A4 所有布局
- **调整记忆**: 可选的单票调整配置持久化 (按文件名匹配,跨会话恢复)

**v2.0.2 — 交互优化**:
- **放大上限 3x**: slotScale 上限从 2.0 放宽到 3.0
- **拖拽约束动态化**: 根据发票实际显示尺寸动态计算可拖范围
- **滚轮缩放单票**: 单击选中槽位后,鼠标滚轮直接调节该票缩放比例 (5%/步)
- **编辑态溢出可见**: 选中或拖拽中的发票临时显示超出 slot 的内容

**数据模型**: `fileObj.{slotScale, slotOffsetX, slotOffsetY}` — 独立于全局排版参数
**持久化**: `perFileAdjustments` Map 按文件名匹配,可选开启/关闭

### 预览加载优化 (v1.10.5)

- **预览 DPI**: 300 → 150,渲染像素减少 75%
- **图片格式**: PNG → JPEG (quality 80%),文件体积减少 60-80%
- **打印不受影响**: 打印/保存走独立矢量流程 (lopdf 直通)

### 智能 PDF 缓存 (v1.10.5)

用深度对象比较替代 dirty flag,精确判断缓存的 PDF 是否可复用。

- `deepEqual(a, b)` — 递归深度比较,比较整个 `LayoutRenderRequest`
- `canUseCachedPdf(currentRequest)` — 只要排版参数没变,任何打印模式/H5导出都复用
- 替代了旧的 `_pdfDirty` / `_lastPdfPath` 简单标记方案

### PDF 印章烘焙 (v2.0.4)

`extract_page_as_form_xobject()` 在提取 PDF 页面为 Form XObject 时,自动将页面标注 (印章/签章) 烘焙到内容流中。

- **标注发现**: 读取页面 `/Annots` 数组,跳过隐藏标注 (F bit 2)
- **外观提取**: `/AP` → `/N` (Normal appearance),经 `deep_copy_object` 完整迁移到输出文档
- **坐标映射**: 标注 Rect [x1,y1,x2,y2] → Form BBox 坐标系的平移+缩放变换矩阵
- **内容流顺序**: 后缀 `Q` (恢复图形状态) **先于**标注绘制命令 `q matrix /__AnnotN Do Q`

### 发票文件批量重命名 (v2.0.5)

**汇总表内嵌重命名面板**,支持预设模板 + 自定义字段,一键批量重命名发票磁盘文件。

- **入口**: 汇总表弹窗底部「🔄 重命名文件」按钮 → 展开内嵌面板
- **模板**: 4 个预设 (金额+销售方 / 销售方+号码 / 金额+日期+号码 / 自定义) + 字段勾选
- **分隔符**: 默认 `_`,可自定义
- **重名**: `resolveNameConflicts()` — 自动加 `_2`、`_3` 序号
- **执行**: `executeRename()` → `POST /api/v1/rename_file` 批量重命名
- **Rust**: `rename_file` handler (async + spawn_blocking),同盘 `fs::rename` (原子),跨盘 `copy + delete` fallback

### 发票汇总表导出 (v2.0.3)

**可编辑预览 + CSV 导出**,用于报销时生成发票明细汇总表。

- **入口**: 侧边栏左下角金额汇总旁 📊 按钮
- **弹窗**: 14 个字段按需勾选,列选择和备注持久化
- **编辑**: 金额/文本双击编辑,`setSummaryCellValue()` 回写 `fileObj`
- **合计**: 三种金额 (含税/不含税/税额) 分别汇总
- **导出**: `exportSummaryCsv()` — UTF-8 BOM + CRLF,`write_text_file` (async+spawn_blocking)
- **数据源**: `getCheckedFiles()` — 不含 `copies` 展开的已勾选文件列表

## 前端模块

| 文件 | 职责 |
|------|------|
| `frontend/js/api.js` | 统一 HTTP 通信层,SSE 监听,fetch 适配 |
| `frontend/js/app.js` | 主入口、状态管理(S)、文件加载、XML 数电票、设置持久化 |
| `frontend/js/ocr.js` | 发票字段提取、金额解析、中文大写解析、类型检测 |
| `frontend/js/layout.js` | 布局计算、预览渲染、单票调整拖拽 |
| `frontend/js/print.js` | 打印/导出、LayoutRenderRequest、智能缓存 (deepEqual) |

- 全部用 `var` 声明顶层变量 (避免与浏览器扩展/注入脚本冲突)
- 无模块打包,`index.html` 按顺序 `<script>` 加载
- **v2.1.0 改动**: `_isDesktop` 检测已删除,纯 `fetch()` 调用

## Feature Flag

- Cargo.toml 定义 `ocr` feature,`lib.rs` 按 `#[cfg(feature = "ocr")]` 条件启用 OCR 命令
- `ocr-rs` 依赖仅在 `--features ocr` 时编译,大幅缩短轻量版构建时间
- Docker 默认 build 包含 OCR,可构建 `Dockerfile.ocr` 变体

## 关键踩坑

### Tauri 2.x → Axum 迁移 (v2.1.0)
- **Tauri 已完全移除**: 仓库无 `tauri` / `tauri.conf.json` / `tauri-cli` 任何引用
- `<input>.click()` 在浏览器中工作正常,无需 `plugin:dialog|open` 兜底
- 后端 `async fn` handler 必须用 `tokio::task::spawn_blocking` 包装 CPU 密集操作
- 桌面版 `Tauri IPC` 替换为 HTTP POST,统一 `__api.call()` 入口
- `serde` 字段命名: HTTP 不走 Tauri 自动 camelCase 转换,Rust 结构体须显式 `#[serde(rename_all = "camelCase")]`

### Web 路径安全
- Web 模式**禁止**接收绝对路径,所有文件操作走 session 目录
- `Session::resolve_path()` 必须 `Path::canonicalize()` 后 `starts_with` session_dir
- 文件上传走 `multipart/form-data`,服务端用 UUID 重命名落盘

### 上传大小
- 100MB 限制由 nginx `client_max_body_size` + Axum `RequestBodyLimitLayer` 双重保护
- 早期 `tower-http limit` feature 引发 413 内部错误 (commit 45e700d),已移除并由 nginx 接管
- `generate_pdf` 含 base64 图片数据,**不要**对全局 body 设限,仅限 upload 端点

### 智能 PDF 缓存
- `deepEqual` 比较整个 `LayoutRenderRequest` 对象
- 保存 PDF 时先生成到临时目录 → `updatePdfCache(req, tempPath)` → `copy_file` 到用户路径

### OFD
- ImageMask 遮罩: 二值图合成主图 alpha 通道
- 自闭合标签不能用 `read_element_text()`
- CJK 拆字问题 (dzcp 格式): 需虚拟标签合成

### 进程生命周期
- 关闭时必须用 `TerminateProcess`,不能用 `process::exit(0)` (MNN/OCR 引擎死锁)

### PDFium 打印
- `libloading::Library` 不能在函数内创建,drop 时 DLL 卸载导致全局状态崩溃 → 用全局 `LazyLock<Mutex<Option<PdfiumState>>>` 持有
- **PDFium 不是线程安全的** → `with_pdfium()` 闭包 + Mutex 串行化 (RwLock 尝试已回退)
- `DEVMODEW` 嵌套匿名结构: `dm.Anonymous1.Anonymous1.dmCopies`,`dmDuplex` 是 `DEVMODE_DUPLEX(i16)`
- **SEH 保护**: 打印机驱动 GDI bug → `seh_wrapper.c` 用 `__try/__except` 捕获,fallback 到位图渲染
- **DEVMODE 截断**: `std::ptr::read` 只复制 `sizeof(DEVMODEW)` 丢弃驱动私有数据 → 返回完整 `Vec<u8>` 缓冲区

### 批量文字提取
- 多 PDF 文件场景下必须按 `pdfPath` 分组调用 `extract_pdf_texts`
- `extract_pdf_texts` 返回 `HashMap<u32, PdfTextResult>` keyed by pageIdx
- 批量失败时自动回退单页 `extract_pdf_text`,再失败则回退 OCR

### EXIF
- `image` crate 不自动应用 EXIF;6=90°CW, 8=90°CCW, 3=180°

### Web 部署
- SSE 必须 `proxy_buffering off` 否则进度推送卡死
- `TICKETCHAN_BIND_ADDR=0.0.0.0` 必须配 `TICKETCHAN_AUTH_TOKEN` 或外层鉴权

## Git 工作流

- 主分支: `master` (稳定) / `dev` (开发) / `feat/*` (功能)
- 当前活跃: `feat/web-cross-platform` (v2.1.0 重构,未合并)
- 小步提交,完成即 push
- 变动大时升版本打 tag 触发 CI
- 会话结束前确保无未提交变更
- `package.json` 不再存在,版本号统一在 `Cargo.toml` (`[package].version`)

## 用户偏好

- 简洁直接,对 Bug 极度敏感,全面修复原则
- 不要主动编译 (耗时),等明确指令
- 分析任务绝对不可修改代码,必须先确认方案

## Release 检查清单

每次 release 前必须完成以下文档更新:

1. **README.md**: 更新功能描述、技术栈版本、部署命令,确保与当前版本一致
2. **CHANGELOG.md**: 补充新版本更新日志,包含新功能/修复/优化/依赖变更等
3. **AGENTS.md**: 更新版本号、架构要点 (如有变更)
4. **其他文档**: 如有新增配置/命令/架构变更,同步更新对应文档

## v2.1.0 重大变更摘要 (对比 v2.0.7)

| 项 | v2.0.7 | v2.1.0 |
|---|---|---|
| 技术栈 | Tauri 2.x + Axum (模式 B) | 纯 Axum (无 Tauri) |
| 目录结构 | `src-tauri/src/` | 根 `src/` |
| 前端位置 | `src/` (Tauri 嵌入) | `frontend/` (浏览器静态) |
| 桌面打包 | NSIS + 便携版 (双版本) | **已废弃** |
| 通信层 | Tauri IPC + HTTP (适配) | 纯 HTTP fetch |
| 二进制 | `ticketchan.exe` + `ticketchan-server.exe` | 仅 `ticketchan-server` |
| `is_desktop` 字段 | 存在 (v2.0.7 适配层) | **已删除** |
| `_isDesktop` 前端检测 | 存在 | **已删除** |
| SumatraPDF 打印 | 第四种模式 | **已删除** (3 模式) |
| PDFium 锁 | Mutex | RwLock (实验) → **回退 Mutex** (崩溃) |
| `package.json` | 存在 | **已删除** (版本号迁移到 Cargo.toml) |
| `tauri.conf.json` | 存在 (2 份) | **已删除** |
| `default-run` | `ticketchan` (桌面) | (无,显式 `--bin`) |
| 编译命令 | `npm run dev/build` | `cargo run --bin ticketchan-server` |
| Docker | 无 | **新增** (多阶段构建) |
| systemd | 无 | **新增** (`ticketchan.service`) |
| Nginx 配置 | 无 | **新增** (`nginx.conf`) |
