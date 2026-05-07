# 发票批量打印工具 — Code Wiki

> **项目名称**: fapiao-print  
> **版本**: v1.9.4  
> **许可证**: MIT  
> **技术栈**: Tauri 2 + Rust + Vanilla JS  
> **目标平台**: Windows (x64)

---

## 目录

1. [项目概述](#1-项目概述)
2. [整体架构](#2-整体架构)
3. [目录结构](#3-目录结构)
4. [前端模块详解](#4-前端模块详解)
5. [后端模块详解](#5-后端模块详解)
6. [OFD 引擎模块详解](#6-ofd-引擎模块详解)
7. [数据流与交互](#7-数据流与交互)
8. [依赖关系](#8-依赖关系)
9. [构建与运行](#9-构建与运行)
10. [CI/CD 流程](#10-cicd-流程)
11. [关键设计决策](#11-关键设计决策)

---

## 1. 项目概述

**发票批量打印工具**是一款 Windows 桌面应用，用于批量处理和打印中国电子发票。支持 PDF、OFD、图片（JPG/PNG/BMP/WebP/TIFF）等多种发票格式，提供排版预览、OCR 识别、白边裁剪、水印添加等功能。

### 核心特性

| 特性 | 说明 |
|------|------|
| 多格式支持 | PDF、OFD（国标 GB/T 33190-2016）、JPG/PNG/BMP/WebP/TIFF |
| 批量排版 | 自定义行列布局（1×1 ~ 10×10），支持 A4/A5/B5/Letter/Legal 及自定义纸张 |
| OCR 识别 | 可选 PP-OCRv5 引擎，自动提取发票金额、销售方、发票号码等结构化信息 |
| OFD 解析 | 纯 Rust 实现的 OFD 解析器，SVG 矢量渲染 + XML 元数据提取，无需 OCR |
| PDF 文字提取 | 直接解析 PDF 内容流文字层（~5ms/页），无需 OCR |
| 打印方式 | 对话框打印 / 静默直打（Windows Print Spooler API） |
| 白边裁剪 | Rust 后端图像处理，10-50 倍速于前端 Canvas |
| 水印 | 自定义文字、透明度、角度、颜色、大小 |
| 双版本发布 | 轻量版（~8MB，无 OCR）+ OCR 版（含 PP-OCRv5 模型） |

---

## 2. 整体架构

```
┌─────────────────────────────────────────────────────────────┐
│                    Tauri 2 Desktop App                       │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐   │
│  │              Frontend (WebView2 / Chromium)           │   │
│  │                                                       │   │
│  │  ┌────────┐ ┌────────┐ ┌────────┐ ┌────────┐       │   │
│  │  │ app.js │ │ ocr.js │ │layout.js│ │print.js│       │   │
│  │  │ 主入口  │ │OCR识别 │ │排版渲染 │ │打印PDF │       │   │
│  │  └────────┘ └────────┘ └────────┘ └────────┘       │   │
│  │         │         │          │          │             │   │
│  │         └─────────┴──────────┴──────────┘             │   │
│  │                      │ invoke()                       │   │
│  └──────────────────────┼───────────────────────────────┘   │
│                         │ Tauri IPC                         │
│  ┌──────────────────────┼───────────────────────────────┐   │
│  │              Backend (Rust / Native)                  │   │
│  │                      │                                │   │
│  │  ┌───────────────────┴──────────────────────┐        │   │
│  │  │            lib.rs (Tauri Commands)        │        │   │
│  │  │  open_invoice_files / render_pdf_pages    │        │   │
│  │  │  generate_pdf_from_layout / print_pdf_file│        │   │
│  │  │  ocr_image / ocr_pdf_page (OCR feature)  │        │   │
│  │  │  parse_ofd / open_ofd_images             │        │   │
│  │  └───────────┬──────────────┬───────────────┘        │   │
│  │              │              │                         │   │
│  │  ┌───────────┴───┐  ┌──────┴──────┐                 │   │
│  │  │  pdf_engine   │  │  ofd-engine │                 │   │
│  │  │  PDF渲染/生成  │  │  OFD解析渲染 │                 │   │
│  │  │  OCR/打印/图像 │  │  发票数据提取 │                 │   │
│  │  └───────────────┘  └─────────────┘                 │   │
│  └──────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

### 架构特点

- **前后端分离**: 前端 Vanilla JS 负责 UI 交互和预览渲染，后端 Rust 负责所有计算密集型操作
- **IPC 通信**: 通过 Tauri 2 的 `invoke()` 机制，前端调用后端 Tauri Command
- **Feature Flag**: OCR 功能通过 Cargo feature `ocr` 控制，编译时决定是否包含
- **进程安全**: 全局 `SHUTTING_DOWN` 原子标志，确保关闭时所有长操作（PDF渲染/OCR）能及时中止

---

## 3. 目录结构

```
fapiao-print/
├── src/                          # 前端源码
│   ├── index.html                # 主页面 HTML
│   ├── styles.css                # 全局样式
│   ├── app.js                    # 主入口：文件管理、UI交互、状态管理
│   ├── ocr.js                    # OCR识别：发票信息提取、坐标分析
│   ├── layout.js                 # 排版计算：布局算法、预览渲染、拖拽交互
│   └── print.js                  # 打印/PDF：构建请求、打印控制、浏览器回退
├── src-tauri/                    # Tauri/Rust 后端
│   ├── Cargo.toml                # Rust 依赖配置
│   ├── tauri.conf.json           # Tauri 轻量版配置
│   ├── tauri.ocr.conf.json       # Tauri OCR版配置（含模型资源）
│   ├── capabilities/
│   │   └── default.json          # Tauri 权限声明
│   ├── src/
│   │   ├── main.rs               # 程序入口
│   │   ├── lib.rs                # Tauri Command 注册 + 应用启动
│   │   └── pdf_engine.rs         # 核心引擎：PDF渲染/生成/OCR/打印
│   ├── ofd-engine/               # OFD 解析子 crate
│   │   ├── Cargo.toml
│   │   └── src/
│   │       └── lib.rs            # OFD 解析器 + SVG 渲染器
│   ├── models/                   # OCR 模型文件（仅OCR版）
│   │   ├── PP-OCRv5_mobile_det.mnn
│   │   ├── PP-OCRv5_mobile_rec.mnn
│   │   └── ppocr_keys_v5.txt
│   └── build.rs                  # Tauri 构建脚本
├── scripts/
│   ├── build-all.js              # 全量构建脚本（4个产物）
│   └── bump-version.js           # 版本号同步脚本
├── .github/
│   └── workflows/
│       └── build.yml             # GitHub Actions CI/CD
└── package.json                  # Node.js 项目配置
```

---

## 4. 前端模块详解

### 4.1 app.js — 主入口模块

**职责**: 应用状态管理、文件上传/加载、文件列表渲染、UI 交互、设置管理

#### 全局状态对象 `S`

```javascript
var S = {
  files: [],           // 发票文件对象数组
  currentPage: 0,      // 当前预览页
  totalPages: 0,       // 总页数
  viewZoom: 0,         // 缩放比例（0=自适应）
  layout: { cols: 1, rows: 1, orient: 'landscape' },  // 排版布局
  editIdx: -1,         // 当前编辑文件索引
  selectedSlot: -1,    // 当前选中槽位
  amtMode: 'tax',      // 金额显示模式
  feat: {              // 功能开关
    cutline, number, border, trimWhite, watermark,
    collate, duplex, pageNum, printDate, confirmPrint,
    autoOpenPdf, ocrEnabled
  }
};
```

#### 文件对象 `createFileObj()`

每个发票文件的核心数据结构，包含：

| 字段 | 类型 | 说明 |
|------|------|------|
| `id` | string | 唯一标识 |
| `name/size/type` | string/number/string | 文件基本信息 |
| `checked` | boolean | 是否选中参与打印 |
| `previewUrl` | string | 预览图 data URL |
| `copies` | number | 打印份数 |
| `rotation` | number | 旋转角度（0/90/180/270） |
| `amountTax/amountNoTax/taxAmount` | number | 含税价/不含税价/税额 |
| `sellerName/sellerCreditCode` | string | 销售方信息 |
| `invoiceNo/invoiceDate` | string | 发票号码/开票日期 |
| `buyerName/buyerCreditCode` | string | 购买方信息 |
| `ow/oh` | number | 原始图像尺寸 |
| `slotScale/slotOffsetX/slotOffsetY` | number | 单票微调参数 |
| `_filePath` | string | 磁盘路径（Rust直接读取，跳过base64） |
| `_pdfPath/_pdfPageIdx` | string/number | PDF源信息（零IPC往返OCR） |
| `_ocrText/_ocrPending` | string/boolean | OCR状态 |

#### 关键函数

| 函数 | 说明 |
|------|------|
| `triggerUpload()` | 通过 Tauri 对话框选择文件 |
| `processFileDataList(fileDataList)` | 批量加载：骨架占位 → 并行加载 → 顺序渲染 |
| `processFilesIncremental(paths)` | 增量加载：逐文件读取，即时预览 |
| `loadFileFromDataUrlFast(fd)` | 快速加载：先显示预览，OCR 后台队列执行 |
| `renderFileList()` | 渲染侧边栏文件列表 |
| `openInvModal(i)` | 打开发票详情编辑弹窗 |
| `updateAmountSummary()` | 更新金额统计汇总 |
| `processTrim()` | 调用 Rust 后端裁剪白边 |

#### OCR 队列管理

```javascript
var _ocrQueue = [];          // OCR 任务队列
var _ocrRunning = 0;         // 当前运行中的 OCR 任务数
var _ocrMaxConcurrent = 1;   // 最大并发数（MNN引擎Mutex限制）
```

- `applyOcrAsync(fileObj, dataUrl)`: 将 OCR 任务加入队列
- `_drainOcrQueue()`: 消费队列，控制并发
- `_onOcrTaskDone()`: 任务完成回调，更新进度

### 4.2 ocr.js — OCR 识别模块

**职责**: OCR 结果处理、发票信息提取（坐标分析 + 文本正则）

#### 核心提取策略

采用**坐标优先**（Coordinate-first）策略，利用 PP-OCRv5 的高精度 bbox 输出：

1. **文本提取**（`_extractByText`）: 基于 OCR 文本中的键值对模式（如"发票号码：XXX"）
2. **坐标提取**（`extractByCoordinates`）: 基于词位置的区域分类和邻近匹配
3. **金额验证**（`_extractAmountsByText`）: 数学验证法——找两个金额之和等于含税价

#### 区域分类

```
发票布局（归一化坐标 0~1）:
  ny 0.00~0.15:  标题 + 发票号码 + 开票日期
  ny 0.15~0.35:  购买方(nx<0.5) | 销售方(nx>0.5)
  ny 0.35~0.45:  明细表头
  ny 0.45~0.60:  明细行
  ny 0.60~0.70:  合计行
  ny 0.70~0.80:  价税合计
  ny 0.80~1.00:  备注 + 开票人
```

#### 关键函数

| 函数 | 说明 |
|------|------|
| `applyOcrResult(fileObj, ocrResult)` | 将 OCR 结果应用到文件对象 |
| `applyPdfTextResult(fileObj, pdfTextResult)` | 将 PDF 文字提取结果应用到文件对象 |
| `applyOcr(fileObj, dataUrl, filePath)` | 图片 OCR（调用 Rust `ocr_image`） |
| `applyOcrPdfPage(fileObj)` | PDF 页 OCR（调用 Rust `ocr_pdf_page`，零IPC） |
| `extractByCoordinates(ocrResult)` | 坐标优先发票信息提取 |
| `_extractByText(fullText)` | 文本正则提取（发票号、日期、买卖方） |
| `_extractAmountsByText(fullText)` | 文本金额提取（含数学验证） |
| `normalizeOcrCurrency(s)` | OCR 货币符号纠错（¥↔1 误识别） |
| `cleanOcrAmtStr(raw)` | 清洗 OCR 金额字符串 |
| `isLikelyYearOrDate(val, rawText)` | 判断是否为年份/日期（排除误识别） |

#### OCR 误识别纠正

OCR 常见误识别模式及纠正策略：

| 误识别 | 原始 | 纠正 | 策略 |
|--------|------|------|------|
| "1" → "¥" | "¥¥72.68" | "¥172.68" | `normalizeOcrCurrency()` |
| "¥" → "1" | "1317.00" | "¥317.00" | 关键词后4+位数字检测 |
| 全角→半角 | "１３１７．００" | "1317.00" | `_normText()` |
| CJK空格 | "购 买 方" | "购买方" | CJK间空格折叠 |

### 4.3 layout.js — 排版计算模块

**职责**: 布局算法、预览渲染、单票微调交互

#### 核心函数

| 函数 | 说明 |
|------|------|
| `calculateLayout(settings, pxPerMm)` | 统一布局计算——纯函数，返回槽位位置/裁切线 |
| `renderPage(pageFiles, pi, total, s)` | HTML/CSS 预览渲染 |
| `getRotation(fileObj, slot, settings)` | 计算旋转角度（支持自动/手动/0°/90°/180°/270°） |
| `initSlotInteraction()` | 绑定槽位拖拽/缩放交互 |
| `buildTransformString(f, s, slot)` | 构建 CSS transform 字符串 |

#### 布局算法

```
槽位宽度 sw = (纸张宽度 - 列数×(左外边距+右外边距) - (列数-1)×水平间距) / 列数
槽位高度 sh = (纸张高度 - 行数×(上外边距+下外边距) - (行数-1)×垂直间距) / 行数
槽位位置 x = 左外边距 + 列号 × (sw + 左外边距 + 右外边距 + 水平间距)
           y = 上外边距 + 行号 × (sh + 上外边距 + 下外边距 + 垂直间距)
```

#### 单票微调

- **拖拽移动**: 鼠标拖拽发票图片，实时更新 `slotOffsetX/slotOffsetY`（mm单位）
- **角点缩放**: 拖拽四角手柄，实时更新 `slotScale`（0.2x ~ 2.0x）
- **应用全部**: 将当前微调参数应用到所有发票

### 4.4 print.js — 打印/PDF 模块

**职责**: 构建 Rust 布局请求、PDF 生成/打印控制、浏览器回退

#### 核心函数

| 函数 | 说明 |
|------|------|
| `buildLayoutRequest(files, settings)` | 构建 `LayoutRenderRequest`（文件去重、旋转计算、逐份/逐份排序） |
| `doPrint()` | 打印发票（支持缓存复用、确认对话框、进度监听） |
| `savePdf()` | 保存为 PDF 文件（支持自动打开） |
| `fallbackPrint(files, s)` | 浏览器回退：新窗口打印 |
| `refreshPrinters()` | 刷新系统打印机列表 |
| `listenPdfProgress()` | 监听 Rust 端 PDF 生成进度事件 |

#### LayoutRenderRequest 结构

```javascript
{
  files: [{ ow, oh, rotation, filePath, dataUrl, sourceType, pdfPath, pdfPageIdx }],
  pages: [{ slots: [{ fileIndex, rotation, scale, offsetX, offsetY }] }],
  settings: { paperW, paperH, cols, rows, margins, gap, fitMode, ... }
}
```

#### PDF 缓存机制

- `_lastPdfPath`: 上次生成/保存的 PDF 路径
- `_pdfDirty`: PDF 内容是否已变更
- 未变更时直接调用 `print_pdf_file` 复用，跳过 PDF 重新生成

---

## 5. 后端模块详解

### 5.1 main.rs — 程序入口

```rust
fn main() {
    app_lib::run();
}
```

仅调用 `app_lib::run()`，设置 `windows_subsystem = "windows"` 隐藏控制台窗口。

### 5.2 lib.rs — Tauri Command 注册

**职责**: 定义所有 Tauri Command，注册插件，处理窗口事件

#### Tauri Commands 列表

| Command | Feature | 说明 |
|---------|---------|------|
| `open_invoice_files` | - | 读取发票文件（返回缩略图+元数据） |
| `parse_ofd` | - | OFD 矢量解析（SVG + 结构化数据） |
| `open_ofd_images` | - | OFD 位图回退（提取图片） |
| `get_printers` | - | 获取系统打印机列表 |
| `render_pdf_pages` | - | WinRT 渲染 PDF 页面为 PNG |
| `render_and_ocr_pdf` | ocr | 渲染+OCR 一步完成（避免IPC往返） |
| `ocr_image` | ocr | 图片 OCR 识别 |
| `ocr_pdf_page` | ocr | PDF 单页 OCR（零IPC往返） |
| `check_ocr_available` | - | 检查 OCR 功能是否可用 |
| `extract_pdf_text` | - | 提取 PDF 文字层（~5ms/页） |
| `generate_pdf_from_layout` | - | Rust 端排版+PDF 生成（异步，带进度） |
| `print_pdf_file` | - | 直接打印已有 PDF |
| `trim_image` | - | 裁剪白边 |
| `open_file/open_url` | - | ShellExecute 打开文件/URL |
| `get_app_version` | - | 获取编译时版本号 |
| `get_config` | - | 获取后端配置（DPI等） |
| `get_temp_dir` | - | 获取系统临时目录 |
| `show_window` | - | 显示主窗口 |

#### 窗口关闭处理

```rust
CloseRequested => {
    api.prevent_close();                    // 阻止默认关闭
    SHUTTING_DOWN.store(true, ...);         // 设置关闭标志
    win.eval("window._tauriCleanup()");     // 通知前端清理
    // 300ms 后 TerminateProcess 强制退出
    // 原因: ExitProcess 的 DLL_PROCESS_DETACH 可能死锁（MNN/OCR引擎）
}
```

#### 拖放处理

Rust 端拦截 `DragDrop` 事件，过滤有效扩展名（pdf/jpg/jpeg/png/bmp/webp/tiff/tif/ofd），通过 `eval()` 调用前端 `window._tauriFileDrop(paths)`。

#### 打印实现

- **对话框打印**: `ShellExecuteW("print", ...)` + `SW_SHOWNORMAL`
- **静默直打**: 优先使用 Windows Print Spooler API（`OpenPrinterW` → `StartDocPrinterW` → `WritePrinter` → `EndDocPrinterW`），失败时回退到 `ShellExecuteW("printto", ...)` + `SW_HIDE`

### 5.3 pdf_engine.rs — 核心引擎

**职责**: PDF 渲染、PDF 生成、OCR、图像处理、打印

#### 常量

| 常量 | 值 | 说明 |
|------|----|------|
| `RENDER_DPI` | 300 | 渲染 DPI（前后端必须一致） |
| `MM_TO_PT` | 72.0/25.4 | 毫米到 PDF 点的转换系数 |

#### 核心类型

| 类型 | 说明 |
|------|------|
| `PdfResult` | PDF 生成/打印结果（success, message, pdf_path） |
| `PrinterInfo` | 打印机信息（name, is_default） |
| `FileData` | 返回前端的文件数据（name, ext, size, dataUrl, path, origW, origH） |
| `RenderedPage` | 渲染后的 PDF 页面（index, imageDataUrl, width, height, renderDpi） |
| `RenderedOcrPage` | 渲染+OCR 的 PDF 页面（OCR feature only） |
| `OcrResult` | OCR 识别结果（text, lines, imgW, imgH） |
| `LayoutRenderRequest` | 前端发来的排版请求（files, pages, settings） |
| `PdfTextResult` | PDF 文字层提取结果 |
| `ImageSource` | 图像来源（Decoded / JpegPassthrough） |

#### PDF 渲染（Windows WinRT）

使用 `windows::Data::Pdf::PdfDocument` API：

1. `StorageFile::GetFileFromPathAsync` 加载文件
2. `PdfDocument::LoadFromFileAsync` 解析 PDF
3. 逐页 `GetPage` → `RenderWithOptionsToStreamAsync` 渲染为 PNG
4. 自适应 DPI：确保最长边 ≥ 3508px（A4@300DPI）

#### PDF 生成（`generate_pdf_from_layout`）

1. 解析 `LayoutRenderRequest` 中的文件和页面规格
2. 使用 `printpdf` crate 创建 PDF 文档
3. 逐页：计算槽位 → 加载图像 → 旋转 → 缩放 → 嵌入
4. 图像嵌入策略：
   - **image source**: 从磁盘/内存读取，解码为 `DynamicImage`
   - **pdf-page source**: WinRT 渲染 PDF 页 → PNG → 解码
   - **ofd-page source**: 使用前端传来的 dataUrl
5. 通过 `ProgressFn` 回调报告进度（`pdf-progress` 事件）
6. 异步执行：`tokio::task::spawn_blocking` 避免阻塞 IPC 线程

#### EXIF 方向处理

```rust
fn read_exif_orientation(bytes: &[u8]) -> u32  // 读取 EXIF 方向标签
fn apply_exif_orientation(img, orientation)      // 应用旋转/翻转
```

支持所有 8 种 EXIF 方向值，确保手机拍摄的照片在 PDF 中正确显示。

#### 白边裁剪

```rust
pub fn trim_white_edges(img: &DynamicImage, threshold: u8) -> DynamicImage
```

从四边向内扫描，裁剪掉亮度 ≥ threshold（默认245）的白色边距。

#### PDF 文字层提取（`extract_pdf_text`）

使用 `lopdf` crate 解析 PDF 内容流，提取文字及其坐标。~5ms/页，无需 OCR，轻量版也可用。

#### OCR（Feature: ocr）

使用 `ocr-rs` crate（基于 MNN 推理引擎 + PP-OCRv5 模型）：

- `ocr_image(data_url, file_path)`: 图片 OCR，支持磁盘直读
- `ocr_pdf_page(pdf_path, page_index, dpi)`: PDF 页 OCR，零IPC往返
- `render_and_ocr_pdf(pdf_path, dpi)`: 渲染+OCR 一步完成

---

## 6. OFD 引擎模块详解

### ofd-engine/src/lib.rs

**职责**: OFD（Open Fixed-layout Document）格式解析与 SVG 渲染，支持中国电子发票结构化数据提取

#### OFD 格式概述

OFD 是中国国家标准 GB/T 33190-2016 定义的版式文档格式，本质是 ZIP 压缩包，包含 XML 页面描述和图片资源。

```
典型 OFD 结构:
  OFD.xml                    — 根元数据 + CustomData（发票字段）
  Doc_0/Document.xml         — 文档结构（模板页 + 内容页）
  Doc_0/PublicRes.xml        — 字体定义 + DrawParam 继承
  Doc_0/DocumentRes.xml      — 图片资源映射
  Doc_0/Tags/CustomTag.xml   — 语义字段→TextObject ID 映射
  Doc_0/Pages/Page_0/Content.xml  — 页面内容
  Doc_0/Tpls/Content.xml     — 模板内容（背景层）
  Doc_0/Annots/Page_0/Annotation.xml — 注释（水印层）
  Doc_0/Res/                 — 图片资源文件
```

#### 公共类型

| 类型 | 说明 |
|------|------|
| `OfdInvoiceInfo` | 发票数据（发票号、日期、买卖方、金额等） |
| `OfdResult` | 解析结果（SVG + invoiceInfo + 页面尺寸） |
| `OfdExtractedImage` | 提取的图片（位图回退路径） |

#### 内部结构

| 结构 | 说明 |
|------|------|
| `OfdFont` | 字体定义（ID、字体名、族名） |
| `OfdDrawParam` | 绘制参数（线宽、描边色、填充色、继承链） |
| `OfdTextObject` | 文本对象（边界、字体、大小、CTM、DeltaX、颜色等） |
| `OfdPathObject` | 路径对象（边界、线宽、颜色、缩写路径数据） |
| `OfdImageObject` | 图像对象（边界、资源ID、CTM、ImageMask） |

#### 解析流程（`parse_ofd_file`）

```
1. 打开 OFD ZIP → 读取 OFD.xml
2. 解析 CustomData（快速发票字段提取）
3. 读取 Document.xml → 获取模板/页面路径
4. 解析 PublicRes.xml → 字体 + DrawParam 继承链
5. 解析 DocumentRes.xml → 图片资源映射
6. 加载图片原始字节（延迟生成 data URL，支持 ImageMask 合成）
7. 解析模板内容（背景层：网格线、静态标签）
8. 解析页面内容（数据层：发票填写内容）
9. 解析注释（水印层，处理 Appearance 偏移）
10. 解析 CustomTag.xml → 语义字段映射
11. 提取发票信息（CustomData → CustomTag → 文本回退）
12. 构建 SVG（3层：template + content + annotations）
```

#### SVG 渲染

- 缩放因子: 3.5（1mm → 3.5 SVG 单位）
- 三层结构: `<g id="template">` + `<g id="content">` + `<g id="annotations">`
- 文本渲染: 使用 `<tspan x>` 绝对定位，正确处理 DeltaX 字间距
- 路径渲染: OFD AbbreviatedData → SVG path data 命令转换
- 字体映射: PostScript 名称 → CSS font-family（含 CJK 回退链）
- DrawParam 继承: 遍历 Relative 链解析颜色/线宽默认值

#### 发票信息提取（三级策略）

1. **CustomData**（OFD.xml）: 最可靠，直接从元数据读取
2. **CustomTag**（CustomTag.xml）: 语义字段 → TextObject ID → 文本内容
3. **文本回退**（`extract_invoice_from_text`）: 扫描文本对象，模式匹配标签和值

#### 位图回退（`extract_ofd_images`）

当矢量解析失败时，从 ZIP 中提取图片：
- 路径过滤：排除 Seals/、Signs/ 目录
- 尺寸过滤：优先 ≥500px 长边（完整发票页），过滤二维码/印章
- 逐页去重：每页保留最大图片

#### ImageMask 合成

OFD 的 ImageMask 机制：主图像 + 蒙版图像 → 合成 RGBA（白色=不透明，黑色=透明）。

---

## 7. 数据流与交互

### 7.1 文件加载流程

```
用户拖放/选择文件
    │
    ▼
Rust: open_invoice_files(paths)
    │  ├─ 图片: 读取 → 生成缩略图(600px) → base64 dataUrl
    │  └─ PDF: 读取 → 返回元数据(path/size)
    ▼
前端: processFileDataList()
    │  ├─ 创建骨架占位符（即时反馈）
    │  ├─ 并行加载 → 顺序渲染
    │  ▼
    ├─ 图片: new Image() → 预览
    │   └─ applyOcrAsync() → OCR队列
    │
    ├─ PDF: invoke('render_pdf_pages') → 预览
    │   ├─ invoke('extract_pdf_text') → 文字提取(~5ms)
    │   └─ applyOcrAsync() → OCR队列(扫描PDF回退)
    │
    └─ OFD: invoke('parse_ofd') → SVG → svgToPngDataUrl() → 预览
        └─ 结构化数据直接提取，无需OCR
        └─ 失败时: invoke('open_ofd_images') → 位图回退
```

### 7.2 打印流程

```
用户点击"打印"
    │
    ▼
前端: doPrint()
    │  ├─ 确认对话框（如启用）
    │  ├─ PDF缓存命中? → print_pdf_file() → 完成
    │  ▼
    ├─ buildLayoutRequest() → 构建请求
    │   ├─ 文件去重（filePath/prewiewUrl 为 key）
    │   ├─ 份数展开（逐份/逐份排序）
    │   └─ 旋转计算（auto/手动）
    │
    ▼
Rust: generate_pdf_from_layout()
    │  ├─ spawn_blocking（异步，不阻塞IPC）
    │  ├─ 逐页: 计算槽位 → 加载图像 → 旋转 → 缩放 → 嵌入PDF
    │  ├─ 进度回调 → emit('pdf-progress')
    │  ▼
    ├─ 对话框打印: ShellExecuteW("print")
    └─ 静默直打: Spooler API → WritePrinter
```

### 7.3 OCR 数据流

```
┌──────────────────────────────────────────────────────┐
│ 图片文件                                              │
│  applyOcr(fileObj, dataUrl, filePath)                 │
│  → Rust: ocr_image(dataUrl, filePath)                 │
│    ├─ 有 filePath: Rust 直接读磁盘（跳过base64编解码）  │
│    └─ 无 filePath: 前端降采样 → base64 IPC → Rust解码  │
├──────────────────────────────────────────────────────┤
│ PDF 文件                                              │
│  applyOcrPdfPage(fileObj)                             │
│  → Rust: ocr_pdf_page(pdfPath, pageIndex)             │
│    └─ Rust 内部: 渲染 → 解码 → OCR → 返回结果          │
│       （零IPC往返，避免 base64 编解码链）               │
├──────────────────────────────────────────────────────┤
│ 结果处理                                              │
│  applyOcrResult(fileObj, ocrResult)                   │
│  → extractByCoordinates(ocrResult)                    │
│    ├─ _extractByText() → 文本正则提取                  │
│    ├─ _extractAmountsByText() → 金额数学验证           │
│    └─ 坐标回退 → 区域分类 → 邻近匹配                   │
└──────────────────────────────────────────────────────┘
```

---

## 8. 依赖关系

### 8.1 Rust 依赖（Cargo.toml）

| 依赖 | 版本 | 用途 |
|------|------|------|
| `tauri` | 2 | 桌面应用框架 |
| `tauri-plugin-dialog` | 2 | 文件选择/保存对话框 |
| `tauri-plugin-fs` | 2 | 文件系统访问 |
| `tauri-plugin-log` | 2 | 日志（仅debug模式） |
| `tauri-plugin-shell` | 2 | Shell 命令 |
| `serde` / `serde_json` | 1 | 序列化/反序列化 |
| `image` | 0.25 | 图像解码/编码/处理 |
| `printpdf` | 0.9 | PDF 生成 |
| `lopdf` | 0.39 | PDF 解析（文字层提取） |
| `rayon` | 1.10 | 并行计算 |
| `ab_glyph` | 0.2 | 字体渲染（水印文字） |
| `base64` | 0.22 | Base64 编解码 |
| `kamadak-exif` | 0.5 | EXIF 方向读取 |
| `flate2` | 1.1 | 压缩/解压 |
| `ocr-rs` | 2.2 | OCR 引擎（可选 feature） |
| `ofd-engine` | local | OFD 解析子 crate |
| `windows` | 0.58 | Windows API（WinRT PDF/Shell/Spooler/COM） |

### 8.2 OFD 引擎依赖

| 依赖 | 版本 | 用途 |
|------|------|------|
| `zip` | 2 | ZIP 压缩包解析（OFD 本质是 ZIP） |
| `quick-xml` | 0.37 | XML 解析（OFD 页面描述） |
| `serde` | 1 | 序列化 |
| `image` | 0.25 | 图片解码/合成 |
| `base64` | 0.22 | Base64 编码 |
| `log` | 0.4 | 日志 |

### 8.3 前端依赖

| 依赖 | 版本 | 用途 |
|------|------|------|
| `@tauri-apps/api` | ^2 | Tauri 前端 API |
| `@tauri-apps/cli` | ^2 | Tauri CLI 工具 |

> **注意**: 前端使用 Vanilla JS（无框架），零运行时依赖。

### 8.4 依赖关系图

```
package.json (Node.js)
  └── @tauri-apps/cli → tauri dev / tauri build

src-tauri/Cargo.toml (Rust)
  ├── tauri (框架)
  ├── printpdf (PDF生成)
  ├── lopdf (PDF解析)
  ├── image (图像处理)
  ├── ocr-rs (可选OCR)
  ├── windows (WinRT/Shell/Spooler)
  └── ofd-engine (子crate)
        ├── zip (ZIP解析)
        ├── quick-xml (XML解析)
        ├── image (图像)
        └── base64 (编码)
```

---

## 9. 构建与运行

### 9.1 开发环境

```bash
# 安装依赖
npm ci

# 轻量版开发
npm run dev

# OCR版开发
npm run dev:ocr
```

### 9.2 生产构建

```bash
# 轻量版构建
npm run build

# OCR版构建
npm run build:ocr

# 全量构建（4个产物）
npm run build:all
```

### 9.3 全量构建产物

| 产物 | 格式 | 说明 |
|------|------|------|
| `发票打印工具_{版本}_x64-setup.exe` | NSIS 安装包 | 轻量版安装包 |
| `发票打印工具_{版本}_x64_绿色版.exe` | 单文件 EXE | 轻量版绿色便携 |
| `发票打印工具_{版本}_x64_OCR版-setup.exe` | NSIS 安装包 | OCR 版安装包 |
| `发票打印工具_{版本}_x64_OCR绿色版.zip` | ZIP 压缩包 | OCR 版绿色便携（exe + models/） |

### 9.4 版本管理

```bash
# 设置新版本号并同步到所有配置文件
node scripts/bump-version.js 1.9.5

# 同步当前 package.json 版本号
node scripts/bump-version.js
```

同步范围: `package.json` → `Cargo.toml` → `tauri.conf.json`

### 9.5 前置条件

- Node.js 20+
- Rust 1.77.2+（`rust-toolchain` 稳定版）
- CMake 3.x（OCR 版构建 MNN 需要）
- Windows SDK（WinRT PDF 渲染）
- WebView2 Runtime（Tauri 2 依赖）

---

## 10. CI/CD 流程

### GitHub Actions（`.github/workflows/build.yml`）

**触发条件**: 推送 `v*` 标签 或 手动触发

**流程**:

1. **Checkout** — 检出代码
2. **Setup UTF-8** — 设置控制台编码（确保中文文件名正确）
3. **Setup Node.js 20** — 安装前端依赖
4. **Setup Rust** — 安装 Rust 工具链 + 缓存
5. **Install CMake** — OCR 版构建依赖
6. **Build all** — 执行 `npm run build:all`（4个产物）
7. **Verify artifacts** — 验证 4 个产物是否齐全
8. **Rename fallback** — CI 编码问题导致中文文件名丢失时强制重命名
9. **Create Release** — 使用 `softprops/action-gh-release@v2` 创建 GitHub Release 并上传产物

---

## 11. 关键设计决策

### 11.1 为什么用 Vanilla JS 而非框架？

- 应用为单页面工具，交互模式固定
- 无需组件复用、路由、状态管理库
- 减少构建复杂度和包体积
- 直接操作 DOM 性能更可控

### 11.2 为什么 OCR 是可选 Feature？

- OCR 模型（PP-OCRv5）约 15MB，显著增加安装包体积
- 大量用户只需打印，不需要识别
- 编译时决定：`Cargo.toml` 的 `[features] ocr = ["ocr-rs"]`
- 运行时检测：`check_ocr_available()` 命令

### 11.3 为什么用 TerminateProcess 而非 ExitProcess？

- `ExitProcess` 先杀所有线程，再执行 `DLL_PROCESS_DETACH`
- MNN/OCR 引擎在 `DLL_PROCESS_DETACH` 中可能死锁
- `TerminateProcess` 跳过 `DLL_PROCESS_DETACH`，立即终止进程
- 300ms 延迟给待处理 I/O 完成时间

### 11.4 为什么 PDF 生成移到 Rust？

- 前端 Canvas → `toDataURL()` → base64 IPC → Rust 解码，链路长且慢
- Rust 端直接读文件、解码、嵌入 PDF，跳过 base64 编解码
- 支持磁盘直读（`filePath`），节省 ~30% 数据传输
- `spawn_blocking` 异步执行，不阻塞 IPC 线程

### 11.5 为什么 OFD 用 SVG 而非位图？

- SVG 矢量渲染保留文字清晰度，任意缩放不失真
- 直接从 XML 提取结构化数据，无需 OCR
- 文件更小，渲染更快
- 位图仅作为回退方案

### 11.6 IPC 优化策略

| 优化 | 说明 |
|------|------|
| 磁盘直读 | 有 `filePath` 时 Rust 直接读文件，跳过 base64 编解码 |
| 零IPC OCR | `ocr_pdf_page` 在 Rust 内部完成渲染+OCR，避免往返 |
| 缩略图 | 图片文件返回 600px 缩略图而非全尺寸，减少 IPC 数据量 |
| OCR 降采样 | 前端降采样到 1280px 再发送 OCR，减少传输量 |
| PDF 缓存 | 未修改时复用上次生成的 PDF，跳过重新生成 |
