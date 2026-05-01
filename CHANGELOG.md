# 📋 更新日志

## v1.8.0 — PDF 引擎优化（JPEG 直通 / 无损压缩 / PDF 全布局直通）

### 🚀 重大变更

- **JPEG 直通**：`ImageSource::JpegPassthrough` + `ExternalXObject{Filter=DCTDecode}`，零质量损失
  - 工具函数 `is_jpeg_bytes` / `parse_jpeg_info`（SOF 标记解析宽高 + 颜色分量）
  - 直通条件：JPEG + 无裁白边 + 无色彩模式变更 + 旋转 0°/180°
  - 90°/270° 回退解码旋转，180° 用 PDF 变换矩阵
- **FlateDecode 无损压缩**：含 PDF/OFD 页面时 `ImageCompression::Flate`
  - FileSpec 扩展：`source_type`（image / pdf-page / ofd-page）、`pdf_path`、`pdf_page_idx`
  - 前端 `buildLayoutRequest()` 传递 sourceType / pdfPath / pdfPageIdx
- **PDF 全布局直通**：lopdf Form XObject + cm/Do 变换矩阵
  - 新增依赖 `lopdf = "0.39"`（printpdf 传递依赖已有）
  - 核心函数：`can_passthrough_pdf` / `extract_page_as_form_xobject` / `build_nup_content_stream` / `generate_pdf_passthrough`
  - 支持所有布局（1×1 到 N×M）+ 任意旋转
  - 资源深拷贝：`deep_copy_object` + `remap_references`（ObjectId 重映射）
  - 任何错误自动回退渲染管道

### 🔧 改进

- `decode_base64_to_bytes` 新函数（与现有 `decode_base64_image` 并存）
- 版本号全面统一至 v1.8.0（UI 右下角、Cargo.toml、package.json、tauri.conf.json）

### 📦 发布产物

| 文件 | 说明 | 大小 |
|------|------|------|
| `发票打印工具_x64-setup.exe` | 轻量版安装包 | ~3.5MB |
| `发票打印工具_x64_绿色版.zip` | 轻量版便携 | ~5MB |
| `发票打印工具_x64_OCR版-setup.exe` | OCR 版安装包 | ~24MB |
| `发票打印工具_x64_OCR绿色版.zip` | OCR 版便携 | ~22MB |

---

## v1.7.7 — OCR Feature Flag（轻量版/OCR版双构建）

### 🚀 重大变更

- **OCR 功能改为可选 Feature Flag**：同一套代码，编译时决定是否包含 OCR
  - 轻量版 `npm run build`：无 OCR，安装包 ~3.5MB
  - OCR 版 `npm run build:ocr`：含 PP-OCRv5，安装包 ~24MB
- **`check_ocr_available` 命令**：前端启动时检测 OCR 可用性，无 OCR 时自动隐藏相关 UI
- **模型文件不再默认打包**：轻量版 `tauri.conf.json` 移除 `models/`，OCR 版通过 `tauri.ocr.conf.json` 注入
- **`ocr-rs` 改为 optional 依赖**：不启用 `ocr` feature 时不编译 MNN 推理引擎
- **一键全量构建** (`scripts/build-all.js`)：`npm run build:all` 产出 4 个发布文件
- **Rust 条件编译**：所有 OCR 代码用 `#[cfg(feature = "ocr")]` 包裹，`invoke_handler` 按 feature 注册

### 🐛 修复

- 修复打印模式分支反转（直接打印/对话框打印函数调用互换）
- 修复关闭时 OCR 队列残留（`_tauriCleanup` 增加 `_ocrRunning=0`，`_drainOcrQueue` 检查 `__TAURI_CLOSING__`）
- 根治关闭时进程残留/死锁：`prevent_close()` 阻止 Tauri 默认关闭 → `exit(0)` 独立线程 200ms 后执行 → `TerminateProcess` 5s 兜底

### 🔧 改进

- 布局预设 3×1 → 3×2；PDF 生成前 shutdown 检查；`dist/` 加入 .gitignore

### 📦 发布产物

| 文件 | 说明 | 大小 |
|------|------|------|
| `发票打印工具_x64-setup.exe` | 轻量版安装包 | ~3.5MB |
| `发票打印工具_x64_绿色版.zip` | 轻量版便携 | ~5MB |
| `发票打印工具_x64_OCR版-setup.exe` | OCR 版安装包 | ~24MB |
| `发票打印工具_x64_OCR绿色版.zip` | OCR 版便携 | ~22MB |

---

## v1.7.6 — 一键识别 + OCR 准确率恢复

- 新增一键识别按钮 🔍（自动识别所有未识别发票，显示进度）
- 单文件 OCR 结果 toast + OCR 按钮 spinner 动画
- OCR_MAX_DIM 恢复为 960（720 对小字识别率不足），resize 滤波器恢复 Triangle

## v1.7.5 — OCR 默认关闭 + 手动识别按钮

- OCR 自动识别默认关闭，设置面板新增"自动识别"开关
- 发票列表每项新增 🔍 手动识别按钮

## v1.7.4 — OCR 速度优化

- `ocr_pdf_page` 零 IPC 往返：Rust 渲染+OCR 一体化，省掉 base64 传输链路
- OCR_MAX_DIM 960→720，resize 滤波器 Triangle→Nearest

## v1.7.3 — PDF 渲染与 OCR 分离 + 文本提取架构

- PDF 渲染与 OCR 分离：`render_and_ocr_pdf` → `render_pdf_pages`（仅渲染）+ 后台异步 OCR 队列
- 文本优先提取架构：正则直接提取 → 坐标回退
- 金额三阶段提取：含税价 → 数学验证配对(A+B=含税) → 区域解析
- 新增字段：invoiceNo、invoiceDate、buyerName、buyerCreditCode
- 发票类型检测 `_detectInvoiceType()`
- 点击跳转预览、OCR 进度 toast

## v1.7.1 — 移除 PDF.js，纯原生渲染

- 移除 PDF.js（节省 ~3.6MB），PDF 渲染完全走 WinRT，文字提取完全走 PP-OCRv5

## v1.7.0 — 含税价同行多金额修复

- 修复含税价匹配到同行不含税价：同行多金额时搜索下方更大金额

## v1.6.9 — OCR 引擎切换：WinRT → ocr-rs (PP-OCRv5 + MNN)

- OCR 引擎从 WinRT 切换为 ocr-rs (PaddleOCR + MNN)，PP-OCRv5 准确率提升约 13%
- 正则优化适配 PP-OCRv5：跨行匹配、数字空格归一化、全角￥归一化、CJK 跨行归一化
- 字符宽度权重模型、四角多边形坐标传递、OCR 置信度传递

## v1.6.8 — 含税价 findLastNum + 年份过滤

- 含税价改用 `findLastNum()`；车票价格过滤年份误匹配

## v1.6.7 — 含税价/不含税价关键字精准匹配

- 含税价上下文感知匹配，新增 `小写` 关键字；不含税价首选 standalone "合计"

## v1.6.6 — 车票位置提取 + 发票含税价修复

- 车票金额移至左半侧位置提取；含税价增加部分关键词匹配+位置反推

## v1.6.5 — OCR ¥符号误识别修复

- ¥↔1 误识别自动修正，新增 `normalizeOcrCurrency()`；修复全文折叠/展开

## v1.6.4 — 车票票种标签

- 车票票种标签显示；车票坐标邻近金额提取

## v1.6.3 — OCR ¥→1 误识别修复

- 金额关键词后"1XXX.XX"自动转为"¥XXX.XX"

## v1.6.1 — 坐标感知增强

- 不含税金额/税额关键词邻近提取；三值交叉验证；新增 `taxAmount` 字段

## v1.6.0 — 坐标感知 OCR 提取

- word 级坐标返回、区域分类、销售方 7 策略、公司后缀补全、车票检测

## v1.5.3 — 关闭后残留进程修复

- `SHUTTING_DOWN` AtomicBool + `std::process::exit(0)` 立即终止

## v1.5.2 — 启动白屏根治

- `"visible": false` 根治白屏；打印机按需加载；销售方识别增强

## v1.5.0 — 前端模块化拆分

- 拆分 ocr.js / layout.js / print.js / app.js；image 0.25 + printpdf 0.9

## v1.2.1 — 打印机选择

- 打印机选择、直接打印模式、打印机列表刷新

## v1.2.0 — 画质优化、OFD 支持、金额识别

- OFD 格式支持、OCR 金额识别、金额统计、版面布局增强、深色模式、自适应 DPI

## v1.1.0 — WinRT PDF 渲染

- `Windows.Data.Pdf` 原生渲染 + PDF.js 回退 + CMap 中文支持

## v1.0.0 — 初始版本

- PDF/JPG/PNG/BMP/WebP/TIFF 多格式、纸张规格、版面布局、拖放排序、打印/导出
