# 📋 更新日志

## v1.9.0 — 单票独立调整 + PDF统一直通 + UI全面优化

### 🚀 新功能

- **单票独立调整大小/位置**：每张发票可在预览中拖拽移动 + 角落 handle 缩放，侧边栏「单票调整」面板 + 发票弹窗参数编辑，PDF 裁剪输出
  - 数据模型 `fileObj.{slotScale, slotOffsetX, slotOffsetY}` — slot 内独立缩放/偏移
  - CSS transform (translate+scale+rotate) 预览 → Rust `SlotSpec.scale/offset_x/offset_y` PDF 生成
- **发票列表排序**：↑↓ 按钮替换拖拽排序（Tauri webview 拖拽卡顿），hover 浮动显示不占空间

### 🔧 重构

- **PDF 生成统一 lopdf 混合直通路径**：图片用 JPEG XObject，删除 `can_passthrough_pdf` 分支，代码路径更简洁
- **文件列表 UI 重构**：操作按钮 + 大小标签合入 meta 行，文件名/销售方独占整行宽度，侧边栏 340px

### 🐛 修复

- **弹窗输入框全面对齐**：双行布局 — 固定行右对齐 140px + 自适应行 flex 填满 + 标签左对齐 + 单票 % 与 mm 对齐
- **meta 行金额过长**：标签区可收缩 + 操作区不换行，金额输入框自适应宽度
- **列表选中抖动**：hover padding-left 变化 + border-left 导致内容右移 → 改用 box-shadow
- **app.js 语法错误**：移除 fallbackCopy 编辑残余的重复代码
- **构建脚本**：OCR 绿色版 zip 内 exe 使用原始文件名 + 修复全量构建误选旧版安装包

---

## v1.8.4 — 页面内容层 DrawParam 继承修复

### 🐛 修复

- **页面内容层 DrawParam 错误继承**：移除页面内容 Layer 无 DrawParam 时自动查找根 DrawParam 的逻辑。页面数据层无 DrawParam 属性时直接使用 OFD 默认黑色（0,0,0），不再错误继承模板层的 DrawParam 颜色（如深红 128,0,0），修复发票文字颜色异常
- **README.md 重写优化**

---

## v1.8.3 — OFD 矢量渲染完善（文字间距 / 红章 / 字重 / 字体回退）

### 🐛 修复

- **OFD 文字 x 定位**：无 DeltaX 或单字符文本（¥符号、数量"台"等）必须在 `<text>` 上设置 `x` 属性，否则渲染在 x=0
- **OFD 注解 Appearance 偏移**：`Annotation.xml` 中 `ImageObject Boundary` 是相对 `Appearance Boundary` 的局部坐标，必须叠加 Appearance 的 x/y 偏移，否则红章渲染在 (0,0)
- **OFD DrawParam 继承链**：页面内容 Layer 无显式 DrawParam 时，自动查找根 DrawParam（被引用但自身无 Relative 的节点）作为全局回退，线条不可见问题修复
- **OFD DrawParam 应用于文本**：`apply_draw_param_defaults` 扩展到 `TextObject`，文本 fill/stroke 颜色也继承 DrawParam
- **OFD 填充颜色回退**：`fill=true` 但无 `fillColor` 时不再回退到 `strokeColor`（避免实心填充遮盖 ¥ 等内部笔画）
- **旋转后图片不自适应 slot**：三处修复

### 🔧 优化

- **OFD 字重解析**：`TextObject Weight` 属性（400=normal, 700=bold）决定加粗，不再硬编码字体名
- **OFD 字体 fallback**：SVG `font-family` 跨平台 fallback 链（楷体→KaiTi→STKaiti→serif 等），Windows/macOS/Linux 均可正确渲染
- **OFD 图片提取优化**：印章过滤 + 尺寸筛选 + 页面去重
- **加载/OCR 进度 Toast 逻辑改进**
- **cargo check warnings 清理**

---

## v1.8.2 — 进程退出安全修复 + CropBox 优先 + 加载进度优化

### 🐛 修复

- **进程退出死锁修复**：用 `TerminateProcess` 替代 `process::exit(0)`（ExitProcess）
  - ExitProcess 先杀所有线程再走 DLL_PROCESS_DETACH，若 MNN/OCR 引擎死锁则进程永远残留
  - TerminateProcess 跳过 DLL_PROCESS_DETACH，立即终止，永不挂起
  - `CloseRequested` 和 `Destroyed` 事件统一使用 TerminateProcess
- **PDF CropBox 优先**：passthrough 模式下 CropBox 优先于 MediaBox（PDF 规范 7.7.3.3）
  - 修复含 CropBox 的 PDF（如仅显示页面下半部分）缩放计算错误
  - CropBox 非零原点时自动添加 `1 0 0 1 -x1 -y1 cm` 平移

### 🔧 优化

- **加载进度显示**：toast 显示 "加载中 X/Y，识别中 X/Y" 格式，进度一目了然
- **并行加载 + 顺序渲染**：所有文件并行加载（快），结果逐个渲染并 yield 给浏览器（用户看到逐个出现）
- **OCR 批次追踪改进**：`_ocrFromButton` 区分单文件/批量 OCR，单文件显示识别金额，批量显示总数

---

## v1.8.1 — OCR 准确率优化 + PDF 空白页修复 + 静默打印

### 🚀 新功能

- **Print Spooler API 静默打印**：绕过 PDF 阅读器，直接将 PDF 字节写入打印队列
  - `OpenPrinterW` → `StartDocPrinterW(RAW)` → `WritePrinter` → `EndDocPrinterW` → `ClosePrinter`
  - `PrinterHandle` RAII guard 确保 `ClosePrinter` 正确调用
  - 失败自动回退 `ShellExecuteW("printto")` + `SW_HIDE`
  - `confirmPrint` 确认弹窗：打印前显示发票数量/版面/纸张/打印机/模式/份数
- **发票查验平台修正**：主平台改回国家税务总局官方 `inv-veri.chinatax.gov.cn`，仿真平台 `fz.chinaive.com` 降为备用

### 🐛 修复

- **PDF 空白页根因修复**：资源继承缺失 + 双重压缩 + MediaBox 继承
  - `get_page_resources` 陷阱：返回 `(Option<&Dictionary>, Vec<ObjectId>)`，第一个只含页面内联 Resources（Reference 被跳过），必须合并 ref_ids + resources_opt 才能得到完整资源
  - Stream 双重压缩：已有 Filter 的流 `allows_compression = false`，避免对已压缩流再次 FlateDecode
  - MediaBox 继承：部分 PDF 的 MediaBox 在父 Pages 节点上，需向上遍历 Parent
  - CropBox 优先于 MediaBox：`get_page_effective_box()` 优先 CropBox → 回退 MediaBox；非零原点时内容流前加 `1 0 0 1 -x1 -y1 cm` 平移
- **PDF 合成遗漏修复**：修复合成 PDF 时部分页面遗漏的问题

### 🔧 OCR 准确率优化

- **OCR_MAX_DIM**: 960 → 1280px，保留更多小字细节
- **Resize 滤波器**: Triangle → Lanczos3，文字边缘更锐利
- **对比度增强**: `enhance_contrast_ocr()` 直方图拉伸（1%-99%），低对比度发票效果显著
- **恢复 v1.6.7 关键修正**: `_normTextForExtract` 中关键词后 ¥→1 误读修正
- **卖家名提取增强**: 恢复信用代码锚定、销方名称缩写、收款单位、公司后缀等策略
- **前端降采样**: 960 → 1280，JPEG 质量 0.85 → 0.92

### 📦 发布产物

| 文件 | 说明 | 大小 |
|------|------|------|
| `发票打印工具_x64-setup.exe` | 轻量版安装包 | ~3.5MB |
| `发票打印工具_x64_绿色版.exe` | 轻量版便携（单文件） | ~5MB |
| `发票打印工具_x64_OCR版-setup.exe` | OCR 版安装包 | ~24MB |
| `发票打印工具_x64_OCR绿色版.zip` | OCR 版便携 | ~22MB |

---

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
| `发票打印工具_x64_绿色版.exe` | 轻量版便携（单文件） | ~5MB |
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
