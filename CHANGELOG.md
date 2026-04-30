# 📋 更新日志

## v1.7.5 — OCR 默认关闭 + 手动识别按钮

### 🆕 新功能

- **OCR 自动识别默认关闭**：添加发票时不再自动 OCR，减少不必要的等待
  - 设置面板新增 "🔍 OCR 识别" 区域，包含"自动识别"开关
  - 开启后恢复之前的行为：添加发票时自动 OCR 识别金额、销售方等信息
  - 关闭时仅提示"已添加 N 张发票..."，不再显示"识别中"
  - 设置持久化到 `localStorage`，重启后保留
- **手动识别按钮 🔍**：发票列表每项新增 OCR 手动触发按钮
  - 点击 🔍 按钮即可对单张发票发起 OCR 识别
  - 识别中（`_ocrPending`）或加载中时按钮禁用
  - 非 Tauri 环境提示"OCR 识别需要桌面版"

### 🔧 改进

- 设置面板拆分：原 "🔍 发票查验" 区域一分为二——"🔍 OCR 识别" + "✅ 发票查验"
- 重置设置时同步清除 `fapiao-ocr-enabled` localStorage 项

---

## v1.7.4 — OCR 速度优化

### 🚀 性能优化

- **`ocr_pdf_page` 零 IPC 往返**：PDF 页面 OCR 不再走 base64 传输
  - Rust 侧新增 `ocr_pdf_page(pdfPath, pageIndex, dpi)` — 渲染+OCR 在 Rust 内存中完成
  - 省掉 `Rust→base64→IPC→前端downsample→base64→IPC→Rust解码→OCR` 整条链路
  - 前端 `applyOcrPdfPage(fileObj)` → `invoke('ocr_pdf_page', {pdfPath, pageIndex})`
  - `createFileObj` 新增 `_pdfPath`/`_pdfPageIdx` 字段，PDF 加载时填充
  - `applyOcrAsync` 三路分支: isPdfPage → applyOcrPdfPage | hasFilePath → applyOcr(filePath) | else → downsample+applyOcr
- **OCR_MAX_DIM 960→720**：发票文字大且清晰，720 足够，检测提速约 40%，识别提速约 25%
- **resize 滤波器 Triangle→Nearest**：OCR 对插值质量不敏感，Nearest 快 3-5x
  - `run_ocr_on_image` 和 `render_and_ocr_pdf` 两处 OCR resize 均改为 Nearest
  - 非 OCR 的 PDF 渲染 resize 保持 Triangle

---

## v1.7.3 — PDF 渲染与 OCR 分离 + 文本提取架构 + 交互增强

### 🚀 重大变更

- **PDF 渲染与 OCR 分离**：`render_and_ocr_pdf` → `render_pdf_pages`（仅渲染），OCR 改为后台队列异步执行
  - 预览即时显示（无需等OCR完成），占位符快速替换为真实预览
  - OCR 通过 `applyOcrAsync` 队列异步执行，不阻塞 UI
  - PDF 页面的 OCR 数据通过 `downsampleForOcr` 缩放到 960px 后 IPC 传输，减少开销
  - 图片文件仍优先使用 `_filePath` 直读磁盘，不走 base64

### 🆕 新功能

- **文本优先提取架构**：OCR 文本格式整齐，用正则直接提取结构化字段（文本提取优先，坐标仅作回退）
  - 发票号码：`发票号码：(\d{8,20})`
  - 开票日期：`开票日期：(YYYY年MM月DD日)`
  - 名称：优先 "购买方名称："/"销售方名称："，回退通用 "名称："（第1=购买方，第2=销售方）
  - 信用代码：优先 "统一社会信用代码"/"纳税人识别号" 后跟代码
  - 使用 `_normTextForExtract()` 保留换行符（普通 `normText` 会折叠 CJK 换行导致多行合并）
- **文本金额提取（三阶段）**：`_extractAmountsByText`
  - Phase 1 — 含税价：`小写¥金额` → `小写bare金额` → `价税合计后¥` → `价税合计后bare`
  - Phase 2 — 数学验证配对（PRIMARY）：含税价已知后，扫描全文所有金额，找 A+B=含税价 的配对（大=不含税，小=税额），数学关系不可伪造
  - Phase 3 — 区域解析（FALLBACK）：截取 standalone "合计" 到 "价税合计" 之间的文本
- **OCR 进度 toast**：「识别中，剩余 N 张...」实时更新，每完成一张递减
- **OCR spinner**：`_ocrPending` 标志 + `.ocr-spinner` CSS 动画，识别中显示旋转图标替代空金额
- **点击跳转预览**：点击左侧发票项自动跳转右侧预览到对应页面
  - 未勾选的发票自动勾选
  - `_activeFileIdx` 跟踪高亮项，`.active-item` CSS 蓝色左边框高亮
  - 翻页时 `syncActiveFileFromPage()` 自动更新侧边栏高亮
- **新增字段**：`invoiceNo`（发票号码）、`invoiceDate`（开票日期）、`buyerName`（购买方名称）、`buyerCreditCode`（购买方信用代码）
- **发票类型检测**：`_detectInvoiceType()` — vat/ticket/ride/unknown

### 🔧 改进

- **坐标提取降级为回退**：`extractByCoordinates` 各步骤加 `if (!field)` 保护，仅填充文本提取未找到的字段
- **旧正则路径删除**：`extractInvoiceInfo()` 已移除

---

## v1.7.2 — PDF 渲染+OCR 一体化（已废弃）

### 🚀 变更（v1.7.3 已废弃此路径）

- **PDF 一步到位**（`render_and_ocr_pdf`）：WinRT 渲染 PDF 页 → 直接 OCR → 一起返回预览图+OCR结果
- v1.7.3 改用 `render_pdf_pages`（仅渲染）+ 后台 OCR 队列，此路径已废弃

---

## v1.7.1 — 移除 PDF.js，纯原生渲染

### 🚀 重大变更

- **移除 PDF.js**（节省 ~3.6MB 安装包体积）
  - 删除 `src/pdf.min.js`、`src/pdf.worker.min.js`
  - 删除 `src/cmaps/` 目录（168 个 bcmap 文件）
  - 删除 `src/standard_fonts/` 目录（14 个字体文件）
  - PDF 渲染完全走 WinRT `Windows.Data.Pdf` 原生路径
  - 文字提取完全走 PP-OCRv5，不再需要 PDF.js 文本层

---

## v1.7.0 — 含税价同行多金额修复 + 下方优先

### 🐛 修复

- **修复含税价仍匹配到同行不含税价**（含税价修复4）
  - **根因**：发票布局中，不含税金额+税额在同一行（如 `¥172.68 ¥5.18`），含税价在更下方（如 `¥177.86`）。`_findNearbyAmount`/`findAmountNearKeyword` 按距离排序会误匹配同行不含税价
  - **修复**：检测匹配金额同行是否有其他金额→如有，搜索下方更大金额作为含税价
  - `extractByCoordinates` Step 1：价税合计关键词找到金额后，同行多金额→搜索下方
  - `extractByCoordinates` Step 1.5：小写关键词彻底重写→优先找下方金额
  - `findAmountNearKeyword` 新增 `preferBelow` 参数→同行多金额+有下方候选时选下方
  - 关键规律：含税价在`（小写）`右侧，比不含税价+税额行更下方

---

## v1.6.9 — OCR 引擎切换：WinRT → ocr-rs (PP-OCRv5 + MNN)

### 🚀 重大变更

- **OCR 引擎从 WinRT `Windows.Media.Ocr` 切换为 `ocr-rs` (PaddleOCR + MNN)**
  - 使用 PP-OCRv5 mobile 模型，识别准确率较 v4 提升约 13%
  - 支持简体中文、繁体中文、英文、日文、中文拼音 5 大文字类型
  - 增强手写体、竖排文本、生僻字识别能力
  - 字符集从 6,623 字（v4）扩展至 18,383 字（v5）
  - 模型文件：`PP-OCRv5_mobile_det.mnn` (4.5MB) + `PP-OCRv5_mobile_rec.mnn` (15.8MB) + `ppocr_keys_v5.txt` (90KB)

### 🔧 正则优化（适配 PP-OCRv5 输出特征）

- **跨行匹配**：所有金额提取正则从 `[^\d]*?` / `[^\n]*?` 改为 `[\s\S]*?`，支持 PP-OCRv5 将关键词和金额拆成多行的情况
- **数字空格归一化增强**：迭代归并数字间空格（`3 1 7.00` → `317.00`），处理 PP-OCRv5 在数字中间插入空格的问题
- **全角￥归一化**：`￥` 后空格归一化（`￥ 317.00` → `￥317.00`）
- **PP-OCRv5 特殊字符归一化**：中间点/项目符号→小数点（`3·00` → `3.00`），数字间O→0（`3O7` → `307`）
- **CJK 跨行归一化**：相邻 CJK 字符间的换行符归并（`价\n税` → `价税`），修复 PP-OCRv5 拆行问题

### 🧮 坐标优化

- **字符宽度权重模型**：`split_line_to_words` 从按字符数均分宽度改为字符宽度权重分配（CJK=2.0, Latin/digit=1.0），大幅提升 word-level 坐标准确度
- **四角多边形坐标传递**：`OcrLine` 新增 `points` 字段（4 个角点），传递 PP-OCRv5 检测模型的多边形坐标到前端
- **OCR 置信度传递**：`OcrLine` 新增 `confidence` 字段，前端可过滤低置信度结果（< 0.3 阈值）

### 📦 模型文件

- 从 RapidOCR/ModelScope 下载 ONNX 格式模型，本地用 MNNConvert 转为 MNN 格式
- 模型打包到安装程序，无需联网下载

### 🧹 清理

- 移除 `Cargo.toml` 中不再需要的 `windows` crate features（`Win32_System_Threading`、`Graphics_Imaging`）
- OCR 后端代码完全移除 WinRT `Windows.Media.Ocr` 依赖，使用纯 Rust 的 `ocr-rs`

---

## v1.6.8 — 含税价 findLastNum + 年份过滤 + 移除进度条

### 🐛 修复

- **修复含税价仍匹配到不含税价**
  - **根因**：`findFirstNum` 取关键词后第一个数字，但含税价在发票 OCR 文本最下方，第一个匹配到的常是不含税价行
  - **修复**：新增 `findLastNum()` 辅助函数，取关键词后最后一个有效数字，含税价 Step 1 正则 fallback 全部改用 `findLastNum`
- **修复车票价格匹配到年份（如 2025）**
  - **根因**：OCR 文本中 "2025年01月" 的 "2025.01" 被 `\d+\.\d{2}` 正则匹配为价格
  - **修复**：新增 `isLikelyYearOrDate()` — 过滤 1900-2099 范围值和 "20XX.XX" 日期格式；应用于 `findFirstNum`、`findLastNum`、`findAmountNearKeyword`、`findNumberNearWord`、`collectAmountWords`、Step 6 车票全部正则
  - 车票价格上限统一 ¥5000，下限 ¥5
- **移除顶部导入进度条**：遮挡标题内容，下方 toast 加载动画已足够

### ⚠️ 已知问题

- **OCR 识别准确率较低**：当前使用 Windows 系统 OCR 引擎（WinRT `Windows.Media.Ocr`），对发票版式适应性有限，金额/销售方/不含税价等字段可能出现误识别。建议用户核对 OCR 自动填充结果。**后续版本计划切换至更精准的 OCR 引擎以提升识别率。**

---

## v1.6.7 — 修复含税价/不含税价关键字精准匹配

### 🐛 修复

- **修复含税价匹配到不含税价的数字**
  - **根因**：`合\s*计` 关键字被直接用于含税价匹配，但发票中 standalone 的"合计"实际是不含税金额行的标签（"合计 100.00"），不是价税合计
  - **修复**：含税价的"合计"匹配改为上下文感知 — 只有左侧有"价"的"合计"才用于含税价（是"价税合计"的 OCR 拆分）
  - **新增**不含税价首选关键字：standalone "合计"（左侧无"价"）匹配不含税金额，放在"金额"之前
  - **新增** `小写` 关键字匹配含税价 — "（小写）"紧挨含税价数字（来自"价税合计（大写）...（小写）¥113.00"），非常特异
  - **新增** `findNumberNearWord()` 辅助函数：接受 word 对象直接查找近邻数字，支持上下文感知匹配

### 🔧 改进

- 含税价匹配优先级：`价税合计` → `价税` → `税合计` → `合计`（仅左侧有价） → `小写` → 全图兜底
- 不含税价匹配优先级：`合计`（standalone） → `金额` → `不含税金额`

---

## v1.6.6 — 修复金额识别：车票位置提取 + 发票含税价

### 🐛 修复

- **修复发票含税价仍显示不含税价**
  - **根因1**：OCR 将"价税合计"拆分为"价税"+"合计"等多个 word，`findAmountNearKeyword` 的完整关键词匹配失败，导致 `amountTax` 未找到，auto-fallback 错误地将 `amountTax = amountNoTax`
  - **修复**：增加部分关键词匹配（"价税"、"税合计"），在金额区域用坐标位置找最大值作为含税价候选
  - **新增** `collectAmountWords()` 函数：收集区域内所有金额 word 并按值排序，取最大值作为含税价
  - **新增** auto-fallback 位置反推：当只有 `amountNoTax` 时，搜索全图中比它大的最小金额作为 `amountTax`
  - `cleanOcrAmtStr` 提取为全局函数，供 `collectAmountWords` 调用
  - 含税价位置兜底：金额区域最大值 = 含税价候选，自动推导不含税价和税额

- **修复车票金额提取不准确**
  - **根因**：车票"票价"在左半侧 40-50% 高度位置，被 `classifyRegion` 归类为 `buyer` 区域，导致通用发票 Step 0 可能在错误区域提取
  - **修复**：车票提取（Step 0）移至发票提取（Step 1）之前，跳过通用发票区域分类
  - **新增**车票位置提取：在左半侧 30-60% 高度区域搜索最大金额作为票价
  - 车票关键词匹配新增"学生价"

### 🔧 改进

- **关闭进程机制简化**：移除双线程延迟退出方案，改为 `CloseRequested` 中立即 `std::process::exit(0)` + `Destroyed` 事件兜底
- **COM 对象显式释放**：OCR 和 PDF 渲染中所有 WinRT COM 对象逆序 drop/Close，确保进程退出前资源释放

---

## v1.6.5 — 修复 OCR 金额识别和全文折叠

### 🐛 修复

- **修复 OCR ¥符号误识别导致金额解析错误**
  - **根因**：OCR 常将数字"1"误读为"¥"符号（两者字形相似），产生"双¥"模式如 `¥¥72.68`（实际应为 `¥172.68`）、`￥¥07.00`（实际应为 `¥107.00`），导致不含税价被错误解析为 72.68 或 7.00
  - 新增 `normalizeOcrCurrency()` 函数，在金额提取前规范化货币符号
  - 在 `buildWordMap` 中对 word 文本也做规范化，确保坐标感知提取正确
  - 关键词正则改用 `\d{3,}` 替代 `\d{2,}`，避免3位数金额的"1"被误判为误读¥
  - 关键词正则排除 ¥ 字符 `[^\d¥￥]*?`，避免对已规范化文本重复替换产生"双¥"

- **修复 OCR 全文折叠/展开功能无效**
  - 添加通用 `.hidden { display: none; }` 规则

---

## v1.6.4 — 车票票种标签 + 坐标邻近金额提取

### 🆕 新功能

- **车票票种标签**：`getTicketTypeLabel()` 自动识别铁路电子客票/出租车票/网约车票，显示为蓝色 `.ticket-badge` 替代空白销售方

### 🔧 改进

- **车票专用坐标邻近金额提取**（Step 0.5）：`findAmountNearKeyword` + 'any' 区域，用于车票"票价"等关键词近邻匹配

---

## v1.6.3 — 修复 OCR ¥→1 误识别

### 🐛 修复

- **修复 OCR 将 ¥ 误识别为"1"**：金额关键词后"1XXX.XX"自动转为"¥XXX.XX"
- **新增 `cleanOcrAmtStr()`**：无¥前缀时剥离疑似误读的1（仅4+位数）
- 过滤裸"1"数字词（`parseInt < 2`）

---

## v1.6.1 — 坐标感知增强

### 🔧 改进

- `findAmountNearKeyword` 增加"下方搜索"：表格布局中关键词在表头、数值在下一行
- maxLineDist 从 30→80，支持跨行邻近匹配；接受1-2位小数
- Step 0 扩展：不含税金额（"金额"关键词→邻近数值）、税额（"税额"关键词→邻近数值）
- Cross-validation：amountTax = amountNoTax + taxAmount，三值互相推导
- 新增 `taxAmount` 字段：编辑弹窗、汇总栏均已支持

---

## v1.6.0 — 坐标感知 OCR 提取

### 🆕 新功能

- **坐标感知提取**：`ocr_image_from_data` 返回 JSON（含 word 级坐标），不再仅返回纯文本
- **区域分类**：`classifyRegion()` 按位置分区 — y<55% 左半=buyer/右半=seller，55-75%=amount，>75%=remark
- **Strategy 0**（最高优先级）：只在 seller 区域提取名称+信用代码，彻底消除买卖方混淆
- **Step 0**（金额最高优先级）：先在 amount 区域提取，坐标感知关键词+数值近邻匹配
- **销售方 7 策略**：坐标感知 → "销售方名称:" → "名称:"位置推导 → 替代标签 → 信用代码近邻 → 关键词后匹配 → 全文兜底
- **公司后缀补全**：有限合伙/合伙企业/个体工商户/工作室/经营部/分公司等
- **独立信用代码匹配**：OCR 遗漏前缀时用 `[0-9][A-Z0-9]{17}` 独立匹配
- **PDF.js 伪 word map**：用 `transform[4,5]` 构建坐标，Y轴翻转
- **车票检测**：`isTicketText()` 检测车票/出租车/网约车关键词，车票跳过销售方提取

---

## v1.5.3 — 彻底修复关闭后残留进程

### 🐛 修复

- **彻底修复关闭后残留进程问题**
  - **根因**：Tauri 同步 command 运行在 tokio 的 `spawn_blocking` 线程池中。OCR/PDF 渲染中的 WinRT `.get()` 调用会阻塞 OS 线程数秒。用户关闭窗口后，tokio 的 `Runtime::drop()` 会**无限等待**所有 `spawn_blocking` 任务完成，导致 WebView2 子进程沦为孤儿进程。
  - **修复**：关闭窗口时设置 `SHUTTING_DOWN` 标记并通知前端停止 OCR 队列，然后直接调用 `std::process::exit(0)`。该 API 调用 Windows `ExitProcess`，立即终止进程及所有线程。
  - `SHUTTING_DOWN` AtomicBool：OCR/PDF渲染/图片解码/PDF生成入口+循环中检查，关闭时拒绝新请求并中止进行中的循环
  - 前端：`beforeunload` 兜底 + `_drainOcrQueue` 递归时检查 `__TAURI_CLOSING__`

---

## v1.5.2 — 启动白屏根治、打印机按需加载、销售方识别增强

### 🐛 修复

- **启动白屏根治**：`tauri.conf.json` 中 `"visible": false` 彻底解决，删除无效的 `window.hide()` + splash 方案
- **打印机按需加载**：首次点击打印面板时才获取打印机列表，避免启动时 Win32 API 调用阻塞

### 🔧 改进

- **销售方识别增强（12策略+公司后缀补全）**

---

## v1.5.1 — 销售方识别增强、金额识别补全

### 🔧 改进

- **销售方识别增强（9策略）**
- **金额识别补全**：价税合计+¥直接匹配、合计金额/金额合计、税价合计(OCR颠倒)、合计附近¥双向、全角￥、定额发票短文本、开票金额/发票金额
- **导入动画优化**

---

## v1.5.0 — image 0.25 + printpdf 0.9 + 前端模块化拆分

### 🆕 新功能

- **前端模块化拆分**（v1.5.0）：
  - `src/ocr.js` — OCR 提取 + `applyOcr()` 统一入口
  - `src/layout.js` — `calculateLayout()` 纯函数 + HTML预览渲染
  - `src/print.js` — 打印/导出 PDF
  - `src/app.js` — 主入口、状态、文件加载

### 🔧 改进

- **image crate 升级到 0.25**：原生支持 WebP，`DynamicImage::from(buf)` 新 API
- **printpdf 升级到 0.9**：`PdfPage::new()` + `Op::UseXobject` 新 API，`doc.save()` 直接返回字节
- **前端 Canvas 转换移除**：`ensureRustCompatibleUrl()` 和 `processTrim()` 中的 Canvas 转换代码已删除，data URL 直接透传 Rust
- **Rust 图片/XObject 缓存**：避免重复解码

---

## v1.4.0 — 修复6项问题

### 🐛 修复

- 销售方识别/车票跳过/布局优化/导入性能/金额识别/进程残留

---

## v1.2.1 — 打印优化、打印机选择

### 🆕 新功能

- **打印机选择**：可在左侧「🖨 打印」面板选择特定打印机
- **直接打印模式优化**：使用 winprint 直接发送 PDF 到 Windows Print Spooler，无需 PDF 阅读器

### 🔧 改进

- **打印机列表刷新**：使用 winprint 原生 API 获取打印机列表
- **默认打印机获取**：使用 Win32 GetDefaultPrinterW API
- **打印模式持久化**：打印模式设置保存到 localStorage

---

## v1.2.0 — 画质优化、OFD支持、金额识别、排版增强

### 🆕 新功能

- **OFD 格式支持**：支持打开 OFD（电子发票标准格式），自动提取嵌入图片
- **OCR 金额自动识别**：基于 WinRT OCR 引擎，零额外依赖
- **金额统计**：文件列表底部实时显示已选发票金额汇总
- **发票查验**：一键跳转国家税务总局发票查验平台
- **版面布局增强**：9 个预设 + 自定义行列，工具栏快捷切换
- **缩放功能增强**：Ctrl+滚轮、双击重置、Ctrl++/-/0 快捷键
- **自适应 DPI 渲染**：小 PDF 页面自动提升渲染 DPI
- **深色模式**：完整的深色主题支持
- **设置面板扩展**

### 🔧 改进

- **DPI 常量统一**：前端和 Rust 端统一为 300
- **Per-slot 边距**：每张发票在各自 slot 内有独立边距

### 🐛 修复

- CID 字体 PDF 拼接缺陷、OCR 多页中断、金额正则过严等

---

## v1.1.0 — WinRT PDF渲染修复，中文发票完美显示

### 🆕 新功能

- **WinRT 原生 PDF 渲染**：使用 `Windows.Data.Pdf` API
- **PDF.js 回退**：WinRT 失败时自动回退
- **PDF.js CMap 配置**：中文 CID 编码字体正常渲染

---

## v1.0.0 — 发票批量打印工具

### 🆕 初始功能

- PDF / JPG / PNG / BMP / WebP / TIFF 多格式支持
- A4 / A5 / B5 / Letter / Legal 纸张规格
- 1×1 / 1×2 / 2×2 版面布局
- 拖放添加、排序、旋转、缩放
- PDF 导出、打印
