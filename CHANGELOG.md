# 📋 更新日志

## v2.0.7 — XML 数电票 + 文件列表记忆 + 打印状态

### 📄 XML 数电票支持

- **独立 XML 解析**：`parse_xml_invoice` 命令解析 `<EInvoice>` 格式，提取发票号码/日期/金额/销售方/购买方/类型
- **文件列表占位**：XML 文件显示 XML 图标 + 发票尾号，无缩略图
- **汇总表兼容**：`_xmlInvoice` 标记自动适配汇总表、CSV 导出、批量重命名
- **不参与排版**：`getActiveFiles()` 过滤，`getFileIndex()` 返回 null，避免预览/打印异常
- **格式限制**：纯数据格式，无版式信息，不可渲染票面

### 💾 文件列表记忆（可选）

- **记忆开关**：设置面板新增「记忆发票列表」，默认关闭
- **恢复机制**：启动时自动重建文件列表、缩略图、字段提取（跳过 OCR）
- **路径校验**：启动时验证文件存在性，自动跳过已删除文件
- **仅记忆文件路径**：不保存金额/OCR 等易变数据，保持轻量

### ✅ 打印状态追踪

- **三种过滤**：侧边栏顶部「全部/未打印/已打印」过滤按钮组
- **自动标记**：四种打印模式成功后自动标记已打印（绿色 ✓ 标识）
- **持久化**：`_printedMap` 始终保存到 localStorage，不受开关影响
- **重置迁移**：`clearAll()` / `executeRename()` / `resetSettings()` 均正确迁移

### 🔧 ofd-engine → invoice-engine

- **Crate 重命名**：`ofd-engine` → `invoice-engine`，整合 OFD + XML 解析
- `parse_ofd` / `parse_xml_invoice` / `open_ofd_images` 均从 invoice-engine 导入

### 🐛 修复

- `togglePref()` 漏调 `saveSettings()`：切换开关后设置不持久化的问题（影响所有开关）

---

## v2.0.6 — 品牌升级：发票酱 + 重命名备注

### 🎨 品牌升级

- **产品更名**：从「电子发票批量打印工具」正式更名为「发票酱 (TicketChan)」
- npm/Cargo 包名：`fapiao-print` → `ticketchan`
- 标识符：`com.fapiao.print` → `com.ticketchan.app`
- 所有配置键、构建产物名、CI 脚本全面更新
- 删除 4 个硬编码本地路径的测试函数

### 🔄 重命名增强

- **备注作为命名字段**：汇总表「备注」列默认显示，可在重命名模板中勾选「备注」作为文件名元素
- **行内备注编辑**：勾选「备注」后预览表新增备注输入列，直接填写或修改每张发票的备注，300ms 防抖自动刷新文件名预览 + 汇总表双向同步
- **模板持久化**：重命名模板勾选字段 + 分隔符自动保存到本地存储，下次打开软件自动恢复
- **旧用户迁移**：已有保存配置的用户自动补上备注列
- **清除按钮**：重命名字段区新增「清除」按钮（统一样式），一键清空所有勾选

### 🛠 设计原则

- 备注复用现有 `_notesMap` / `summaryNotes` 持久化机制，重命名后 key 自动迁移
- 汇总表编辑备注 → 重命名预览自动刷新；重命名预览编辑备注 → 汇总表自动刷新
- `executeRename` 前检测备注变更，防止防抖期间的竞态导致文件名不含新备注

---


### 🔄 批量重命名（新增）

- **汇总表内嵌面板**：点击「🔄 重命名文件」展开配置面板，与汇总表在同一弹窗中操作
- **预设模板**：金额+销售方+号码 / 销售方+号码 / 金额+日期+号码，默认模板含发票号码防止重名
- **自定义字段**：勾选 10 个字段（含税金额/不含税金额/税额/销售方名称/税号/购买方名称/税号/发票号码/开票日期/发票类型），后勾选的字段排在文件名后面
- **分隔符可调**：默认 `_`，支持自定义（最多 4 字符）
- **实时预览**：原文件名 → 新文件名，状态标识（✓ 正常 / ⚠ 重名加序号 / ✗ 字段为空跳过）
- **重名自动序号**：`94.00_某某公司.pdf` → `94.00_某某公司_2.pdf`
- **多页 PDF 安全**：同一 PDF 多页只重命名一次，后续页面自动跳过并同步路径
- **持久化迁移**：`_fileAdjMap`（单票调整）+ `_notesMap`（备注）的 key 自动同步到新文件名

### 🛠 Rust 后端

- `rename_file` 命令：`async fn` + `spawn_blocking`，同盘 `fs::rename`（原子），跨盘 `copy + delete` fallback

### 🐛 修复

- OFD 文件加载时设置 `filePath` 以支持重命名（`print.js` 的 `sourceType` 判断兼容此改动）
- 汇总表列宽优化：金额列固定 80px，文本列 `min-width: 100px` 自动撑开

---

## v2.0.4 — PDF 印章烘焙

### 🏷️ 印章烘焙（新增）

- **PDF 标注保留**：生成合成 PDF 时，原票 PDF 的印章/签章/标注（`/Annots`）自动烘焙到 Form XObject 内容流，输出 PDF 印章不丢失
- **隐藏标注过滤**：自动跳过 PDF 标记为隐藏（F bit 2）的标注，避免误烘焙
- **外观流完整迁移**：标注外观（`/AP` → `/N`）及其嵌套 XObject 资源经 deep copy 完整迁移到输出文档，确保印章保真
- **BBox 坐标映射**：标注 Rect 坐标系自动映射到 Form BBox 空间，矩阵组合（平移+缩放）正确，印章位置/大小与原票一致

### 🧪 测试覆盖

- 4 个自动化测试（含 E2E 直通流程），验证 7.22底座.pdf 双印章和 320101.pdf 单印章烘焙

---

## v2.0.3 — 发票汇总表导出

### 📊 汇总表导出（新增）

- **可编辑预览**：弹窗展示所有发票字段表格，支持 14 个字段勾选（序号/发票号码/日期/类型/销售方/税号/购买方/税号/含税金额/不含税/税额/文件名/份数/备注）
- **单元格编辑**：金额、日期、名称可直接双击编辑，改动自动回写 fileObj 并同步到文件列表、预览和底部汇总栏
- **三种金额合计**：含税金额、不含税金额、税额分别汇总，合计行固定在表格底部
- **CSV 导出**：UTF-8 BOM 编码，Excel 直接打开无乱码，导出后自动打开所在文件夹
- **记忆配置**：列勾选和备注统一持久化到 `ticketchan-settings`，关闭软件后自动恢复

### 🛠 代码优化

- `write_text_file` Rust 命令改为 `async fn` + `spawn_blocking`（符合 IPC 异步化规范）
- 汇总栏布局重构：销售方数量 `div→span`，flex 精确控制单行不换行
- `getCheckedFiles()` 辅助函数：不含 copies 展开，供汇总表等数据型操作使用

### 📋 开发质量

- 汇总表功能涉及 4 文件（HTML/CSS/JS/Rust），零新依赖
- 14 个开发迭代提交合并为 1 个功能提交

---

## v2.0.2 — 单票调整交互优化

### 🎯 单票调整增强

- **放大上限 3x**：slotScale 上限从 2.0 放宽到 3.0，解决地铁行程单等窄长发票放大不够的问题
- **拖拽约束动态化**：根据发票实际显示尺寸（兼容 contain/fill/original/custom 四种适配模式）动态计算可拖范围，放越大可拖越远，不再受固定约束限制
- **滚轮缩放单票**：单击选中槽位后，鼠标滚轮直接调节该票缩放比例（5%/步），无需 Ctrl，无需去侧边面板
- **编辑态溢出可见**：选中或拖拽中的发票临时显示超出 slot 的内容，方便判断调整方向和幅度；非编辑态保持 overflow:hidden 与打印输出一致
- **双击重置单票**：双击已选中的单票 → 重置大小+位置，双击空白区域保留预览尺寸复原

### 🛠 设计原则

- **不溢出原则**：发票内容始终 ≤ slot 边界，Rust 端 clip 不变，编辑态临时溢出仅辅助调整
- **超 A5 发票适配**：通过放大 + 拖动将重要内容调整到 slot 可见区，边缘裁切由 Rust clip 自动处理
- 偏移量面板输入范围同步扩展到 ±150mm

### 📋 开发质量

- 变动仅涉及 JS/HTML/CSS（4 文件），不触发 Rust 重编译
- 拖拽约束数学推导覆盖所有 fitMode，两轮复核通过
- 方案文档：`docs/plan-slot-enhancement.md`

---

## v2.0.1 — 单票调整 UI 完善 + 稳定性修复

### 🎯 单票调整（重大 UI 改进）

- **快速对齐九宫格**：一键贴边/居中，9 种对齐方向（↖↑↗←⊙→↙↓↘）
- **鼠标滚轮增减**：所有数字输入框和滑块支持滚轮微调
- **输入框加宽**：52→72px，适配 ±三位数带负号
- **拖动修复**：CSS transform 应用到 wrapper div（与渲染一致），消除拖动错位
- **偏移范围扩展**：±50→±150，覆盖 A3/A4 所有布局
- **调整记忆**：可选的单票调整配置持久化（按文件名匹配，跨会话恢复）

### 🛠 稳定性修复

- **savePdf 路径失效**：savedDir 被删除/移动后自动回退对话框；缓存 PDF 源文件校验
- **保存 PDF 默认路径**：优先使用用户下载目录

### 📋 开发质量

- 对齐计算 3 轮迭代修复（旋转适配、scale 适配、范围适配）
- 死代码修复（createFileObj return 提前退出）
- 前后端同步复核通过

---

## v2.0.0 — 稳定大版本

历经 v1.10.x 五个子版本迭代，ticketchan 在打印稳定性、性能优化、用户体验方面趋于成熟，升级为 2.0.0 大版本。

### v1.10.x 迭代总结

**稳定性**：PDFium 打印 SEH 保护（v1.10.3）、DEVMODE 完整缓冲区、打印机兼容性修复（v1.10.2）、打印流程解耦（v1.10.4）

**性能**：预览速度 2-3x 提升（150DPI + JPEG）、批量文件加载 + 批量 PDF 文字提取并行化、智能 PDF 缓存（deepEqual）

**可用性**：设置持久化（v1.10.1）、金额校验可视化（v1.10.4）、排版份数批量设置（v1.10.4）、PDFium 打印自动降级（v1.10.5）

**代码质量**：IPC 异步化（spawn_blocking）、死代码清理、refactor 消除重复调用

### 从 v1.10.5 以来的改动

- PDF 缓存键排除打印机参数，切换打印机/份数不再重新生成 PDF
- PDFium 打印 copies/duplex 独立参数化
- 边框跟随发票视觉边界（preview + PDF）
- 修复多 PDF 文件分组文字提取
- 回归测试全面通过

---

## v1.10.5 — 预览加速 + 批量加载 + 智能缓存 + IPC 异步化

### 🚀 新功能

- **批量 PDF 文字提取**：新增 `extract_pdf_texts` Rust 命令 + `applyPdfTextToResults` 前端函数，一次打开 PDF，rayon 并行提取多页文字；前端按 PDF 路径分组批量调用，失败时自动回退单页模式
- **copy_file Rust 命令**：新增文件复制命令（`std::fs::copy`），用于缓存 PDF 复用到保存路径

### ⚡ 优化

- **预览加载速度大幅提升**（2-3x）：PDF 文件预览加载性能优化，包括降低预览 DPI（150 vs 300）和 JPEG 编码支持，渲染像素减少 75%，文件体积减少 60-80%
  - 预览 DPI：300 → 150（屏幕显示足够清晰）
  - 图像格式：PNG → JPEG（质量 80%，文件体积更小）
  - 移除预览时的强制 DPI 自适应缩放
  - 打印质量完全不受影响：打印/保存走独立矢量流程（lopdf 直通）
- **批量文件加载优化**：重构 `processFilesIncremental`，一次批量 IPC 调用 `open_invoice_files` + `Promise.all` 并行渲染 + `setInterval` 定时批量刷新 UI，Toast 防抖 100ms，大幅减少等待时间
- **智能 PDF 缓存**：引入 `deepEqual` 深度对象比较 + `canUseCachedPdf` + `updatePdfCache`，只要有有效缓存就复用，替代旧的 dirty flag 方案
- **保存 PDF 缓存复用**：`savePdf` 先在临时目录生成 PDF 作为缓存，再 `copy_file` 到用户路径；布局不变时直接复制缓存文件
- **PDFium 打印自动降级**：PDFium 打印失败/异常时自动 fallback 到 `doSumatraPrint`，用户无感知兜底
- **IPC 异步化**：`render_pdf_pages` / `render_pdf_pages_pdfium` / `extract_pdf_text` / `extract_pdf_texts` 改为 `async fn` + `spawn_blocking`，防止 IPC 消息泵饥饿导致 `ERR_CONNECTION_REFUSED`
- **文字提取日志增强**：`extract_pdf_text_from_doc` 新增 `pdf_path` 参数，错误信息包含文件名

### ⚠️ 已知限制

- PDF 阅读器模式仍不支持选择打印机（由阅读器实现决定）

---

## v1.10.4 — 打印流程解耦 + 金额校验可视化 + 排版份数

### 🚀 新功能

- **金额校验失败可视化**：OCR 和 PDF 文字提取金额求和校验失败时，发票卡片金额徽章显示 ⚠ 警告标识；hover 警告徽章可查看含税/不含税/税额/验证计算详情；汇总栏新增校验异常发票计数提示
- **排版份数批量设置**：文件列表新增 ② 按钮，支持批量设置选中发票排版份数（×1/×2/×3）；区分「排版份数」（每张发票在版面中重复几次）与「打印份数」（整版打印几份），模态框和设置面板分别标注

### 🐛 修复

- **打印流程解耦**：各打印模式独立调用打印命令，不再经 `generate_pdf_from_layout` 隐式降级。此前 SumatraPDF 模式重新生成时会 fallback 到 `shell_execute_print`，PDF 阅读器模式会经 SumatraPDF 路径，现已修正为各模式显式调用对应打印命令
- **PDF 阅读器模式提示消息修正**：`is_direct=false` 时消息从「已弹出打印对话框」改为「已打开PDF，请在阅读器中打印」

### 🔧 优化

- **PDF 阅读器模式提示完善**：UI 提示补充「不支持选择打印机」说明，避免用户误以为该模式可控制打印机；AGENTS.md 同步补充该已知限制的技术原因（`ShellExecuteW printto` 动词依赖阅读器实现，多数阅读器不支持指定打印机，fallback 到默认打印机）

---

## v1.10.3 — PDFium 打印闪退修复 + DEVMODE 完整缓冲区

### 🐛 修复

- **PDFium 打印闪退（部分电脑）**：`FPDF_RenderPage` 直打某些打印机 DC 时，驱动 GDI 实现有 bug 导致原生访问违例（ACCESS_VIOLATION），Rust 无法捕获直接闪退。现新增 SEH（结构化异常处理）包装器 `seh_wrapper.c`，用 C 的 `__try/__except` 捕获原生崩溃，自动 fallback 到位图渲染（`FPDF_RenderPageBitmap` + `StretchDIBits`），不再闪退
- **DEVMODE 驱动私有数据截断**：`get_printer_default_devmode()` 原先用 `std::ptr::read` 只复制 `sizeof(DEVMODEW)` 字节，丢弃了 `dmDriverExtra` 字节的驱动私有数据，导致 `CreateDCW` 访问违例。现改为返回完整 `Vec<u8>` 缓冲区，保留全部驱动配置
- **`shell_fallback_open` 非 Windows 编译失败**：变量在 `#[cfg(target_os = "windows")]` 块内声明但在块外使用，非 Windows 平台编译报错。已添加 `#[cfg(not(target_os = "windows"))]` 分支修复

### 🔧 优化

- **PDF阅读器模式提示完善**：UI 提示补充「不支持选择打印机」说明，避免用户误以为该模式可控制打印机；AGENTS.md 同步补充该已知限制的技术原因（`ShellExecuteW printto` 动词依赖阅读器实现，多数阅读器不支持指定打印机，fallback 到默认打印机）

- **矢量优先 + 位图 fallback**：PDFium 打印始终先尝试矢量直打 DC（零质量损失），仅在 SEH 捕获异常时自动 fallback 到位图渲染，兼顾打印质量和稳定性
- **PDF阅读器模式 Strategy 3 fallback**：`shell_execute_print()` 新增 `ShellExecuteW("open")` fallback，当 `printto`/`print` 均失败时自动打开 PDF 供用户手动打印

### 📦 新增依赖

- `cc = "1"`（build-dependencies）— 编译 `seh_wrapper.c` 为静态库

---

## v1.10.2 — 打印兼容性修复

### 🐛 修复

- **PDF阅读器模式打印报错（错误码31）**：PDF-XChange Editor 等第三方阅读器未注册 `printto`/`print` Shell 动词时，`ShellExecuteW` 返回 SE_ERR_NOASSOC(31)。现增加 Strategy 3 fallback：自动回退 `ShellExecuteW("open")` 打开 PDF，提示用户手动打印
- **PDFium/弹窗确认模式打印机有声音但不进纸**：`build_dev_mode()` 从零构建 DEVMODE 缺少打印机驱动默认配置（纸盒选择、纸张来源等），导致 HP 等网络打印机无法正确匹配纸盒。现通过 `DocumentPropertiesW` API 获取打印机默认 DEVMODE 作为基础配置，再覆盖自定义字段
- **打印机 DC 无效尺寸防御**：`GetDeviceCaps` 返回 HORZRES/VERTRES/LOGPIXELSX ≤ 0 时提前报错，避免空页打印

---

## v1.10.1 — 设置持久化

### 🚀 新功能

- **设置持久化**：关闭软件后自动记住用户设置，下次打开自动恢复，无需重复调整
  - 排版布局（行列数，如常用 2×1）
  - 纸张规格、方向、自定义纸张尺寸
  - 四边边距、列间距/行间距
  - 缩放模式与自定义缩放比例
  - 全局旋转、份数、颜色模式、页面顺序、打印模式
  - 辅助开关（裁切线、编号、边框、裁剪白边、水印、逐份打印、双面打印、打印页码、打印日期、自定义页脚、自动打开PDF、自定义下边距）
  - 水印参数（文字、透明度、颜色、角度、字号）
  - 页脚文本、下边距数值
  - 统一 `saveSettings()`/`loadSettings()` 管理，`ticketchan-settings` JSON 存储
  - `updatePreview()` 防抖 500ms 自动保存，无文件时也能保存
  - 恢复默认设置时清除所有持久化数据

---

## v1.10.0 — PDF 渲染双引擎 + WinRT fallback

### 🚀 新功能

- **PDF 渲染双引擎**：新增 PDFium 渲染 fallback，解决部分电脑（企业精简版/LTSC）WinRT PDF 组件不可用导致 PDF 文件加载失败的问题
  - 启动检测 `check_winrt_pdf_available()`：创建临时 PDF 测试 WinRT `PdfDocument` API 可用性
  - PDFium 位图渲染 `render_pdf_pages_pdfium()`：`FPDF_LoadMemDocument` + `FPDF_RenderPageBitmap` → BGRA→RGBA → PNG
  - 前端 fallback 链：`_winrtPdfAvailable` 标志控制，WinRT 失败自动切换 PDFium，后续直接走 PDFium
  - `pdfium_print.rs` 扩展位图 API：`FPDFBitmap_Create` / `FPDFBitmap_FillRect` / `FPDF_RenderPageBitmap` / `FPDFBitmap_GetBuffer` / `FPDFBitmap_GetStride` / `FPDFBitmap_Destroy`
- **PDFium DLL 下载提示优化**：区分"PDF 预览"和"打印"场景的提示文案，启动时 WinRT 不可用 + DLL 不存在自动弹出下载提示
  - `showPdfiumMissing(reason)` 支持自定义原因文案，黄色警告框突出显示
  - 下载成功后提示"请重新添加 PDF 文件"
  - 打印场景专用提示："PDFium 打印引擎需要 pdfium.dll 才能工作。"

### 🐛 修复

- **PDF 文件加载失败**（部分电脑）：WinRT `PdfDocument` API 在企业精简版/LTSC 系统上不可用时，PDF 文件无法渲染预览，图片文件正常。现自动 fallback 到 PDFium 渲染引擎
- **Alpha 通道处理不安全**：`FPDFBitmap_Create(alpha=0)` 第 4 字节未定义，强制 alpha=255 避免垃圾值

### 🔧 优化

- **错误路径资源泄漏防护**：PDFium 渲染失败时 `close_document` 改为 `let _ =` 避免吞掉原始错误
- **错误信息精确匹配**：DLL 缺失提示仅匹配 `pdfium.dll` 和 `不可用` 关键词，不再误匹配 PDF 文件损坏错误
- **cfg 条件简化**：移除冗余 `#[cfg(target_os = "windows")]`（项目 Windows-only）

---

## v1.9.8 — PDFium 矢量打印 + XPS 清理 + 打印体验优化

### 🚀 新功能

- **PDFium 矢量静默打印**：新增 Chromium PDFium 引擎直打打印机 DC，打印清晰，Spool 体积极小
  - `pdfium_print.rs` 新模块：`FPDF_LoadMemDocument` 内存加载 → `FPDF_RenderPage(printer_dc)` 直打，无需 EMF 中间层
  - DLL 全局生命周期管理：`LazyLock<Mutex<Option<PdfiumState>>>` + `with_pdfium()` 闭包串行化
  - 按需下载：`download_pdfium_dll` 命令，`gh-proxy.com` 加速国内下载，`cancel_download` 命令支持取消
  - 缺失弹窗：自动下载 / 手动下载（GitHub Releases）/ 切换到 PDF 阅读器模式
- **弹窗确认模式增加引擎选择**：确认弹窗内可选 PDFium（推荐）或 SumatraPDF
- **下载取消机制**：`AtomicBool DOWNLOAD_CANCELLED` 全局标志，前端取消时 Rust 端立即停止下载并清理临时文件（PDFium + SumatraPDF 均支持）

### 🐛 修复

- **PDFium 打印模糊**：首版 EMF 中间层方案坐标映射错误（0.01mm 画框 vs 像素内容），改为直接渲染到打印机 DC，以打印机原生 DPI 输出
- **`cancelPdfiumDownload` / `cancelSumatraDownload` 仅设 JS 标志**：Rust 端继续下载浪费带宽，现通过 `cancel_download` 命令通知 Rust 端停止
- **`doSumatraPrint` 参数默认值缺失**：`copies`/`duplex`/`paperW`/`paperH`/`colorMode` 添加 `||` 兜底
- **`colorMode` 默认值为空字符串**：改为 `'color'`，三个打印渠道统一为 `s.colorMode || 'color'`

### 🔧 优化

- **打印模式重命名**：选项名即说明，`静默打印（PDFium）⭐` / `静默打印（SumatraPDF）` / `弹窗确认` / `PDF阅读器`
- **弹窗确认精简**：15 行合并为 6 行，缩小行间距，不再需要滚动
- **PDF 缓存复用统一**：`_pdfDirty` + `_lastPdfPath` 跨三个打印渠道共享，PDFium 命中缓存时跳过生成
- **`pdfium_print_pdf` 参数类型安全**：`Option<u32>` → `u32`，`Option<f32>` → `f32`
- **临时文件固定命名**：`pdfium_cache.pdf` 替代时间戳命名，不再累积
- **`pdfium_vector_print` 简化**：生成 PDF 后直接调用 `pdfium_print_pdf`，消除冗余磁盘读取
- **DLL 统一存放**：`pdfium.dll` 移至 `{exe}/tools/` 目录，与 `SumatraPDF.exe` 一致
- **下载源更新**：`niclas/pdfium-binaries` → `bblanchon/pdfium-binaries`（更主流），`gh-proxy.com` 加速
- **默认打印模式**：改为 PDF 阅读器（无需下载依赖即可使用）
- **UI 措辞**：去掉 Spool/矢量/光栅化等技术术语，改为"推荐"/"需下载"等用户友好描述
- **手动下载提示**：弹窗提示文件放置路径（tools 文件夹）和重命名要求

### 🗑️ 清理

- **XPS 打印代码全部移除**：`pdf_engine.rs` 删除 ~770 行（`xps_print`、`build_xps_bytes`、`build_xps_doc_xml` 等 12 个函数）
- **`render_footer_text_png` 死代码移除**：仅 XPS 使用的页脚 PNG 渲染函数（~56 行）
- **`doXpsPrint` 前端函数移除**
- **`xps_print` Tauri 命令移除**（Windows + 非 Windows 版本 + invoke_handler 注册）
- **净减少 ~370 行代码**

### 📦 新增依赖

- `libloading = "0.8"` — 动态加载 pdfium.dll
- `tar = "0.4"` — 解压 pdfium tgz 包

---

## v1.9.7 — dzcp/iloveofd 格式兼容修复 + 非税发票金额修复

### 🚀 新功能

- **PDF 文字提取可配置**：新增 `enable_pdf_text_extraction` 配置项（默认开启），关闭后可强制走 OCR 路径，用于对比测试或规避特定 PDF 文字层异常

### 🐛 修复

- **dzcp 格式 PDF 名称/信用代码提取失败**：CJK 拆字导致"名称:"/"统一社会信用代码:"标签被拆成单字，`_extractNamesByCoords` 和信用代码坐标提取无法匹配。新增虚拟标签合成逻辑（"名"+"称"配对、"统"+"一"+"社"配对），合成完整标签后正常提取
- **dzcp 格式发票类型检测返回 unknown**：同上，类型检测依赖"价税合计"等关键词，拆字后无法匹配。新增空间位置回退检测（`_detectInvoiceTypeFallback`），直接扫描所有文本块内容
- **dzcp 格式金额提取失败（数字跨行折叠）**："3.15" 和 "30" 被拆到两行导致拼接错误。新增数字换行折叠合并逻辑，`_extractAmountsByCoordinates` 中先合并再解析
- **dzcp 格式金额提取失败（中文大写金额换行折叠）**："（小写）" 和 "¥70000.00" 之间有 "柒万圆整" 等中文大写字符，Pattern1 的 `\s*` 无法跨越。新增 Pattern1c，允许中文大写字符（壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整正）
- **dzcp 格式 ¥ 符号提取为 ·（中点）**：部分 PDF 文字提取将 ¥ 符号提取为 `·`（U+00B7），导致 Pattern1/2 无法匹配。修复：所有金额正则统一使用 `[¥·•]` 字符类
- **dzcp 格式金额提取（小数点后空格）**："94. 00" 因 kern 断词导致小数点与小数部分分离。新增小数点后空格折叠处理
- **dzcp 格式购方信用代码提取失败**：标签与值之间因拆字产生换行，`_extractCreditCodesByCoords` 正则改为 `[\s\S]*?` 跨行匹配
- **非税发票金额异常（中文大写金额误解析）**：`parseChineseNumeral` 将非税票据的"柒万圆整"等解析为超大金额。修复：`_isNonTax=true` 时 `amountNoTax` 强制等于 `amountTax`，忽略中文大写金额解析结果
- **0% 税率发票税额错误**：税额为 0 时 OCR/文字提取可能返回异常值。多处加守卫：`extractByCoordinates` 中税额为 0 时跳过坐标提取、`_extractAmountsByText` 中 `amountTax - amountNoTax != taxAmount` 时重置 `taxAmount = 0`、`_applyOcrTextResult` 中 `taxAmount > amountTax` 时重置
- **话费发票金额异常**："价税合计（大写）柒万圆整（小写）¥70000.00" 格式导致金额提取取到极大值。修复：`_extractAmountsByText` 中 Pattern1c 允许中文大写字符、`_chineseNumeralToNumber` 兜底返回 0 而非 NaN
- **虚拟打印机（Microsoft Print to PDF 等）无法使用**：SumatraPDF 的 `-silent` 参数会阻止虚拟打印机的保存对话框弹出，导致无法选择保存位置。修复：检测虚拟打印机（Print to PDF / Microsoft Print to PDF / OneNote / Fax），对其不使用 `-silent` 参数，让用户可以正常交互

### 🔧 优化

- **虚拟标签合成**：遍历所有文本块，将相邻的单字文本块（"名"+"称"、"统"+"一"+"社"+"会"+"信"+"用"+"代"+"码"）合成为虚拟标签，用于名称/信用代码/发票类型检测
- **`merge_adjacent_words` 增强**：处理 kern 断词导致的数字碎片（如 "94. " + "00"）

---

## v1.9.5 — GBK/Big5 编码支持 + 火车票识别修复

### 🐛 修复

- **GBK-EUC-H / Big5 编码乱码**：lopdf 对 `GBK-EUC-H`、`ETen-B5-H` 等 CJK SimpleEncoding 返回 `CharacterEncoding` 错误，导致文字型 PDF 提取结果为乱码。现新增 `encoding_rs` 回退解码：lopdf 失败后自动按编码名选择 GBK/Big5 解码，中文文本正确输出
- **火车票类型检测不准**：`_detectInvoiceType()` 改为扫描全部文本（不再仅前 60%），新增"电子客票号"/"铁路电子客票"关键词，增加"有购买方无销售方"二级检测，减少误判
- **火车票票价提取为 0**：新增文本正则 Method 0（`票价:94.00` 模式），扩大坐标搜索半径（maxDx 500, maxDy 50），放宽金额上限至 50000，新增 Method 3 兜底（全文最大合理金额），扩展纵向范围 ny 0.2~0.8
- **车票信用代码归属**：车票唯一信用代码归购买方，不再误赋给销售方
- **扫描件 OCR 回退**：PDF 文字层为空（无 CMap/扫描件）或提取失败时，自动回退 OCR

---

## v1.9.4 — PDF 文字层提取（轻量版也能识别文字型PDF）

### 🚀 新功能

- **PDF 文字层提取**：解析 PDF 内容流（Tm/Td/Tj/TJ/T* 操作），从字体 Encoding 解码文本，输出带坐标的词级 bounding box。5ms/页，无需 OCR 引擎
- **轻量版也能识别**：文字型 PDF 无需 OCR 版即可提取发票字段，轻量版从"完全不能识别"→"大部分可识别"
- **OCR 自动跳过**：PDF 文字提取覆盖关键字段（销售方 + 金额）后自动跳过 OCR（省 1-3s/页）

### 🐛 修复

- **金额求和验证**：`applyPdfTextResult` 补齐 `含税≈不含税+税额` 校验
- **车票标记**：`applyPdfTextResult` 补齐 `_isTicket` 设置
- **Tm 字号**：改为取矩阵垂直分量 `d`（而非水平 `a`），避免旋转文本字号误判
- **Td 行首偏移**：`cur_x` 改为从 `line_start_x` 偏移（PDF 规范），不再累加
- **BT 状态重置**：补齐 `leading` 归零
- **Tj/TJ 间空格分隔**：相邻文本片段间插入空格，防止粘合破坏正则匹配
- **TJ kern 断词阈值**：`kern < -80` 才断词，小 kern 不再误拆词

### 🔧 优化

- **TL 操作支持**：解析 `TL` 设置行距，`T*` 使用实际 leading 值
- **sellerCreditCode 车票过滤**：车票不设 `sellerCreditCode`
- **ToUnicode CMap 预留**：`decode_bytes_with_encoding` 增加 `_pdf_doc` 参数

---

## v1.9.3 — OFD ImageMask 遮罩兼容

### 🐛 修复

- **OFD ImageMask 遮罩不生效**：iloveofd 等二次转换工具生成的 OFD 文件，红章图片使用 `ImageMask` 属性引用 BMP 遮罩图（0=透明，255=不透明）。此前引擎未解析该属性，导致遮罩不生效、BMP 黑色区域变成图片黑色背景。现已在解析阶段提取 ImageMask，图片加载时将遮罩合成为主图 alpha 通道，输出 RGBA PNG
- **BMP 图片格式支持**：`image` crate 新增 `bmp` feature，可解码 BMP 遮罩图；BMP 文件统一以 PNG 格式输出（浏览器不原生支持 BMP data URL）

### 📝 说明

- 不同厂商/转换工具生成的 OFD 发票格式存在差异（如税务原版 OFD、iloveofd 转换、dzcp 公共服务平台等），如遇解析渲染问题请及时反馈，我们会持续适配

---

## v1.9.2 — OFD 发票字段提取修复

### 🐛 修复

- **OFD 自闭合标签字段丢失**：`parse_ofd_custom_data` 中 `Event::Empty`（自闭合标签如 `<CustomData Name="购买方纳税人识别号"/>`）错误调用 `read_element_text()`，该函数会消费后续兄弟节点的 XML 事件，导致购方税号为空时销售方税号也被清空，且后续字段级联错位
- **CustomData 空白值过滤**：空字符串不再覆盖已有有效值
- **BuyerTaxID/SellerTaxID 回退**：增加 CustomTag.xml 中 `BuyerTaxID` / `SellerTaxID` 的 ObjectRef 查找回退

---

## v1.9.1 — PDF解析兼容性修复 + 导入体验优化

### 🐛 修复

- **EXIF 方向**：修复图片（车票等）导入后方向颠倒的问题，读取 EXIF Orientation 自动旋转像素
- **PDF /Rotate 属性**：修复带旋转标记的 PDF 页面直通后内容颠倒，`get_page_rotation()` 正确处理
- **CropBox 坐标归一化**：修复部分 PDF 的 CropBox 为倒序（y1 > y2）导致负高度内容翻转
- **OFD lopdf 压缩**：已有 Filter 的流不再二次压缩，避免 OFD 文件变黑
- **FlateDecode 可靠性**：恢复 JPEG DCTDecode 直通策略（OFD 文字变淡是图片型 vs 矢量型的区别）

### ⚡ 优化

- **导入加载体验**：骨架屏秒出 + 逐文件渐进加载 + 持久进度 toast 替代全屏遮罩

---

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

| 文件 | 说明 |
|------|------|
| `发票酱_x64-setup.exe` | 轻量版安装包 |
| `发票酱_x64_绿色版.exe` | 轻量版便携（单文件） |
| `发票酱_x64_OCR版-setup.exe` | OCR 版安装包 |
| `发票酱_x64_OCR绿色版.zip` | OCR 版便携 |

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

| 文件 | 说明 |
|------|------|
| `发票酱_x64-setup.exe` | 轻量版安装包 |
| `发票酱_x64_绿色版.exe` | 轻量版便携（单文件） |
| `发票酱_x64_OCR版-setup.exe` | OCR 版安装包 |
| `发票酱_x64_OCR绿色版.zip` | OCR 版便携 |

---

## v1.7.7 — OCR Feature Flag（轻量版/OCR版双构建）

### 🚀 重大变更

- **OCR 功能改为可选 Feature Flag**：同一套代码，编译时决定是否包含 OCR
  - 轻量版 `npm run build`：无 OCR
  - OCR 版 `npm run build:ocr`：含 PP-OCRv5
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

| 文件 | 说明 |
|------|------|
| `发票酱_x64-setup.exe` | 轻量版安装包 |
| `发票酱_x64_绿色版.zip` | 轻量版便携 |
| `发票酱_x64_OCR版-setup.exe` | OCR 版安装包 |
| `发票酱_x64_OCR绿色版.zip` | OCR 版便携 |

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
