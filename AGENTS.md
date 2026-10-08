# 发票酱 — Agent 指南

> 版本历史见 CHANGELOG.md，本文只描述**当前**架构与约定，不记录演进过程。

## 项目概览

- **版本**: v2.6.11（数据源 `package.json`，`npm run bump` 同步到 Cargo.toml + tauri.conf.json；`Cargo.lock` 的 `ticketchan` 包版本行需手动同步）
- **技术栈**: Tauri 2.x (Rust) + 原生 HTML/CSS/JS（无框架、无打包）
- **双版本**: 轻量版 / OCR 版（PP-OCRv6）；Cargo.toml 定义 `ocr` feature，`lib.rs` 按 `#[cfg(feature = "ocr")]` 条件注册命令，OCR 构建用 `tauri.ocr.conf.json` 叠加配置（仅追加 bundle.resources）
- **目录结构**:

| 路径 | 内容 |
| --- | --- |
| `src/` | 桌面版前端：`index.html` `styles.css` `ocr.js` `layout.js` `print.js` `app.js` |
| `src-tauri/src/` | Rust 后端：`main.rs` `lib.rs`（IPC 命令层）`pdf_engine.rs`（PDF 生成/渲染）`pdfium_print.rs`（直打引擎） |
| `src-tauri/invoice-engine/` | 独立 crate：OFD / XML 数电票解析 |
| `web` 分支 | 纯浏览器版（pdf-lib 实现），目录为 `js/` `css/` `vendor/` |

## 常用命令

```bash
npm run dev             # 轻量版开发
npm run dev:ocr         # OCR 版开发
npm run build           # 轻量版构建
npm run build:ocr       # OCR 版构建
npm run build:all       # 全量构建，产物输出到 dist/
npm run bump <版本号>    # 同步版本号
```

- **编译缓存**: 只改 HTML/JS/CSS 不触发 Rust 重编译；改 Rust 文件才会完整重编译
- **CI/CD**: GitHub Actions，push tag `v*` 触发，产出 4 个安装包（轻量/OCR × setup/绿色版）
- **⚠️ 产物文件名必须全 ASCII**：GitHub Release 上传时会剥掉文件名里的非 ASCII 字符（中文 → 下划线），下载得到的文件名会残缺（曾出现 `_2.6.4_x64_.exe`）。`scripts/build-all.js` 的 `FINAL_FILES` 统一用 `TicketChan_<version>_x64[_ocr][-setup|_portable]` 命名，CI 的 `Verify artifacts` 会逐个断言这些名字；`productName`（`发票酱`）只作程序显示名，**不得**介入产物命名

## 架构总览

核心数据流（桌面版）：

```
用户文件 (jpg/png/pdf/ofd/xml)
  │  app.js 批量加载（open_invoice_files 一次 IPC；PDF 经 WinRT/PDFium 渲染缩略图）
  ▼
S.files[] — fileObj { previewUrl, ow/oh, rotation, slotScale/Offset, checked, ... }
  │  layout.js 预览渲染（calculateLayout + renderPage，CSS transform）
  ▼
print.js buildLayoutRequest() — files 去重 specs + pages 槽位矩阵 + settings
  ▼
Rust generate_pdf_from_layout() — lopdf 直通管道 → 失败回退 printpdf 管道
  ▼
四种打印模式输出（PDF阅读器 / 弹窗确认 / PDFium / SumatraPDF）
```

关键分层约定：

- **预览与导出必须语义一致**：预览（CSS）与 PDF 生成（Rust/web）对旋转、缩放、偏移、适配的计算公式互为镜像，任何一方改动必须同步另一方（见「旋转与适配语义」）
- **桌面/web 双分支同步**：字段提取（`ocr.js` ↔ `js/pdf-text.js`）、预览布局（`layout.js` ↔ `js/layout.js`）逻辑一致，识别与排版改动必须双向同步
- **坐标体系**: JS 端 top-down（y 向下），PDF 端 bottom-up（y 向上，PDF 标准），各自独立计算，**不做相互转换**

## 核心机制

### PDF 生成双管道

入口 `generate_pdf_from_layout()`：首选 **lopdf 直通管道**（矢量无损）→ 任何错误自动回退 **printpdf 渲染管道**。

**lopdf 直通** `generate_pdf_passthrough()`：

- PDF 页面 → `extract_page_as_form_xobject()` 提取为 Form XObject（矢量保真），标注（印章/签章）从 `/Annots` 经 `/AP /N` 烘焙进内容流（跳过隐藏标注；后缀 `Q` 必须先于标注绘制命令）
- 图片/OFD → `image_to_lopdf_xobject()` 编码为 JPEG DCTDecode Image XObject；旋转烘焙进像素（`RENDER_DPI=300` 换算 pt）
- 源 PDF `/Rotate` 属性烘焙进内容流前缀（PDF spec：显示时顺时针旋转），90=`[0 -1 1 0 0 w]`、180=`[-1 0 0 -1 w h]`、270=`[0 1 -1 0 h 0]`
- 页面组装：`build_nup_content_stream()` 按 slot 计算 cm 矩阵 + `q re W n` 槽位裁剪 + Do；页脚/编号/水印为 PNG Image XObject 叠加

**printpdf 回退** `build_page_ops()`：

- `get_cached_xobj()` 按 `(file_idx, rotation)` 缓存 XObject；Decoded 图片像素级烘焙旋转；JpegPassthrough 仅 0°/180° 直通（180° 用 PDF 层 `rotate_op` 补转，**仅限 JpegPassthrough**——Decoded 已烘焙，再转会双重旋转抵消）

**老打印机兼容模式**（`settings.compat_flat`，issue #46，默认关）：

- 生成完成后 `apply_compat_flat()` 把整份 PDF **逐页栅格化**（原地覆写）：PDFium 白底渲染 300dpi（`COMPAT_FLAT_DPI`，含标注/签章）→ PNG dataURL → JPEG → 重建 PDF 1.4 + **经典 xref 表**、每页一个 Image XObject 的纯位图 PDF
- 目的：消除 SMask 透明、嵌入字体、对象流/交叉引用流——只认图元的老 RIP / 老打印机驱动不再丢字丢章，等效「Ghostscript 重写为 PDF 1.3」；代价是失去矢量锐度、文件变大，故只在设置里显式开启
- 打印与「保存为 PDF」共用生成链，开启后一并生效；`compatFlat` 属**内容参数**，留在 `print.js` 缓存比较内（切换即失效缓存，勿加入 `_cacheExclude`）
- 依赖 PDFium 组件（与静默打印同一个 DLL），未下载时返回可读错误

### 打印体系

**四种模式**（`print.js doPrint` 分发，各自独立调用命令，不经隐式降级）：

| 模式 | 实现 | 说明 |
| --- | --- | --- |
| PDF 阅读器（默认） | `ShellExecuteW` print/printto | ⚠️ 无法可靠控制打印机选择（Edge/Chrome 查看器不支持 printto），选了具体打印机时 toast 引导改用 PDFium |
| 弹窗确认 | 自定义对话框 → PDFium/SumatraPDF | — |
| 静默打印 PDFium（推荐） | `pdfium_print.rs` 矢量直打 DC | 失败自动降级 SumatraPDF（`doPdfiumPrint` 兜底） |
| 静默打印 SumatraPDF | 命令行 `-print-settings` | — |

**PDFium 细节**（`pdfium_print.rs`）：

- DLL: `{exe}/tools/pdfium.dll`，下载源 `bblanchon/pdfium-binaries` via `gh-proxy.com`，`AtomicBool DOWNLOAD_CANCELLED` 支持取消
- 渲染：`FPDF_LoadMemDocument` → 逐页 `FPDF_RenderPage(printer_dc)` 原生 DPI；SEH 捕获异常时 fallback `FPDF_RenderPageBitmap` + `StretchDIBits` 位图（`seh_wrapper.c`，`cc` 编译为静态库）
- DEVMODE: `build_dev_mode()` + `infer_paper_size()`；`get_printer_default_devmode()` 返回**完整** `Vec<u8>` 缓冲区（含 `dmDriverExtra` 驱动私有数据）

**智能 PDF 缓存**（`print.js`）：

- `deepEqual()` 深比较整个 `LayoutRenderRequest`，`canUseCachedPdf()` 判断复用——统一三个打印渠道 + 保存 PDF
- `savePdf` 先生成到临时目录 → `updatePdfCache()` → `copy_file` 到用户路径；布局不变时直接复制缓存
- 缓存比较排除纯打印机参数（printerName/copies/duplex/collate），避免切打印机误判缓存失效

### 旋转与适配语义（全链路约定，issue #29）

统一约定：**正值 = 顺时针**（与 CSS `rotate(N deg)` 一致），**先旋转后适配**（旋转后的视觉宽高 contain-fit 槽位）。

- **PDF 坐标系 y 向上**：cm 矩阵 `[0 sy -sx 0]` 是逆时针、`[0 -sy sx 0]` 是顺时针。图片像素路径 `image crate rotate90()` 恰为顺时针，与 CSS 天然一致
- **预览**（`layout.js renderPage`）：90°/270° 时 wrapper 取视觉盒的**转置**，CSS 旋转落位后即视觉盒。同步点三处——`renderPage`、拖拽偏移约束（`onSlotMouseMove`）、`setSlotAlignment`（九宫格对齐），**必须三处同步**
- **Rust lopdf**：图片走像素烘焙 + `adjustment.rotation=0`；PDF 页面走 cm 矩阵（rotation 保留），两路径语义等价
- **web**（pdf-lib）：`drawImage/drawPage` 的 rotate 绕 **(x,y) 锚点**（未旋转盒左下角）且正角度为逆时针——须传 `degrees(-rot)` 并换算锚点 `(cx,cy) - R·(w/2,h/2)`；水印角度同样取负
- **验证方法**：红色象限测试源 + WinRT `render_pdf_pages` 渲染输出找红色质心，断言落点象限

### 预览与版面交互

**布局计算** `layout.js calculateLayout()`（纯函数，预览/打印共用）：

- slot 尺寸由纸张、行列、边距、间距、页脚扣除推导；cutLines 基于 slot 实际边界
- **每页 slot 数一律走 `getPerPage(s)`**（报销单单列分段 vs 网格 cols×rows），禁止直接写 `cols * rows`
- 页脚边距模型：footerMargin 是纸底额外独立空间，不影响 slot 边距

**报销单分段模式**（`S.feat.reimburse`，默认段高 120mm）：单列 N 段（N=⌊paperH÷seg⌋），mt/mb 为段内安全边距，裁切线在 k×seg 绝对位置强制绘制（不经 cutline 开关）；发票**左上对齐**——JS `renderPage`、Rust `build_nup_content_stream`/`build_page_ops`、`setSlotAlignment` 基准三处同步；rows/cols/gap/footerMargin 扣除均忽略（UI 置灰 `syncReimburseUI()`），关闭后网格布局原样恢复。

**粘贴单模式**（`S.feat.pasteMode`，v2.6.6）：表单式整页版式，票据只排在装订线上方、签字栏下方；**与报销单分段互斥**（开启互关，`enablePasteMode`/`enableReimburse` 同步 UI）。独立边距（`pasteTop/Bottom/Left/Right`，默认上 20.8mm 装订区 / 下·左·右 4.1mm）+ 装订线（`pasteBindLine` 开关 + `pasteBindText` 文字 + `pasteBindSize` 字号）+ 右下角签字栏（`pasteShowSig` + `pasteSigCols` 列名 / `pasteSigWidth` 宽 / `pasteSigRowH` 表头行高 / `pasteSigBodyH` 填写行高 / `pasteSigGap` 区间距），签字栏留空打印后手写；粘贴单下页码/日期/自定义页脚停用（`getPageSettings` 置空）；Rust `calculate_layout_mm` 与 JS 解析规则逐字对应（预览/PDF 同源）；`applyPastePreset()` 一键恢复标准边距，`resetSettings` 的 `layoutOnly` 分支须同步粘贴单 UI。

**单票独立调整**：`fileObj.{slotScale, slotOffsetX, slotOffsetY}`，CSS transform 预览 + Rust `SlotSpec` 参数输出。九宫格快速对齐、数字框/滑块滚轮微调、选中后滚轮缩放单票（5%/步）、拖拽约束按实际显示尺寸动态计算、放大上限 5x、编辑态溢出可见（`.selected/.dragging` 时 `overflow:visible`）。持久化：`perFileAdjustments` Map 按文件名匹配，可选开关。

**选中票面浮动工具条**（v2.6.1，v2.6.8 钳制 + 删除，v2.6.9 按格子归位）：`app.js syncSlotToolbar()` 把 `#slotToolbar` 锚在选中槽位上方 36px 居中（CSS `translateX(-50%)`），随 `previewWrap` 滚动与窗口 resize 重算。⚠️ 四条硬约定：①上方那块位置**一旦与别的格子相交就贴回本格子内侧顶部**（下排格子、报销单第 2 段起、3×3 中下排都命中）——否则工具条骑在裁切线上压住邻格票面，就是 issue #43① 的「跑到框外」；②**水平必须左右钳制**（内侧放置先钳在本格子框内，再统一钳在预览区可视宽度内各留 8px）——否则右列槽位的右半截会跑出预览框；③**先 `classList.remove('hidden')` 再读 `offsetWidth`**，`display:none` 时宽度为 0 会让钳制失效；且**旋转按钮文案要在测量之前刷新**（`旋转` → `90°` 宽约 +12px，否则钳制用过期宽度）。工具条宽于格子时自动加 `.compact`（只留图标）并重测宽度。按钮：重置 / 居中 / 应用到全部 / 旋转 / **✕ 删除**（`deleteSlotInvoice()`，删完清空选中）。

**预览滚轮交互**（`previewWrap` wheel 三分支，按优先级）：选中槽位+悬停 → 缩放单票；Ctrl+滚轮 → 缩放整体视图；普通滚轮 → 滚动内容，触顶/触底翻页。`_wheelFlipTs` 150ms 节流。

**版面拖拽排序**（v2.5.0+，与单票偏移拖拽共用 `_slotDrag` 状态机 `mode:'move'`）：

- 双手势：`updateDropTarget()` 按目标槽位主轴 25%/50%/25% 分区——边缘 before/after 顺位插入（`moveSlotInvoice`），中间对调（`swapSlotInvoices`）；落点 `elementFromPoint`
- 有效落点仅限装有发票（含占位）的槽位，尾部空槽拒收；倒序打印时 before/after 反映射；跨槽松手回滚拖动产生的偏移（排序手势不得产生单票偏移）
- ⚠️ `_dragHintShown` 必须在 layout.js 顶层声明——未声明时 ReferenceError 被 mousemove 事件边界吞掉，落点判定静默全灭
- 尾部空槽临时占位（v2.5.1）：按下尾部第一个空槽拖动即 `insertTempPlaceholder()` 临时创建占位进链路，松手落实体之间=中间留白；无效落点/拖回尾部则销毁；`_slotSuppressClick` capture 阶段吞一次合成 click

**槽位精准上传与留白**：空槽点击上传精准落位（`prepareSlotInsertion()` 返回 `{insertAt, blankCount, replaceIdx, reverse}`）；空白占位 `fileObj._placeholder` 只占槽位不打印不统计，`getActiveFiles()` 过滤条件 `(f.checked || f._placeholder)`，其余消费点全部排除；占位无 `_filePath` 不持久化。三条加载路径（`processFileDataList`/`processFiles`/`processFilesIncremental`）同步支持插入替换；`_slotUploadActive` + `_loadingBatchActive` 并发锁。

**列表与版面双向联动**：`clickFileItem()` 正向（activeIdx → 翻页+选槽），`syncSidebarToSelectedSlot()` 反向（高亮+滚动定位）。`clickFileItem` 第三参 `opts.autoCheck:false` 供右键联动复用（只同步选中态不改勾选）。

**右键菜单**（v2.6.0，`_ctxIdx` 状态机）：全局 `contextmenu` 分发——`.file-item/.file-card` 命中则 preventDefault 并弹 `#ctxMenu`（份数 ×1/×2/×3、旋转、OCR（`hasOcr` 显隐）、复制发票信息、删除，作用于被右键单项）；`input/textarea/contenteditable` 放行系统菜单；其余区域仅屏蔽。联动经 `clickFileItem(idx, null, {autoCheck:false})`，右键不改勾选。菜单定位防溢出翻转，click 外点与 fileList scroll 关闭。`ctxCopyInfo()` 非空字段逐行拼接（`label：value`），无信息时 toast 不复制。

**预览副本标记**（v2.6.0，`S.feat.copyBadge` 默认关）：layout.js `renderPage` 按全局展开序列预计算 `copySeq[n]/copyTotal`，同发票副本槽位显示 `n/N`（左上角，仅预览 DOM，不进打印 PDF）；新增 feat 开关必须四处同步——`S.feat` 默认值、`saveSettings` featKeys、`loadSettings` featMap、`getSettings` 显式传递，纯预览参数须加入 print.js `_cacheExclude` 防缓存误失效，`resetSettings` 的 `S.feat` 字面量与按钮 `.on` 重置同步。

**快捷布局**：工具栏由 `S.quickLayouts` 动态生成；默认值必须经 `defaultQuickLayouts()` 深拷贝（禁止 `slice()` 共享）；`normalizeQuickLayoutValue()` 限 1-10；允许空列表（`loadSettings` 按 `Array.isArray` 恢复）；不得按内容强制迁移旧配置（无法区分用户自定义）。

### 文件加载与列表管理

**批量加载** `processFilesIncremental`：`open_invoice_files({paths})` 一次 IPC 读全部 → `Promise.all` 并行渲染 → `setInterval` 定时批量刷 DOM + toast 100ms 防抖。图片缩略图带 EXIF 烘焙与原始尺寸（`origW/origH`）；预览 PDF 用 `PDF_PREVIEW_DPI=150` + JPEG（`useJpeg:true`），打印/保存独立走 300 DPI 矢量管道互不影响。

**PDF 渲染双引擎**：首选 WinRT（`render_pdf_pages`，`check_winrt_pdf_available()` 启动检测）→ 失败回退 PDFium（`render_pdf_pages_pdfium`）。

**筛选体系**（侧边栏，可折叠）：类型（专票/普票/车票/通行费/医疗/非税/其他 `S.typeFilter`）× 格式（PDF/OFD/图片/XML `S.formatFilter`）× 状态（全部/未打印/已打印/重复 `S.printedFilter`/`S.fileFilter`）三维正交；类型匹配与票种 chip 共用 `resolveInvoiceType` 单一真源（车票/通行费/医疗直判 `_isTicket`/`_isToll`/`_isMedical` 分类标记），「其他」= 不渲染票种 chip 的文件（未识别 + 无法归专普的粗粒度串），非税/专票/普票对 XML/OFD/PDF 文字层等非 OCR 来源同样生效；类型/格式切换即切换打印批次，`clearInvisibleChecks()` 清除不可见勾选（防筛选切换后残留勾选重复打印）。列表上方工具栏（`.file-header`）与筛选区（`.filter-section`）sticky 钉在滚动容器顶部，仅列表滚动。

**文件列表双视图**：`S.fileView`（list/grid），`renderFileList()` grid 分支输出 `.file-card`；`updateFileItem()` 按视图增量更新。

**份数概念区分**（易混淆）：「排版份数」= 每张发票在版面中重复几次（`fileObj.copies`，列表 ② 按钮批量设 ×1/×2/×3，`getActiveFiles()` 展开）；「打印份数」= 整版打印几份（全局 copies，SumatraPDF 经 `-print-settings Nx` 处理，不展开）。

**打印状态追踪**：四种模式成功后 `markFilesAsPrinted()` → ✓ 标识；`_printedMap` 持久化 localStorage；`clearAll()/executeRename()/resetSettings()` 迁移 key。

**文件列表记忆**（可选 `S.feat.fileListMemory`）：启动 `restoreFiles()` 批量恢复路径，`check_path_exists` 校验；`_isRestoringFiles` 阻止恢复期触发 OCR。

**重复发票识别**：`getDupKey()` 按发票号（`no:`）或 销售方+金额+日期（`sum:` 疑似）生成 key 标记 `_dup`。**安全边界**：自动删除只信任 `no:` key，`sum:` 一律跳过（同日同销售方同金额的真发票会被误判，仅标记交人工核对）；「重复」筛选会覆盖原有勾选（toast 明示）。

**图片文本增强**（`toggleTextEnhance`，纯本地）：Rust `enhance_image()` 读原图全分辨率 → EXIF 烘焙 → 直方图 1%/99% 色阶拉伸 + gamma 1.4 + USM 锐化 → JPEG q92；同一 LUT 应用 RGB 三通道（红章保色），退化图恒等映射防噪点放大。`f._enhanced` + 备份可还原；打印链路走 dataUrl 分支（去重 key 改用 previewUrl）；限图片文件且有 `_filePath`（web 未移植）。**核心 `enhance_rgb_inplace()` 与打印自动增强共用（单一真源）**，`normalize_for_jpeg()` 统一 16 位/32F 高精度图。

**清晰度体检与打印自动增强**（v2.6.6，issue #39）：`audit_clarity`（`async fn` + `spawn_blocking`）只读文件头毫秒级算每张发票折算打印 DPI（矢量电子发票 `kind='vector'` 与分辨率无关刻意不参与），低于阈值前端打 `clarity-badge` ⚠ 徽章；`S.feat.autoEnhance` 开启后打印/保存时对折算 DPI < `enhanceMinDpi`（默认 250）的图片自动增强（`EnhanceParams{minDpi, gamma, amountPct, quality}`，滑块越界 clamp 而非拒绝），读原图全分辨率、预览缩略图永不作增强源。

**白边裁剪坐标换算**（v2.6.6，issue #38）：`trimmedBox` 基于预览缩略图（`THUMB_MAX_DIM=600`）坐标，而图片文件 `ow/oh` 是原图尺寸——进 `SlotSpec.trimBox` 前必须按 `ow/tw`、`oh/th` 比例换算到原图坐标（否则 Rust 读全分辨率原图执行裁剪时整体偏移，打印与预览不一致）；PDF/OFD 页面 `ow/oh` 即渲染位图尺寸，无需换算。裁剪走 `trim_images_batch`（`async fn` + `spawn_blocking`，Rust 内 rayon 并行，一次 IPC 返回整批 `[Option<TrimImageResult>]`，单张失败为 `null` 由前端跳过）。

### 发票识别与数据提取

**路径优先级**: PDF 文字层 > OFD XML > XML 数电票 > OCR。OCR 跳过条件：`_pdfTextExtracted && sellerName && amountTax > 0`。

**类型检测** `_detectInvoiceType()`（ocr.js）：ticket > toll > medical > nontax > vat > ride > unknown。

- 专票 / 普票：`_detectVatSubtype()`（仅 vat 路径判定）—— 票头标题区 `ny < 0.18` 优先、全文兜底；关键词 `普通发票|增值税普通|电子普通` → 普票、`专用发票|增值税专用` → 专票，**「普通」优先于「专用」**（票面其它位置的「专用」字样不致误判）；结果写回 `fileObj.invoiceType`（不覆盖已有结构化类型）
- ticket：强标记（铁路电子客票/电子客票号）直判 + 弱信号 `_countTicketSignalGroups()` 13 组关键词 ≥2 组确认，防增值税票误判；`getTicketTypeLabel()` 细分标签
- toll（通行费）：「通行费」强标记；「车牌号/车牌颜色+通行日期」弱标记双组确认；复用 VAT 提取链路（销售方=路桥公司），不走 ticket/nontax 早退分支；老式纸质票无价税合计时两段式金额兜底
- **medical（医疗收费票据）**：强标记「医疗门诊/住院/急诊收费票据|医疗收费明细」，**必须在 nontax 判定前**（否则「票据代码/票据号码/交款人」先抢走判成非税）；复用非税提取链路（金额合计/票据号码/交款人），`getMedicalLabel()` 细分门诊/住院/急诊标签；报销口径只取金额合计，不识别自费/统筹等支付分解字段。明细页（第二页起「医疗收费明细」长清单）由 `isMedicalDetailPage()` 双确认（`医疗收费明细` + `所属电子票据号码`，「所属」前缀是明细页独有特征），`finalizeMedicalDetailPages()` 从列表移除——只留汇总首页、不参与排版/统计/去重，主票挂「附N页明细」徽章并 toast（幂等 `_medDetailHandled`，与逻辑票聚合 issue #40 共存，移除后 `updateDuplicateMarks()` 重排）；web 分支识别同步（js/pdf-text.js）

**金额提取**：含税价 → 数学验证配对 → 区域解析三阶段；中文大写 `parseChineseNumeral()` 兜底；金额求和校验失败时卡片 ⚠ 徽章 + hover 详情 + 汇总栏计数。

**多页发票逻辑票聚合**（v2.6.7，issue #40）：**物理层逐页输出不变**，逻辑层新增分组——`rebuildPdfInvoiceGroups()` 把同批次同 PDF（`getPdfBatchKey` = `_pdfPath|_batchId`）且组内非空发票号去重后**唯一**的多页归为一张逻辑票，写回 `_invoiceGroupId`（= 批次 key）+ `_multiPageInvoice{total,pageNo,isSummary}`；**合计页** = 组内最后一个识别到含税金额的页（无则取末页），统计口径（底部汇总 / 汇总表合计行 / CSV）经 `isMultiPageDetail()` 跳过非合计页，明细页仅回填合计页票种（只补空不覆盖，金额不回填）；去重 `updateDuplicateMarks()` / `removeDuplicates` 按 `getLogicalInvoiceId()` **整组计数 + 整组删除**，杜绝续页被当成重复自动删除导致打印缺页；续页徽章 `buildMultiPageBadge()`「续 n/N」「共 N 页」列表 / 卡片双视图共用。分组**幂等**，依赖已识别的 `invoiceNo`，识别异步完成后须重跑；`_placeholder` / `_xmlInvoice` 不参与。

**购销方识别**（表头锚点 + 交叉验证）：`_determineLabelSide()` 用「购买方/销售方」表头 x 坐标作区域锚点（支持融合词与 CJK 拆字）；`_getSideBoundary()` 动态边界（双表头中点/单表头±0.25/无表头 0.5）；`_crossValidateBuyerSeller()` 四规则（同名清空 sellerName、位置反了交换、信用代码位置交换、同侧迁移）；`_headerCache` 按 words 引用缓存防重复扫描。

**CJK 拆字兜底**（dzcp/iloveofd 格式）：信用代码全文拼接匹配 18 位；正则支持字母夹数字 + 15/18 位校验；名称括号保留；「年/MM/月/DD/日」序列合并；`_cleanName` 清理日期碎片。

**PDF 文字层提取**：`extract_pdf_text(s)` 解析 lopdf content stream（批量版一次打开 + rayon 并行，按 pdfPath 分组调用，失败回退单页→OCR）；前端 `applyPdfTextResult()` 复用 `extractByCoordinates()`。坑：Form XObject 需展开、GBK-EUC-H 需 `encoding_rs::GBK.decode()`、`Content::encode()` 尾部无换行、内容流顺序≠视觉顺序（金额取最大 ¥）。

**XML 数电票**（`invoice-engine::parse_xml_invoice`）：纯结构化数据无版式，`fileObj._xmlInvoice=true`，不参与排版打印（`getActiveFiles()` 过滤）；用于列表展示/统计/汇总/重命名。

### 导出与工具命令

**汇总表**（侧边栏 📊）：14 字段按需勾选、双击编辑回写全 UI 同步、三金额合计行 sticky；`exportSummaryCsv()` UTF-8 BOM + CRLF 手写 CSV → `write_text_file`；数据源 `getCheckedFiles()`（不含 copies 展开）。「发票类型」单一真源 `resolveInvoiceType(f)`（分类标记 通行费/车票/非税 → 结构化或 OCR 类型 → 兜底「发票」），`normalizeInvoiceType()` 归一为「专票 / 普票」短标签，列表徽章 / 复制发票信息 / 重命名 / CSV 共用同一口径；**禁止再写死「增值税发票」**（issue #35：普票全被显示成该串）。内嵌批量重命名面板：3 预设模板 + 自定义字段（勾选顺序=文件名顺序）、`resolveNameConflicts()` 自动 `_2` 序号、`executeRename()` → `rename_file` 命令并同步 `S.files` 共享路径与 `_fileAdjMap`/`_notesMap` key；OFD 的 dedup key 排除 `_filePath`。⚠️ **长数字必须按文本导出**：≥15 位纯数字（20 位发票号、18 位税号）用 Excel 文本公式 `="..."` 包裹，否则 Excel/WPS 打开会转科学计数法并丢精度（issue #49）。

**文件命令**（均为 `async fn` + `spawn_blocking`）：`copy_file`、`rename_file`（同盘原子 rename，跨盘 copy+delete）。

**诊断导出**（右键菜单「📤 导出诊断信息」，issue #47 配套）：`diagnostics.rs` 的 `export_diagnostics` 生成「脱敏结构报告」供用户直接贴进 issue——用于替代「截图猜 + 反复要样本」。硬约束：**任何用户原文不得出现**——文本经 `invoice_engine::sanitize_text`（汉字→汉 / 数字→9 / 字母→A / 空格→·，标点保留、长度与字符类型序列完整）、文件名脱敏（去目录、保留扩展名）；结构数据原样（OFD 的 TextObject/ΔX/ΔY 数组、PDF 词级坐标、XML 字段有无、图片尺寸）。前端 `buildDiagSummary` 只允许布尔/枚举等固定词汇（该部分不脱敏）；报告 512KB 截断（按 char 边界）；报告头部含版本/构建/导出时间（时间由前端格式化传入）。

### 设置持久化与更新检查

**设置持久化**：`saveSettings()`/`loadSettings()` — `ticketchan-settings` JSON，覆盖排版/纸张/边距/缩放/旋转/水印/页脚/筛选/视图等；`updatePreview()` 500ms 防抖自动保存；恢复默认清空全部。⚠️ **var 提升坑**：被 `loadSettings()` 恢复的 JS 变量的 `var x = 默认值` 声明必须在调用点之前（声明提升、赋值不提升，曾致 issue #7）。

**导出 / 导入配置**（`exportSettings()`/`importSettings()`，issue #50）：数据同源——`collectSettings()`（saveSettings 的收集逻辑抽出共用）+ 独立存储的偏好键（theme/amtMode/ocrEnabled/pdfTextEnabled/ocrPrecision/saveDir）；导出经原生「另存为」写 JSON（含 `_meta` 头）；导入校验 `_meta.app`（或 layout+feat）→ 写回同名 localStorage 键 → **`location.reload()` 重载生效**（不重载会被下次自动保存覆盖回去），Rust 侧 `read_text_file` 与之配套。

**更新检查**：`check_for_updates`（reqwest 调 GitHub Releases API，主源 `api.github.com` 失败回退 `gh-proxy.com`）；启动 5 秒后静默检查（1 小时缓存 `ticketchan-update-cache` 防速率限制），状态栏版本号/关于面板可手动触发；更新弹窗 `#updateModal`。**忽略体系**（v2.6.0）：`shouldAutoShowUpdate()` 只拦静默弹窗——`ticketchan-update-ignore`（忽略此版本，弹窗按钮）+ `ticketchan-update-ignore-all`（忽略所有，弹窗按钮 + 设置→关于「自动检查更新」开关 `toggleAutoUpdateCheck`，`showApp` 里 `syncAutoUpdateCheckUI()` 同步）；手动检查不受忽略影响。未用 Tauri Updater（4 产物 + 无签名证书，引导用户去 Release 自选）。Release Notes 由 CI 从 CHANGELOG.md 提取 `## v<tag>` 段落写入 `release_body.txt`。

## 前端模块

| 文件 | 职责 |
| --- | --- |
| `app.js` | 主入口、状态管理(S)、批量文件加载、文件列表双视图与筛选、Tauri IPC 分发、设置持久化、XML 数电票加载 |
| `ocr.js` | 发票字段提取、金额解析、中文大写解析、类型检测、金额校验 |
| `layout.js` | 布局计算、预览渲染、单票调整拖拽、slot 交互、版面拖拽排序 |
| `print.js` | 打印/导出、构建 LayoutRenderRequest、智能 PDF 缓存、四模式分发与降级 |

- 顶层变量全部用 `var`（避免与 Tauri 注入脚本冲突）
- 无模块打包，`index.html` 按序 `<script>` 加载

## 关键踩坑

### Tauri 2.x

- **同步命令阻塞 IPC 线程**：非 `async fn` 命令阻塞 IPC 消息泵 → `ERR_CONNECTION_REFUSED`。所有 CPU 密集命令必须 `async fn` + `spawn_blocking`
- `<input>.click()` 无效 → 用 `plugin:dialog|open`
- 关闭必须 `TerminateProcess`，不能用 `process::exit(0)`（MNN/OCR 引擎死锁）

### PDFium / Win32

- `libloading::Library` 不能在函数内创建（drop 时 DLL 卸载致全局崩溃）→ 全局 `LazyLock<Mutex<Option<PdfiumState>>>`，`_lib` 字段持有
- PDFium 非线程安全 → `with_pdfium()` 闭包 + Mutex 串行化
- `DEVMODEW` 嵌套匿名结构 `dm.Anonymous1.Anonymous1.dmCopies`；`dmDuplex` 是 `DEVMODE_DUPLEX(i16)`；`std::ptr::read` 只复制 `sizeof(DEVMODEW)` 会丢驱动私有数据 → 完整 `Vec<u8>`
- `DOCINFOW`/`StartDocW`/`StartPage`/`EndPage` 在 `Win32::Storage::Xps` 模块（不是 Gdi）
- `windows` crate 0.58：`HENHMETAFILE` 是 CopyType，`DeleteEnhMetaFile(h)` 不需要 `&`
- `CreateEnhMetaFileW` 的 `lpRect` 是 0.01mm 单位（直打 DC 时无需 EMF 中间层）

### OFD

- ImageMask 遮罩：二值图合成主图 alpha 通道
- 自闭合标签不能用 `read_element_text()`
- CJK 拆字（dzcp 格式）：需虚拟标签合成
- **TextCode 转义与占位符**（issue #44/#47）：`\XXXX` 四位十六进制转义（标准要求空格等一律转义）必须解码，否则按字面 5 字符渲染；占位符参与 ΔX 定位（占一个字符槽位）但**不渲染字形**（直接输出会与相邻字符叠字），判定集中在 `is_placeholder()`——**只认国标 `¤`（U+00A4）＋ PUA 码位（U+E000–U+F8FF 等，系统字体渲染必乱码）**。⚠️ 不要把 Ø/∅/Φ 等圆圈符号拉进黑名单——issue #47 的「月Ø」已查明是「数字 0 被 ΔX 错位挤到『月』的竖笔上」，与占位符无关，乱扩名单会误伤正文。**不要**把占位符从字符序列里剔除（剔除会与 ΔX 错位），也不要按普通字符渲染
- **一个 TextObject 可含 1..N 个 TextCode**（issue #47，国标 11.3：各带 X/Y/ΔX，标准示例就是两个）：解析为 `OfdTextObject.segments`，渲染逐段输出各自的 `<text>`；对象级 `text` 仍按序拼接（CustomTag 按 ID 取值、文本提取沿用旧口径）、`text_x/text_y` 保留末段（数电票 body 坐标提取依赖）。**禁止**再把多段拼成一行、用末段 X/ΔX 渲染（ΔX 用尽会重复末值，整行被均匀拉开）
- **DeltaX 逐字定位的口径判定**（issue #44/#47）：空格是否参与 ΔX 各生成器不一（规范口径=空格是转义内容的一部分、参与定位，`单··位` 4 字符配 3 个 ΔX；数电票表头把空格当列分隔=不参与，属厂商偏差）。判据优先级：**① ΔX 含负值（换行回退）→ 直接按完整字符流**（`multiline && fits_all`）——多行文本 ΔX 累加和天然 ≠ Boundary 宽（每行宽 ≤ boundary、行间为负回退）使 Boundary 自校验失效，而 ΔX 条数又常同时吻合两口径（不动产证附记栏 131 字符含 1 空格配 130 ΔX：含空格 131-1=130 ✓、去空格后 130 字符 ⚠️ 也"凑巧"= 130）；此时误选去空格会让换行点**后移一个字符**（「2」留在上行行尾、「、」跑到下行行首）。② **Boundary 宽自校验**——ΔX 累加和应≈文字总宽（残差一个末字宽），两口径误差差 2 倍以上才切换；**CTM 含缩放不得跳过本校验**（表46：X/Y 是「对象坐标系」坐标，ΔX/ΔY 与 Boundary 同坐标系）。③ 两口径误差都很大（ΔX 与 Boundary 自身不自洽，如数电票表头样本 184.6 vs 172.6）时退回长度拟合（`invoice-engine/src/lib.rs` `build_svg_text`）
- **表46 的缺省语义**（issue #47）：`DeltaX` 缺省 = 字型在 X 方向**不偏移**（`char_dx()`；勿用字号顶替——纯 ΔY 定位的竖排/垂直文本会被斜着排）；`X/Y` 缺省 = 沿用**上一个 TextCode** 的坐标（对象内首个必需，解析处经 `t.text_x/t.text_y` 镜像继承）

### OCR / MNN

- **本地手工放置的 MNN 预编译缓存会把 exe 静默链成动态依赖**（v2.6.10 本地事故）：ocr-rs 的 `build.rs` 只在「**全新下载解压后**」把 `MNN.lib` 替换为静态库 `MNN_static.lib`；一旦命中已存在的缓存目录就直接早退、跳过替换。若把官方预编译 zip 手工解压到 `~/.cargo/registry/src/*/ocr-rs-*/3rd_party/prebuilt/mnn-dev-windows-x86_64/lib/`，那里的 `MNN.lib`（438KB）是 **DLL 导入库**，于是 `cargo:rustc-link-lib=static=MNN` 实际产出依赖 `MNN.dll` 的 exe —— 装上启动即报「找不到 MNN.dll」（build.rs 的 `remove_dynamic_libs()` 还会把该目录里的 DLL 删掉，更找不到）。修法：解压后 `MNN.lib` → `MNN_import.lib`，再复制 `MNN_static.lib` → `MNN.lib`（即 build.rs 下载路径做的事），然后 **`cargo clean -p ocr-rs` 强制重链接**（cargo 不知道 lib 被换过，否则重建是空操作）。校验：OCR 版 exe 二进制内**不得**出现 `MNN.dll` 字符串（CI 产物为静态链接，无此依赖；轻量版不受影响）

### 其他

- **EXIF**：`image` crate 不自动应用；6=90°CW、8=90°CCW、3=180°
- **批量文字提取**：多 PDF 必须按 pdfPath 分组调 `extract_pdf_texts`；返回 `HashMap<u32, PdfTextResult>` keyed by pageIdx，前端按 `r._pdfPageIdx` 取结果
- **旋转方向**：全链路约定见「旋转与适配语义」小节——最易错点是 PDF 矩阵方向与 CSS 相反、pdf-lib 绕锚点旋转
- **ureq 的 TLS 后端必须显式注入**（v2.6.4 线上事故，issue #37①）：`features = ["native-tls"]` 只是让 native-tls 适配器可用，**不会**成为默认 TLS 后端。未启用 `tls`(rustls) feature 时 `default_tls_config()` 返回一个直接报错的桩，于是**全部 https 请求**都以 `cannot make HTTPS request because no TLS backend is configured` 失败（更新检查、PDFium / SumatraPDF 下载同时报废）。所有 http 请求一律经 `build_http_agent()` 建 agent，不要裸建 `ureq::AgentBuilder`
- **新建 lopdf 文档必须显式设经典 xref 表**（issue #46）：`Document::new()` / `with_version()` 的 `reference_table.cross_reference_type` 默认是 `CrossReferenceStream`（PDF 1.5 特性），而 `save_to()` 会照此写出——文件头声明 1.4 却在用 1.5 结构，只认经典 `xref` 表的老 RIP / 老打印机驱动解析不了整个文件（丢元素甚至空白页）。直通管道已强制 `lopdf::xref::XrefType::CrossReferenceTable`；printpdf 回退管道自身就写 1.3 + xref 表（printpdf 0.9 `serialize.rs`），两条管道需保持一致

## 硬性规则速查

改动前自查，违反即引入 bug：

1. 每页 slot 数一律 `getPerPage(s)`，禁止 `cols * rows`（所有分页计算点）
2. 旋转适配改动必须同步：预览 renderPage / 拖拽约束 / setSlotAlignment / Rust 矩阵 / web 分支，共五处
3. CPU 密集 IPC 命令必须 `async fn` + `spawn_blocking`
4. 顶层变量用 `var`；被 `loadSettings()` 恢复的变量声明必须在调用点之前
5. `defaultQuickLayouts()` 取深拷贝；空列表按 `Array.isArray` 恢复；不迁移旧默认布局
6. 占位（`_placeholder`）只占槽位：打印/统计/汇总/重命名/OCR 全部排除
7. 自动去重只删 `no:` key，`sum:` 仅标记
8. 类型/格式筛选切换后必须重算勾选集合（`selectFilteredOnly()` / `applyFilterSelection()`）
9. `deepEqual` 缓存比较排除纯打印机参数（printerName/copies/duplex/collate）
10. 桌面/web 双分支：识别（ocr.js ↔ js/pdf-text.js）与预览布局（layout.js ↔ js/layout.js）改动双向同步
11. 新增 http 请求一律走 `build_http_agent()`；裸建 `ureq::AgentBuilder` 会丢 TLS 后端（见「关键踩坑」）

## Git 工作流

- 开发在 `dev` 分支，完成后合并到 `master`；web 版单独 `web` 分支
- 小步提交，完成即 push；变动大时升版本打 tag 触发 CI
- 会话结束前确保无未提交变更

## 用户偏好

- 简洁直接，对 Bug 极度敏感，全面修复原则
- **Issue 回复只说人话**：结论 + 让用户做什么（下载哪个文件、看哪几处、回报什么），不写规范条款、推导过程与术语细节——太复杂没人看；技术细节留在代码注释、CHANGELOG 与本文件里
- 不要主动编译（耗时），等明确指令
- 分析任务绝对不可修改代码，必须先确认方案

## Release 检查清单

每次 release 前完成以下文档更新：

1. **README.md**：功能描述、技术栈版本与当前版本一致
2. **CHANGELOG.md**：新版本更新日志（新功能/修复/优化/依赖变更）
3. **AGENTS.md**：版本号、架构要点（如有变更）
4. **其他文档**：新增配置/命令/架构变更同步更新
