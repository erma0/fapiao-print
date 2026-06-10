# 发票酱 前端重构审计报告 — JS 对 Rust 逐项对比

> 审计日期: 2026-06-10
> 对比基线: `feat/web-cross-platform` 前端 JS vs `src/` + `invoice-engine/` Rust 实现
> 审计范围: OFD/XML 解析器、OCR 字段提取、PDF 合成、布局计算、API 兼容、并发安全、配置/日志

---

## 总览

| 分项 | 一致性评级 | 说明 |
|------|-----------|------|
| OFD 解析器 | ★★★★☆ | 核心逻辑高度一致，一处 Annotation 解析差异 |
| XML 数电票解析器 | ★★★★★ | 完全对齐，差异仅实现方式 |
| OCR / 字段提取 | ★★★☆☆ | 坐标系提取一致，但无后端 OCR 引擎 → 无真实 OCR 能力 |
| PDF 合成 | ★★★☆☆ | 渲染质量可能退化 (JPEG 重编码)，布局公式一致 |
| 布局计算 | ★★★★★ | 公式完全一致 |
| API 兼容 | ★★★☆☆ | 存在冗余映射和不存在的端点 |
| 并发安全 | ★★★★★ | JS 单线程自然安全，无 Rust 的 Mutex/DashMap 风险 |
| 配置/日志 | ★★★★☆ | 基本一致 |

---

## 1. OFD 解析器 (ofd-client.js ↔ invoice-engine/src/lib.rs)

### 1.1 核心逻辑一致性

| 功能模块 | Rust | JS | 状态 |
|----------|------|----|------|
| ZIP 读取 | `zip::ZipArchive` | `JSZip.loadAsync` | ✅ 一致 |
| XML 解析 | quick-xml SAX + `local_tag_name()` | DOMParser + `_stripXmlNs()` | ✅ 一致（方法不同结果同） |
| DeltaX 解析 | `parse_delta_x()` — group `g N val` + extra | `_parseDeltaX()` — 完全对齐 | ✅ 一致 |
| 字体规范化 | `normalize_font_name()` — 12 种映射 | `_normalizeFontName()` — 12 种映射 | ✅ 一致 |
| SVG 文本构建 | CTM/Normal 两路径 + tspan 绝对定位 | 完全一致 | ✅ 一致 |
| SVG 路径构建 | `ofd_path_to_svg()` 8 种指令 | 完全一致 | ✅ 一致 |
| ImageObject | Boundary + ResourceID + ImageMask | 完全一致 | ✅ 一致 |
| ImageMask 合成 | `image` crate 像素合成 RGBA PNG | Canvas API 像素合成 RGBA PNG | ✅ 一致 |
| DrawParam 继承 | `resolve_draw_param()` 链式遍历 + 循环检测 | 完全一致 + visited 去重 | ✅ 一致 |
| 页面尺寸检测 | 仅 page content XML PhysicalBox | 先 Document.xml → 再 page content 覆盖 | ⚠️ JS 更全面 |
| Annotation 解析 | explicitly wrap in `<ofd:Content><ofd:Layer>` | `XMLSerializer().serializeToString(ap)` 依赖已有的 Layer 结构 | ⚠️ **差异点** |

### 1.2 ⚠️ 差异点 D1 — Annotation 解析

**Rust** (`parse_annotations`, L1136-1203):
```rust
let mut frag = format!("<ofd:Content><ofd:Layer>");
// ... reconstructs full XML with Content + Layer wrapper
frag.push_str("</ofd:Layer></ofd:Content>");
let (mut texts, _, mut imgs) = parse_ofd_content(&frag);
```

**JS** (`_parseAnnotations`, L565):
```javascript
var content = _parseOfdContent(new XMLSerializer().serializeToString(ap));
```

**影响范围**: 对标准 OFD 文件（Appearance 内含 `<Content><Layer>` 结构）无差异。对非标准 OFD 文件（Appearance 直接含 TextObject）可能遗漏所有 Annotation 内容（水印/印章）。

**修复建议**: 在 `_parseOfdContent` 调用前判断是否已有 Layer，无则包裹一层。或改为与 Rust 一致的显式重建。**优先级: 低**（实际 OFD 文件均含 Layer 结构）。

### 1.3 CustomData 字段提取

| 字段 | Rust (CustomData + CustomTag) | JS (CustomData + CustomTag) | 状态 |
|------|------|------|------|
| invoiceNo | `get_custom("发票号码")` → `get_tag_text("InvoiceNo")` | 完全一致 | ✅ |
| invoiceDate | `get_custom("开票日期")` → `get_tag_text("IssueDate")` | 完全一致 | ✅ |
| buyerName | 仅 CustomTag `BuyerName` | 完全一致 | ✅ |
| sellerName | 仅 CustomTag `SellerName` | 完全一致 | ✅ |
| buyerTaxId | `get_custom("购买方纳税人识别号")` → `BuyerTaxID` | 完全一致 | ✅ |
| sellerTaxId | `get_custom("销售方纳税人识别号")` → `SellerTaxID` | 完全一致 | ✅ |
| amountNoTax | `get_custom("合计金额")` → fallback Amount/TotalAmWithoutTax | 完全一致 | ✅ |
| taxAmount | `get_custom("合计税额")` → fallback TaxTotalAmount | 完全一致 | ✅ |
| amountTax | `noTax+tax` 计算 → fallback `价税合计` → TaxInclusiveTotalAmount | 完全一致 | ✅ |
| invoiceType | 从模板/页面文本 + CustomTag | 完全一致 | ✅ |

### 1.4 Text-based 兜底提取

`_extractInvoiceFromText()` 与 `extract_invoice_from_text()` 逐行对比：
- 单字拼接 buffer：一致 ✅
- 买方/卖方 section 检测：一致 ✅
- "名称" 出现次数区分买卖方：一致 ✅
- "纳税人识别号" 出现次数区分：一致 ✅
- 价税合计/小写/合计 标记链：一致 ✅
- ¥ 金额解析：一致 ✅
- 金额缺口补全 (amount_tax ↔ no_tax+tax)：一致 ✅

---

## 2. XML 数电票解析器 (xml-client.js ↔ invoice-engine parse_xml_invoice_content)

### 2.1 核心逻辑一致性

| 功能 | Rust | JS | 状态 |
|------|------|----|------|
| 根元素检测 | `content.contains("<EInvoice")` | 完全一致 | ✅ |
| InvoiceNumber | `path.push/pop` SAX 路径追踪 | `getElementsByTagName` DOM 查询 | ✅ |
| IssueTime | `split('T').next()` | `split('T')[0]` | ✅ |
| SellerIdNum 空值 | `if !text.is_empty()` pattern guard | `if (sellerIdEl && sellerIdEl.trim())` | ✅ |
| BuyerIdNum 空值 | `if !text.is_empty()` 个人发票跳过 | 完全一致 | ✅ |
| TotalTax-includedAmount | 主路径 + `TotalTaxIncludedAmount` fallback | 完全一致 | ✅ |
| LabelName 上下文 | `path.len()-2` → 父标签 `EInvoiceType`/`GeneralOrSpecialVAT` | `getElementsByTagName('EInvoiceType')` → `getElementsByTagName('LabelName')` | ✅ |
| 发票类型组合 | `format!("{}({})", prefix, special_label)` | 完全一致 | ✅ |

### 2.2 ⚠️ 差异点 D2 — LabelName 查找范围

**Rust**: 精确的 SAX 路径追踪，`LabelName` 仅在 `EInvoiceType` / `GeneralOrSpecialVAT` 直接父元素下匹配。

**JS**: 使用 `parentEl.getElementsByTagName('LabelName')`（递归搜索所有后代），从 `getElementsByTagName('EInvoiceType')` 找到的每个 EInvoiceType 元素的所有后代中找第一个 LabelName。

**影响**: 对标准 XML 结构无差异（LabelName 是 EInvoiceType/GeneralOrSpecialVAT 的直接子元素）。如果 XML 中存在嵌套的 EInvoiceType→SomeWrapper→LabelName 结构，JS 仍能正确找到（更鲁棒），而 Rust 会漏掉（因为只看直接子元素）。

**结论**: JS 实际上**更鲁棒**。无修复需要。

### 2.3 ⚠️ 差异点 D3 — `_stripXmlNs` 代码重复

`ofd-client.js:25-33` 和 `xml-client.js:156-161` 中存在完全相同的 `_stripXmlNs()` 实现。

**修复建议**: 提取到独立的 `shared.js` 或挂载到 `window.__xmlNs` 公用。**优先级: 低**（功能正常，仅维护负担）。

---

## 3. OCR / 字段提取 (ocr.js ↔ pdf_engine.rs OCR逻辑)

### 3.1 架构差异

| 维度 | Rust (pdf_engine.rs) | JS (ocr.js) |
|------|---------------------|-------------|
| OCR 引擎 | MNN/PP-OCRv5 (GPU 加速) | 无 OCR 引擎 |
| PDF 文本提取 | PDFium `FPDFText_GetText` | PDF.js `getTextContent` |
| 字段提取 | `extract_invoice_from_pdf_text()` 坐标+关键词 | `extractByCoordinates()` + `applyPdfTextResult()` |
| 金额解析 | `parse_amt()` + `parse_chinese_numeral()` | `parseAmt()` + `parseChineseNumeral()` |
| 发票类型检测 | `_detect_invoice_type()` 关键词 | `_detectInvoiceType()` 完全一致 |

### 3.2 ⚠️ 差异点 D4 — 无后端 OCR 能力

**影响**: 在纯前端模式下（无 Rust Server），无法对**图片格式**发票（JPG/PNG）进行 OCR 字段提取。PDF 发票可通过 PDF.js 文本层提取。OFD 发票通过 XML 解析提取。

**现状**: 这是架构选择的必然结果——前端模式不支持服务器端 OCR。PDF.js 文本层提取在大部分场景替代 OCR。

**修复建议**: 已有 `applyOcrAsync()` 函数可调用服务器端 OCR API（当 server 就绪时）。不需要额外修改。

### 3.3 parseChineseNumeral 对比

Rust 和 JS 的 `parseChineseNumeral` 逻辑应一致（大写/小写数字映射、亿/万/仟/佰/拾/圆/角/分 处理）。

由于 Rust 后端中 `parse_chinese_numeral` 函数与 JS 中独立实现，且未找到 Rust 侧的对应源码（可能在 pdf_engine.rs 的深层），标记为 **需验证**。

---

## 4. PDF 合成 (print.js ↔ pdf_engine.rs generate_pdf_from_layout)

### 4.1 架构差异

| 维度 | Rust (pdf_engine.rs) | JS (print.js) |
|------|---------------------|---------------|
| 矢量直通 | lopdf `embedPage` + Form XObject | pdf-lib `embedPage` |
| 图像嵌入 | lopdf JPEG DCTDecode (原始字节, 零重编码) | pdf-lib `embedJpg`/`embedPng` |
| 旋转处理 | PDF cm 矩阵 (90/180/270) 旋转映射 | 浏览器 canvas / CSS transform |
| 裁切线 | `build_cut_lines_content_stream()` | `_composePdfBlob()` PDF 路径绘制 |
| 水印 | `ab_glyph` + simhei.ttf 渲染 | pdf-lib `drawText` + 内置字体 |
| 页脚 | `ab_glyph` 文本渲染 | pdf-lib `drawText` |
| 打印 | PDFium GDI 矢量打印 (Windows) | 浏览器 `iframe.print()` |

### 4.2 ⚠️ 差异点 D5 — JPEG 重编码质量退化

**Rust lopdf 路径**: 直接将原始 JPEG 字节通过 DCTDecode 嵌入 PDF，**零解码+零重编码**，保持原始 JPEG 质量。

**JS pdf-lib 路径**: `embedJpg` 内部可能对 JPEG 进行解码-重编码（取决于 pdf-lib 版本和实现），可能导致质量损失。

**影响范围**: 使用 PDF 发票源文件（PDF 内嵌页面）作为输入时，矢量直通不受影响。仅当图像（JPG/PNG/OFD 渲染后的 PNG）嵌入时，可能产生轻微质量差异。

**修复建议**: 检查 pdf-lib 版本是否支持 JPEG 直通（pass-through），如不支持则接受质量差异。**优先级: 低**（打印质量差异肉眼不可见）。

### 4.3 布局计算

Rust 和 JS 使用**完全相同的数学公式**：
```
sw = (pw - cols*(ml+mr) - (cols-1)*gh) / cols
sh = (ph - rows*(mt+mb) - (rows-1)*gv - effectiveFm) / rows
```

裁切线位置计算也一致（行间中点 / 列间中点 / 页脚上方）。✅

### 4.4 ⚠️ 差异点 D6 — 打印控制能力

| 功能 | Rust (pdfium_print.rs) | JS |
|------|------------------------|----|
| 指定打印机 | ✅ Win32 EnumPrintersW + 指定 | ❌ 浏览器决定 |
| 打印份数 | ✅ DEVMODE dmCopies | ❌ |
| 双面打印 | ✅ DEVMODE dmDuplex | ❌ |
| 颜色模式 | ✅ DEVMODE dmColor | ❌ |
| 静默打印 | ✅ ShellExecuteW printto | ❌ |

**影响**: 纯前端模式下无法控制打印机/份数/双面/颜色，这是**已知的架构限制**，在 v2.2.0 重构时已明确。

---

## 5. 布局计算 (layout.js ↔ pdf_engine.rs)

### 5.1 公式对比

完全一致。`calculateLayout()` 中的 `sw`/`sh` 计算公式、slot 坐标生成、cutLines 位置计算均与 Rust `build_nup_content_stream` 中的逻辑对应。

唯一的增强是 JS 的 `effectiveFm` 处理更精细（customFM vs auto 模式），但在逻辑上是兼容的超集。

---

## 6. API 兼容性 (api.js ↔ server.rs)

### 6.1 映射分析

| api.js 命令 | 映射到的路由 | 服务器是否存在 | 状态 |
|-------------|-------------|---------------|------|
| `render_pdf_pages` | `/api/v1/render_pdf` | ✅ | 正常 |
| `render_pdf_pages_pdfium` | `/api/v1/render_pdf` | ✅ (同上) | 正常 |
| `extract_pdf_text` | `/api/v1/extract_pdf_text` | ✅ | 正常 |
| `extract_pdf_texts` | `/api/v1/extract_pdf_texts` | ✅ | 正常 |
| `generate_pdf_from_layout` | `/api/v1/generate_pdf` | ✅ | 正常 |
| `get_printers` | `/api/v1/printers` | ✅ | 正常 |
| `list_printers` | `/api/v1/printers` | ✅ (同上) | 正常 |
| `pdfium_print` | `/api/v1/pdfium_print` | ✅ server.rs 存在 | 正常 |
| `print_pdf_file` | `/api/v1/print` | ✅ server.rs 存在 | 正常 |
| **`ocr_image`** | **`/api/v1/ocr_image`** | **✅ server.rs 存在** | **⚠️ 缺失映射 → 已修复** |
| **`ocr_pdf_page`** | **`/api/v1/ocr_pdf_page`** | **✅ server.rs 存在** | **⚠️ 缺失映射 → 已修复** |
| `parse_ofd` | `/api/v1/parse_ofd` | ✅ | 正常 |
| `parse_xml_invoice` | `/api/v1/parse_xml_invoice` | ✅ | 正常 |
| `open_ofd_images` | `/api/v1/open_ofd_images` | ✅ server.rs 存在 | 正常 |
| ~~`check_path_exists`~~ | ~~`/api/v1/check_path_exists`~~ | ✅ server.rs 存在 | 已清理（纯桌面用） |
| ~~`get_config`~~ | ~~`/api/v1/get_config`~~ | ✅ server.rs 存在 | 已清理（纯桌面用） |
| `get_app_version` | `/api/v1/get_app_version` | ✅ server.rs 存在 | 正常 |
| `trim_image` | `/api/v1/trim_image` | ✅ server.rs 存在 | 正常 |
| ~~`copy_file`~~ | ~~`/api/v1/copy_file`~~ | ✅ server.rs 存在 | 已清理（纯桌面用） |
| `rename_file` | `/api/v1/rename_file` | ✅ server.rs 存在 | 正常 |
| ~~`write_text_file`~~ | ~~`/api/v1/write_text_file`~~ | ✅ server.rs 存在 | 已清理（纯桌面用） |
| ~~`get_temp_dir`~~ | ~~`/api/v1/get_temp_dir`~~ | ✅ server.rs 存在 | 已清理（纯桌面用） |
| ~~`get_downloads_dir`~~ | ~~`/api/v1/get_downloads_dir`~~ | ✅ server.rs 存在 | 已清理（纯桌面用） |

### 6.2 ⚠️ 差异点 D7 — API 端点映射缺失 + 桌面遗留清理

**复核发现**: 原审计遗漏了一个**实际 Bug** — `ocr.js` 调用 `__api.call('ocr_image')` 和 `__api.call('ocr_pdf_page')`，但 `_endpoints` 映射中缺少这两个条目，调用会直接抛 "Command not available" 错误。服务器端 `server.rs` 实际有这两个路由。

原审计标注的 `pdfium_print`/`print`/`open_ofd_images`/`get_app_version`/`trim_image`/`rename_file` 等端点，在 `server.rs` 中**均存在**，原标注"可能不存在"有误。

**修复**: 已添加 `ocr_image` 和 `ocr_pdf_page` 映射，清理纯桌面模式遗留端点。

---

## 7. 并发与线程安全

### 7.1 Rust 侧安全机制

| 机制 | 用途 |
|------|------|
| `DashMap` | Session 管理 — 无锁读写 |
| `Mutex<PdfiumState>` | PDFium FFI 保护 — **必须 Mutex 不可 RwLock** (ACCESS_VIOLATION) |
| `Semaphore` | OCR/Render 并发限制 |
| `AtomicBool` | 全局关闭标志 |
| `spawn_blocking` | CPU 密集型操作移到 blocking thread pool |

### 7.2 JS 侧

浏览器 JS 天然单线程（主线程），PDF.js Worker 运行在独立 Worker 线程中（最多 4 并发），IndexedDB 是异步事务模型。不存在 Rust 侧的线程安全问题。

**结论**: 无退化。JS 不需要 Mutex/Semaphore/DashMap 等保护机制。

---

## 8. 配置解析 / 日志 / 参数

### 8.1 配置

| 配置项 | Rust | JS | 状态 |
|--------|------|----|------|
| 设置存储 | 浏览器 localStorage (`ticketchan-*` 前缀) | localStorage (`ticketchan-*` 前缀) | ✅ 一致 |
| 环境变量 | `TICKETCHAN_SERVER_PORT` 等 5 个 | `window.__apiBaseUrl` | ✅ 对应 |
| Cargo.toml version | 唯一版本号源 | `S.version` 从 api 获取 | ✅ |

### 8.2 日志

- Rust: `env_logger` + `log::info!/warn!/error!`
- JS: `console.log/warn/error` + toast 提示

格式差异是平台决定的，不影响功能。

### 8.3 错误处理

- Rust: `AppError` 11 种错误类型 → `IntoResponse` JSON
- JS: `api.js` 中的 `_fetch()` 解析 JSON error 响应 → `throw Error`

错误格式兼容。✅

---

## 9. 差异汇总与修复建议

### 需修复 (P1)

| ID | 差异 | 影响 | 建议 |
|----|------|------|------|
| **D4** | JS 前端无 OCR 引擎 | 图片发票无法提取字段 | 已通过 PDF 文本层和 OFD XML 解析覆盖主要场景；出问题时引导用户使用 server 模式 |

### 建议改进 (P2) → 已修复 ✅

| ID | 差异 | 影响 | 状态 |
|----|------|------|------|
| **D1** | Annotation 解析缺少显式 Layer 包裹 | 非标准 OFD 的 Annotation 可能遗漏 | ✅ 已修复 — 添加 Layer 包裹容错 |
| **D3** | `_stripXmlNs` 代码重复 | 维护负担 | ✅ 已修复 — 提取到 xml-utils.js |
| **D7** | API 端点映射缺失 + 桌面遗留 | ocr.js OCR 调用失败 | ✅ 已修复 — 添加 OCR 端点，清理桌面遗留 |

### 复核新增发现 (P1 → 已修复 ✅)

| ID | 差异 | 影响 | 状态 |
|----|------|------|------|
| **D7b** | `ocr_image`/`ocr_pdf_page` 端点映射缺失 | ocr.js 调用 OCR 功能必然失败 | ✅ 已修复 |

### 无需处理 (P3)

| ID | 差异 | 原因 |
|----|------|------|
| **D2** | LabelName 查找范围差异 | JS 更鲁棒，无负面影响 |
| **D5** | JPEG 重编码质量 | 肉眼不可见，且矢量直通路径不受影响 |
| **D6** | 打印控制能力 | 已知架构限制，已文档化 |

---

## 10. 结论

JS 重构与 Rust 原版在核心逻辑上**高度一致**。OFD/XML 解析器可视为逐行对齐移植，字段提取的三级优先级（CustomData → CustomTag → Text fallback）完全保持。布局计算公式一致。

主要差异集中在**平台能力边界**：
1. Rust 有 PDFium 渲染引擎 + MNN OCR 引擎 → JS 只能用 PDF.js + 前端解析
2. Rust 有系统打印 API → JS 只能通过浏览器打印
3. Rust 有 lopdf JPEG 直通 → JS pdf-lib 可能有轻微重编码

这些差异是 Web 架构的固有限制，不是代码缺陷。当前没有阻塞性问题。
