# PDF 解析引擎重构尝试记录

**分支**: `archive/pdf_oxide-refactor-attempt`
**创建时间**: 2025年
**状态**: 已归档（不推荐合并到主分支）
**结论**: 性能退化、代码量增加、匹配率无改善，不建议采用

---

## 背景

项目原有的 `pdf_engine.rs` 使用 lopdf 进行 PDF 文本提取，但存在以下问题：
- lopdf 的 CJK 编码支持不完整，需要大量补丁代码（约 976 行）
- 火车票等使用 GBK-EUC-H 编码的 PDF 无法直接解码
- 代码维护困难，需要手工维护 CMap 解析器和 GBK/Big5 回退逻辑

本重构尝试使用 `pdf_oxide` 库替代 lopdf 进行文本提取，以期：
- 减少/消除 CJK 补丁代码
- 提升火车票等 GBK 编码 PDF 的解析质量
- 改善代码可维护性

---

## 重构内容

### 阶段一：文本提取替换

**变更**:
- 删除 `pdf_engine.rs` 中约 976 行的 lopdf CJK 补丁代码
- 新增 `pdf_engine_text_extract.rs`，使用 `pdf_oxide::PdfDocument::extract_page_text()` 提取文本
- 移除 `encoding_rs` 和 `flate2` 依赖（pdf_oxide 自带）

**依赖变更**:
```toml
[删除]
encoding_rs = "0.8"
flate2 = "1.1.9"

[新增]
pdf_oxide = "0.3"

[变更]
rust-version = "1.77.2" → "1.88"
```

### 阶段二：PDF 生成回退替换

**变更**:
- 移除 `printpdf` 依赖
- 使用 `pdf_oxide::PdfDocumentBuilder` 替代 printpdf 生成 PDF

### 阶段三：文本格式优化

**变更**:
- 修复 pdf_oxide 输出的 CJK 字符被拆分为单字的问题
- 添加按 Y 坐标分行、合并相邻词的处理逻辑

### 阶段四：发票解析问题修复

由于 pdf_oxide 文本提取质量不如预期，添加了大量前端补丁：

**后端修复** (`pdf_engine_text_extract.rs`):
- `is_garbled_text()`: 检测乱码文本，触发 OCR 回退
- `try_split_fused_span()`: 拆分跨越页面中心的融合 span
- `decode_gbk_span()`: GBK 字节对解码（对韩文区域字符有效）
- `extract_gbk_text_lopdf()`: lopdf 内容流 GBK 解码（火车票完整回退）

**前端修复** (`ocr.js`):
- `_normTextForExtract`: 多点数字修复（`3.99.00` → `399.00`）
- `_cleanName`: 剥离页码/日期/标签后缀
- `_splitFusedCompanyNames()`: 融合公司名称拆分
- `_extractNamesByCoords`: 融合词预拆分逻辑
- 融合信用代码拆分（3处）
- 非税发票金额提取修复

---

## 性能对比（实测数据）

**测试环境**: Windows, release 模式

### 普通增值税发票（1页）

| 方案 | 平均耗时 |
|------|---------|
| lopdf 单独 | **1.5ms** |
| pdf_oxide 单独 | 2.8ms |
| pdf_oxide + lopdf（当前方案） | 2.9ms |

### 火车票 PDF（1页，GBK 编码）

| 方案 | 平均耗时 |
|------|---------|
| lopdf 单独 | **1.2ms** |
| pdf_oxide 单独 | 1.8ms |
| pdf_oxide + lopdf（当前方案） | 3.0ms |

**结论**: pdf_oxide 比 lopdf 慢 1.5~2 倍。对于 GBK 编码 PDF，当前双重加载方案比纯 lopdf 慢 2.5 倍。

---

## 代码量对比

| 文件 | 原版 | 重构后 | 变化 |
|------|------|--------|------|
| `pdf_engine.rs` | 4,194 行 | 3,477 行 | -717 行 |
| `pdf_engine_text_extract.rs` | 0 | 528 行 | +528 行 |
| `ocr.js` | 2,746 行 | 3,107 行 | +361 行 |
| Rust 合计 | 4,194 行 | 4,005 行 | **-189 行** |
| **总代码量** | 6,940 行 | 7,112 行 | **+172 行** |

后端减少了 189 行，但前端增加了 361 行补丁，净增 172 行。

---

## 匹配率对比

| 发票类型 | 原版 lopdf | pdf_oxide 方案 | 变化 |
|---------|-----------|---------------|------|
| 普通增值税发票 | ✅ | ✅ | 持平 |
| 增值税专用发票 | ✅ | ⚠️ 多余小数点 | **变差** |
| 打车发票 | ⚠️ 名称融合 | ⚠️ 名称融合（前端拆分） | 持平 |
| 非税发票 | ⚠️ 卖方名称错误 | ⚠️ 卖方名称错误 | 持平 |
| 火车票 | ✅ GBK 直接解码 | ✅ lopdf 回退解码 | **持平** |
| 页码前缀问题 | ⚠️ | ✅ 前端已修复 | 改善 |

**关键发现**: 原版 lopdf 已经内置了 GBK-EUC-H 和 Big5 解码支持（通过 `encoding_rs`），火车票可以正确解析。本重构的 lopdf 回退路径实际上绕过了原版已有的功能。

---

## 问题分析

### 1. pdf_oxide 文本提取质量不如 lopdf

pdf_oxide 的文本提取存在以下问题：
- **CJK 字符拆分**: 部分字体将每个汉字拆分为单独的 span
- **GBK 编码误处理**: 将 GBK 字节对当作 Unicode 码点输出（产生韩文乱码）
- **CID 编码问题**: 对 Identity-H 编码字体的处理不完整（数字/标点字符未正确映射）

### 2. 双重加载导致性能退化

当前方案对 GBK 编码 PDF 的处理流程：
1. `pdf_oxide::PdfDocument::open()` — 加载一次
2. `doc.extract_page_text()` — 提取文本
3. `calc_garbled_ratio()` + `test_gbk_decode()` — 乱码检测
4. `decode_gbk_span()` × N — GBK 解码每个 span
5. `lopdf::Document::load()` — **再加载一次**
6. 内容流解析 + GBK 解码 — **再提取一次**

相当于对同一个 PDF 加载了两次，提取了两次。

### 3. lopdf 无法完全移除

重构目标是"不再依赖 lopdf"，但实际上 lopdf 仍需保留：
- N-up 合并功能（Form XObject 嵌入）仍使用 lopdf
- GBK 内容流回退解析需要 lopdf

### 4. 依赖变更

pdf_oxide 引入了 37 个子依赖，包括 `image`、`brotli`、`qcms`、`subsetter` 等，其中很多与项目已有依赖重复但版本不同。

---

## 结论

本次重构尝试**未达到预期目标**：

1. **性能退化**: pdf_oxide 比 lopdf 慢 1.5~2 倍，GBK PDF 慢 2.5 倍
2. **代码量增加**: 净增 172 行（前端补丁增加）
3. **匹配率无改善**: 原版 lopdf 已支持 GBK 解码，本重构只是绕过了已有功能
4. **依赖更重**: pdf_oxide 引入 37 个子依赖，MSRV 提升到 1.88
5. **lopdf 无法移除**: N-up 合并和 GBK 回退仍需要 lopdf

---

## 经验教训

1. **lopdf 的 CJK 支持比预期好**: 原版 lopdf 已经内置了 GBK/Big5 解码支持，通过 `encoding_rs` 实现
2. **pdf_oxide 的文本提取质量不如预期**: 对于复杂的发票 PDF，lopdf 的直接内容流解析反而更可靠
3. **文本提取库的选择**: lopdf 虽然 API 底层，但控制力强；pdf_oxide 封装度高但灵活性差
4. **回退策略**: 如果要使用 pdf_oxide，应该以它为主，遇到问题再回退 lopdf，而不是反过来

---

## 后续建议

如果未来需要改进 PDF 文本提取，建议：

1. **保留原版 lopdf 方案**: 当前方案稳定可靠，无需更换
2. **如果火车票解析有问题**: 检查原版 lopdf 的 GBK 解码路径是否被绕过
3. **如果要换库**: 考虑 `pdf-rs`（轻量、纯 Rust）或直接使用系统工具（如 `pdftotext`）
4. **如果要提升解析质量**: 重点应该在 OCR 回退策略和前端正则匹配优化上，而不是底层库
