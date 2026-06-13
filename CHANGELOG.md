# 📋 更新日志

> 本分支（`web`）独立发布，与主分支（`master`，Tauri 桌面版）并行维护。
> 下文仅记录本分支的纯前端 Web 版变更历史。

## v3.1.0 — 功能补齐 + 可靠性修复

> 补齐 v3.0.0 遗漏的功能说明，修复多项 PDF 打印可靠性问题。

### ✨ 新功能

- **保存图片 PDF 按钮**：非矢量模式，跳过 pdf-lib 字体嵌入，无需下载 CJK 字体即可输出带水印/页脚的 PDF（内容为位图）

### 🛠️ 修复

- **水印/序号 PDF 位置对齐预览**：水印改为 per-slot 居中、序号改为 slot 右上角黑底白字，与预览完全一致
- **CJK 字体嵌入增强**：新增 TTC (TrueType Collection) 字体解包支持，系统本地字体为 TTC 格式时自动提取第一个字体
- **CJK subset 失败兜底**：不再只处理 CFF/OTF 格式异常，任何 subset 失败均无条件重试无 subset 嵌入
- **模块异步加载竞态修复**：新增 `waitForGlobal()` 轮询等待机制，消除 `__pdfClient`/`__ofdClient` 等异步模块未就绪的时序问题
- **批量加载进度优化**：并发加载每张发票，逐个 toast 更新实时进度而非定时轮询
- **vendor 路径修复**：pdf-client.js 中 CMap/worker 路径改为正确的 `../vendor/` 相对路径
- **PDF 缓存键补全**：新增 `number`/`watermarkColor`/`watermarkAngle`/`customFM` 字段到缓存键，消除开关切换时返回过期 PDF 的问题

### 📝 文档

- **README**: 补充打印状态追踪、日期排序、排版份数、调整记忆、设置持久化、主题切换等功能描述
- **README**: 版本号更新至 v3.1.0，桌面版引用更新至 v2.0.8
- **README**: 新增「与桌面版的差异」章节，明确纯 Web 版因浏览器限制不可用的功能
- **README**: 移除不准确的「文件记忆」描述（纯前端无法恢复本地路径）

---

## v3.0.0 — 纯前端（零后端）

> 完成从 Rust 版到纯前端版的最终迁移。所有功能在浏览器内自给自足，无需任何服务器。

### 🗑️ 删除

- **Rust 后端全部代码**: `src-tauri/` 目录、`Cargo.toml`、`Cargo.lock`、所有 Rust 依赖
- **invoice-engine OFD 解析器**: Rust 版 OFD 引擎（已完全移植到 ofd-client.js）
- **OCR 引擎**: MNN / PP-OCRv5 模型文件全部删除
- **Docker / Nginx / systemd**: 部署配置文件全部删除（不再需要服务器）

### 🧹 偏好设置清理

- **删除「保存后自动打开」**(autoOpenPdf): 死开关
- **删除「记忆发票列表」**(fileListMemory): 纯前端无法恢复本地文件路径
- **删除「发票查验」板块**: 仅跳转外部链接无实际功能

### 🔄 前端解析器（完全移植自 Rust）

- **ofd-client.js**: OFD ZIP→XML→SVG→PNG 300DPI，对齐 Rust invoice-engine 全部逻辑
- **xml-client.js**: XML 数电票 DOMParser 解析
- **xml-utils.js**: 共享 XML 工具，处理命名空间前缀

### ✅ 功能完整性

所有核心功能与 Rust v2.0.7 版对齐：
- PDF/OFD/XML/Image 文件加载与渲染
- 发票字段提取（PDF 文字层 > OFD XML > XML 数电票）
- PDF 矢量直通合成（pdf-lib embedPage，保留印章/签章）
- 汇总表（14 字段可勾选、CSV UTF-8 BOM）
- 单票调整（slotScale/slotOffset、九宫格对齐、滚轮缩放）
- 全部辅助功能（裁切线/编号/边框/水印/页码/页脚）
- IndexedDB 文件缓存
- 设置持久化（localStorage）

---

## 关联项目

- **主分支 `master`**: 独立维护的 [Tauri 桌面版 v2.0.8](https://github.com/erma0/fapiao-print/tree/master)，使用 Rust 后端 + PDFium 引擎，支持静默打印、打印机选择、批量重命名等桌面级功能。需要本地安装 Windows 客户端的用户请使用该版本。
