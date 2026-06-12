# 📋 更新日志

> 本分支（`web`）独立发布，与主分支（`master`，Tauri 桌面版）并行维护。
> 下文仅记录本分支的纯前端 Web 版变更历史。

## v3.1.0 — PNG 导出

- 🖼 **保存 PNG**: 替换打印按钮，canvas 渲染版面到 300 DPI PNG 图片
  - 支持全部页面批量导出、当前页单独导出
  - 保持所有排版设置（缩放、旋转、偏移、裁切线、水印、页脚）
  - 快捷键 Ctrl+S 保存全部页面

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

### 🛠️ 修复 & 优化

- `savePdf is not defined` 拼写错误 → 补上 `_buildPage` 缺失闭合
- `detached ArrayBuffer` → `buffer.slice(0)` 提前复制
- 水印/页脚/序号坐标 → 统一 PDF 坐标系（72dpi）而非屏幕（96dpi）
- 字段名 `wmText` → `watermarkText` 修正
- CJK 字体按需加载：① 本地系统字体（Local Font Access API）→ ② IndexedDB 缓存 → ③ CDN 下载
- 纯英文水印/页脚不触发字体下载
- OTF/CFF 字体嵌入失败时自动降级为无 subset 嵌入

---

## 关联项目

- **主分支 `master`**: 独立维护的 [Tauri 桌面版 v2.0.7](https://github.com/erma0/fapiao-print/tree/master)，使用 Rust 后端 + PDFium 引擎，支持静默打印、打印机选择、虚拟打印机等桌面级功能。需要本地安装 Windows 客户端的用户请使用该版本。
