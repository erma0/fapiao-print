# 📄 电子发票批量打印工具

一个基于 Tauri 2.x 的桌面应用，专为批量打印电子发票设计。支持 PDF / JPG / PNG 等常见发票格式，提供排版、拼版、批量打印等功能。

## ✨ 功能特性

- **多格式支持**：PDF、JPG、PNG、BMP、WebP、TIFF
- **批量添加**：拖放文件或点击选择，一次添加多张发票
- **排版设置**：纸张规格、方向、边距、间距全面可调
- **拼版打印**：1×1 / 1×2 / 2×2 布局，一张纸打印多张发票
- **缩放模式**：自适应 / 拉伸填充 / 原始大小 / 自定义百分比
- **辅助选项**：裁切线、编号、边框、裁剪白边
- **水印功能**：自定义文字、透明度、颜色、角度、字号
- **打印设置**：份数、逐份/逐页、双面打印、颜色模式
- **PDF 导出**：保存为 PDF，支持自动打开和自定义保存目录
- **深色模式**：完整的深色主题支持
- **键盘快捷键**：Ctrl+O 添加、Ctrl+P 打印、方向键翻页

## 🛠 技术栈

| 层级 | 技术 |
|------|------|
| 前端 | 原生 HTML / CSS / JS（无框架） |
| PDF 渲染 | PDF.js（本地 + CDN 回退） |
| 后端 | Tauri 2.x (Rust) |
| PDF 生成 | printpdf |
| 打印 | Windows ShellExecuteW |
| 图像处理 | image (Rust) |

## 📦 项目结构

```
fapiao-print/
├── src/                    # 前端文件
│   ├── index.html          # 主页面（单文件应用）
│   ├── pdf.min.js          # PDF.js 本地副本
│   └── pdf.worker.min.js   # PDF.js Worker
├── src-tauri/              # Tauri / Rust 后端
│   ├── src/
│   │   ├── lib.rs          # 应用入口、命令定义、拖放处理
│   │   └── pdf_engine.rs   # PDF 生成、文件读取、打印机列表
│   ├── capabilities/
│   │   └── default.json    # Tauri 权限配置
│   ├── icons/              # 应用图标
│   ├── Cargo.toml          # Rust 依赖
│   ├── build.rs            # Tauri 构建脚本
│   └── tauri.conf.json     # Tauri 配置
├── package.json            # Node.js 依赖（Tauri CLI）
└── .gitignore
```

## 🚀 开发

### 环境要求

- [Node.js](https://nodejs.org/) 18+
- [Rust](https://www.rust-lang.org/tools/install) 1.77+
- Windows 10/11（打印功能依赖 Windows Shell API）

### 安装依赖

```bash
npm install
```

### 开发模式

```bash
npm run dev
```

### 构建发布

```bash
npm run build
```

构建产物位于 `src-tauri/target/release/fapiao-print.exe`。

## 📋 使用说明

1. **添加发票**：点击「➕ 添加」按钮或拖放文件到窗口
2. **排版设置**：在左侧「⚙ 排版」面板调整纸张、布局、边距等
3. **预览检查**：主区域实时预览打印效果，支持缩放和翻页
4. **打印**：点击「🖨 打印」按钮，选择打印模式
5. **保存 PDF**：点击「📥 PDF」按钮导出 PDF 文件

### 打印模式

| 模式 | 行为 |
|------|------|
| 弹出打印对话框 | 生成 PDF 并打开预览，用户在阅读器中确认打印 |
| 直接打印 | 生成 PDF 并直接发送到默认打印机 |

### 保存目录

首次保存 PDF 时会弹出目录选择对话框，选择后自动记住目录，后续保存直接使用，无需重复选择。可在「🔧 设置」面板中修改或清除。

## 📄 许可证

MIT License
