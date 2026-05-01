# 📄 电子发票批量打印工具

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)]()
[![Tauri 2.x](https://img.shields.io/badge/Tauri-2.x-orange.svg)]()

轻量桌面应用，专为批量打印电子发票设计。提供**轻量版**（~3.5MB，无 OCR）和 **OCR 版**（~24MB，含 PP-OCRv5 智能识别），单文件 exe 即开即用。

> 🆕 v1.8.0：PDF 引擎三大优化 — JPEG 零质量损失直通、FlateDecode 无损压缩、PDF 全布局直通（lopdf Form XObject）

## ✨ 功能特性

### 📥 文件管理
- **多格式支持**：PDF、OFD、JPG、PNG、BMP、WebP、TIFF
- **WinRT 原生渲染**：`Windows.Data.Pdf`，支持中文系统字体
- **PP-OCRv5 智能识别**（OCR 版）：文本优先+坐标回退双重架构，含税价/不含税价/税额数学验证配对，车票专用提取，发票号码/日期/买卖方信息自动提取
- **发票查验**：一键跳转国家税务总局查验平台
- **批量添加**：拖放或点击选择，拖拽排序，双击单独设置份数/旋转

### 📐 排版设置
- **纸张**：A4 / A5 / B5 / Letter / Legal / 自定义
- **布局**：6 预设（1×1 / 2×1 / 3×2 / 1×2 / 2×2 / 3×3）+ 自定义行列（1-10 × 1-10），自动横纵方向
- **边距/间距**：独立可调，预设快捷按钮
- **缩放**：自适应 / 拉伸填充 / 原始大小 / 自定义百分比
- **旋转**：全局 0°/90°/180°/270°/自动 + 单张旋转

### ✂ 辅助功能
- 裁切线、编号标记、边框显示、裁剪白边、自定义水印

### 🖨 打印与导出
- **打印模式**：弹出预览 或 直接打印到指定打印机
- **份数**：全局+单张，逐份/逐页打印，双面打印，彩色/灰度/黑白
- **PDF 导出**：自动打开或自定义保存目录

### 🎨 界面
- 深色模式、实时预览（缩放+翻页）、金额统计、车票票种标签
- **快捷键**：`Ctrl+O` 添加 | `Ctrl+P` 打印 | `Ctrl++/-` 缩放 | `Ctrl+0` 自适应 | `←→` 翻页

## 📸 界面预览

<table>
  <tr>
    <td align="center">☀️ 浅色模式</td>
    <td align="center">🌙 深色模式</td>
  </tr>
  <tr>
    <td><img src="screenshots/light.png" alt="浅色模式" width="480"/></td>
    <td><img src="screenshots/dark.png" alt="深色模式" width="480"/></td>
  </tr>
</table>

## 🛠 技术栈

| 层级 | 技术 | 说明 |
|------|------|------|
| 前端 | 原生 HTML/CSS/JS | 模块化（app/ocr/layout/print），无框架 |
| PDF 渲染 | WinRT `Windows.Data.Pdf` | 原生渲染，支持中文系统字体 |
| OCR | ocr-rs 2.2 (PP-OCRv5 + MNN) | 文本优先+坐标回退（OCR 版可选） |
| 后端 | Tauri 2.x (Rust) | 轻量桌面框架 |
| PDF 生成 | printpdf 0.9 + lopdf 0.39 | JPEG 直通零质量损失、PDF 页面 Form XObject 直通 |
| 打印 | ShellExecuteW (Win32) | 对话框/直接打印 |
| 图像处理 | image 0.25 (Rust) | 原生 WebP 支持 |

## 📦 项目结构

```
fapiao-print/
├── src/                        # 前端
│   ├── index.html / styles.css
│   ├── app.js                  # 主入口、状态、文件加载
│   ├── ocr.js                  # OCR 提取（文本优先+坐标回退）
│   ├── layout.js               # calculateLayout() + 预览渲染
│   └── print.js                # 打印/导出 PDF
├── src-tauri/                  # Tauri / Rust 后端
│   ├── src/
│   │   ├── main.rs             # 入口
│   │   ├── lib.rs              # 命令、拖放、进程管理
│   │   └── pdf_engine.rs       # PDF 生成（JPEG 直通/FlateDecode/PDF 全布局直通）、WinRT 渲染、OCR
│   ├── models/                 # PP-OCRv5 MNN 模型（OCR 版打包用）
│   ├── Cargo.toml              # ocr-rs 为 optional 依赖，lopdf 0.39
│   ├── tauri.conf.json         # 轻量版配置（无 models）
│   └── tauri.ocr.conf.json     # OCR 版配置（含 models）
├── scripts/
│   └── build-all.js            # 一键全量构建
└── package.json
```

## 🚀 开发

**环境**：Node.js 18+、Rust 1.77+、Windows 10/11

```bash
npm install

# 开发
npm run dev          # 轻量版
npm run dev:ocr      # OCR 版

# 构建
npm run build        # 轻量版
npm run build:ocr    # OCR 版
npm run build:all    # 一键全量构建（4 产物）
```

## 📥 下载

从 [Releases](../../releases) 下载最新版本：

| 文件 | 说明 |
|------|------|
| `发票打印工具_x64-setup.exe` | 轻量版安装包（~3.5MB，无 OCR） |
| `发票打印工具_x64_绿色版.exe` | 轻量版便携（单文件 exe，无需安装） |
| `发票打印工具_x64_OCR版-setup.exe` | OCR 版安装包（~24MB，含 PP-OCRv5） |
| `发票打印工具_x64_OCR绿色版.zip` | OCR 版便携（exe + models/） |

> 💡 只需打印选轻量版；需要自动识别金额/销售方信息选 OCR 版。

**运行依赖**：Windows 11 直接运行；Windows 10 较新版本大部分已预装 WebView2；老版本可能需安装 [WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/)

## 📋 使用说明

1. **添加发票**：点击「➕ 添加」或拖放文件
2. **排版设置**：左侧「⚙ 排版」面板调整纸张、布局、边距
3. **预览检查**：主区域实时预览，支持缩放翻页
4. **打印**：点击「🖨 打印」，选择弹出预览或直接打印
5. **保存 PDF**：点击「📥 PDF」导出

## 🗺 路线图

- [ ] 重写打印流程：直接调用 Win32 Print Spooler API，绕过 PDF 阅读器
- [ ] 完善 OCR：全电发票版式、通行费字段、识别缓存、准确率优化
- [ ] 发票去重检测（发票号码+开票日期）
- [ ] 批量打印进度反馈
- [ ] 国际化支持

## 🤖 关于此项目

本项目由 [WorkBuddy](https://www.codebuddy.cn/) AI 辅助生成，历经 60+ 轮迭代。主要攻克：Tauri 2.x 对话框死锁、WebView2 拖放失效、WinRT COM 接口适配、ocr-rs 集成与条件编译、文本优先+坐标回退 OCR 架构、金额三值数学验证、零 IPC 往返优化、进程残留根治、PDF 引擎 JPEG 直通与 lopdf Form XObject 全布局直通等。

## 📄 许可证

[MIT License](LICENSE)
