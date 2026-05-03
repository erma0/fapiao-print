# 📄 电子发票批量打印工具

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)]()
[![Tauri 2.x](https://img.shields.io/badge/Tauri-2.x-orange.svg)]()

轻量桌面应用，专为批量打印电子发票设计。支持 PDF、OFD、图片等多格式导入，智能排版，一键打印或导出。

提供 **轻量版**（~3.5MB，纯打印）和 **OCR 版**（~24MB，含 PP-OCRv5 智能识别），单文件 exe 即开即用。

## ✨ 功能特性

### 🏆 OFD 完整支持

OFD（开放版式文档）是国家标准电子发票格式，本工具提供原生完整支持 — 矢量渲染、发票信息直提、印章保真，拖入即用，无需 OCR。

### 📥 文件管理

- **多格式支持**：PDF、OFD、JPG、PNG、BMP、WebP、TIFF
- **WinRT 原生 PDF 渲染**：`Windows.Data.Pdf`，支持中文系统字体，自适应 DPI（小页面自动提升至 1200）
- **PP-OCRv5 智能识别**（OCR 版）：文本优先 + 坐标回退双重架构，含税价 / 不含税价 / 税额数学验证配对，发票号码 / 日期 / 买卖方信息自动提取
- **发票查验**：一键跳转国家税务总局查验平台
- **批量操作**：拖放或点击选择，拖拽排序，双击单独设置份数 / 旋转

### 📐 排版设置

- **纸张**：A4 / A5 / B5 / Letter / Legal / 自定义
- **布局**：6 预设（1×1 / 2×1 / 3×2 / 1×2 / 2×2 / 3×3）+ 自定义行列（1-10 × 1-10），自动横纵方向
- **边距 / 间距**：独立可调，预设快捷按钮
- **缩放**：自适应 / 拉伸填充 / 原始大小 / 自定义百分比
- **旋转**：全局 0° / 90° / 180° / 270° / 自动 + 单张旋转

### ✂️ 辅助功能

- 裁切线、编号标记、边框显示、裁剪白边、自定义水印
- 金额统计、车票票种标签、发票类型自动检测

### 🖨️ 打印与导出

- **打印模式**：弹出预览（调用系统 PDF 阅读器）或静默直接打印（Print Spooler API，零窗口弹出）
- **份数控制**：全局 + 单张份数，逐份 / 逐页打印，双面打印，彩色 / 灰度 / 黑白
- **PDF 导出**：自动打开或自定义保存目录
- **确认弹窗**：打印前显示发票数量 / 版面 / 纸张 / 打印机 / 模式 / 份数，防止误操作

### 🎨 界面

- 深色 / 浅色模式、实时预览（缩放 + 翻页）
- **快捷键**：`Ctrl+O` 添加 · `Ctrl+P` 打印 · `Ctrl++/-` 缩放 · `Ctrl+0` 自适应 · `←→` 翻页

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

## 📦 下载

从 [Releases](../../releases) 下载最新版本：

| 文件 | 说明 |
|------|------|
| `发票打印工具_x64-setup.exe` | 轻量版安装包（~3.5MB） |
| `发票打印工具_x64_绿色版.exe` | 轻量版便携（单文件 exe，无需安装） |
| `发票打印工具_x64_OCR版-setup.exe` | OCR 版安装包（~24MB，含 PP-OCRv5） |
| `发票打印工具_x64_OCR绿色版.zip` | OCR 版便携（exe + models/） |

> 💡 只需排版打印选轻量版；需要自动识别金额 / 销售方信息选 OCR 版。

**运行依赖**：Windows 10 1803+ / Windows 11 均可直接运行（系统已预装或自动获取 WebView2）。Windows 7 需手动安装 [WebView2 Runtime v109](https://developer.microsoft.com/en-us/microsoft-edge/webview2/)，安装后应该可以运行，但微软已停止对 Win7/8 的 WebView2 更新支持。

## 📋 使用说明

1. **添加发票**：点击「➕ 添加」或拖放文件（支持 PDF / OFD / 图片混选）
2. **排版设置**：左侧「⚙ 排版」面板调整纸张、布局、边距
3. **预览检查**：主区域实时预览，支持缩放翻页；OCR 版可查看自动识别的金额信息
4. **打印**：点击「🖨 打印」，选择弹出预览或直接打印
5. **保存 PDF**：点击「📥 PDF」导出合成 PDF

## 🛠 技术栈

| 层级 | 技术 | 说明 |
|------|------|------|
| 前端 | 原生 HTML/CSS/JS | 模块化（app / ocr / layout / print），零依赖框架 |
| 后端 | Tauri 2.x (Rust) | 轻量桌面框架，Rust 条件编译管理功能开关 |
| PDF 渲染 | WinRT `Windows.Data.Pdf` | 原生渲染，自适应 DPI，支持中文系统字体 |
| PDF 生成 | printpdf 0.9 + lopdf 0.39 | JPEG 直通零质量损失、PDF 页面 Form XObject 全布局直通 |
| OFD 解析 | Rust 原生 XML 解析 | 矢量 SVG 输出 + 发票信息直提，FlateDecode 无损嵌入 PDF |
| OCR | ocr-rs 2.2 (PP-OCRv5 + MNN) | 文本优先 + 坐标回退，对比度增强，Lanczos3 锐化（OCR 版可选） |
| 打印 | Print Spooler API + ShellExecuteW (Win32) | 静默打印 / 对话框模式，自动获取默认打印机 |
| 图像处理 | image 0.25 (Rust) | 原生 WebP/TIFF 支持 |

## 📁 项目结构

```
fapiao-print/
├── src/                            # 前端
│   ├── index.html / styles.css
│   ├── app.js                      # 主入口、状态、文件加载
│   ├── ocr.js                      # OCR 提取（文本优先 + 坐标回退）
│   ├── layout.js                   # calculateLayout() + 预览渲染
│   └── print.js                    # 打印 / 导出 PDF
├── src-tauri/                      # Tauri / Rust 后端
│   ├── src/
│   │   ├── main.rs                 # 入口
│   │   ├── lib.rs                  # 命令、拖放、进程管理、OFD 解析
│   │   └── pdf_engine.rs           # PDF 生成（JPEG 直通 / FlateDecode / 全布局直通）、WinRT 渲染、OCR
│   ├── models/                     # PP-OCRv5 MNN 模型（OCR 版打包用）
│   ├── Cargo.toml                  # ocr feature flag + lopdf 0.39
│   ├── tauri.conf.json             # 轻量版配置
│   └── tauri.ocr.conf.json         # OCR 版配置（含 models）
├── scripts/
│   ├── build-all.js                # 一键全量构建（4 产物）
│   └── bump-version.js             # 版本号同步
└── package.json
```

## 🚀 开发

**环境要求**：Node.js 18+、Rust 1.77+、Windows 10/11

```bash
npm install

# 开发
npm run dev          # 轻量版
npm run dev:ocr      # OCR 版

# 构建
npm run build        # 轻量版
npm run build:ocr    # OCR 版
npm run build:all    # 一键全量构建（4 产物）

# 版本号
npm run bump 1.9.2   # 同步 package.json → Cargo.toml → tauri.conf.json
```

## 🗺 路线图

- [x] OFD 完整支持（矢量渲染 + 信息直提 + 印章 + 字体保真）
- [x] PDF 全布局直通（JPEG 零损失 + lopdf Form XObject）
- [x] Print Spooler API 静默打印
- [x] OCR Feature Flag 双版本构建
- [ ] 全电发票版式完善 + 通行费字段
- [ ] 发票去重检测（发票号码 + 开票日期）

## 🤖 关于此项目

本项目由 [WorkBuddy](https://www.codebuddy.cn/) AI 辅助生成，历经 60+ 轮迭代。主要攻克：Tauri 2.x 对话框死锁、WebView2 拖放失效、WinRT COM 接口适配、ocr-rs 条件编译集成、OFD 矢量渲染（DrawParam 继承链 / 文字排版 / 印章偏移）、PDF 引擎 JPEG 直通与 lopdf Form XObject 全布局直通、进程残留根治等。

## 📄 许可证

[MIT License](LICENSE)
