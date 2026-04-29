# 📄 电子发票批量打印工具

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)]()
[![Tauri 2.x](https://img.shields.io/badge/Tauri-2.x-orange.svg)]()

一个轻量级的桌面应用，专为批量打印电子发票设计。单文件 exe，无需安装，即开即用。

## ✨ 功能特性

### 📥 文件管理
- **多格式支持**：PDF、OFD、JPG、PNG、BMP、WebP、TIFF
- **智能渲染**：WinRT 原生渲染优先（支持中文系统字体），PDF.js 回退
- **金额识别**：自动识别发票金额（OCR + 文字提取），支持价税合计、合计、金额等关键字段
- **发票查验**：一键跳转国家税务总局发票查验平台
- **批量添加**：拖放文件或点击选择，一次添加多张发票
- **文件排序**：拖拽排序，调整打印顺序
- **单张设置**：双击发票可单独设置份数和旋转角度

### 📐 排版设置
- **纸张规格**：A4 / A5 / B5 / Letter / Legal / 自定义尺寸
- **版面布局**：9 个预设（1×1 到 3×3）+ 自定义行列（1-10行 × 1-10列），工具栏快捷切换
- **自动方向**：根据行列比自动选择横向/纵向
- **边距控制**：上下左右边距独立可调，预设 0/5/10mm 快捷按钮
- **间距调整**：列间距、行间距滑块微调
- **缩放模式**：自适应 / 拉伸填充 / 原始大小 / 自定义百分比
- **旋转控制**：全局旋转（0°/90°/180°/270°/自动适配）+ 单张旋转

### ✂ 辅助功能
- **裁切线**：多页拼版时显示虚线裁切标记
- **编号标记**：每个发票位标注序号
- **边框显示**：为每个发票位添加边框
- **裁剪白边**：自动检测并裁除发票周围的白边
- **水印**：自定义文字、透明度、颜色、角度、字号

### 🖨 打印与导出
- **打印模式**：弹出预览（在 PDF 阅读器中确认打印）或直接打印
- **份数设置**：全局份数 + 单张份数，支持逐份/逐页打印
- **双面打印**：选项切换
- **颜色模式**：彩色 / 灰度 / 黑白
- **页面顺序**：正向 / 反向
- **PDF 导出**：保存为 PDF，自动打开或自定义保存目录

### 🎨 界面
- **深色模式**：完整的深色主题支持
- **实时预览**：主区域实时预览打印效果，支持缩放和翻页
- **无感缩放**：Ctrl+滚轮缩放（聚焦指针处放大缩小），双击重置自适应
- **金额统计**：实时显示已选发票金额汇总
- **键盘快捷键**：`Ctrl+O` 添加 | `Ctrl+P` 打印 | `Ctrl++/-` 缩放 | `Ctrl+0` 自适应 | `←→` 翻页

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
| 前端 | 原生 HTML / CSS / JS | 单文件应用，无框架依赖 |
| PDF 渲染 | WinRT + PDF.js | WinRT 优先（支持系统字体），PDF.js 回退 |
| 后端 | Tauri 2.x (Rust) | 轻量桌面框架 |
| PDF 生成 | printpdf | Rust 原生 PDF 生成 |
| 打印 | winprint (Win32 Print Spooler) | 直接发送到打印机，可指定打印机 |
| 图像处理 | image (Rust) | 高性能图像解码 |

## 📦 项目结构

```
fapiao-print/
├── src/                        # 前端文件
│   ├── index.html              # 主页面（单文件应用）
│   ├── pdf.min.js              # PDF.js 本地副本
│   └── pdf.worker.min.js       # PDF.js Worker
├── src-tauri/                  # Tauri / Rust 后端
│   ├── src/
│   │   ├── main.rs             # 入口（隐藏控制台窗口）
│   │   ├── lib.rs              # 命令定义、拖放处理
│   │   └── pdf_engine.rs       # PDF 生成、WinRT 渲染、文件读取
│   ├── capabilities/
│   │   └── default.json        # Tauri 权限配置
│   ├── icons/                  # 应用图标
│   ├── Cargo.toml              # Rust 依赖
│   ├── build.rs                # Tauri 构建脚本
│   └── tauri.conf.json         # Tauri 配置
├── package.json                # Node.js 依赖（Tauri CLI）
├── LICENSE                     # MIT 许可证
└── .gitignore
```

## 🚀 开发

### 环境要求

- [Node.js](https://nodejs.org/) 18+
- [Rust](https://www.rust-lang.org/tools/install) 1.77+
- Windows 10/11（打印功能依赖 Windows Print Spooler API）

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

1. **添加发票**：点击「➕ 添加」按钮或直接拖放文件到窗口
2. **排版设置**：在左侧「⚙ 排版」面板调整纸张、布局、边距等
3. **预览检查**：主区域实时预览打印效果，支持缩放和翻页
4. **打印**：点击「🖨 打印」按钮，选择打印模式
5. **保存 PDF**：点击「📥 PDF」按钮导出 PDF 文件

### 打印模式

| 模式 | 行为 |
|------|------|
| 弹出打印对话框 | 生成 PDF 并用系统默认阅读器打开，用户确认后打印 |
| 直接打印 | 生成 PDF 并直接发送到指定打印机（或系统默认打印机），无需 PDF 阅读器 |

### 打印机选择

在左侧「🖨 打印」面板中可选择特定打印机，未选择时使用系统默认打印机。支持点击「🔄 刷新打印机列表」更新可用打印机。

### 保存目录

首次保存 PDF 时会弹出目录选择对话框，选择后自动记住，后续保存无需重复选择。可在「🔧 设置」面板中修改或清除。

## 📥 下载与运行

从 [Releases](../../releases) 下载最新版本：

| 文件 | 说明 |
|------|------|
| `发票打印工具_x64-setup.exe` | NSIS 安装包（推荐，自动处理 WebView2 安装） |
| `fapiao-print.exe` | 免安装绿色版（需系统已安装 WebView2） |

**运行依赖**：
- Windows 11：✅ 直接运行（自带 WebView2）
- Windows 10（较新版本）：✅ 大部分已预装 WebView2
- Windows 10（老版本）/ Windows Server：⚠️ 可能需要安装 [WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/)

## 🤖 关于此项目

本项目由 [WorkBuddy](https://www.codebuddy.cn/) AI 辅助生成，从零开始到可发布版本，历经 **60+ 轮** 调试迭代。主要攻克的技术难点包括：

- Tauri 2.x 文件对话框死锁问题（主线程同步调用导致）
- WebView2 拖放文件事件失效（`dataTransfer.files` 为空）
- Tauri 注入脚本与前端 `const` 变量冲突
- 缩放按钮在非均匀步进选项下的跳转逻辑
- Windows 子进程隐藏命令行窗口
- CSP 安全策略与 PDF.js CDN 回退的兼容
- WinRT `IBufferByteAccess` COM 接口查询失败（`E_NOINTERFACE`），改用 `DataReader` 读取渲染数据
- PDF.js CMap 配置，解决中文 CID 编码字体渲染问题
- CID 字体 PDF 金额提取失败（`join(' ')` → `join('')`，空格破坏正则匹配）
- WinRT OCR 金额自动识别（`Windows.Media.Ocr`，零额外依赖）
- 自适应 DPI 渲染（小页面自动提升 DPI，确保打印清晰度）
- OFD 格式解析（ZIP + XML + 图片提取）
- PDF 画质优化：300 DPI + PNG 无损输出 + 自适应渲染
- 多发票排版边距独立计算（per-slot margin）
- 打印机选择与直接打印（winprint + Win32 GetDefaultPrinterW，绕过 PDF 阅读器）

## 📄 许可证

[MIT License](LICENSE)
