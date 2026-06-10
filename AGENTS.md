# 发票酱 — Agent 指南

> v2.2.0 纯前端（feat/web-cross-platform 分支）

## 项目概览

- **版本**: v2.2.0
- **本质**: 单页 HTML 应用，零后端、零构建
- **技术栈**: 原生 HTML/CSS/JS（无框架）、PDF.js + pdf-lib + JSZip
- **部署**: 静态文件托管，`index.html` 直接打开或 HTTP server

## 目录结构

```
fapiao/
├── index.html
├── css/styles.css
└── js/
    ├── app.js              # 主逻辑: 文件加载/状态管理/汇总表/CSV/设置
    ├── layout.js           # 排版计算 + 预览渲染
    ├── print.js            # PDF 合成 (pdf-lib) + iframe.print()
    ├── pdf-client.js       # PDF.js 封装 (ESM): 渲染+文字提取
    ├── ofd-client.js       # OFD 前端解析 (~960行)
    ├── xml-client.js       # XML 数电票解析
    ├── xml-utils.js        # 共享: _stripXmlNs/_parseXml
    ├── idb-store.js        # IndexedDB 文件存储 (ESM)
    └── vendor/
        ├── pdf.min.mjs / pdf.worker.min.mjs   # PDF.js 4.10
        ├── pdf-lib.min.js                     # pdf-lib 1.17
        ├── jszip.min.js                       # JSZip
        └── cmaps/                             # CJK bcmap
```

## 文件加载

`loadFileFromFile()` 根据扩展名分发:

| 类型 | 函数 | 解析器 |
|------|------|--------|
| `.pdf` | `loadPdfFromFile()` | PDF.js 渲染 + `applyPdfTextResult` 字段提取 |
| `.ofd` | `loadOfdFromFile()` | ofd-client.js → SVG → Canvas → PNG 300DPI |
| `.xml` | `loadXmlFromFile()` | xml-client.js DOMParser → `_xmlInvoice:true` |
| 图片 | `loadImageFromFile()` | FileReader → Image |

### 发票字段提取优先级
PDF文字层 > OFD XML > XML 数电票

### PDF 矢量嵌入
- pdf-lib `embedPage()` 直接嵌入原始 PDF 页面，保留矢量
- `srcPdfBytes`/`srcPageIndex`/`srcPageWidthPt`/`srcPageHeightPt` 记录源信息
- 失败 fallback 位图

### fileObj 结构 (createFileObj)
```
id, name, size, type, previewUrl
ow, oh, renderDpi                     # 原始尺寸
srcPdfBytes, srcPageIndex             # PDF 矢量嵌入源
srcPageWidthPt, srcPageHeightPt
slotScale, slotOffsetX, slotOffsetY   # 单票调整
copies, rotation, printed
// 发票字段
invoiceNo, invoiceDate, invoiceType
sellerName, sellerCreditCode
buyerName, buyerCreditCode
amountTax, amountNoTax, taxAmount
// 内部标记
_isTicket, _xmlInvoice, _pdfTextExtracted
_ofdSvg, _ofdPageWidth, _ofdPageHeight
```

## 打印流程

1. `print.js` 收集 `getActiveFiles()`（跳过 `_xmlInvoice`）
2. `_embedForFile()` pdf-lib 嵌入（embedPage 优先 → fallback 位图）
3. `_buildPage()` 按版面组装
4. iframe `contentWindow.print()` 浏览器 PDF 引擎

## 状态管理 (app.js S 对象)

```js
var S = {
  files: [], currentPage: 0, totalPages: 0,
  viewZoom: 0, editIdx: -1, selectedSlot: -1,
  amtMode: 'tax', printedFilter: 'all',
  layout: { cols, rows, orient },
  feat: { cutline, number, border, trimWhite, watermark,
          footer, pageNum, printDate, autoOpenPdf,
          pdfTextEnabled, customFM, slotAdjMemory, fileListMemory },
  // localStorage: ticketchan-settings / ticketchan-filelist
};
```

## 已知限制

- 打印走浏览器 PDF 引擎，无打印机控制（份数/双面/逐份无效）
- 依赖浏览器 `<iframe>.print()`，Safari/移动端可能有差异
- OFD 打印走位图路径（非矢量）
- XML 数电票不可打印/不可渲染（纯数据格式）
