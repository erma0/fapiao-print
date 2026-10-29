# 回归与真实样本验证

## 浏览器回归

```powershell
npm test
```

需要 Node.js 22+ 与 Chrome / Edge。脚本直接通过 Chromium 调试接口加载当前 `src/index.html`，使用独立临时配置，不依赖额外 npm 测试库。`TICKETCHAN_TEST_BROWSER` 可指定浏览器可执行文件。

默认覆盖购销方兜底、备注边界、份数变动后的选中位置、实际鼠标菜单与双击操作，以及更新渠道切换时响应/错误/缓存的并发处理。

## Rust PDF 提取验证

以下命令会编译测试程序，按需执行：

```powershell
$sampleVerificationDir = Join-Path $env:TEMP 'ticketchan-sample-verification'
New-Item -ItemType Directory -Path $sampleVerificationDir -Force | Out-Null
$env:TICKETCHAN_SAMPLE_DIR = (Resolve-Path -LiteralPath '.\sample').Path
$env:TICKETCHAN_SAMPLE_OUTPUT = Join-Path $sampleVerificationDir 'pdf-inputs.json'
Set-Location -LiteralPath '.\src-tauri'
cargo test --lib --no-default-features -- --nocapture
```

不设置样本变量时执行合成 PDF 回归用例；同时设置以上两个变量时，额外读取样本目录顶层的每份 PDF，通过生产批量提取函数输出全部页面。覆盖 CTM 拼接、旋转/缩放包围盒、CropBox、嵌套与重复 Form、字体作用域、资源继承及循环引用。

## 真实样本字段核对

样本原文、人工预期值与结果报告放在仓库外。Rust 输出每页的实际提取结果；图片可用 OCR 版桌面程序的 `ocr_image`，OFD/XML 用 `parse_ofd` / `parse_xml_invoice` 获取实际结果。合并后的输入数组格式为：

```json
[
  {"name":"example.pdf","page":0,"kind":"pdf","input":{"text":"","imgW":1000,"imgH":600,"lines":[],"hasTextLayer":true}},
  {"name":"example.jpg","page":0,"kind":"ocr","input":{"text":"","imgW":1000,"imgH":600,"lines":[]}},
  {"name":"example.xml","page":0,"kind":"xml","info":{"invoiceNo":"","buyerName":"","sellerName":"","amountTax":0,"remark":""}}
]
```

每个 `input.lines` 使用后端返回的完整词坐标，不替换为模拟文字；OFD 的 `info` 使用 `parse_ofd` 返回值中的 `invoiceInfo`，`kind` 为 `ofd`。

人工预期文件按实际票面核对，格式为 `[{"name":"example.pdf","page":0,"fields":{"invoiceNo":"…","buyerName":"…","sellerName":"…","amountTax":100,"invoiceRemark":"…","invoiceType":"普票"}}]`。同类文件名可能对应不同发票，跨格式比较必须先核对票号。备注为空的 XML 保持空值，不拼接独立银行字段。

在仓库根目录运行：

```powershell
$env:TICKETCHAN_SAMPLE_INPUT = Join-Path $sampleVerificationDir 'inputs.json'
$env:TICKETCHAN_SAMPLE_EXPECTED = Join-Path $sampleVerificationDir 'expected.json'
$env:TICKETCHAN_SAMPLE_REPORT = Join-Path $sampleVerificationDir 'results.json' # 可选
npm test
```

测试同时比较字段提取结果与 `applyPdfTextResult` / `applyOcrResult` 写入文件对象后的结果；票种经 `resolveInvoiceType` 核对。所有提供的样本均须有预期值，缺样本或字段不一致时失败。

2026-10-29 复核使用 24 个真实文件（14 PDF / 2 图片 / 4 OFD / 4 XML），共 25 页/结果，275 项字段检查通过；另有 21 项浏览器回归与 8 项 Rust 测试通过。
