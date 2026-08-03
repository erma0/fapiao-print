// =====================================================
// PDF.js 4.x Client Wrapper
// =====================================================
// Pure-frontend replacement for the v2.1.0 server-side
// render_pdf_pages_pdfium + extract_pdf_texts endpoints.
// Loads ArrayBuffer → renders pages to JPEG dataURL → extracts
// text content with word-level coordinates for extractByCoordinates().
//
// Multi-worker concurrency: spawns multiple PDF.js worker instances
// based on navigator.hardwareConcurrency for parallel document parsing.

import * as pdfjsLib from '../vendor/pdf.min.mjs';

const CMAP_URL = new URL('../vendor/cmaps/', import.meta.url).toString();
const PDF_RENDER_DPI = 300;
const PDF_PREVIEW_DPI = 300;
const MIN_RENDER_PX = 3508;  // A4 long side at 300 DPI — minimum rendered pixels
const MAX_RENDER_DPI = 600;
const JPEG_QUALITY = 0.82;

// Auto-detect concurrency: half the logical cores, min 1, max 4
const MAX_CONCURRENT = Math.min(4, Math.max(1, Math.floor((navigator.hardwareConcurrency || 4) / 2)));

// Pre-create worker source URLs for parallel parsing
var _workerSrcUrl = new URL('../vendor/pdf.worker.min.mjs', import.meta.url).toString();
pdfjsLib.GlobalWorkerOptions.workerSrc = _workerSrcUrl;

async function loadDocument(arrayBuffer, opts) {
  opts = opts || {};
  var data = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
  var task = pdfjsLib.getDocument({
    data: data,
    cMapUrl: CMAP_URL,
    cMapPacked: true,
    standardFontDataUrl: CMAP_URL.replace(/cmaps\/$/, ''),
    isEvalSupported: false,
    workerSrc: _workerSrcUrl,
    // 暴露 font.missingFile / toUnicode / cidEncoding，用于检测 12306 等字体未嵌入的 PDF
    fontExtraProperties: opts.fontExtraProperties === true
  });
  return await task.promise;
}

async function renderPageToCanvas(pdfPage, dpi) {
  var vp1 = pdfPage.getViewport({ scale: 1.0 });
  var longestSide = Math.max(vp1.width, vp1.height);
  var minDpiFromPx = Math.ceil(MIN_RENDER_PX / (longestSide / 72));
  var targetDpi = Math.max(dpi, minDpiFromPx);
  targetDpi = Math.min(targetDpi, MAX_RENDER_DPI);
  var scale = targetDpi / 72;
  var viewport = pdfPage.getViewport({ scale: scale });
  var canvas = document.createElement('canvas');
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);
  var ctx = canvas.getContext('2d', { alpha: false });
  await pdfPage.render({ canvasContext: ctx, viewport: viewport }).promise;
  return { canvas: canvas, width: canvas.width, height: canvas.height, dpi: targetDpi };
}

async function extractTextWithCoords(pdfPage) {
  var tc = await pdfPage.getTextContent();
  var vp = pdfPage.getViewport({ scale: 1.0 });
  var pageW = vp.width;
  var pageH = vp.height;
  var scale = PDF_RENDER_DPI / 72;
  var imgW = Math.round(pageW * scale);
  var imgH = Math.round(pageH * scale);

  var LINE_Y_THRESHOLD = 2;
  var lines = [];
  var currentWords = [];
  var lastY = null;
  var fullTextParts = [];

  // ---- 字体未嵌入检测层 1：文本信号 ----
  // items 存在但大量空串 → ToUnicode 缺失的强信号（如 12306 铁路电子客票）
  var totalItems = tc.items.length;
  var emptyItems = 0;
  var distinctFonts = {};
  for (var ei = 0; ei < tc.items.length; ei++) {
    var _it = tc.items[ei];
    if (!_it.str || _it.str.length === 0) emptyItems++;
    if (_it.fontName) distinctFonts[_it.fontName] = (distinctFonts[_it.fontName] || 0) + 1;
  }
  var emptyRatio = totalItems > 0 ? emptyItems / totalItems : 0;

  for (var idx = 0; idx < tc.items.length; idx++) {
    var item = tc.items[idx];
    if (!item.str || !item.transform) continue;
    // 跳过纯空白/控制字符项，避免污染坐标提取
    if (!item.str.replace(/\s/g, '').length) continue;

    var tx = item.transform[4];
    var ty = item.transform[5];
    var fontSize = Math.abs(item.transform[3]) || Math.abs(item.transform[0]) || 10;
    if (fontSize < 1) fontSize = 10;

    var textW = item.width > 0 ? item.width : fontSize * item.str.length * 0.5;

    var px = tx * scale;
    var py = (pageH - ty - fontSize) * scale;
    var pw = textW * scale;
    var ph = fontSize * scale;

    var word = {
      text: item.str,
      x: px,
      y: Math.max(0, py),
      w: pw,
      h: ph
    };

    if (lastY === null || Math.abs(ty - lastY) > LINE_Y_THRESHOLD) {
      if (currentWords.length > 0) {
        lines.push({ words: currentWords, confidence: 1.0 });
        fullTextParts.push(currentWords.map(function(w) { return w.text; }).join(''));
      }
      currentWords = [];
    }

    currentWords.push(word);
    lastY = ty;
  }

  if (currentWords.length > 0) {
    lines.push({ words: currentWords, confidence: 1.0 });
    fullTextParts.push(currentWords.map(function(w) { return w.text; }).join(''));
  }

  // 阈值收紧：空串占比 >= 50% 即判定为文字层缺失
  var hasTextLayer = totalItems > 0 && emptyRatio < 0.5;

  // ---- 字体未嵌入检测层 2：字体元信息（best-effort，pdf.js 私有 API，失败时静默）----
  var fontDiag = null;
  if (!hasTextLayer && totalItems > 0) {
    fontDiag = _inspectPageFonts(pdfPage, distinctFonts);
  }

  return {
    text: fullTextParts.join('\n'),
    lines: lines,
    imgW: imgW,
    imgH: imgH,
    hasTextLayer: hasTextLayer,
    // 文字层缺失诊断（下游可选使用，不破坏现有调用）
    textLayerDiagnostic: {
      totalItems: totalItems,
      emptyItems: emptyItems,
      emptyRatio: emptyRatio,
      distinctFonts: Object.keys(distinctFonts),
      fontInspect: fontDiag,
      suspectMissingToUnicode: !hasTextLayer && totalItems > 0
    }
  };
}

/**
 * 通过 pdfPage.commonObjs 检查页面字体嵌入状态。
 * 注意：commonObjs 在 getTextContent/render 之后才被填充，调用时机必须在那些 API 之后。
 * 字段名基于 pdf.js 4.x，未来版本可能变化，全部 try/catch 保护。
 */
function _inspectPageFonts(pdfPage, distinctFonts) {
  var result = { missingFonts: [], noToUnicode: [], cidEncodings: [], sampledNames: [] };
  try {
    var commonObjs = pdfPage.commonObjs;
    if (!commonObjs) return result;
    var fontIds = Object.keys(distinctFonts);
    for (var fi = 0; fi < fontIds.length; fi++) {
      var fid = fontIds[fi];
      var fontObj = null;
      try { fontObj = commonObjs.get(fid); } catch (e) { fontObj = null; }
      if (!fontObj) continue;
      if (fontObj.name) result.sampledNames.push(fontObj.name);
      // missingFile=true → 字体未嵌入（如 12306 的 SimSun）
      if (fontObj.missingFile) result.missingFonts.push(fontObj.name || fid);
      // toUnicode=null/undefined → 没有 ToUnicode CMap，pdf.js 无法将 CID 映射回 Unicode
      if (!fontObj.toUnicode) result.noToUnicode.push(fontObj.name || fid);
      // cidEncoding（如 "GBK-EUC-H" / "Identity-H"）
      if (fontObj.cidEncoding) result.cidEncodings.push(fontObj.cidEncoding);
    }
  } catch (e) {
    console.warn('[pdf-client] 字体检测异常:', e);
  }
  return result;
}

export async function loadPdfFromArrayBuffer(arrayBuffer) {
  // fontExtraProperties:true 让 commonObjs 暴露 font.missingFile / toUnicode / cidEncoding
  var pdf = await loadDocument(arrayBuffer, { fontExtraProperties: true });
  var pages = [];
  for (var i = 1; i <= pdf.numPages; i++) {
    var page = await pdf.getPage(i);
    var rendered = await renderPageToCanvas(page, PDF_PREVIEW_DPI);
    var previewUrl = rendered.canvas.toDataURL('image/jpeg', JPEG_QUALITY);
    var pdfTextResult = await extractTextWithCoords(page);
    var vp = page.getViewport({ scale: 1.0 });
    pages.push({
      pageIndex: i,
      previewUrl: previewUrl,
      width: rendered.width,
      height: rendered.height,
      pdfWidthPt: vp.width,
      pdfHeightPt: vp.height,
      renderDpi: rendered.dpi,
      pdfTextResult: pdfTextResult
    });
    page.cleanup();
  }
  pdf.destroy();
  return { numPages: pages.length, pages: pages };
}

// Semaphore-based concurrent loader: limits how many PDFs are parsed simultaneously
var _activeCount = 0;
var _waitQueue = [];

function _acquire() {
  if (_activeCount < MAX_CONCURRENT) {
    _activeCount++;
    return Promise.resolve();
  }
  return new Promise(function(resolve) { _waitQueue.push(resolve); });
}

function _release() {
  _activeCount--;
  if (_waitQueue.length > 0) {
    _activeCount++;
    var next = _waitQueue.shift();
    next();
  }
}

export async function loadPdfConcurrent(arrayBuffer) {
  await _acquire();
  try {
    return await loadPdfFromArrayBuffer(arrayBuffer);
  } finally {
    _release();
  }
}

export async function getPageBytesForLayout(pdf, pageIndex) {
  return null;
}

window.__pdfClient = {
  loadPdfFromArrayBuffer: loadPdfFromArrayBuffer,
  loadPdfConcurrent: loadPdfConcurrent,
  maxConcurrent: MAX_CONCURRENT
};
