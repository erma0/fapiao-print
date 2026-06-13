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

async function loadDocument(arrayBuffer, workerIdx) {
  var data = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
  var task = pdfjsLib.getDocument({
    data: data,
    cMapUrl: CMAP_URL,
    cMapPacked: true,
    standardFontDataUrl: CMAP_URL.replace(/cmaps\/$/, ''),
    isEvalSupported: false,
    workerSrc: _workerSrcUrl
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

  for (var idx = 0; idx < tc.items.length; idx++) {
    var item = tc.items[idx];
    if (!item.str || !item.transform) continue;

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

  var hasTextLayer = tc.items.length > 0 && tc.items.some(function(item) { return item.str && item.str.length > 0; });

  return {
    text: fullTextParts.join('\n'),
    lines: lines,
    imgW: imgW,
    imgH: imgH,
    hasTextLayer: hasTextLayer
  };
}

export async function loadPdfFromArrayBuffer(arrayBuffer) {
  var pdf = await loadDocument(arrayBuffer);
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
