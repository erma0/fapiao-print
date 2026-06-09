// =====================================================
// PDF.js 4.x Client Wrapper
// =====================================================
// Pure-frontend replacement for the v2.1.0 server-side
// render_pdf_pages_pdfium + extract_pdf_texts endpoints.
// Loads ArrayBuffer → renders pages to JPEG dataURL → extracts
// text content with word-level coordinates for extractByCoordinates().

import * as pdfjsLib from './vendor/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./vendor/pdf.worker.min.mjs', import.meta.url).toString();

const CMAP_URL = new URL('./vendor/cmaps/', import.meta.url).toString();
const PDF_RENDER_DPI = 300;
const PDF_PREVIEW_DPI = 150;
const JPEG_QUALITY = 0.82;

async function loadDocument(arrayBuffer) {
  var data = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
  var task = pdfjsLib.getDocument({
    data: data,
    cMapUrl: CMAP_URL,
    cMapPacked: true,
    standardFontDataUrl: CMAP_URL.replace(/cmaps\/$/, ''),
    isEvalSupported: false
  });
  return await task.promise;
}

async function renderPageToCanvas(pdfPage, dpi) {
  var vp1 = pdfPage.getViewport({ scale: 1.0 });
  var longestSide = Math.max(vp1.width, vp1.height);
  var targetDpi = Math.max(dpi, Math.ceil((dpi * 0.5) / longestSide * 72));
  targetDpi = Math.min(targetDpi, 600);
  var scale = targetDpi / 72;
  var viewport = pdfPage.getViewport({ scale: scale });
  var canvas = document.createElement('canvas');
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);
  var ctx = canvas.getContext('2d', { alpha: false });
  await pdfPage.render({ canvasContext: ctx, viewport: viewport }).promise;
  return { canvas: canvas, width: canvas.width, height: canvas.height, dpi: targetDpi };
}

/**
 * Extract text with word-level coordinates from a PDF page via PDF.js.
 * Produces a PdfTextResult-compatible structure that applyPdfTextResult()
 * in ocr.js can consume — same shape as the Rust extract_pdf_text output.
 *
 * PDF.js getTextContent() items have:
 *   str: text string
 *   transform: [scaleX, skewX, skewY, scaleY, translateX, translateY]
 *   width: width of the text string in PDF points
 *   height: font size (approximate)
 *
 * Coordinate conversion:
 *   PDF.js transform[4,5] are in PDF pt space (origin bottom-left, y-up).
 *   We convert to the same pixel coordinate system as Rust extract_pdf_text:
 *   scale = RENDER_DPI / 72
 *   x = transform[4] * scale
 *   y = (pageHeight - transform[5] - fontSize) * scale  (flip y, baseline→top)
 *   w = width * scale (or approximate from character count)
 *   h = fontSize * scale
 */
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

export async function getPageBytesForLayout(pdf, pageIndex) {
  return null;
}

window.__pdfClient = { loadPdfFromArrayBuffer: loadPdfFromArrayBuffer };
