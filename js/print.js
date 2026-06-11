// =====================================================
// Print & PDF Functions — pure client-side (pdf-lib + iframe.print)
// =====================================================
// No backend. All PDF generation runs in the browser via pdf-lib.
// Print uses an embedded <iframe> to invoke the browser's native PDF print.
//
// PDF sources use vector embedding (embedPage) to preserve text/line quality.
// Image sources (JPG/PNG/OFD) fall back to raster embedding.

var _printCacheKey = null;
var _printCacheBlob = null;

// Source PDF document cache: avoids re-loading the same ArrayBuffer for
// multi-page PDFs. Cleared after each compose call.
var _srcPdfDocs = {};

// pdf-lib UMD exports as `PDFLib` (uppercase). Use a window alias so all
// internal `pdfLib.X` references resolve consistently.
if (typeof window !== 'undefined' && !window.pdfLib) {
  window.pdfLib = window.PDFLib || {};
}

function _loadSettings() { try { return JSON.parse(localStorage.getItem('ticketchan-settings') || '{}'); } catch (e) { return {}; } }
function _settingsKey(s) {
  var k = {
    layout: s.layout, fitMode: s.fitMode, customScale: s.customScale,
    paperW: s.paperW, paperH: s.paperH, paperSize: s.paperSize, orientation: s.orientation,
    marginTop: s.marginTop, marginBottom: s.marginBottom, marginLeft: s.marginLeft, marginRight: s.marginRight,
    gapH: s.gapH, gapV: s.gapV, colorMode: s.colorMode, cutline: s.cutline, border: s.border,
    watermark: s.watermark, wmText: s.wmText, wmSize: s.wmSize, wmOpacity: s.wmOpacity,
    globalRotation: s.globalRotation, footerText: s.footerText, pageNum: s.pageNum,
    printDate: s.printDate, footerMargin: s.footerMargin
  };
  return JSON.stringify(k);
}

function _pngBlobFromDataUrl(dataUrl) {
  var bin = atob(dataUrl.split(',')[1]);
  var buf = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
  return buf;
}

function _jpegBlobFromDataUrl(dataUrl) {
  return _pngBlobFromDataUrl(dataUrl);
}

// Load a source PDF with pdf-lib (cached per file ID prefix)
async function _getOrLoadSrcPdf(fileObj) {
  var key = fileObj.id.replace(/_p\d+$/, '');
  if (_srcPdfDocs[key]) return _srcPdfDocs[key];
  // Copy the buffer — pdf-lib may transfer/detach the original ArrayBuffer,
  // causing subsequent reads to fail with "detached ArrayBuffer".
  var raw = fileObj.srcPdfBytes;
  var bytes = raw instanceof Uint8Array
    ? new Uint8Array(raw)
    : new Uint8Array(raw.slice(0));
  var doc = await pdfLib.PDFDocument.load(bytes, { ignoreEncryption: true });
  _srcPdfDocs[key] = doc;
  return doc;
}

// Embed a file into the output PDF.
// For PDF sources: returns { type:'pdfPage', embedded, width(pt), height(pt) }
// For image sources: returns { type:'image', embedded, width(px), height(px) }
// On failure or unsupported type: returns null.
async function _embedForFile(pdfDoc, fileObj) {
  if (!fileObj) return null;
  if (fileObj._xmlInvoice) return null;

  // Vector path: embed original PDF page
  if (fileObj.srcPdfBytes && fileObj.srcPageIndex != null) {
    try {
      var srcDoc = await _getOrLoadSrcPdf(fileObj);
      var srcPage = srcDoc.getPage(fileObj.srcPageIndex);
      var embedded = await pdfDoc.embedPage(srcPage);
      return {
        type: 'pdfPage',
        embedded: embedded,
        width: srcPage.getWidth(),
        height: srcPage.getHeight()
      };
    } catch (e) {
      console.warn('embedPage failed, falling back to raster:', e);
    }
  }

  // Raster path: embed as image
  if (fileObj.previewUrl) {
    var isJpeg = fileObj.previewUrl.indexOf('data:image/jpeg') === 0
      || fileObj.previewUrl.indexOf('data:image/jpg') === 0;
    var bytes = isJpeg ? _jpegBlobFromDataUrl(fileObj.previewUrl) : _pngBlobFromDataUrl(fileObj.previewUrl);
    var img = isJpeg ? await pdfDoc.embedJpg(bytes) : await pdfDoc.embedPng(bytes);
    return { type: 'image', embedded: img, width: img.width, height: img.height };
  }
  return null;
}

async function _buildPage(pdfDoc, pageFiles, pageIdx, settings) {
  var layout = calculateLayout(settings);
  var pw = layout.pw;
  var ph = layout.ph;
  var page = pdfDoc.addPage([pw, ph]);
  var bgColor = settings.colorMode === 'bw' ? [1, 1, 1] : [1, 1, 1];
  page.drawRectangle({ x: 0, y: 0, width: pw, height: ph, color: pdfLib.rgb(bgColor[0], bgColor[1], bgColor[2]) });

  for (var i = 0; i < layout.slots.length; i++) {
    var slot = layout.slots[i];
    var f = pageFiles ? pageFiles[i] : null;
    if (!f) continue;
    var embedResult = await _embedForFile(pdfDoc, f);
    if (!embedResult) continue;

    var rot = getRotation(f, slot, settings);
    var perScale = f.slotScale || 1;
    var perOffX = f.slotOffsetX || 0;
    var perOffY = f.slotOffsetY || 0;

    // Use pt dimensions for PDF pages, pixel dimensions for images.
    // Both give correct aspect ratio for fit calculation.
    var objW, objH;
    if (embedResult.type === 'pdfPage') {
      objW = f.srcPageWidthPt || embedResult.width;
      objH = f.srcPageHeightPt || embedResult.height;
    } else {
      objW = f.ow || embedResult.width;
      objH = f.oh || embedResult.height;
    }

    var fitScale = Math.min(slot.w / objW, slot.h / objH);
    if (settings.fitMode === 'fill') fitScale = Math.max(slot.w / objW, slot.h / objH);
    else if (settings.fitMode === 'original') fitScale = 1;
    if (settings.fitMode === 'custom' && settings.customScale) fitScale *= settings.customScale;
    fitScale *= perScale;

    var drawW = objW * fitScale;
    var drawH = objH * fitScale;
    var cx = slot.x + slot.w / 2 + perOffX;
    var cy = ph - (slot.y + slot.h / 2 + perOffY);

    var drawOpts = {
      x: cx - drawW / 2,
      y: cy - drawH / 2,
      width: drawW,
      height: drawH,
      rotate: pdfLib.degrees(rot),
      opacity: settings.colorMode === 'bw' ? 1 : 1
    };

    if (embedResult.type === 'pdfPage') {
      page.drawPage(embedResult.embedded, drawOpts);
    } else {
      page.drawImage(embedResult.embedded, drawOpts);
    }

    if (settings.border) {
      page.drawRectangle({
        x: cx - drawW / 2,
        y: cy - drawH / 2,
        width: drawW,
        height: drawH,
        borderColor: pdfLib.rgb(0, 0, 0),
        borderWidth: 0.2
      });
    }
  }

  if (settings.cutline && layout.cutLines.length > 0) {
    for (var cl = 0; cl < layout.cutLines.length; cl++) {
      var line = layout.cutLines[cl];
      if (line.type === 'horizontal') {
        page.drawLine({
          start: { x: 0, y: ph - line.pos },
          end: { x: pw, y: ph - line.pos },
          color: pdfLib.rgb(0.7, 0.7, 0.7),
          dashArray: [1, 1],
          thickness: 0.1
        });
      } else if (line.type === 'vertical') {
        var vEndY = line.endY !== undefined ? line.endY : ph;
        page.drawLine({
          start: { x: line.pos, y: ph },
          end: { x: line.pos, y: ph - vEndY },
          color: pdfLib.rgb(0.7, 0.7, 0.7),
          dashArray: [1, 1],
          thickness: 0.1
        });
      }
    }
  }

  // pdf-lib StandardFonts (Helvetica) only support WinAnsiEncoding (Latin-1).
  // CJK characters (Chinese, Japanese, Korean) must be replaced with ASCII equivalents
  // to avoid "WinAnsi cannot encode" errors.
  function _safeText(s) {
    if (!s) return '';
    var result = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c <= 0xFF) { result += s.charAt(i); }
    }
    return result.replace(/\s+/g, ' ').trim();
  }

  if (settings.watermark && settings.wmText) {
    var wmText = _safeText(settings.wmText);
    if (wmText) {
      var wmSize = settings.wmSize || 60;
      var wmOpacity = settings.wmOpacity != null ? settings.wmOpacity : 0.15;
      page.drawText(wmText, {
        x: pw / 2 - wmText.length * wmSize * 0.15,
        y: ph / 2,
        size: wmSize,
        font: settings._fontBold,
        color: pdfLib.rgb(0.6, 0.6, 0.6),
        opacity: wmOpacity,
        rotate: pdfLib.degrees(30)
      });
    }
  }

  if (settings.pageNum || settings.printDate || (settings.footerText || '').trim()) {
    var dateStr = new Date().toISOString().slice(0, 10);
    var line1 = '';
    if (settings.pageNum) line1 = 'Page ' + (pageIdx + 1);
    if (settings.printDate) line1 += (line1 ? '  ' : '') + 'Printed ' + dateStr;
    if (line1) {
      page.drawText(line1, {
        x: pw / 2 - line1.length * 1.5,
        y: 8,
        size: 8,
        font: settings._font,
        color: pdfLib.rgb(0.5, 0.5, 0.5)
      });
    }
    if (settings.footerText && settings.footerText.trim()) {
      var ftText = _safeText(settings.footerText);
      if (ftText) {
        page.drawText(ftText, {
          x: pw / 2 - ftText.length * 1.5,
        y: 2,
        size: 8,
        font: settings._font,
        color: pdfLib.rgb(0.5, 0.5, 0.5)
      });
    }
  }
}
}

async function _composePdfBlob(files, settings, onProgress) {
  _srcPdfDocs = {};
  var key = _settingsKey(settings) + '|' + files.map(function(f) { return f.id + ':' + f.copies + ':' + f.rotation; }).join(',');
  if (_printCacheKey === key && _printCacheBlob) {
    return _printCacheBlob;
  }
  if (typeof pdfLib === 'undefined' || !pdfLib.PDFDocument) {
    throw new Error('pdf-lib 未加载');
  }
  showLoading('正在生成 PDF...');
  try {
    var pdfDoc = await pdfLib.PDFDocument.create();
    settings._font = await pdfDoc.embedFont(pdfLib.StandardFonts.Helvetica);
    settings._fontBold = await pdfDoc.embedFont(pdfLib.StandardFonts.HelveticaBold);
    var pages = buildPages(files, settings);
    for (var i = 0; i < pages.length; i++) {
      if (onProgress) onProgress(i + 1, pages.length);
      await _buildPage(pdfDoc, pages[i], i, settings);
    }
    var bytes = await pdfDoc.save();
    var blob = new Blob([bytes], { type: 'application/pdf' });
    _printCacheKey = key;
    _printCacheBlob = blob;
    return blob;
  } finally {
    _srcPdfDocs = {};
    hideLoading();
  }
}
function _downloadBlob(blob, filename) {
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(function() {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 1000);
}

function _printBlob(blob) {
  var url = URL.createObjectURL(blob);
  var iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.right = '-9999px';
  iframe.style.bottom = '0';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = '0';
  iframe.src = url;
  document.body.appendChild(iframe);
  iframe.onload = function() {
    setTimeout(function() {
      try {
        iframe.contentWindow.focus();
        iframe.contentWindow.print();
      } catch (e) {
        console.error('print failed:', e);
        toast('打印失败：' + e.message);
      }
    }, 200);
  };
  setTimeout(function() {
    if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
    URL.revokeObjectURL(url);
  }, 60000);
}

async function doPrint() {
  var files = getActiveFiles();
  if (!files.length) { toast('请先添加发票！'); return; }
  var settings = getSettings();
  try {
    var blob = await _composePdfBlob(files, settings);
    _printBlob(blob);
    markFilesAsPrinted(files);
  } catch (e) {
    console.error('doPrint error:', e);
    toast('打印失败：' + e.message);
  }
}

async function savePdf() {
  var files = getActiveFiles();
  if (!files.length) { toast('请先添加发票！'); return; }
  var settings = getSettings();
  try {
    var blob = await _composePdfBlob(files, settings);
    var ts = new Date().toISOString().slice(0, 19).replace(/[T:]/g, '-');
    _downloadBlob(blob, '发票-' + ts + '.pdf');
    markFilesAsPrinted(files);
  } catch (e) {
    console.error('savePdf error:', e);
    toast('保存失败：' + e.message);
  }
}

// Backward-compat: old code may still call doPdfiumPrint / doPdfReaderPrint
async function doPdfiumPrint() { return doPrint(); }
async function doPdfReaderPrint() { return doPrint(); }
function showPrintConfirm() { doPrint(); }
function confirmPrint() { doPrint(); }
