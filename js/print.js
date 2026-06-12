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

// CJK font priority: ① Local Font Access API (instant) → ② IndexedDB cache → ③ CDN download
var _cjkFontCacheName = 'ticketchan-cjk-font';
var _cjkFontCacheKey = 'NotoSansSC-Regular';
// CDN fallback: 4MB subset OTF (Regular only, no subsetting needed)
var _cjkFontCdnUrl = 'https://cdn.jsdelivr.net/gh/googlefonts/noto-cjk/main/Sans/SubsetOTF/SC/NotoSansSC-Regular.otf';

// CJK font names to try via Local Font Access API (ordered by priority)
var _cjkLocalFontNames = [
  'Microsoft YaHei', '微软雅黑',       // Windows
  'Noto Sans SC', 'Noto Sans CJK SC',  // Linux / cross-platform
  'Source Han Sans SC', '思源黑体',     // Adobe
  'PingFang SC', '苹方',               // macOS
  'SimHei', '黑体',                     // fallback Windows
  'SimSun', '宋体',                     // ultimate fallback
];

function _hasCjk(text) {
  if (!text) return false;
  // CJK Unified Ideographs + CJK Extension + Hangul + Hiragana/Katakana
  return /[\u4e00-\u9fff\u3400-\u4dbf\uac00-\ud7af\u3040-\u309f\u30a0-\u30ff]/.test(text);
}

function _needsCjkFont(settings) {
  if (settings.watermark && settings.watermarkText && _hasCjk(settings.watermarkText)) return true;
  if (settings.footer && settings.footerText && _hasCjk(settings.footerText)) return true;
  if (settings.number) return true; // numbers are ASCII, but user might customize
  return false;
}

async function _loadCjkFontFromCache() {
  return new Promise(function(resolve) {
    var req = indexedDB.open(_cjkFontCacheName, 1);
    req.onupgradeneeded = function(e) {
      var db = e.target.result;
      if (!db.objectStoreNames.contains('fonts')) {
        db.createObjectStore('fonts');
      }
    };
    req.onsuccess = function(e) {
      var db = e.target.result;
      try {
        var tx = db.transaction('fonts', 'readonly');
        var store = tx.objectStore('fonts');
        var getReq = store.get(_cjkFontCacheKey);
        getReq.onsuccess = function() {
          resolve(getReq.result || null);
        };
        getReq.onerror = function() { resolve(null); };
      } catch (err) { resolve(null); }
    };
    req.onerror = function() { resolve(null); };
  });
}

async function _saveCjkFontToCache(bytes) {
  return new Promise(function(resolve) {
    var req = indexedDB.open(_cjkFontCacheName, 1);
    req.onupgradeneeded = function(e) {
      var db = e.target.result;
      if (!db.objectStoreNames.contains('fonts')) {
        db.createObjectStore('fonts');
      }
    };
    req.onsuccess = function(e) {
      var db = e.target.result;
      try {
        var tx = db.transaction('fonts', 'readwrite');
        var store = tx.objectStore('fonts');
        store.put(bytes, _cjkFontCacheKey);
        tx.oncomplete = function() { resolve(true); };
        tx.onerror = function() { resolve(false); };
      } catch (err) { resolve(false); }
    };
    req.onerror = function() { resolve(false); };
  });
}

// ① Try Local Font Access API — read system CJK fonts directly, zero download.
async function _queryLocalCjkFont() {
  // Chrome 103+, needs 'local-fonts' permission
  if (!window.queryLocalFonts) { console.log('[print] Local Font Access API not available'); return null; }
  try {
    var all = await window.queryLocalFonts();
    console.log('[print] queried', all.length, 'local fonts, looking for CJK...');
    for (var fj = 0; fj < _cjkLocalFontNames.length; fj++) {
      var targetName = _cjkLocalFontNames[fj].toLowerCase();
      for (var fi = 0; fi < all.length; fi++) {
        var f = all[fi];
        var full = (f.fullName || '').toLowerCase();
        var family = (f.family || '').toLowerCase();
        if (full.indexOf(targetName) !== -1 || family.indexOf(targetName) !== -1) {
          console.log('[print] using local font:', f.fullName);
          var blob = await f.blob();
          var buf = await blob.arrayBuffer();
          _saveCjkFontToCache(buf); // cache for next time
          return buf;
        }
      }
    }
    console.warn('[print] no matching CJK font found among', all.length, 'local fonts');
  } catch (e) {
    console.warn('[print] Local Font Access failed:', e.message || e);
  }
  return null;
}

// ② CDN download fallback
async function _fetchCjkFont() {
  try {
    var resp = await fetch(_cjkFontCdnUrl, { mode: 'cors' });
    if (resp.ok) {
      var bytes = await resp.arrayBuffer();
      if (bytes && bytes.byteLength > 100000) {
        _saveCjkFontToCache(bytes);
        return bytes;
      }
    }
  } catch (e) { console.warn('[print] CDN font load failed:', e); }
  return null;
}

async function _getCjkFontBytes() {
  // 1. Check memory cache
  if (_cjkFontBytesCache) return _cjkFontBytesCache;
  // 2. Check IndexedDB cache
  var cached = await _loadCjkFontFromCache();
  if (cached) {
    _cjkFontBytesCache = cached;
    return cached;
  }
  // 3. Try Local Font Access API (instant, user's system fonts)
  var localBytes = await _queryLocalCjkFont();
  if (localBytes) {
    _cjkFontBytesCache = localBytes;
    return localBytes;
  }
  // 4. Download from CDN (one-time, cached thereafter)
  var bytes = await _fetchCjkFont();
  if (bytes) {
    _cjkFontBytesCache = bytes;
  }
  return bytes || null;
}

var _cjkFontBytesCache = null;

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
    watermark: s.watermark, watermarkText: s.watermarkText, watermarkSize: s.watermarkSize, watermarkOpacity: s.watermarkOpacity,
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
  var ptPerMm = 72 / 25.4;
  var layout = calculateLayout(settings, ptPerMm);
  var pw = layout.pw;
  var ph = layout.ph;
  var page = pdfDoc.addPage([pw, ph]);
  page.drawRectangle({ x: 0, y: 0, width: pw, height: ph, color: pdfLib.rgb(1, 1, 1) });

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

    // Use pt dimensions for PDF pages, convert pixel→pt for images.
    var objW, objH;
    if (embedResult.type === 'pdfPage') {
      objW = f.srcPageWidthPt || embedResult.width;
      objH = f.srcPageHeightPt || embedResult.height;
    } else {
      var imgDpi = f.renderDpi || 300;
      objW = (f.ow || embedResult.width) * 72 / imgDpi;
      objH = (f.oh || embedResult.height) * 72 / imgDpi;
    }

    var fitScale = Math.min(slot.w / objW, slot.h / objH);
    if (settings.fitMode === 'fill') fitScale = Math.max(slot.w / objW, slot.h / objH);
    else if (settings.fitMode === 'original') fitScale = 1;
    if (settings.fitMode === 'custom' && settings.customScale) fitScale *= settings.customScale;
    fitScale *= perScale;

    var drawW = objW * fitScale;
    var drawH = objH * fitScale;
    var cx = slot.x + slot.w / 2 + perOffX * ptPerMm;
    var cy = ph - (slot.y + slot.h / 2 + perOffY * ptPerMm);

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

    if (settings.number) {
      var numStr = String(pageIdx * settings.cols * settings.rows + i + 1);
      page.drawText(numStr, {
        x: slot.x + 2,
        y: ph - slot.y - 2 - 8,
        size: 8,
        font: settings._font,
        color: pdfLib.rgb(0.5, 0.5, 0.5)
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

  function _safeText(s) {
    if (!s) return '';
    return String(s).replace(/\s+/g, ' ').trim();
  }

  if (settings.watermark && settings.watermarkText) {
    var wmText = _safeText(settings.watermarkText);
    if (wmText) {
      var wmSize = (settings.watermarkSize || 60) * ptPerMm;
      var wmOpacity = settings.watermarkOpacity != null ? settings.watermarkOpacity : 0.15;
      var wmAngle = settings.watermarkAngle || 30;
      page.drawText(wmText, {
        x: pw / 2 - wmText.length * wmSize * 0.15,
        y: ph / 2,
        size: wmSize,
        font: settings._fontBold,
        color: pdfLib.rgb(0.6, 0.6, 0.6),
        opacity: wmOpacity,
        rotate: pdfLib.degrees(wmAngle)
      });
    }
  }

  if (settings.number) {
    for (var si = 0; si < layout.slots.length; si++) {
      var sn = layout.slots[si];
      var numStr = String(pageIdx * settings.cols * settings.rows + si + 1);
      page.drawText(numStr, {
        x: sn.x + 4,
        y: ph - sn.y - 4 - 8,
        size: 8,
        font: settings._font,
        color: pdfLib.rgb(0.5, 0.5, 0.5)
      });
    }
  }

  if (settings.pageNum || settings.printDate || (settings.footerText || '').trim()) {
    var fm = layout.fm || 0;
    var lineHeight = 5 * ptPerMm;
    var footerFontSize = 8;
    var footerColor = pdfLib.rgb(0.5, 0.5, 0.5);

    var ftText = _safeText(settings.footerText);
    var pageNumStr = '';
    if (settings.pageNum) pageNumStr = 'Page ' + (pageIdx + 1);
    var dateStr = '';
    if (settings.printDate) dateStr = new Date().toISOString().slice(0, 10);

    var textY = 3 * ptPerMm;
    if (ftText) {
      page.drawText(ftText, {
        x: pw / 2 - ftText.length * footerFontSize * 0.3,
        y: textY,
        size: footerFontSize,
        font: settings._font,
        color: footerColor
      });
      textY += lineHeight;
    }
    if (pageNumStr && dateStr) {
      page.drawText(pageNumStr, { x: 10, y: textY, size: footerFontSize, font: settings._font, color: footerColor });
      page.drawText(dateStr, { x: pw - dateStr.length * footerFontSize * 0.6 - 10, y: textY, size: footerFontSize, font: settings._font, color: footerColor });
    } else if (pageNumStr) {
      page.drawText(pageNumStr, { x: pw / 2 - pageNumStr.length * footerFontSize * 0.3, y: textY, size: footerFontSize, font: settings._font, color: footerColor });
    } else if (dateStr) {
      page.drawText(dateStr, { x: pw / 2 - dateStr.length * footerFontSize * 0.3, y: textY, size: footerFontSize, font: settings._font, color: footerColor });
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
    if (typeof fontkit !== 'undefined') pdfDoc.registerFontkit(fontkit);
    var fontBytes = null;
    // Only download CJK font when watermark/footer contains CJK characters.
    // Font is cached in IndexedDB after first download (~8-17MB one-time cost).
    if (_needsCjkFont(settings)) {
      console.log('[print] CJK text detected, loading font...');
      fontBytes = await _getCjkFontBytes();
      if (!fontBytes) console.warn('[print] CJK font download failed, will fallback to Helvetica');
    }
    if (fontBytes) {
      try {
        // Try subsetting first (works for TTF/glyf fonts)
        settings._font = await pdfDoc.embedFont(fontBytes, { subset: true });
        settings._fontBold = settings._font;
      } catch (e) {
        // CFF/OTF fonts don't support subsetting — retry without subset
        if (e.message && e.message.indexOf('CFF') !== -1) {
          try {
            settings._font = await pdfDoc.embedFont(fontBytes);
            settings._fontBold = settings._font;
            console.log('[print] CJK font embedded without subset (OTF/CFF)');
          } catch (e2) {
            console.warn('[print] CJK embed failed, fallback to Helvetica:', e2);
            settings._font = await pdfDoc.embedFont(pdfLib.StandardFonts.Helvetica);
            settings._fontBold = await pdfDoc.embedFont(pdfLib.StandardFonts.HelveticaBold);
          }
        } else {
          console.warn('[print] CJK embed failed, fallback to Helvetica:', e);
          settings._font = await pdfDoc.embedFont(pdfLib.StandardFonts.Helvetica);
          settings._fontBold = await pdfDoc.embedFont(pdfLib.StandardFonts.HelveticaBold);
        }
      }
    } else {
      settings._font = await pdfDoc.embedFont(pdfLib.StandardFonts.Helvetica);
      settings._fontBold = await pdfDoc.embedFont(pdfLib.StandardFonts.HelveticaBold);
    }
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
