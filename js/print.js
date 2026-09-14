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

// Extract the first individual font from a TTC (TrueType Collection) buffer.
// pdf-lib cannot handle TTC directly — it needs a standalone OTF/TTF.
function _extractFontFromTtc(ttcBuffer) {
  try {
    var src = new Uint8Array(ttcBuffer);
    var view = new DataView(src.buffer, src.byteOffset, src.byteLength);
    var tag = String.fromCharCode(view.getUint8(0), view.getUint8(1), view.getUint8(2), view.getUint8(3));
    if (tag !== 'ttcf') return ttcBuffer; // Not a TTC, return as-is

    var numFonts = view.getUint32(8);
    if (numFonts < 1) return ttcBuffer;
    var firstFontOffset = view.getUint32(12);

    // Read the individual font's table directory
    var fv = new DataView(src.buffer, src.byteOffset + firstFontOffset, src.byteLength - firstFontOffset);
    var sfVersion = fv.getUint32(0);
    var numTables = fv.getUint16(4);

    // Collect table records (tag, checksum, offset, length)
    var tables = [];
    for (var i = 0; i < numTables; i++) {
      var ro = 12 + i * 16;
      var tTag = String.fromCharCode(fv.getUint8(ro), fv.getUint8(ro+1), fv.getUint8(ro+2), fv.getUint8(ro+3));
      var tChecksum = fv.getUint32(ro + 4);
      var tOffset = fv.getUint32(ro + 8);
      var tLength = fv.getUint32(ro + 12);
      tables.push({ tag: tTag, checksum: tChecksum, offset: tOffset, length: tLength });
    }

    // Calculate output size: header (12) + table records (16 * numTables) + table data (4-byte aligned)
    var headerSize = 12 + numTables * 16;
    var totalSize = headerSize;
    for (var i = 0; i < tables.length; i++) {
      totalSize += (tables[i].length + 3) & ~3;
    }

    var out = new Uint8Array(totalSize);
    var ov = new DataView(out.buffer);

    // Write font header
    ov.setUint32(0, sfVersion);
    ov.setUint16(4, numTables);
    var pow2 = 1;
    while (pow2 * 2 <= numTables) pow2 *= 2;
    ov.setUint16(6, pow2 * 16);
    ov.setUint16(8, Math.log2(pow2) | 0);
    ov.setUint16(10, numTables * 16 - pow2 * 16);

    // Write table records and copy table data
    var dataOff = headerSize;
    for (var i = 0; i < tables.length; i++) {
      var t = tables[i];
      var ro = 12 + i * 16;
      // Write tag as 4 bytes
      for (var c = 0; c < 4; c++) ov.setUint8(ro + c, t.tag.charCodeAt(c));
      ov.setUint32(ro + 4, t.checksum);
      ov.setUint32(ro + 8, dataOff);
      ov.setUint32(ro + 12, t.length);
      // Copy table data from source
      out.set(new Uint8Array(src.buffer, src.byteOffset + t.offset, t.length), dataOff);
      dataOff += (t.length + 3) & ~3;
    }

    console.log('[print] extracted font from TTC, tables:', numTables, 'size:', totalSize);
    return out.buffer;
  } catch (e) {
    console.warn('[print] TTC extraction failed, returning raw buffer:', e);
    return ttcBuffer;
  }
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
          // TTC fonts must be extracted to individual OTF/TTF for pdf-lib
          buf = _extractFontFromTtc(buf);
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
    cached = _extractFontFromTtc(cached); // Handle cached TTC fonts
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
    gapH: s.gapH, gapV: s.gapV, colorMode: s.colorMode, cutline: s.cutline, border: s.border, number: s.number,
    trimWhite: s.trimWhite,
    watermark: s.watermark, watermarkText: s.watermarkText, watermarkSize: s.watermarkSize, watermarkOpacity: s.watermarkOpacity,
    watermarkColor: s.watermarkColor, watermarkAngle: s.watermarkAngle,
    globalRotation: s.globalRotation, footerText: s.footerText, pageNum: s.pageNum, customFM: s.customFM,
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
// For PDF sources: returns { type:'pdfPage', embedded, width(pt), height(pt), ptW, ptH }
// For image sources: returns { type:'image', embedded, width(px), height(px) }
// On failure or unsupported type: returns null.
// settings._rasterOnly=true → skip vector embedding, always use raster image.
// settings.trimWhite + fileObj.trimmedBox → 源 PDF 走 embedPage 裁剪框矢量裁切，
// 无法安全换算坐标时退回裁剪后的位图，保证与预览一致。
// 页面带 /Rotate 时把旋转烘焙进 XObject Matrix（与桌面版同款），无法烘焙时退回位图。
// ptW/ptH: 实际绘制尺寸（点）。裁剪（尺寸变化）或 /Rotate 烘焙（宽高互换）时给出。
// 页面 /Rotate 度数。pdf-lib 返回 Rotation 对象 { type, angle }（非数值），
// radians 需换算；读取失败返回 NaN（调用方按"不可信"处理）。
function _pageRotationDeg(srcPage) {
  try {
    var rot = srcPage.getRotation();
    if (typeof rot === 'number') return rot;
    if (rot) return rot.type === 'radians' ? (rot.angle || 0) * 180 / Math.PI : (rot.angle || 0);
    return 0;
  } catch (e) { return NaN; }
}

// /Rotate 90/180/270 的烘焙矩阵（与桌面版 Rust extract_page_as_form_xobject 同款语义）：
//   90:  (x,y) → (y, pageW - x)          180: (x,y) → (pageW - x, pageH - y)
//   270: (x,y) → (pageH - y, x)
// outW/outH 是旋转后的页面尺寸（drawPage 的目标宽高）。
// ⚠️ pdf-lib drawPage 会先按 width/embWidth、height/embHeight 归一化缩放（emb=BBox
// 尺寸，未旋转），XMatrix 在缩放之内，所以这里预除该缩放（XMatrix = S⁻¹ · T），
// 使内容经 Matrix + drawPage 缩放后正好铺满 outW × outH。
function _rotateMatrix(rot, pageW, pageH, outW, outH) {
  var ix = pageW / outW, iy = pageH / outH;
  if (rot === 90) return [0, -iy, ix, 0, 0, iy * pageW];
  if (rot === 180) return [-ix, 0, 0, -iy, ix * pageW, iy * pageH];
  if (rot === 270) return [0, iy, -ix, 0, ix * pageH, 0];
  return null;
}

// 能否把 /Rotate 烘焙进 XObject：要求 MediaBox 与 CropBox 一致且原点在 0。
// pdf-lib 的 embedPage BBox 以 0 为基准、忽略原点偏移；不一致时烘焙会错位，退回位图更安全。
function _canBakeRotation(srcPage) {
  try {
    var mb = srcPage.getMediaBox(), cb = srcPage.getCropBox();
    if (!mb || !cb) return false;
    if (Math.abs(mb.x) > 0.01 || Math.abs(mb.y) > 0.01) return false;
    return Math.abs(mb.x - cb.x) < 0.01 && Math.abs(mb.y - cb.y) < 0.01
      && Math.abs(mb.width - cb.width) < 0.01 && Math.abs(mb.height - cb.height) < 0.01;
  } catch (e) { return false; }
}

function _trimBBoxPt(srcPage, fileObj, box) {
  // 预览位图 → 源页面用户空间（PDF 点）的裁剪框换算。
  // PDF.js 预览渲染的是 CropBox（view）区域，基准必须用 CropBox：
  // 用 MediaBox 会在「CropBox 与 MediaBox 等尺寸但原点偏移」时整体错位，
  // 且会错杀 CropBox < MediaBox 的页面（其实可正确换算）。
  // 只有当预览渲染范围与 CropBox 完全对应时才成立，否则返回 null 交由位图路径。
  var dpi = fileObj.renderDpi || 300;
  var cb;
  try { cb = srcPage.getCropBox(); } catch (e) { return null; }
  if (!cb || !cb.width || !cb.height) return null;
  var angle = _pageRotationDeg(srcPage);
  if (!isFinite(angle) || angle % 360 !== 0) return null; // /Rotate 页面预览为旋转后视图，坐标不可直接换算
  // 预览按 CropBox 渲染，宽高对不上（如 /Rotate 导致的宽高互换）说明不可直接换算
  if (Math.abs(cb.width - fileObj.srcPageWidthPt) > 1.5) return null;
  if (Math.abs(cb.height - fileObj.srcPageHeightPt) > 1.5) return null;
  var ptPerPx = 72 / dpi;
  return {
    left: cb.x + box.x * ptPerPx,
    right: cb.x + (box.x + box.w) * ptPerPx,
    bottom: cb.y + cb.height - (box.y + box.h) * ptPerPx,
    top: cb.y + cb.height - box.y * ptPerPx
  };
}

// Convert a dataURL image to grayscale or B&W via canvas
function _convertImageDataUrl(dataUrl, colorMode) {
  if (!colorMode || colorMode === 'color') return dataUrl;
  return new Promise(function(resolve) {
    var img = new Image();
    img.onload = function() {
      var c = document.createElement('canvas');
      c.width = img.width; c.height = img.height;
      var ctx = c.getContext('2d');
      ctx.drawImage(img, 0, 0);
      var id = ctx.getImageData(0, 0, c.width, c.height);
      var d = id.data;
      for (var i = 0; i < d.length; i += 4) {
        var gray = 0.299 * d[i] + 0.587 * d[i+1] + 0.114 * d[i+2];
        if (colorMode === 'bw') gray = gray > 128 ? 255 : 0;
        d[i] = d[i+1] = d[i+2] = gray;
      }
      ctx.putImageData(id, 0, 0);
      resolve(c.toDataURL('image/jpeg', 0.92));
    };
    img.src = dataUrl;
  });
}

async function _embedForFile(pdfDoc, fileObj, settings) {
  if (!fileObj) return null;
  if (fileObj._xmlInvoice) return null;

  var colorMode = settings.colorMode || 'color';
  var trimBox = settings.trimWhite ? fileObj.trimmedBox : null;

  // Vector path: embed original PDF page (skip if rasterOnly or colorMode needs conversion)
  if (!settings._rasterOnly && colorMode === 'color'
      && fileObj.srcPdfBytes && fileObj.srcPageIndex != null) {
    try {
      var srcDoc = await _getOrLoadSrcPdf(fileObj);
      var srcPage = srcDoc.getPage(fileObj.srcPageIndex);
      if (trimBox) {
        var bbox = _trimBBoxPt(srcPage, fileObj, trimBox);
        if (bbox) {
          var cropped = await pdfDoc.embedPage(srcPage, bbox);
          return {
            type: 'pdfPage',
            embedded: cropped,
            width: cropped.width,
            height: cropped.height,
            ptW: cropped.width,
            ptH: cropped.height
          };
        }
        // 坐标无法安全换算 → 不裁切的矢量图会与预览不一致，改为走裁剪后的位图
      } else {
        var bake = ((_pageRotationDeg(srcPage) % 360) + 360) % 360;
        if (bake === 0) {
          var embedded = await pdfDoc.embedPage(srcPage);
          return {
            type: 'pdfPage',
            embedded: embedded,
            width: srcPage.getWidth(),
            height: srcPage.getHeight()
          };
        }
        // /Rotate 页面：把旋转烘焙进 XObject 的 Matrix（与桌面版同款矩阵），
        // 否则未旋转的内容会被按旋转后的尺寸拉伸；无法安全烘焙时落到
        // 下方位图路径，保证与预览一致
        if (isFinite(bake) && _canBakeRotation(srcPage)) {
          var pw = srcPage.getWidth(), ph = srcPage.getHeight();
          var outW = bake === 180 ? pw : ph;
          var outH = bake === 180 ? ph : pw;
          var matrix = _rotateMatrix(bake, pw, ph, outW, outH);
          if (matrix) {
            var emb = await pdfDoc.embedPage(srcPage, { left: 0, bottom: 0, right: pw, top: ph }, matrix);
            return {
              type: 'pdfPage',
              embedded: emb,
              width: outW,
              height: outH,
              ptW: outW,
              ptH: outH
            };
          }
        }
      }
    } catch (e) {
      console.warn('embedPage failed, falling back to raster:', e);
    }
  }

  // Raster path: embed as image
  var srcUrl = (settings.trimWhite && fileObj.trimmedUrl) ? fileObj.trimmedUrl : fileObj.previewUrl;
  if (!srcUrl) return null;
  // Apply color mode conversion
  if (colorMode && colorMode !== 'color') {
    srcUrl = await _convertImageDataUrl(srcUrl, colorMode);
  }
  var isJpeg = srcUrl.indexOf('data:image/jpeg') === 0
    || srcUrl.indexOf('data:image/jpg') === 0;
  var bytes = isJpeg ? _jpegBlobFromDataUrl(srcUrl) : _pngBlobFromDataUrl(srcUrl);
  var img = isJpeg ? await pdfDoc.embedJpg(bytes) : await pdfDoc.embedPng(bytes);
  return { type: 'image', embedded: img, width: img.width, height: img.height };
}

async function _buildPage(pdfDoc, pageFiles, pageIdx, totalPages, settings) {
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
    var embedResult = await _embedForFile(pdfDoc, f, settings);
    if (!embedResult) continue;

    var rot = getRotation(f, slot, settings);
    var perScale = f.slotScale || 1;
    var perOffX = f.slotOffsetX || 0;
    var perOffY = f.slotOffsetY || 0;

    // Use pt dimensions for PDF pages, convert pixel→pt for images.
    // 裁剪白边时统一用裁剪后的尺寸（getObjDims），与预览适配保持一致。
    var objW, objH;
    if (embedResult.type === 'pdfPage') {
      objW = embedResult.ptW || f.srcPageWidthPt || embedResult.width;
      objH = embedResult.ptH || f.srcPageHeightPt || embedResult.height;
    } else {
      var imgDpi = f.renderDpi || 300;
      var objDims = getObjDims(f, settings);
      objW = objDims.w * 72 / imgDpi;
      objH = objDims.h * 72 / imgDpi;
    }

    // 先旋转后适配（与预览/桌面端一致）：90°/270° 按旋转后视觉宽高 fit 槽位
    var isRot90 = (rot === 90 || rot === 270);
    var fitW = isRot90 ? objH : objW;
    var fitH = isRot90 ? objW : objH;

    var fitScale = Math.min(slot.w / fitW, slot.h / fitH);
    if (settings.fitMode === 'fill') fitScale = Math.max(slot.w / fitW, slot.h / fitH);
    else if (settings.fitMode === 'original') fitScale = 1;
    if (settings.fitMode === 'custom' && settings.customScale) fitScale *= settings.customScale;
    fitScale *= perScale;

    var visW = fitW * fitScale;
    var visH = fitH * fitScale;
    var cx = slot.x + slot.w / 2 + perOffX * ptPerMm;
    // 「裁剪白边」开启时垂直贴顶（视觉盒顶 = 槽位顶）：中心 y = slot.y + visH/2；
    // 否则在槽位内居中。水平方向始终居中 —— 左右裁剪本来就对得准。
    var cy = settings.trimWhite
      ? ph - (slot.y + visH / 2 + perOffY * ptPerMm)
      : ph - (slot.y + slot.h / 2 + perOffY * ptPerMm);

    // pdf-lib drawImage/drawPage 的 rotate 绕 (x,y) 锚点（未旋转盒左下角）且正角度为逆时针，
    // 与 CSS 旋转（绕中心、顺时针为正）不同。这里换算锚点使旋转后视觉盒以 (cx,cy) 为中心：
    //   anchor = (cx,cy) - R_θ·(unrotW/2, unrotH/2)，方向取 degrees(-rot)
    var unrotW = objW * fitScale;
    var unrotH = objH * fitScale;
    var drawOpts;
    if (rot === 90) {
      drawOpts = { x: cx - visW / 2, y: cy + visH / 2, width: unrotW, height: unrotH, rotate: pdfLib.degrees(-90) };
    } else if (rot === 180) {
      drawOpts = { x: cx + visW / 2, y: cy + visH / 2, width: unrotW, height: unrotH, rotate: pdfLib.degrees(180) };
    } else if (rot === 270) {
      drawOpts = { x: cx + visW / 2, y: cy - visH / 2, width: unrotW, height: unrotH, rotate: pdfLib.degrees(90) };
    } else {
      drawOpts = { x: cx - visW / 2, y: cy - visH / 2, width: unrotW, height: unrotH };
    }

    if (embedResult.type === 'pdfPage') {
      page.drawPage(embedResult.embedded, drawOpts);
    } else {
      page.drawImage(embedResult.embedded, drawOpts);
    }

    if (settings.border) {
      page.drawRectangle({
        x: cx - visW / 2,
        y: cy - visH / 2,
        width: visW,
        height: visH,
        borderColor: pdfLib.rgb(0, 0, 0),
        borderWidth: 0.2
      });
    }

    if (settings.number) {
      var numStr = String(pageIdx * settings.cols * settings.rows + i + 1);
      // Match preview: slot-num is at top-right of slot (top:3px, right:3px)
      var numX = slot.x + slot.w - 3 * ptPerMm - numStr.length * 8 * 0.5;
      var numY = ph - slot.y - 3 * ptPerMm - 8;
      page.drawRectangle({
        x: numX - 2,
        y: numY - 1,
        width: numStr.length * 8 * 0.5 + 5 * ptPerMm,
        height: 8 + 2 * ptPerMm,
        color: pdfLib.rgb(0, 0, 0),
        opacity: 0.55
      });
      page.drawText(numStr, {
        x: numX,
        y: numY,
        size: 8,
        font: settings._font,
        color: pdfLib.rgb(1, 1, 1)
      });
    }

    // Watermark per slot (matches preview: centered within each slot)
    if (settings.watermark && settings.watermarkText) {
      var wmText = _safeText(settings.watermarkText);
      if (wmText) {
        var wmSize = (settings.watermarkSize || 60) * ptPerMm;
        var wmOpacity = settings.watermarkOpacity != null ? settings.watermarkOpacity : 0.15;
        var wmAngle = settings.watermarkAngle || 30;
        var slotCx = slot.x + slot.w / 2;
        var slotCy = ph - (slot.y + slot.h / 2);
        // Estimate text width for centering: CJK char ≈ fontSize, ASCII ≈ fontSize * 0.55
        var wmTextWidth = 0;
        for (var ci = 0; ci < wmText.length; ci++) {
          wmTextWidth += wmText.charCodeAt(ci) > 0x2E80 ? wmSize : wmSize * 0.55;
        }
        var wmColor = settings.watermarkColor || '#ff0000';
        var wmR = parseInt(wmColor.slice(1, 3), 16) / 255;
        var wmG = parseInt(wmColor.slice(3, 5), 16) / 255;
        var wmB = parseInt(wmColor.slice(5, 7), 16) / 255;
        page.drawText(wmText, {
          x: slotCx - wmTextWidth / 2,
          y: slotCy,
          size: wmSize,
          font: settings._fontBold,
          color: pdfLib.rgb(wmR, wmG, wmB),
          opacity: wmOpacity,
          // pdf-lib 正角度为逆时针，预览 CSS rotate 为顺时针，取负保持一致
          rotate: pdfLib.degrees(-(wmAngle || 0))
        });
      }
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

  if (settings.pageNum || settings.printDate || (settings.footerText || '').trim()) {
    var fm = layout.fm || 0;
    var lineHeight = 5 * ptPerMm;
    var footerFontSize = 8;
    var footerColor = pdfLib.rgb(0.58, 0.64, 0.72); // match preview #94a3b8

    var ftText = _safeText(settings.footerText);
    var pageNumStr = '';
    if (settings.pageNum) pageNumStr = '第 ' + (pageIdx + 1) + ' 页 / 共 ' + totalPages + ' 页';
    var dateStr = '';
    if (settings.printDate) {
      var now = new Date();
      dateStr = '打印日期 ' + now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
    }

    // Match preview: text stacks from bottom up
    var textBottomPx = 3 * ptPerMm;
    var pageNumBottomPx = textBottomPx;
    if (ftText) pageNumBottomPx += lineHeight;

    if (ftText) {
      page.drawText(ftText, {
        x: pw / 2 - ftText.length * footerFontSize * 0.3,
        y: textBottomPx,
        size: footerFontSize,
        font: settings._font,
        color: footerColor
      });
    }
    if (pageNumStr && dateStr) {
      page.drawText(pageNumStr, { x: 10, y: pageNumBottomPx, size: footerFontSize, font: settings._font, color: footerColor });
      page.drawText(dateStr, { x: pw - dateStr.length * footerFontSize * 0.55 - 10, y: pageNumBottomPx, size: footerFontSize, font: settings._font, color: footerColor });
    } else if (pageNumStr) {
      page.drawText(pageNumStr, { x: pw / 2 - pageNumStr.length * footerFontSize * 0.3, y: pageNumBottomPx, size: footerFontSize, font: settings._font, color: footerColor });
    } else if (dateStr) {
      page.drawText(dateStr, { x: pw / 2 - dateStr.length * footerFontSize * 0.3, y: pageNumBottomPx, size: footerFontSize, font: settings._font, color: footerColor });
    }
  }
}

async function _composePdfBlob(files, settings, onProgress) {
  _srcPdfDocs = {};
  var rasterKey = settings._rasterOnly ? 'R' : 'V';
  // 裁剪白边是异步生成的，打印时可能尚未就绪；把每票的裁剪状态纳入缓存键，
  // 避免裁剪完成前的旧结果被复用。
  var key = _settingsKey(settings) + '|' + rasterKey + '|' + files.map(function(f) {
    return f.id + ':' + f.copies + ':' + f.rotation + ':' + (f.trimmedBox ? 'T' : 'n');
  }).join(',');
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
        // Try subsetting first (smaller PDF, works for TTF/glyf fonts)
        settings._font = await pdfDoc.embedFont(fontBytes, { subset: true });
        settings._fontBold = settings._font;
        console.log('[print] CJK font embedded with subset');
      } catch (e) {
        // Subsetting can fail for many reasons (CFF/OTF, TTC, fontkit issues)
        // Always retry without subsetting
        console.warn('[print] CJK subset failed, retrying without subset:', e.message || e);
        try {
          settings._font = await pdfDoc.embedFont(fontBytes);
          settings._fontBold = settings._font;
          console.log('[print] CJK font embedded without subset');
        } catch (e2) {
          console.warn('[print] CJK embed failed, fallback to Helvetica:', e2);
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
      await _buildPage(pdfDoc, pages[i], i, pages.length, settings);
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

async function saveRasterPdf() {
  var files = getActiveFiles();
  if (!files.length) { toast('请先添加发票！'); return; }
  var settings = getSettings();
  settings._rasterOnly = true;
  try {
    var blob = await _composePdfBlob(files, settings);
    var ts = new Date().toISOString().slice(0, 19).replace(/[T:]/g, '-');
    _downloadBlob(blob, '发票-' + ts + '.pdf');
    markFilesAsPrinted(files);
  } catch (e) {
    console.error('saveRasterPdf error:', e);
    toast('保存失败：' + e.message);
  }
}

