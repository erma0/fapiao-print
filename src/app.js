// =====================================================
// 发票批量打印工具 — 主入口
// v1.3.0 — 重构版本
// =====================================================

// Detect Tauri — use var to avoid conflict with Tauri's injected scripts
var isTauri = window.__TAURI_INTERNALS__ !== undefined;
var invoke  = isTauri ? window.__TAURI_INTERNALS__.invoke : null;
console.log('发票批量打印 v1.5.1 | isTauri:', isTauri);

// =====================================================
// Constants
// =====================================================
var PAPER = { A4:{w:210,h:297}, A5:{w:148,h:210}, B5:{w:176,h:250}, letter:{w:216,h:279}, legal:{w:216,h:356} };
var MM2PX = 96 / 25.4;
var PDF_RENDER_DPI = 300;  // Must match Rust RENDER_DPI — validated at startup via get_config
var MIN_RENDER_PX = 3508;  // A4 long side at 300 DPI — minimum rendered pixels
var WHITE_THRESHOLD = 245; // Pixel value threshold for white-edge trimming

// CMap & font base URL — local first, CDN fallback
// Set to local paths; if files not found, PDF.js will fall back gracefully
var CMAP_BASE_URL = 'cmaps/';
var STD_FONT_BASE_URL = 'standard_fonts/';

// =====================================================
// State
// =====================================================
var S = {
  files: [],
  currentPage: 0,
  totalPages: 0,
  viewZoom: 0,
  layout: { cols: 1, rows: 1, orient: 'landscape' },
  editIdx: -1,
  amtMode: 'tax',
  feat: {
    cutline: true, number: false, border: false, trimWhite: false,
    watermark: false, collate: true, duplex: false, pageNum: false,
    printDate: false, confirmPrint: true,
    autoOpenPdf: true
  }
};

// =====================================================
// File Object Factory — unified creation with defaults
// =====================================================
function createFileObj(opts) {
  return {
    id: opts.id || ('f' + Date.now() + Math.random().toString(36).slice(2)),
    name: opts.name || '',
    size: opts.size || 0,
    type: opts.type || '',
    checked: true,
    previewUrl: opts.previewUrl || '',
    copies: 1,
    rotation: 0,
    note: '',
    amount: opts.amount || 0,
    amountTax: opts.amountTax || 0,
    amountNoTax: opts.amountNoTax || 0,
    img: opts.img || null,
    ow: opts.img ? opts.img.naturalWidth : (opts.ow || 0),
    oh: opts.img ? opts.img.naturalHeight : (opts.oh || 0),
    renderDpi: opts.renderDpi || PDF_RENDER_DPI,
    sellerName: opts.sellerName || '',
    sellerCreditCode: opts.sellerCreditCode || '',
    _ocrText: opts._ocrText || ''
  };
}

// =====================================================
// Helpers
// =====================================================
var toastT = null;
function toast(msg, dur) { dur = dur || 2500; var e = document.getElementById('toast'); e.textContent = msg; e.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(function() { e.classList.remove('show'); }, dur); }
function syncSlider(s, n) { document.getElementById(n).value = s.value; }
function syncRange(n, s) { document.getElementById(s).value = n.value; }
function showLoading(t) { document.getElementById('loadingText').textContent = t || '处理中...'; document.getElementById('loading').classList.remove('hidden'); }
function hideLoading() { document.getElementById('loading').classList.add('hidden'); }
function fmtSize(b) { return b < 1024 ? b + 'B' : b < 1048576 ? (b / 1024).toFixed(1) + 'KB' : (b / 1048576).toFixed(1) + 'MB'; }
function escHtml(s) { return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

// Convert data URL to Uint8Array
function dataUrlToUint8Array(dataUrl) {
  var base64 = dataUrl.split(',')[1] || dataUrl;
  var binaryStr = atob(base64);
  var bytes = new Uint8Array(binaryStr.length);
  for (var i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
  return bytes;
}

// =====================================================
// FILE UPLOAD — via Tauri dialog plugin
// =====================================================
async function triggerUpload() {
  if (isTauri && invoke) {
    try {
      var result = await invoke('plugin:dialog|open', {
        options: {
          multiple: true,
          title: '选择发票文件',
          filters: [{ name: '发票文件', extensions: ['pdf', 'jpg', 'jpeg', 'png', 'bmp', 'webp', 'tiff', 'tif', 'ofd'] }]
        }
      });
      if (!result) return;
      var paths = typeof result === 'string' ? [result] : (Array.isArray(result) ? result : []);
      if (paths.length === 0) return;

      showLoading('读取 ' + paths.length + ' 个文件...');
      var fileDataList = await invoke('open_invoice_files', { paths: paths });
      hideLoading();

      if (fileDataList && fileDataList.length > 0) {
        await processFileDataList(fileDataList);
      }
    } catch (err) {
      console.error('Dialog error:', err);
      hideLoading();
      toast('打开文件对话框失败: ' + String(err));
    }
  } else {
    document.getElementById('fileInput').click();
  }
}

async function handleFileInput(fl) {
  if (!fl || !fl.length) return;
  await processFiles(Array.from(fl));
  document.getElementById('fileInput').value = '';
}

// Process FileData array from Rust backend — fast mode: show preview first, OCR in background
async function processFileDataList(fileDataList) {
  var added = 0;
  var total = fileDataList.length;
  var completed = 0;

  // Use lightweight progress indicator instead of blocking overlay
  var progressEl = document.getElementById('importProgress');
  if (progressEl) {
    progressEl.style.display = 'flex';
    updateImportProgress(0, total);
  }

  var promises = fileDataList.map(function(fd) {
    return loadFileFromDataUrlFast(fd.name, fd.dataUrl, fd.size, fd.ext, fd.path).then(function(r) {
      completed++;
      updateImportProgress(completed, total);
      if (r) {
        var items = Array.isArray(r) ? r : [r];
        items.forEach(function(item) {
          if (item) { S.files.push(item); added++; }
        });
        renderFileList(); updatePreview(); updatePrintBtn();
      }
    }).catch(function(err) {
      completed++;
      updateImportProgress(completed, total);
      console.error('Load file error:', fd.name, err);
    });
  });

  await Promise.all(promises);
  if (progressEl) progressEl.style.display = 'none';
  if (added > 0) toast('已添加 ' + added + ' 张发票，识别中...');
}

function updateImportProgress(done, total) {
  var bar = document.getElementById('importProgressBar');
  var text = document.getElementById('importProgressText');
  if (bar) bar.style.width = Math.round(done / total * 100) + '%';
  if (text) text.textContent = done + '/' + total;
}

// Process an array of File objects (browser fallback) — fast mode
async function processFiles(files) {
  showLoading('载入 ' + files.length + ' 个文件...');
  var added = 0;
  var total = files.length;
  var completed = 0;

  var promises = files.map(function(file) {
    return loadFileFast(file).then(function(r) {
      completed++;
      document.getElementById('loadingText').textContent = '载入 ' + completed + '/' + total + '...';
      if (r) {
        var items = Array.isArray(r) ? r : [r];
        items.forEach(function(item) {
          if (item) { S.files.push(item); added++; }
        });
        renderFileList(); updatePreview(); updatePrintBtn();
      }
    }).catch(function(err) {
      completed++;
      console.error('Load file error:', file.name, err);
    });
  });

  await Promise.all(promises);
  hideLoading(); renderFileList(); updatePreview(); updatePrintBtn();
  if (added > 0) toast('已添加 ' + added + ' 张发票，识别中...');
}

// NOTE: loadFile(), loadFileFromDataUrl(), loadPdfFromDataUrl() removed in v1.4.1
// Replaced by Fast variants (loadFileFast, loadFileFromDataUrlFast, loadPdfFromDataUrlFast)
// which show preview first and run OCR in background for better UX.

// =====================================================
// Fast loading functions — show preview first, OCR in background
// =====================================================

/**
 * Cleanup function called by Rust before closing the window.
 * Clears OCR queues and sets closing flag to prevent new work.
 */
window._tauriCleanup = function() {
  window.__TAURI_CLOSING__ = true;
  _ocrQueue = [];
  console.log('[Cleanup] OCR queue cleared, closing flag set');
};
var _ocrQueue = [];
var _ocrRunning = 0;
var _ocrMaxConcurrent = 2;

function _drainOcrQueue() {
  while (_ocrRunning < _ocrMaxConcurrent && _ocrQueue.length > 0) {
    var task = _ocrQueue.shift();
    _ocrRunning++;
    task().then(function() {
      _ocrRunning--;
      _drainOcrQueue();
    }).catch(function() {
      _ocrRunning--;
      _drainOcrQueue();
    });
  }
}

function applyOcrAsync(fileObj, dataUrl) {
  if (!isTauri || !invoke || window.__TAURI_CLOSING__) return;
  _ocrQueue.push(function() {
    return applyOcr(fileObj, dataUrl).then(function() {
      updateFileItem(fileObj);
      updateAmountSummary();
    }).catch(function(e) {
      console.warn('[OCR] 后台识别失败:', e);
    });
  });
  _drainOcrQueue();
}

/**
 * Background OCR for PDF pages + PDF.js text extraction fallback.
 * Uses OCR queue for throttling. Updates UI incrementally.
 */
function backgroundOcrPdf(results, dataUrl) {
  if (window.__TAURI_CLOSING__) return;
  if (!isTauri || !invoke) {
    // Non-Tauri: try PDF.js text extraction as the only source
    if (dataUrl) {
      tryExtractPdfInfo(dataUrl, results.length).then(function(pdfInfoList) {
        var updated = false;
        for (var k = 0; k < pdfInfoList.length && k < results.length; k++) {
          var pi = pdfInfoList[k];
          var piEffAmt = pi.amountTax > 0 ? pi.amountTax : pi.amountNoTax;
          if (piEffAmt > 0 && results[k].amount <= 0) {
            results[k].amount = piEffAmt;
            results[k].amountTax = pi.amountTax;
            results[k].amountNoTax = pi.amountNoTax;
            updated = true;
          }
          // Skip seller info for tickets
          if (!pi.isTicket) {
            if (pi.sellerName && !results[k].sellerName) { results[k].sellerName = pi.sellerName; updated = true; }
            if (pi.sellerCreditCode && !results[k].sellerCreditCode) { results[k].sellerCreditCode = pi.sellerCreditCode; updated = true; }
          }
          if (pi.isTicket) { results[k]._isTicket = true; }
        }
        if (updated) { renderFileList(); updateAmountSummary(); }
      }).catch(function() {});
    }
    return;
  }

  // Tauri: use throttled OCR queue for pages missing info
  for (var p = 0; p < results.length; p++) {
    // Queue OCR for pages that are missing either amount OR seller info
    if (results[p].amount > 0 && results[p].sellerName && !results[p]._isTicket) continue;
    (function(idx) {
      _ocrQueue.push(function() {
        return applyOcr(results[idx], results[idx].previewUrl).then(function() {
          updateFileItem(results[idx]);
          updateAmountSummary();
        }).catch(function(e) {
          console.warn('[OCR] PDF页后台识别失败:', e);
        });
      });
    })(p);
  }
  _drainOcrQueue();

  // Always try PDF.js text extraction for seller info — OCR on rendered images
  // often misses the seller section, but PDF.js text extraction preserves layout better
  if (dataUrl) {
    tryExtractPdfInfo(dataUrl, results.length).then(function(pdfInfoList) {
      var updated = false;
      for (var k = 0; k < pdfInfoList.length && k < results.length; k++) {
        var pi = pdfInfoList[k];
        var piEffAmt = pi.amountTax > 0 ? pi.amountTax : pi.amountNoTax;
        if (piEffAmt > 0 && results[k].amount <= 0) {
          results[k].amount = piEffAmt;
          results[k].amountTax = pi.amountTax;
          results[k].amountNoTax = pi.amountNoTax;
          updated = true;
        }
        // Skip seller info for tickets
        if (!pi.isTicket && !results[k]._isTicket) {
          if (pi.sellerName && !results[k].sellerName) { results[k].sellerName = pi.sellerName; updated = true; }
          if (pi.sellerCreditCode && !results[k].sellerCreditCode) { results[k].sellerCreditCode = pi.sellerCreditCode; updated = true; }
        }
        if (pi.isTicket) { results[k]._isTicket = true; }
      }
      if (updated) { renderFileList(); updateAmountSummary(); }
    }).catch(function(e) { console.warn('[信息提取] PDF.js提取失败:', e); });
  }
}

/**
 * Incrementally update a single file item's badges in the sidebar
 */
function updateFileItem(fileObj) {
  var idx = S.files.indexOf(fileObj);
  if (idx < 0) return;
  var list = document.getElementById('fileList');
  var items = list.querySelectorAll('.file-item');
  if (!items[idx]) { renderFileList(); return; }
  var f = fileObj;
  var cb = f.copies > 1 ? '<span class="copy-badge">' + f.copies + '份</span>' : '';
  var rb = f.rotation ? '<span class="rot-badge">' + f.rotation + '°</span>' : '';
  var ab = (f.amountTax > 0 || f.amountNoTax > 0) ? '<span class="amt-badge">\u00A5' + (f.amountTax || f.amountNoTax).toFixed(2) + '</span>' : '';
  var sb = f.sellerName ? '<span class="seller-badge" title="' + escHtml(f.sellerCreditCode || '') + '">' + escHtml(f.sellerName.length > 16 ? f.sellerName.substring(0, 16) + '\u2026' : f.sellerName) + '</span>' : '';
  var metaEl = items[idx].querySelector('.file-meta');
  var sellerEl = items[idx].querySelector('.file-seller');
  if (metaEl) metaEl.innerHTML = fmtSize(f.size) + cb + rb + ab;
  if (sellerEl) {
    sellerEl.innerHTML = sb;
    sellerEl.style.display = sb ? '' : 'none';
  }
}

/**
 * Fast load from data URL — show preview immediately, OCR in background
 */
function loadFileFromDataUrlFast(name, dataUrl, size, ext, filePath) {
  return new Promise(function(resolve) {
    var id = 'f' + Date.now() + Math.random().toString(36).slice(2);

    if (ext === 'pdf') {
      if (isTauri && invoke && filePath) {
        invoke('render_pdf_pages', { pdfPath: filePath, dpi: PDF_RENDER_DPI }).then(async function(pages) {
          if (pages && pages.length > 0) {
            var results = [];
            for (var p = 0; p < pages.length; p++) {
              var pg = pages[p];
              var img = new Image(); img.src = pg.imageDataUrl;
              await new Promise(function(r) { img.onload = r; });
              results.push(createFileObj({
                id: id + '_p' + (p + 1),
                name: pages.length > 1 ? name.replace(/\.pdf$/i, '') + '_第' + (p + 1) + '页.pdf' : name,
                size: size, type: 'pdf', previewUrl: pg.imageDataUrl,
                img: img, renderDpi: pg.renderDpi || PDF_RENDER_DPI
              }));
            }
            resolve(results.length === 1 ? results[0] : results);
            // Background: OCR + PDF.js text extraction
            backgroundOcrPdf(results, dataUrl);
            return;
          }
          loadPdfFromDataUrlFast(name, dataUrl, size, id, resolve);
        }).catch(function(err) {
          console.error('[PDF] WinRT rendering failed, falling back to PDF.js:', err);
          loadPdfFromDataUrlFast(name, dataUrl, size, id, resolve);
        });
        return;
      }
      loadPdfFromDataUrlFast(name, dataUrl, size, id, resolve);
    }
    else {
      var img = new Image(); img.src = dataUrl;
      img.onload = function() {
        var result = createFileObj({
          id: id, name: name, size: size, type: ext,
          previewUrl: dataUrl, img: img
        });
        resolve(result);
        // Background OCR
        applyOcrAsync(result, dataUrl);
      };
      img.onerror = function() { toast('图片加载失败: ' + name); resolve(null); };
    }
  });
}

/**
 * Fast load File object (browser mode) — show preview first, OCR in background
 */
function loadFileFast(file) {
  return new Promise(function(resolve) {
    var ext = file.name.split('.').pop().toLowerCase();
    var id = 'f' + Date.now() + Math.random().toString(36).slice(2);

    if (ext === 'pdf') {
      if (!window.pdfjsLib) {
        toast('PDF.js 尚未加载，请稍后重试');
        resolve(null); return;
      }
      var url = URL.createObjectURL(file);
      pdfjsLib.getDocument({ url: url, cMapUrl: CMAP_BASE_URL, cMapPacked: true, standardFontDataUrl: STD_FONT_BASE_URL, disableFontFace: true, useSystemFonts: false }).promise.then(async function(pdf) {
        var results = [];
        for (var p = 1; p <= pdf.numPages; p++) {
          var page = await pdf.getPage(p);
          var vp1 = page.getViewport({ scale: 1.0 });
          var longestSide = Math.max(vp1.width, vp1.height);
          var targetDpi = Math.max(PDF_RENDER_DPI, Math.ceil(MIN_RENDER_PX / longestSide * 72));
          targetDpi = Math.min(targetDpi, 1200);
          var vp = page.getViewport({ scale: targetDpi / 72 });
          var canvas = document.createElement('canvas');
          canvas.width = vp.width; canvas.height = vp.height;
          await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
          var previewUrl = canvas.toDataURL('image/png');
          var img = new Image(); img.src = previewUrl;
          await new Promise(function(r) { img.onload = r; });
          var textContent = await page.getTextContent();
          var info = extractInvoiceInfo(textContent);
          var effAmt = info.amountTax > 0 ? info.amountTax : info.amountNoTax;
          results.push(createFileObj({
            id: id + '_p' + p,
            name: pdf.numPages > 1 ? file.name.replace(/\.pdf$/i, '') + '_第' + p + '页.pdf' : file.name,
            size: file.size, type: 'pdf', previewUrl: previewUrl,
            img: img, amountTax: info.amountTax, amountNoTax: info.amountNoTax,
            amount: effAmt, renderDpi: targetDpi,
            sellerName: info.sellerName, sellerCreditCode: info.sellerCreditCode,
            _ocrText: info._ocrText
          }));
        }
        URL.revokeObjectURL(url);
        resolve(results.length === 1 ? results[0] : results);
        // Background OCR for pages missing info
        backgroundOcrPdf(results, null);
      }).catch(function(err) {
        console.error('PDF load error:', err);
        toast('PDF 加载失败: ' + file.name);
        resolve(null);
      });
    }
    else if (['jpg', 'jpeg', 'png', 'bmp', 'webp', 'tiff', 'tif'].indexOf(ext) >= 0) {
      var reader = new FileReader();
      reader.onload = async function(e) {
        var img = new Image(); img.src = e.target.result;
        await new Promise(function(r) { img.onload = r; });
        var fileObj = createFileObj({
          id: id, name: file.name, size: file.size, type: ext,
          previewUrl: e.target.result, img: img
        });
        resolve(fileObj);
        applyOcrAsync(fileObj, e.target.result);
      };
      reader.onerror = function() { toast('读取失败: ' + file.name); resolve(null); };
      reader.readAsDataURL(file);
    }
    else if (ext === 'ofd') {
      toast('OFD 格式请使用桌面版打开');
      resolve(null);
    }
    else {
      toast('不支持的格式: ' + ext);
      resolve(null);
    }
  });
}

/**
 * Fast PDF.js fallback — render + text extract, then resolve; OCR in background
 */
function loadPdfFromDataUrlFast(name, dataUrl, size, id, resolve) {
  if (!window.pdfjsLib) {
    toast('PDF.js 尚未加载，请稍后重试');
    resolve(null); return;
  }
  var raw = dataUrlToUint8Array(dataUrl);
  pdfjsLib.getDocument({ data: raw, cMapUrl: CMAP_BASE_URL, cMapPacked: true, standardFontDataUrl: STD_FONT_BASE_URL, disableFontFace: true, useSystemFonts: false }).promise.then(async function(pdf) {
    var results = [];
    for (var p = 1; p <= pdf.numPages; p++) {
      var page = await pdf.getPage(p);
      var vp1 = page.getViewport({ scale: 1.0 });
      var longestSide = Math.max(vp1.width, vp1.height);
      var targetDpi = Math.max(PDF_RENDER_DPI, Math.ceil(MIN_RENDER_PX / longestSide * 72));
      targetDpi = Math.min(targetDpi, 1200);
      var vp = page.getViewport({ scale: targetDpi / 72 });
      var canvas = document.createElement('canvas');
      canvas.width = vp.width; canvas.height = vp.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
      var previewUrl = canvas.toDataURL('image/png');
      var img = new Image(); img.src = previewUrl;
      await new Promise(function(r) { img.onload = r; });
      var textContent = await page.getTextContent();
      var info = extractInvoiceInfo(textContent);
      var effAmt = info.amountTax > 0 ? info.amountTax : info.amountNoTax;
      results.push(createFileObj({
        id: id + '_p' + p,
        name: pdf.numPages > 1 ? name.replace(/\.pdf$/i, '') + '_第' + p + '页.pdf' : name,
        size: size, type: 'pdf', previewUrl: previewUrl,
        img: img, amountTax: info.amountTax, amountNoTax: info.amountNoTax,
        amount: effAmt, renderDpi: targetDpi,
        sellerName: info.sellerName, sellerCreditCode: info.sellerCreditCode,
        _ocrText: info._ocrText
      }));
    }
    resolve(results.length === 1 ? results[0] : results);
    // Background OCR for pages missing info
    backgroundOcrPdf(results, dataUrl);
  }).catch(function(err) {
    console.error('PDF load error:', err);
    toast('PDF 加载失败: ' + name);
    resolve(null);
  });
}

// Drag & Drop (browser fallback)
function handleDragOver(e) { e.preventDefault(); e.stopPropagation(); document.getElementById('dropZone').classList.add('drag-over'); }
function handleDragLeave(e) { e.preventDefault(); e.stopPropagation(); document.getElementById('dropZone').classList.remove('drag-over'); }
async function handleDrop(e) {
  e.preventDefault(); e.stopPropagation();
  document.getElementById('dropZone').classList.remove('drag-over');
  if (e.dataTransfer.files && e.dataTransfer.files.length) {
    await processFiles(Array.from(e.dataTransfer.files));
  }
}

// =====================================================
// File list management
// =====================================================
function renderFileList() {
  var list = document.getElementById('fileList');
  var sel = S.files.filter(function(f) { return f.checked; }).length;
  document.getElementById('fileCount').textContent = S.files.length + ' 张，已选 ' + sel;
  var summaryEl = document.getElementById('amountSummary');
  if (!S.files.length) { list.innerHTML = ''; if (summaryEl) summaryEl.style.display = 'none'; updateAmountSummary(); return; }
  if (summaryEl) summaryEl.style.display = 'flex';
  list.innerHTML = S.files.map(function(f, i) {
    var cb = f.copies > 1 ? '<span class="copy-badge">' + f.copies + '份</span>' : '';
    var rb = f.rotation ? '<span class="rot-badge">' + f.rotation + '°</span>' : '';
    var ab = (f.amountTax > 0 || f.amountNoTax > 0) ? '<span class="amt-badge">\u00A5' + (f.amountTax || f.amountNoTax).toFixed(2) + '</span>' : '';
    var sb = f.sellerName ? '<span class="seller-badge" title="' + escHtml(f.sellerCreditCode || '') + '">' + escHtml(f.sellerName.length > 16 ? f.sellerName.substring(0, 16) + '\u2026' : f.sellerName) + '</span>' : '';
    // XSS FIX: escHtml(f.name) in both title and display text
    // XSS FIX: escHtml(f.previewUrl) in img src, escHtml(f.type) in type-badge
    var safePreviewUrl = escHtml(f.previewUrl || '');
    var safeType = escHtml(f.type === 'jpeg' ? 'jpg' : f.type);
    return '<div class="file-item" draggable="true" ondragstart="dStart(event,' + i + ')" ondragover="dOver(event)" ondrop="dDrop(event,' + i + ')" ondblclick="openInvModal(' + i + ')">' +
      '<div class="file-check ' + (f.checked ? 'checked' : '') + '" onclick="togCheck(' + i + ')"></div>' +
      '<div class="file-thumb">' + (f.previewUrl ? '<img src="' + safePreviewUrl + '">' : '\uD83D\uDCC4') + '<div class="type-badge">' + safeType + '</div></div>' +
      '<div class="file-info"><div class="file-name" title="' + escHtml(f.name) + '">' + escHtml(f.name) + '</div><div class="file-meta">' + fmtSize(f.size) + cb + rb + ab + '</div>' + (sb ? '<div class="file-seller">' + sb + '</div>' : '<div class="file-seller" style="display:none"></div>') + '</div>' +
      '<div style="display:flex;gap:2px"><button class="ib" onclick="rotFile(' + i + ')" title="旋转90°">\u21BB</button><button class="ib danger" onclick="rmFile(' + i + ')">\u2715</button></div>' +
    '</div>';
  }).join('');
  updateAmountSummary();
}
function togCheck(i) { S.files[i].checked = !S.files[i].checked; renderFileList(); updatePreview(); }
function selectAll() { S.files.forEach(function(f) { f.checked = true; }); renderFileList(); updatePreview(); }
function deselectAll() { S.files.forEach(function(f) { f.checked = false; }); renderFileList(); updatePreview(); }
function deleteSelected() { if (!S.files.some(function(f) { return f.checked; })) return; S.files = S.files.filter(function(f) { return !f.checked; }); renderFileList(); updatePreview(); updatePrintBtn(); }
function rmFile(i) { S.files.splice(i, 1); renderFileList(); updatePreview(); updatePrintBtn(); }
function rotFile(i) { S.files[i].rotation = (S.files[i].rotation + 90) % 360; renderFileList(); updatePreview(); }
function clearAll() { if (!S.files.length) return; if (!confirm('确认清除所有发票？')) return; S.files = []; renderFileList(); updatePreview(); updatePrintBtn(); }
var dSrc = null;
function dStart(e, i) { dSrc = i; e.dataTransfer.effectAllowed = 'move'; }
function dOver(e) { e.preventDefault(); }
function dDrop(e, i) { e.preventDefault(); if (dSrc === null || dSrc === i) return; var item = S.files.splice(dSrc, 1)[0]; S.files.splice(i, 0, item); dSrc = null; renderFileList(); updatePreview(); }

// Amount statistics
function updateAmountSummary() {
  var el = document.getElementById('amountSummary');
  if (!el) return;
  var checked = S.files.filter(function(f) { return f.checked; });
  var taxTotal = checked.reduce(function(s, f) { return s + (f.amountTax || 0); }, 0);
  var noTaxTotal = checked.reduce(function(s, f) { return s + (f.amountNoTax || 0); }, 0);
  var withAmt = checked.filter(function(f) { return (f.amountTax || f.amountNoTax) > 0; }).length;

  el.style.display = checked.length > 0 ? '' : 'none';
  if (checked.length === 0) return;

  var countHtml = '<span class="amt-count">' + withAmt + '/' + checked.length + ' 张已识别</span>';
  var mode = S.amtMode || 'tax';
  var amtHtml = '';
  if (mode === 'tax') {
    amtHtml = '<span class="amt-total">\u00A5' + taxTotal.toFixed(2) + '</span>';
  } else if (mode === 'notax') {
    amtHtml = '<span class="amt-total">\u00A5' + noTaxTotal.toFixed(2) + '</span>';
  } else {
    amtHtml = '<span class="amt-total" style="font-size:12px;display:flex;flex-direction:column;align-items:flex-end;gap:1px"><span>含税 \u00A5' + taxTotal.toFixed(2) + '</span><span style="font-size:11px;color:var(--text-muted);font-weight:400">不含税 \u00A5' + noTaxTotal.toFixed(2) + '</span></span>';
  }
  var sellers = {}, sellerNames = [], sellerHtml = '';
  checked.forEach(function(f) {
    if (f.sellerName) { var n = f.sellerName.trim(); sellers[n] = (sellers[n] || 0) + 1; }
  });
  sellerNames = Object.keys(sellers);
  if (sellerNames.length > 0) {
    var sellerList = sellerNames.slice(0, 5).map(function(n) {
      var cnt = sellers[n];
      return '<span class="seller-badge" style="margin:0 1px">' + escHtml(n.length > 10 ? n.substring(0, 10) + '\u2026' : n) + (cnt > 1 ? '\u00D7' + cnt : '') + '</span>';
    }).join('');
    if (sellerNames.length > 5) sellerList += ' <span style="font-size:10px;color:var(--text-muted)">+' + (sellerNames.length - 5) + '</span>';
    sellerHtml = '<div style="display:flex;align-items:center;gap:2px;flex-wrap:wrap;font-size:10px;color:var(--text-muted);margin-top:2px;max-width:300px"><span style="white-space:nowrap">' + sellerNames.length + '个销售方</span>' + sellerList + '</div>';
  }
  el.innerHTML = countHtml + amtHtml + sellerHtml;

  var stAmt = document.getElementById('stAmount');
  if (stAmt) {
    var mainTotal = mode === 'notax' ? noTaxTotal : taxTotal;
    stAmt.textContent = mainTotal > 0 ? '\u00A5' + mainTotal.toFixed(2) : '';
  }
}

// Invoice modal
function openInvModal(i) {
  S.editIdx = i; var f = S.files[i];
  var ocrHtml = f._ocrText ? '<details style="margin-top:8px;font-size:11px"><summary style="cursor:pointer;color:var(--text-muted)">\uD83D\uDD0D OCR识别原文</summary><pre style="margin-top:4px;padding:6px 8px;background:var(--surface2);border-radius:4px;max-height:200px;overflow:auto;white-space:pre-wrap;word-break:break-all;font-size:10px;line-height:1.4">' + escHtml(f._ocrText) + '</pre></details>' : '';
  document.getElementById('invModalBody').innerHTML =
    '<div style="font-size:13px;padding:8px 10px;background:var(--surface2);border-radius:6px;margin-bottom:10px">\uD83D\uDCC4 ' + escHtml(f.name) + '</div>' +
    '<div class="row"><label class="lbl">份数</label><div style="display:flex;gap:4px;align-items:center"><button class="btn btn-sm btn-icon" onclick="changeModalCopies(-1)">\u2212</button><input type="number" id="mCopies" value="' + f.copies + '" min="1" max="99" style="width:52px;text-align:center"><button class="btn btn-sm btn-icon" onclick="changeModalCopies(1)">+</button></div></div>' +
    '<div class="row"><label class="lbl">含税价</label><div style="display:flex;gap:4px;align-items:center"><span style="font-size:14px;font-weight:600;color:var(--success)">\u00A5</span><input type="number" id="mAmountTax" value="' + (f.amountTax || '') + '" min="0" step="0.01" placeholder="0.00" style="width:100px;text-align:right"></div></div>' +
    '<div class="row"><label class="lbl">不含税</label><div style="display:flex;gap:4px;align-items:center"><span style="font-size:14px;font-weight:600;color:var(--text-muted)">\u00A5</span><input type="number" id="mAmountNoTax" value="' + (f.amountNoTax || '') + '" min="0" step="0.01" placeholder="0.00" style="width:100px;text-align:right"></div></div>' +
    '<div class="row"><label class="lbl">销售方</label><div style="flex:1;display:flex;gap:4px;align-items:center"><input type="text" id="mSeller" value="' + escHtml(f.sellerName || '') + '" placeholder="自动识别" style="flex:1;font-size:11px;min-width:0"></div></div>' +
    '<div class="row"><label class="lbl">信用代码</label><div style="flex:1;display:flex;gap:4px;align-items:center"><input type="text" id="mCreditCode" value="' + escHtml(f.sellerCreditCode || '') + '" placeholder="自动识别" style="flex:1;font-size:11px;min-width:0;font-family:monospace"></div></div>' +
    '<div class="row"><label class="lbl">旋转</label><div class="ctrl"><select id="mRot"><option value="0" ' + (f.rotation === 0 ? 'selected' : '') + '>不旋转</option><option value="90" ' + (f.rotation === 90 ? 'selected' : '') + '>90\u00B0</option><option value="180" ' + (f.rotation === 180 ? 'selected' : '') + '>180\u00B0</option><option value="270" ' + (f.rotation === 270 ? 'selected' : '') + '>270\u00B0</option></select></div></div>' +
    ocrHtml;
  document.getElementById('invModal').classList.remove('hidden');
}
function changeModalCopies(d) { var e = document.getElementById('mCopies'); e.value = Math.max(1, Math.min(99, parseInt(e.value) + d)); }
function closeInvModal() { document.getElementById('invModal').classList.add('hidden'); }
function confirmInvModal() {
  if (S.editIdx < 0) return;
  var f = S.files[S.editIdx];
  f.copies = Math.max(1, parseInt(document.getElementById('mCopies').value) || 1);
  f.rotation = parseInt(document.getElementById('mRot').value) || 0;
  var at = parseFloat(document.getElementById('mAmountTax').value);
  var an = parseFloat(document.getElementById('mAmountNoTax').value);
  f.amountTax = isNaN(at) || at < 0 ? 0 : Math.round(at * 100) / 100;
  f.amountNoTax = isNaN(an) || an < 0 ? 0 : Math.round(an * 100) / 100;
  f.amount = f.amountTax || f.amountNoTax;
  f.sellerName = document.getElementById('mSeller').value;
  f.sellerCreditCode = document.getElementById('mCreditCode').value;
  closeInvModal(); renderFileList(); updatePreview(); updateAmountSummary();
}

// =====================================================
// Layout / Settings
// =====================================================
function setLayout(c, r, el) {
  S.layout = { cols: c, rows: r };
  document.querySelectorAll('.go').forEach(function(e) { e.classList.remove('active'); });
  if (el && el.classList.contains('go')) el.classList.add('active');
  else {
    document.querySelectorAll('.go').forEach(function(e) {
      if (parseInt(e.dataset.cols) === c && parseInt(e.dataset.rows) === r) e.classList.add('active');
    });
  }
  syncToolbarHighlight(c, r);
  document.getElementById('customRows').value = r;
  document.getElementById('customCols').value = c;
  updatePreview();
}
function quickLayout(c, r) {
  var orient = r > c ? 'portrait' : 'landscape';
  document.getElementById('orientation').value = orient;
  var goEl = null;
  document.querySelectorAll('.go').forEach(function(e) {
    if (parseInt(e.dataset.cols) === c && parseInt(e.dataset.rows) === r) goEl = e;
  });
  setLayout(c, r, goEl);
  document.getElementById('customRows').value = r;
  document.getElementById('customCols').value = c;
}
function toggleFeature(k, btn) {
  S.feat[k] = !S.feat[k];
  btn.classList.toggle('on', S.feat[k]);
  if (k === 'watermark') document.getElementById('wmOpts').style.display = S.feat[k] ? 'block' : 'none';
  if (k === 'trimWhite' && S.feat[k]) processTrim();
  updatePreview();
}
function setLayoutPreset(c, r, orient, el) {
  if (!orient) orient = r > c ? 'portrait' : 'landscape';
  document.getElementById('orientation').value = orient;
  S.layout = { cols: c, rows: r };
  document.querySelectorAll('.go').forEach(function(e) { e.classList.remove('active'); });
  if (el) el.classList.add('active');
  syncToolbarHighlight(c, r);
  document.getElementById('customRows').value = r;
  document.getElementById('customCols').value = c;
  updatePreview();
}
function applyCustomLayout() {
  var r = Math.max(1, Math.min(10, parseInt(document.getElementById('customRows').value) || 1));
  var c = Math.max(1, Math.min(10, parseInt(document.getElementById('customCols').value) || 1));
  document.getElementById('customRows').value = r;
  document.getElementById('customCols').value = c;
  var orient = r > c ? 'portrait' : 'landscape';
  document.getElementById('orientation').value = orient;
  S.layout = { cols: c, rows: r };
  document.querySelectorAll('.go').forEach(function(e) {
    e.classList.remove('active');
    if (parseInt(e.dataset.cols) === c && parseInt(e.dataset.rows) === r) e.classList.add('active');
  });
  syncToolbarHighlight(c, r);
  updatePreview();
}
function showCustomLayoutModal() {
  var r = S.layout.rows, c = S.layout.cols;
  document.getElementById('customRows').value = r;
  document.getElementById('customCols').value = c;
  switchTab('settings', document.querySelectorAll('.sidebar-tab')[1]);
  setTimeout(function() { document.getElementById('customRows').focus(); document.getElementById('customRows').select(); }, 100);
}
function syncToolbarHighlight(c, r) {
  document.querySelectorAll('.ql-btn').forEach(function(e) {
    e.classList.remove('active');
    if (!e.classList.contains('ql-custom') && parseInt(e.dataset.cols) === c && parseInt(e.dataset.rows) === r) {
      e.classList.add('active');
    }
  });
}
function syncLayoutHighlight() {
  var c = S.layout.cols, r = S.layout.rows;
  document.querySelectorAll('.go').forEach(function(e) {
    e.classList.remove('active');
    if (parseInt(e.dataset.cols) === c && parseInt(e.dataset.rows) === r) {
      e.classList.add('active');
    }
  });
  syncToolbarHighlight(c, r);
}
function switchTab(n, el) {
  document.querySelectorAll('.sidebar-tab').forEach(function(t) { t.classList.remove('active'); });
  document.querySelectorAll('.sidebar-panel').forEach(function(p) { p.classList.add('hidden'); });
  el.classList.add('active');
  document.getElementById('panel-' + n).classList.remove('hidden');
}
function onPaperChange() { document.getElementById('customPaperRow').style.display = document.getElementById('paperSize').value === 'custom' ? 'flex' : 'none'; updatePreview(); }
function onFitChange() { document.getElementById('customScaleRow').style.display = document.getElementById('fitMode').value === 'custom' ? 'flex' : 'none'; updatePreview(); }
function setMP(t, b, l, r) {
  [['marginTop', 'marginTopN', t], ['marginBottom', 'marginBottomN', b], ['marginLeft', 'marginLeftN', l], ['marginRight', 'marginRightN', r]].forEach(function(arr) {
    document.getElementById(arr[0]).value = arr[2]; document.getElementById(arr[1]).value = arr[2];
  });
  updatePreview();
}
function changeCopies(d) { var e = document.getElementById('copies'); e.value = Math.max(1, Math.min(99, parseInt(e.value) + d)); updatePreview(); }

// Trim whitespace — now delegates to Rust backend (10-50x faster)
async function processTrim() {
  if (!isTauri || !invoke) {
    toast('白边裁剪需要桌面版');
    return;
  }
  showLoading('裁剪白边...');
  try {
    for (var i = 0; i < S.files.length; i++) {
      var f = S.files[i];
      if (f.previewUrl && !f.trimmedUrl) {
        f.trimmedUrl = await invoke('trim_image', { dataUrl: f.previewUrl });
      }
    }
    hideLoading();
    updatePreview();
    toast('裁剪完成');
  } catch (err) {
    hideLoading();
    console.error('[Trim] 裁剪失败:', err);
    toast('裁剪失败: ' + String(err));
  }
}

// =====================================================
// Get settings
// =====================================================
function getSettings() {
  var ps = document.getElementById('paperSize').value;
  var pw, ph;
  if (ps === 'custom') { pw = parseFloat(document.getElementById('customW').value) || 210; ph = parseFloat(document.getElementById('customH').value) || 297; }
  else { var p = PAPER[ps] || PAPER.A4; pw = p.w; ph = p.h; }
  if (document.getElementById('orientation').value === 'landscape') { var tmp = pw; pw = ph; ph = tmp; }
  return {
    paperW: pw, paperH: ph, cols: S.layout.cols, rows: S.layout.rows,
    marginTop: parseFloat(document.getElementById('marginTop').value),
    marginBottom: parseFloat(document.getElementById('marginBottom').value),
    marginLeft: parseFloat(document.getElementById('marginLeft').value),
    marginRight: parseFloat(document.getElementById('marginRight').value),
    gapH: parseFloat(document.getElementById('gapH').value),
    gapV: parseFloat(document.getElementById('gapV').value),
    fitMode: document.getElementById('fitMode').value,
    customScale: parseFloat(document.getElementById('customScale').value) / 100,
    colorMode: document.getElementById('colorMode').value,
    globalRotation: document.getElementById('globalRotation').value,
    cutline: S.feat.cutline, number: S.feat.number, border: S.feat.border,
    borderWidth: 1, borderColor: '#000000', trimWhite: S.feat.trimWhite,
    watermark: S.feat.watermark,
    watermarkText: document.getElementById('wmText').value,
    watermarkOpacity: parseFloat(document.getElementById('wmOpacity').value) / 100,
    watermarkColor: document.getElementById('wmColor').value,
    watermarkAngle: parseFloat(document.getElementById('wmAngle').value),
    watermarkSize: parseFloat(document.getElementById('wmSize').value),
    pageNum: S.feat.pageNum, printDate: S.feat.printDate,
    copies: parseInt(document.getElementById('copies').value) || 1,
    collate: S.feat.collate, duplex: S.feat.duplex,
    printerName: document.getElementById('printerSel').value || null
  };
}

function getActiveFiles() {
  var files = S.files.filter(function(f) { return f.checked; });
  if (document.getElementById('pageOrder').value === 'reverse') files = files.slice().reverse();
  var exp = [];
  files.forEach(function(f) { for (var c = 0; c < Math.max(1, f.copies); c++) exp.push(f); });
  return exp;
}

function buildPages(files, settings) {
  var perPage = settings.cols * settings.rows;
  var pages = [];
  for (var i = 0; i < files.length; i += perPage) pages.push(files.slice(i, i + perPage));
  return pages;
}

// =====================================================
// Preview & Navigation
// =====================================================
function updatePreview() {
  var files = getActiveFiles();
  document.getElementById('stFiles').textContent = S.files.filter(function(f) { return f.checked; }).length + ' 张';
  document.getElementById('stLayout').textContent = S.layout.rows + '\u00D7' + S.layout.cols;
  var ps = document.getElementById('paperSize').value;
  document.getElementById('stPaper').textContent = ps + ' ' + (document.getElementById('orientation').value === 'portrait' ? '纵' : '横');

  if (!files.length) {
    document.getElementById('emptyState').style.display = 'flex';
    document.getElementById('previewPages').style.display = 'none';
    document.getElementById('pageNav').style.display = 'none';
    document.getElementById('pageInfo').textContent = '\u2014 / \u2014';
    document.getElementById('prevBtn').disabled = true; document.getElementById('nextBtn').disabled = true;
    document.getElementById('stPages').textContent = '0 页'; return;
  }
  var settings = getSettings();
  var pages = buildPages(files, settings);
  S.totalPages = pages.length;
  S.currentPage = Math.max(0, Math.min(S.currentPage, pages.length - 1));
  document.getElementById('stPages').textContent = pages.length + ' 页';
  renderPage(pages[S.currentPage], S.currentPage, pages.length, settings);
  updatePageDots(pages.length);
}

function updatePageDots(t) { var d = document.getElementById('pageDots'); if (t <= 1) { d.innerHTML = ''; return; } var m = Math.min(t, 12); d.innerHTML = Array.from({ length: m }, function(_, i) { return '<div class="page-dot ' + (i === S.currentPage ? 'active' : '') + '" onclick="gotoPage(' + i + ')"></div>'; }).join(''); }
function prevPage() { if (S.currentPage > 0) { S.currentPage--; updatePreview(); } }
function nextPage() { if (S.currentPage < S.totalPages - 1) { S.currentPage++; updatePreview(); } }
function gotoPage(i) { S.currentPage = i; updatePreview(); }
function getFitZoom() {
  var wrap = document.getElementById('previewWrap');
  if (!wrap) return 100;
  var ps = document.getElementById('paperSize').value;
  var pw, ph;
  if (ps === 'custom') { pw = parseFloat(document.getElementById('customW').value) || 210; ph = parseFloat(document.getElementById('customH').value) || 297; }
  else { var p = PAPER[ps] || PAPER.A4; pw = p.w; ph = p.h; }
  if (document.getElementById('orientation').value === 'landscape') { var tmp = pw; pw = ph; ph = tmp; }
  var fitScale = Math.min((wrap.clientWidth - 40) / (pw * MM2PX), (wrap.clientHeight - 40) / (ph * MM2PX), 1.2);
  return Math.round(fitScale * 100);
}
function updateZoomDisplay() {
  var label = document.getElementById('zoomLabel');
  if (!label) return;
  label.textContent = S.viewZoom === 0 ? '自适应' : S.viewZoom + '%';
}
function changeZoom(d) {
  var cur = S.viewZoom === 0 ? getFitZoom() : S.viewZoom;
  var newVal = Math.max(10, Math.min(500, cur + d));
  if (newVal === cur) return;
  S.viewZoom = newVal;
  updateZoomDisplay();
  updatePreview();
}
function setZoom(v) {
  if (v === 'fit' || v === 0) { S.viewZoom = 0; }
  else { S.viewZoom = Math.max(10, Math.min(500, parseInt(v) || 100)); }
  updateZoomDisplay();
  updatePreview();
  document.getElementById('zoomMenu').classList.add('hidden');
}
function toggleZoomMenu() {
  document.getElementById('zoomMenu').classList.toggle('hidden');
}
document.addEventListener('click', function(e) {
  if (!e.target.closest('.zoom-ctrl')) {
    var m = document.getElementById('zoomMenu');
    if (m) m.classList.add('hidden');
  }
});
function updatePrintBtn() { document.getElementById('printBtn').disabled = !S.files.some(function(f) { return f.checked; }); }

// =====================================================
// Save settings & Preferences
// =====================================================
function togglePref(k, btn) {
  S.feat[k] = !S.feat[k];
  btn.classList.toggle('on', S.feat[k]);
}

function getSaveDir() {
  try { return localStorage.getItem('fapiao-save-dir') || ''; } catch(e) { return ''; }
}
function setSaveDir(dir) {
  try { localStorage.setItem('fapiao-save-dir', dir); } catch(e) {}
  document.getElementById('saveDir').value = dir;
}
async function pickSaveDir() {
  if (isTauri && invoke) {
    try {
      var result = await invoke('plugin:dialog|open', {
        options: { directory: true, title: '选择PDF保存目录' }
      });
      if (result) { setSaveDir(result); toast('保存目录已设置'); }
    } catch(e) { toast('选择目录失败: ' + String(e)); }
  }
}
function clearSaveDir() { setSaveDir(''); toast('已清除保存目录'); }

async function verifyInvoice() {
  var url = 'https://inv-veri.chinatax.gov.cn/';
  if (isTauri && invoke) {
    try { await invoke('open_url', { url: url }); } catch(e) { toast('打开查验网站失败: ' + String(e)); }
  } else { window.open(url, '_blank'); }
}

function applyTheme() {
  var theme = document.getElementById('themeMode').value;
  if (theme === 'dark') { document.documentElement.classList.add('dark'); }
  else { document.documentElement.classList.remove('dark'); }
  try { localStorage.setItem('fapiao-theme', theme); } catch(e) {}
}

function exportSettings() {
  var data = { layout: S.layout, feat: S.feat, paperSize: document.getElementById('paperSize').value, orientation: document.getElementById('orientation').value, copies: document.getElementById('copies').value, colorMode: document.getElementById('colorMode').value, printMode: document.getElementById('printMode').value, saveDir: getSaveDir() };
  var blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = '发票打印设置.json'; a.click();
  toast('设置已导出');
}

function resetSettings() {
  if (!confirm('确认恢复所有默认设置？')) return;
  S.layout = { cols: 1, rows: 1 };
  S.feat = { cutline: true, number: false, border: false, trimWhite: false, watermark: false, collate: true, duplex: false, pageNum: false, printDate: false, confirmPrint: true, autoOpenPdf: true };
  S.viewZoom = 0;
  document.getElementById('paperSize').value = 'A4';
  document.getElementById('orientation').value = 'landscape';
  document.getElementById('customRows').value = 1;
  document.getElementById('customCols').value = 1;
  document.getElementById('marginTop').value = 5; document.getElementById('marginTopN').value = 5;
  document.getElementById('marginBottom').value = 5; document.getElementById('marginBottomN').value = 5;
  document.getElementById('marginLeft').value = 5; document.getElementById('marginLeftN').value = 5;
  document.getElementById('marginRight').value = 5; document.getElementById('marginRightN').value = 5;
  document.getElementById('gapH').value = 3; document.getElementById('gapHN').value = 3;
  document.getElementById('gapV').value = 3; document.getElementById('gapVN').value = 3;
  document.getElementById('fitMode').value = 'fit';
  document.getElementById('globalRotation').value = '0';
  document.getElementById('copies').value = 1;
  document.getElementById('colorMode').value = 'color';
  document.getElementById('pageOrder').value = 'normal';
  document.getElementById('customPaperRow').style.display = 'none';
  document.getElementById('customScaleRow').style.display = 'none';
  document.getElementById('wmOpts').style.display = 'none';
  updateZoomDisplay();
  document.getElementById('toggleCutline').classList.add('on');
  document.getElementById('toggleNumber').classList.remove('on');
  document.getElementById('toggleBorder').classList.remove('on');
  document.getElementById('toggleTrimWhite').classList.remove('on');
  document.getElementById('toggleWatermark').classList.remove('on');
  document.getElementById('toggleCollate').classList.add('on');
  document.getElementById('toggleDuplex').classList.remove('on');
  document.getElementById('togglePageNum').classList.remove('on');
  document.getElementById('toggleDate').classList.remove('on');
  document.getElementById('toggleConfirm').classList.add('on');
  document.getElementById('toggleAutoOpenPdf').classList.add('on');
  document.getElementById('printMode').value = 'dialog';
  document.getElementById('themeMode').value = 'light';
  document.documentElement.classList.remove('dark');
  try { localStorage.removeItem('fapiao-theme'); } catch(e) {}
  try { localStorage.removeItem('fapiao-save-dir'); } catch(e) {}
  try { localStorage.removeItem('fapiao-amt-mode'); } catch(e) {}
  document.getElementById('saveDir').value = '';
  document.getElementById('amtMode').value = 'tax';
  S.amtMode = 'tax';
  syncLayoutHighlight();
  updatePreview();
  toast('已恢复默认设置');
}

// =====================================================
// Keyboard shortcuts
// =====================================================
document.addEventListener('keydown', function(e) {
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;
  if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') { e.preventDefault(); prevPage(); }
  if (e.key === 'ArrowRight' || e.key === 'ArrowDown') { e.preventDefault(); nextPage(); }
  if ((e.ctrlKey || e.metaKey) && e.key === 'p') { e.preventDefault(); doPrint(); }
  if ((e.ctrlKey || e.metaKey) && e.key === 'o') { e.preventDefault(); triggerUpload(); }
  if ((e.ctrlKey || e.metaKey) && (e.key === '=' || e.key === '+')) { e.preventDefault(); changeZoom(5); }
  if ((e.ctrlKey || e.metaKey) && e.key === '-') { e.preventDefault(); changeZoom(-5); }
  if ((e.ctrlKey || e.metaKey) && e.key === '0') { e.preventDefault(); setZoom('fit'); }
});

// Ctrl+Wheel zoom
document.getElementById('previewWrap').addEventListener('wheel', function(e) {
  if (!e.ctrlKey) return;
  e.preventDefault();
  var step = 5;
  var curZoom = S.viewZoom === 0 ? getFitZoom() : S.viewZoom;
  var delta = e.deltaY > 0 ? -step : step;
  if (curZoom > 200) delta = delta * 2;
  var newZoom = Math.max(10, Math.min(500, curZoom + delta));
  if (newZoom === curZoom) return;

  var oldScale = curZoom / 100;
  var newScale = newZoom / 100;

  var container = document.querySelector('.preview-container');
  var logicalX = 0, logicalY = 0;
  if (container) {
    var cRect = container.getBoundingClientRect();
    logicalX = (e.clientX - cRect.left) / oldScale;
    logicalY = (e.clientY - cRect.top) / oldScale;
  }

  S.viewZoom = newZoom;
  updateZoomDisplay();
  updatePreview();

  var newContainer = document.querySelector('.preview-container');
  if (newContainer) {
    var ncRect = newContainer.getBoundingClientRect();
    var dx = (ncRect.left + logicalX * newScale) - e.clientX;
    var dy = (ncRect.top + logicalY * newScale) - e.clientY;
    var wrap = document.getElementById('previewWrap');
    wrap.scrollLeft += dx;
    wrap.scrollTop += dy;
  }
}, { passive: false });

// Double-click to reset zoom
document.getElementById('previewWrap').addEventListener('dblclick', function() {
  if (S.viewZoom !== 0) { setZoom('fit'); }
});

// Global drag & drop (browser fallback)
document.body.addEventListener('dragover', function(e) { e.preventDefault(); });
document.body.addEventListener('drop', function(e) { e.preventDefault(); if (e.dataTransfer.files.length) processFiles(Array.from(e.dataTransfer.files)); });
window.addEventListener('resize', function() { if (S.files.length) updatePreview(); });

// Tauri drag & drop — Rust calls window._tauriFileDrop(paths) via eval()
window._tauriFileDrop = function(paths) {
  if (!Array.isArray(paths) || paths.length === 0) return;
  (async function() {
    try {
      showLoading('读取 ' + paths.length + ' 个文件...');
      var fileDataList = await invoke('open_invoice_files', { paths: paths });
      hideLoading();
      if (fileDataList && fileDataList.length > 0) {
        await processFileDataList(fileDataList);
      }
    } catch(err) {
      hideLoading();
      toast('拖放文件读取失败: ' + String(err));
    }
  })();
};

// =====================================================
// Load PDF.js — local first, CDN fallback
// =====================================================
(function() {
  var s = document.createElement('script');
  s.onerror = function() {
    console.warn('Local PDF.js not found, trying CDN...');
    var s2 = document.createElement('script');
    s2.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    s2.onload = function() {
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      // CDN fallback also means CMap/standard_fonts must come from CDN
      CMAP_BASE_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/';
      STD_FONT_BASE_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/standard_fonts/';
      console.log('PDF.js loaded from CDN');
    };
    s2.onerror = function() { console.error('PDF.js failed to load'); };
    document.head.appendChild(s2);
  };
  s.onload = function() {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'pdf.worker.min.js';
    console.log('PDF.js loaded from local');
  };
  s.src = 'pdf.min.js';
  document.head.appendChild(s);
})();

// Auto-refresh printers in Tauri — delayed to avoid blocking startup
if (isTauri) setTimeout(function() { refreshPrinters(); }, 800);

// =====================================================
// DPI Runtime Validation — verify frontend matches Rust
// =====================================================
if (isTauri && invoke) {
  invoke('get_config').then(function(config) {
    if (config && config.renderDpi && config.renderDpi !== PDF_RENDER_DPI) {
      console.error('[DPI] 前后端 DPI 不一致！前端=' + PDF_RENDER_DPI + ', Rust=' + config.renderDpi + '，请检查代码');
      toast('警告：渲染DPI配置不一致，打印质量可能受影响', 5000);
    } else if (config && config.renderDpi) {
      console.log('[DPI] 前后端 DPI 一致: ' + config.renderDpi);
    }
  }).catch(function() {
    // get_config command not available in older versions — skip silently
  });
}

// =====================================================
// Initialization — restore saved preferences
// =====================================================
(function() {
  try {
    var saved = localStorage.getItem('fapiao-theme');
    if (saved === 'dark') {
      document.getElementById('themeMode').value = 'dark';
      document.documentElement.classList.add('dark');
    }
  } catch(e) {}
})();

document.getElementById('orientation').value = 'landscape';

(function() {
  try {
    var dir = localStorage.getItem('fapiao-save-dir') || '';
    document.getElementById('saveDir').value = dir;
  } catch(e) {}
})();

(function() {
  try {
    var m = localStorage.getItem('fapiao-amt-mode');
    if (m && (m === 'tax' || m === 'notax' || m === 'both')) {
      S.amtMode = m;
      document.getElementById('amtMode').value = m;
    }
  } catch(e) {}
})();

(function() {
  try {
    var pm = localStorage.getItem('fapiao-print-mode');
    if (pm && (pm === 'dialog' || pm === 'direct')) {
      document.getElementById('printMode').value = pm;
    }
  } catch(e) {}
})();

// =====================================================
// Remove splash screen after everything is loaded
// =====================================================
(function() {
  function removeSplash() {
    var splash = document.getElementById('splash');
    if (splash) {
      splash.classList.add('hide');
      setTimeout(function() { splash.remove(); }, 350);
    }
    // Tell Rust to show the window now that content is rendered (prevents white flash)
    if (isTauri && invoke) {
      try { invoke('show_window'); } catch(e) {}
    }
  }
  // Remove splash after a minimum display time (prevents flash) or when DOM is ready
  if (document.readyState === 'complete') {
    setTimeout(removeSplash, 300);
  } else {
    window.addEventListener('load', function() { setTimeout(removeSplash, 300); });
    // Fallback: remove after 2s no matter what
    setTimeout(removeSplash, 2000);
  }
})();
