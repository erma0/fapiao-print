// =====================================================
// 发票批量打印工具 — 主入口
// v1.8.2 — 进程退出修复 + CropBox优先 + 加载进度优化
// =====================================================

// Detect Tauri — use var to avoid conflict with Tauri's injected scripts
var isTauri = window.__TAURI_INTERNALS__ !== undefined;
var invoke  = isTauri ? window.__TAURI_INTERNALS__.invoke : null;
var hasOcr  = false; // Set to true at startup if OCR feature is available
var APP_VERSION = ''; // Filled at startup from Rust get_app_version()

// =====================================================
// Constants
// =====================================================
var PAPER = { A4:{w:210,h:297}, A5:{w:148,h:210}, B5:{w:176,h:250}, letter:{w:216,h:279}, legal:{w:216,h:356} };
var MM2PX = 96 / 25.4;
var PDF_RENDER_DPI = 300;  // Must match Rust RENDER_DPI — validated at startup via get_config
var MIN_RENDER_PX = 3508;  // A4 long side at 300 DPI — minimum rendered pixels
var WHITE_THRESHOLD = 245; // Pixel value threshold for white-edge trimming

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
    autoOpenPdf: true,
    ocrEnabled: false
  }
};

// Track newly added file IDs for entrance animation
var _newFileIds = {};

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
    taxAmount: opts.taxAmount || 0,
    img: opts.img || null,
    // Original dimensions: prefer explicit ow/oh (from Rust FileData.origW/origH for thumbnails),
    // fall back to img.naturalWidth/naturalHeight (full-size images and rendered PDF pages).
    ow: opts.ow || (opts.img ? opts.img.naturalWidth : 0),
    oh: opts.oh || (opts.img ? opts.img.naturalHeight : 0),
    renderDpi: opts.renderDpi || PDF_RENDER_DPI,
    sellerName: opts.sellerName || '',
    sellerCreditCode: opts.sellerCreditCode || '',
    invoiceNo: opts.invoiceNo || '',
    invoiceDate: opts.invoiceDate || '',
    buyerName: opts.buyerName || '',
    buyerCreditCode: opts.buyerCreditCode || '',
    _ocrText: opts._ocrText || '',
    _isTicket: opts._isTicket || false,
    _loading: opts._loading || false,
    _ocrPending: false,
    // Disk path for the original file (when available).
    // Used by Rust to read bytes directly, skipping base64 encode/decode.
    _filePath: opts.filePath || '',
    // PDF source info for ocr_pdf_page command (zero IPC round-trip OCR).
    // Set when this fileObj represents a PDF page rendered via render_pdf_pages.
    _pdfPath: opts.pdfPath || '',
    _pdfPageIdx: opts.pdfPageIdx != null ? opts.pdfPageIdx : -1
  };
}

// =====================================================
// Helpers
// =====================================================
var toastT = null;
function toast(msg, dur) { dur = dur || 2500; var e = document.getElementById('toast'); e.textContent = msg; e.classList.add('show'); clearTimeout(toastT); if (dur > 0) toastT = setTimeout(function() { e.classList.remove('show'); }, dur); else clearTimeout(toastT); }
function toastHtml(msg, dur) { dur = dur || 2500; var e = document.getElementById('toast'); e.innerHTML = msg; e.classList.add('show'); clearTimeout(toastT); if (dur > 0) toastT = setTimeout(function() { e.classList.remove('show'); }, dur); else clearTimeout(toastT); }
function toastLoading(msg) { _ocrToastActive = true; toastHtml('<span class="toast-spinner"></span>' + msg, 0); }
function toastDone(msg) { toast(msg, 2500); }
function hideToast() { var e = document.getElementById('toast'); e.classList.remove('show'); clearTimeout(toastT); }
function syncSlider(s, n) { document.getElementById(n).value = s.value; }
function syncRange(n, s) { document.getElementById(s).value = n.value; }
function showLoading(t) { document.getElementById('loadingText').textContent = t || '处理中...'; document.getElementById('loadingProgress').classList.add('hidden'); document.getElementById('loadingDetail').classList.add('hidden'); document.getElementById('loading').classList.remove('hidden'); }
function hideLoading() { document.getElementById('loading').classList.add('hidden'); document.getElementById('loadingProgress').classList.add('hidden'); document.getElementById('loadingDetail').classList.add('hidden'); }
function updateLoadingProgress(phase, current, total) {
  var pct = total > 0 ? Math.round(current / total * 100) : 0;
  var bar = document.getElementById('loadingBar');
  var prog = document.getElementById('loadingProgress');
  var detail = document.getElementById('loadingDetail');
  var text = document.getElementById('loadingText');
  if (bar) bar.style.width = pct + '%';
  if (prog) prog.classList.remove('hidden');
  if (detail) {
    if (phase === 'build') {
      detail.textContent = current + ' / ' + total + ' 页';
      if (text) text.textContent = '正在排版...';
    } else if (phase === 'save') {
      detail.textContent = '';
      if (text) text.textContent = '正在写入PDF...';
    } else {
      detail.textContent = current + ' / ' + total;
      if (text) text.textContent = '正在处理...';
    }
    if (detail.textContent) detail.classList.remove('hidden'); else detail.classList.add('hidden');
  }
}
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

// Downsample a data URL image for faster OCR IPC transfer.
// Renders to a canvas at max `maxDim` pixels on the longest side, exports as JPEG.
// Returns a Promise<string> with the downsampled data URL.
function downsampleForOcr(dataUrl, maxDim) {
  return new Promise(function(resolve) {
    if (!dataUrl || dataUrl.length < 100000) { resolve(dataUrl); return; }
    try {
      var img = new Image();
      img.onload = function() {
        var longest = Math.max(img.naturalWidth, img.naturalHeight);
        if (longest <= maxDim) { resolve(dataUrl); return; }
        var scale = maxDim / longest;
        var w = Math.round(img.naturalWidth * scale);
        var h = Math.round(img.naturalHeight * scale);
        var canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        var ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, w, h);
        // Quality 0.92: higher quality preserves text edges for better OCR accuracy
        resolve(canvas.toDataURL('image/jpeg', 0.92));
      };
      img.onerror = function() { resolve(dataUrl); };
      img.src = dataUrl;
    } catch(e) { resolve(dataUrl); }
  });
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

// Process FileData array from Rust backend — instant placeholders, then load in parallel + render sequentially
async function processFileDataList(fileDataList) {
  var total = fileDataList.length;
  var completed = 0;
  var added = 0;
  _loadingBatchActive = true;

  // 1. Create placeholder entries immediately for instant visual feedback
  fileDataList.forEach(function(fd) {
    var ph = createFileObj({
      name: fd.name,
      size: fd.size,
      type: fd.ext,
      _loading: true
    });
    ph._placeholderKey = ph.id;
    fd._phKey = ph._placeholderKey;
    S.files.push(ph);
    _newFileIds[ph.id] = true;
  });

  // Render placeholders immediately — user sees skeleton items right away
  renderFileList(); updatePreview(); updatePrintBtn();

  // Show "加载中" toast immediately with spinner
  toastLoading('加载中 0/' + total);

  // Count how many files will need OCR (for batch tracking)
  var ocrEligibleCount = S.feat.ocrEnabled ? fileDataList.length : 0;
  if (ocrEligibleCount >= 1) {
    _ocrBatchTotal = ocrEligibleCount;
  }

  // 2. Start all loads in parallel (efficient for PDF IPC), then process results sequentially for incremental rendering
  var loadPromises = fileDataList.map(function(fd) {
    return loadFileFromDataUrlFast(fd).catch(function(err) {
      console.error('Load file error:', fd.name, err);
      return null;
    });
  });

  for (var fdIdx = 0; fdIdx < fileDataList.length; fdIdx++) {
    var fd = fileDataList[fdIdx];
    var r = await loadPromises[fdIdx];
    completed++;

    // Find placeholder by key
    var phIdx = -1;
    for (var i = 0; i < S.files.length; i++) {
      if (S.files[i]._placeholderKey === fd._phKey) { phIdx = i; break; }
    }

    if (phIdx >= 0 && r) {
      var items = Array.isArray(r) ? r : [r];
      items.forEach(function(it) { _newFileIds[it.id] = true; });
      S.files.splice.apply(S.files, [phIdx, 1].concat(items));
      added += items.length;
    } else if (phIdx >= 0) {
      // Remove placeholder for failed/empty file
      S.files.splice(phIdx, 1);
    }

    // Update loading progress toast
    var ocrRemaining = _ocrQueue.length + _ocrRunning;
    var isLast = (completed >= total);
    if (isLast) {
      // Last file loaded — check if OCR still running
      if (ocrRemaining > 0 && S.feat.ocrEnabled) {
        var ocrDone2 = _ocrBatchTotal > 0 ? _ocrBatchTotal - ocrRemaining : 0;
        toastLoading('加载完成，识别中 ' + ocrDone2 + '/' + _ocrBatchTotal);
      }
      // else: will be handled after the loop (toastDone)
    } else {
      if (ocrRemaining > 0 && S.feat.ocrEnabled) {
        var ocrDone = _ocrBatchTotal > 0 ? _ocrBatchTotal - ocrRemaining : 0;
        toastLoading('加载中 ' + completed + '/' + total + '，识别中 ' + ocrDone + '/' + _ocrBatchTotal);
      } else {
        toastLoading('加载中 ' + completed + '/' + total);
      }
    }

    renderFileList(); updatePreview(); updatePrintBtn();

    // Yield to browser for painting — ensures user sees each file appear incrementally
    await nextFrame();
  }

  // Loading batch complete
  _loadingBatchActive = false;

  // If no OCR queued, dismiss toast now
  if (_ocrQueue.length === 0 && _ocrRunning === 0) {
    _ocrToastActive = false;
    _ocrBatchTotal = 0;
    _ocrBatchAddedCount = 0;
    toastDone('已加载 ' + added + ' 张发票');
  } else {
    // OCR still running — save added count for _drainOcrQueue's final toast
    _ocrBatchAddedCount = added;
  }
}

// Process an array of File objects (browser fallback) — instant placeholders, then load in parallel + render sequentially
async function processFiles(files) {
  var total = files.length;
  var completed = 0;
  var added = 0;
  _loadingBatchActive = true;

  // Create placeholder entries immediately
  files.forEach(function(file) {
    var ext = file.name.split('.').pop().toLowerCase();
    var ph = createFileObj({
      name: file.name,
      size: file.size,
      type: ext,
      _loading: true
    });
    ph._placeholderKey = ph.id;
    file._phKey = ph._placeholderKey;
    S.files.push(ph);
    _newFileIds[ph.id] = true;
  });
  renderFileList(); updatePreview(); updatePrintBtn();

  // Show "加载中" toast immediately with spinner
  toastLoading('加载中 0/' + total);

  // Count how many files will need OCR (for batch tracking)
  var ocrEligibleCount = S.feat.ocrEnabled ? files.length : 0;
  if (ocrEligibleCount >= 1) {
    _ocrBatchTotal = ocrEligibleCount;
  }

  // Start all loads in parallel (efficient for FileReader I/O), then process results sequentially for incremental rendering
  var loadPromises = files.map(function(file) {
    return loadFileFast(file).catch(function(err) {
      console.error('Load file error:', file.name, err);
      return null;
    });
  });

  for (var fIdx = 0; fIdx < files.length; fIdx++) {
    var file = files[fIdx];
    var r = await loadPromises[fIdx];
    completed++;

    var phIdx = -1;
    for (var i = 0; i < S.files.length; i++) {
      if (S.files[i]._placeholderKey === file._phKey) { phIdx = i; break; }
    }

    if (phIdx >= 0 && r) {
      var items = Array.isArray(r) ? r : [r];
      items.forEach(function(it) { _newFileIds[it.id] = true; });
      S.files.splice.apply(S.files, [phIdx, 1].concat(items));
      added += items.length;
    } else if (phIdx >= 0) {
      S.files.splice(phIdx, 1);
    }

    // Update loading progress toast
    var ocrRemaining = _ocrQueue.length + _ocrRunning;
    var isLast = (completed >= total);
    if (isLast) {
      // Last file loaded — check if OCR still running
      if (ocrRemaining > 0 && S.feat.ocrEnabled) {
        var ocrDone2 = _ocrBatchTotal > 0 ? _ocrBatchTotal - ocrRemaining : 0;
        toastLoading('加载完成，识别中 ' + ocrDone2 + '/' + _ocrBatchTotal);
      }
      // else: will be handled after the loop (toastDone)
    } else {
      if (ocrRemaining > 0 && S.feat.ocrEnabled) {
        var ocrDone = _ocrBatchTotal > 0 ? _ocrBatchTotal - ocrRemaining : 0;
        toastLoading('加载中 ' + completed + '/' + total + '，识别中 ' + ocrDone + '/' + _ocrBatchTotal);
      } else {
        toastLoading('加载中 ' + completed + '/' + total);
      }
    }

    renderFileList(); updatePreview(); updatePrintBtn();

    // Yield to browser for painting — ensures user sees each file appear incrementally
    await nextFrame();
  }

  // Loading batch complete
  _loadingBatchActive = false;

  // If no OCR queued, dismiss toast now
  if (_ocrQueue.length === 0 && _ocrRunning === 0) {
    _ocrToastActive = false;
    _ocrBatchTotal = 0;
    _ocrBatchAddedCount = 0;
    toastDone('已加载 ' + added + ' 张发票');
  } else {
    // OCR still running — save added count for _drainOcrQueue's final toast
    _ocrBatchAddedCount = added;
  }
}

// NOTE: loadFile(), loadFileFromDataUrl(), loadPdfFromDataUrl(), loadPdfFromDataUrlFast() removed.
// PDF.js removed in v1.7.1 — all PDF rendering via WinRT native, all text extraction via PP-OCRv5.

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
  _ocrRunning = 0;
  _ocrToastActive = false;
  _ocrFromButton = false;
  _loadingBatchActive = false;
  console.log('[Cleanup] OCR queue cleared, closing flag set');
};
var _loadingBatchActive = false; // True while batch loading is in progress — prevents OCR from dismissing toast
var _ocrQueue = [];
var _ocrRunning = 0;
var _ocrMaxConcurrent = 1; // OCR引擎是Mutex，同时只有1个请求能执行
var _ocrToastActive = false; // track if "识别中" toast is showing
var _ocrFromButton = false;  // true = OCR triggered by single-file button click (show per-file result toast)
var _ocrBatchTotal = 0;     // Total files in current batch (for progress display)
var _ocrBatchAddedCount = 0; // Total added files in current loading batch (for final toast message)
var _lastPdfPath = null;   // Path of last generated/saved PDF (for print cache)
var _pdfDirty = true;      // Whether PDF content has changed since last generation

/** Yield to browser for reliable painting — double rAF ensures at least one frame is painted */
function nextFrame() { return new Promise(function(r) { requestAnimationFrame(function() { requestAnimationFrame(r); }); }); }
var _activeFileIdx = -1;   // Index of currently active/highlighted file in sidebar

function _onOcrTaskDone() {
  _ocrRunning--;
  var remaining = _ocrQueue.length + _ocrRunning;
  // Only update OCR toast when batch loading is NOT active (loading loop handles its own toast)
  if (remaining > 0 && _ocrToastActive && !_loadingBatchActive) {
    var done = _ocrBatchTotal > 0 ? _ocrBatchTotal - remaining : 0;
    if (_ocrBatchTotal > 0) {
      toastLoading('识别中 ' + done + '/' + _ocrBatchTotal);
    } else {
      toastLoading('识别中，剩余 ' + remaining + ' 张');
    }
  }
  updateOcrAllBtn();
  if (!window.__TAURI_CLOSING__) _drainOcrQueue();
}

function _drainOcrQueue() {
  if (window.__TAURI_CLOSING__) return;
  while (_ocrRunning < _ocrMaxConcurrent && _ocrQueue.length > 0) {
    var task = _ocrQueue.shift();
    _ocrRunning++;
    task().then(_onOcrTaskDone).catch(_onOcrTaskDone);
  }
  // All OCR done — dismiss loading toast (but NOT if batch loading is still active)
  if (_ocrQueue.length === 0 && _ocrRunning === 0 && _ocrToastActive && !_loadingBatchActive) {
    _ocrToastActive = false;
    var wasBatchTotal = _ocrBatchTotal;
    var wasAddedCount = _ocrBatchAddedCount;
    var wasFromButton = _ocrFromButton;
    _ocrBatchTotal = 0;
    _ocrBatchAddedCount = 0;
    _ocrFromButton = false;
    updateOcrAllBtn();
    // Single-file OCR from button click shows its own result toast in applyOcrAsync
    // For batch operations (loading or ocrAll), show completion toast here
    if (!wasFromButton) {
      if (wasAddedCount > 0) {
        toastDone('已加载并识别 ' + wasAddedCount + ' 张发票');
      } else if (wasBatchTotal > 0) {
        toastDone('识别完成');
      }
    }
  }
}

function updateOcrAllBtn() {
  var btn = document.getElementById('ocrAllBtn');
  if (!btn) return;
  var remaining = _ocrQueue.length + _ocrRunning;
  if (remaining > 0) {
    var done = _ocrBatchTotal > 0 ? _ocrBatchTotal - remaining : 0;
    btn.innerHTML = _ocrBatchTotal > 0
      ? '<span class="ocr-spinner"></span> ' + done + '/' + _ocrBatchTotal
      : '<span class="ocr-spinner"></span> ' + remaining;
    btn.disabled = true;
    btn.title = '识别中 ' + (_ocrBatchTotal > 0 ? done + '/' + _ocrBatchTotal : '剩余' + remaining);
  } else {
    btn.textContent = '\uD83D\uDD0D';
    btn.disabled = false;
    btn.title = '一键识别';
  }
}

function applyOcrAsync(fileObj, dataUrl) {
  if (!hasOcr || !isTauri || !invoke || window.__TAURI_CLOSING__) return;
  fileObj._ocrPending = true;
  updateFileItem(fileObj);
  updateOcrAllBtn();
  var hasFilePath = !!(fileObj._filePath);
  var isPdfPage = !!(fileObj._pdfPath && fileObj._pdfPageIdx >= 0);
  _ocrQueue.push(function() {
    var ocrPromise;
    if (isPdfPage) {
      // PDF page: use ocr_pdf_page — Rust renders + OCRs in one pass (zero IPC round-trip)
      ocrPromise = applyOcrPdfPage(fileObj);
    } else if (hasFilePath) {
      // Image file: Rust reads from disk directly — fast, no base64 round-trip
      ocrPromise = applyOcr(fileObj, '', fileObj._filePath);
    } else {
      // No file path (e.g. browser-mode image) — downsample for faster IPC
      ocrPromise = downsampleForOcr(dataUrl, 1280).then(function(ocrDataUrl) {
        return applyOcr(fileObj, ocrDataUrl);
      });
    }
    return ocrPromise.then(function() {
      fileObj._ocrPending = false;
      updateFileItem(fileObj);
      updateAmountSummary();
      // Show result toast only for single-file OCR triggered by button click
      // (_ocrFromButton === true means user clicked OCR on one file)
      // During batch loading or ocrAll, progress is shown via _onOcrTaskDone
      if (_ocrFromButton && _ocrQueue.length === 0 && _ocrRunning <= 1) {
        var amt = fileObj.amountTax || fileObj.amountNoTax;
        toast(amt > 0 ? '识别成功 \u00A5' + amt.toFixed(2) : '识别完成，未识别到金额', 2500);
      }
    }).catch(function(e) {
      fileObj._ocrPending = false;
      console.warn('[OCR] 后台识别失败:', e);
      if (_ocrFromButton && _ocrQueue.length === 0 && _ocrRunning <= 1) {
        toast('识别失败', 2500);
      }
    });
  });
  // Show toast with remaining count
  var remaining = _ocrQueue.length + _ocrRunning;
  if (_ocrToastActive) {
    var done = _ocrBatchTotal > 0 ? _ocrBatchTotal - remaining : 0;
    toastLoading(_ocrBatchTotal > 0 ? '识别中 ' + done + '/' + _ocrBatchTotal : '识别中，剩余 ' + remaining + ' 张');
  }
  _drainOcrQueue();
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
  var ab = (f.amountTax > 0 || f.amountNoTax > 0) ? '<span class="amt-badge">\u00A5' + (f.amountTax || f.amountNoTax).toFixed(2) + '</span>' : (f._ocrPending ? '<span class="ocr-spinner" title="识别中"></span>' : '');
  var sb = f.sellerName ? '<span class="' + (f._isTicket ? 'ticket-badge' : 'seller-badge') + '" title="' + escHtml(f.sellerCreditCode || '') + '">' + escHtml(f.sellerName.length > 16 ? f.sellerName.substring(0, 16) + '\u2026' : f.sellerName) + '</span>' : '';
  var metaEl = items[idx].querySelector('.file-meta');
  var sellerEl = items[idx].querySelector('.file-seller');
  if (metaEl) metaEl.innerHTML = fmtSize(f.size) + cb + rb + ab;
  if (sellerEl) {
    sellerEl.innerHTML = sb;
    sellerEl.style.display = sb ? '' : 'none';
  }
  // Update per-file OCR button state
  var ocrBtn = items[idx].querySelector('.ocr-btn');
  if (ocrBtn) {
    if (f._ocrPending) {
      ocrBtn.innerHTML = '<span class="ocr-spinner"></span>';
      ocrBtn.disabled = true;
      ocrBtn.title = '识别中';
      ocrBtn.onclick = null;
    } else {
      ocrBtn.textContent = '\uD83D\uDD0D';
      ocrBtn.disabled = false;
      ocrBtn.title = 'OCR识别';
      ocrBtn.onclick = (function(i) { return function() { ocrFile(i); }; })(idx);
    }
  }
}

/**
 * Render SVG string to PNG data URL via Canvas.
 * @param {string} svgString - SVG markup
 * @param {number} pageWidthMm - page width in mm
 * @param {number} pageHeightMm - page height in mm
 * @returns {Promise<string>} PNG data URL at 300 DPI
 */
function svgToPngDataUrl(svgString, pageWidthMm, pageHeightMm) {
  return new Promise(function(resolve, reject) {
    // OFD SVG scale=3.5, so viewBox = pageWidth * 3.5
    var svgScale = 3.5;
    var svgW = pageWidthMm * svgScale;
    var svgH = pageHeightMm * svgScale;
    // Target: 300 DPI
    var pxW = Math.round(pageWidthMm * PDF_RENDER_DPI / 25.4);
    var pxH = Math.round(pageHeightMm * PDF_RENDER_DPI / 25.4);

    var blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var img = new Image();
    img.onload = function() {
      var canvas = document.createElement('canvas');
      canvas.width = pxW;
      canvas.height = pxH;
      var ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, pxW, pxH);
      ctx.drawImage(img, 0, 0, svgW, svgH, 0, 0, pxW, pxH);
      URL.revokeObjectURL(url);
      try {
        var pngUrl = canvas.toDataURL('image/png');
        resolve(pngUrl);
      } catch(e) {
        reject(new Error('Canvas toDataURL failed: ' + e.message));
      }
    };
    img.onerror = function() {
      URL.revokeObjectURL(url);
      reject(new Error('SVG image load failed'));
    };
    img.src = url;
  });
}

/**
 * Fast load from FileData — show preview immediately, OCR in background.
 * @param {Object} fd - FileData from Rust: { name, dataUrl, size, ext, path, origW, origH }
 */
function loadFileFromDataUrlFast(fd) {
  var name = fd.name, dataUrl = fd.dataUrl, size = fd.size, ext = fd.ext, filePath = fd.path;
  return new Promise(function(resolve) {
    var id = 'f' + Date.now() + Math.random().toString(36).slice(2);

    if (ext === 'pdf') {
      if (isTauri && invoke && filePath) {
        // Render PDF pages only (fast, no OCR) — preview appears immediately.
        // OCR runs in background queue, so the user sees previews right away.
        invoke('render_pdf_pages', { pdfPath: filePath, dpi: PDF_RENDER_DPI }).then(async function(pages) {
          if (pages && pages.length > 0) {
            var results = [];
            for (var p = 0; p < pages.length; p++) {
              var pg = pages[p];
              var img = new Image(); img.src = pg.imageDataUrl;
              await new Promise(function(r) { img.onload = r; });
              var fileObj = createFileObj({
                id: id + '_p' + (p + 1),
                name: pages.length > 1 ? name.replace(/\.pdf$/i, '') + '_第' + (p + 1) + '页.pdf' : name,
                size: size, type: 'pdf', previewUrl: pg.imageDataUrl,
                img: img, renderDpi: pg.renderDpi || PDF_RENDER_DPI,
                pdfPath: filePath, pdfPageIdx: p
              });
              results.push(fileObj);
            }
            resolve(results.length === 1 ? results[0] : results);
            // Queue OCR for each page in background — no blocking!
            results.forEach(function(r) {
              if (S.feat.ocrEnabled) applyOcrAsync(r, r.previewUrl);
            });
            return;
          }
          toast('PDF 渲染结果为空: ' + name);
          resolve(null);
        }).catch(function(err) {
          console.error('[PDF] WinRT rendering failed:', err);
          toast('PDF 渲染失败: ' + name);
          resolve(null);
        });
        return;
      }
      // Non-Tauri: PDF files require native rendering
      toast('PDF 格式请使用桌面版打开');
      resolve(null);
    }
    // OFD: SVG vector rendering + structured invoice data from XML (skips OCR)
    else if (ext === 'ofd' && isTauri && invoke && filePath) {
      invoke('parse_ofd', { ofdPath: filePath }).then(function(result) {
        return svgToPngDataUrl(result.svg, result.pageWidth, result.pageHeight).then(function(pngUrl) {
          var img = new Image(); img.src = pngUrl;
          return new Promise(function(r) { img.onload = function() { r({img: img, pngUrl: pngUrl, info: result.invoiceInfo}); }; });
        });
      }).then(function(payload) {
        var info = payload.info || {};
        var fileObj = createFileObj({
          id: id, name: name, size: size, type: 'ofd',
          previewUrl: payload.pngUrl, img: payload.img,
          // Don't set filePath — OFD is a ZIP, not an image.
          // print.js needs sourceType='ofd-page' (FlateDecode), which requires _filePath to be empty.
          // The OFD path is only needed for parse_ofd (already done above).
          // Structured data from OFD XML — skip OCR
          amountTax: info.amountTax || 0,
          amountNoTax: info.amountNoTax || 0,
          taxAmount: info.taxAmount || 0,
          sellerName: info.sellerName || '',
          sellerCreditCode: info.sellerTaxId || '',
          invoiceNo: info.invoiceNo || '',
          invoiceDate: info.invoiceDate || '',
          buyerName: info.buyerName || '',
          buyerCreditCode: info.buyerTaxId || '',
          // OFD page dimensions for layout
          ow: payload.img.naturalWidth,
          oh: payload.img.naturalHeight,
          // Mark as OFD source for PDF generation (FlateDecode)
          _ofdPage: true
        });
        resolve(fileObj);
      }).catch(function(err) {
        // Fallback: call open_ofd_images for bitmap extraction
        console.warn('[OFD] parse_ofd failed, falling back to bitmap:', err);
        invoke('open_ofd_images', { ofdPath: filePath }).then(function(fileDataList) {
          if (fileDataList && fileDataList.length > 0) {
            // Load the first page as bitmap fallback
            var fd0 = fileDataList[0];
            var img = new Image(); img.src = fd0.dataUrl;
            img.onload = function() {
              var fileObj = createFileObj({
                id: id, name: fd0.name, size: fd0.size, type: fd0.ext,
                previewUrl: fd0.dataUrl, img: img,
                ow: fd0.origW || 0, oh: fd0.origH || 0
              });
              resolve(fileObj);
              if (S.feat.ocrEnabled) applyOcrAsync(fileObj, fd0.dataUrl);
            };
            img.onerror = function() { resolve(null); };
          } else {
            resolve(null);
          }
        }).catch(function() { resolve(null); });
      });
      return;
    }
    else if (ext === 'ofd') {
      toast('OFD 格式请使用桌面版打开');
      resolve(null);
    }
    else {
      var img = new Image(); img.src = dataUrl;
      img.onload = function() {
        var result = createFileObj({
          id: id, name: name, size: size, type: ext,
          previewUrl: dataUrl, img: img, filePath: filePath || '',
          // When Rust provides original dimensions (thumbnail mode), use them
          // instead of the thumbnail's naturalWidth/naturalHeight.
          // This ensures correct layout rotation and PDF sizing.
          ow: fd.origW || 0,
          oh: fd.origH || 0
        });
        resolve(result);
        // Background OCR — pass filePath to skip base64 round-trip
        if (S.feat.ocrEnabled) applyOcrAsync(result, dataUrl);
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
      // Browser mode: PDF files require native rendering, not available here
      toast('PDF 格式请使用桌面版打开');
      resolve(null);
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
        if (S.feat.ocrEnabled) applyOcrAsync(fileObj, e.target.result);
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
  var scrollTop = list.scrollTop;
  var sel = S.files.filter(function(f) { return f.checked; }).length;
  document.getElementById('fileCount').textContent = S.files.length + ' 张，已选 ' + sel;
  var summaryEl = document.getElementById('amountSummary');
  if (!S.files.length) { list.innerHTML = ''; if (summaryEl) summaryEl.style.display = 'none'; updateAmountSummary(); return; }
  if (summaryEl) summaryEl.style.display = 'flex';

  // Snapshot and clear new-file IDs so animation only plays once
  var currentNewIds = _newFileIds;
  _newFileIds = {};

  list.innerHTML = S.files.map(function(f, i) {
    var cls = 'file-item';
    if (currentNewIds[f.id]) cls += ' entering';
    if (f._loading) cls += ' loading-item';
    if (i === _activeFileIdx) cls += ' active-item';
    var cb = f.copies > 1 ? '<span class="copy-badge">' + f.copies + '份</span>' : '';
    var rb = f.rotation ? '<span class="rot-badge">' + f.rotation + '°</span>' : '';
    var ab = (f.amountTax > 0 || f.amountNoTax > 0) ? '<span class="amt-badge">\u00A5' + (f.amountTax || f.amountNoTax).toFixed(2) + '</span>' : (f._ocrPending ? '<span class="ocr-spinner" title="识别中"></span>' : '');
    var sb = f.sellerName ? '<span class="' + (f._isTicket ? 'ticket-badge' : 'seller-badge') + '" title="' + escHtml(f.sellerCreditCode || '') + '">' + escHtml(f.sellerName.length > 16 ? f.sellerName.substring(0, 16) + '\u2026' : f.sellerName) + '</span>' : '';
    // XSS FIX: escHtml(f.name) in both title and display text
    // XSS FIX: escHtml(f.previewUrl) in img src, escHtml(f.type) in type-badge
    var safePreviewUrl = escHtml(f.previewUrl || '');
    var safeType = escHtml(f.type === 'jpeg' ? 'jpg' : f.type);
    var thumbContent = f._loading ? '' : (f.previewUrl ? '<img src="' + safePreviewUrl + '">' : '\uD83D\uDCC4');
    var ocrBtnHtml = hasOcr
      ? (f._ocrPending
        ? '<button class="ib ocr-btn" disabled title="识别中"><span class="ocr-spinner"></span></button>'
        : '<button class="ib ocr-btn" onclick="ocrFile(' + i + ')" title="OCR识别">\uD83D\uDD0D</button>')
      : '';
    var actionBtns = f._loading
      ? '<button class="ib danger" onclick="rmFile(' + i + ')">\u2715</button>'
      : ocrBtnHtml + '<button class="ib" onclick="rotFile(' + i + ')" title="旋转90°">\u21BB</button><button class="ib danger" onclick="rmFile(' + i + ')">\u2715</button>';
    return '<div class="' + cls + '" draggable="true" ondragstart="dStart(event,' + i + ')" ondragover="dOver(event)" ondrop="dDrop(event,' + i + ')" onclick="clickFileItem(' + i + ',event)" ondblclick="openInvModal(' + i + ')">' +
      '<div class="file-check ' + (f.checked ? 'checked' : '') + '" onclick="togCheck(' + i + ')"></div>' +
      '<div class="file-thumb">' + thumbContent + '<div class="type-badge">' + safeType + '</div></div>' +
      '<div class="file-info"><div class="file-name" title="' + escHtml(f.name) + '">' + escHtml(f.name) + '</div><div class="file-meta">' + fmtSize(f.size) + cb + rb + ab + '</div>' + (sb ? '<div class="file-seller">' + sb + '</div>' : '<div class="file-seller" style="display:none"></div>') + '</div>' +
      '<div style="display:flex;gap:2px">' + actionBtns + '</div>' +
    '</div>';
  }).join('');

  // Apply staggered animation delay for entering items
  var enteringItems = list.querySelectorAll('.file-item.entering');
  enteringItems.forEach(function(el, idx) {
    el.style.animationDelay = (idx * 30) + 'ms';
  });

  list.scrollTop = scrollTop;
  updateAmountSummary();
}
function togCheck(i) { S.files[i].checked = !S.files[i].checked; renderFileList(); updatePreview(); }
function selectAll() { S.files.forEach(function(f) { f.checked = true; }); renderFileList(); updatePreview(); }
function deselectAll() { S.files.forEach(function(f) { f.checked = false; }); renderFileList(); updatePreview(); }
function deleteSelected() { if (!S.files.some(function(f) { return f.checked; })) return; S.files = S.files.filter(function(f) { return !f.checked; }); renderFileList(); updatePreview(); updatePrintBtn(); }
function rmFile(i) { S.files.splice(i, 1); if (_activeFileIdx === i) _activeFileIdx = -1; else if (_activeFileIdx > i) _activeFileIdx--; renderFileList(); updatePreview(); updatePrintBtn(); }
function rotFile(i) { S.files[i].rotation = (S.files[i].rotation + 90) % 360; renderFileList(); updatePreview(); }
function ocrFile(i) {
  var f = S.files[i];
  if (f._loading || f._ocrPending) return;
  if (!hasOcr) { toast('此版本不支持 OCR 识别'); return; }
  if (!isTauri || !invoke) { toast('OCR 识别需要桌面版'); return; }
  // Mark as single-file OCR from button click so per-file result toast shows correctly
  _ocrBatchTotal = 1;
  _ocrFromButton = true;
  _ocrToastActive = true;
  applyOcrAsync(f, f.previewUrl);
}
function ocrAll() {
  if (!hasOcr) { toast('此版本不支持 OCR 识别'); return; }
  if (!isTauri || !invoke) { toast('OCR 识别需要桌面版'); return; }
  var running = _ocrQueue.length + _ocrRunning;
  if (running > 0) { toast('正在识别中，请稍候'); return; }
  var targets = S.files.filter(function(f) {
    return !f._loading && !f._ocrPending && !(f.amountTax > 0 || f.amountNoTax > 0);
  });
  if (targets.length === 0) { toast('没有需要识别的发票'); return; }
  _ocrBatchTotal = targets.length;
  updateOcrAllBtn();
  toastLoading('识别中，共 ' + targets.length + ' 张...');
  targets.forEach(function(f) { applyOcrAsync(f, f.previewUrl); });
}
function clearAll() { if (!S.files.length) return; if (!confirm('确认清除所有发票？')) return; S.files = []; _activeFileIdx = -1; renderFileList(); updatePreview(); updatePrintBtn(); }

// Click file item → navigate preview to the page containing this invoice
function clickFileItem(idx, event) {
  // Ignore clicks on checkbox and action buttons
  if (event && (event.target.closest('.file-check') || event.target.closest('button'))) return;
  var f = S.files[idx];
  if (f._loading) return;

  _activeFileIdx = idx;

  // Auto-check if unchecked so the file appears in preview
  if (!f.checked) {
    f.checked = true;
  }

  // Find which page this file is on
  var activeFiles = getActiveFiles();
  var perPage = S.layout.cols * S.layout.rows;
  var activeIdx = -1;
  for (var i = 0; i < activeFiles.length; i++) {
    if (activeFiles[i].id === f.id) { activeIdx = i; break; }
  }
  if (activeIdx >= 0) {
    S.currentPage = Math.floor(activeIdx / perPage);
    updatePreview();
  }

  updateActiveFileHighlight();
  renderFileList();
}

// Update sidebar highlight to match _activeFileIdx
function updateActiveFileHighlight() {
  var list = document.getElementById('fileList');
  if (!list) return;
  var items = list.querySelectorAll('.file-item');
  items.forEach(function(el, i) {
    el.classList.toggle('active-item', i === _activeFileIdx);
  });
}

// Sync _activeFileIdx with current preview page (called from updatePreview)
function syncActiveFileFromPage() {
  var activeFiles = getActiveFiles();
  var perPage = S.layout.cols * S.layout.rows;
  var pageStart = S.currentPage * perPage;
  if (pageStart < activeFiles.length) {
    var firstFileOnPage = activeFiles[pageStart];
    var newIdx = S.files.indexOf(firstFileOnPage);
    if (newIdx !== _activeFileIdx) {
      _activeFileIdx = newIdx;
      updateActiveFileHighlight();
    }
  }
}
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
  var taxAmtTotal = checked.reduce(function(s, f) { return s + (f.taxAmount || 0); }, 0);
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
    var detailLines = '<span>含税 \u00A5' + taxTotal.toFixed(2) + '</span>';
    if (taxAmtTotal > 0) {
      detailLines += '<span style="font-size:11px;color:var(--text-muted);font-weight:400">不含税 \u00A5' + noTaxTotal.toFixed(2) + ' | 税额 \u00A5' + taxAmtTotal.toFixed(2) + '</span>';
    } else {
      detailLines += '<span style="font-size:11px;color:var(--text-muted);font-weight:400">不含税 \u00A5' + noTaxTotal.toFixed(2) + '</span>';
    }
    amtHtml = '<span class="amt-total" style="font-size:12px;display:flex;flex-direction:column;align-items:flex-end;gap:1px">' + detailLines + '</span>';
  }
  var sellerNames = [];
  checked.forEach(function(f) {
    if (f.sellerName) { var n = f.sellerName.trim(); if (sellerNames.indexOf(n) < 0) sellerNames.push(n); }
  });
  var sellerHtml = sellerNames.length > 0
    ? '<div style="font-size:10px;color:var(--text-muted);margin-top:2px">' + sellerNames.length + '个销售方</div>'
    : '';
  el.innerHTML = countHtml + amtHtml + sellerHtml;

  // Total amount is already shown in amountSummary (bottom-left), no need to duplicate in statusbar
}

// Invoice modal
function openInvModal(i) {
  if (S.files[i]._loading) return; // Don't open modal for loading placeholders
  S.editIdx = i; var f = S.files[i];
  var ocrText = f._ocrText || '';
  var ocrHtml = ocrText ? '<div style="margin-top:12px;border-top:1px solid var(--border);padding-top:10px"><div style="display:flex;align-items:center;gap:6px;cursor:pointer;margin-bottom:6px" onclick="this.nextElementSibling.classList.toggle(\'hidden\');this.querySelector(\'.arrow\').textContent=this.nextElementSibling.classList.contains(\'hidden\')?\'▶\':\'▼\'"><span class="arrow" style="font-size:10px;color:var(--text-muted)">▶</span><span style="font-size:12px;font-weight:600;color:var(--primary)">🔍 OCR识别全文</span><span style="font-size:10px;color:var(--text-muted)">(点击展开)</span></div><div class="hidden" style="position:relative"><pre style="margin:0;padding:8px 10px;background:var(--surface2);border-radius:6px;max-height:260px;overflow:auto;white-space:pre-wrap;word-break:break-all;font-size:11px;line-height:1.5;font-family:Consolas,monospace;border:1px solid var(--border)">' + escHtml(ocrText) + '</pre><button class="btn btn-sm" style="position:absolute;top:6px;right:6px;padding:3px 8px;font-size:11px;opacity:0.7" onclick="event.stopPropagation();copyOcrText(this)" title="复制OCR文本">📋 复制</button></div></div>' : '<div style="margin-top:12px;border-top:1px solid var(--border);padding-top:10px;font-size:11px;color:var(--text-muted)">⏳ OCR 全文尚未识别</div>';
  document.getElementById('invModalBody').innerHTML =
    '<div style="font-size:13px;padding:8px 10px;background:var(--surface2);border-radius:6px;margin-bottom:10px">\uD83D\uDCC4 ' + escHtml(f.name) + '</div>' +
    '<div class="row"><label class="lbl">份数</label><div style="display:flex;gap:4px;align-items:center"><button class="btn btn-sm btn-icon" onclick="changeModalCopies(-1)">\u2212</button><input type="number" id="mCopies" value="' + f.copies + '" min="1" max="99" style="width:52px;text-align:center"><button class="btn btn-sm btn-icon" onclick="changeModalCopies(1)">+</button></div></div>' +
    '<div class="row"><label class="lbl">含税价</label><div style="display:flex;gap:4px;align-items:center"><span style="font-size:14px;font-weight:600;color:var(--success)">\u00A5</span><input type="number" id="mAmountTax" value="' + (f.amountTax || '') + '" min="0" step="0.01" placeholder="0.00" style="width:100px;text-align:right"></div></div>' +
    '<div class="row"><label class="lbl">不含税</label><div style="display:flex;gap:4px;align-items:center"><span style="font-size:14px;font-weight:600;color:var(--text-muted)">\u00A5</span><input type="number" id="mAmountNoTax" value="' + (f.amountNoTax || '') + '" min="0" step="0.01" placeholder="0.00" style="width:100px;text-align:right"></div></div>' +
    '<div class="row"><label class="lbl">税额</label><div style="display:flex;gap:4px;align-items:center"><span style="font-size:14px;font-weight:600;color:var(--warning,orange)">\u00A5</span><input type="number" id="mTaxAmount" value="' + (f.taxAmount || '') + '" min="0" step="0.01" placeholder="0.00" style="width:100px;text-align:right"></div></div>' +
    '<div class="row"><label class="lbl">发票号码</label><div style="flex:1;display:flex;gap:4px;align-items:center"><input type="text" id="mInvoiceNo" value="' + escHtml(f.invoiceNo || '') + '" placeholder="自动识别" style="flex:1;font-size:11px;min-width:0;font-family:monospace"></div></div>' +
    '<div class="row"><label class="lbl">开票日期</label><div style="flex:1;display:flex;gap:4px;align-items:center"><input type="text" id="mInvoiceDate" value="' + escHtml(f.invoiceDate || '') + '" placeholder="自动识别" style="flex:1;font-size:11px;min-width:0"></div></div>' +
    '<div class="row"><label class="lbl">购买方</label><div style="flex:1;display:flex;gap:4px;align-items:center"><input type="text" id="mBuyer" value="' + escHtml(f.buyerName || '') + '" placeholder="自动识别" style="flex:1;font-size:11px;min-width:0"></div></div>' +
    '<div class="row"><label class="lbl">购方代码</label><div style="flex:1;display:flex;gap:4px;align-items:center"><input type="text" id="mBuyerCreditCode" value="' + escHtml(f.buyerCreditCode || '') + '" placeholder="自动识别" style="flex:1;font-size:11px;min-width:0;font-family:monospace"></div></div>' +
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
  var ta = parseFloat(document.getElementById('mTaxAmount').value);
  f.amountTax = isNaN(at) || at < 0 ? 0 : Math.round(at * 100) / 100;
  f.amountNoTax = isNaN(an) || an < 0 ? 0 : Math.round(an * 100) / 100;
  f.taxAmount = isNaN(ta) || ta < 0 ? 0 : Math.round(ta * 100) / 100;
  f.amount = f.amountTax || f.amountNoTax;
  f.sellerName = document.getElementById('mSeller').value;
  f.sellerCreditCode = document.getElementById('mCreditCode').value;
  f.invoiceNo = document.getElementById('mInvoiceNo').value;
  f.invoiceDate = document.getElementById('mInvoiceDate').value;
  f.buyerName = document.getElementById('mBuyer').value;
  f.buyerCreditCode = document.getElementById('mBuyerCreditCode').value;
  closeInvModal(); renderFileList(); updatePreview(); updateAmountSummary();
}

function copyOcrText(btn) {
  var pre = btn.parentElement.querySelector('pre');
  if (!pre) return;
  var text = pre.textContent || pre.innerText;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(function() {
      btn.textContent = '✓ 已复制';
      setTimeout(function() { btn.innerHTML = '📋 复制'; }, 1500);
    }).catch(function() { fallbackCopy(text, btn); });
  } else {
    fallbackCopy(text, btn);
  }
}
function fallbackCopy(text, btn) {
  var ta = document.createElement('textarea');
  ta.value = text; ta.style.position = 'fixed'; ta.style.left = '-9999px';
  document.body.appendChild(ta); ta.select();
  try { document.execCommand('copy'); btn.textContent = '✓ 已复制'; setTimeout(function() { btn.innerHTML = '📋 复制'; }, 1500); }
  catch(e) { toast('复制失败'); }
  document.body.removeChild(ta);
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
var _printersLoaded = false;
function switchTab(n, el) {
  document.querySelectorAll('.sidebar-tab').forEach(function(t) { t.classList.remove('active'); });
  document.querySelectorAll('.sidebar-panel').forEach(function(p) { p.classList.add('hidden'); });
  el.classList.add('active');
  document.getElementById('panel-' + n).classList.remove('hidden');
  // Lazy-load printers on first visit to print tab
  if (n === 'print' && !_printersLoaded && isTauri && invoke) {
    _printersLoaded = true;
    refreshPrinters();
  }
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
  var files = S.files.filter(function(f) { return f.checked && !f._loading; });
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
  _pdfDirty = true;
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
  syncActiveFileFromPage();
}

function updatePageDots(t) {
  var d = document.getElementById('pageDots');
  if (t <= 1) { d.innerHTML = ''; return; }
  var MAX_DOTS = 9;
  if (t <= MAX_DOTS) {
    // All pages fit — show every dot
    d.innerHTML = Array.from({ length: t }, function(_, i) {
      return '<div class="page-dot ' + (i === S.currentPage ? 'active' : '') + '" onclick="gotoPage(' + i + ')"></div>';
    }).join('');
  } else {
    // Sliding window: show dots around current page with ellipsis indicators
    var cur = S.currentPage;
    var half = Math.floor((MAX_DOTS - 2) / 2); // dots on each side of center (reserve 2 for ellipsis)
    var start = Math.max(1, cur - half);
    var end = Math.min(t - 2, start + MAX_DOTS - 3);
    start = Math.max(1, end - (MAX_DOTS - 3));
    var html = '<div class="page-dot ' + (cur === 0 ? 'active' : '') + '" onclick="gotoPage(0)"></div>';
    if (start > 1) html += '<div class="page-dot ellipsis" title="更多页">···</div>';
    for (var i = start; i <= end; i++) {
      html += '<div class="page-dot ' + (i === cur ? 'active' : '') + '" onclick="gotoPage(' + i + ')"></div>';
    }
    if (end < t - 2) html += '<div class="page-dot ellipsis" title="更多页">···</div>';
    html += '<div class="page-dot ' + (cur === t - 1 ? 'active' : '') + '" onclick="gotoPage(' + (t - 1) + ')"></div>';
    d.innerHTML = html;
  }
}
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
  // Persist OCR enabled setting
  if (k === 'ocrEnabled') {
    try { localStorage.setItem('fapiao-ocr-enabled', S.feat[k] ? '1' : '0'); } catch(e) {}
  }
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

async function verifyInvoice(backup) {
  // 主：国家税务总局官方查验平台；备：仿真平台（证书有效）
  var urls = {
    primary: 'https://inv-veri.chinatax.gov.cn/',
    backup: 'https://fz.chinaive.com/fpcy/'
  };
  var url = backup ? urls.backup : urls.primary;
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
  S.feat = { cutline: true, number: false, border: false, trimWhite: false, watermark: false, collate: true, duplex: false, pageNum: false, printDate: false, confirmPrint: true, autoOpenPdf: true, ocrEnabled: false };
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
  document.getElementById('toggleOcrEnabled').classList.remove('on');
  document.getElementById('printMode').value = 'dialog';
  document.getElementById('themeMode').value = 'light';
  document.documentElement.classList.remove('dark');
  try { localStorage.removeItem('fapiao-theme'); } catch(e) {}
  try { localStorage.removeItem('fapiao-save-dir'); } catch(e) {}
  try { localStorage.removeItem('fapiao-amt-mode'); } catch(e) {}
  try { localStorage.removeItem('fapiao-ocr-enabled'); } catch(e) {}
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

// beforeunload safety net — stop all work if the window is being destroyed
// (covers cases where _tauriCleanup() wasn't called or didn't execute in time)
window.addEventListener('beforeunload', function() {
  window.__TAURI_CLOSING__ = true;
  _ocrQueue = [];
  _ocrRunning = 0;
  _loadingBatchActive = false;
});

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

// Printers are loaded on-demand when user opens the print tab (see switchTab)

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

// Restore OCR enabled setting
(function() {
  try {
    var v = localStorage.getItem('fapiao-ocr-enabled');
    if (v === '1') {
      S.feat.ocrEnabled = true;
      document.getElementById('toggleOcrEnabled').classList.add('on');
    }
  } catch(e) {}
})();

// =====================================================
// Show main window after DOM is ready (window starts hidden via visible:false)
// =====================================================
(function() {
  function showApp() {
    if (isTauri && invoke) {
      // Check OCR availability at startup
      invoke('check_ocr_available').then(function(available) {
        hasOcr = !!available;
        // Hide OCR-specific UI if OCR is not available
        if (!hasOcr) {
          var ocrAllBtn = document.getElementById('ocrAllBtn');
          if (ocrAllBtn) ocrAllBtn.style.display = 'none';
          var ocrSection = document.getElementById('ocrSection');
          if (ocrSection) ocrSection.style.display = 'none';
        }
      }).catch(function() {});
      // Get app version from Rust (compiled from Cargo.toml)
      invoke('get_app_version').then(function(v) {
        APP_VERSION = v;
        var el = document.getElementById('stVersion');
        if (el) el.textContent = 'v' + v;
        console.log('发票批量打印 v' + v + ' | isTauri:', isTauri);
      }).catch(function() {});
      try { invoke('show_window'); } catch(e) {}
    } else {
      // Non-Tauri (browser) fallback
      var el = document.getElementById('stVersion');
      if (el) el.textContent = 'web';
    }
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', showApp);
  } else {
    showApp();
  }
  setTimeout(showApp, 2000);
})();
