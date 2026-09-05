// =====================================================
// 发票酱 — Web 入口
// =====================================================
var APP_VERSION = '';

// =====================================================
// Constants
// =====================================================
var PAPER = { A4:{w:210,h:297}, A5:{w:148,h:210}, B5:{w:176,h:250}, letter:{w:216,h:279}, legal:{w:216,h:356} };
var MM2PX = 96 / 25.4;
var PDF_RENDER_DPI = 300;  // Render/print DPI
var PDF_PREVIEW_DPI = 300;  // Preview DPI
var _loadingBatchActive = false;
var _printedMap = {};
var WHITE_THRESHOLD = 245; // Pixel value threshold for white-edge trimming

function nextFrame() { return new Promise(function(r) { requestAnimationFrame(function() { requestAnimationFrame(r); }); }); }

// Wait for async module scripts (pdf-client.js, idb-store.js) to be ready
function waitForGlobal(name, timeout) {
  timeout = timeout || 15000;
  if (window[name]) return Promise.resolve(window[name]);
  return new Promise(function(resolve, reject) {
    var start = Date.now();
    var timer = setInterval(function() {
      if (window[name]) { clearInterval(timer); resolve(window[name]); }
      else if (Date.now() - start > timeout) { clearInterval(timer); reject(new Error(name + ' not available after ' + timeout + 'ms')); }
    }, 50);
  });
}

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
  selectedSlot: -1,  // Index of currently selected slot in preview (for per-slot adjustment)
  amtMode: 'tax',
  printedFilter: 'all',
  fileFilter: 'all',   // 'all' | 'duplicates'
  fileView: 'list',    // 'list' | 'grid'
  feat: {
    cutline: true, number: false, border: false, trimWhite: false,
    watermark: false, pageNum: false,
    printDate: false, footer: false,
    pdfTextEnabled: true,
    customFM: false
  }
};

// Track newly added file IDs for entrance animation
var _newFileIds = {};

var _activeFileIdx = -1;

// =====================================================
// File Object Factory — unified creation with defaults
// =====================================================
function createFileObj(opts) {
  var obj = {
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
    invoiceType: opts.invoiceType || '',
    _isTicket: opts._isTicket || false,
    _loading: opts._loading || false,
    _xmlInvoice: opts._xmlInvoice || false,
    // Disk path for the original file (when available).
    // Used by Rust to read bytes directly, skipping base64 encode/decode.
    _filePath: opts.filePath || '',
    _pdfPath: opts.pdfPath || '',
    _pdfPageIdx: opts.pdfPageIdx != null ? opts.pdfPageIdx : -1,
    // Per-slot adjustment: scale & position within the layout slot
    slotScale: opts.slotScale || 1,        // 1.0 = default (contain-fit size)
    slotOffsetX: opts.slotOffsetX || 0,    // X offset in mm (0 = centered)
    slotOffsetY: opts.slotOffsetY || 0,    // Y offset in mm (0 = centered)
    _printed: false,                       // True after successful print
    // Vector print: embed original PDF page via pdf-lib embedPage
    srcPdfBytes: opts.srcPdfBytes || null,
    srcPageIndex: opts.srcPageIndex != null ? opts.srcPageIndex : -1,
    srcPageWidthPt: opts.srcPageWidthPt || 0,
    srcPageHeightPt: opts.srcPageHeightPt || 0
  };

  // Apply saved per-file adjustments if memory is enabled
  if (S.feat.slotAdjMemory && S._fileAdjMap) {
    var saved = S._fileAdjMap[obj.name];
    if (saved) {
      obj.slotScale = saved.scale != null ? saved.scale : obj.slotScale;
      obj.slotOffsetX = saved.offX != null ? saved.offX : obj.slotOffsetX;
      obj.slotOffsetY = saved.offY != null ? saved.offY : obj.slotOffsetY;
    }
  }

  // Restore saved note for this file
  if (S._notesMap && S._notesMap[obj.name]) {
    obj.note = S._notesMap[obj.name];
  }
  // Restore printed state
  var printKey = obj._filePath || obj._pdfPath;
  if (printKey && _printedMap && _printedMap[printKey]) {
    obj._printed = true;
  }

  return obj;
}

// =====================================================
// Helpers
// =====================================================
var toastT = null;
function toast(msg, dur) { dur = dur || 2500; var e = document.getElementById('toast'); e.textContent = msg; e.classList.add('show'); clearTimeout(toastT); if (dur > 0) toastT = setTimeout(function() { e.classList.remove('show'); }, dur); else clearTimeout(toastT); }
function toastHtml(msg, dur) { dur = dur || 2500; var e = document.getElementById('toast'); e.innerHTML = msg; e.classList.add('show'); clearTimeout(toastT); if (dur > 0) toastT = setTimeout(function() { e.classList.remove('show'); }, dur); else clearTimeout(toastT); }
function toastLoading(msg) { toastHtml('<span class="toast-spinner"></span>' + msg, 0); }
function toastDone(msg) { toast(msg, 2500); }
function hideToast() { var e = document.getElementById('toast'); e.classList.remove('show'); clearTimeout(toastT); }
function syncSlider(s, n) { document.getElementById(n).value = s.value; }
function syncRange(n, s) { document.getElementById(s).value = n.value; }

/**
 * Enable mouse wheel to increment/decrement number inputs and range sliders.
 * Delegated to the sidebar; covers all settings panel inputs and adj panel inputs.
 */
function setupInputWheelSupport() {
  var sidebar = document.querySelector('.sidebar');
  if (!sidebar) return;
  sidebar.addEventListener('wheel', function(e) {
    var t = e.target;
    if (t.tagName !== 'INPUT') return;
    if (t.type !== 'number' && t.type !== 'range') return;
    e.preventDefault();

    var step = parseFloat(t.step) || 1;
    if (t.type === 'range') step = parseFloat(t.step) || 1;
    var min = t.hasAttribute('min') ? parseFloat(t.min) : -Infinity;
    var max = t.hasAttribute('max') ? parseFloat(t.max) : Infinity;
    var val = parseFloat(t.value);
    if (isNaN(val)) val = 0;

    if (e.deltaY < 0) val += step;
    else if (e.deltaY > 0) val -= step;

    val = Math.max(min, Math.min(max, val));
    // Round to step precision to avoid floating-point noise
    var decimals = (step.toString().split('.')[1] || '').length;
    val = parseFloat(val.toFixed(Math.max(decimals, 0)));

    t.value = val;
    t.dispatchEvent(new Event('input', { bubbles: true }));
    t.dispatchEvent(new Event('change', { bubbles: true }));
  }, { passive: false });
}
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
    } else if (phase === 'print') {
      detail.textContent = current + ' / ' + total + ' 页';
      if (text) text.textContent = '正在渲染打印...';
    } else {
      detail.textContent = current + ' / ' + total;
      if (text) text.textContent = '正在处理...';
    }
    if (detail.textContent) detail.classList.remove('hidden'); else detail.classList.add('hidden');
  }
}
function fmtSize(b) { return b < 1024 ? b + 'B' : b < 1048576 ? (b / 1024).toFixed(1) + 'KB' : (b / 1048576).toFixed(1) + 'MB'; }
function escHtml(s) { return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function openExternal(url) {
  window.open(url, '_blank');
}



function showPdfiumMissing(reason) {
  toast('PDF 渲染引擎缺失，请确保服务端已安装 PDFium');
}

// Convert data URL to Uint8Array
function dataUrlToUint8Array(dataUrl) {
  var base64 = dataUrl.split(',')[1] || dataUrl;
  var binaryStr = atob(base64);
  var bytes = new Uint8Array(binaryStr.length);
  for (var i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
  return bytes;
}

// =====================================================
async function triggerUpload() {
  document.getElementById('fileInput').click();
}

// 版面空槽点击上传：新文件精准落在点击的槽位，中间空隙自动补空白占位
function addFileToSlot(slotIdx) {
  if (_slotUploadActive || _loadingBatchActive) {
    toast('当前仍在加载发票，请稍候再补传');
    return;
  }
  _slotUploadActive = true;
  _insertSlotIdx = slotIdx;
  // 原生 input 取消选择不会触发 change，窗口重新获得焦点时释放槽位锁
  var inputEl = document.getElementById('fileInput');
  var onPickerFocus = function() {
    window.removeEventListener('focus', onPickerFocus);
    setTimeout(function() {
      if (!inputEl.files || inputEl.files.length === 0) {
        _insertSlotIdx = -1;
        _slotUploadActive = false;
      }
    }, 100);
  };
  window.addEventListener('focus', onPickerFocus);
  inputEl.click();
}

async function handleFileInput(fl) {
  if (!fl || !fl.length) return;
  try {
    await processFileList(Array.from(fl));
  } catch(e) {
    toast('加载失败: ' + String(e));
  }
  document.getElementById('fileInput').value = '';
}

// Process File objects purely in the browser — IndexedDB for persistence
async function processFileList(fileList) {
  var total = fileList.length;
  var completed = 0;
  var added = 0;
  _loadingBatchActive = true;

  fileList.forEach(function(file) {
    var ph = createFileObj({
      name: file.name,
      size: file.size,
      type: (file.name.split('.').pop() || '').toLowerCase(),
      _loading: true
    });
    ph._placeholderKey = ph.id;
    file._phKey = ph._placeholderKey;
    S.files.push(ph);
    _newFileIds[ph.id] = true;
  });

  renderFileList(); updatePreview(); updatePdfBtn(); updateSummaryBtn();
  toastLoading('加载中 0/' + total);

  // Launch all file loads concurrently for speed
  var resolvedCount = 0;
  var loadPromises = fileList.map(function(file) {
    return loadFileFromFile(file).catch(function(err) {
      console.error('Load file error:', file.name, err);
      return null;
    }).then(function(r) {
      resolvedCount++;
      // Update toast as each file resolves (even before for-loop consumes it)
      toastLoading('加载中 ' + resolvedCount + '/' + total);
      return r;
    });
  });

  var startTime = Date.now();
  for (var fdIdx = 0; fdIdx < fileList.length; fdIdx++) {
    var r = await loadPromises[fdIdx];
    completed++;

    var file = fileList[fdIdx];

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

    renderFileList(); updatePreview(); updatePdfBtn(); updateSummaryBtn();
    await nextFrame();
  }
  if (slotInsert) locateInsertedFile(_lastInsertedId);
  toastLoading('加载完成');
  _loadingBatchActive = false;
  _insertSlotIdx = -1;
  _lastInsertedId = null;

  var elapsed = Date.now() - startTime;
  var minToastDelay = Math.max(300, 800 - elapsed);
  if (added > 0) {
    setTimeout(function() { toast('已加载 ' + added + ' 张发票', 2500); }, minToastDelay);
  } else {
    toast('文件加载失败');
  }
}
// =====================================================

function buildAmtBadge(f) {
  if (f.amountTax > 0 || f.amountNoTax > 0) {
    return '<span class="amt-badge">\u00A5' + (f.amountTax || f.amountNoTax).toFixed(2) + '</span>';
  }
  if (f._amtValidationFail) {
    var v = f._amtValidationFail;
    var tip = '\u26A0 金额校验失败\n含税: \u00A5' + v.amountTax.toFixed(2) +
      '\n不含税: \u00A5' + v.amountNoTax.toFixed(2) +
      '\n税额: \u00A5' + v.taxAmount.toFixed(2) +
      '\n验证: \u00A5' + v.amountNoTax.toFixed(2) + ' + \u00A5' + v.taxAmount.toFixed(2) + ' = \u00A5' + (Math.round((v.amountNoTax + v.taxAmount) * 100) / 100).toFixed(2) + ' \u2260 \u00A5' + v.amountTax.toFixed(2);
    return '<span class="amt-warn-badge" title="' + escHtml(tip) + '">\u26A0\u00A5' + v.amountTax.toFixed(2) + '</span>';
  }
  return '';
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
  var ab = buildAmtBadge(f);
  var pd = f._printed ? '<span class="printed-dot" title="已打印">✓</span>' : '';
  if (S.fileView === 'grid') {
    // grid 卡片：更新 card-meta（与 renderFileList grid 分支字段顺序一致）
    var cardMetaEl = items[idx].querySelector('.card-meta');
    if (cardMetaEl) {
      var gdupb = f._dup ? '<span class="dup-badge" title="检测到重复发票">⚠</span>' : '';
      cardMetaEl.innerHTML = pd + ab + cb + rb + gdupb + '<span class="card-size" title="文件大小">' + fmtSize(f.size) + '</span>';
    }
    var gsellerHtml = f.sellerName ? '<span class="' + (f._isTicket ? 'ticket-badge' : f._isNonTax ? 'nontax-badge' : f._isToll ? 'toll-badge' : 'seller-badge') + '">' + escHtml(f.sellerName) + '</span>' : '';
    var sellerLine = items[idx].querySelector('.card-seller');
    if (sellerLine) {
      sellerLine.innerHTML = gsellerHtml;
      sellerLine.title = f.sellerName || '';
      sellerLine.style.display = gsellerHtml ? '' : 'none';
    } else if (gsellerHtml) {
      var cardNameEl = items[idx].querySelector('.card-name');
      if (cardNameEl && cardNameEl.parentElement) {
        var newSellerLine = document.createElement('div');
        newSellerLine.className = 'card-seller';
        newSellerLine.title = f.sellerName || '';
        newSellerLine.innerHTML = gsellerHtml;
        cardNameEl.parentElement.insertBefore(newSellerLine, cardNameEl.nextSibling);
      }
    }
  } else {
    var sb = f.sellerName ? '<span class="' + (f._isTicket ? 'ticket-badge' : f._isNonTax ? 'nontax-badge' : f._isToll ? 'toll-badge' : 'seller-badge') + '" title="' + escHtml(f.sellerCreditCode || f.sellerName) + '">' + escHtml(f.sellerName) + '</span>' : '';
    // 只更新 .file-meta-left，保留 file-meta-right 操作按钮与布局结构
    var leftEl = items[idx].querySelector('.file-meta-left');
    if (leftEl) {
      var dupb = f._dup ? '<span class="dup-badge" title="检测到重复发票：点击左上角「重复」筛选可一键勾选删除">⚠重复</span>' : '';
      leftEl.innerHTML = pd + '<span class="file-size">' + fmtSize(f.size) + '</span>' + cb + rb + dupb + ab;
    }
    var sellerEl = items[idx].querySelector('.file-seller');
    if (sellerEl) {
      sellerEl.innerHTML = sb;
      sellerEl.title = f.sellerName || '';
      sellerEl.style.display = sb ? '' : 'none';
    } else if (sb) {
      // .file-seller didn't exist at render time (no sellerName yet), insert it now
      var nameEl = items[idx].querySelector('.file-name');
      if (nameEl && nameEl.parentElement) {
        var newSeller = document.createElement('div');
        newSeller.className = 'file-seller';
        newSeller.title = f.sellerName || '';
        newSeller.innerHTML = sb;
        nameEl.parentElement.insertBefore(newSeller, nameEl.nextSibling);
      }
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

function readFileAsArrayBuffer(file) {
  return new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onload = function() { resolve(reader.result); };
    reader.onerror = function() { reject(reader.error); };
    reader.readAsArrayBuffer(file);
  });
}

async function loadXmlFromFile(file, id, name, size) {
  var xmlClient = await waitForGlobal('__xmlClient');
  var info = await xmlClient.parseXmlInvoice(file);
  if (!info.invoiceNo && !info.sellerName) {
    toast('无法识别的 XML 发票: ' + name);
    return null;
  }
  var fileObj = createFileObj({
    id: id,
    name: name,
    size: size,
    type: 'xml',
    previewUrl: '',
    img: null,
    ow: 0,
    oh: 0,
    _xmlInvoice: true
  });
  if (info.invoiceNo) fileObj.invoiceNo = info.invoiceNo;
  if (info.invoiceDate) fileObj.invoiceDate = info.invoiceDate;
  if (info.sellerName) fileObj.sellerName = info.sellerName;
  if (info.sellerTaxId) fileObj.sellerCreditCode = info.sellerTaxId;
  if (info.buyerName) fileObj.buyerName = info.buyerName;
  if (info.buyerTaxId) fileObj.buyerCreditCode = info.buyerTaxId;
  if (info.amountTax != null) fileObj.amountTax = info.amountTax;
  if (info.amountNoTax != null) fileObj.amountNoTax = info.amountNoTax;
  if (info.taxAmount != null) fileObj.taxAmount = info.taxAmount;
  if (info.invoiceType) fileObj.invoiceType = info.invoiceType;
  fileObj._pdfTextExtracted = true;
  return fileObj;
}

async function loadOfdFromFile(file, id, name, size) {
  var buffer = await readFileAsArrayBuffer(file);
  var ofdClient = await waitForGlobal('__ofdClient');
  var result = await ofdClient.parseOfdFromArrayBuffer(buffer);
  var previewUrl = await svgToPngDataUrl(result.svg, result.pageWidth, result.pageHeight);
  var pxW = Math.round(result.pageWidth * PDF_RENDER_DPI / 25.4);
  var pxH = Math.round(result.pageHeight * PDF_RENDER_DPI / 25.4);
  var fileObj = createFileObj({
    id: id,
    name: name,
    size: size,
    type: 'ofd',
    previewUrl: previewUrl,
    img: null,
    ow: pxW,
    oh: pxH,
    renderDpi: PDF_RENDER_DPI
  });
  var info = result.invoiceInfo || {};
  if (info.invoiceNo) fileObj.invoiceNo = info.invoiceNo;
  if (info.invoiceDate) fileObj.invoiceDate = info.invoiceDate;
  if (info.sellerName) fileObj.sellerName = info.sellerName;
  if (info.sellerCreditCode) fileObj.sellerCreditCode = info.sellerCreditCode;
  if (info.sellerTaxId) fileObj.sellerCreditCode = info.sellerTaxId;
  if (info.buyerName) fileObj.buyerName = info.buyerName;
  if (info.buyerTaxId) fileObj.buyerCreditCode = info.buyerTaxId;
  if (info.amountTax != null) fileObj.amountTax = info.amountTax;
  if (info.amountNoTax != null) fileObj.amountNoTax = info.amountNoTax;
  if (info.taxAmount != null) fileObj.taxAmount = info.taxAmount;
  if (info.invoiceType) fileObj.invoiceType = info.invoiceType;
  fileObj._ofdSvg = result.svg;
  fileObj._ofdPageWidth = result.pageWidth;
  fileObj._ofdPageHeight = result.pageHeight;
  fileObj._pdfTextExtracted = true;
  return fileObj;
}

var _pdfMissingFontToastShown = false;  // 防止批量加载时反复弹 toast

async function loadPdfFromFile(file, id, name, size) {
  var buffer = await readFileAsArrayBuffer(file);
  var srcBuffer = buffer.slice(0);
  var idb = await waitForGlobal('__idb');
  var pdfClient = await waitForGlobal('__pdfClient');
  await idb.putFile(id, name, file.type || 'application/pdf', buffer);
  var loaded = await pdfClient.loadPdfConcurrent(buffer);
  var results = [];
  var hadMissingFontInBatch = false;
  for (var p = 0; p < loaded.pages.length; p++) {
    var pg = loaded.pages[p];
    var fileObj = createFileObj({
      id: id + '_p' + (p + 1),
      name: loaded.pages.length > 1 ? name.replace(/\.pdf$/i, '') + '_第' + (p + 1) + '页.pdf' : name,
      size: size,
      type: 'pdf',
      previewUrl: pg.previewUrl,
      img: null,
      ow: pg.width || 0,
      oh: pg.height || 0,
      renderDpi: pg.renderDpi || PDF_RENDER_DPI,
      srcPdfBytes: srcBuffer,
      srcPageIndex: p,
      srcPageWidthPt: pg.pdfWidthPt || 0,
      srcPageHeightPt: pg.pdfHeightPt || 0
    });

    // ---- 字体未嵌入检测（如 12306 铁路电子客票）----
    // pdf.js 依赖 ToUnicode CMap 做字形→Unicode 映射，字体未嵌入且无 ToUnicode 时
    // getTextContent() 返回空串，字段识别会全失败。这里检测并提示用户。
    var diag = pg.pdfTextResult && pg.pdfTextResult.textLayerDiagnostic;
    if (diag && diag.suspectMissingToUnicode) {
      fileObj._pdfTextMissingFonts = true;
      fileObj._pdfTextDiag = diag;
      hadMissingFontInBatch = true;
      console.warn('[PDF加载] 文字层缺失（字体未嵌入）:', name, diag);
    }

    if (pg.pdfTextResult && pg.pdfTextResult.hasTextLayer && pg.pdfTextResult.lines.length > 0) {
      if (typeof applyPdfTextResult === 'function') {
        applyPdfTextResult(fileObj, pg.pdfTextResult);
      }
    }
    results.push(fileObj);
  }

  // 批量加载时只弹一次 toast
  if (hadMissingFontInBatch && !_pdfMissingFontToastShown) {
    _pdfMissingFontToastShown = true;
    toastHtml('<span style="font-weight:600">\u26A0 部分发票字体未嵌入</span><br>' +
      '<span style="font-size:12px;opacity:.9">字段无法自动识别（如 12306 电子客票），请手动填写金额/销售方/发票号</span>', 5000);
    setTimeout(function() { _pdfMissingFontToastShown = false; }, 6000);
  }

  return results;
}

async function loadImageFromFile(file, id, name, size, ext) {
  var dataUrl = await new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onload = function() { resolve(reader.result); };
    reader.onerror = function() { reject(reader.error); };
    reader.readAsDataURL(file);
  });
  var img = await new Promise(function(resolve, reject) {
    var i = new Image();
    i.onload = function() { resolve(i); };
    i.onerror = reject;
    i.src = dataUrl;
  });
  return createFileObj({
    id: id, name: name, size: size, type: ext,
    previewUrl: dataUrl, img: img,
    ow: img.naturalWidth, oh: img.naturalHeight
  });
}

async function loadFileFromFile(file) {
  var name = file.name;
  var size = file.size;
  var ext = (name.split('.').pop() || '').toLowerCase();
  var id = 'f' + Date.now() + Math.random().toString(36).slice(2);

  try {
    if (ext === 'pdf') {
      return await loadPdfFromFile(file, id, name, size);
    }
    if (['jpg','jpeg','png','bmp','webp','gif'].indexOf(ext) >= 0) {
      return await loadImageFromFile(file, id, name, size, ext);
    }
    if (ext === 'ofd' || ext === 'ofx') {
      return await loadOfdFromFile(file, id, name, size);
    }
    if (ext === 'xml') {
      return await loadXmlFromFile(file, id, name, size);
    }
    toast('不支持的格式: ' + ext);
    return null;
  } catch (e) {
    console.error('loadFileFromFile error:', name, e);
    toast('加载失败: ' + name);
    return null;
  }
}

// Drag & Drop (browser native)
function handleDragOver(e) { e.preventDefault(); e.stopPropagation(); document.getElementById('dropZone').classList.add('drag-over'); }
function handleDragLeave(e) { e.preventDefault(); e.stopPropagation(); document.getElementById('dropZone').classList.remove('drag-over'); }
async function handleDrop(e) {
  e.preventDefault(); e.stopPropagation();
  document.getElementById('dropZone').classList.remove('drag-over');
  if (e.dataTransfer.files && e.dataTransfer.files.length) {
    try {
      await processFileList(Array.from(e.dataTransfer.files));
    } catch(err) {
      toast('加载失败: ' + String(err));
    }
  }
}

// =====================================================
// File list management
// =====================================================
function setPrintedFilter(filter) {
  S.printedFilter = filter;
  S.fileFilter = 'all';
  syncFilterButtons();
  renderFileList();
}

// 按 S.fileFilter / S.printedFilter 统一同步筛选按钮高亮
function syncFilterButtons() {
  var active = S.fileFilter === 'duplicates' ? 'duplicates' : S.printedFilter;
  document.querySelectorAll('.pf-btn').forEach(function(b) {
    b.classList.toggle('pf-active', b.dataset.filter === active);
  });
}

function setFileFilter(filter) {
  if (filter === 'duplicates') { selectDuplicateExtras(); return; }
  S.fileFilter = 'all';
  syncFilterButtons();
  renderFileList();
}

function getFilteredFiles() {
  if (S.fileFilter === 'duplicates') return S.files.filter(function(f) { return f._dup; });
  if (S.printedFilter === 'all') return S.files;
  return S.files.filter(function(f) {
    if (S.printedFilter === 'printed') return f._printed;
    if (S.printedFilter === 'unprinted') return !f._printed;
    return true;
  });
}

// 生成发票去重key：优先发票号，回退到 销售方+含税金额+日期（针对重复下载被改名的文件）
function getDupKey(f) {
  if (f.invoiceNo) return 'no:' + String(f.invoiceNo).replace(/\s+/g, '').trim().toUpperCase();
  if (f.sellerName && f.amountTax > 0) {
    return 'sum:' + String(f.sellerName).replace(/\s+/g, '').toUpperCase() + '|' + Number(f.amountTax).toFixed(2) + '|' + String(f.invoiceDate || '').replace(/\D/g, '');
  }
  return null;
}

// 标记重复发票：同 key 出现多次的文件置 _dup=true（保留第一份为原迹）
function updateDuplicateMarks() {
  var counts = {};
  for (var i = 0; i < S.files.length; i++) {
    var k = getDupKey(S.files[i]);
    if (k) counts[k] = (counts[k] || 0) + 1;
  }
  for (var i = 0; i < S.files.length; i++) {
    var k = getDupKey(S.files[i]);
    S.files[i]._dup = !!k && counts[k] > 1;
  }
  var dupCount = S.files.filter(function(f) { return f._dup; }).length;
  var dupEl = document.getElementById('duplicateCount');
  if (dupEl) dupEl.textContent = dupCount ? '(' + dupCount + ')' : '';
}

// 删除每组第一份之后的重复项。仅处理按发票号判定的可靠重复（no: key）；
// sum:（同销售方+金额+日期）疑似重复可能是同日同额的两张真发票，只标记不删除，
// 交由人工勾选处理。无 key、加载骨架一律不动。
// silent=true 为自动去重路径（识别完成后），删除后仍会 toast 告知用户。
function removeDuplicates(silent) {
  var seen = {};
  var removed = 0;
  var active = _activeFileIdx >= 0 ? S.files[_activeFileIdx] : null;
  S.files = S.files.filter(function(f) {
    if (f._placeholder || f._loading) return true;
    var key = getDupKey(f);
    if (!key || key.indexOf('no:') !== 0) return true;
    if (seen[key]) { removed++; return false; }
    seen[key] = true;
    return true;
  });
  _activeFileIdx = active ? S.files.indexOf(active) : -1;
  updateDuplicateMarks();
  if (!silent) {
    S.fileFilter = 'all';
    S.printedFilter = 'all';
    syncFilterButtons();
    renderFileList(); updatePreview(); updatePdfBtn(); updateSummaryBtn();
    toast(removed ? '已删除 ' + removed + ' 个重复项，保留每组第一份' : '未发现可删除的重复项');
  } else if (removed) {
    toast('已自动去重：删除 ' + removed + ' 个重复项（每组保留第一份）');
  }
  return removed;
}

// 一键勾选每组第一份之后的重复项，配合删除按钮安全去重。
// 仅勾选按发票号判定的可靠重复（no: key）；sum:（同销售方+金额+日期）疑似重复
// 只保留 ⚠ 标记供人工核对——同日同销售方同金额的两张真发票会被误判，不能自动勾选删除。
// 注意：此操作会覆盖用户原有勾选，toast 中明确提示。
function selectDuplicateExtras() {
  updateDuplicateMarks();
  if (!S.files.some(function(f) { return f._dup; })) {
    toast('未发现重复项');
    return 0;
  }
  var seen = {};
  var selected = 0;
  var suspected = 0;
  S.files.forEach(function(f) {
    if (f._loading || !f._dup) {
      f.checked = false;
      return;
    }
    var key = getDupKey(f);
    if (seen[key]) {
      if (key.indexOf('no:') === 0) { f.checked = true; selected++; }
      else { f.checked = false; suspected++; }
    } else {
      seen[key] = true;
      f.checked = false;
    }
  });
  S.fileFilter = 'duplicates';
  S.printedFilter = 'all';
  syncFilterButtons();
  renderFileList();
  if (selected) {
    toast('已覆盖原有勾选：选中 ' + selected + ' 个重复项（每组保留第一份），点击删除按钮即可去重' +
      (suspected ? '；另有 ' + suspected + ' 个疑似重复仅标记 ⚠，请人工核对' : ''));
  } else {
    toast('未发现可靠重复' + (suspected ? '；' + suspected + ' 个疑似重复（同销售方+金额+日期）仅标记 ⚠，请人工核对' : ''));
  }
  return selected;
}

function renderFileList() {
  updateDuplicateMarks();
  var list = document.getElementById('fileList');
  var scrollTop = list.scrollTop;
  var filtered = getFilteredFiles();
  var sel = filtered.filter(function(f) { return f.checked; }).length;
  document.getElementById('fileCount').textContent = filtered.length + ' 张，已选 ' + sel;
  var summaryEl = document.getElementById('amountSummary');
  if (!S.files.length) { list.innerHTML = ''; if (summaryEl) summaryEl.style.display = 'none'; updateAmountSummary(); return; }
  if (summaryEl) summaryEl.style.display = 'flex';

  // Snapshot and clear new-file IDs so animation only plays once
  var currentNewIds = _newFileIds;
  _newFileIds = {};

  var grid = S.fileView === 'grid';
  list.classList.toggle('grid', grid);

  list.innerHTML = S.files.map(function(f, i) {
    var cls = 'file-item';
    if (currentNewIds[f.id]) cls += ' entering';
    if (f._loading) cls += ' loading-item';
    if (i === _activeFileIdx) cls += ' active-item';
    var hidden = (S.fileFilter === 'duplicates' && !f._dup) ||
      (S.fileFilter !== 'duplicates' && ((S.printedFilter === 'printed' && !f._printed) || (S.printedFilter === 'unprinted' && f._printed)));
    var hideStyle = hidden ? ' style="display:none"' : '';
    if (grid) {
      var gcb = f.copies > 1 ? '<span class="copy-badge">' + f.copies + '份</span>' : '';
      var grb = f.rotation ? '<span class="rot-badge">' + f.rotation + '°</span>' : '';
      var gdupb = f._dup ? '<span class="dup-badge" title="检测到重复发票">⚠</span>' : '';
      var gab = buildAmtBadge(f);
      var gpd = f._printed ? '<span class="printed-dot" title="已打印">✓</span>' : '';
      var gsize = '<span class="card-size" title="文件大小">' + fmtSize(f.size) + '</span>';
      var gseller = f.sellerName ? '<div class="card-seller" title="' + escHtml(f.sellerName) + '"><span class="' + (f._isTicket ? 'ticket-badge' : f._isNonTax ? 'nontax-badge' : f._isToll ? 'toll-badge' : 'seller-badge') + '">' + escHtml(f.sellerName) + '</span></div>' : '';
      var gthumb = f._loading ? '' : (f.previewUrl ? '<img src="' + escHtml(f.previewUrl) + '">' : (f._xmlInvoice ? '<div class="xml-placeholder"><span class="xml-icon">XML</span>' + (f.invoiceNo ? '<span class="xml-no">' + escHtml(f.invoiceNo.slice(-4)) + '</span>' : '') + '</div>' : '\uD83D\uDCC4'));
      var gtype = f._xmlInvoice && f.invoiceType ? escHtml(f.invoiceType.replace(/^[^(]*\(/, '').replace(/\)$/, '') || f.invoiceType) : (f.type === 'jpeg' ? 'jpg' : escHtml(f.type));
      var gacts = '';
      if (!f._loading) {
        gacts = '<button class="ib card-ib' + (i === 0 ? ' disabled' : '') + '" onclick="moveFile(' + i + ',-1)" title="上移">\u25B2</button>' +
          '<button class="ib card-ib' + (i === S.files.length - 1 ? ' disabled' : '') + '" onclick="moveFile(' + i + ',1)" title="下移">\u25BC</button>' +
          '<button class="ib card-ib" onclick="rotFile(' + i + ')" title="旋转90°">\u21BB</button>' +
          '<button class="ib card-ib danger" onclick="rmFile(' + i + ')" title="删除">\u2715</button>';
      } else {
        gacts = '<button class="ib card-ib danger" onclick="rmFile(' + i + ')" title="删除">\u2715</button>';
      }
      return '<div class="' + cls + ' file-card" data-idx="' + i + '"' + hideStyle + ' onclick="clickFileItem(' + i + ',event)" ondblclick="openInvModal(' + i + ')">' +
        '<div class="file-thumb">' + gthumb + '<div class="type-badge">' + gtype + '</div>' +
        '<div class="file-check ' + (f.checked ? 'checked' : '') + '" onclick="togCheck(' + i + ')"></div>' +
        '<div class="card-actions">' + gacts + '</div></div>' +
        '<div class="card-name" title="' + escHtml(f.name) + '">' + escHtml(f.name) + '</div>' +
        gseller +
        '<div class="card-meta">' + gpd + gab + gcb + grb + gdupb + gsize + '</div></div>';
    }
    var cb = f.copies > 1 ? '<span class="copy-badge">' + f.copies + '份</span>' : '';
    var rb = f.rotation ? '<span class="rot-badge">' + f.rotation + '°</span>' : '';
    var dupb = f._dup ? '<span class="dup-badge" title="检测到重复发票：点击左上角「重复」筛选可一键勾选删除">⚠重复</span>' : '';
    var ab = buildAmtBadge(f);
    var sb = f.sellerName ? '<span class="' + (f._isTicket ? 'ticket-badge' : f._isNonTax ? 'nontax-badge' : f._isToll ? 'toll-badge' : 'seller-badge') + '" title="' + escHtml(f.sellerCreditCode || f.sellerName) + '">' + escHtml(f.sellerName) + '</span>' : '';
    // XSS FIX: escHtml(f.name) in both title and display text
    // XSS FIX: escHtml(f.previewUrl) in img src, escHtml(f.type) in type-badge
    var safePreviewUrl = escHtml(f.previewUrl || '');
    var safeType = escHtml(f.type === 'jpeg' ? 'jpg' : f.type);
    var typeBadgeText = f._xmlInvoice && f.invoiceType ? escHtml(f.invoiceType.replace(/^[^(]*\(/, '').replace(/\)$/, '') || f.invoiceType) : safeType;
    var thumbContent = f._loading ? '' : (f.previewUrl ? '<img src="' + safePreviewUrl + '">' : (f._xmlInvoice ? '<div class="xml-placeholder"><span class="xml-icon">XML</span>' + (f.invoiceNo ? '<span class="xml-no">' + escHtml(f.invoiceNo.slice(-4)) + '</span>' : '') + '</div>' : '\uD83D\uDCC4'));
    var pd = f._printed ? '<span class="printed-dot" title="已打印">✓</span>' : '';
    var metaActions = f._loading
      ? '<button class="ib danger" onclick="rmFile(' + i + ')">\u2715</button>'
      : '<div class="file-meta-left">' + pd + '<span class="file-size">' + fmtSize(f.size) + '</span>' + cb + rb + dupb + ab + '</div>' +
        '<div class="file-meta-sep"></div>' +
        '<div class="file-meta-right">' +
        '<button class="ib sort-btn' + (i === 0 ? ' disabled' : '') + '" onclick="moveFile(' + i + ',-1)" title="上移">\u25B2</button>' +
        '<button class="ib sort-btn' + (i === S.files.length - 1 ? ' disabled' : '') + '" onclick="moveFile(' + i + ',1)" title="下移">\u25BC</button>' +
        '<button class="ib" onclick="rotFile(' + i + ')" title="旋转90°">\u21BB</button><button class="ib danger" onclick="rmFile(' + i + ')">\u2715</button></div>';
    return '<div class="' + cls + '" data-idx="' + i + '" data-printed="' + (f._printed ? '1' : '0') + '"' + hideStyle + ' onclick="clickFileItem(' + i + ',event)" ondblclick="openInvModal(' + i + ')">' +
      '<div class="file-check ' + (f.checked ? 'checked' : '') + '" onclick="togCheck(' + i + ')"></div>' +
      '<div class="file-thumb">' + thumbContent + '<div class="type-badge">' + typeBadgeText + '</div></div>' +
      '<div class="file-info"><div class="file-name" title="' + escHtml(f.name) + '">' + escHtml(f.name) + '</div>' + (sb ? '<div class="file-seller" title="' + escHtml(f.sellerName) + '">' + sb + '</div>' : '') + '<div class="file-meta">' + metaActions + '</div></div>' +
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
function toggleFileView() {
  S.fileView = S.fileView === 'grid' ? 'list' : 'grid';
  syncFileViewBtn();
  saveSettings();
  renderFileList();
}
function syncFileViewBtn() {
  var btn = document.getElementById('fileViewBtn');
  if (!btn) return;
  var grid = S.fileView === 'grid';
  btn.textContent = grid ? '\u2630' : '\u25A6';
  btn.title = grid ? '切换列表视图' : '切换缩略图视图';
}
function toggleCopyMenu() {
  var menu = document.getElementById('copyMenu');
  menu.classList.toggle('hidden');
}
function toggleSortMenu() {
  var menu = document.getElementById('sortMenu');
  menu.classList.toggle('hidden');
}
function sortByDate(dir) {
  document.getElementById('sortMenu').classList.add('hidden');
  if (!S.files.length) return;
  S.files.sort(function(a, b) {
    var da = _parseDate(a.invoiceDate);
    var db = _parseDate(b.invoiceDate);
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    if (da < db) return -dir;
    if (da > db) return dir;
    return 0;
  });
  _activeFileIdx = -1;
  renderFileList();
  updatePreview();
}
function _parseDate(s) {
  if (!s || typeof s !== 'string') return null;
  var m = s.match(/(\d{4})[^\d]*(\d{1,2})[^\d]*(\d{1,2})/);
  if (!m) return null;
  var d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  if (isNaN(d.getTime())) return null;
  return d;
}
function setAllCopies(e, n) {
  e.stopPropagation();
  var sel = S.files.filter(function(f) { return f.checked; });
  if (!sel.length) { toast('请先选择发票'); document.getElementById('copyMenu').classList.add('hidden'); return; }
  sel.forEach(function(f) { f.copies = n; });
  document.getElementById('copyMenu').classList.add('hidden');
  renderFileList();
  updatePreview();
}
function togCheck(i) { S.files[i].checked = !S.files[i].checked; renderFileList(); updatePreview(); updateSummaryBtn(); }
function selectAll() { S.files.forEach(function(f) { f.checked = true; }); renderFileList(); updatePreview(); updateSummaryBtn(); }
function deselectAll() { S.files.forEach(function(f) { f.checked = false; }); renderFileList(); updatePreview(); updateSummaryBtn(); }
function deleteSelected() {
  if (!S.files.some(function(f) { return f.checked; })) return;
  var active = _activeFileIdx >= 0 ? S.files[_activeFileIdx] : null;
  S.files = S.files.filter(function(f) { return !f.checked; });
  _activeFileIdx = active ? S.files.indexOf(active) : -1;
  renderFileList(); updatePreview(); updatePdfBtn(); updateSummaryBtn();
}
function rmFile(i) { S.files.splice(i, 1); if (_activeFileIdx === i) _activeFileIdx = -1; else if (_activeFileIdx > i) _activeFileIdx--; renderFileList(); updatePreview(); updatePdfBtn(); updateSummaryBtn(); }
function rotFile(i) { S.files[i].rotation = (S.files[i].rotation + 90) % 360; renderFileList(); updatePreview(); }
function clearAll() {
  if (!S.files.length) return;
  if (!confirm('确认清除所有发票？')) return;
  S.files = [];
  _activeFileIdx = -1;
  _printedMap = {};
  saveSettings();
  renderFileList();
  updatePreview();
  updatePdfBtn();
  updateSummaryBtn();
}

// Click file item → navigate preview to the page containing this invoice
function clickFileItem(idx, event) {
  // Ignore clicks on checkbox, sort buttons, and action buttons
  if (event && (event.target.closest('.file-check') || event.target.closest('.sort-btn') || event.target.closest('button'))) return;
  var f = S.files[idx];
  if (f._loading || f._placeholder) return;

  _activeFileIdx = idx;

  // Auto-check if unchecked so the file appears in preview
  if (!f.checked) {
    f.checked = true;
  }

  // Find which page this file is on
  var activeFiles = getActiveFiles();
  var perPage = getPerPage(getSettings());
  var activeIdx = -1;
  for (var i = 0; i < activeFiles.length; i++) {
    if (activeFiles[i].id === f.id) { activeIdx = i; break; }
  }
  if (activeIdx >= 0) {
    S.currentPage = Math.floor(activeIdx / perPage);
    S.selectedSlot = activeIdx % perPage;
    updatePreview();
  } else {
    // 不参与排版的文件（如 XML 数电票）：清除预览槽位选中态并刷新面板
    S.selectedSlot = -1;
    var selEl = document.querySelector('.invoice-slot.selected');
    if (selEl) selEl.classList.remove('selected');
    updateAdjPanel();
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
  var perPage = getPerPage(getSettings());
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
// =====================================================
// File list sorting — move up / move down
// =====================================================
function moveFile(i, dir) {
  var target = i + dir;
  if (target < 0 || target >= S.files.length) return;
  var tmp = S.files[i];
  S.files[i] = S.files[target];
  S.files[target] = tmp;
  // Update active file index to follow the moved item
  if (_activeFileIdx === i) { _activeFileIdx = target; }
  else if (_activeFileIdx === target) { _activeFileIdx = i; }
  renderFileList();
  updatePreview();
  // Scroll to keep the moved item visible
  var list = document.getElementById('fileList');
  var items = list.querySelectorAll('.file-item');
  if (items[target]) items[target].scrollIntoView({ block: 'nearest' });
}

// =====================================================
// File list drag & drop sorting (#26)
// =====================================================
var _listDrag = null;
var _listDragBound = false;
var _listDragSuppressClick = false;
var _listDragHintShown = false;

function canListDrag() { return S.fileFilter === 'all' && S.printedFilter === 'all'; }

function initListDrag() {
  if (_listDragBound) return;
  var list = document.getElementById('fileList');
  if (!list) return;
  _listDragBound = true;
  list.addEventListener('mousedown', onListMouseDown);
  // 拖拽松手后浏览器派发的合成 click 会误触选中/弹窗，capture 阶段吞掉一次
  document.addEventListener('click', function(e) {
    if (!_listDragSuppressClick) return;
    _listDragSuppressClick = false;
    e.stopPropagation();
    e.preventDefault();
  }, true);
}

function onListMouseDown(e) {
  if (e.button !== 0 || !canListDrag()) return;
  var itemEl = e.target.closest ? e.target.closest('#fileList .file-item') : null;
  if (!itemEl) return;
  // 勾选框/操作按钮区域保持原有点击行为，不启动拖拽
  if (e.target.closest('.file-check') || e.target.closest('button')) return;
  var idx = parseInt(itemEl.dataset.idx);
  if (isNaN(idx)) return;
  var f = S.files[idx];
  if (!f || f._loading) return;
  _listDrag = { itemEl: itemEl, idx: idx, startX: e.clientX, startY: e.clientY, moved: false, dropEl: null, dropIdx: -1, dropZone: '' };
  e.preventDefault();
  document.addEventListener('mousemove', onListMouseMove);
  document.addEventListener('mouseup', onListMouseUp);
}

function onListMouseMove(e) {
  if (!_listDrag) return;
  if (!_listDrag.moved) {
    var dx = e.clientX - _listDrag.startX;
    var dy = e.clientY - _listDrag.startY;
    if (Math.abs(dx) < 4 && Math.abs(dy) < 4) return;
    _listDrag.moved = true;
    _listDrag.itemEl.classList.add('dragging');
    showListDragHint();
  }
  e.preventDefault();
  updateListDropTarget(e);
}

function onListMouseUp(e) {
  if (!_listDrag) return;
  var d = _listDrag;
  _listDrag = null;
  document.removeEventListener('mousemove', onListMouseMove);
  document.removeEventListener('mouseup', onListMouseUp);
  d.itemEl.classList.remove('dragging');
  clearListDropTarget(d.dropEl);
  if (!d.moved) return;
  _listDragSuppressClick = true;
  setTimeout(function() { _listDragSuppressClick = false; }, 0);
  if (d.dropIdx >= 0 && d.dropIdx !== d.idx) {
    if (d.dropZone === 'before' || d.dropZone === 'after') {
      moveFileTo(d.idx, d.dropIdx, d.dropZone);
    } else {
      swapFiles(d.idx, d.dropIdx);
    }
  }
}

function findNearestListItem(x, y) {
  var best = null, bestD = Infinity;
  var items = document.querySelectorAll('#fileList .file-item');
  for (var i = 0; i < items.length; i++) {
    var r = items[i].getBoundingClientRect();
    var dx = Math.max(r.left - x, 0, x - r.right);
    var dy = Math.max(r.top - y, 0, y - r.bottom);
    var d = dx * dx + dy * dy;
    if (d < bestD) { bestD = d; best = items[i]; }
  }
  if (!best || bestD > 6400) return null;
  if (best.dataset.idx === undefined || parseInt(best.dataset.idx) === _listDrag.idx) return null;
  return best;
}

function updateListDropTarget(e) {
  var el = document.elementFromPoint(e.clientX, e.clientY);
  el = el && el.closest ? el.closest('.file-item') : null;
  if (el && !el.closest('#fileList')) el = null;
  var idx = el ? parseInt(el.dataset.idx) : -1;
  if (isNaN(idx)) idx = -1;
  if (idx < 0 || idx === _listDrag.idx) {
    var near = findNearestListItem(e.clientX, e.clientY);
    if (near) { el = near; idx = parseInt(el.dataset.idx); }
  }
  if (idx < 0) { el = null; idx = -1; }
  var zone = '';
  if (el) {
    var r = el.getBoundingClientRect();
    var ratio = (e.clientY - r.top) / (r.height || 1);
    zone = ratio < 0.25 ? 'before' : ratio > 0.75 ? 'after' : 'swap';
  }
  if (el === _listDrag.dropEl && zone === _listDrag.dropZone) return;
  clearListDropTarget(_listDrag.dropEl);
  _listDrag.dropEl = el;
  _listDrag.dropIdx = idx;
  _listDrag.dropZone = zone;
  if (!el) return;
  if (zone === 'swap') { el.classList.add('drop-target'); }
  else {
    el.classList.add('drop-insert');
    el.classList.add(zone === 'before' ? 'drop-at-start' : 'drop-at-end');
  }
}

function clearListDropTarget(el) {
  if (!el) return;
  el.classList.remove('drop-target', 'drop-insert', 'drop-at-start', 'drop-at-end');
}

function showListDragHint() {
  if (_listDragHintShown) return;
  _listDragHintShown = true;
  try {
    if (localStorage.getItem('ticketchan-list-drag-hint')) return;
    localStorage.setItem('ticketchan-list-drag-hint', '1');
  } catch (err) { return; }
  toast('拖到列表项边缘 = 顺位插入，拖到中间 = 两张对调', 4000);
}

function swapFiles(ia, ib) {
  if (ia === ib || ia < 0 || ib < 0 || ia >= S.files.length || ib >= S.files.length) return;
  var tmp = S.files[ia]; S.files[ia] = S.files[ib]; S.files[ib] = tmp;
  if (_activeFileIdx === ia) { _activeFileIdx = ib; }
  else if (_activeFileIdx === ib) { _activeFileIdx = ia; }
  renderFileList(); updatePreview(); scrollToListItem(ib);
}

function moveFileTo(fromIdx, toIdx, zone) {
  if (fromIdx === toIdx || fromIdx < 0 || toIdx < 0 || fromIdx >= S.files.length || toIdx >= S.files.length) return;
  var a = S.files[fromIdx], b = S.files[toIdx];
  if (!a || !b || a === b) return;
  S.files.splice(fromIdx, 1);
  var insertAt = zone === 'after' ? S.files.indexOf(b) + 1 : S.files.indexOf(b);
  S.files.splice(insertAt, 0, a);
  if (insertAt === fromIdx) return;
  _activeFileIdx = S.files.indexOf(a);
  renderFileList(); updatePreview(); scrollToListItem(insertAt);
}

function scrollToListItem(idx) {
  var list = document.getElementById('fileList');
  if (!list) return;
  var items = list.querySelectorAll('.file-item');
  if (items[idx]) items[idx].scrollIntoView({ block: 'nearest' });
}

// Amount statistics
function updateAmountSummary() {
  var el = document.getElementById('amountSummary');
  if (!el) return;
  var checked = S.files.filter(function(f) { return f.checked; });
  var taxTotal = checked.reduce(function(s, f) { return s + (f.amountTax || 0); }, 0);
  var noTaxTotal = checked.reduce(function(s, f) { return s + (f.amountNoTax || 0); }, 0);
  var taxAmtTotal = checked.reduce(function(s, f) { return s + (f.taxAmount || 0); }, 0);
  var withAmt = checked.filter(function(f) { return (f.amountTax || f.amountNoTax) > 0; }).length;
  var warnAmt = checked.filter(function(f) { return f._amtValidationFail; }).length;

  // Container visibility: show when files exist, hide when empty
  // (renderFileList handles the initial show/hide; we only override when truly empty)
  if (!S.files.length) { el.style.display = 'none'; return; }
  el.style.display = '';

  if (checked.length === 0) {
    var textEl = document.getElementById('amountSummaryText');
    if (textEl) textEl.innerHTML = '';
    return;
  }

  var countHtml = '<span class="amt-count">' + withAmt + '/' + checked.length + ' 张已识别</span>';
  if (warnAmt > 0) {
    countHtml += '<span class="amt-warn-count" title="' + warnAmt + ' 张发票金额校验失败（含税≠不含税+税额）">' + warnAmt + ' 张校验异常</span>';
  }
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
    ? '<span style="font-size:10px;color:var(--text-muted);margin-left:6px">' + sellerNames.length + '个销售方</span>'
    : '';
  var textEl = document.getElementById('amountSummaryText');
  if (textEl) textEl.innerHTML = countHtml + amtHtml + sellerHtml;

  // Total amount is already shown in amountSummary (bottom-left), no need to duplicate in statusbar
}

// Invoice modal
function openInvModal(i) {
  if (S.files[i]._loading) return; // Don't open modal for loading placeholders
  S.editIdx = i; var f = S.files[i];
  var _fw = 'width:140px;flex:none;text-align:right;font-size:12px';
  var _fwm = _fw + ';font-family:monospace';
  var mRF = function(label, html) { return '<div class="modal-row"><label class="modal-lbl">' + label + '</label><div class="modal-ctrl end">' + html + '</div></div>'; };
  var mRA = function(label, html) { return '<div class="modal-row"><label class="modal-lbl">' + label + '</label><div class="modal-ctrl">' + html + '</div></div>'; };
  document.getElementById('invModalBody').innerHTML =
    '<div style="font-size:13px;padding:8px 10px;background:var(--surface2);border-radius:6px;margin-bottom:10px">\uD83D\uDCC4 ' + escHtml(f.name) + '</div>' +
    mRF('排版份数', '<button class="btn btn-sm btn-icon" onclick="changeModalCopies(-1)">\u2212</button><input type="number" id="mCopies" value="' + f.copies + '" min="1" max="99" style="width:52px;text-align:center;flex:none"><button class="btn btn-sm btn-icon" onclick="changeModalCopies(1)">+</button>') +
    '<div style="font-size:10px;color:var(--text-muted);margin:-6px 0 8px 76px">同一发票在布局中占几个位置</div>' +
    mRF('含税价', '<span style="font-size:14px;font-weight:600;color:var(--success);flex-shrink:0">\u00A5</span><input type="number" id="mAmountTax" value="' + (f.amountTax || '') + '" min="0" step="0.01" placeholder="0.00" style="' + _fw + '">') +
    mRF('不含税', '<span style="font-size:14px;font-weight:600;color:var(--text-muted);flex-shrink:0">\u00A5</span><input type="number" id="mAmountNoTax" value="' + (f.amountNoTax || '') + '" min="0" step="0.01" placeholder="0.00" style="' + _fw + '">') +
    mRF('税额', '<span style="font-size:14px;font-weight:600;color:var(--warning,orange);flex-shrink:0">\u00A5</span><input type="number" id="mTaxAmount" value="' + (f.taxAmount || '') + '" min="0" step="0.01" placeholder="0.00" style="' + _fw + '">') +
    mRA('发票号码', '<input type="text" id="mInvoiceNo" value="' + escHtml(f.invoiceNo || '') + '" placeholder="自动识别" class="mono-input">') +
    mRA('开票日期', '<input type="text" id="mInvoiceDate" value="' + escHtml(f.invoiceDate || '') + '" placeholder="自动识别">') +
    mRA('购买方', '<input type="text" id="mBuyer" value="' + escHtml(f.buyerName || '') + '" placeholder="自动识别">') +
    mRA('购方代码', '<input type="text" id="mBuyerCreditCode" value="' + escHtml(f.buyerCreditCode || '') + '" placeholder="自动识别" class="mono-input">') +
    mRA('销售方', '<input type="text" id="mSeller" value="' + escHtml(f.sellerName || '') + '" placeholder="自动识别">') +
    mRA('信用代码', '<input type="text" id="mCreditCode" value="' + escHtml(f.sellerCreditCode || '') + '" placeholder="自动识别" class="mono-input">') +
    mRF('旋转', '<select id="mRot" style="width:140px;flex:none"><option value="0" ' + (f.rotation === 0 ? 'selected' : '') + '>不旋转</option><option value="90" ' + (f.rotation === 90 ? 'selected' : '') + '>90\u00B0</option><option value="180" ' + (f.rotation === 180 ? 'selected' : '') + '>180\u00B0</option><option value="270" ' + (f.rotation === 270 ? 'selected' : '') + '>270\u00B0</option></select>') +
    '<div style="border-top:1px dashed var(--border);margin-top:4px;padding-top:8px">' +
    '<div style="font-size:11px;font-weight:700;color:var(--text-secondary);margin-bottom:6px">🎯 单票调整</div>' +
    mRF('缩放', '<input type="number" id="mSlotScale" value="' + Math.round((f.slotScale || 1) * 100) + '" min="20" max="300" style="' + _fw + '"><span style="font-size:11px;color:var(--text-muted);width:16px;flex-shrink:0;text-align:left">%</span>') +
    mRF('X偏移', '<input type="number" id="mSlotOffX" value="' + (f.slotOffsetX || 0) + '" min="-50" max="50" step="0.5" style="' + _fw + '"><span style="font-size:11px;color:var(--text-muted);width:16px;flex-shrink:0;text-align:left">mm</span>') +
    mRF('Y偏移', '<input type="number" id="mSlotOffY" value="' + (f.slotOffsetY || 0) + '" min="-50" max="50" step="0.5" style="' + _fw + '"><span style="font-size:11px;color:var(--text-muted);width:16px;flex-shrink:0;text-align:left">mm</span>') +
    '</div>';
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
  // Per-slot adjustments
  f.slotScale = Math.max(0.2, Math.min(3.0, (parseInt(document.getElementById('mSlotScale').value) || 100) / 100));
  f.slotOffsetX = parseFloat(document.getElementById('mSlotOffX').value) || 0;
  f.slotOffsetY = parseFloat(document.getElementById('mSlotOffY').value) || 0;
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
// Per-slot Adjustment
// =====================================================
function selectSlot(idx) {
  S.selectedSlot = idx;
  updateAdjPanel();
  // Highlight in preview
  document.querySelectorAll('.invoice-slot').forEach(function(el) { el.classList.remove('selected'); });
  if (idx >= 0) {
    var slotEl = document.querySelector('.invoice-slot[data-slot-idx="' + idx + '"]');
    if (slotEl) slotEl.classList.add('selected');
  }
  syncSidebarToSelectedSlot();
}

// 选中槽位 → 同步左侧列表高亮并滚动定位（列表与版面双向联动）
function syncSidebarToSelectedSlot() {
  var f = getSelectedFileObj();
  if (!f) return;
  var idx = S.files.indexOf(f);
  if (idx < 0) return;
  _activeFileIdx = idx;
  updateActiveFileHighlight();
  var list = document.getElementById('fileList');
  if (!list) return;
  var el = list.querySelector('.file-item[data-idx="' + idx + '"]');
  if (el) el.scrollIntoView({ block: 'nearest' });
}

function getSelectedFileObj() {
  if (S.selectedSlot < 0) return null;
  var files = getActiveFiles();
  var settings = getSettings();
  var perPage = getPerPage(settings);
  var pageStart = S.currentPage * perPage;
  var fileIdx = pageStart + S.selectedSlot;
  return fileIdx < files.length ? files[fileIdx] : null;
}

function updateAdjPanel() {
  var f = getSelectedFileObj();
  var empty = document.getElementById('adjEmpty');
  var content = document.getElementById('adjContent');
  if (!f) {
    empty.style.display = '';
    content.style.display = 'none';
    return;
  }
  empty.style.display = 'none';
  content.style.display = '';
  document.getElementById('adjFileName').textContent = f.name || '未命名';
  document.getElementById('adjScale').value = Math.round((f.slotScale || 1) * 100);
  document.getElementById('adjScaleN').value = Math.round((f.slotScale || 1) * 100);
  document.getElementById('adjOffX').value = f.slotOffsetX || 0;
  document.getElementById('adjOffXN').value = f.slotOffsetX || 0;
  document.getElementById('adjOffY').value = f.slotOffsetY || 0;
  document.getElementById('adjOffYN').value = f.slotOffsetY || 0;
}

function onAdjScaleChange() {
  var f = getSelectedFileObj();
  if (!f) return;
  f.slotScale = Math.max(0.2, Math.min(3.0, parseInt(document.getElementById('adjScale').value) / 100));
  updatePreview();
}

function onAdjOffsetChange() {
  var f = getSelectedFileObj();
  if (!f) return;
  f.slotOffsetX = parseFloat(document.getElementById('adjOffX').value) || 0;
  f.slotOffsetY = parseFloat(document.getElementById('adjOffY').value) || 0;
  updatePreview();
}

function resetSlotAdj() {
  var f = getSelectedFileObj();
  if (!f) return;
  f.slotScale = 1;
  f.slotOffsetX = 0;
  f.slotOffsetY = 0;
  updateAdjPanel();
  updatePreview();
}

function applySlotAdjToAll() {
  var f = getSelectedFileObj();
  if (!f) return;
  var scale = f.slotScale, ox = f.slotOffsetX, oy = f.slotOffsetY;
  S.files.forEach(function(file) {
    file.slotScale = scale;
    file.slotOffsetX = ox;
    file.slotOffsetY = oy;
  });
  updatePreview();
  toast('已应用到全部 ' + S.files.length + ' 张发票');
}

/**
 * Quick alignment: snap the selected invoice to a slot edge or center.
 * @param {string} alignH - 'left' | 'center' | 'right'
 * @param {string} alignV - 'top' | 'center' | 'bottom'
 */
function setSlotAlignment(alignH, alignV) {
  var f = getSelectedFileObj();
  if (!f) return;

  var settings = getSettings();
  var layout = calculateLayout(settings);
  var slot = layout.slots[S.selectedSlot];
  if (!slot) return;

  // Use unrotated image dimensions — same as renderPage.
  // renderPage computes wrapper box size from f.ow/f.oh (unrotated),
  // then applies rotation as a CSS transform. Alignment must match.
  var imgObjW = f.ow || 1;
  var imgObjH = f.oh || 1;

  var slotW_mm = slot.w / MM2PX;
  var slotH_mm = slot.h / MM2PX;

  // Calculate contained wrapper dimensions in mm (mirrors renderPage)
  var containedW_mm, containedH_mm;
  if (settings.fitMode === 'original') {
    // original mode: image displays at native resolution; for alignment
    // we convert native px→mm using the render DPI the image was produced at.
    // If renderDpi is not set, fall back to PDF_RENDER_DPI (300).
    var rDpi = f.renderDpi || 300;
    var oPxPerMm = rDpi / 25.4;
    containedW_mm = imgObjW / oPxPerMm;
    containedH_mm = imgObjH / oPxPerMm;
  } else if (settings.fitMode === 'fill') {
    containedW_mm = slotW_mm;
    containedH_mm = slotH_mm;
  } else {
    // contain / custom: aspect-ratio fit inside slot
    // Both slot.w and imgObjW are in CSS coordinate space; ratio is correct.
    var fitScale = Math.min(slot.w / imgObjW, slot.h / imgObjH);
    containedW_mm = (imgObjW * fitScale) / MM2PX;
    containedH_mm = (imgObjH * fitScale) / MM2PX;
  }

  // Effective visual size = contained wrapper size × per-slot scale × custom scale.
  // CSS scale() transforms from center; the wrapper box stays at containedW_mm×containedH_mm
  // but the visible content is containedW_mm × effectiveScale.
  // Alignment must account for the actual visual footprint.
  var perScale = f.slotScale || 1;
  var customScale = (settings.fitMode === 'custom') ? (settings.customScale || 1) : 1;
  var effectiveScale = perScale * customScale;
  var gapX = (slotW_mm - containedW_mm * effectiveScale) / 2;
  var gapY = (slotH_mm - containedH_mm * effectiveScale) / 2;

  // Offset to move wrapper from centered position to target alignment
  var offsetX = 0, offsetY = 0;
  if (alignH === 'left')  offsetX = -gapX;
  if (alignH === 'right') offsetX =  gapX;
  if (alignV === 'top')   offsetY = -gapY;
  if (alignV === 'bottom') offsetY =  gapY;

  f.slotOffsetX = Math.round(offsetX * 10) / 10;
  f.slotOffsetY = Math.round(offsetY * 10) / 10;

  updateAdjPanel();
  updatePreview();
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
  saveSettings();
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
  var isOn = !S.feat[k]; // 切换后的状态
  S.feat[k] = isOn;
  btn.classList.toggle('on', isOn);

  var targets = ['pageNum', 'printDate', 'footer', 'customFM'];
  var isTarget = targets.indexOf(k) >= 0;

  if (isTarget) {
    // 按行数计算页脚边距：pageNum+printDate 共享一行，footerText 单独一行
    var lineCount = (S.feat.pageNum || S.feat.printDate ? 1 : 0) + (S.feat.footer ? 1 : 0);
    var fmRow = document.getElementById('footerMarginRow');
    var cfmRow = document.getElementById('customFMRow');

    // "自定义下边距"开关行：任何页脚功能开启时显示
    if (cfmRow) cfmRow.style.display = lineCount > 0 ? 'flex' : 'none';

    if (S.feat.customFM && lineCount > 0) {
      // 自定义下边距模式：显示滑块，自动设置最小值
      var minFM = lineCount >= 2 ? 16 : 8;
      var currentFM = parseFloat(document.getElementById('footerMargin').value) || 0;
      if (currentFM < minFM) {
        document.getElementById('footerMargin').value = minFM;
        document.getElementById('footerMarginN').value = minFM;
      }
      if (fmRow) fmRow.style.display = 'flex';
    } else {
      // 默认模式或全部关闭：隐藏滑块
      if (fmRow) fmRow.style.display = 'none';
    }
  }

  if (k === 'watermark') document.getElementById('wmOpts').style.display = S.feat[k] ? 'block' : 'none';
  if (k === 'trimWhite' && S.feat[k]) processTrim();
  if (k === 'footer') {
    document.getElementById('footerOpts').style.display = S.feat[k] ? 'block' : 'none';
  }
  saveSettings();
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
  saveSettings();
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
  saveSettings();
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
function onFitChange() {
  var isCustom = document.getElementById('fitMode').value === 'custom';
  document.getElementById('customScaleRow').style.display = isCustom ? 'flex' : 'none';
  document.getElementById('customScaleHint').style.display = isCustom ? 'block' : 'none';
  updatePreview();
}
function setMP(t, b, l, r) {
  [['marginTop', 'marginTopN', t], ['marginBottom', 'marginBottomN', b], ['marginLeft', 'marginLeftN', l], ['marginRight', 'marginRightN', r]].forEach(function(arr) {
    document.getElementById(arr[0]).value = arr[2]; document.getElementById(arr[1]).value = arr[2];
  });
  updatePreview();
}
function changeCopies(d) { var e = document.getElementById('copies'); e.value = Math.max(1, Math.min(99, parseInt(e.value) + d)); updatePreview(); }

// Trim whitespace — client-side canvas implementation
async function trimOneImage(dataUrl) {
  return new Promise(function(resolve) {
    var img = new Image();
    img.onload = function() {
      var canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      var ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      try {
        var data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
        var top = 0, bottom = canvas.height - 1, left = 0, right = canvas.width - 1;
        var threshold = 245;
        function rowBlank(y) {
          for (var x = 0; x < canvas.width; x++) {
            var i = (y * canvas.width + x) * 4;
            if (data[i] < threshold || data[i+1] < threshold || data[i+2] < threshold) return false;
          }
          return true;
        }
        function colBlank(x) {
          for (var y = 0; y < canvas.height; y++) {
            var i = (y * canvas.width + x) * 4;
            if (data[i] < threshold || data[i+1] < threshold || data[i+2] < threshold) return false;
          }
          return true;
        }
        while (top < canvas.height && rowBlank(top)) top++;
        while (bottom > top && rowBlank(bottom)) bottom--;
        while (left < canvas.width && colBlank(left)) left++;
        while (right > left && colBlank(right)) right--;
        if (top >= bottom || left >= right) { resolve(dataUrl); return; }
        var pad = 4;
        top = Math.max(0, top - pad); bottom = Math.min(canvas.height - 1, bottom + pad);
        left = Math.max(0, left - pad); right = Math.min(canvas.width - 1, right + pad);
        var w = right - left + 1, h = bottom - top + 1;
        var out = document.createElement('canvas');
        out.width = w; out.height = h;
        var octx = out.getContext('2d');
        octx.drawImage(canvas, left, top, w, h, 0, 0, w, h);
        resolve(out.toDataURL('image/jpeg', 0.9));
      } catch (e) {
        resolve(dataUrl);
      }
    };
    img.onerror = function() { resolve(dataUrl); };
    img.src = dataUrl;
  });
}

async function processTrim() {
  showLoading('裁剪白边...');
  try {
    for (var i = 0; i < S.files.length; i++) {
      var f = S.files[i];
      if (f.previewUrl && !f.trimmedUrl) {
        f.trimmedUrl = await trimOneImage(f.previewUrl);
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

// Auto-calculate footer margin based on line count
// Must be >= actual text height (3mm bottom + lineCount * 5mm line height)
function _autoFooterMargin() {
  var lineCount = (S.feat.pageNum || S.feat.printDate ? 1 : 0) + (S.feat.footer ? 1 : 0);
  return 3 + lineCount * 5; // matches text layout: 3mm bottom padding + 5mm per line
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
    colorMode: 'color',
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
    footerText: S.feat.footer ? document.getElementById('footerText').value : '',
    footerMargin: (S.feat.pageNum || S.feat.printDate || S.feat.footer) ? (S.feat.customFM ? parseFloat(document.getElementById('footerMargin').value) || 0 : _autoFooterMargin()) : 0,
    customFM: S.feat.customFM,
    copies: parseInt(document.getElementById('copies').value) || 1
  };
}

// Get checked files WITHOUT copies expansion (for summary table, etc.)
function getCheckedFiles() {
  return S.files.filter(function(f) { return f.checked && !f._loading; });
}

function markFilesAsPrinted(files) {
  files.forEach(function(f) {
    f._printed = true;
    var key = f._filePath || f._pdfPath;
    if (key) _printedMap[key] = true;
  });
  saveSettings();
  renderFileList();
}

function getActiveFiles() {
  // 占位对象（_placeholder）虽未勾选，也参与排版占槽位（版面留白）
  var files = S.files.filter(function(f) { return (f.checked || f._placeholder) && !f._loading && !f._xmlInvoice; });
  var exp = [];
  files.forEach(function(f) { for (var c = 0; c < Math.max(1, f.copies); c++) exp.push(f); });
  return exp;
}

function buildPages(files, settings) {
  var perPage = getPerPage(settings);
  var pages = [];
  for (var i = 0; i < files.length; i += perPage) pages.push(files.slice(i, i + perPage));
  return pages;
}

// =====================================================
// Preview & Navigation
// =====================================================
var _saveTimer = null;
function updatePreview() {
  if (_saveTimer) clearTimeout(_saveTimer);
  _saveTimer = setTimeout(saveSettings, 500);
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
  if (typeof updateAdjPanel === 'function') updateAdjPanel();
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
function prevPage() { if (S.currentPage > 0) { S.currentPage--; S.selectedSlot = -1; updatePreview(); } }
function nextPage() { if (S.currentPage < S.totalPages - 1) { S.currentPage++; S.selectedSlot = -1; updatePreview(); } }
function gotoPage(i) { S.currentPage = i; S.selectedSlot = -1; updatePreview(); }
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
  if (!e.target.closest('.copy-ctrl')) {
    var cm = document.getElementById('copyMenu');
    if (cm) cm.classList.add('hidden');
  }
  if (!e.target.closest('.zoom-ctrl')) {
    var zm = document.getElementById('zoomMenu');
    if (zm) zm.classList.add('hidden');
  }
});
function updatePdfBtn() { var has = S.files.some(function(f) { return f.checked; }); document.getElementById('pdfBtn').disabled = !has; document.getElementById('rasterPdfBtn').disabled = !has; }
function updateSummaryBtn() { var btn = document.getElementById('summaryBtn'); if (btn) btn.disabled = !S.files.some(function(f) { return f.checked; }); }

// =====================================================
// Save settings & Preferences
// =====================================================
function saveSettings() {
  var o = {
    layout: { cols: S.layout.cols, rows: S.layout.rows },
    fileView: S.fileView,
    paperSize: document.getElementById('paperSize').value,
    orientation: document.getElementById('orientation').value,
    customW: document.getElementById('customW').value,
    customH: document.getElementById('customH').value,
    marginTop: document.getElementById('marginTop').value,
    marginBottom: document.getElementById('marginBottom').value,
    marginLeft: document.getElementById('marginLeft').value,
    marginRight: document.getElementById('marginRight').value,
    gapH: document.getElementById('gapH').value,
    gapV: document.getElementById('gapV').value,
    fitMode: document.getElementById('fitMode').value,
    customScale: document.getElementById('customScale').value,
    globalRotation: document.getElementById('globalRotation').value,
    copies: document.getElementById('copies').value,
    feat: {}
  };
  var featKeys = ['cutline','number','border','trimWhite','watermark','pageNum','printDate','footer','customFM','slotAdjMemory'];
  featKeys.forEach(function(k) { o.feat[k] = S.feat[k]; });
  // Save per-file slot adjustments when memory is enabled
  if (S.feat.slotAdjMemory) {
    var adjMap = {};
    S.files.forEach(function(f) {
      if (f.name && (f.slotScale !== undefined || f.slotOffsetX !== undefined || f.slotOffsetY !== undefined)) {
        adjMap[f.name] = {
          scale: f.slotScale || 1,
          offX: f.slotOffsetX || 0,
          offY: f.slotOffsetY || 0
        };
      }
    });
    if (Object.keys(adjMap).length > 0) {
      o.fileAdjustments = adjMap;
    }
  }
  // Always save watermark/footer values so they survive feature toggles
  o.wmText = document.getElementById('wmText').value;
  o.wmOpacity = document.getElementById('wmOpacity').value;
  o.wmColor = document.getElementById('wmColor').value;
  o.wmAngle = document.getElementById('wmAngle').value;
  o.wmSize = document.getElementById('wmSize').value;
  o.footerText = document.getElementById('footerText').value;
  o.footerMargin = document.getElementById('footerMargin').value;
  if (_summaryActiveCols && _summaryActiveCols.length > 0) {
    o.summaryCols = _summaryActiveCols;
  }
  // Persist per-file notes (keyed by file name)
  var notesMap = {};
  S.files.forEach(function(f) { if (f.note && f.name) notesMap[f.name] = f.note; });
  if (Object.keys(notesMap).length > 0) o.summaryNotes = notesMap;
  // Save printed state
  var printedMap = {};
  S.files.forEach(function(f) {
    var key = f._filePath || f._pdfPath;
    if (key && f._printed) printedMap[key] = true;
  });
  o.printedMap = printedMap;
  try { localStorage.setItem('ticketchan-settings', JSON.stringify(o)); } catch(e) {}
}

function loadSettings() {
  var raw;
  try { raw = localStorage.getItem('ticketchan-settings'); } catch(e) { return; }
  if (!raw) return;
  var o;
  try { o = JSON.parse(raw); } catch(e) { return; }
  if (o.layout) {
    S.layout = { cols: o.layout.cols || 1, rows: o.layout.rows || 1 };
    document.getElementById('customRows').value = S.layout.rows;
    document.getElementById('customCols').value = S.layout.cols;
    document.querySelectorAll('.go').forEach(function(e) {
      e.classList.remove('active');
      if (parseInt(e.dataset.cols) === S.layout.cols && parseInt(e.dataset.rows) === S.layout.rows) e.classList.add('active');
    });
    syncToolbarHighlight(S.layout.cols, S.layout.rows);
  }
  if (o.fileView === 'grid' || o.fileView === 'list') {
    S.fileView = o.fileView;
    syncFileViewBtn();
  }
  if (o.paperSize) { document.getElementById('paperSize').value = o.paperSize; onPaperChange(); }
  if (o.orientation) document.getElementById('orientation').value = o.orientation;
  if (o.customW) document.getElementById('customW').value = o.customW;
  if (o.customH) document.getElementById('customH').value = o.customH;
  var sliders = ['marginTop','marginBottom','marginLeft','marginRight','gapH','gapV','customScale'];
  sliders.forEach(function(id) {
    if (o[id] != null) {
      document.getElementById(id).value = o[id];
      var nId = id + 'N';
      var nEl = document.getElementById(nId);
      if (nEl) nEl.value = o[id];
    }
  });
  if (o.fitMode) { document.getElementById('fitMode').value = o.fitMode; onFitChange(); }
  if (o.globalRotation) document.getElementById('globalRotation').value = o.globalRotation;
  if (o.copies) document.getElementById('copies').value = o.copies;
  if (o.feat) {
    var featMap = {
      cutline: 'toggleCutline', number: 'toggleNumber', border: 'toggleBorder',
      trimWhite: 'toggleTrimWhite', watermark: 'toggleWatermark',
      pageNum: 'togglePageNum', printDate: 'toggleDate',
      footer: 'toggleFooter', customFM: 'toggleCustomFM',
      slotAdjMemory: 'toggleSlotAdjMemory'
    };
    Object.keys(featMap).forEach(function(k) {
      if (o.feat[k] != null) {
        S.feat[k] = o.feat[k];
        var btn = document.getElementById(featMap[k]);
        if (btn) btn.classList.toggle('on', S.feat[k]);
      }
    });
    if (S.feat.watermark) {
      document.getElementById('wmOpts').style.display = 'block';
    }
    if (S.feat.footer) {
      document.getElementById('footerOpts').style.display = 'block';
    }
    var lineCount = (S.feat.pageNum || S.feat.printDate ? 1 : 0) + (S.feat.footer ? 1 : 0);
    if (S.feat.customFM && lineCount > 0) {
      document.getElementById('customFMRow').style.display = 'flex';
      document.getElementById('footerMarginRow').style.display = 'flex';
    } else if (lineCount > 0) {
      document.getElementById('customFMRow').style.display = 'flex';
    }
  }
  // Always restore watermark/footer values (even when features are off,
  // so the values are ready when user enables them later)
  if (o.wmText != null) document.getElementById('wmText').value = o.wmText;
  if (o.wmOpacity != null) { document.getElementById('wmOpacity').value = o.wmOpacity; document.getElementById('wmOpacityN').value = o.wmOpacity; }
  if (o.wmColor) document.getElementById('wmColor').value = o.wmColor;
  if (o.wmAngle != null) { document.getElementById('wmAngle').value = o.wmAngle; document.getElementById('wmAngleN').value = o.wmAngle; }
  if (o.wmSize != null) { document.getElementById('wmSize').value = o.wmSize; document.getElementById('wmSizeN').value = o.wmSize; }
  if (o.footerText != null) document.getElementById('footerText').value = o.footerText;
  if (o.footerMargin != null) {
    document.getElementById('footerMargin').value = o.footerMargin;
    document.getElementById('footerMarginN').value = o.footerMargin;
  }
  // Restore summary table column selection
  if (o.summaryCols && Array.isArray(o.summaryCols) && o.summaryCols.length > 0) {
    _summaryActiveCols = o.summaryCols;
    // v2.0.6 migration: ensure note column is included for existing users
    if (_summaryActiveCols.indexOf('note') < 0) _summaryActiveCols.push('note');
  }
  // Restore per-file notes (applied when files are added)
  S._notesMap = o.summaryNotes || {};
  // Load saved per-file slot adjustments (applied when files are added)
  S._fileAdjMap = (o.fileAdjustments && S.feat.slotAdjMemory) ? o.fileAdjustments : {};
  // Restore printed state (always, regardless of switch)
  if (o.printedMap) _printedMap = o.printedMap;
  else _printedMap = {};
}

function togglePref(k, btn) {
  S.feat[k] = !S.feat[k];
  btn.classList.toggle('on', S.feat[k]);
  if (k === 'pdfTextEnabled') {
    try { localStorage.setItem('ticketchan-pdf-text-enabled', S.feat[k] ? '1' : '0'); } catch(e) {}
  }
  saveSettings();
}

function applyTheme() {
  var theme = document.getElementById('themeMode').value;
  if (theme === 'dark') { document.documentElement.classList.add('dark'); }
  else { document.documentElement.classList.remove('dark'); }
  try { localStorage.setItem('ticketchan-theme', theme); } catch(e) {}
}

function exportSettings() {
  var data = { layout: S.layout, feat: S.feat, paperSize: document.getElementById('paperSize').value, orientation: document.getElementById('orientation').value, copies: document.getElementById('copies').value };
  var blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = '发票酱设置.json'; a.click();
  toast('设置已导出');
}

function resetSettings() {
  if (!confirm('确认恢复所有默认设置？')) return;
  S.layout = { cols: 1, rows: 1 };
  S.feat = { cutline: true, number: false, border: false, trimWhite: false, watermark: false, footer: false, customFM: false, pageNum: false, printDate: false, pdfTextEnabled: true, slotAdjMemory: false };
  S.viewZoom = 0;
  S.fileView = 'list';
  S.fileFilter = 'all';
  S.printedFilter = 'all';
  syncFileViewBtn();
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
  document.getElementById('customW').value = 210;
  document.getElementById('customH').value = 297;
  document.getElementById('customScale').value = 100; document.getElementById('customScaleN').value = 100;
  document.getElementById('customPaperRow').style.display = 'none';
  document.getElementById('customScaleRow').style.display = 'none';
  document.getElementById('wmOpts').style.display = 'none';
  document.getElementById('wmText').value = '已打印';
  document.getElementById('wmOpacity').value = 20; document.getElementById('wmOpacityN').value = 20;
  document.getElementById('wmColor').value = '#ff0000';
  document.getElementById('wmAngle').value = -30; document.getElementById('wmAngleN').value = -30;
  document.getElementById('wmSize').value = 15; document.getElementById('wmSizeN').value = 15;
  document.getElementById('footerText').value = '';
  updateZoomDisplay();
  document.getElementById('toggleCutline').classList.add('on');
  document.getElementById('toggleNumber').classList.remove('on');
  document.getElementById('toggleBorder').classList.remove('on');
  document.getElementById('toggleTrimWhite').classList.remove('on');
  document.getElementById('toggleWatermark').classList.remove('on');
  document.getElementById('togglePageNum').classList.remove('on');
  document.getElementById('toggleDate').classList.remove('on');
  document.getElementById('togglePdfText').classList.add('on');
  document.getElementById('toggleFooter').classList.remove('on');
  document.getElementById('toggleCustomFM').classList.remove('on');
  document.getElementById('footerOpts').style.display = 'none';
  document.getElementById('customFMRow').style.display = 'none';
  document.getElementById('footerMarginRow').style.display = 'none';
  document.getElementById('footerMargin').value = 8; document.getElementById('footerMarginN').value = 8;
  document.getElementById('themeMode').value = 'light';
  document.documentElement.classList.remove('dark');
  try { localStorage.removeItem('ticketchan-theme'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-amt-mode'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-pdf-text-enabled'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-settings'); } catch(e) {}
  _printedMap = {};
  _summaryActiveCols = [];
  S._fileAdjMap = {};
  S._notesMap = {};
  S.printedFilter = 'all';
  document.querySelectorAll('.pf-btn').forEach(function(b) {
    b.classList.toggle('pf-active', b.dataset.filter === 'all');
  });
  renderFileList();
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
  if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); savePdf(); }
  if ((e.ctrlKey || e.metaKey) && e.key === 'o') { e.preventDefault(); triggerUpload(); }
  if ((e.ctrlKey || e.metaKey) && (e.key === '=' || e.key === '+')) { e.preventDefault(); changeZoom(5); }
  if ((e.ctrlKey || e.metaKey) && e.key === '-') { e.preventDefault(); changeZoom(-5); }
  if ((e.ctrlKey || e.metaKey) && e.key === '0') { e.preventDefault(); setZoom('fit'); }
  if (e.key === 'Escape') {
    var sm = document.getElementById('summaryModal');
    if (sm && !sm.classList.contains('hidden')) {
      e.preventDefault();
      closeSummaryModal();
    }
  }
});

// Wheel: selected slot + cursor over it → zoom slot; plain → flip page; Ctrl → zoom view
var _wheelFlipTs = 0; // 上次滚轮翻页时间，节流防触控板惯性连翻
document.getElementById('previewWrap').addEventListener('wheel', function(e) {
  if (!e.ctrlKey && S.selectedSlot >= 0) {
    var slotEl = e.target.closest('.invoice-slot');
    if (slotEl && parseInt(slotEl.dataset.slotIdx) === S.selectedSlot) {
      e.preventDefault();
      var f = getSelectedFileObj();
      if (f) {
        var step = 5;
        var curPct = Math.round((f.slotScale || 1) * 100);
        var newPct = e.deltaY > 0 ? curPct - step : curPct + step;
        f.slotScale = Math.max(0.2, Math.min(3.0, newPct / 100));
        updatePreview();
        updateAdjPanel();
        return;
      }
    }
  }
  if (!e.ctrlKey) {
    // Plain wheel: flip pages, unless the zoomed view still has content to scroll
    if (e.deltaY !== 0 && S.totalPages > 1) {
      var wrap = this;
      var canScroll = wrap.scrollHeight > wrap.clientHeight + 1;
      var atTop = wrap.scrollTop <= 0;
      var atBottom = wrap.scrollTop + wrap.clientHeight >= wrap.scrollHeight - 1;
      if (!canScroll || (e.deltaY > 0 && atBottom) || (e.deltaY < 0 && atTop)) {
        e.preventDefault();
        var now = Date.now();
        if (now - _wheelFlipTs > 150) {
          _wheelFlipTs = now;
          if (e.deltaY > 0) nextPage(); else prevPage();
        }
      }
    }
    return;
  }
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

// Double-click: on selected slot → reset per-slot adj (size+position); elsewhere → reset preview zoom
document.getElementById('previewWrap').addEventListener('dblclick', function(e) {
  if (S.selectedSlot >= 0) {
    var slotEl = e.target.closest('.invoice-slot');
    if (slotEl && parseInt(slotEl.dataset.slotIdx) === S.selectedSlot) {
      resetSlotAdj();
      return;
    }
  }
  if (S.viewZoom !== 0) { setZoom('fit'); }
});

// Global drag & drop (browser fallback)
document.body.addEventListener('dragover', function(e) { e.preventDefault(); });

window.addEventListener('resize', function() { if (S.files.length) updatePreview(); });

// beforeunload safety net — stop all work if the window is being destroyed
window.addEventListener('beforeunload', function() {
  _loadingBatchActive = false;
});


// =====================================================
// Initialization — restore saved preferences
// =====================================================
(function() {
  try {
    var saved = localStorage.getItem('ticketchan-theme');
    if (saved === 'dark') {
      document.getElementById('themeMode').value = 'dark';
      document.documentElement.classList.add('dark');
    }
  } catch(e) {}
})();

document.getElementById('orientation').value = 'landscape';

(function() {
  try {
    var m = localStorage.getItem('ticketchan-amt-mode');
    if (m && (m === 'tax' || m === 'notax' || m === 'both')) {
      S.amtMode = m;
      document.getElementById('amtMode').value = m;
    }
  } catch(e) {}
})();

// Restore PDF text extraction setting
(function() {
  try {
    var v = localStorage.getItem('ticketchan-pdf-text-enabled');
    var btn = document.getElementById('togglePdfText');
    if (v === '0') {
      S.feat.pdfTextEnabled = false;
      if (btn) btn.classList.remove('on');
    } else {
      S.feat.pdfTextEnabled = true;
      if (btn) btn.classList.add('on');
    }
  } catch(e) {}
})();

// Restore all layout & feature settings
loadSettings();

// =====================================================
// App initialization
// =====================================================
(function() {
  function showApp() {
    APP_VERSION = '3.3.0';
    var el = document.getElementById('stVersion');
    if (el) el.textContent = 'v' + APP_VERSION;
    console.log('发票酱 ' + APP_VERSION);
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function() { showApp(); bindFooterTextEvent(); setupInputWheelSupport(); initListDrag(); });
  } else {
    showApp(); bindFooterTextEvent(); setupInputWheelSupport(); initListDrag();
  }
})();

// =====================================================
// 发票汇总表 — 可编辑预览 + CSV 导出
// =====================================================

var SUMMARY_FIELDS = [
  { key: 'seq',       label: '序号',     type: 'seq',     default: true, editable: false },
  { key: 'invoiceNo', label: '发票号码',  type: 'text',    default: true, editable: true },
  { key: 'invoiceDate',label: '开票日期', type: 'text',    default: true, editable: true },
  { key: 'invoiceType',label:'发票类型',  type: 'text',    default: false, editable: false },
  { key: 'sellerName',label:'销售方名称', type: 'text',    default: true, editable: true },
  { key: 'sellerCreditCode',label:'销售方税号', type:'text',default: false, editable: true },
  { key: 'buyerName', label: '购买方名称',type: 'text',    default: false, editable: true },
  { key: 'buyerCreditCode',label:'购买方税号',type:'text', default: false, editable: true },
  { key: 'amountTax', label: '含税金额',  type: 'amount',  default: true, editable: true },
  { key: 'amountNoTax',label:'不含税金额',type: 'amount',  default: false, editable: true },
  { key: 'taxAmount', label: '税额',      type: 'amount',  default: false, editable: true },
  { key: 'name',      label: '文件名',    type: 'text',    default: false, editable: true },
  { key: 'copies',    label: '份数',      type: 'copies',  default: false, editable: true },
  { key: 'note',      label: '备注',      type: 'text',    default: true, editable: true }
];

var _summaryActiveCols = []; // keys of currently visible columns
var _summaryOriginalData = []; // snapshot of original values when modal opens

function openSummaryModal() {
  var files = getCheckedFiles();
  if (!files.length) { toast('没有发票数据'); return; }

  // Snapshot original values for edited-cell highlighting
  _summaryOriginalData = files.map(function(f) {
    var snap = {};
    SUMMARY_FIELDS.forEach(function(field) {
      if (field.editable) snap[field.key] = getSummaryCellValue(f, field, 0);
    });
    return snap;
  });

  // Use persisted column selection (restored by loadSettings), or fall back to defaults
  if (!_summaryActiveCols || _summaryActiveCols.length === 0) {
    _summaryActiveCols = [];
    SUMMARY_FIELDS.forEach(function(f) { if (f.default) _summaryActiveCols.push(f.key); });
  }

  renderSummaryColumns();
  renderSummaryTable();

  document.getElementById('summaryModal').classList.remove('hidden');
}

function closeSummaryModal() {
  // Persist column selection via unified settings
  saveSettings();
  document.getElementById('summaryModal').classList.add('hidden');
}

// Render the column checkbox bar
function renderSummaryColumns() {
  var html = '';
  SUMMARY_FIELDS.forEach(function(f) {
    if (f.key === 'seq') return; // seq always shown, no toggle
    var checked = _summaryActiveCols.indexOf(f.key) >= 0 ? ' checked' : '';
    html += '<label class="summary-col-label"><input type="checkbox" data-key="' + f.key + '" ' + checked + ' onchange="onSummaryColToggle(this)">' + f.label + '</label>';
  });
  html += '<span class="summary-col-actions"><a onclick="summarySelectAll()">全选</a><a onclick="summaryDeselectAll()">取消全选</a></span>';
  document.getElementById('summaryColumns').innerHTML = html;
}

function onSummaryColToggle(cb) {
  var key = cb.dataset.key;
  var idx = _summaryActiveCols.indexOf(key);
  if (cb.checked && idx < 0) _summaryActiveCols.push(key);
  if (!cb.checked && idx >= 0) _summaryActiveCols.splice(idx, 1);
  renderSummaryTable();
}

function summarySelectAll() {
  _summaryActiveCols = [];
  SUMMARY_FIELDS.forEach(function(f) { if (f.key !== 'seq') _summaryActiveCols.push(f.key); });
  renderSummaryColumns();
  renderSummaryTable();
}

function summaryDeselectAll() {
  _summaryActiveCols = ['seq', 'invoiceNo'];
  renderSummaryColumns();
  renderSummaryTable();
}

// Get display value for a field on a fileObj
function getSummaryCellValue(fileObj, field, idx) {
  switch (field.key) {
    case 'seq': return String(idx + 1);
    case 'invoiceType':
      if (fileObj.invoiceType) return fileObj.invoiceType;
      if (fileObj._isToll) return '通行费发票';
      if (fileObj._isTicket) return fileObj.sellerName || '车票'; // sellerName holds ticket label
      return '增值税发票';
    case 'amountTax': return fileObj.amountTax > 0 ? fileObj.amountTax.toFixed(2) : '';
    case 'amountNoTax': return fileObj.amountNoTax > 0 ? fileObj.amountNoTax.toFixed(2) : '';
    case 'taxAmount': return fileObj.taxAmount > 0 ? fileObj.taxAmount.toFixed(2) : '';
    case 'copies': return String(fileObj.copies || 1);
    default: return String(fileObj[field.key] || '');
  }
}

// Sync edited value back to fileObj
function setSummaryCellValue(fileObj, field, value) {
  switch (field.key) {
    case 'amountTax': fileObj.amountTax = parseFloat(value) || 0; break;
    case 'amountNoTax': fileObj.amountNoTax = parseFloat(value) || 0; break;
    case 'taxAmount': fileObj.taxAmount = parseFloat(value) || 0; break;
    case 'copies': fileObj.copies = Math.max(1, parseInt(value) || 1); break;
    case 'invoiceType': break; // doesn't sync back (derived field)
    default: fileObj[field.key] = value; break;
  }
}

// Enter: next row same column / Shift+Enter: previous row same column
function onSummaryKeyNav(e, input) {
  if (e.key !== 'Enter') return;
  var shift = e.shiftKey;
  var idx = parseInt(input.dataset.idx);
  var key = input.dataset.key;
  var files = getCheckedFiles();
  if (shift ? idx <= 0 : idx >= files.length - 1) return;
  e.preventDefault();
  input.blur(); // triggers onchange → renderSummaryTable (sync) if value changed
  var target = document.querySelector('#summaryTable input[data-idx="' + (idx + (shift ? -1 : 1)) + '"][data-key="' + key + '"]');
  if (target) { target.focus(); target.select(); }
}

// Render the data table based on current column selection
function renderSummaryTable() {
  var files = getCheckedFiles();
  var visibleFields = SUMMARY_FIELDS.filter(function(f) { return _summaryActiveCols.indexOf(f.key) >= 0; });
  if (visibleFields.length === 0) { _summaryActiveCols = ['seq', 'invoiceNo', 'amountTax']; visibleFields = SUMMARY_FIELDS.filter(function(f) { return _summaryActiveCols.indexOf(f.key) >= 0; }); }

  // Table header
  var html = '<thead><tr>';
  visibleFields.forEach(function(f) {
    var cls = '';
    if (f.key === 'seq') cls = 'col-seq';
    else if (f.type === 'amount' || f.type === 'copies') cls = 'col-' + (f.type === 'amount' ? 'amount' : 'copies');
    else if (f.type === 'text') cls = 'col-text';
    html += '<th class="' + cls + '">' + f.label + '</th>';
  });
  html += '</tr></thead><tbody>';

  var totalAmountTax = 0, totalAmountNoTax = 0, totalTaxAmount = 0;
  files.forEach(function(fileObj, idx) {
    html += '<tr>';
    visibleFields.forEach(function(f) {
      var val = getSummaryCellValue(fileObj, f, idx);
      var cls = '';
      if (f.key === 'seq') cls = 'col-seq';
      else if (f.type === 'amount') cls = 'col-amount';
      else if (f.key === 'copies') cls = 'col-copies';
      else if (f.type === 'text') cls = 'col-text';

      if (!f.editable) {
        html += '<td class="' + cls + ' summary-cell-static" style="padding:6px 10px">' + escHtml(val) + '</td>';
      } else {
        var inputCls = 'summary-cell-input' + (f.type === 'amount' || f.key === 'copies' ? ' number' : '');
        var isEdited = _summaryOriginalData[idx] && _summaryOriginalData[idx][f.key] !== undefined && _summaryOriginalData[idx][f.key] !== val;
        if (isEdited) inputCls += ' edited';
        html += '<td class="' + cls + '"><input class="' + inputCls + '" value="' + escHtml(val) + '" data-idx="' + idx + '" data-key="' + f.key + '" onchange="onSummaryCellEdit(this)" onfocus="this.select()" onkeydown="onSummaryKeyNav(event, this)"></td>';
      }

      if (f.key === 'amountTax' && fileObj.amountTax > 0) totalAmountTax += fileObj.amountTax;
      if (f.key === 'amountNoTax' && fileObj.amountNoTax > 0) totalAmountNoTax += fileObj.amountNoTax;
      if (f.key === 'taxAmount' && fileObj.taxAmount > 0) totalTaxAmount += fileObj.taxAmount;
    });
    html += '</tr>';
  });

  // Total row
  html += '<tr class="summary-total-row">';
  visibleFields.forEach(function(f, ci) {
    if (f.key === 'amountTax') {
      html += '<td class="col-amount"><span class="summary-total-cell">¥' + totalAmountTax.toFixed(2) + '</span></td>';
    } else if (f.key === 'amountNoTax') {
      html += '<td class="col-amount"><span class="summary-total-cell">¥' + totalAmountNoTax.toFixed(2) + '</span></td>';
    } else if (f.key === 'taxAmount') {
      html += '<td class="col-amount"><span class="summary-total-cell">¥' + totalTaxAmount.toFixed(2) + '</span></td>';
    } else if (ci === 0) {
      html += '<td class="col-seq summary-total-cell" style="padding:8px 10px">合计</td>';
    } else {
      html += '<td class="summary-total-cell" style="padding:8px 10px"></td>';
    }
  });
  html += '</tr>';

  html += '</tbody>';
  document.getElementById('summaryTable').innerHTML = html;

  // Update total below table
  var totalEl = document.getElementById('summaryTotal');
  totalEl.textContent = '共 ' + files.length + ' 张发票';
}

// Handle cell edit — sync back to fileObj + refresh all UI
function onSummaryCellEdit(input) {
  var idx = parseInt(input.dataset.idx);
  var key = input.dataset.key;
  var newVal = input.value;

  var files = getCheckedFiles();
  if (idx < 0 || idx >= files.length) return;

  var field = null;
  SUMMARY_FIELDS.forEach(function(f) { if (f.key === key) field = f; });
  if (!field) return;

  setSummaryCellValue(files[idx], field, newVal);

  // Rebuild table to sync all cells (including total row)
  renderSummaryTable();

  // Sync file list badges + bottom amount summary
  renderFileList();

  // Refresh preview in case amounts are overlaid
  updatePreview();


}

// Export to CSV (UTF-8 BOM for Excel compatibility)
async function exportSummaryCsv() {
  var files = getCheckedFiles();
  if (!files.length) { toast('没有发票数据可导出'); return; }

  var visibleFields = SUMMARY_FIELDS.filter(function(f) { return _summaryActiveCols.indexOf(f.key) >= 0; });
  if (visibleFields.length === 0) return;

  // Build CSV content
  var rows = [];
  // Header
  rows.push(visibleFields.map(function(f) { return csvEscape(f.label); }).join(','));
  // Data rows
  files.forEach(function(fileObj, idx) {
    rows.push(visibleFields.map(function(f) {
      return csvEscape(getSummaryCellValue(fileObj, f, idx));
    }).join(','));
  });
  // Total row
  var totalAmountTax = files.reduce(function(s, f) { return s + (f.amountTax || 0); }, 0);
  var totalAmountNoTax = files.reduce(function(s, f) { return s + (f.amountNoTax || 0); }, 0);
  var totalTaxAmount = files.reduce(function(s, f) { return s + (f.taxAmount || 0); }, 0);
  rows.push(visibleFields.map(function(f, ci) {
    if (f.key === 'amountTax') return csvEscape(totalAmountTax.toFixed(2));
    if (f.key === 'amountNoTax') return csvEscape(totalAmountNoTax.toFixed(2));
    if (f.key === 'taxAmount') return csvEscape(totalTaxAmount.toFixed(2));
    if (ci === 0) return csvEscape('合计');
    return '';
  }).join(','));

  var csvContent = '\uFEFF' + rows.join('\r\n');

  // Download via Blob
  var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a'); a.href = url; a.download = '发票汇总表.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
  closeSummaryModal();
  toast('汇总表已导出');
}

function csvEscape(val) {
  var s = String(val || '');
  if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function bindFooterTextEvent() {
  var el = document.getElementById('footerText');
  if (el) el.addEventListener('input', function() { updatePreview(); });
}
