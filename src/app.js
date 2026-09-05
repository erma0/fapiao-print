// =====================================================
// 发票酱 — 主入口
// v1.10.5 — 预览加速 + 批量加载 + 智能缓存 + IPC 异步化
// =====================================================

// Detect Tauri — use var to avoid conflict with Tauri's injected scripts
var isTauri = window.__TAURI_INTERNALS__ !== undefined;
var invoke  = isTauri ? window.__TAURI_INTERNALS__.invoke : null;
var hasOcr  = false; // Set to true at startup if OCR feature is available
var APP_VERSION = ''; // Filled at startup from Rust get_app_version()
var _winrtPdfAvailable = true; // Set to false at startup if WinRT PDF component is missing

// =====================================================
// Constants
// =====================================================
var PAPER = { A4:{w:210,h:297}, A5:{w:148,h:210}, B5:{w:176,h:250}, letter:{w:216,h:279}, legal:{w:216,h:356} };
var MM2PX = 96 / 25.4;
var PDF_RENDER_DPI = 300;  // Print/save DPI — Must match Rust RENDER_DPI
var PDF_PREVIEW_DPI = 150;  // Preview DPI — faster loading, lower resolution
var MIN_RENDER_PX = 3508;  // A4 long side at 300 DPI — minimum rendered pixels
var WHITE_THRESHOLD = 245; // Pixel value threshold for white-edge trimming

// =====================================================
// State
// =====================================================
// 默认快捷布局（顶部工具栏按钮，可插拔）：原 6 个常用 + 4×5 一页 20 格总览（#11）
var DEFAULT_QUICK_LAYOUTS = [
  { cols: 1, rows: 1 },   // 1×1
  { cols: 2, rows: 1 },   // 1×2
  { cols: 3, rows: 2 },   // 2×3
  { cols: 1, rows: 2 },   // 2×1
  { cols: 2, rows: 2 },   // 2×2
  { cols: 3, rows: 3 },   // 3×3
  { cols: 5, rows: 4 }    // 4×5
];
function normalizeQuickLayoutValue(value) {
  return Math.max(1, Math.min(10, parseInt(value) || 1));
}
function normalizeQuickLayoutMax(value) {
  return Math.max(0, Math.min(20, parseInt(value) || 0));
}
function cloneQuickLayouts(items) {
  if (!Array.isArray(items)) return [];
  return items.map(function(q) {
    return {
      cols: normalizeQuickLayoutValue(q && q.cols),
      rows: normalizeQuickLayoutValue(q && q.rows)
    };
  });
}
function defaultQuickLayouts() {
  return cloneQuickLayouts(DEFAULT_QUICK_LAYOUTS);
}
var S = {
  files: [],
  currentPage: 0,
  totalPages: 0,
  viewZoom: 0,
  layout: { cols: 1, rows: 1, orient: 'landscape' },
  // 顶部工具栏快捷排版按钮（可插拔）：内置 7 个，可增删改、排序
  quickLayouts: defaultQuickLayouts(),
  quickLayoutMax: 0,  // 顶部显示数量，0=全部
  editIdx: -1,
  selectedSlot: -1,  // Index of currently selected slot in preview (for per-slot adjustment)
  amtMode: 'tax',
  printedFilter: 'all',
  fileFilter: 'all',
  fileView: 'list',
  ocrPrecision: 'standard',
  feat: {
    cutline: true, number: false, border: false, trimWhite: false,
    watermark: false, collate: true, duplex: false, pageNum: false,
    printDate: false, footer: false,
    autoOpenPdf: true,
    ocrEnabled: false,
    pdfTextEnabled: true,
    customFM: false,
    fileListMemory: false,
    autoDedup: false
  }
};

// Track newly added file IDs for entrance animation
var _newFileIds = {};

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
    _ocrText: opts._ocrText || '',
    _isTicket: opts._isTicket || false,
    _isToll: opts._isToll || false,
    _loading: opts._loading || false,
    _placeholder: opts._placeholder || false,   // 版面空白占位（只占槽位，不打印不统计）
    _ocrPending: false,
    _xmlInvoice: opts._xmlInvoice || false,
    // Disk path for the original file (when available).
    // Used by Rust to read bytes directly, skipping base64 encode/decode.
    _filePath: opts.filePath || '',
    // PDF source info for ocr_pdf_page command (zero IPC round-trip OCR).
    // Set when this fileObj represents a PDF page rendered via render_pdf_pages.
    _pdfPath: opts.pdfPath || '',
    _pdfPageIdx: opts.pdfPageIdx != null ? opts.pdfPageIdx : -1,
    // Per-slot adjustment: scale & position within the layout slot
    slotScale: opts.slotScale || 1,        // 1.0 = default (contain-fit size)
    slotOffsetX: opts.slotOffsetX || 0,    // X offset in mm (0 = centered)
    slotOffsetY: opts.slotOffsetY || 0,    // Y offset in mm (0 = centered)
    _printed: false                        // True after successful print
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
// v2.1.1: Toast 防抖,避免批量加载时频繁 innerHTML 写入
var _toastLoadMsg = '';
var _toastLoadTimer = null;
var _toastLoadDirty = false;
function toast(msg, dur) { dur = dur || 2500; var e = document.getElementById('toast'); e.textContent = msg; e.classList.add('show'); clearTimeout(toastT); if (dur > 0) toastT = setTimeout(function() { e.classList.remove('show'); }, dur); else clearTimeout(toastT); }
function toastHtml(msg, dur) { dur = dur || 2500; var e = document.getElementById('toast'); e.innerHTML = msg; e.classList.add('show'); clearTimeout(toastT); if (dur > 0) toastT = setTimeout(function() { e.classList.remove('show'); }, dur); else clearTimeout(toastT); }
function toastLoading(msg) {
  _ocrToastActive = true;
  _toastLoadMsg = msg;
  _toastLoadDirty = true;
  // 防抖: 100ms 内多次调用只更新一次 DOM
  if (_toastLoadTimer === null) {
    _toastLoadTimer = setTimeout(function() {
      _toastLoadTimer = null;
      if (_toastLoadDirty) {
        _toastLoadDirty = false;
        toastHtml('<span class="toast-spinner"></span>' + _toastLoadMsg, 0);
      }
    }, 100);
  }
}
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

// Open URL in external browser (reuse verifyInvoice pattern)
function openExternal(url) {
  if (isTauri && invoke) {
    invoke('open_url', { url: url }).catch(function(e) {
      console.warn('[openExternal] Tauri open failed:', e);
      toast('打开浏览器失败，请手动访问: ' + url);
    });
  } else {
    window.open(url, '_blank');
  }
}

function showSumatraPdfMissing() {
  var existing = document.getElementById('sumatraPdfModal');
  if (existing) { existing.classList.remove('hidden'); return; }
  var div = document.createElement('div');
  div.id = 'sumatraPdfModal';
  div.className = 'modal-bg';
  div.innerHTML = '<div class="modal" onclick="event.stopPropagation()">' +
    '<div class="modal-title">未检测到 SumatraPDF</div>' +
    '<div class="modal-body" style="padding:8px 0;font-size:13px;line-height:1.6;color:var(--text-secondary)">' +
    'SumatraPDF 是一款免费轻量的 PDF 阅读器，支持静默打印。<br>' +
    '<span style="font-size:12px;color:var(--text-muted)">手动下载后请将 exe 重命名为 SumatraPDF.exe，放到程序目录下的 tools 文件夹</span></div>' +
    '<div class="modal-actions" style="flex-direction:column;gap:8px">' +
    '<button class="btn btn-primary" style="width:100%" onclick="downloadSumatraPdf()">自动下载</button>' +
    '<button class="btn btn-sm" style="width:100%" onclick="openExternal(\'https://www.sumatrapdfreader.org/download-free-pdf-viewer\')">手动下载（官网）</button>' +
    '<button class="btn btn-sm" style="width:100%" onclick="switchToPdfMode()">切换到「PDF阅读器」模式</button>' +
    '</div>' +
    '<div class="modal-actions" style="margin-top:8px;justify-content:flex-end">' +
    '<button class="btn btn-sm" onclick="document.getElementById(\'sumatraPdfModal\').classList.add(\'hidden\')">取消</button>' +
    '</div></div></div>';
  div.onclick = function() { div.classList.add('hidden'); };
  document.body.appendChild(div);
}

async function downloadSumatraPdf() {
  if (!isTauri || !invoke) return;
  var modal = document.getElementById('sumatraPdfModal');
  if (modal) {
    var body = modal.querySelector('.modal-body');
    var actions = modal.querySelectorAll('.modal-actions');
    if (body) body.innerHTML = '<div style="text-align:center;padding:16px 0">' +
      '<div class="spinner" style="width:40px;height:40px;border-width:3px;margin:0 auto 12px"></div>' +
      '<div style="font-size:14px;color:var(--text-secondary);margin-bottom:10px">正在下载 SumatraPDF，请稍候...</div>' +
      '<div style="width:100%;height:14px;background:var(--bg-secondary);border-radius:7px;overflow:hidden">' +
        '<div id="sumatraDownloadProgress" style="height:100%;width:0%;background:var(--accent);border-radius:7px;transition:width 0.2s"></div>' +
      '</div>' +
      '<div id="sumatraDownloadPercent" style="font-size:13px;color:var(--text-muted);margin-top:6px">0%</div>' +
    '</div>';
    if (actions[0]) actions[0].innerHTML = '';
    if (actions[1]) actions[1].innerHTML = '<button class="btn btn-sm" onclick="cancelSumatraDownload()">取消下载</button>';
    modal.classList.remove('hidden');
  }
  _sumatraDownloadAborted = false;
  var unlistenProgress = null;
  try {
    if (isTauri && window.__TAURI_INTERNALS__) {
      var callbackId = window.__TAURI_INTERNALS__.transformCallback(function(evt) {
        var progress = evt.payload;
        var bar = document.getElementById('sumatraDownloadProgress');
        var percent = document.getElementById('sumatraDownloadPercent');
        if (bar) bar.style.width = Math.min(100, progress.percent).toFixed(0) + '%';
        if (percent) percent.textContent = Math.min(100, progress.percent).toFixed(0) + '%';
      });
      var eventId = await invoke('plugin:event|listen', {
        event: 'sumatra-download-progress',
        target: { kind: 'Any' },
        handler: callbackId
      });
      unlistenProgress = function() {
        try { invoke('plugin:event|unlisten', { event: 'sumatra-download-progress', eventId: eventId }); } catch(e) {}
      };
    }
    var result = await invoke('download_sumatrapdf');
    if (unlistenProgress) unlistenProgress();
    if (_sumatraDownloadAborted) return;
    if (modal) modal.classList.add('hidden');
    if (result.success) {
      toast('\u2705 ' + result.message);
    } else {
      showSumatraDownloadError(result.message);
    }
  } catch(e) {
    if (unlistenProgress) unlistenProgress();
    if (_sumatraDownloadAborted) return;
    if (modal) modal.classList.add('hidden');
    showSumatraDownloadError(String(e));
  }
}

var _sumatraDownloadAborted = false;

function cancelSumatraDownload() {
  _sumatraDownloadAborted = true;
  if (isTauri && invoke) { try { invoke('cancel_download'); } catch(e) {} }
  var modal = document.getElementById('sumatraPdfModal');
  if (modal) modal.classList.add('hidden');
  toast('下载已取消');
}

function showSumatraDownloadError(errMsg) {
  var modal = document.getElementById('sumatraPdfModal');
  if (!modal) { showSumatraPdfMissing(); modal = document.getElementById('sumatraPdfModal'); }
  if (!modal) return;
  var body = modal.querySelector('.modal-body');
  var actions = modal.querySelectorAll('.modal-actions');
  if (body) body.innerHTML = '<div style="padding:12px 16px;background:var(--danger-light);border-radius:8px;border-left:4px solid var(--danger);margin-bottom:4px">' +
    '<div style="font-size:15px;font-weight:600;color:var(--danger);margin-bottom:6px">\u274c 下载失败</div>' +
    '<div style="font-size:12px;line-height:1.5;color:var(--text-secondary);word-break:break-all">' + escHtml(errMsg) + '</div>' +
  '</div>';
  if (actions[0]) actions[0].innerHTML =
    '<button class="btn btn-primary" style="width:100%" onclick="downloadSumatraPdf()">重试下载</button>' +
    '<button class="btn btn-sm" style="width:100%" onclick="openExternal(\'https://www.sumatrapdfreader.org/download-free-pdf-viewer\')">手动下载（官网）</button>' +
    '<button class="btn btn-sm" style="width:100%" onclick="switchToPdfMode()">切换到「PDF阅读器」模式</button>';
  if (actions[1]) actions[1].innerHTML = '<button class="btn btn-sm" onclick="document.getElementById(\'sumatraPdfModal\').classList.add(\'hidden\')">关闭</button>';
  modal.classList.remove('hidden');
}

function switchToPdfMode() {
  document.getElementById('printMode').value = 'pdf';
  try { localStorage.setItem('ticketchan-print-mode', 'pdf'); } catch(e) {}
  var modal1 = document.getElementById('sumatraPdfModal');
  if (modal1) modal1.classList.add('hidden');
  var modal2 = document.getElementById('pdfiumModal');
  if (modal2) modal2.classList.add('hidden');
  toast('已切换到 PDF 阅读器模式');
}

function showPdfiumMissing(reason) {
  var existing = document.getElementById('pdfiumModal');
  if (existing) { existing.classList.remove('hidden'); return; }
  var reasonHtml = reason
    ? '<div style="padding:10px 14px;background:var(--warning-light, #fff8e1);border-radius:8px;border-left:4px solid var(--warning, #f59e0b);margin-bottom:8px;font-size:13px;line-height:1.5;color:var(--text-secondary)">' + escHtml(reason) + '</div>'
    : '';
  var div = document.createElement('div');
  div.id = 'pdfiumModal';
  div.className = 'modal-bg';
  div.innerHTML = '<div class="modal" onclick="event.stopPropagation()">' +
    '<div class="modal-title">需要下载 PDF 渲染组件</div>' +
    '<div class="modal-body" style="padding:8px 0;font-size:13px;line-height:1.6;color:var(--text-secondary)">' +
    reasonHtml +
    'PDFium 是 Chromium 内核的 PDF 渲染引擎，用于加载和预览 PDF 发票。<br>' +
    '<span style="font-size:12px;color:var(--text-muted)">下载后自动生效，无需重启。也可手动将 pdfium.dll 放到程序目录下的 tools 文件夹</span></div>' +
    '<div class="modal-actions" style="flex-direction:column;gap:8px">' +
    '<button class="btn btn-primary" style="width:100%" onclick="downloadPdfiumDll()">自动下载（约 7MB）</button>' +
    '<button class="btn btn-sm" style="width:100%" onclick="openExternal(\'https://github.com/bblanchon/pdfium-binaries/releases\')">手动下载（GitHub Releases）</button>' +
    '</div>' +
    '<div class="modal-actions" style="margin-top:8px;justify-content:flex-end">' +
    '<button class="btn btn-sm" onclick="document.getElementById(\'pdfiumModal\').classList.add(\'hidden\')">取消</button>' +
    '</div></div></div>';
  div.onclick = function() { div.classList.add('hidden'); };
  document.body.appendChild(div);
}

var _pdfiumDownloadAborted = false;

async function downloadPdfiumDll() {
  if (!isTauri || !invoke) return;
  var modal = document.getElementById('pdfiumModal');
  if (modal) {
    var body = modal.querySelector('.modal-body');
    var actions = modal.querySelectorAll('.modal-actions');
    if (body) body.innerHTML = '<div style="text-align:center;padding:16px 0">' +
      '<div class="spinner" style="width:40px;height:40px;border-width:3px;margin:0 auto 12px"></div>' +
      '<div style="font-size:14px;color:var(--text-secondary);margin-bottom:10px">正在下载 pdfium.dll，请稍候...</div>' +
      '<div style="width:100%;height:14px;background:var(--bg-secondary);border-radius:7px;overflow:hidden">' +
        '<div id="pdfiumDownloadProgress" style="height:100%;width:0%;background:var(--accent);border-radius:7px;transition:width 0.2s"></div>' +
      '</div>' +
      '<div id="pdfiumDownloadPercent" style="font-size:13px;color:var(--text-muted);margin-top:6px">0%</div>' +
    '</div>';
    if (actions[0]) actions[0].innerHTML = '';
    if (actions[1]) actions[1].innerHTML = '<button class="btn btn-sm" onclick="cancelPdfiumDownload()">取消下载</button>';
    modal.classList.remove('hidden');
  }
  _pdfiumDownloadAborted = false;
  var unlistenProgress = null;
  try {
    if (isTauri && window.__TAURI_INTERNALS__) {
      var callbackId = window.__TAURI_INTERNALS__.transformCallback(function(evt) {
        var progress = evt.payload;
        var bar = document.getElementById('pdfiumDownloadProgress');
        var percent = document.getElementById('pdfiumDownloadPercent');
        if (bar) bar.style.width = Math.min(100, progress.percent).toFixed(0) + '%';
        if (percent) percent.textContent = Math.min(100, progress.percent).toFixed(0) + '%';
      });
      var eventId = await invoke('plugin:event|listen', {
        event: 'pdfium-download-progress',
        target: { kind: 'Any' },
        handler: callbackId
      });
      unlistenProgress = function() {
        try { invoke('plugin:event|unlisten', { event: 'pdfium-download-progress', eventId: eventId }); } catch(e) {}
      };
    }
    var result = await invoke('download_pdfium_dll');
    if (unlistenProgress) unlistenProgress();
    if (_pdfiumDownloadAborted) return;
    if (modal) modal.classList.add('hidden');
    if (result.success) {
      toast('\u2705 ' + result.message + '，请重新添加 PDF 文件');
    } else {
      showPdfiumDownloadError(result.message);
    }
  } catch(e) {
    if (unlistenProgress) unlistenProgress();
    if (_pdfiumDownloadAborted) return;
    if (modal) modal.classList.add('hidden');
    showPdfiumDownloadError(String(e));
  }
}

function cancelPdfiumDownload() {
  _pdfiumDownloadAborted = true;
  if (isTauri && invoke) { try { invoke('cancel_download'); } catch(e) {} }
  var modal = document.getElementById('pdfiumModal');
  if (modal) modal.classList.add('hidden');
  toast('下载已取消');
}

function showPdfiumDownloadError(errMsg) {
  var modal = document.getElementById('pdfiumModal');
  if (!modal) { showPdfiumMissing('下载失败，请重试。'); modal = document.getElementById('pdfiumModal'); }
  if (!modal) return;
  var body = modal.querySelector('.modal-body');
  var actions = modal.querySelectorAll('.modal-actions');
  if (body) body.innerHTML = '<div style="padding:12px 16px;background:var(--danger-light);border-radius:8px;border-left:4px solid var(--danger);margin-bottom:4px">' +
    '<div style="font-size:15px;font-weight:600;color:var(--danger);margin-bottom:6px">\u274c 下载失败</div>' +
    '<div style="font-size:12px;line-height:1.5;color:var(--text-secondary);word-break:break-all">' + escHtml(errMsg) + '</div>' +
  '</div>';
  if (actions[0]) actions[0].innerHTML =
    '<button class="btn btn-primary" style="width:100%" onclick="downloadPdfiumDll()">重试下载</button>' +
    '<button class="btn btn-sm" style="width:100%" onclick="openExternal(\'https://github.com/bblanchon/pdfium-binaries/releases\')">手动下载（GitHub Releases）</button>' +
    '<button class="btn btn-sm" style="width:100%" onclick="switchToPdfMode()">切换到「PDF阅读器」模式</button>';
  if (actions[1]) actions[1].innerHTML = '<button class="btn btn-sm" onclick="document.getElementById(\'pdfiumModal\').classList.add(\'hidden\')">关闭</button>';
  modal.classList.remove('hidden');
}

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
        resolve(canvas.toDataURL('image/jpeg', 0.92));
      };
      img.onerror = function() { resolve(dataUrl); };
      img.src = dataUrl;
    } catch(e) { resolve(dataUrl); }
  });
}

function ocrMaxDim() {
  var p = S.ocrPrecision || 'standard';
  if (p === 'fast') return 1280;
  if (p === 'precise') return 2800;
  return 1920;
}

// =====================================================
// FILE UPLOAD — via Tauri dialog plugin
// =====================================================
async function restoreFiles(paths) {
  _isRestoringFiles = true;
  var checks = await Promise.all(paths.map(function(p) {
    return invoke('check_path_exists', { path: p })
      .then(function(info) { return { path: p, valid: !!(info && info.exists && info.isFile) }; })
      .catch(function() { return { path: p, valid: false }; });
  }));
  var valid = checks.filter(function(c) { return c.valid; }).map(function(c) { return c.path; });
  var skipped = paths.length - valid.length;
  if (!valid.length) {
    _isRestoringFiles = false;
    renderFileList();
    if (skipped > 0) toast('上次的 ' + skipped + ' 个文件已不存在，已自动跳过');
    return;
  }
  try {
    if (valid.length <= 3) {
      toastLoading('恢复 ' + valid.length + ' 个文件...');
      var fileDataList = await invoke('open_invoice_files', { paths: valid });
      if (fileDataList && fileDataList.length > 0) {
        await processFileDataList(fileDataList);
      }
    } else {
      await processFilesIncremental(valid);
    }
  } catch(e) {
    toast('恢复发票列表失败: ' + String(e));
  }
  // Delay to allow async applyPdfTextToResults callbacks to finish
  // before clearing the OCR-skip flag (they fire after processFileDataList returns)
  setTimeout(function() { _isRestoringFiles = false; }, 3000);
  if (skipped > 0) toast('上次的 ' + skipped + ' 个文件已不存在，已自动跳过');
}

var _insertSlotIdx = -1;   // 点击版面空白槽位加号上传时的目标槽位，-1 表示普通追加
var _slotUploadActive = false; // Prevent concurrent uploads from sharing insertion state

async function openFileDialog() {
  if (isTauri && invoke) {
    try {
      var result = await invoke('plugin:dialog|open', {
        options: {
          multiple: true,
          title: '选择发票文件',
          filters: [{ name: '发票文件', extensions: ['pdf', 'jpg', 'jpeg', 'png', 'bmp', 'webp', 'tiff', 'tif', 'ofd', 'xml'] }]
        }
      });
      if (!result) return [];
      return typeof result === 'string' ? [result] : (Array.isArray(result) ? result : []);
    } catch (err) {
      console.error('Dialog error:', err);
      hideToast();
      toast('打开文件对话框失败: ' + String(err));
      return [];
    }
  }
  document.getElementById('fileInput').click();
  return null;
}

async function addPaths(paths) {
  if (paths.length <= 3) {
    toastLoading('读取 ' + paths.length + ' 个文件...');
    var fileDataList = await invoke('open_invoice_files', { paths: paths });
    if (fileDataList && fileDataList.length > 0) {
      await processFileDataList(fileDataList);
    } else {
      toast('无法读取所选文件');
    }
  } else {
    await processFilesIncremental(paths);
  }
}

async function triggerUpload() {
  if (_slotUploadActive || _loadingBatchActive) {
    toast('当前仍在加载发票，请稍候再添加');
    return;
  }
  _slotUploadActive = true;
  _insertSlotIdx = -1;
  var browserPending = false;
  try {
    var paths = await openFileDialog();
    if (paths === null) {
      browserPending = true;
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
      return;
    }
    if (paths.length === 0) return;
    await addPaths(paths);
  } catch (err) {
    console.error('Add files error:', err);
    toast('文件加载失败: ' + String(err));
  } finally {
    if (!browserPending) {
      _insertSlotIdx = -1;
      _slotUploadActive = false;
    }
  }
}

// 点击版面空白槽位的加号：上传发票并插入到该槽位对应位置
async function addFileToSlot(slotIdx) {
  if (_slotUploadActive || _loadingBatchActive) {
    toast('当前仍在加载发票，请稍候再补传');
    return;
  }
  _slotUploadActive = true;
  _insertSlotIdx = slotIdx;
  var browserPending = false;
  try {
    var paths = await openFileDialog();
    // Browser fallback completes through handleFileInput(), which consumes the
    // insertion state after the native file picker fires its change event.
    if (paths === null) {
      browserPending = true;
      // A native input does not fire `change` when the picker is cancelled.
      // Release the slot lock on the focus return path in that case.
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
      return;
    }
    if (paths.length === 0) return;
    await addPaths(paths);
  } catch (err) {
    console.error('Add files error:', err);
    toast('文件加载失败: ' + String(err));
  } finally {
    if (!browserPending) {
      _insertSlotIdx = -1;
      _slotUploadActive = false;
    }
  }
}

// 槽位上传的插入准备：返回 { insertAt, blankCount, replaceIdx }
// - 目标槽位已有空白占位 → replaceIdx = 占位在 S.files 中的索引（升级替换）
// - 目标槽位在当前 active 文件之后 → 中间空槽位用占位补齐（insertAt + blankCount）
// 实现"点击哪个格子，新文件就固定出现在哪个格子"的精准定位
function prepareSlotInsertion() {
  var files = getActiveFiles();
  var perPage = getPerPage(getSettings());
  var pos = S.currentPage * perPage + _insertSlotIdx;
  var reverse = document.getElementById('pageOrder').value === 'reverse';
  if (pos < 0) return { insertAt: S.files.length, blankCount: 0, replaceIdx: -1 };
  var target = pos < files.length ? files[pos] : null;
  if (target) {
    if (target._placeholder) {
      // 目标槽位是空白占位：升级替换（保留其他留白）
      return { insertAt: -1, blankCount: 0, replaceIdx: S.files.indexOf(target), reverse: reverse };
    }
    // 目标槽位已有文件（防御，正常不会点到这里）：按显示顺序插入
    var targetIdx = S.files.indexOf(target);
    return { insertAt: reverse ? targetIdx : targetIdx + 1, blankCount: 0, replaceIdx: -1, reverse: reverse };
  }
  // 目标槽位在当前 active 文件之后：补足中间空白
  var blankCount = pos - files.length;
  // Reverse mode renders the tail of S.files first, so append-to-display means
  // inserting at the head of the underlying array.
  var insertAt = reverse ? 0 : S.files.length;
  if (files.length > 0) {
    var last = S.files.indexOf(files[files.length - 1]);
    if (last >= 0 && !reverse) insertAt = last + 1;
  }
  return { insertAt: insertAt, blankCount: blankCount, replaceIdx: -1, reverse: reverse };
}

// 在指定位置插入空白占位对象，返回占位数量（供加载函数同步偏移 placeholder 插入点）
function insertBlankSlots(insertAt, count) {
  for (var b = 0; b < count; b++) {
    var blank = createFileObj({ name: '空白占位', _placeholder: true });
    blank.checked = false;
    S.files.splice(insertAt + b, 0, blank);
  }
}

var _lastInsertedId = null;   // 槽位上传后用于定位新文件的 id

// 槽位上传完成后定位到新插入的文件：跳页 + 高亮列表 + 选中槽位
function locateInsertedFile(id) {
  if (!id) return;
  var f = null;
  for (var i = 0; i < S.files.length; i++) {
    if (S.files[i].id === id) { f = S.files[i]; break; }
  }
  if (!f) return;
  var files = getActiveFiles();
  var activeIdx = files.indexOf(f);
  if (activeIdx < 0) return;
  var perPage = getPerPage(getSettings());
  S.currentPage = Math.floor(activeIdx / perPage);
  _activeFileIdx = S.files.indexOf(f);
  S.selectedSlot = activeIdx % perPage;
}

// 左侧文件列表滚动到当前激活项（renderFileList 之后调用）
function scrollActiveFileIntoView() {
  if (_activeFileIdx < 0) return;
  var list = document.getElementById('fileList');
  if (!list) return;
  var el = list.querySelector('.file-item[data-idx="' + _activeFileIdx + '"]');
  if (el) el.scrollIntoView({ block: 'nearest' });
}

async function handleFileInput(fl) {
  try {
    if (!fl || !fl.length) return;
    await processFiles(Array.from(fl));
  } finally {
    _insertSlotIdx = -1;
    _slotUploadActive = false;
    document.getElementById('fileInput').value = '';
  }
}

// Process FileData array from Rust backend — instant placeholders, then load in parallel + render sequentially
async function processFileDataList(fileDataList) {
  var total = fileDataList.length;
  var completed = 0;
  var added = 0;
  var slotInsert = _insertSlotIdx >= 0;
  var firstPlaceholder = null;
  _loadingBatchActive = true;

  // 1. Create placeholder entries immediately for instant visual feedback
  var insertAt = S.files.length;
  var replaceIdx = -1;
  if (_insertSlotIdx >= 0) {
    var prep = prepareSlotInsertion();
    replaceIdx = prep.replaceIdx;
    insertAt = prep.insertAt;
  }
  if (replaceIdx >= 0) {
    // Replace exactly the clicked placeholder. Additional files are inserted
    // in display order, independent of whether placeholders are contiguous.
    var targetPh = S.files[replaceIdx];
    var firstFd = fileDataList[0];
    targetPh._placeholder = false;
    targetPh.checked = true;
    targetPh.name = firstFd.name;
    targetPh.size = firstFd.size;
    targetPh.type = firstFd.ext;
    targetPh._loading = true;
    targetPh._placeholderKey = targetPh.id;
    firstFd._phKey = targetPh._placeholderKey;
    firstPlaceholder = targetPh;
    _newFileIds[targetPh.id] = true;
    for (var ri = 1; ri < fileDataList.length; ri++) {
      var rfd = fileDataList[ri];
      var rph = createFileObj({ name: rfd.name, size: rfd.size, type: rfd.ext, _loading: true });
      rph._placeholderKey = rph.id;
      rfd._phKey = rph._placeholderKey;
      S.files.splice(prep.reverse ? replaceIdx : replaceIdx + ri, 0, rph);
      _newFileIds[rph.id] = true;
    }
  } else {
    if (_insertSlotIdx >= 0) {
      insertBlankSlots(insertAt, prep.blankCount);
      if (!prep.reverse) insertAt += prep.blankCount;
    }
    fileDataList.forEach(function(fd, fi) {
      var ph = createFileObj({
        name: fd.name,
        size: fd.size,
        type: fd.ext,
        _loading: true
      });
      ph._placeholderKey = ph.id;
      fd._phKey = ph._placeholderKey;
      if (fi === 0) firstPlaceholder = ph;
      S.files.splice(prep && prep.reverse ? insertAt : insertAt + fi, 0, ph);
      _newFileIds[ph.id] = true;
    });
  }

  // Render placeholders immediately — user sees skeleton items right away
  renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();

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

  var startTime = Date.now();
  var updateIntervalMs = Math.max(50, Math.min(150, Math.floor(500 / total)));
  var hasNewResults = false;

  var updateInterval = setInterval(function() {
    if (hasNewResults) {
      renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();
      hasNewResults = false;
    }
  }, updateIntervalMs);

  var lastToastUpdate = 0;
  for (var fdIdx = 0; fdIdx < fileDataList.length; fdIdx++) {
    var r = await loadPromises[fdIdx];
    completed++;

    var fd = fileDataList[fdIdx];
    var phIdx = -1;
    for (var i = 0; i < S.files.length; i++) {
      if (S.files[i]._placeholderKey === fd._phKey) { phIdx = i; break; }
    }

    if (phIdx >= 0 && r) {
      var items = Array.isArray(r) ? r : [r];
      items.forEach(function(it) { _newFileIds[it.id] = true; });
      if (slotInsert && !_lastInsertedId && firstPlaceholder && fd._phKey === firstPlaceholder._placeholderKey) _lastInsertedId = items[0].id;
      S.files.splice.apply(S.files, [phIdx, 1].concat(items));
      added += items.length;
    } else if (phIdx >= 0) {
      S.files.splice(phIdx, 1);
    }

    var now = Date.now();
    if (now - lastToastUpdate > 100 || completed >= total) {
      lastToastUpdate = now;
      var ocrRemaining = _ocrQueue.length + _ocrRunning;
      var isLast = (completed >= total);
      if (isLast) {
        if (ocrRemaining > 0 && S.feat.ocrEnabled) {
          var ocrDone2 = _ocrBatchTotal > 0 ? _ocrBatchTotal - ocrRemaining : 0;
          toastLoading('加载完成，识别中 ' + ocrDone2 + '/' + _ocrBatchTotal);
        }
      } else {
        if (ocrRemaining > 0 && S.feat.ocrEnabled) {
          var ocrDone = _ocrBatchTotal > 0 ? _ocrBatchTotal - ocrRemaining : 0;
          toastLoading('加载中 ' + completed + '/' + total + '，识别中 ' + ocrDone + '/' + _ocrBatchTotal);
        } else {
          toastLoading('加载中 ' + completed + '/' + total);
        }
      }
    }

    hasNewResults = true;
    await nextFrame();
  }

  clearInterval(updateInterval);
  if (slotInsert) locateInsertedFile(_lastInsertedId);
  if (S.feat.autoDedup) removeDuplicates(true);
  renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();
  scrollActiveFileIntoView();

  _loadingBatchActive = false;
  _insertSlotIdx = -1;
  _lastInsertedId = null;

  if (_ocrQueue.length === 0 && _ocrRunning === 0) {
    _ocrToastActive = false;
    _ocrBatchTotal = 0;
    _ocrBatchAddedCount = 0;
    var elapsed = Date.now() - startTime;
    var minToastDelay = Math.max(300, 800 - elapsed);
    if (added > 0) {
      var doneMsg = '已加载 ' + added + ' 张发票';
      setTimeout(function() { toast(doneMsg, 2500); }, minToastDelay);
    } else {
      toast('文件加载失败');
    }
  } else {
    _ocrBatchAddedCount = added;
  }
}

// Process an array of File objects (browser fallback) — instant placeholders, then load in parallel + render sequentially
async function processFiles(files) {
  var total = files.length;
  var completed = 0;
  var added = 0;
  var slotInsert = _insertSlotIdx >= 0;
  var firstPlaceholder = null;
  _loadingBatchActive = true;

  // Create placeholder entries immediately
  var insertAt = S.files.length;
  var replaceIdx = -1;
  if (_insertSlotIdx >= 0) {
    var prep = prepareSlotInsertion();
    replaceIdx = prep.replaceIdx;
    insertAt = prep.insertAt;
  }
  if (replaceIdx >= 0) {
    var targetPh = S.files[replaceIdx];
    var firstFile = files[0];
    var firstExt = firstFile.name.split('.').pop().toLowerCase();
    targetPh._placeholder = false;
    targetPh.checked = true;
    targetPh.name = firstFile.name;
    targetPh.size = firstFile.size;
    targetPh.type = firstExt;
    targetPh._loading = true;
    targetPh._placeholderKey = targetPh.id;
    firstFile._phKey = targetPh._placeholderKey;
    firstPlaceholder = targetPh;
    _newFileIds[targetPh.id] = true;
    for (var ri = 1; ri < files.length; ri++) {
      var rfile = files[ri];
      var rext = rfile.name.split('.').pop().toLowerCase();
      var rph = createFileObj({ name: rfile.name, size: rfile.size, type: rext, _loading: true });
      rph._placeholderKey = rph.id;
      rfile._phKey = rph._placeholderKey;
      S.files.splice(prep.reverse ? replaceIdx : replaceIdx + ri, 0, rph);
      _newFileIds[rph.id] = true;
    }
  } else {
    if (_insertSlotIdx >= 0) {
      insertBlankSlots(insertAt, prep.blankCount);
      if (!prep.reverse) insertAt += prep.blankCount;
    }
    files.forEach(function(file, fi) {
      var ext = file.name.split('.').pop().toLowerCase();
      var ph = createFileObj({
        name: file.name,
        size: file.size,
        type: ext,
        _loading: true
      });
      ph._placeholderKey = ph.id;
      file._phKey = ph._placeholderKey;
      if (fi === 0) firstPlaceholder = ph;
      S.files.splice(prep && prep.reverse ? insertAt : insertAt + fi, 0, ph);
      _newFileIds[ph.id] = true;
    });
  }
  renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();

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
      if (slotInsert && !_lastInsertedId && firstPlaceholder && file._phKey === firstPlaceholder._placeholderKey) _lastInsertedId = items[0].id;
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

    renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();

    // Yield to browser for painting — ensures user sees each file appear incrementally
    await nextFrame();
  }

  // Loading batch complete
  if (slotInsert) locateInsertedFile(_lastInsertedId);
  if (S.feat.autoDedup) removeDuplicates(true);
  renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();
  scrollActiveFileIntoView();

  _loadingBatchActive = false;
  _insertSlotIdx = -1;
  _lastInsertedId = null;

  if (_ocrQueue.length === 0 && _ocrRunning === 0) {
    _ocrToastActive = false;
    _ocrBatchTotal = 0;
    _ocrBatchAddedCount = 0;
    toastDone(added > 0 ? '已加载 ' + added + ' 张发票' : '文件加载失败');
  } else {
    _ocrBatchAddedCount = added;
  }
}

// Incremental loading: read files one-by-one, render in small batches.
// Strategy: skeleton placeholders (stable layout) + parallel background load + batch render every 3 files.
async function processFilesIncremental(paths) {
  var total = paths.length;
  var added = 0;
  var startTime = Date.now();
  var slotInsert = _insertSlotIdx >= 0;
  _loadingBatchActive = true;

  // 1. Create ALL skeleton placeholders immediately
  var placeholders = [];
  var insertAt = S.files.length;
  var replaceIdx = -1;
  if (_insertSlotIdx >= 0) {
    var prep = prepareSlotInsertion();
    replaceIdx = prep.replaceIdx;
    insertAt = prep.insertAt;
  }
  if (replaceIdx >= 0) {
    var targetPh = S.files[replaceIdx];
    var firstPath = paths[0];
    var firstParts = firstPath.split(/[\\/]/);
    targetPh._placeholder = false;
    targetPh.checked = true;
    targetPh.name = firstParts[firstParts.length - 1];
    targetPh.size = 0;
    targetPh.type = '';
    targetPh._loading = true;
    targetPh._placeholderKey = targetPh.id;
    _newFileIds[targetPh.id] = true;
    placeholders.push(targetPh);
    for (var ri = 1; ri < paths.length; ri++) {
      var nameParts = paths[ri].split(/[/\\]/);
      var rph = createFileObj({ name: nameParts[nameParts.length - 1], size: 0, type: '', _loading: true });
      rph._placeholderKey = rph.id;
      S.files.splice(prep.reverse ? replaceIdx : replaceIdx + ri, 0, rph);
      _newFileIds[rph.id] = true;
      placeholders.push(rph);
    }
  } else {
    if (_insertSlotIdx >= 0) {
      insertBlankSlots(insertAt, prep.blankCount);
      if (!prep.reverse) insertAt += prep.blankCount;
    }
    paths.forEach(function(p, pi) {
      var nameParts = p.split(/[/\\]/);
      var name = nameParts[nameParts.length - 1];
      var ph = createFileObj({ name: name, size: 0, type: '', _loading: true });
      ph._placeholderKey = ph.id;
      S.files.splice(prep && prep.reverse ? insertAt : insertAt + pi, 0, ph);
      _newFileIds[ph.id] = true;
      placeholders.push(ph);
    });
  }
  renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();

  document.getElementById('fileList').classList.add('batch-loading');
  toastLoading('加载中 0/' + total);

  if (S.feat.ocrEnabled) { _ocrBatchTotal = total; }

  // 2. Batch read all files in one IPC call
  var fileDataMap = {};
  try {
    var allFileData = await invoke('open_invoice_files', { paths: paths });
    if (allFileData && allFileData.length > 0) {
      for (var ai = 0; ai < allFileData.length; ai++) {
        fileDataMap[allFileData[ai].path || ''] = allFileData[ai];
      }
    }
  } catch (err) {
    console.error('Batch read error:', err);
  }

  // 3. Start all renders in parallel, then process results incrementally
  var loadPromises = placeholders.map(function(ph, pi) {
    var path = paths[pi];
    var fd = fileDataMap[path];
    if (!fd) return Promise.resolve(null);
    return loadFileFromDataUrlFast(fd).catch(function(err) {
      console.error('Load error:', fd.name, err);
      return null;
    });
  });

  // 处理任意完成的 Promise，而不是按顺序
  var remaining = placeholders.slice();
  var promises = loadPromises.slice();
  var completedCount = 0;

  while (remaining.length > 0) {
    // 等待任意一个完成
    var winner = await Promise.race(
      promises.map(function(p, i) {
        return p
          .then(function(r) { return { result: r, idx: i, success: true }; })
          .catch(function() { return { idx: i, success: false }; });
      })
    );

    // 找到对应的索引并处理
    var ph = remaining[winner.idx];
    var phIdx = S.files.indexOf(ph);
    remaining.splice(winner.idx, 1);
    promises.splice(winner.idx, 1);
    completedCount++;

    if (phIdx >= 0 && winner.success && winner.result) {
      var items = Array.isArray(winner.result) ? winner.result : [winner.result];
      items.forEach(function(it) { _newFileIds[it.id] = true; });
      if (slotInsert && !_lastInsertedId && ph === placeholders[0]) _lastInsertedId = items[0].id;
      S.files.splice.apply(S.files, [phIdx, 1].concat(items));
      added += items.length;
    } else if (phIdx >= 0) {
      S.files.splice(phIdx, 1);
    }

    var ocrRemaining = _ocrQueue.length + _ocrRunning;
    var isLast = (completedCount >= total);
    if (isLast) {
      if (ocrRemaining > 0 && S.feat.ocrEnabled) {
        var ocrDone2 = _ocrBatchTotal > 0 ? _ocrBatchTotal - ocrRemaining : 0;
        toastLoading('加载完成，识别中 ' + ocrDone2 + '/' + _ocrBatchTotal);
      }
    } else {
      if (ocrRemaining > 0 && S.feat.ocrEnabled) {
        var ocrDone = _ocrBatchTotal > 0 ? _ocrBatchTotal - ocrRemaining : 0;
        toastLoading('加载中 ' + completedCount + '/' + total + '，识别中 ' + ocrDone + '/' + _ocrBatchTotal);
      } else {
        toastLoading('加载中 ' + completedCount + '/' + total);
      }
    }

    renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();
    await nextFrame();
  }

  if (slotInsert) locateInsertedFile(_lastInsertedId);
  if (S.feat.autoDedup) removeDuplicates(true);
  renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();
  scrollActiveFileIntoView();

  _loadingBatchActive = false;
  _insertSlotIdx = -1;
  _lastInsertedId = null;
  document.getElementById('fileList').classList.remove('batch-loading');

  if (_ocrQueue.length === 0 && _ocrRunning === 0) {
    _ocrToastActive = false;
    _ocrBatchTotal = 0;
    _ocrBatchAddedCount = 0;
    var elapsed = Date.now() - startTime;
    var minToastDelay = Math.max(300, 800 - elapsed);
    if (added > 0) {
      var doneMsg = '已加载 ' + added + ' 张发票';
      setTimeout(function() { toast(doneMsg, 2500); }, minToastDelay);
    } else {
      toast('文件加载失败');
    }
  } else {
    _ocrBatchAddedCount = added;
  }
}

// NOTE: loadFile(), loadFileFromDataUrl(), loadPdfFromDataUrl(), loadPdfFromDataUrlFast() removed.
// PDF.js removed in v1.7.1 — all PDF rendering via WinRT native, all text extraction via PP-OCRv6.

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


/** Yield to browser for reliable painting — double rAF ensures at least one frame is painted */
function nextFrame() { return new Promise(function(r) { requestAnimationFrame(function() { requestAnimationFrame(r); }); }); }
var _activeFileIdx = -1;   // Index of currently active/highlighted file in sidebar
var _printedMap = {};      // Printed state cache: {filePath: true}
var _restoreFilePaths = null; // File paths to restore on startup
var _isRestoringFiles = false; // True while restoring files (skip OCR)
var _listDrag = null;          // 左侧列表拖拽排序状态机 (#26)
var _listDragBound = false;    // 列表拖拽事件只绑定一次
var _listDragSuppressClick = false; // 拖拽松手后吞掉浏览器派发的 click
var _listDragHintShown = false; // 本次会话是否已提示过列表拖拽手势

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
  if (_isRestoringFiles) return; // Skip OCR during file list restoration
  // Skip OCR if PDF text extraction already covered all key fields
  if (fileObj._pdfTextExtracted && fileObj.sellerName && fileObj.amountTax > 0) {
    console.log('[OCR] PDF文字提取已覆盖关键字段，跳过OCR');
    return;
  }
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
      ocrPromise = applyOcr(fileObj, '', fileObj._filePath);
    } else {
      ocrPromise = downsampleForOcr(dataUrl, ocrMaxDim()).then(function(ocrDataUrl) {
        return applyOcr(fileObj, ocrDataUrl);
      });
    }
    return ocrPromise.then(function() {
      fileObj._ocrPending = false;
      if (S.feat.autoDedup) {
        var autoRemoved = removeDuplicates(true);
        if (autoRemoved) { updatePreview(); updatePrintBtn(); updateSummaryBtn(); }
      }
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
  if (f._ocrPending) {
    return '<span class="ocr-spinner" title="识别中"></span>';
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
  var pd = f._printed ? '<span class="printed-dot" title="已打印">\u2713</span>' : '';
  if (S.fileView === 'grid') {
    // grid 卡片：更新 card-meta（与 renderFileList grid 分支字段顺序一致）
    var cardMetaEl = items[idx].querySelector('.card-meta');
    if (cardMetaEl) {
      var gdupb = f._dup ? '<span class="dup-badge" title="检测到重复发票">\u26A0</span>' : '';
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
  // Update per-file OCR button state (both views share .ocr-btn)
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
function applyPdfTextToResults(results, pdfPath) {
  if (!results || results.length === 0) return;
  if (!S.feat.pdfTextEnabled) return;
  var pageIndices = results.map(function(r) { return r._pdfPageIdx; });
  invoke('extract_pdf_texts', {
    pdfPath: pdfPath,
    pageIndices: pageIndices
  }).then(function(pdfTextMap) {
    results.forEach(function(r) {
      var pdfText = pdfTextMap[r._pdfPageIdx];
      if (pdfText && pdfText.lines && pdfText.lines.length > 0) {
        applyPdfTextResult(r, pdfText);
        updateFileItem(r);
        updateAmountSummary();
      } else if (hasOcr && S.feat.ocrEnabled) {
        console.log('[PDF文字提取] 文本层为空(无CMap/扫描件)，自动回退OCR');
        applyOcrAsync(r, r.previewUrl);
      }
    });
  }).catch(function(err) {
    console.warn('[PDF文字提取] 批量提取失败，回退单页模式:', err);
    results.forEach(function(r) {
      invoke('extract_pdf_text', {
        pdfPath: r._pdfPath,
        pageIdx: r._pdfPageIdx
      }).then(function(pdfText) {
        if (pdfText && pdfText.lines && pdfText.lines.length > 0) {
          applyPdfTextResult(r, pdfText);
          updateFileItem(r);
          updateAmountSummary();
        } else if (hasOcr && S.feat.ocrEnabled) {
          applyOcrAsync(r, r.previewUrl);
        }
      }).catch(function() {
        if (hasOcr && S.feat.ocrEnabled) applyOcrAsync(r, r.previewUrl);
      });
    });
  });
}

function buildPdfResults(pages, id, name, size, filePath) {
  var results = [];
  for (var p = 0; p < pages.length; p++) {
    var pg = pages[p];
    var fileObj = createFileObj({
      id: id + '_p' + (p + 1),
      name: pages.length > 1 ? name.replace(/\.pdf$/i, '') + '_第' + (p + 1) + '页.pdf' : name,
      size: size, type: 'pdf', previewUrl: pg.imageDataUrl,
      ow: pg.width || 0, oh: pg.height || 0,
      renderDpi: pg.renderDpi || PDF_RENDER_DPI,
      pdfPath: filePath, pdfPageIdx: p
    });
    results.push(fileObj);
  }
  return results;
}

function loadPdfImages(results) {
  return Promise.all(results.map(function(r) {
    return new Promise(function(resolve) {
      var img = new Image();
      img.src = r.previewUrl;
      img.onload = function() { r.img = img; resolve(r); };
      img.onerror = function() { resolve(r); };
    });
  }));
}

function loadFileFromDataUrlFast(fd) {
  var name = fd.name, dataUrl = fd.dataUrl, size = fd.size, ext = fd.ext, filePath = fd.path;
  return new Promise(function(resolve) {
    var id = 'f' + Date.now() + Math.random().toString(36).slice(2);

    if (ext === 'pdf') {
      if (isTauri && invoke && filePath) {
        var renderFn = _winrtPdfAvailable ? 'render_pdf_pages' : 'render_pdf_pages_pdfium';
        var renderLabel = _winrtPdfAvailable ? 'WinRT' : 'PDFium';
        invoke(renderFn, { pdfPath: filePath, dpi: PDF_PREVIEW_DPI, useJpeg: true }).then(async function(pages) {
          if (pages && pages.length > 0) {
            var results = buildPdfResults(pages, id, name, size, filePath);
            resolve(results.length === 1 ? results[0] : results);

            loadPdfImages(results);
            applyPdfTextToResults(results, filePath);

            results.forEach(function(r) {
              if (S.feat.ocrEnabled) applyOcrAsync(r, r.previewUrl);
            });
            return;
          }
          toast('PDF 渲染结果为空: ' + name);
          resolve(null);
        }).catch(function(err) {
          console.error('[PDF] ' + renderLabel + ' rendering failed:', err);
          if (renderFn === 'render_pdf_pages') {
            _winrtPdfAvailable = false;
            console.warn('[PDF] WinRT failed, trying PDFium fallback...');
            invoke('render_pdf_pages_pdfium', { pdfPath: filePath, dpi: PDF_PREVIEW_DPI, useJpeg: true }).then(async function(pages2) {
              if (pages2 && pages2.length > 0) {
                var results2 = buildPdfResults(pages2, id, name, size, filePath);
                resolve(results2.length === 1 ? results2[0] : results2);

                loadPdfImages(results2);
                applyPdfTextToResults(results2, filePath);

                results2.forEach(function(r) {
                  if (S.feat.ocrEnabled) applyOcrAsync(r, r.previewUrl);
                });
                return;
              }
              toast('PDF 渲染失败: ' + name);
              resolve(null);
            }).catch(function(err2) {
              console.error('[PDF] PDFium fallback also failed:', err2);
              var errMsg = String(err2 || '');
              if (errMsg.indexOf('pdfium.dll') >= 0 || errMsg.indexOf('不可用') >= 0) {
                showPdfiumMissing('当前系统的 PDF 组件不可用，需要下载 PDFium 渲染引擎才能加载 PDF 文件。');
              } else {
                toast('PDF 渲染失败: ' + name);
              }
              resolve(null);
            });
          } else {
            var errMsg2 = String(err || '');
            if (errMsg2.indexOf('pdfium.dll') >= 0 || errMsg2.indexOf('不可用') >= 0) {
              showPdfiumMissing('当前系统的 PDF 组件不可用，需要下载 PDFium 渲染引擎才能加载 PDF 文件。');
            } else {
              toast('PDF 渲染失败: ' + name);
            }
            resolve(null);
          }
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
          filePath: filePath || '',  // needed for batch rename
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
          invoiceType: info.invoiceType || '',
          _isToll: !!info.isToll,
          // OFD page dimensions for layout
          ow: payload.img.naturalWidth,
          oh: payload.img.naturalHeight,
          // Mark as OFD source for PDF generation (FlateDecode)
          _ofdPage: true
        });
        resolve(fileObj);
        // Fallback OCR: OFD XML 未提取到有效数据时，以 OCR 作补充
        if (S.feat.ocrEnabled && !info.amountTax && !info.amountNoTax && !info.sellerName) {
          applyOcrAsync(fileObj, payload.pngUrl);
        }
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
    // XML 数电票: structured data only, no visual layout
    else if (ext === 'xml' && isTauri && invoke && filePath) {
      invoke('parse_xml_invoice', { xmlPath: filePath }).then(function(info) {
        var fileObj = createFileObj({
          id: id, name: name, size: size, type: 'xml',
          filePath: filePath || '',
          // Structured data from XML — skip OCR
          amountTax: info.amountTax || 0,
          amountNoTax: info.amountNoTax || 0,
          taxAmount: info.taxAmount || 0,
          sellerName: info.sellerName || '',
          sellerCreditCode: info.sellerTaxId || '',
          invoiceNo: info.invoiceNo || '',
          invoiceDate: info.invoiceDate || '',
          buyerName: info.buyerName || '',
          buyerCreditCode: info.buyerTaxId || '',
          invoiceType: info.invoiceType || '',
          _isToll: !!info.isToll,
          // XML has no preview image — use placeholder dimensions
          ow: 0, oh: 0,
          _xmlInvoice: true
        });
        resolve(fileObj);
      }).catch(function(err) {
        console.warn('[XML] parse_xml_invoice failed:', err);
        toast('XML 发票解析失败: ' + String(err));
        resolve(null);
      });
      return;
    }
    else if (ext === 'xml') {
      toast('XML 格式请使用桌面版打开');
      resolve(null);
    }
    else {
      if (!dataUrl) { resolve(null); return; }
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
    else if (ext === 'xml') {
      toast('XML 格式请使用桌面版打开');
      resolve(null);
    }
    else {
      toast('不支持的格式: ' + ext);
      resolve(null);
    }
  });
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

function setFileFilter(filter) {
  if (filter === 'duplicates') { selectDuplicateExtras(); return; }
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
    if (f._placeholder || f._loading || !f._dup) {
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

function getFilteredFiles() {
  var files = S.files;
  if (S.fileFilter === 'duplicates') return files.filter(function(f) { return f._dup; });
  if (S.printedFilter === 'all') return files;
  return files.filter(function(f) {
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
  var dupCount = S.files.filter(function(f) { return f._dup && !f._placeholder; }).length;
  var dupEl = document.getElementById('duplicateCount');
  if (dupEl) dupEl.textContent = dupCount ? '(' + dupCount + ')' : '';
}

// 删除每组第一份之后的重复项。仅处理按发票号判定的可靠重复（no: key）；
// sum:（同销售方+金额+日期）疑似重复可能是同日同额的两张真发票，只标记不删除，
// 交由人工勾选处理。无 key、加载骨架、排版占位一律不动。
// silent=true 为自动去重路径（加载完成/OCR 识别后），删除后仍会 toast 告知用户。
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
    renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();
    toast(removed ? '已删除 ' + removed + ' 个重复项，保留每组第一份' : '未发现可删除的重复项');
  } else if (removed) {
    toast('已自动去重：删除 ' + removed + ' 个重复项（每组保留第一份）');
  }
  return removed;
}

function renderFileList() {
  updateDuplicateMarks();
  var list = document.getElementById('fileList');
  var scrollTop = list.scrollTop;
  var filtered = getFilteredFiles();
  var realCount = filtered.filter(function(f) { return !f._placeholder; }).length;
  var sel = filtered.filter(function(f) { return f.checked; }).length;
  document.getElementById('fileCount').textContent = realCount + ' 张，已选 ' + sel;
  syncSelectAllBtn();
  syncDeleteBtn();
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
      if (f._placeholder) {
        return '<div class="file-item file-card placeholder-item" data-idx="' + i + '"' + hideStyle + '>' +
          '<div class="file-thumb"><div class="blank-thumb">\u25A6</div></div>' +
          '<div class="card-name">空白占位</div>' +
          '<div class="card-meta"><button class="ib card-ib danger" onclick="rmFile(' + i + ')" title="删除空白占位">\u2715</button></div></div>';
      }
      var gcb = f.copies > 1 ? '<span class="copy-badge">' + f.copies + '\u4efd</span>' : '';
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
          (hasOcr ? (f._ocrPending
            ? '<button class="ib card-ib ocr-btn" disabled title="识别中"><span class="ocr-spinner"></span></button>'
            : '<button class="ib card-ib ocr-btn" onclick="ocrFile(' + i + ')" title="OCR识别">\uD83D\uDD0D</button>') : '') +
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
    if (f._placeholder) {
      var pMeta = '<div class="file-meta-left"><span class="blank-badge">空白</span></div>' +
        '<div class="file-meta-sep"></div>' +
        '<div class="file-meta-right">' +
        '<button class="ib sort-btn' + (i === 0 ? ' disabled' : '') + '" onclick="moveFile(' + i + ',-1)" title="上移">\u25B2</button>' +
        '<button class="ib sort-btn' + (i === S.files.length - 1 ? ' disabled' : '') + '" onclick="moveFile(' + i + ',1)" title="下移">\u25BC</button>' +
        '<button class="ib danger" onclick="rmFile(' + i + ')" title="删除空白占位">\u2715</button></div>';
      return '<div class="file-item placeholder-item" data-idx="' + i + '"' + hideStyle + '>' +
        '<div class="file-check disabled"></div>' +
        '<div class="file-thumb"><div class="blank-thumb">\u25A6</div></div>' +
        '<div class="file-info"><div class="file-name">空白占位</div><div class="file-meta">' + pMeta + '</div></div></div>';
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
    var ocrBtnHtml = hasOcr
      ? (f._ocrPending
        ? '<button class="ib ocr-btn" disabled title="识别中"><span class="ocr-spinner"></span></button>'
        : '<button class="ib ocr-btn" onclick="ocrFile(' + i + ')" title="OCR识别">\uD83D\uDD0D</button>')
      : '';
    var pd = f._printed ? '<span class="printed-dot" title="已打印">✓</span>' : '';
    var metaActions = f._loading
      ? '<button class="ib danger" onclick="rmFile(' + i + ')">\u2715</button>'
      : '<div class="file-meta-left">' + pd + '<span class="file-size">' + fmtSize(f.size) + '</span>' + cb + rb + dupb + ab + '</div>' +
        '<div class="file-meta-sep"></div>' +
        '<div class="file-meta-right">' +
        '<button class="ib sort-btn' + (i === 0 ? ' disabled' : '') + '" onclick="moveFile(' + i + ',-1)" title="上移">\u25B2</button>' +
        '<button class="ib sort-btn' + (i === S.files.length - 1 ? ' disabled' : '') + '" onclick="moveFile(' + i + ',1)" title="下移">\u25BC</button>' +
        ocrBtnHtml + '<button class="ib" onclick="rotFile(' + i + ')" title="旋转90°">\u21BB</button><button class="ib danger" onclick="rmFile(' + i + ')">\u2715</button></div>';
    return '<div class="' + cls + '" data-idx="' + i + '"' + hideStyle + ' onclick="clickFileItem(' + i + ',event)" ondblclick="openInvModal(' + i + ')">' +
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
function togCheck(i) { if (S.files[i]._placeholder) return; S.files[i].checked = !S.files[i].checked; renderFileList(); updatePreview(); updateSummaryBtn(); }
function selectAll() { S.files.forEach(function(f) { if (!f._placeholder) f.checked = true; }); renderFileList(); updatePreview(); updateSummaryBtn(); }
function deselectAll() { S.files.forEach(function(f) { f.checked = false; }); renderFileList(); updatePreview(); updateSummaryBtn(); }
function toggleSelectAll() {
  var selectable = S.files.filter(function(f) { return !f._placeholder; });
  var all = selectable.length > 0 && selectable.every(function(f) { return f.checked; });
  if (all) deselectAll(); else selectAll();
}
function syncSelectAllBtn() {
  var btn = document.getElementById('selectAllBtn');
  if (!btn) return;
  var selectable = S.files.filter(function(f) { return !f._placeholder; });
  var all = selectable.length > 0 && selectable.every(function(f) { return f.checked; });
  btn.textContent = all ? '\u25FB' : '\u2611';
  btn.title = all ? '取消全选' : '全选';
}
function syncDeleteBtn() {
  var btn = document.getElementById('deleteBtn');
  if (!btn) return;
  var n = S.files.filter(function(f) { return f.checked; }).length;
  btn.disabled = n === 0;
  btn.title = n > 0 ? '删除选中 ' + n + ' 张' : '未勾选发票';
}
function deleteSelected() {
  if (!S.files.some(function(f) { return f.checked; })) return;
  var active = _activeFileIdx >= 0 ? S.files[_activeFileIdx] : null;
  S.files = S.files.filter(function(f) { return !f.checked; });
  _activeFileIdx = active ? S.files.indexOf(active) : -1;
  renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn();
}
function rmFile(i) { S.files.splice(i, 1); if (_activeFileIdx === i) _activeFileIdx = -1; else if (_activeFileIdx > i) _activeFileIdx--; renderFileList(); updatePreview(); updatePrintBtn(); updateSummaryBtn(); }
function rotFile(i) { S.files[i].rotation = (S.files[i].rotation + 90) % 360; renderFileList(); updatePreview(); }
function rotateSelected() {
  var f = getSelectedFileObj();
  if (!f) { toast('请先选中版面中的发票'); return; }
  var i = S.files.indexOf(f);
  if (i < 0) return;
  rotFile(i);
}
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
    return !f._placeholder && !f._loading && !f._ocrPending && !(f.amountTax > 0 || f.amountNoTax > 0);
  });
  if (targets.length === 0) { toast('没有需要识别的发票'); return; }
  _ocrBatchTotal = targets.length;
  updateOcrAllBtn();
  toastLoading('识别中，共 ' + targets.length + ' 张...');
  targets.forEach(function(f) { applyOcrAsync(f, f.previewUrl); });
}
function clearAll() {
  if (!S.files.length) return;
  if (!confirm('确认清除所有发票？')) return;
  S.files = [];
  _activeFileIdx = -1;
  _printedMap = {};
  saveSettings();
  renderFileList();
  updatePreview();
  updatePrintBtn();
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
    updatePrintBtn();
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
    // 若当前 _activeFileIdx 已在当前页范围内,保持不变(用户点击列表项时不应被覆盖)
    if (_activeFileIdx >= 0) {
      var curFile = S.files[_activeFileIdx];
      var curActiveIdx = activeFiles.indexOf(curFile);
      if (curActiveIdx >= pageStart && curActiveIdx < pageStart + perPage) {
        return;
      }
    }
    var firstFileOnPage = activeFiles[pageStart];
    var newIdx = S.files.indexOf(firstFileOnPage);
    if (newIdx !== _activeFileIdx) {
      _activeFileIdx = newIdx;
      updateActiveFileHighlight();
    }
  }
}
// =====================================================
// File list sorting — move up / move down + drag reorder (#26)
// =====================================================
function moveFile(i, dir) {
  swapFiles(i, i + dir);
}

function swapFiles(ia, ib) {
  if (ia === ib || ia < 0 || ib < 0 || ia >= S.files.length || ib >= S.files.length) return;
  var tmp = S.files[ia];
  S.files[ia] = S.files[ib];
  S.files[ib] = tmp;
  // Update active file index to follow the moved item
  if (_activeFileIdx === ia) { _activeFileIdx = ib; }
  else if (_activeFileIdx === ib) { _activeFileIdx = ia; }
  renderFileList();
  updatePreview();
  scrollToListItem(ib);
}

// Insert file at neighbor position (before/after target instead of swapping)
function moveFileTo(fromIdx, toIdx, zone) {
  if (fromIdx === toIdx || fromIdx < 0 || toIdx < 0 || fromIdx >= S.files.length || toIdx >= S.files.length) return;
  var a = S.files[fromIdx];
  var b = S.files[toIdx];
  if (!a || !b || a === b) return;
  S.files.splice(fromIdx, 1);
  var insertAt = zone === 'after' ? S.files.indexOf(b) + 1 : S.files.indexOf(b);
  S.files.splice(insertAt, 0, a);
  if (insertAt === fromIdx) return; // 落点即原位，顺序未变
  _activeFileIdx = S.files.indexOf(a);
  renderFileList();
  updatePreview();
  scrollToListItem(insertAt);
}

function scrollToListItem(idx) {
  var list = document.getElementById('fileList');
  if (!list) return;
  var items = list.querySelectorAll('.file-item');
  if (items[idx]) items[idx].scrollIntoView({ block: 'nearest' });
}

// 列表拖拽排序仅在无筛选时启用：筛选态显示序 ≠ 底层序，拖拽会乱序
function canListDrag() {
  return S.fileFilter === 'all' && S.printedFilter === 'all';
}

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
  _listDrag = {
    itemEl: itemEl,
    idx: idx,
    startX: e.clientX,
    startY: e.clientY,
    moved: false,
    dropEl: null,
    dropIdx: -1,
    dropZone: ''
  };
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

// 落点分区：目标项上下各 25% 为顺位插入（指示线），中间 50% 为对调（虚线框）
// 光标落在列表项之间空白时，吸附到距离最近的列表项（gap 盲区兜底，80px 内有效）
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
  if (zone === 'swap') {
    el.classList.add('drop-target');
  } else {
    el.classList.add('drop-insert');
    el.classList.add(zone === 'before' ? 'drop-at-start' : 'drop-at-end');
  }
}

function clearListDropTarget(el) {
  if (!el) return;
  el.classList.remove('drop-target', 'drop-insert', 'drop-at-start', 'drop-at-end');
}

// 首次拖拽列表时提示两种手势的区别
function showListDragHint() {
  if (_listDragHintShown) return;
  _listDragHintShown = true;
  try {
    if (localStorage.getItem('ticketchan-list-drag-hint')) return;
    localStorage.setItem('ticketchan-list-drag-hint', '1');
  } catch (err) { return; }
  toast('拖到列表项边缘 = 顺位插入，拖到中间 = 两张对调', 4000);
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
  var ocrText = f._ocrText || '';
  var ocrHtml = ocrText ? '<div style="margin-top:12px;border-top:1px solid var(--border);padding-top:10px"><div style="display:flex;align-items:center;gap:6px;cursor:pointer;margin-bottom:6px" onclick="this.nextElementSibling.classList.toggle(\'hidden\');this.querySelector(\'.arrow\').textContent=this.nextElementSibling.classList.contains(\'hidden\')?\'▶\':\'▼\'"><span class="arrow" style="font-size:10px;color:var(--text-muted)">▶</span><span style="font-size:12px;font-weight:600;color:var(--primary)">🔍 OCR识别全文</span><span style="font-size:10px;color:var(--text-muted)">(点击展开)</span></div><div class="hidden" style="position:relative"><pre style="margin:0;padding:8px 10px;background:var(--surface2);border-radius:6px;max-height:260px;overflow:auto;white-space:pre-wrap;word-break:break-all;font-size:11px;line-height:1.5;font-family:Consolas,monospace;border:1px solid var(--border)">' + escHtml(ocrText) + '</pre><button class="btn btn-sm" style="position:absolute;top:6px;right:6px;padding:3px 8px;font-size:11px;opacity:0.7" onclick="event.stopPropagation();copyOcrText(this)" title="复制OCR文本">📋 复制</button></div></div>' : '<div style="margin-top:12px;border-top:1px solid var(--border);padding-top:10px;font-size:11px;color:var(--text-muted)">⏳ OCR 全文尚未识别</div>';
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
    '</div>' +
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
    syncSidebarToSelectedSlot();
  }
}

// 右侧选中版面槽位时，左侧文件列表同步高亮并滚动到对应发票
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
  syncEnhanceBtn(f);
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

// =====================================================
// Text Enhancement — 浅色/模糊图片发票增强（Rust 全分辨率处理）
// =====================================================
var ENHANCE_IMAGE_TYPES = ['jpg', 'jpeg', 'png', 'bmp', 'webp', 'tiff', 'tif'];

/** 仅图片文件（有磁盘路径，走 Rust 原图处理）可增强；PDF/OFD/XML 不支持 */
function canEnhanceFile(f) {
  return !!(f && f._filePath && !f._xmlInvoice && ENHANCE_IMAGE_TYPES.indexOf(f.type) >= 0);
}

/** 单票调整面板「文本增强」开关：增强作用于原图全分辨率，打印清晰度不受影响 */
function toggleTextEnhance() {
  var f = getSelectedFileObj();
  if (!canEnhanceFile(f) || f._enhancing) return;

  // 已增强 → 还原原图
  if (f._enhanced) {
    f.previewUrl = f._origPreviewUrl;
    f.img = f._origImg;
    f._origPreviewUrl = '';
    f._origImg = null;
    f._enhanced = false;
    f.trimmedUrl = null; // 白边缓存基于原图，需按还原后的图重算
    updateAdjPanel();
    updatePreview();
    renderFileList();
    if (S.feat.trimWhite) processTrim();
    return;
  }

  f._enhancing = true;
  updateAdjPanel();
  invoke('enhance_image', { filePath: f._filePath }).then(function(dataUrl) {
    var img = new Image();
    img.onload = function() {
      f._origPreviewUrl = f.previewUrl;
      f._origImg = f.img;
      f.previewUrl = dataUrl;
      f.img = img;
      // ow/oh 不变：增强不改动尺寸（EXIF 方向已在加载时与增强时一致烘焙）
      f._enhanced = true;
      f._enhancing = false;
      f.trimmedUrl = null; // 白边缓存基于原图，需按增强后的图重算
      updateAdjPanel();
      updatePreview();
      renderFileList();
      if (S.feat.trimWhite) processTrim();
      toast('已增强：' + f.name);
    };
    img.onerror = function() {
      f._enhancing = false;
      updateAdjPanel();
      toast('增强结果加载失败');
    };
    img.src = dataUrl;
  }).catch(function(e) {
    f._enhancing = false;
    updateAdjPanel();
    toast('文本增强失败: ' + String(e));
  });
}

/** 同步「文本增强」按钮状态（由 updateAdjPanel 调用） */
function syncEnhanceBtn(f) {
  var btn = document.getElementById('btnTextEnhance');
  if (!btn) return;
  if (!canEnhanceFile(f)) {
    btn.disabled = true;
    btn.classList.remove('btn-primary');
    btn.textContent = '✨ 文本增强';
    btn.title = '仅支持图片文件（PDF/OFD 内嵌图像暂不支持）';
  } else if (f._enhancing) {
    btn.disabled = true;
    btn.classList.remove('btn-primary');
    btn.textContent = '增强中…';
  } else if (f._enhanced) {
    btn.disabled = false;
    btn.classList.add('btn-primary');
    btn.textContent = '✓ 已增强（点击还原）';
  } else {
    btn.disabled = false;
    btn.classList.remove('btn-primary');
    btn.textContent = '✨ 文本增强';
    btn.title = '浅色/模糊图片发票：文字加深、边缘锐化（仅作用于图片文件，不影响原文件）';
  }
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
    // If renderDpi is not set, fall back to PDF_PREVIEW_DPI (150).
    var rDpi = f.renderDpi || 150;
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

  // Offset to move wrapper from base position to target alignment.
  // 常规模式基准=居中；报销单模式基准=左上（renderPage 中 wrapper 定位分支保持一致）。
  var offsetX = 0, offsetY = 0;
  if (settings.reimburseMode) {
    if (alignH === 'center') offsetX = gapX;
    else if (alignH === 'right') offsetX = 2 * gapX;
    if (alignV === 'center') offsetY = gapY;
    else if (alignV === 'bottom') offsetY = 2 * gapY;
  } else {
    if (alignH === 'left')  offsetX = -gapX;
    if (alignH === 'right') offsetX =  gapX;
    if (alignV === 'top')   offsetY = -gapY;
    if (alignV === 'bottom') offsetY =  gapY;
  }

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
// 渲染顶部工具栏快捷排版按钮（可插拔，数量受 quickLayoutMax 限制）
function renderQuickLayoutBar() {
  var bar = document.getElementById('quickLayoutBar');
  if (!bar) return;
  var items = S.quickLayouts || [];
  var max = parseInt(S.quickLayoutMax) || 0;
  var html = '';
  items.forEach(function(q, i) {
    if (max > 0 && i >= max) return;
    var c = parseInt(q.cols) || 1, r = parseInt(q.rows) || 1;
    html += '<button class="btn btn-sm ql-btn" data-cols="' + c + '" data-rows="' + r + '" onclick="quickLayout(' + c + ',' + r + ')" title="' + r + '行' + c + '列">' + r + '×' + c + '</button>';
  });
  html += '<button class="btn btn-sm ql-btn ql-custom" onclick="openQuickLayoutManager()" title="管理快捷布局" style="font-size:12px">⚙</button>';
  bar.innerHTML = html;
  syncToolbarHighlight(S.layout.cols, S.layout.rows);
}

// 打开快捷布局管理（切换到排版 tab 并滚动到管理区块）
function openQuickLayoutManager() {
  switchTab('settings', document.querySelectorAll('.sidebar-tab')[1]);
  renderQuickLayoutList();
  var sec = document.getElementById('quickLayoutSec');
  if (sec) sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// 渲染管理区块的布局列表
function renderQuickLayoutList() {
  var list = document.getElementById('quickLayoutList');
  if (!list) return;
  var items = S.quickLayouts || [];
  var html = '';
  items.forEach(function(q, i) {
    var c = parseInt(q.cols) || 1, r = parseInt(q.rows) || 1;
    html += '<div class="ql-item">' +
      '<input type="number" min="1" max="10" value="' + r + '" class="ql-in" onchange="updateQuickLayout(' + i + ',\'rows\',this.value)" title="行数">' +
      '<span class="ql-x">×</span>' +
      '<input type="number" min="1" max="10" value="' + c + '" class="ql-in" onchange="updateQuickLayout(' + i + ',\'cols\',this.value)" title="列数">' +
      '<button class="ib" onclick="moveQuickLayout(' + i + ',-1)" title="上移"' + (i === 0 ? ' disabled' : '') + '>⬆</button>' +
      '<button class="ib" onclick="moveQuickLayout(' + i + ',1)" title="下移"' + (i === items.length - 1 ? ' disabled' : '') + '>⬇</button>' +
      '<button class="ib danger" onclick="removeQuickLayout(' + i + ')" title="删除">✕</button>' +
      '</div>';
  });
  if (!html) html = '<div style="font-size:12px;color:var(--text-muted);padding:4px 0">暂无快捷布局，点击下方添加</div>';
  list.innerHTML = html;
}

function addQuickLayout() {
  if (!S.quickLayouts) S.quickLayouts = [];
  S.quickLayouts.push({ cols: 1, rows: 1 });
  renderQuickLayoutList();
  renderQuickLayoutBar();
  saveSettings();
}

function updateQuickLayout(i, field, val) {
  if (!S.quickLayouts || !S.quickLayouts[i]) return;
  var q = S.quickLayouts[i];
  var oldC = parseInt(q.cols) || 1, oldR = parseInt(q.rows) || 1;
  var v = normalizeQuickLayoutValue(val);
  q[field] = v;
  var c = parseInt(q.cols) || 1, r = parseInt(q.rows) || 1;
  // 若当前版面恰好是被编辑的布局，同步应用新值
  if (S.layout.cols === oldC && S.layout.rows === oldR) {
    document.getElementById('orientation').value = r > c ? 'portrait' : 'landscape';
    S.layout = { cols: c, rows: r };
    updatePreview();
  }
  renderQuickLayoutList();
  renderQuickLayoutBar();
  saveSettings();
}

function removeQuickLayout(i) {
  if (!S.quickLayouts || !S.quickLayouts[i]) return;
  S.quickLayouts.splice(i, 1);
  renderQuickLayoutList();
  renderQuickLayoutBar();
  saveSettings();
}

function moveQuickLayout(i, dir) {
  var arr = S.quickLayouts;
  if (!arr || !arr[i]) return;
  var j = i + dir;
  if (j < 0 || j >= arr.length) return;
  var tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
  renderQuickLayoutList();
  renderQuickLayoutBar();
  saveSettings();
}

function setQuickLayoutMax(v) {
  S.quickLayoutMax = normalizeQuickLayoutMax(v);
  document.getElementById('quickLayoutMax').value = S.quickLayoutMax;
  renderQuickLayoutBar();
  saveSettings();
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
var _savedPrinterName = null;
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
    reimburseMode: !!S.feat.reimburse,
    reimburseHeight: parseFloat(document.getElementById('reimburseHeight').value) || 120,
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
    footerText: S.feat.footer ? document.getElementById('footerText').value : '',
    footerMargin: (S.feat.pageNum || S.feat.printDate || S.feat.footer) ? (S.feat.customFM ? parseFloat(document.getElementById('footerMargin').value) || 0 : _autoFooterMargin()) : 0,
    customFm: S.feat.customFM,
    copies: parseInt(document.getElementById('copies').value) || 1,
    collate: S.feat.collate, duplex: S.feat.duplex,
    printerName: document.getElementById('printerSel').value || null
  };
}

// Get checked files WITHOUT copies expansion (for summary table, etc.)
function getCheckedFiles() {
  return S.files.filter(function(f) { return f.checked && !f._loading; });
}

function markFilesAsPrinted(files) {
  files.forEach(function(f) {
    if (f._placeholder) return;
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
  if (document.getElementById('pageOrder').value === 'reverse') files = files.slice().reverse();
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
  document.getElementById('stLayout').textContent = S.feat.reimburse ? '报销单' : (S.layout.rows + '×' + S.layout.cols);
  var ps = document.getElementById('paperSize').value;
  document.getElementById('stPaper').textContent = ps + ' ' + (document.getElementById('orientation').value === 'portrait' ? '纵' : '横');
  updatePrintBtn();

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
function updatePrintBtn() { document.getElementById('printBtn').disabled = !S.files.some(function(f) { return f.checked; }); }
function updateSummaryBtn() { var btn = document.getElementById('summaryBtn'); if (btn) btn.disabled = !S.files.some(function(f) { return f.checked; }); }

// =====================================================
// Save settings & Preferences
// =====================================================
function saveSettings() {
  var o = {
    layout: { cols: S.layout.cols, rows: S.layout.rows },
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
    colorMode: document.getElementById('colorMode').value,
    pageOrder: document.getElementById('pageOrder').value,
    printMode: document.getElementById('printMode').value,
    printerName: document.getElementById('printerSel').value || null,
    feat: {}
  };
  var featKeys = ['cutline','number','border','trimWhite','watermark','collate','duplex','pageNum','printDate','footer','autoOpenPdf','customFM','slotAdjMemory','fileListMemory','autoDedup','reimburse'];
  featKeys.forEach(function(k) { o.feat[k] = S.feat[k]; });
  o.reimburseHeight = document.getElementById('reimburseHeight').value;
  o.quickLayouts = cloneQuickLayouts(S.quickLayouts);
  o.quickLayoutMax = normalizeQuickLayoutMax(S.quickLayoutMax);
  o.fileView = S.fileView;
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
  // Persist rename template and separator
  if (_renameTemplate && _renameTemplate.length > 0) o.renameTemplate = _renameTemplate;
  if (_renameSeparator) o.renameSeparator = _renameSeparator;
  // Persist per-file notes (keyed by file name)
  var notesMap = {};
  S.files.forEach(function(f) { if (f.note && f.name) notesMap[f.name] = f.note; });
  if (Object.keys(notesMap).length > 0) o.summaryNotes = notesMap;
  // Save printed state (always, regardless of fileListMemory switch)
  var printedMap = {};
  S.files.forEach(function(f) {
    var key = f._filePath || f._pdfPath;
    if (key && f._printed) printedMap[key] = true;
  });
  o.printedMap = printedMap;
  // Save file paths only when memory is enabled (always write to clear stale data)
  if (S.feat.fileListMemory) {
    var filePaths = [];
    S.files.forEach(function(f) {
      var p = f._filePath || f._pdfPath;
      if (p && filePaths.indexOf(p) < 0) filePaths.push(p);
    });
    o.filePaths = filePaths;
  } else {
    o.filePaths = [];
  }
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
  // Restore the exact saved list, including an intentionally empty list.
  // Older configs have no reliable marker distinguishing the old defaults
  // from a user-customized three-item list, so do not overwrite them.
  if (Array.isArray(o.quickLayouts)) S.quickLayouts = cloneQuickLayouts(o.quickLayouts);
  if (o.quickLayoutMax != null) S.quickLayoutMax = normalizeQuickLayoutMax(o.quickLayoutMax);
  if (o.fileView === 'grid') S.fileView = 'grid';
  syncFileViewBtn();
  document.getElementById('quickLayoutMax').value = S.quickLayoutMax;
  renderQuickLayoutBar();
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
  if (o.colorMode) document.getElementById('colorMode').value = o.colorMode;
  if (o.pageOrder) document.getElementById('pageOrder').value = o.pageOrder;
  if (o.printMode) document.getElementById('printMode').value = o.printMode;
  if (o.printerName) _savedPrinterName = o.printerName;
  if (o.feat) {
    var featMap = {
      cutline: 'toggleCutline', number: 'toggleNumber', border: 'toggleBorder',
      trimWhite: 'toggleTrimWhite', watermark: 'toggleWatermark', collate: 'toggleCollate',
      duplex: 'toggleDuplex', pageNum: 'togglePageNum', printDate: 'toggleDate',
      footer: 'toggleFooter', autoOpenPdf: 'toggleAutoOpenPdf', customFM: 'toggleCustomFM',
      slotAdjMemory: 'toggleSlotAdjMemory',
      fileListMemory: 'toggleFileListMemory',
      autoDedup: 'toggleAutoDedup',
      reimburse: 'toggleReimburse'
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
  if (o.reimburseHeight != null) document.getElementById('reimburseHeight').value = o.reimburseHeight;
  syncReimburseUI();
  // Restore summary table column selection
  if (o.summaryCols && Array.isArray(o.summaryCols) && o.summaryCols.length > 0) {
    _summaryActiveCols = o.summaryCols;
    // v2.0.6 migration: ensure note column is included for existing users
    if (_summaryActiveCols.indexOf('note') < 0) _summaryActiveCols.push('note');
  }
  // Restore rename template and separator
  if (o.renameTemplate && Array.isArray(o.renameTemplate) && o.renameTemplate.length > 0) {
    _renameTemplate = o.renameTemplate;
  }
  if (o.renameSeparator) _renameSeparator = o.renameSeparator;
  // Restore per-file notes (applied when files are added)
  S._notesMap = o.summaryNotes || {};
  // Load saved per-file slot adjustments (applied when files are added)
  S._fileAdjMap = (o.fileAdjustments && S.feat.slotAdjMemory) ? o.fileAdjustments : {};
  // Restore printed state (always, regardless of switch)
  if (o.printedMap) _printedMap = o.printedMap;
  else _printedMap = {};
  // Restore file paths only when memory is enabled
  if (o.filePaths && o.filePaths.length > 0 && S.feat.fileListMemory) {
    _restoreFilePaths = o.filePaths;
  }
}

function togglePref(k, btn) {
  S.feat[k] = !S.feat[k];
  btn.classList.toggle('on', S.feat[k]);
  if (k === 'ocrEnabled') {
    try { localStorage.setItem('ticketchan-ocr-enabled', S.feat[k] ? '1' : '0'); } catch(e) {}
  }
  if (k === 'pdfTextEnabled') {
    try { localStorage.setItem('ticketchan-pdf-text-enabled', S.feat[k] ? '1' : '0'); } catch(e) {}
  }
  saveSettings();
}

// 报销单分段模式开关：开启后忽略网格/行列/间距/裁切线开关，
// 按固定段高纵向分段，段边界处强制绘制裁切线
function toggleReimburseMode(btn) {
  S.feat.reimburse = !S.feat.reimburse;
  btn.classList.toggle('on', S.feat.reimburse);
  syncReimburseUI();
  saveSettings();
  updatePreview();
}

// 同步报销单模式相关 UI 状态（选项显隐 + 网格/间距/裁切线控件置灰）
function syncReimburseUI() {
  var on = !!S.feat.reimburse;
  document.getElementById('reimburseOpts').style.display = on ? 'block' : 'none';
  ['gridSel', 'customLayoutRow', 'gapSec', 'cutlineRow'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) {
      el.style.opacity = on ? '0.4' : '';
      el.style.pointerEvents = on ? 'none' : '';
    }
  });
}

function toggleFileListMemory(btn) {
  S.feat.fileListMemory = !S.feat.fileListMemory;
  btn.classList.toggle('on', S.feat.fileListMemory);
  saveSettings();
}

function setOcrPrecision(val) {
  S.ocrPrecision = val;
  try { localStorage.setItem('ticketchan-ocr-precision', val); } catch(e) {}
}

function getSaveDir() {
  try { return localStorage.getItem('ticketchan-save-dir') || ''; } catch(e) { return ''; }
}
function setSaveDir(dir) {
  try { localStorage.setItem('ticketchan-save-dir', dir); } catch(e) {}
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
  try { localStorage.setItem('ticketchan-theme', theme); } catch(e) {}
}

function exportSettings() {
  var data = {
    layout: S.layout,
    feat: S.feat,
    ocrPrecision: S.ocrPrecision,
    paperSize: document.getElementById('paperSize').value,
    orientation: document.getElementById('orientation').value,
    copies: document.getElementById('copies').value,
    colorMode: document.getElementById('colorMode').value,
    printMode: document.getElementById('printMode').value,
    saveDir: getSaveDir(),
    quickLayouts: cloneQuickLayouts(S.quickLayouts),
    quickLayoutMax: normalizeQuickLayoutMax(S.quickLayoutMax)
  };
  var blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = '发票酱设置.json'; a.click();
  toast('设置已导出');
}

function resetSettings() {
  if (!confirm('确认恢复所有默认设置？')) return;
  S.layout = { cols: 1, rows: 1 };
  S.feat = { cutline: true, number: false, border: false, trimWhite: false, watermark: false, footer: false, customFM: false, collate: true, duplex: false, pageNum: false, printDate: false, autoOpenPdf: true, ocrEnabled: false, pdfTextEnabled: true, slotAdjMemory: false, fileListMemory: false, autoDedup: false, reimburse: false };
  S.ocrPrecision = 'standard';
  S.viewZoom = 0;
  S.quickLayouts = defaultQuickLayouts();
  S.quickLayoutMax = 0;
  document.getElementById('quickLayoutMax').value = 0;
  renderQuickLayoutBar();
  renderQuickLayoutList();
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
  document.getElementById('customW').value = 210;
  document.getElementById('customH').value = 297;
  document.getElementById('customScale').value = 100; document.getElementById('customScaleN').value = 100;
  document.getElementById('pageOrder').value = 'normal';
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
  document.getElementById('toggleCollate').classList.add('on');
  document.getElementById('toggleDuplex').classList.remove('on');
  document.getElementById('togglePageNum').classList.remove('on');
  document.getElementById('toggleDate').classList.remove('on');
  document.getElementById('toggleAutoOpenPdf').classList.add('on');
  document.getElementById('toggleOcrEnabled').classList.remove('on');
  document.getElementById('togglePdfText').classList.add('on');
  document.getElementById('toggleFooter').classList.remove('on');
  document.getElementById('toggleCustomFM').classList.remove('on');
  document.getElementById('toggleReimburse').classList.remove('on');
  document.getElementById('reimburseHeight').value = 120;
  syncReimburseUI();
  document.getElementById('footerOpts').style.display = 'none';
  document.getElementById('customFMRow').style.display = 'none';
  document.getElementById('footerMarginRow').style.display = 'none';
  document.getElementById('footerMargin').value = 8; document.getElementById('footerMarginN').value = 8;
  document.getElementById('ocrPrecision').value = 'standard';
  document.getElementById('printMode').value = 'pdfium';
  document.getElementById('themeMode').value = 'light';
  document.documentElement.classList.remove('dark');
  try { localStorage.removeItem('ticketchan-theme'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-save-dir'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-amt-mode'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-ocr-enabled'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-ocr-precision'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-pdf-text-enabled'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-settings'); } catch(e) {}
  try { localStorage.removeItem('ticketchan-print-mode'); } catch(e) {}
  _renameTemplate = ['amountTax', 'sellerName', 'invoiceNo'];
  _renameSeparator = '_';
  _summaryActiveCols = [];
  _savedPrinterName = null;
  _printedMap = {};
  S._fileAdjMap = {};
  S._notesMap = {};
  S.printedFilter = 'all';
  S.fileFilter = 'all';
  syncFilterButtons();
  renderFileList();
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
document.body.addEventListener('drop', function(e) {
  e.preventDefault();
  if (_slotUploadActive || _loadingBatchActive) { toast('当前仍在加载发票，请稍候再添加'); return; }
  if (e.dataTransfer.files.length) processFiles(Array.from(e.dataTransfer.files));
});

// 拖入提示浮层 — 桌面版由 Rust DragDropEvent::Enter/Leave 调 _tauriDragHover，
// 浏览器版 HTML5 drag 事件不触发（WebView2 原生接管），走 dragenter/dragleave 计数。
window._tauriDragHover = function(show) {
  var el = document.getElementById('dropOverlay');
  if (el) el.classList.toggle('on', !!show);
};
var _dragDepth = 0;
window.addEventListener('dragenter', function(e) {
  var types = e.dataTransfer && e.dataTransfer.types;
  if (types && Array.prototype.indexOf.call(types, 'Files') >= 0) {
    _dragDepth++;
    window._tauriDragHover(true);
  }
});
window.addEventListener('dragleave', function() {
  if (_dragDepth > 0) _dragDepth--;
  if (_dragDepth === 0) window._tauriDragHover(false);
});
window.addEventListener('drop', function() {
  _dragDepth = 0;
  window._tauriDragHover(false);
});
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
  if (!Array.isArray(paths)) return;
  if (paths.length === 0) {
    toast('不支持的文件格式，请拖入 PDF/JPG/PNG/OFD/XML 等发票文件');
    return;
  }
  (async function() {
    try {
      if (paths.length <= 3) {
        toastLoading('读取 ' + paths.length + ' 个文件...');
        var fileDataList = await invoke('open_invoice_files', { paths: paths });
        if (fileDataList && fileDataList.length > 0) {
          await processFileDataList(fileDataList);
        } else {
          toast('无法读取拖放的文件');
        }
      } else {
        await processFilesIncremental(paths);
      }
    } catch(err) {
      hideToast();
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
    var dir = localStorage.getItem('ticketchan-save-dir') || '';
    document.getElementById('saveDir').value = dir;
  } catch(e) {}
})();

(function() {
  try {
    var m = localStorage.getItem('ticketchan-amt-mode');
    if (m && (m === 'tax' || m === 'notax' || m === 'both')) {
      S.amtMode = m;
      document.getElementById('amtMode').value = m;
    }
  } catch(e) {}
})();

(function() {
  try {
    var pm = localStorage.getItem('ticketchan-print-mode');
    if (pm && (pm === 'confirm' || pm === 'direct' || pm === 'pdfium' || pm === 'pdf')) {
      document.getElementById('printMode').value = pm;
    } else {
      // 默认 PDFium 静默打印 — PDF 阅读器模式无法可靠切换打印机(Edge/Chrome 内置查看器不支持 printto)
      document.getElementById('printMode').value = 'pdfium';
    }
  } catch(e) {}
})();

// Restore OCR enabled setting
(function() {
  try {
    var v = localStorage.getItem('ticketchan-ocr-enabled');
    if (v === '1') {
      S.feat.ocrEnabled = true;
      document.getElementById('toggleOcrEnabled').classList.add('on');
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

// Restore OCR precision setting
(function() {
  try {
    var p = localStorage.getItem('ticketchan-ocr-precision');
    if (p && (p === 'fast' || p === 'standard' || p === 'precise')) {
      S.ocrPrecision = p;
      document.getElementById('ocrPrecision').value = p;
    }
  } catch(e) {}
})();

// Defaults for summary/rename — MUST be initialized before loadSettings()
// (var hoisting: declaration hoists but assignment does not; a later
// `var x = default` would overwrite the value restored by loadSettings)
var _summaryActiveCols = []; // keys of currently visible columns
var _renameTemplate = ['amountTax', 'sellerName', 'invoiceNo'];
var _renameSeparator = '_';

// Restore all layout & feature settings
loadSettings();

// Render quick layout buttons — also covers first run (loadSettings returns early with no saved data)
renderQuickLayoutBar();
renderQuickLayoutList();

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
      invoke('check_winrt_pdf').then(function(available) {
        _winrtPdfAvailable = !!available;
        if (!_winrtPdfAvailable) {
          console.warn('[PDF] WinRT PDF 组件不可用，将使用 PDFium fallback');
          invoke('check_pdfium_available').then(function(pdfiumAvail) {
            if (!pdfiumAvail) {
              showPdfiumMissing('当前系统的 PDF 组件不可用，需要下载 PDFium 渲染引擎才能加载 PDF 文件。');
            }
          }).catch(function() {});
        }
      }).catch(function() {});
      // Get app version from Rust (compiled from Cargo.toml)
      invoke('get_app_version').then(function(v) {
        APP_VERSION = v;
        var el = document.getElementById('stVersion');
        if (el) {
          el.innerHTML = 'v' + v + ' <span class="ver-check">🔄 检查更新</span>';
          el.style.cursor = 'pointer';
          el.title = '点击检查更新';
          el.onclick = function() { checkForUpdates(false); };
        }
        // Update version display in prefs panel too
        var pv = document.getElementById('prefsCurrentVersion');
        if (pv) pv.textContent = 'v' + v;
        console.log('发票酱 v' + v + ' | isTauri:', isTauri);
        // Auto-check for updates 5s after startup (silent — only shows modal if update exists)
        setTimeout(function() { checkForUpdates(true); }, 5000);
      }).catch(function() {});
      try { invoke('show_window'); } catch(e) {}
      // Restore file list from last session if memory is enabled
      if (_restoreFilePaths && _restoreFilePaths.length) {
        var pathsToRestore = _restoreFilePaths;
        _restoreFilePaths = null;
        restoreFiles(pathsToRestore);
      }
    } else {
      // Non-Tauri (browser) fallback
      var el = document.getElementById('stVersion');
      if (el) el.textContent = 'web';
    }
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function() { showApp(); bindFooterTextEvent(); setupInputWheelSupport(); initListDrag(); });
  } else {
    showApp(); bindFooterTextEvent(); setupInputWheelSupport(); initListDrag();
  }
  setTimeout(showApp, 2000);
})();

// =====================================================
// 更新检查 — GitHub Release
// =====================================================

var _UPDATE_CACHE_KEY = 'ticketchan-update-cache';
var _UPDATE_CACHE_TTL = 60 * 60 * 1000; // 1 hour
var _updateChecking = false;

/**
 * Check for updates via GitHub Releases API.
 * @param {boolean} silent — silent mode (auto-check): only show modal if update is available,
 *                           and use cached result to avoid hitting GitHub rate limits.
 *                           Manual click (silent=false) always bypasses cache.
 */
function checkForUpdates(silent) {
  if (!isTauri || !invoke) return;
  if (_updateChecking) return;
  _updateChecking = true;

  // Silent auto-check: respect cache TTL to avoid rate limits
  if (silent) {
    try {
      var cached = localStorage.getItem(_UPDATE_CACHE_KEY);
      if (cached) {
        var data = JSON.parse(cached);
        if (Date.now() - data.ts < _UPDATE_CACHE_TTL) {
          _updateChecking = false;
          if (data.info && data.info.has_update) {
            showUpdateModal(data.info);
          }
          return;
        }
      }
    } catch(e) {}
  }

  if (!silent) toast('正在检查更新...', 1500);

  invoke('check_for_updates').then(function(info) {
    _updateChecking = false;
    // Cache result for silent auto-check
    try {
      localStorage.setItem(_UPDATE_CACHE_KEY, JSON.stringify({ ts: Date.now(), info: info }));
    } catch(e) {}

    if (info.has_update) {
      showUpdateModal(info);
    } else if (!silent) {
      toast('已是最新版本 v' + info.current_version, 2500);
    }
  }).catch(function(err) {
    _updateChecking = false;
    if (!silent) toast('检查更新失败: ' + err, 4000);
    console.warn('[Update] check failed:', err);
  });
}

/**
 * Render the update modal with release info.
 */
function showUpdateModal(info) {
  var modal = document.getElementById('updateModal');
  if (!modal) return;

  document.getElementById('updateCurrentVersion').textContent = 'v' + info.current_version;
  document.getElementById('updateLatestVersion').textContent = 'v' + info.latest_version;

  var dateStr = '';
  if (info.published_at) {
    try {
      var d = new Date(info.published_at);
      dateStr = d.toLocaleDateString('zh-CN');
    } catch(e) { dateStr = info.published_at; }
  }
  document.getElementById('updatePubDate').textContent = dateStr;

  // Release notes — escape HTML then convert basic markdown
  var notes = info.release_notes || '（无更新说明）';
  var escaped = notes
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\r\n/g, '\n')
    .replace(/\n/g, '<br>');
  document.getElementById('updateNotes').innerHTML = escaped;

  // Build asset list (download links)
  var assetsHtml = '';
  if (info.assets && info.assets.length) {
    assetsHtml = '<div class="update-assets-title">下载文件：</div><ul class="update-assets">';
    info.assets.forEach(function(a) {
      var sizeStr = a.size > 1048576
        ? (a.size / 1048576).toFixed(1) + ' MB'
        : a.size > 1024
          ? (a.size / 1024).toFixed(0) + ' KB'
          : a.size + ' B';
      assetsHtml += '<li class="update-asset" onclick="openUpdateAsset(\'' + a.download_url.replace(/'/g, "\\'") + '\')">' +
        '<span class="update-asset-name">📥 ' + a.name + '</span>' +
        '<span class="update-asset-size">' + sizeStr + '</span></li>';
    });
    assetsHtml += '</ul>';
  }
  document.getElementById('updateAssets').innerHTML = assetsHtml;

  // Open release page button
  var openBtn = document.getElementById('updateOpenBtn');
  if (openBtn) {
    openBtn.onclick = function() {
      if (info.release_url && isTauri && invoke) {
        invoke('open_url', { url: info.release_url }).catch(function(){});
      }
    };
  }

  modal.classList.remove('hidden');
}

function closeUpdateModal() {
  var modal = document.getElementById('updateModal');
  if (modal) modal.classList.add('hidden');
}

// Asset list click → open download URL in browser
function openUpdateAsset(url) {
  if (isTauri && invoke) {
    invoke('open_url', { url: url }).catch(function(){});
  }
}

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

  // Reset rename panel UI
  document.getElementById('summaryRenamePanel').classList.add('hidden');
  document.getElementById('summaryRenameBtn').classList.remove('active');
  _renamePreview = [];
  document.getElementById('srpSep').value = _renameSeparator || '_';
  document.getElementById('srpError').style.display = 'none';
  // Highlight the matching preset button, or clear all if custom
  document.querySelectorAll('.srp-preset').forEach(function(p) { p.classList.remove('active'); });
  document.querySelectorAll('.srp-preset').forEach(function(p) {
    var keys = p.getAttribute('onclick');
    if (keys) {
      var match = keys.match(/\[([^\]]+)\]/);
      if (match) {
        var presetKeys = match[1].replace(/'/g, '').split(',');
        if (presetKeys.length === _renameTemplate.length && presetKeys.every(function(k, i) { return k === _renameTemplate[i]; })) {
          p.classList.add('active');
        }
      }
    }
  });

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
      if (fileObj._xmlInvoice && fileObj.invoiceType) return fileObj.invoiceType;
      if (fileObj._isToll) return '通行费发票';
      if (fileObj._isTicket) return fileObj.sellerName || '车票'; // sellerName holds ticket label
      if (fileObj._ocrText && /非税/.test(fileObj._ocrText)) return '非税票据';
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

  // Auto-refresh rename preview if panel is open
  if (!document.getElementById('summaryRenamePanel').classList.contains('hidden')) {
    updateRenamePreview();
  }
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

  var csvContent = '\uFEFF' + rows.join('\r\n'); // UTF-8 BOM + CRLF for Excel

  if (isTauri && invoke) {
    try {
      var defaultDir = '';
      try { defaultDir = await invoke('get_downloads_dir'); } catch(e) {}
      var ts = new Date();
      var tsStr = ts.getFullYear() + String(ts.getMonth()+1).padStart(2,'0') + String(ts.getDate()).padStart(2,'0') + '_' + String(ts.getHours()).padStart(2,'0') + String(ts.getMinutes()).padStart(2,'0');
      var defaultName = '发票汇总表_' + tsStr + '.csv';
      var savePath = await invoke('plugin:dialog|save', {
        options: {
          title: '保存汇总表',
          defaultPath: defaultDir ? (defaultDir + (defaultDir.endsWith('\\')||defaultDir.endsWith('/')?'':'\\') + defaultName) : defaultName,
          filters: [{ name: 'CSV 文件', extensions: ['csv'] }]
        }
      });
      if (!savePath) return;
      await invoke('write_text_file', { path: savePath, content: csvContent });
      closeSummaryModal();
      // Open containing folder so user can find the file
      var dirPath = savePath.substring(0, Math.max(savePath.lastIndexOf('\\'), savePath.lastIndexOf('/')));
      try { await invoke('open_file', { path: dirPath }); } catch(e) {}
      toast('已保存: ' + savePath);
      // Update saveDir for future use
      if (dirPath) localStorage.setItem('ticketchan-save-dir', dirPath);
    } catch(e) {
      toast('导出失败: ' + e);
    }
  } else {
    // Browser fallback: download via Blob
    var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href = url; a.download = '发票汇总表.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    closeSummaryModal();
    toast('汇总表已导出');
  }
}

function csvEscape(val) {
  var s = String(val || '');
  if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

// =====================================================
// Batch File Rename (v2.0.5)
// =====================================================
var _renamePreview = [];

var RENAME_FIELDS = [
  { key: 'amountTax',       label: '含税金额'   },
  { key: 'amountNoTax',     label: '不含税金额' },
  { key: 'taxAmount',       label: '税额'       },
  { key: 'sellerName',      label: '销售方名称' },
  { key: 'sellerCreditCode',label: '销售方税号' },
  { key: 'buyerName',       label: '购买方名称' },
  { key: 'buyerCreditCode', label: '购买方税号' },
  { key: 'invoiceNo',       label: '发票号码'   },
  { key: 'invoiceDate',     label: '开票日期'   },
  { key: 'invoiceType',     label: '发票类型'   },
  { key: 'note',           label: '备注'       },
];

function toggleSummaryRename() {
  var panel = document.getElementById('summaryRenamePanel');
  var btn = document.getElementById('summaryRenameBtn');
  var isHidden = panel.classList.toggle('hidden');
  btn.classList.toggle('active', !isHidden);
  if (!isHidden) {
    renderRenameFields();
    updateRenamePreview();
    // Scroll panel into view
    setTimeout(function() { panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }, 50);
  }
}

function onRenamePresetClick(templateKeys, btn) {
  var presets = document.querySelectorAll('.srp-preset');
  presets.forEach(function(p) { p.classList.remove('active'); });
  if (btn) btn.classList.add('active');
  _renameTemplate = templateKeys;
  renderRenameFields();
  updateRenamePreview();
}

function renderRenameFields() {
  var html = '';
  RENAME_FIELDS.forEach(function(f) {
    var checked = _renameTemplate.indexOf(f.key) >= 0 ? ' checked' : '';
    html += '<label class="srp-field-item"><input type="checkbox" id="srpChk_' + f.key + '"' + checked + ' onchange="onRenameFieldToggle(\'' + f.key + '\')">' + escHtml(f.label) + '</label>';
  });
  html += '<span class="srp-field-actions"><a onclick="renameFieldsClear()">清除</a></span>';
  // Show current template order as hint
  var orderHint = _renameTemplate.map(function(key) {
    var f = RENAME_FIELDS.find(function(rf) { return rf.key === key; });
    return f ? f.label : key;
  }).join(' → ');
  html += '<div class="srp-order-hint">' + (orderHint || '请勾选字段') + '</div>';
  document.getElementById('srpFields').innerHTML = html;
}

function onRenameFieldToggle(key) {
  var cb = document.getElementById('srpChk_' + key);
  if (cb.checked) {
    // Append to end — later checked = later in filename
    if (_renameTemplate.indexOf(key) < 0) _renameTemplate.push(key);
  } else {
    _renameTemplate = _renameTemplate.filter(function(k) { return k !== key; });
  }
  renderRenameFields();
  updateRenamePreview();
}

function renameFieldsClear() {
  _renameTemplate = [];
  renderRenameFields();
  updateRenamePreview();
}

function sanitizeFileName(str) {
  if (!str) return '';
  var s = String(str);
  // Replace illegal characters for Windows filenames
  s = s.replace(/[\\/:*?"<>|]/g, '-');
  // Collapse repeated dots (path traversal safeguard)
  s = s.replace(/\.\.+/g, '.');
  // Remove leading/trailing whitespace and dots
  s = s.replace(/^[\s.]+/, '').replace(/[\s.]+$/, '');
  // Truncate to 200 chars (leaves room for extension + conflict suffix)
  if (s.length > 200) s = s.substring(0, 200);
  return s;
}

function buildNewFileName(fileObj) {
  if (!fileObj || !fileObj.name) return null;
  var parts = [];
  _renameTemplate.forEach(function(key) {
    var fieldDef = RENAME_FIELDS.find(function(f) { return f.key === key; }) ||
                   SUMMARY_FIELDS.find(function(f) { return f.key === key; });
    if (!fieldDef) return;
    var val = getSummaryCellValue(fileObj, fieldDef, 0);
    var clean = sanitizeFileName(val);
    if (clean) parts.push(clean);
  });
  if (parts.length === 0) return null;
  var newBase = parts.join(_renameSeparator);
  if (!newBase) return null;
  var extMatch = fileObj.name.match(/\.([^.]+)$/i);
  var ext = extMatch ? '.' + extMatch[1].toLowerCase() : '';
  return newBase + ext;
}

function updateRenamePreview() {
  _renameSeparator = document.getElementById('srpSep').value || '_';
  var files = getCheckedFiles();
  var preview = [];
  var okCount = 0, warnCount = 0, skipCount = 0;
  var seenPaths = {}; // dedup by source path — same PDF file has multiple pages

  files.forEach(function(fileObj) {
    // Use _filePath (for images/OFD) or _pdfPath (for PDF pages) as rename source
    var srcPath = fileObj._filePath || fileObj._pdfPath || '';
    if (!srcPath) {
      preview.push({ fileObj: fileObj, oldName: fileObj.name, newName: null, status: 'skip', reason: '无文件路径' });
      skipCount++;
      return;
    }
    // Same source file already processed (multi-page PDF)? Skip subsequent pages
    if (seenPaths[srcPath]) {
      preview.push({ fileObj: fileObj, oldName: fileObj.name, newName: null, status: 'skip', reason: '同文件已处理' });
      skipCount++;
      return;
    }
    var newName = buildNewFileName(fileObj);
    if (!newName) {
      preview.push({ fileObj: fileObj, oldName: fileObj.name, newName: null, status: 'skip', reason: '字段为空' });
      skipCount++;
      return;
    }
    if (newName === fileObj.name) {
      preview.push({ fileObj: fileObj, oldName: fileObj.name, newName: newName, status: 'ok', reason: '已匹配' });
      okCount++;
      seenPaths[srcPath] = true;
      return;
    }
    // Check for conflicts among preview entries
    var conflict = preview.find(function(p) { return p.newName === newName && p.status === 'ok'; });
    if (conflict) {
      warnCount++;
      preview.push({ fileObj: fileObj, oldName: fileObj.name, newName: newName, status: 'conflict' });
      seenPaths[srcPath] = true;
      return;
    }
    preview.push({ fileObj: fileObj, oldName: fileObj.name, newName: newName, status: 'ok' });
    okCount++;
    seenPaths[srcPath] = true;
  });

  // Resolve conflicts with sequence numbers
  resolveNameConflicts(preview);

  _renamePreview = preview;

  // Render preview table
  var hasNote = _renameTemplate.indexOf('note') >= 0;
  var html = '<table class="srp-preview-table"><thead><tr><th class="srp-status"></th><th>原文件名</th>'
    + (hasNote ? '<th class="srp-note-col">备注</th>' : '')
    + '<th>新文件名</th></tr></thead><tbody>';
  var execCount = 0;
  preview.forEach(function(p, pIdx) {
    var statusIcon = '', statusCls = '';
    switch (p.status) {
      case 'ok':     statusIcon = (p.reason === '已匹配' ? '✓' : '→'); statusCls = p.reason === '已匹配' ? 'srp-status-skip' : 'srp-status-ok'; break;
      case 'conflict': statusIcon = '⚠'; statusCls = 'srp-status-warn'; break;
      case 'skip':   statusIcon = '✗'; statusCls = 'srp-status-error'; break;
    }
    var noteCell = '';
    if (hasNote) {
      var noteVal = p.fileObj.note || '';
      noteCell = '<td class="srp-note-col"><input type="text" value="' + escHtml(noteVal) + '" class="srp-note-input" data-idx="' + pIdx + '" placeholder="备注" oninput="onRenameNoteInput(this)"></td>';
    }
    html += '<tr><td class="srp-status ' + statusCls + '">' + statusIcon + '</td><td>' + escHtml(p.oldName) + '</td>'
      + noteCell
      + '<td>' + escHtml(p.newName || '— 跳过 —') + '</td></tr>';
    if (p.status === 'ok' && p.reason !== '已匹配') execCount++;
    if (p.status === 'conflict') execCount++;
  });
  html += '</tbody></table>';

  // If no files or all files are skipped, show guide tip
  if (preview.length === 0) {
    html += '<div class="srp-guide">没有勾选的发票。请先在文件列表中勾选需要重命名的发票</div>';
  } else if (okCount === 0 && execCount === 0) {
    var allNoPath = preview.every(function(p) { return p.reason === '无文件路径'; });
    if (allNoPath) {
      html += '<div class="srp-guide">未找到文件路径，请通过「拖入文件」方式加载发票</div>';
    } else {
      html += '<div class="srp-guide">暂无可用字段。请先在汇总表中核对金额和销售方，编辑后预览自动刷新</div>';
    }
  }
  document.getElementById('srpPreview').innerHTML = html;

  var execBtn = document.getElementById('srpExecBtn');
  execBtn.textContent = '执行重命名 (' + execCount + ')';
  execBtn.disabled = execCount === 0;

  // Hide error div
  document.getElementById('srpError').style.display = 'none';
}

var _renameNoteTimer = 0;
var _renameNoteDirty = false;
function onRenameNoteInput(input) {
  var idx = parseInt(input.dataset.idx);
  var pv = _renamePreview[idx];
  if (!pv || !pv.fileObj) return;
  pv.fileObj.note = input.value;
  _renameNoteDirty = true;
  clearTimeout(_renameNoteTimer);
  _renameNoteTimer = setTimeout(function() {
    updateRenamePreview();
    if (!document.getElementById('summaryModal').classList.contains('hidden')) {
      renderSummaryTable();
    }
    _renameNoteDirty = false;
  }, 300);
}

function resolveNameConflicts(preview) {
  var seen = {};
  // First pass: mark all existing names (from non-skip entries)
  preview.forEach(function(p) {
    if (p.newName && p.status !== 'skip') {
      seen[p.newName] = (seen[p.newName] || 0) + 1;
    }
  });
  // Second pass: for names that appear >1 times, add _2, _3 suffixes
  var counter = {};
  preview.forEach(function(p) {
    if (p.status === 'skip' || !p.newName) return;
    if (seen[p.newName] <= 1) return;
    counter[p.newName] = (counter[p.newName] || 0) + 1;
    if (counter[p.newName] > 1) {
      var extMatch = p.newName.match(/\.([^.]+)$/i);
      var base = extMatch ? p.newName.substring(0, p.newName.length - extMatch[0].length) : p.newName;
      var ext = extMatch ? extMatch[0] : '';
      p.newName = base + '_' + counter[p.newName] + ext;
      p.status = 'conflict';
    }
  });
}

async function executeRename() {
  if (_renameNoteDirty) { updateRenamePreview(); _renameNoteDirty = false; }
  var execList = _renamePreview.filter(function(p) {
    return (p.status === 'ok' && p.reason !== '已匹配') || p.status === 'conflict';
  });
  if (!execList.length) { toast('没有需要重命名的文件'); return; }

  var execBtn = document.getElementById('srpExecBtn');
  execBtn.disabled = true;
  execBtn.textContent = '重命名中...';

  var successCount = 0, failCount = 0;
  var errors = [];

  for (var i = 0; i < execList.length; i++) {
    var p = execList[i];
    // Use _filePath (images/OFD) or _pdfPath (PDF pages) as rename source
    var srcPath = p.fileObj._filePath || p.fileObj._pdfPath || '';
    var srcDir = srcPath.substring(0, Math.max(srcPath.lastIndexOf('\\'), srcPath.lastIndexOf('/')));
    var newPath = srcDir + (srcDir.endsWith('\\') || srcDir.endsWith('/') ? '' : '\\') + p.newName;

    try {
      var isBrowserMode = !isTauri || !invoke;
      if (!isBrowserMode) {
        await invoke('rename_file', { srcPath: srcPath, destPath: newPath });
      } else {
        // Browser testing fallback — only update display name, not file paths
        console.log('[rename]', p.oldName, '→', p.newName);
      }

      // Success — update all references
      var oldName = p.fileObj.name;
      p.fileObj.name = p.newName;
      if (!isBrowserMode) {
        if (p.fileObj._filePath) p.fileObj._filePath = newPath;
        // Sync _pdfPath for all pages sharing the same source (multi-page PDF)
        var oldPdf = srcPath;
        S.files.forEach(function(f) {
          if (f._pdfPath === oldPdf) f._pdfPath = newPath;
          if (f._filePath === oldPdf && f !== p.fileObj) f._filePath = newPath;
        });
      }

      // Migrate _fileAdjMap (per-file slot adjustments)
      if (S._fileAdjMap && S._fileAdjMap[oldName]) {
        S._fileAdjMap[p.newName] = S._fileAdjMap[oldName];
        delete S._fileAdjMap[oldName];
      }
      // Migrate _notesMap (per-file notes)
      if (S._notesMap && S._notesMap[oldName] !== undefined) {
        S._notesMap[p.newName] = S._notesMap[oldName];
        delete S._notesMap[oldName];
      }
      // Migrate _printedMap (per-file printed state, keyed by file path)
      if (!isBrowserMode && _printedMap[srcPath] !== undefined) {
        _printedMap[newPath] = _printedMap[srcPath];
        delete _printedMap[srcPath];
      }

      successCount++;
    } catch(e) {
      failCount++;
      errors.push({ oldName: p.oldName, error: String(e) });
    }
  }

  // Refresh UI
  renderFileList();
  renderSummaryTable();
  saveSettings();

  // Report result
  var msg = '重命名完成：成功 ' + successCount + ' 个';
  if (failCount > 0) msg += '，失败 ' + failCount + ' 个';
  toast(msg, failCount > 0 ? 5000 : 3000);

  if (errors.length > 0) {
    var errDiv = document.getElementById('srpError');
    errDiv.style.display = 'block';
    errDiv.innerHTML = '<strong>失败详情：</strong><br>' + errors.map(function(e) { return escHtml(e.oldName) + ': ' + escHtml(e.error); }).join('<br>');
  }

  // Update preview to reflect new state
  updateRenamePreview();
}


function bindFooterTextEvent() {
  var el = document.getElementById('footerText');
  if (el) el.addEventListener('input', function() { updatePreview(); });
}
