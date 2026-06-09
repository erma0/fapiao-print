// =====================================================
// Print & PDF Functions (layout → server-side PDF generation)
// =====================================================
// Dependencies (global): __api, S, getSettings, getActiveFiles, buildPages, showLoading, hideLoading, toast, escHtml, calculateLayout

// 智能缓存：保存上一次的 LayoutRequest 和 PDF 路径
var _lastLayoutRequest = null;
var _lastPdfPath = null;

/**
 * 深度比较两个对象是否相等
 */
function deepEqual(a, b) {
  if (a === b) return true;
  if (typeof a !== typeof b) return false;
  if (typeof a !== 'object' || a === null || b === null) return false;
  
  var keysA = Object.keys(a);
  var keysB = Object.keys(b);
  
  if (keysA.length !== keysB.length) return false;
  
  for (var ki = 0; ki < keysA.length; ki++) {
    var key = keysA[ki];
    if (keysB.indexOf(key) === -1) return false;
    if (!deepEqual(a[key], b[key])) return false;
  }
  
  return true;
}

/**
 * 检查是否可以使用缓存（比较当前和上次的 LayoutRequest）
 */
function canUseCachedPdf(currentRequest) {
  if (!_lastLayoutRequest || !_lastPdfPath) return false;
  return deepEqual(currentRequest, _lastLayoutRequest);
}

/**
 * 更新缓存（生成 PDF 后调用）
 */
function updatePdfCache(request, pdfPath) {
  _lastLayoutRequest = request;
  _lastPdfPath = pdfPath;
}

/**
 * Build a LayoutRenderRequest for the new Rust backend.
 * Replaces the old approach of renderPageToCanvas + generate_and_print/save_pdf.
 *
 * **Optimization**: When a file has a disk path (_filePath), we pass the path
 * instead of the base64 dataUrl. Rust reads the file directly from disk,
 * skipping the expensive base64 encode→IPC→decode round-trip.
 * For rendered PDF pages and OFD images (no disk path), dataUrl is used.
 */
function buildLayoutRequest(files, settings) {
  // 1. Collect unique file specs
  var fileMap = {};
  var fileSpecs = [];

  function getFileIndex(fileObj) {
    if (!fileObj) return null;
    // XML 数电票 has no visual layout — skip from print layout
    if (fileObj._xmlInvoice) return null;
    // Use _filePath as dedup key when available (more stable than previewUrl).
    // For OFD, fall back to previewUrl since _filePath is shared across pages.
    var key = (fileObj.type !== 'ofd' && fileObj._filePath) ? fileObj._filePath : (fileObj.previewUrl || '');
    if (!key) return null;
    if (!(key in fileMap)) {
      fileMap[key] = fileSpecs.length;
      var spec = {
        ow: fileObj.ow || 0,
        oh: fileObj.oh || 0,
        rotation: fileObj.rotation || 0,
      };
      // Determine source type and path strategy
      // PDF pages have _pdfPath; OFD pages have type='ofd' even with _filePath
      if (fileObj._pdfPath) {
        spec.sourceType = 'pdf-page';
        spec.dataUrl = fileObj.previewUrl || '';
        spec.filePath = null;
      } else if (fileObj.type === 'ofd') {
        spec.sourceType = 'ofd-page';
        spec.dataUrl = fileObj.previewUrl || '';
        spec.filePath = null;
      } else if (fileObj._filePath) {
        spec.filePath = fileObj._filePath;
        spec.dataUrl = ''; // not needed — Rust reads from file
        spec.sourceType = 'image';
      } else {
        spec.dataUrl = fileObj.previewUrl || '';
        spec.filePath = null;
        spec.sourceType = 'image';
      }
      // Pass PDF source info for passthrough optimization
      if (fileObj._pdfPath) {
        spec.pdfPath = fileObj._pdfPath;
        spec.pdfPageIdx = fileObj._pdfPageIdx >= 0 ? fileObj._pdfPageIdx : null;
      }
      fileSpecs.push(spec);
    }
    return fileMap[key];
  }

  // 2. Build pages (per-file copies already expanded in getActiveFiles,
  // global copies handled by print command copies parameter, not expanded here)
  var pages = buildPages(files, settings);
  var expanded = pages;

  // 3. Build page specs with effective rotation
  var pageSpecs = [];
  // Pre-calculate layout so we know slot dimensions
  var layout = calculateLayout(settings);

  var perPage = settings.cols * settings.rows;
  for (var i = 0; i < expanded.length; i++) {
    var slots = [];
    var pageFiles = expanded[i];
    for (var j = 0; j < perPage; j++) {
      var f = j < pageFiles.length ? pageFiles[j] : null;
      if (f) {
        var rot = getEffectiveRotation(f, j, settings, layout);
        var slotSpec = { fileIndex: getFileIndex(f), rotation: rot };
        // Per-slot adjustment: pass scale/offset if non-default
        if (f.slotScale && f.slotScale !== 1) slotSpec.scale = f.slotScale;
        if (f.slotOffsetX) slotSpec.offsetX = f.slotOffsetX;
        if (f.slotOffsetY) slotSpec.offsetY = f.slotOffsetY;
        slots.push(slotSpec);
      } else {
        slots.push({ fileIndex: null, rotation: 0 });
      }
    }
    pageSpecs.push({ slots: slots });
  }

  // 排除纯打印机参数，避免切换打印机/份数等操作错误地导致缓存失效
  // 这些字段仅影响打印方式，不影响 PDF 内容布局
  var _cacheExclude = { printerName:1, copies:1, duplex:1, collate:1 };
  var layoutSettings = {};
  for (var k in settings) {
    if (!(k in _cacheExclude)) layoutSettings[k] = settings[k];
  }
  return { files: fileSpecs, pages: pageSpecs, settings: layoutSettings };
}

/**
 * Compute effective rotation for a file in a slot.
 * Mirrors the logic from layout.js getRotation().
 */
function getEffectiveRotation(fileObj, slotIdx, settings, layout) {
  var slot = layout.slots[slotIdx];
  if (settings.globalRotation === 'auto') {
    var isSlotL = slot.w > slot.h;
    var isImgL = (fileObj.ow || 1) > (fileObj.oh || 1);
    return (isSlotL !== isImgL) ? (fileObj.rotation + 90) % 360 : fileObj.rotation;
  }
  return ((parseInt(settings.globalRotation) || 0) + (fileObj.rotation || 0)) % 360;
}

/**
 * Listen for PDF generation progress via SSE.
 * Returns { taskId, close } — taskId is sent with the generate_pdf request,
 * close() shuts down the SSE connection.
 */
function listenPdfProgressSse() {
  if (!__api) return null;
  var taskId = __api.generateTaskId();
  var listener = null;
  try {
    listener = __api.listen(taskId, function(evt) {
      if (evt && evt.current !== undefined && evt.total !== undefined) {
        updateLoadingProgress(evt.phase || '', evt.current, evt.total);
      }
    });
  } catch(e) {
    console.warn('SSE progress listen failed:', e);
    return null;
  }
  return { taskId: taskId, close: function() { if (listener) listener.close(); } };
}

/**
 * Print invoices — browser print in new window.
 */
async function doPrint() {
  var files = getActiveFiles();
  if (!files.length) { toast('请先添加发票！'); return; }
  var s = getSettings();
  var printMode = document.getElementById('printMode').value;

  if (printMode === 'confirm') {
    showPrintConfirm(files, s);
  } else if (printMode === 'pdfium') {
    await doPdfiumPrint(files, s);
  } else if (printMode === 'pdf') {
    await doPdfReaderPrint(files, s);
  } else {
    toast('此打印模式已移除');
  }
}

function showPrintConfirm(files, s) {
  var printerName = s.printerName || '默认打印机';
  var layout = S.layout.rows + '\u00D7' + S.layout.cols;
  var ps = document.getElementById('paperSize').value;
  var orient = document.getElementById('orientation').value === 'portrait' ? '纵向' : '横向';
  var paper = ps === 'custom' ? (s.paperW + '\u00D7' + s.paperH + 'mm') : ps.toUpperCase();
  var copies = s.copies || 1;
  var colorMode = document.getElementById('colorMode').value;
  var colorLabel = colorMode === 'color' ? '彩色' : colorMode === 'grayscale' ? '灰度' : '黑白';
  var activeCount = files.length;

  var pages = buildPages(files, s);
  var totalPages = pages.length;

  var fitLabel = s.fitMode === 'contain' ? '适应' : s.fitMode === 'cover' ? '填充' : (Math.round(s.customScale * 100) + '%');
  var rotLabel = s.globalRotation === 'auto' ? '自动' : s.globalRotation + '\u00B0';

  var row = function(lbl, val, cls) {
    return '<div class="modal-row compact' + (cls ? ' ' + cls : '') + '"><span class="modal-lbl">' + lbl + '</span><span class="modal-val">' + val + '</span></div>';
  };

  var engineSelect = '<select id="confirmPrintEngine" style="padding:1px 4px;border-radius:4px;border:1px solid var(--border);background:var(--bg-primary);color:var(--text-primary);font-size:12px">'
    + '<option value="pdfium">PDFium（推荐）</option>'
    + '</select>';

  var html = row('打印机', escHtml(printerName))
    + row('引擎', engineSelect)
    + row('发票', activeCount + ' 张 \u00D7 ' + copies + ' 份 \u2192 ' + totalPages + ' 页', 'highlight')
    + row('布局', layout + ' \u00B7 ' + paper + orient + ' \u00B7 ' + fitLabel + ' \u00B7 ' + rotLabel)
    + row('边距', s.marginTop + '/' + s.marginBottom + '/' + s.marginLeft + '/' + s.marginRight + 'mm \u00B7 间距 ' + s.gapH + '/' + s.gapV + 'mm')
    + row('选项', colorLabel + ' \u00B7 ' + (s.duplex ? '双面' : '单面') + ' \u00B7 ' + (s.cutline ? '切割线' : '无切割线'));

  document.getElementById('printConfirmBody').innerHTML = html;
  document.getElementById('printConfirmModal').classList.remove('hidden');
}

function closePrintConfirm() {
  document.getElementById('printConfirmModal').classList.add('hidden');
}

async function confirmPrint() {
  closePrintConfirm();
  var files = getActiveFiles();
  var s = getSettings();
  await doPdfiumPrint(files, s);
}

async function doPdfiumPrint(files, s) {
  fallbackPrint(files, s);
}

async function doPdfReaderPrint(files, s) {
  fallbackPrint(files, s);
}

/**
 * Save invoices as PDF file — Rust does layout + PDF generation.
 */
async function savePdf() {
  var files = getActiveFiles();
  if (!files.length) { toast('请先添加发票！'); return; }

  var s = getSettings();
  var currentRequest = buildLayoutRequest(files, s);

  showLoading('正在准备保存...');
  var progress = listenPdfProgressSse();
  try {
    document.getElementById('loadingText').textContent = '正在生成PDF...';
    var result = await __api.call('generate_pdf_from_layout', {
      request: currentRequest,
      progressTaskId: progress ? progress.taskId : '',
      outputPath: '',
      directPrint: false,
      printerName: null,
      printAfter: false
    });
    if (progress) progress.close();
    hideLoading();

    if (result && result.taskId) {
      __api.downloadFile(__api.getSessionId(), result.fileName || (result.taskId + '.pdf'));
    } else {
      fallbackPrint(files, s);
    }
  } catch (err) {
    if (progress) progress.close();
    hideLoading();
    console.error('PDF error:', err);
    toast('PDF生成出错：' + String(err));
  }
}

/**
 * Browser fallback: open print dialog in new window
 */
function fallbackPrint(files, s) {
  var w = window.open('', '_blank');
  if (!w) { alert('弹出窗口被阻止'); return; }
  var pages = buildPages(files, s);
  var expanded = s.collate ? Array(s.copies).fill(pages).flat() : pages.flatMap(function(p) { return Array(s.copies).fill(p); });
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>发票酱</title><style>*{margin:0;padding:0;box-sizing:border-box}@page{size:' + s.paperW + 'mm ' + s.paperH + 'mm;margin:0}body{background:white}.page{width:' + s.paperW + 'mm;height:' + s.paperH + 'mm;position:relative;page-break-after:always;background:white;overflow:hidden}.slot{position:absolute;overflow:hidden;display:flex;align-items:center;justify-content:center}.slot img{max-width:100%;max-height:100%;object-fit:contain}';
  // Add cut line styles if enabled
  if (s.cutline) {
    html += '.cutline-v{position:absolute;top:0;bottom:0;width:0;border-right:1px dashed #ccc}.cutline-h{position:absolute;left:0;right:0;height:0;border-bottom:1px dashed #ccc}';
  }
  html += '</style></head><body>';
  expanded.forEach(function(page, pi) {
    html += '<div class="page">';
    var mt = s.marginTop, mb = s.marginBottom, ml = s.marginLeft, mr = s.marginRight;
    var fm = s.footerMargin || 0;
    var slotW = (s.paperW - s.cols * (ml + mr) - (s.cols - 1) * s.gapH) / s.cols;
    var slotH = (s.paperH - s.rows * (mt + mb) - (s.rows - 1) * s.gapV - fm) / s.rows;
    // Draw cut lines (vertical + horizontal) between slots
    if (s.cutline && (s.cols > 1 || s.rows > 1)) {
      var hasFb = s.pageNum || s.printDate || (s.footerText || '').trim();
      var vLineH = hasFb ? (s.paperH - fm) : s.paperH;
      for (var c = 1; c < s.cols; c++) {
        var x = ml + c * (slotW + ml + mr + s.gapH) - s.gapH / 2;
        html += '<div class="cutline-v" style="left:' + x + 'mm;height:' + vLineH + 'mm"></div>';
      }
      for (var r = 1; r < s.rows; r++) {
        var y = mt + r * (slotH + mt + mb + s.gapV) - s.gapV / 2;
        html += '<div class="cutline-h" style="top:' + y + 'mm"></div>';
      }
    }
    for (var r = 0; r < s.rows; r++) for (var c = 0; c < s.cols; c++) {
      var f = page[r * s.cols + c];
      var x = ml + c * (slotW + ml + mr + s.gapH), y = mt + r * (slotH + mt + mb + s.gapV);
      if (f && f.previewUrl) {
        var src = S.feat.trimWhite && f.trimmedUrl ? f.trimmedUrl : f.previewUrl;
        // Compute effective rotation (same logic as layout.js getRotation)
        var rot = 0;
        var slot = { w: slotW, h: slotH };
        if (s.globalRotation === 'auto') {
          var isSlotL = slotW > slotH;
          var isImgL = (f.ow || 1) > (f.oh || 1);
          rot = (isSlotL !== isImgL) ? ((f.rotation || 0) + 90) % 360 : (f.rotation || 0);
        } else {
          rot = ((parseInt(s.globalRotation) || 0) + (f.rotation || 0)) % 360;
        }
        // For 90°/270° rotation, swap max-width/max-height constraints (same as preview fix)
        var isRotated90 = (rot === 90 || rot === 270);
        var sizeStyle = isRotated90
          ? 'max-width:' + slotH + 'mm;max-height:' + slotW + 'mm;'
          : 'max-width:100%;max-height:100%;';
        // Per-slot adjustment: scale & offset combined with rotation
        var perScale = f.slotScale || 1;
        var perOffX = f.slotOffsetX || 0;
        var perOffY = f.slotOffsetY || 0;
        var transforms = '';
        if (perOffX !== 0 || perOffY !== 0) transforms += 'translate(' + perOffX + 'mm,' + perOffY + 'mm) ';
        if (perScale !== 1) transforms += 'scale(' + perScale + ') ';
        if (rot) transforms += 'rotate(' + rot + 'deg) ';
        var transformStyle = transforms ? 'transform:' + transforms + ';' : '';
        html += '<div class="slot" style="left:' + x + 'mm;top:' + y + 'mm;width:' + slotW + 'mm;height:' + slotH + 'mm"><img src="' + escHtml(src) + '" style="' + sizeStyle + transformStyle + '"></div>';
      }
    }
    // 文本位置：距页面底部 5mm 处（在页边距区域内）
    var textBottomMm = 5;
    var lineHeightMm = 7; // 行高 ~7mm，用于多行文本垂直偏移

    // Page number and print date (per-page, same as preview)
    if (s.pageNum) html += '<div style="position:absolute;bottom:' + textBottomMm + 'mm;left:0;right:0;text-align:center;font-size:10px;color:#94a3b8">第 ' + (pi + 1) + ' 页 / 共 ' + expanded.length + ' 页</div>';
    if (s.printDate) {
      var now = new Date();
      var dateStr = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
      var dateBottomMm = textBottomMm;
      var dateStyle = s.pageNum ? 'position:absolute;bottom:' + dateBottomMm + 'mm;right:10mm;font-size:10px;color:#94a3b8' : 'position:absolute;bottom:' + dateBottomMm + 'mm;left:0;right:0;text-align:center;font-size:10px;color:#94a3b8';
      html += '<div style="' + dateStyle + '">打印日期 ' + dateStr + '</div>';
    }
    // Footer text (per-page, custom text)
    if (s.footerText) {
      var footerBottomMm = textBottomMm;
      if (s.pageNum || s.printDate) footerBottomMm += lineHeightMm;
      html += '<div style="position:absolute;bottom:' + footerBottomMm + 'mm;left:0;right:0;text-align:center;font-size:10px;color:#94a3b8">' + escHtml(s.footerText) + '</div>';
    }
    html += '</div>';
  });
  html += '</body></html>';
  w.document.write(html);
  w.document.close();
  w.onload = function() { setTimeout(function() { w.print(); }, 500); };
}

/**
 * Refresh printer list from system
 */
async function refreshPrinters() {
  toast('Web 模式不支持打印机选择');
}
