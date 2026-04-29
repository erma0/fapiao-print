// =====================================================
// Print & PDF Functions (Rust layout backend)
// =====================================================
// Dependencies (global): isTauri, invoke, S, getSettings, getActiveFiles, buildPages, showLoading, hideLoading, toast, getSaveDir, setSaveDir, escHtml, calculateLayout

/**
 * Build a LayoutRenderRequest for the new Rust backend.
 * Replaces the old approach of renderPageToCanvas + generate_and_print/save_pdf.
 */
function buildLayoutRequest(files, settings) {
  // 1. Collect unique file specs
  var fileMap = {};
  var fileSpecs = [];

  function getFileIndex(fileObj) {
    if (!fileObj) return null;
    var key = fileObj.previewUrl || '';
    if (!key) return null;
    if (!fileMap[key]) {
      fileMap[key] = fileSpecs.length;
      fileSpecs.push({
        dataUrl: key,
        ow: fileObj.ow || 0,
        oh: fileObj.oh || 0,
        rotation: fileObj.rotation || 0,
      });
    }
    return fileMap[key];
  }

  // 2. Build expanded pages (with colate/copies handling)
  var pages = buildPages(files, settings);
  var copies = settings.copies || 1;
  var expanded = settings.collate
    ? Array(copies).fill(pages).flat()
    : pages.flatMap(function(p) { return Array(copies).fill(p); });

  // 3. Build page specs with effective rotation
  var pageSpecs = [];
  // Pre-calculate layout so we know slot dimensions
  var layout = calculateLayout(settings);

  for (var i = 0; i < expanded.length; i++) {
    var slots = [];
    var pageFiles = expanded[i];
    for (var j = 0; j < pageFiles.length; j++) {
      var f = pageFiles[j];
      if (f) {
        var rot = getEffectiveRotation(f, j, settings, layout);
        slots.push({ fileIndex: getFileIndex(f), rotation: rot });
      } else {
        slots.push({ fileIndex: null, rotation: 0 });
      }
    }
    pageSpecs.push({ slots: slots });
  }

  return { files: fileSpecs, pages: pageSpecs, settings: settings };
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
 * Print invoices — Rust does layout + PDF generation.
 */
async function doPrint() {
  var files = getActiveFiles();
  if (!files.length) { toast('请先添加发票！'); return; }

  showLoading('正在准备打印...');
  try {
    var s = getSettings();
    var layoutReq = buildLayoutRequest(files, s);

    if (isTauri && invoke) {
      document.getElementById('loadingText').textContent = '正在生成PDF，请稍候...';
      var printMode = document.getElementById('printMode').value;
      var result = await invoke('generate_pdf_from_layout', {
        request: layoutReq,
        outputPath: 'C:\\temp\\fapiao_print_output.pdf',
        directPrint: printMode === 'direct',
        printerName: s.printerName || null
      });
      hideLoading();
      if (result.success) {
        toast('\uD83D\uDCA8 ' + result.message);
      } else {
        toast('打印失败：' + result.message);
      }
    } else {
      hideLoading();
      fallbackPrint(files, s);
    }
  } catch (err) {
    hideLoading();
    console.error('Print error:', err);
    toast('打印出错：' + String(err));
  }
}

/**
 * Save invoices as PDF file — Rust does layout + PDF generation.
 */
async function savePdf() {
  var files = getActiveFiles();
  if (!files.length) { toast('请先添加发票！'); return; }

  var savePath = null;
  var savedDir = getSaveDir();
  if (isTauri && invoke) {
    try {
      var now = new Date();
      var ts = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0') + '_' + String(now.getHours()).padStart(2, '0') + String(now.getMinutes()).padStart(2, '0') + String(now.getSeconds()).padStart(2, '0');
      var defaultName = '发票打印_' + ts + '.pdf';
      if (savedDir) {
        savePath = savedDir + (savedDir.endsWith('\\') || savedDir.endsWith('/') ? '' : '\\') + defaultName;
      } else {
        savePath = await invoke('plugin:dialog|save', {
          options: {
            title: '保存发票PDF',
            defaultPath: defaultName,
            filters: [{ name: 'PDF文件', extensions: ['pdf'] }]
          }
        });
        if (!savePath) return;
        var lastSep = Math.max(savePath.lastIndexOf('\\'), savePath.lastIndexOf('/'));
        var dir = lastSep >= 0 ? savePath.substring(0, lastSep) : '';
        if (dir) setSaveDir(dir);
      }
    } catch(e) { savePath = null; }
  }

  showLoading('正在准备保存...');
  try {
    var s = getSettings();
    var layoutReq = buildLayoutRequest(files, s);

    if (isTauri && invoke) {
      document.getElementById('loadingText').textContent = '正在生成PDF...';
      var result = await invoke('generate_pdf_from_layout', {
        request: layoutReq,
        outputPath: savePath,
        directPrint: false,
        printerName: null
      });
      hideLoading();
      if (result.success) {
        toast('\u2705 PDF已保存: ' + result.pdfPath);
        // Auto-open
        if (S.feat.autoOpenPdf && result.pdfPath) {
          try { invoke('open_url', { url: 'file:///' + result.pdfPath.replace(/\\/g, '/') }); } catch(e) {}
        }
      } else {
        toast('PDF生成失败：' + result.message);
      }
    } else {
      hideLoading();
      fallbackPrint(files, s);
    }
  } catch (err) {
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
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>发票打印</title><style>*{margin:0;padding:0;box-sizing:border-box}@page{size:' + s.paperW + 'mm ' + s.paperH + 'mm;margin:0}body{background:white}.page{width:' + s.paperW + 'mm;height:' + s.paperH + 'mm;position:relative;page-break-after:always;background:white;overflow:hidden}.slot{position:absolute;overflow:hidden;display:flex;align-items:center;justify-content:center}.slot img{max-width:100%;max-height:100%;object-fit:contain}</style></head><body>';
  expanded.forEach(function(page) {
    html += '<div class="page">';
    var mt = s.marginTop, mb = s.marginBottom, ml = s.marginLeft, mr = s.marginRight;
    var slotW = (s.paperW - s.cols * (ml + mr) - (s.cols - 1) * s.gapH) / s.cols;
    var slotH = (s.paperH - s.rows * (mt + mb) - (s.rows - 1) * s.gapV) / s.rows;
    for (var r = 0; r < s.rows; r++) for (var c = 0; c < s.cols; c++) {
      var f = page[r * s.cols + c];
      var x = ml + c * (slotW + ml + mr + s.gapH), y = mt + r * (slotH + mt + mb + s.gapV);
      if (f && f.previewUrl) {
        var src = S.feat.trimWhite && f.trimmedUrl ? f.trimmedUrl : f.previewUrl;
        html += '<div class="slot" style="left:' + x + 'mm;top:' + y + 'mm;width:' + slotW + 'mm;height:' + slotH + 'mm"><img src="' + src + '"></div>';
      }
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
  if (!isTauri || !invoke) { toast('仅在桌面模式下可用'); return; }
  try {
    var printers = await invoke('get_printers');
    var sel = document.getElementById('printerSel');
    sel.innerHTML = '<option value="">默认打印机</option>';
    printers.forEach(function(p) {
      sel.innerHTML += '<option value="' + escHtml(p.name) + '" ' + (p.isDefault ? 'selected' : '') + '>' + escHtml(p.name) + (p.isDefault ? ' (默认)' : '') + '</option>';
    });
    toast('已刷新打印机列表');
  } catch(e) { toast('获取打印机列表失败'); }
}
