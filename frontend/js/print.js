// =====================================================
// Print & PDF Functions — pure client-side (pdf-lib + iframe.print)
// =====================================================
// No backend. All PDF generation runs in the browser via pdf-lib.
// Print uses an embedded <iframe> to invoke the browser's native PDF print.

var _printCacheKey = null;
var _printCacheBlob = null;

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

async function _embedForFile(pdfDoc, fileObj) {
  if (!fileObj) return null;
  if (fileObj._xmlInvoice) return null;
  if (fileObj.previewUrl) {
    var isJpeg = fileObj.previewUrl.indexOf('data:image/jpeg') === 0
      || fileObj.previewUrl.indexOf('data:image/jpg') === 0;
    var bytes = isJpeg ? _jpegBlobFromDataUrl(fileObj.previewUrl) : _pngBlobFromDataUrl(fileObj.previewUrl);
    return isJpeg ? await pdfDoc.embedJpg(bytes) : await pdfDoc.embedPng(bytes);
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
    var img = await _embedForFile(pdfDoc, f);
    if (!img) continue;

    var rot = getRotation(f, slot, settings);
    var perScale = f.slotScale || 1;
    var perOffX = f.slotOffsetX || 0;
    var perOffY = f.slotOffsetY || 0;

    var imgObjW = f.ow || img.width;
    var imgObjH = f.oh || img.height;
    var fitScale = Math.min(slot.w / imgObjW, slot.h / imgObjH);
    if (settings.fitMode === 'fill') fitScale = Math.max(slot.w / imgObjW, slot.h / imgObjH);
    else if (settings.fitMode === 'original') fitScale = 1;
    if (settings.fitMode === 'custom' && settings.customScale) fitScale *= settings.customScale;
    fitScale *= perScale;

    var drawW = imgObjW * fitScale;
    var drawH = imgObjH * fitScale;
    var cx = slot.x + slot.w / 2 + perOffX;
    var cy = ph - (slot.y + slot.h / 2 + perOffY);

    page.drawImage(img, {
      x: cx - drawW / 2,
      y: cy - drawH / 2,
      width: drawW,
      height: drawH,
      rotate: pdfLib.degrees(rot),
      opacity: settings.colorMode === 'bw' ? 1 : 1
    });

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

  if (settings.watermark && settings.wmText) {
    var wmSize = settings.wmSize || 60;
    var wmOpacity = settings.wmOpacity != null ? settings.wmOpacity : 0.15;
    page.drawText(settings.wmText, {
      x: pw / 2 - settings.wmText.length * wmSize * 0.15,
      y: ph / 2,
      size: wmSize,
      font: pdfLib.StandardFonts.HelveticaBold,
      color: pdfLib.rgb(0.6, 0.6, 0.6),
      opacity: wmOpacity,
      rotate: pdfLib.degrees(30)
    });
  }

  if (settings.pageNum || settings.printDate || (settings.footerText || '').trim()) {
    var dateStr = new Date().toISOString().slice(0, 10);
    var line1 = '';
    if (settings.pageNum) line1 = '第 ' + (pageIdx + 1) + ' 页';
    if (settings.printDate) line1 += (line1 ? '   ' : '') + '打印日期 ' + dateStr;
    if (line1) {
      page.drawText(line1, {
        x: pw / 2 - line1.length * 1.5,
        y: 8,
        size: 8,
        font: pdfLib.StandardFonts.Helvetica,
        color: pdfLib.rgb(0.5, 0.5, 0.5)
      });
    }
    if (settings.footerText && settings.footerText.trim()) {
      page.drawText(settings.footerText, {
        x: pw / 2 - settings.footerText.length * 1.5,
        y: 2,
        size: 8,
        font: pdfLib.StandardFonts.Helvetica,
        color: pdfLib.rgb(0.5, 0.5, 0.5)
      });
    }
  }
}

async function _composePdfBlob(files, settings, onProgress) {
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
