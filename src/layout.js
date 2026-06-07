// =====================================================
// Layout Calculation & Rendering
// =====================================================
// Dependencies (global): S, MM2PX, PDF_RENDER_DPI, MIN_RENDER_PX

/**
 * Unified layout calculation — pure function used by both preview and print rendering.
 * Returns slot positions, dimensions, and cut-line positions.
 * @param {Object} settings - From getSettings()
 * @param {number} pxPerMm - Pixels per mm (MM2PX for screen, PDF_RENDER_DPI/25.4 for print)
 * @returns {Object} Layout data with slots[], pw, ph, sw, sh, margins, cutLines
 */
function calculateLayout(settings, pxPerMm) {
  pxPerMm = pxPerMm || MM2PX;

  var pw = settings.paperW * pxPerMm;
  var ph = settings.paperH * pxPerMm;
  var mt = settings.marginTop * pxPerMm;
  var mb = settings.marginBottom * pxPerMm;
  var fm = (settings.footerMargin || 0) * pxPerMm; // 页脚边距（独立于发票边距）
  var ml = settings.marginLeft * pxPerMm;
  var mr = settings.marginRight * pxPerMm;
  var gh = settings.gapH * pxPerMm;
  var gv = settings.gapV * pxPerMm;

  // The fm area is reserved purely for footer text below all rows.
  // Only deduct footer margin from slot height when there is footer content.
  // In customFM mode: deduct the explicit footerMargin value.
  // In auto mode: deduct the auto-computed footer height (auto_fm).
  // When there is no footer content: no deduction (no footer to collide with).
  var hasFooterContent = settings.pageNum || settings.printDate || (settings.footerText || '').trim();
  var autoFm = 3 + ((settings.pageNum || settings.printDate ? 1 : 0) + ((settings.footerText || '').trim() ? 1 : 0)) * 5;
  var effectiveFm = hasFooterContent ? (settings.customFM ? fm : autoFm * pxPerMm) : 0;
  var sw = (pw - settings.cols * (ml + mr) - (settings.cols - 1) * gh) / settings.cols;
  var sh = (ph - settings.rows * (mt + mb) - (settings.rows - 1) * gv - effectiveFm) / settings.rows;

  // Calculate slot positions
  var slots = [];
  for (var r = 0; r < settings.rows; r++) {
    for (var c = 0; c < settings.cols; c++) {
      slots.push({
        row: r, col: c,
        x: ml + c * (sw + ml + mr + gh),
        y: mt + r * (sh + mt + mb + gv),
        w: sw, h: sh
      });
    }
  }

  // Cut line positions — based on actual slot boundaries (not page averages)
  var cutLines = [];
  if (settings.cutline && (settings.cols > 1 || settings.rows > 1 || hasFooterContent)) {
    // Horizontal cut lines: between adjacent rows
    for (var r = 1; r < settings.rows; r++) {
      // slot[r-1] bottom edge (top-down) and slot[r] top edge (top-down)
      var slotTopY = mt + r * (sh + mt + mb + gv);       // slot[r].y
      var slotPrevBottomY = mt + (r - 1) * (sh + mt + mb + gv) + sh; // slot[r-1].y + sh
      cutLines.push({ type: 'horizontal', pos: (slotPrevBottomY + slotTopY) / 2 });
    }
    // Footer cut line: between bottom row and footer area
    if (hasFooterContent) {
      if (settings.customFM && fm > 0) {
        // 自定义下边距模式：分割线在用户指定的 fm 位置
        cutLines.push({ type: 'horizontal', pos: ph - fm });
      } else {
        // 默认模式：分割线在页脚文本顶部 + 2mm 间隙，避免贴文字
        // 文本布局（从底部起）：3mm底部间距 + 行数×5mm行高
        var footerLineCount = (settings.pageNum || settings.printDate ? 1 : 0) + ((settings.footerText || '').trim() ? 1 : 0);
        var footerTextTopMm = 3 + footerLineCount * 5 + 2; // 从页面底部算起（mm）
        cutLines.push({ type: 'horizontal', pos: ph - footerTextTopMm * pxPerMm });
      }
    }
    // Vertical cut lines: between adjacent columns (stop at footer area if present)
    var vLineEndY = hasFooterContent ? ph - effectiveFm : ph;
    for (var c = 1; c < settings.cols; c++) {
      var slotLeftX = ml + c * (sw + ml + mr + gh);       // slot[c].x
      var slotPrevRightX = ml + (c - 1) * (sw + ml + mr + gh) + sw; // slot[c-1].x + sw
      cutLines.push({ type: 'vertical', pos: (slotPrevRightX + slotLeftX) / 2, endY: vLineEndY });
    }
  }

  return { pw: pw, ph: ph, mt: mt, mb: mb, fm: fm, ml: ml, mr: mr, gh: gh, gv: gv, sw: sw, sh: sh, slots: slots, cutLines: cutLines };
}

/**
 * Calculate rotation for a file in a slot.
 * @param {Object} fileObj - File object with ow, oh, rotation
 * @param {Object} slot - Slot with w, h
 * @param {Object} settings - Settings with globalRotation
 * @returns {number} Rotation in degrees
 */
function getRotation(fileObj, slot, settings) {
  if (settings.globalRotation === 'auto') {
    var isSlotL = slot.w > slot.h;
    var isImgL = (fileObj.ow || 1) > (fileObj.oh || 1);
    return (isSlotL !== isImgL) ? (fileObj.rotation + 90) % 360 : fileObj.rotation;
  }
  return ((parseInt(settings.globalRotation) || 0) + fileObj.rotation) % 360;
}

// =====================================================
// Preview Rendering (HTML/CSS)
// =====================================================

function renderPage(pageFiles, pi, total, s) {
  var layout = calculateLayout(s);
  var wrap = document.getElementById('previewWrap');
  var scale;
  if (S.viewZoom === 0) {
    scale = Math.min((wrap.clientWidth - 40) / layout.pw, (wrap.clientHeight - 40) / layout.ph, 1.2);
  } else {
    scale = S.viewZoom / 100;
  }
  var dw = Math.round(layout.pw * scale);
  var dh = Math.round(layout.ph * scale);

  var html = '';
  for (var i = 0; i < layout.slots.length; i++) {
    var slot = layout.slots[i];
    var f = pageFiles ? pageFiles[i] : null;
    var imgX = slot.x * scale;
    var imgY = slot.y * scale;
    var imgW = slot.w * scale;
    var imgH = slot.h * scale;
    var inner = '';
    // Per-slot adjustment: scale and offset
    var perScale = f ? (f.slotScale || 1) : 1;
    var perOffX = f ? (f.slotOffsetX || 0) : 0;
    var perOffY = f ? (f.slotOffsetY || 0) : 0;
    var isSelected = (S.selectedSlot === i);
    var selClass = isSelected ? ' selected' : '';

    if (f && f.previewUrl) {
      var src = S.feat.trimWhite && f.trimmedUrl ? f.trimmedUrl : f.previewUrl;
      var rot = getRotation(f, slot, s);
      var filt = s.colorMode === 'grayscale' ? 'filter:grayscale(1);' : s.colorMode === 'bw' ? 'filter:grayscale(1) contrast(1.5);' : '';
      var fit = 'contain';
      if (s.fitMode === 'fill') fit = 'cover';
      else if (s.fitMode === 'original') fit = 'none';
      else if (s.fitMode === 'custom') fit = 'contain';
      var transforms = '';
      // Apply per-slot scale first (before fit-mode custom scale and rotation)
      if (perScale !== 1) transforms += 'scale(' + perScale + ') ';
      if (s.fitMode === 'custom' && s.customScale !== 1) transforms += 'scale(' + s.customScale + ') ';
      if (rot) transforms += 'rotate(' + rot + 'deg) ';
      // Per-slot offset via translate (applied before other transforms)
      if (perOffX !== 0 || perOffY !== 0) {
        // Convert mm to preview pixels: mm * MM2PX(screen px per mm) * scale(preview factor)
        var txPx = perOffX * MM2PX * scale;
        var tyPx = perOffY * MM2PX * scale;
        transforms = 'translate(' + txPx.toFixed(1) + 'px, ' + tyPx.toFixed(1) + 'px) ' + transforms;
      }
      // Calculate contained image dimensions for border to follow invoice
      var imgObjW = f.ow || 1;
      var imgObjH = f.oh || 1;
      var containedW, containedH;
      if (s.fitMode === 'original') {
        containedW = imgObjW;
        containedH = imgObjH;
      } else if (s.fitMode === 'fill') {
        containedW = imgW;
        containedH = imgH;
      } else {
        // contain / custom: image fits in slot maintaining aspect ratio
        var fitScale = Math.min(imgW / imgObjW, imgH / imgObjH);
        containedW = imgObjW * fitScale;
        containedH = imgObjH * fitScale;
      }
      // Image wrapper: explicit dimensions, same transforms, optional border
      var wrapperStyle = 'width:' + containedW.toFixed(1) + 'px;height:' + containedH.toFixed(1) + 'px;';
      wrapperStyle += 'position:absolute;';
      wrapperStyle += 'left:' + ((imgW - containedW) / 2).toFixed(1) + 'px;';
      wrapperStyle += 'top:' + ((imgH - containedH) / 2).toFixed(1) + 'px;';
      wrapperStyle += 'transform-origin:center center;';
      if (transforms) wrapperStyle += 'transform:' + transforms + ';';
      if (s.border) wrapperStyle += 'outline:1px solid #000;outline-offset:-1px;';
      // Image fills wrapper
      var imgStyle = 'width:100%;height:100%;object-fit:' + fit + ';' + filt;
      inner = '<div style="' + wrapperStyle + '"><img src="' + src + '" style="' + imgStyle + '"></div>';
      if (s.number) inner += '<div class="slot-num">' + (pi * s.rows * s.cols + i + 1) + '</div>';
      if (s.watermark && s.watermarkText) {
        var ws = s.watermarkSize * MM2PX * scale;
        inner += '<div class="watermark" style="color:' + s.watermarkColor + ';opacity:' + s.watermarkOpacity + ';font-size:' + ws + 'px;transform:translate(-50%,-50%) rotate(' + s.watermarkAngle + 'deg);top:50%;left:50%">' + s.watermarkText + '</div>';
      }
      // Resize handles (visible only when selected)
      inner += '<div class="slot-handle slot-handle-tl" data-handle="tl"></div>';
      inner += '<div class="slot-handle slot-handle-tr" data-handle="tr"></div>';
      inner += '<div class="slot-handle slot-handle-bl" data-handle="bl"></div>';
      inner += '<div class="slot-handle slot-handle-br" data-handle="br"></div>';
      html += '<div class="invoice-slot' + selClass + '" data-slot-idx="' + i + '" style="position:absolute;left:' + imgX + 'px;top:' + imgY + 'px;width:' + imgW + 'px;height:' + imgH + 'px;">' + inner + '</div>';
    } else {
      inner = '<div class="slot-empty">空</div>';
      html += '<div class="invoice-slot' + selClass + '" data-slot-idx="' + i + '" style="position:absolute;left:' + imgX + 'px;top:' + imgY + 'px;width:' + imgW + 'px;height:' + imgH + 'px">' + inner + '</div>';
    }
  }

  // Cut lines
  for (var cl = 0; cl < layout.cutLines.length; cl++) {
    var line = layout.cutLines[cl];
    if (line.type === 'horizontal') {
      html += '<div class="cut-line" style="top:' + (line.pos * scale) + 'px"></div>';
    } else {
      var vStyle = 'left:' + (line.pos * scale) + 'px';
      if (line.endY !== undefined) vStyle += ';height:' + (line.endY * scale) + 'px';
      html += '<div class="cut-line-v" style="' + vStyle + '"></div>';
    }
  }

  // 页脚文本行序（从下到上）：自定义页脚 → 页码/日期
  // 所有位置和字号必须乘以 scale，与 slot 坐标系一致
  var textBottomPx = 3 * MM2PX * scale;
  var lineHeightPx = 5 * MM2PX * scale;
  var footerFontSize = Math.max(8, 10 * scale);

  // 自定义页脚：最下面一行
  var footerText = (s.footerText || '').trim();
  if (footerText) {
    html += '<div style="position:absolute;bottom:' + textBottomPx + 'px;left:0;right:0;text-align:center;font-size:' + footerFontSize + 'px;color:#94a3b8">' + escHtml(footerText) + '</div>';
  }

  // 页码/日期：在自定义页脚上方
  var pageNumBottomPx = textBottomPx;
  if (footerText) pageNumBottomPx += lineHeightPx;
  if (s.pageNum) {
    var pageNumStyle = s.printDate ? 'position:absolute;bottom:' + pageNumBottomPx + 'px;left:' + (10 * MM2PX * scale) + 'px;font-size:' + footerFontSize + 'px;color:#94a3b8' : 'position:absolute;bottom:' + pageNumBottomPx + 'px;left:0;right:0;text-align:center;font-size:' + footerFontSize + 'px;color:#94a3b8';
    html += '<div style="' + pageNumStyle + '">第 ' + (pi + 1) + ' 页 / 共 ' + total + ' 页</div>';
  }
  if (s.printDate) {
    var now = new Date();
    var dateStr = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
    var dateStyle = s.pageNum ? 'position:absolute;bottom:' + pageNumBottomPx + 'px;right:' + (10 * MM2PX * scale) + 'px;font-size:' + footerFontSize + 'px;color:#94a3b8' : 'position:absolute;bottom:' + pageNumBottomPx + 'px;left:0;right:0;text-align:center;font-size:' + footerFontSize + 'px;color:#94a3b8';
    html += '<div style="' + dateStyle + '">打印日期 ' + dateStr + '</div>';
  }

  document.getElementById('previewPages').style.display = 'block';
  document.getElementById('emptyState').style.display = 'none';
  document.getElementById('previewPages').innerHTML = '<div class="preview-container" style="width:' + dw + 'px;height:' + dh + 'px"><div style="width:' + dw + 'px;height:' + dh + 'px;background:white;position:relative">' + html + '</div></div>';
  document.getElementById('pageInfo').textContent = (pi + 1) + ' / ' + total;
  document.getElementById('prevBtn').disabled = pi === 0;
  document.getElementById('nextBtn').disabled = pi === total - 1;
  document.getElementById('pageNav').style.display = 'flex';

  // Re-apply selection highlight and bind interaction
  if (S.selectedSlot >= 0) {
    var selEl = document.querySelector('.invoice-slot[data-slot-idx="' + S.selectedSlot + '"]');
    if (selEl) selEl.classList.add('selected');
  }
  initSlotInteraction();
}

// =====================================================
// Per-slot Interaction — drag & resize in preview
// =====================================================

var _slotDrag = null; // Current drag/resize state

var _slotInteractionBound = false;

/**
 * Bind mousedown on invoice-slot elements for drag-move and corner-resize.
 * Called after each renderPage(). Only binds once.
 */
function initSlotInteraction() {
  var container = document.getElementById('previewPages');
  if (!container) return;
  if (_slotInteractionBound) return;
  _slotInteractionBound = true;
  container.addEventListener('mousedown', onSlotMouseDown);
  // Click on empty area deselects
  document.getElementById('previewWrap').addEventListener('mousedown', function(e) {
    if (!e.target.closest('.invoice-slot') && !e.target.closest('.slot-handle')) {
      selectSlot(-1);
    }
  });
}

function onSlotMouseDown(e) {
  var slotEl = e.target.closest('.invoice-slot');
  if (!slotEl || slotEl.querySelector('.slot-empty')) return;

  var idx = parseInt(slotEl.dataset.slotIdx);
  if (isNaN(idx)) return;

  // Check if clicking a resize handle
  var handle = e.target.closest('.slot-handle');
  if (handle) {
    e.preventDefault();
    e.stopPropagation();
    startResize(e, idx, slotEl, handle.dataset.handle);
    return;
  }

  // Otherwise: click to select + drag to move
  e.preventDefault();
  selectSlot(idx);

  var files = getActiveFiles();
  var settings = getSettings();
  var layout = calculateLayout(settings);
  var perPage = settings.cols * settings.rows;
  var fileIdx = S.currentPage * perPage + idx;
  var f = fileIdx < files.length ? files[fileIdx] : null;
  if (!f) return;

  _slotDrag = {
    mode: 'move',
    slotEl: slotEl,
    wrapperEl: slotEl.querySelector(':scope > div'),
    fileObj: f,
    idx: idx,
    startX: e.clientX,
    startY: e.clientY,
    startOffX: f.slotOffsetX || 0,
    startOffY: f.slotOffsetY || 0,
    previewScale: getCurrentPreviewScale(),
    // Cache settings/layout for perf (avoid getSettings() every mousemove)
    cachedSettings: settings,
    cachedLayout: layout
  };
  slotEl.classList.add('dragging');

  document.addEventListener('mousemove', onSlotMouseMove);
  document.addEventListener('mouseup', onSlotMouseUp);
}

function startResize(e, idx, slotEl, corner) {
  selectSlot(idx);

  var files = getActiveFiles();
  var settings = getSettings();
  var layout = calculateLayout(settings);
  var perPage = settings.cols * settings.rows;
  var fileIdx = S.currentPage * perPage + idx;
  var f = fileIdx < files.length ? files[fileIdx] : null;
  if (!f) return;

  var slot = layout.slots[idx];

  _slotDrag = {
    mode: 'resize',
    corner: corner,
    slotEl: slotEl,
    wrapperEl: slotEl.querySelector(':scope > div'),
    fileObj: f,
    idx: idx,
    startX: e.clientX,
    startY: e.clientY,
    startScale: f.slotScale || 1,
    startDist: Math.max(10, Math.hypot(
      e.clientX - (slotEl.getBoundingClientRect().left + slotEl.offsetWidth / 2),
      e.clientY - (slotEl.getBoundingClientRect().top + slotEl.offsetHeight / 2)
    )),
    previewScale: getCurrentPreviewScale(),
    cachedSettings: settings,
    cachedLayout: layout
  };

  document.addEventListener('mousemove', onSlotMouseMove);
  document.addEventListener('mouseup', onSlotMouseUp);
}

function onSlotMouseMove(e) {
  if (!_slotDrag) return;
  e.preventDefault();
  _slotDrag.moved = true;  // Track actual mouse movement

  var settings = _slotDrag.cachedSettings;
  var layout = _slotDrag.cachedLayout;

  if (_slotDrag.mode === 'move') {
    var dx = e.clientX - _slotDrag.startX;
    var dy = e.clientY - _slotDrag.startY;
    // Convert pixel delta to mm
    var ps = _slotDrag.previewScale;
    var dxMm = dx / (MM2PX * ps);
    var dyMm = dy / (MM2PX * ps);
    var newOffX = _slotDrag.startOffX + dxMm;
    var newOffY = _slotDrag.startOffY + dyMm;
    // Clamp: limit offset so invoice doesn't go fully outside slot
    var slot = layout.slots[_slotDrag.idx];
    var f = _slotDrag.fileObj;
    var imgW = f.ow || 1;
    var imgH = f.oh || 1;
    var s = _slotDrag.cachedSettings;
    var displayW, displayH;
    if (s.fitMode === 'fill') {
      displayW = slot.w * (f.slotScale || 1);
      displayH = slot.h * (f.slotScale || 1);
    } else if (s.fitMode === 'original') {
      displayW = imgW * (f.slotScale || 1);
      displayH = imgH * (f.slotScale || 1);
    } else {
      var fitScale = Math.min(slot.w / imgW, slot.h / imgH);
      var perScale = f.slotScale || 1;
      displayW = imgW * fitScale * perScale;
      displayH = imgH * fitScale * perScale;
      if (s.fitMode === 'custom' && s.customScale !== 1) {
        displayW *= s.customScale;
        displayH *= s.customScale;
      }
    }
    var extraX = Math.max(0, (displayW - slot.w) / 2 / MM2PX);
    var extraY = Math.max(0, (displayH - slot.h) / 2 / MM2PX);
    var minRangeX = (slot.w / MM2PX) * 0.5;
    var minRangeY = (slot.h / MM2PX) * 0.5;
    var maxOffX = Math.max(minRangeX, extraX);
    var maxOffY = Math.max(minRangeY, extraY);
    newOffX = Math.max(-maxOffX, Math.min(maxOffX, newOffX));
    newOffY = Math.max(-maxOffY, Math.min(maxOffY, newOffY));
    _slotDrag.fileObj.slotOffsetX = Math.round(newOffX * 10) / 10;
    _slotDrag.fileObj.slotOffsetY = Math.round(newOffY * 10) / 10;

    // Real-time visual feedback: update CSS transform directly on wrapper div
    if (_slotDrag.wrapperEl) {
      var transforms = buildTransformString(_slotDrag.fileObj, settings, layout.slots[_slotDrag.idx]);
      _slotDrag.wrapperEl.style.transform = transforms;
    }
  } else if (_slotDrag.mode === 'resize') {
    var slotRect = _slotDrag.slotEl.getBoundingClientRect();
    var cx = slotRect.left + slotRect.width / 2;
    var cy = slotRect.top + slotRect.height / 2;
    var dist = Math.hypot(e.clientX - cx, e.clientY - cy);
    var ratio = dist / _slotDrag.startDist;
    var newScale = Math.max(0.2, Math.min(3.0, _slotDrag.startScale * ratio));
    _slotDrag.fileObj.slotScale = Math.round(newScale * 100) / 100;

    // Real-time visual feedback
    if (_slotDrag.wrapperEl) {
      var transforms = buildTransformString(_slotDrag.fileObj, settings, layout.slots[_slotDrag.idx]);
      _slotDrag.wrapperEl.style.transform = transforms;
    }
  }
}

function onSlotMouseUp(e) {
  if (!_slotDrag) return;
  _slotDrag.slotEl.classList.remove('dragging');
  var didMove = !!_slotDrag.moved;
  _slotDrag = null;
  document.removeEventListener('mousemove', onSlotMouseMove);
  document.removeEventListener('mouseup', onSlotMouseUp);
  if (didMove) {
    updatePreview();
    updateAdjPanel();
  }
}

/**
 * Build CSS transform string for per-slot adjustments.
 * Mirrors the logic in renderPage.
 */
function buildTransformString(f, s, slot) {
  var perScale = f.slotScale || 1;
  var perOffX = f.slotOffsetX || 0;
  var perOffY = f.slotOffsetY || 0;
  var rot = getRotation(f, slot, s);
  var ps = getCurrentPreviewScale();
  var transforms = '';

  if (perOffX !== 0 || perOffY !== 0) {
    var txPx = perOffX * MM2PX * ps;
    var tyPx = perOffY * MM2PX * ps;
    transforms += 'translate(' + txPx.toFixed(1) + 'px, ' + tyPx.toFixed(1) + 'px) ';
  }
  if (perScale !== 1) transforms += 'scale(' + perScale + ') ';
  if (s.fitMode === 'custom' && s.customScale !== 1) transforms += 'scale(' + s.customScale + ') ';
  if (rot) transforms += 'rotate(' + rot + 'deg) ';
  return transforms || 'none';
}

/**
 * Get the current preview scale factor (preview pixels / mm).
 */
function getCurrentPreviewScale() {
  var wrap = document.getElementById('previewWrap');
  if (!wrap) return 1;
  if (S.viewZoom === 0) {
    var settings = getSettings();
    var layout = calculateLayout(settings);
    return Math.min((wrap.clientWidth - 40) / layout.pw, (wrap.clientHeight - 40) / layout.ph, 1.2);
  }
  return S.viewZoom / 100;
}

// =====================================================
// Canvas Rendering — REMOVED in v1.4.2
// =====================================================
// PDF generation now goes through Rust generate_pdf_from_layout command.
// The browser fallback (fallbackPrint in print.js) uses HTML/CSS, not canvas.
// The <canvas id="renderCanvas"> element is also removed from index.html.
