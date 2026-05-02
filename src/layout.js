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
  var ml = settings.marginLeft * pxPerMm;
  var mr = settings.marginRight * pxPerMm;
  var gh = settings.gapH * pxPerMm;
  var gv = settings.gapV * pxPerMm;

  // Per-slot margins: each invoice has its own margin within its slot
  var sw = (pw - settings.cols * (ml + mr) - (settings.cols - 1) * gh) / settings.cols;
  var sh = (ph - settings.rows * (mt + mb) - (settings.rows - 1) * gv) / settings.rows;

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

  // Cut line positions
  var cutLines = [];
  if (settings.cutline && (settings.cols > 1 || settings.rows > 1)) {
    for (var r = 1; r < settings.rows; r++) {
      cutLines.push({ type: 'horizontal', pos: r * (sh + mt + mb + gv) - gv / 2 });
    }
    for (var c = 1; c < settings.cols; c++) {
      cutLines.push({ type: 'vertical', pos: c * (sw + ml + mr + gh) - gh / 2 });
    }
  }

  return { pw: pw, ph: ph, mt: mt, mb: mb, ml: ml, mr: mr, gh: gh, gv: gv, sw: sw, sh: sh, slots: slots, cutLines: cutLines };
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
      var bdr = s.border ? 'box-shadow:inset 0 0 0 1px #000;' : 'box-shadow:inset 0 0 0 0.5px rgba(0,0,0,0.1);';
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
      // Rotation sizing fix:
      var isRotated90 = (rot === 90 || rot === 270);
      var sizeStyle = '';
      if (s.fitMode === 'original') {
        sizeStyle = '';
      } else if (isRotated90) {
        sizeStyle = 'max-width:' + imgH + 'px;max-height:' + imgW + 'px;';
      } else {
        sizeStyle = 'max-width:100%;max-height:100%;';
      }
      inner = '<img src="' + src + '" style="' + sizeStyle + 'object-fit:' + fit + ';' + filt + (transforms ? 'transform:' + transforms + ';' : '') + '">';
      if (s.number) inner += '<div class="slot-num">' + (pi * s.rows * s.cols + i + 1) + '</div>';
      if (s.watermark && s.watermarkText) {
        var ws = Math.min(slot.w * scale, slot.h * scale) * 0.15;
        inner += '<div class="watermark" style="color:' + s.watermarkColor + ';opacity:' + s.watermarkOpacity + ';font-size:' + ws + 'px;transform:translate(-50%,-50%) rotate(' + s.watermarkAngle + 'deg);top:50%;left:50%">' + s.watermarkText + '</div>';
      }
      // Resize handles (visible only when selected)
      inner += '<div class="slot-handle slot-handle-tl" data-handle="tl"></div>';
      inner += '<div class="slot-handle slot-handle-tr" data-handle="tr"></div>';
      inner += '<div class="slot-handle slot-handle-bl" data-handle="bl"></div>';
      inner += '<div class="slot-handle slot-handle-br" data-handle="br"></div>';
      html += '<div class="invoice-slot' + selClass + '" data-slot-idx="' + i + '" style="position:absolute;left:' + imgX + 'px;top:' + imgY + 'px;width:' + imgW + 'px;height:' + imgH + 'px;' + bdr + '">' + inner + '</div>';
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
      html += '<div class="cut-line-v" style="left:' + (line.pos * scale) + 'px"></div>';
    }
  }

  if (s.pageNum) html += '<div style="position:absolute;bottom:5px;left:0;right:0;text-align:center;font-size:10px;color:#94a3b8">第 ' + (pi + 1) + ' 页 / 共 ' + total + ' 页</div>';

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
    imgEl: slotEl.querySelector('img'),
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
    imgEl: slotEl.querySelector('img'),
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
    var maxOffX = (slot.w / MM2PX) * 0.5;
    var maxOffY = (slot.h / MM2PX) * 0.5;
    newOffX = Math.max(-maxOffX, Math.min(maxOffX, newOffX));
    newOffY = Math.max(-maxOffY, Math.min(maxOffY, newOffY));
    _slotDrag.fileObj.slotOffsetX = Math.round(newOffX * 10) / 10;
    _slotDrag.fileObj.slotOffsetY = Math.round(newOffY * 10) / 10;

    // Real-time visual feedback: update CSS transform directly
    if (_slotDrag.imgEl) {
      var transforms = buildTransformString(_slotDrag.fileObj, settings, layout.slots[_slotDrag.idx]);
      _slotDrag.imgEl.style.transform = transforms;
    }
  } else if (_slotDrag.mode === 'resize') {
    var slotRect = _slotDrag.slotEl.getBoundingClientRect();
    var cx = slotRect.left + slotRect.width / 2;
    var cy = slotRect.top + slotRect.height / 2;
    var dist = Math.hypot(e.clientX - cx, e.clientY - cy);
    var ratio = dist / _slotDrag.startDist;
    var newScale = Math.max(0.2, Math.min(2.0, _slotDrag.startScale * ratio));
    _slotDrag.fileObj.slotScale = Math.round(newScale * 100) / 100;

    // Real-time visual feedback
    if (_slotDrag.imgEl) {
      var transforms = buildTransformString(_slotDrag.fileObj, settings, layout.slots[_slotDrag.idx]);
      _slotDrag.imgEl.style.transform = transforms;
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
    _pdfDirty = true;
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
