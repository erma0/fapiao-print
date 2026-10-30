// =====================================================
// Layout Calculation & Rendering
// =====================================================
// Dependencies (global): S, MM2PX, PDF_RENDER_DPI, MIN_RENDER_PX

/**
 * Slots per page — photo mode: photo grid capacity; reimburse: fixed-height segments; otherwise cols*rows.
 */
function getPerPage(s) {
  if (s.photoMode) {
    var pg = getPhotoGrid(s);
    return Math.max(1, pg.cols * pg.rows);
  }
  if (s.reimburseMode) {
    return Math.max(1, Math.floor(s.paperH / (s.reimburseHeight || 120)));
  }
  return s.cols * s.rows;
}

/**
 * 相纸模式网格（mm）：按证件照实际尺寸推算每页列/行、间距、留边与居中栅格。
 * 手动列/行（>0）超出可容纳数量时收敛到上限，防止栅格溢出纸张。
 */
function getPhotoGrid(s) {
  var w = Math.max(1, parseFloat(s.photoW) || 25);
  var h = Math.max(1, parseFloat(s.photoH) || 35);
  var gap = Math.max(0, parseFloat(s.photoGap) || 0);
  var margin = Math.max(0, parseFloat(s.photoMargin) || 0);
  var pw = s.paperW, ph = s.paperH;
  // 留边不能把照片挤出纸张；纸张本身小于照片时留边归零（UI 另行提示可能被裁切）
  var mx = Math.min(margin, Math.max(0, (pw - w) / 2));
  var my = Math.min(margin, Math.max(0, (ph - h) / 2));
  var availW = Math.max(1, pw - 2 * mx);
  var availH = Math.max(1, ph - 2 * my);
  // 自动铺满：按留边计算；手动列/行：只受纸张物理尺寸约束（留边可被压缩，栅格仍居中）
  var maxCols = Math.max(1, Math.floor((availW + gap) / (w + gap)));
  var maxRows = Math.max(1, Math.floor((availH + gap) / (h + gap)));
  var fitCols = Math.max(1, Math.floor((pw + gap) / (w + gap)));
  var fitRows = Math.max(1, Math.floor((ph + gap) / (h + gap)));
  var cols = s.photoCols > 0 ? Math.min(s.photoCols, fitCols) : maxCols;
  var rows = s.photoRows > 0 ? Math.min(s.photoRows, fitRows) : maxRows;
  var gridW = cols * w + (cols - 1) * gap;
  var gridH = rows * h + (rows - 1) * gap;
  return { w: w, h: h, gap: gap, cols: cols, rows: rows, maxCols: maxCols, maxRows: maxRows, gridW: gridW, gridH: gridH };
}

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

  // 相纸模式：槽位严格等于证件照实际尺寸，整版居中；普通行列/边距/间距/页脚全部忽略。
  // 与 Rust calculate_layout_mm 的 photo 分支逐字对应。
  if (settings.photoMode) {
    var pg = getPhotoGrid(settings);
    var startX = (settings.paperW - pg.gridW) / 2;
    var startY = (settings.paperH - pg.gridH) / 2;
    var pSlots = [];
    for (var pr = 0; pr < pg.rows; pr++) {
      for (var pc = 0; pc < pg.cols; pc++) {
        pSlots.push({
          row: pr, col: pc,
          x: (startX + pc * (pg.w + pg.gap)) * pxPerMm,
          y: (startY + pr * (pg.h + pg.gap)) * pxPerMm,
          w: pg.w * pxPerMm,
          h: pg.h * pxPerMm
        });
      }
    }
    var pCuts = [];
    if (settings.cutline) {
      for (var pcx = 1; pcx < pg.cols; pcx++) {
        pCuts.push({ type: 'vertical', pos: (startX + pcx * (pg.w + pg.gap) - pg.gap / 2) * pxPerMm, endY: ph });
      }
      for (var pry = 1; pry < pg.rows; pry++) {
        pCuts.push({ type: 'horizontal', pos: (startY + pry * (pg.h + pg.gap) - pg.gap / 2) * pxPerMm });
      }
    }
    return {
      pw: pw, ph: ph,
      mt: startY * pxPerMm, mb: (settings.paperH - startY - pg.gridH) * pxPerMm, fm: 0,
      ml: startX * pxPerMm, mr: (settings.paperW - startX - pg.gridW) * pxPerMm,
      gh: pg.gap * pxPerMm, gv: pg.gap * pxPerMm,
      sw: pg.w * pxPerMm, sh: pg.h * pxPerMm,
      slots: pSlots, cutLines: pCuts, photo: true
    };
  }

  var mt = settings.marginTop * pxPerMm;
  var mb = settings.marginBottom * pxPerMm;
  var fm = (settings.footerMargin || 0) * pxPerMm; // 页脚边距（独立于发票边距）
  var ml = settings.marginLeft * pxPerMm;
  var mr = settings.marginRight * pxPerMm;
  var gh = settings.gapH * pxPerMm;
  var gv = settings.gapV * pxPerMm;

  // 粘贴单模式：改用独立边距（上边距 = 装订区高度，装订线即画在该处）。
  // 与 Rust calculate_layout_mm 的解析规则逐字对应，保证预览/PDF 同源。
  if (settings.pasteMode) {
    mt = (settings.pasteTop != null ? settings.pasteTop : 20.8) * pxPerMm;
    mb = (settings.pasteBottom != null ? settings.pasteBottom : 4.1) * pxPerMm;
    ml = (settings.pasteLeft != null ? settings.pasteLeft : 4.1) * pxPerMm;
    mr = (settings.pasteRight != null ? settings.pasteRight : 4.1) * pxPerMm;
  }

  // 报销单分段模式：固定段高（默认120mm），裁切线位于段边界绝对位置（k×段高），
  // 不受任何边距影响；mt/mb 仅作为段内上下安全边距，ml/mr 决定发票区域左右边界。
  // 忽略 rows/cols/gap/footerMargin；裁切线强制绘制（本模式的核心目的）。
  if (settings.reimburseMode) {
    var segMm = settings.reimburseHeight || 120;
    var segCount = Math.max(1, Math.floor(settings.paperH / segMm));
    var rsw = pw - ml - mr;
    var rsh = Math.max(10 * pxPerMm, (segMm - settings.marginTop - settings.marginBottom) * pxPerMm);
    var rSlots = [];
    for (var si = 0; si < segCount; si++) {
      rSlots.push({
        row: si, col: 0,
        x: ml,
        y: (si * segMm + settings.marginTop) * pxPerMm,
        w: rsw, h: rsh
      });
    }
    var rCutLines = [];
    // 段底裁切线：k = 1..segCount，每段底部一条（含最后一段），位置 = k × 段高（绝对位置）
    for (var k = 1; k <= segCount; k++) {
      rCutLines.push({ type: 'horizontal', pos: k * segMm * pxPerMm });
    }
    return { pw: pw, ph: ph, mt: mt, mb: mb, fm: fm, ml: ml, mr: mr, gh: gh, gv: gv, sw: rsw, sh: rsh, slots: rSlots, cutLines: rCutLines, reimburse: true };
  }

  // The fm area is reserved purely for footer text below all rows.
  // Only deduct footer margin from slot height when there is footer content.
  // In customFM mode: deduct the explicit footerMargin value.
  // In auto mode: deduct the auto-computed footer height (auto_fm).
  // When there is no footer content: no deduction (no footer to collide with).
  var hasFooterContent = settings.pageNum || settings.printDate || (settings.footerText || '').trim();
  var autoFm = 3 + ((settings.pageNum || settings.printDate ? 1 : 0) + ((settings.footerText || '').trim() ? 1 : 0)) * 5;
  var footerFm = hasFooterContent ? (settings.customFM ? fm : autoFm * pxPerMm) : 0;
  // 粘贴单：在页面底部为右下角签字栏预留一整条带（表头行 + 填写行 + 与票据区的间距），
  // 保证票据只排版在装订线下方、签字栏上方。签字栏底边落在下边距处，故此处不含 mb。
  var sigPx = (settings.pasteMode && settings.pasteShowSig !== false)
    ? ((settings.pasteSigRowH != null ? settings.pasteSigRowH : 8)
      + (settings.pasteSigBodyH != null ? settings.pasteSigBodyH : 14)
      + (settings.pasteSigGap != null ? settings.pasteSigGap : 2)) * pxPerMm
    : 0;
  var effectiveFm = footerFm + sigPx;
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
  if (settings.cutline && (settings.cols > 1 || settings.rows > 1 || hasFooterContent || sigPx > 0)) {
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
    // Vertical cut lines: between adjacent columns (stop at footer/signature area if present)
    // 粘贴单：额外多算一个下边距，让垂直线与票据区底边齐平，不切进右下角签字栏
    var vLineReserve = sigPx > 0 ? sigPx + mb : effectiveFm;
    var vLineEndY = (hasFooterContent || sigPx > 0) ? ph - vLineReserve : ph;
    for (var c = 1; c < settings.cols; c++) {
      var slotLeftX = ml + c * (sw + ml + mr + gh);       // slot[c].x
      var slotPrevRightX = ml + (c - 1) * (sw + ml + mr + gh) + sw; // slot[c-1].x + sw
      cutLines.push({ type: 'vertical', pos: (slotPrevRightX + slotLeftX) / 2, endY: vLineEndY });
    }
  }

  return { pw: pw, ph: ph, mt: mt, mb: mb, fm: fm, ml: ml, mr: mr, gh: gh, gv: gv, sw: sw, sh: sh, slots: slots, cutLines: cutLines, sig: sigPx };
}

/**
 * 粘贴单：装订线 + 右下角签字栏的 HTML（只画线，数据留空手写）。
 * @param {Object} s - getSettings()
 * @param {number} k - 毫米 → 输出单位的换算系数（预览传 MM2PX*scale，HTML 打印兜底传 1）
 * @param {string} unit - 输出单位 'px' 或 'mm'
 */
function buildPasteSheetHtml(s, k, unit) {
  if (!s.pasteMode) return '';
  var u = function (mm) { return (mm * k).toFixed(2) + unit; };
  var uf = function (mm) { return unit === 'px' ? Math.max(8, mm * k).toFixed(1) + 'px' : mm.toFixed(2) + 'mm'; };
  // 缺省回落值必须与 calculateLayout / Rust paste_sheet_geom 完全一致；
  // 用 != null 而非 || ，否则用户把边距设成 0 会被默认值覆盖（线画在 20.8 而槽位按 0 算）。
  var d = function (v, def) { return v != null ? v : def; };
  var html = '';

  // 装订线：上边距处一条横贯全页的实线，线上方居中「装 订 线」
  if (s.pasteBindLine !== false) {
    var bt = (s.pasteBindText == null ? '装 订 线' : s.pasteBindText).trim();
    var btFsMm = d(s.pasteBindSize, 4);
    var btTopMm = d(s.pasteTop, 20.8);
    var btFs = unit === 'px' ? Math.max(8, btFsMm * k) : btFsMm * k;
    html += '<div style="position:absolute;left:0;right:0;top:' + u(btTopMm)
      + ';border-top:1px solid #334155;pointer-events:none"></div>';
    if (bt) {
      // 文字底边与装订线留 1.5mm 间隙
      html += '<div style="position:absolute;left:0;right:0;top:' + u(btTopMm - btFsMm - 1.5)
        + ';height:' + u(btFsMm) + ';line-height:' + u(btFsMm)
        + ';text-align:center;font-size:' + (unit === 'px' ? btFs.toFixed(1) + 'px' : btFs.toFixed(2) + 'mm')
        + ';color:#334155;pointer-events:none">' + escHtml(bt) + '</div>';
    }
  }
  // 右下角签字栏：表头行 + 一个空白填写行，右边距/下边距对齐
  if (s.pasteShowSig !== false) {
    var colsRaw = s.pasteSigCols == null ? '票据张数,金额,报销人' : s.pasteSigCols;
    var cols = colsRaw.split(/[,，]/).map(function (x) { return x.trim(); }).filter(Boolean);
    if (cols.length > 0) {
      var fsMm = 3.5;
      var cell = 'border:1px solid #334155;text-align:center;vertical-align:middle;padding:0;';
      html += '<table style="position:absolute;right:' + u(d(s.pasteRight, 4.1))
        + ';bottom:' + u(d(s.pasteBottom, 4.1))
        + ';width:' + u(d(s.pasteSigWidth, 80))
        + ';border-collapse:collapse;table-layout:fixed;pointer-events:none">';
      html += '<tr>';
      cols.forEach(function (c) {
        html += '<td style="' + cell + 'height:' + u(d(s.pasteSigRowH, 8))
          + ';font-size:' + uf(fsMm) + ';color:#334155">' + escHtml(c) + '</td>';
      });
      html += '</tr><tr>';
      cols.forEach(function () {
        html += '<td style="' + cell + 'height:' + u(d(s.pasteSigBodyH, 14)) + '"></td>';
      });
      html += '</tr></table>';
    }
  }
  return html;
}

/**
 * 有效显示尺寸（px）。开启「裁剪白边」且已生成裁剪缓存时用裁剪后的尺寸，
 * 否则用文件原始尺寸（ow/oh）。预览适配、拖拽约束、九宫格对齐与 PDF 导出
 * 必须统一走这里，否则各处算出的适配比例会不一致。
 * @param {Object} fileObj - File object with ow, oh, trimmedW, trimmedH
 * @param {Object} settings - Settings with trimWhite
 * @returns {{w:number,h:number}}
 */
function getObjDims(fileObj, settings) {
  if (!fileObj) return { w: 1, h: 1 };
  // 开启「裁剪白边」且已生成裁剪缓存时才用裁剪后尺寸
  var trimOn = settings && settings.trimWhite;
  if (trimOn && fileObj.trimmedW > 0 && fileObj.trimmedH > 0) {
    return { w: fileObj.trimmedW, h: fileObj.trimmedH };
  }
  return { w: fileObj.ow || 1, h: fileObj.oh || 1 };
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
    var dims = getObjDims(fileObj, settings);
    var isImgL = dims.w > dims.h;
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
  // 副本标记：按全局展开序列预计算 每槽位第几份 / 总份数（仅预览显示，不进打印/PDF）
  var copySeq = null, copyTotal = null;
  if (s.copyBadge) {
    copySeq = [];
    copyTotal = {};
    getActiveFiles().forEach(function(af) {
      copyTotal[af.id] = (copyTotal[af.id] || 0) + 1;
      copySeq.push(copyTotal[af.id]);
    });
  }
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
      // 用 settings.trimWhite 而非 S.feat.trimWhite：相纸模式在 getSettings 里强制关闭裁剪，
      // 若这里仍看功能开关，会出现「预览是裁剪图、PDF 未裁剪」的不一致。
      var src = s.trimWhite && f.trimmedUrl ? f.trimmedUrl : f.previewUrl;
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
      var _dispDims = getObjDims(f, s);
      var imgObjW = _dispDims.w;
      var imgObjH = _dispDims.h;
      // 旋转 90°/270°：与 PDF 导出一致，按旋转后的视觉宽高适配槽位（先旋转后适配）。
      // wrapper 是旋转前的盒子，取视觉盒的转置，CSS 旋转后正好落在视觉盒上。
      var isRot90 = (rot === 90 || rot === 270);
      var fitW = isRot90 ? imgObjH : imgObjW;
      var fitH = isRot90 ? imgObjW : imgObjH;
      var containedW, containedH, visW, visH;
      if (s.fitMode === 'original') {
        containedW = imgObjW;
        containedH = imgObjH;
        visW = isRot90 ? containedH : containedW;
        visH = isRot90 ? containedW : containedH;
      } else if (s.fitMode === 'fill') {
        containedW = isRot90 ? imgH : imgW;
        containedH = isRot90 ? imgW : imgH;
        visW = isRot90 ? containedH : containedW;
        visH = isRot90 ? containedW : containedH;
      } else {
        // contain / custom: 旋转后视觉宽高 contain-fit 槽位，wrapper 转置
        var fitScale = Math.min(imgW / fitW, imgH / fitH);
        visW = fitW * fitScale;
        visH = fitH * fitScale;
        containedW = isRot90 ? visH : visW;
        containedH = isRot90 ? visW : visH;
      }
      // Image wrapper: explicit dimensions, same transforms, optional border
      // 报销单模式：左上对齐（贴段内区域左上角），常规模式：居中
      var wrapperStyle = 'width:' + containedW.toFixed(1) + 'px;height:' + containedH.toFixed(1) + 'px;';
      wrapperStyle += 'position:absolute;';
      if (layout.reimburse) {
        if (isRot90) {
          // 旋转后视觉盒以 wrapper 中心为中心，须把 wrapper 中心移到视觉盒中心位置，
          // 旋转落位后视觉盒才贴住槽位左上角（与 Rust 端报销单左上对齐一致）
          wrapperStyle += 'left:' + ((containedH - containedW) / 2).toFixed(1) + 'px;';
          wrapperStyle += 'top:' + ((containedW - containedH) / 2).toFixed(1) + 'px;';
        } else {
          wrapperStyle += 'left:0;top:0;';
        }
      } else {
        // 「裁剪白边」开启时垂直贴顶（视觉盒顶 = 槽位顶）：wrapper 中心对齐到 visH/2，
        // 否则居中。水平方向始终居中 —— 左右裁剪本来就对得准。
        var wrapTop = ((imgH - containedH) / 2);
        if (s.trimWhite) wrapTop = visH / 2 - containedH / 2;
        wrapperStyle += 'left:' + ((imgW - containedW) / 2).toFixed(1) + 'px;';
        wrapperStyle += 'top:' + wrapTop.toFixed(1) + 'px;';
      }
      wrapperStyle += 'transform-origin:center center;';
      if (transforms) wrapperStyle += 'transform:' + transforms + ';';
      if (s.border) wrapperStyle += 'outline:1px solid #000;outline-offset:-1px;';
      // Image fills wrapper
      var imgStyle = 'width:100%;height:100%;object-fit:' + fit + ';' + filt;
      inner = '<div style="' + wrapperStyle + '"><img src="' + src + '" style="' + imgStyle + '"></div>';
      if (s.number) inner += '<div class="slot-num">' + (pi * getPerPage(s) + i + 1) + '</div>';
      if (s.copyBadge && f.copies > 1 && copySeq && copyTotal[f.id]) {
        var gIdx = pi * getPerPage(s) + i;
        inner += '<div class="slot-copy-badge">' + copySeq[gIdx] + '/' + copyTotal[f.id] + '</div>';
      }
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
    } else if (f && f._loading) {
      inner = '<div class="slot-empty"><span class="plus-icon" style="font-size:14px;color:var(--text-muted)">加载中…</span></div>';
      html += '<div class="invoice-slot' + selClass + '" data-slot-idx="' + i + '" style="position:absolute;left:' + imgX + 'px;top:' + imgY + 'px;width:' + imgW + 'px;height:' + imgH + 'px">' + inner + '</div>';
    } else if (f && f._placeholder) {
      // 空白占位：点击上传可替换该占位（保留其他留白）
      inner = '<div class="slot-empty slot-blank" onclick="addFileToSlot(' + i + ')"><span class="plus-icon">＋</span><span>空白</span></div>';
      html += '<div class="invoice-slot' + selClass + '" data-slot-idx="' + i + '" style="position:absolute;left:' + imgX + 'px;top:' + imgY + 'px;width:' + imgW + 'px;height:' + imgH + 'px">' + inner + '</div>';
    } else {
      inner = '<div class="slot-empty" onclick="addFileToSlot(' + i + ')"><span class="plus-icon">＋</span><span>点击添加发票</span></div>';
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

  // 粘贴单：装订线 + 右下角签字栏（页面级 overlay，与 slot 同坐标系）
  html += buildPasteSheetHtml(s, MM2PX * scale, 'px');

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

var _dragHintShown = false; // 本次会话是否已提示过拖拽手势

var _slotTempDrag = null; // 尾部空槽拖拽的临时占位对象（未移动则销毁）

var _slotSuppressClick = false; // 槽位拖拽松手后吞掉浏览器合成的 click

var DRAG_THRESHOLD_PX = 4; // 位移阈值：小于它不算拖拽，避免微小抖动吞掉 click

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
  // 拖拽松手后浏览器会向空槽派发合成 click，capture 阶段吞掉防止误触上传
  document.addEventListener('click', function(e) {
    if (!_slotSuppressClick) return;
    _slotSuppressClick = false;
    e.stopPropagation();
    e.preventDefault();
  }, true);
  // Click on empty area deselects (floating slot toolbar excluded, issue #33)
  document.getElementById('previewWrap').addEventListener('mousedown', function(e) {
    if (!e.target.closest('.invoice-slot') && !e.target.closest('.slot-handle') && !e.target.closest('#slotToolbar') && !e.target.closest('#slotCopyMenu')) {
      selectSlot(-1);
    }
  });
}

// 尾部空槽按下即拖：临时创建空白占位进入拖拽链路，松手落到实体之间 = 中间留白，未移动则销毁
function insertTempPlaceholder() {
  var temp = createFileObj({ name: '空白占位', _placeholder: true });
  var active = getActiveFiles();
  if (document.getElementById('pageOrder').value === 'reverse') {
    // 倒序渲染：显示序末尾 = 底层数组 active 区头部
    var insertAt = 0;
    for (var i = 0; i < S.files.length; i++) {
      var f = S.files[i];
      if (f.checked || f._placeholder) { insertAt = i; break; }
    }
    S.files.splice(insertAt, 0, temp);
  } else {
    // 正序：插到最后一个 active 文件之后
    var n = active.length;
    var last = n > 0 ? active[n - 1] : null;
    var iLast = last ? S.files.indexOf(last) : -1;
    S.files.splice(iLast >= 0 ? iLast + 1 : S.files.length, 0, temp);
  }
  return temp;
}

function onSlotMouseDown(e) {
  if (e.button !== 0) return;
  var slotEl = e.target.closest('.invoice-slot');
  if (!slotEl) return;

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

  var files = getActiveFiles();
  var settings = getSettings();
  var layout = calculateLayout(settings);
  var perPage = getPerPage(settings);
  var fileIdx = S.currentPage * perPage + idx;
  var f = fileIdx < files.length ? files[fileIdx] : null;

  // 尾部第一个空槽按下即拖：占位延迟到确认拖拽后再落 —— 若在 mousedown 就插占位，
  // 纯点击会在 mouseup 销毁占位并重建预览 DOM，浏览器随之不再派发 click（issue #37②）
  var pendingTail = false;
  if (slotEl.querySelector('.slot-empty')) {
    if (f && f._loading) return;
    if (!f) {
      if (fileIdx !== files.length) return;
      pendingTail = true;
    }
    // 空白占位（slot-blank）：与常规文件一样可拖拽排序，click 仍触发上传
  }
  if (!f && !pendingTail) return;

  // 选中票面后由浮动工具条设置份数（issue #60：双击入口已移除——首次点击会重建预览 DOM，
  // 第二下常落到重建后的元素或压在票面上的工具条里，手势判定无法稳定命中原票面）

  // Otherwise: click to select + drag to move
  // 尾部空槽的纯点击不进拖拽：不 preventDefault、不加 dragging，保证 click 正常派发
  if (!pendingTail) e.preventDefault();
  if (!pendingTail) selectSlot(idx);

  _slotDrag = {
    mode: 'move',
    slotEl: slotEl,
    wrapperEl: slotEl.querySelector(':scope > div'),
    fileObj: f,
    tempPlaceholder: null,
    pendingTail: pendingTail,
    idx: idx,
    startX: e.clientX,
    startY: e.clientY,
    startOffX: f ? (f.slotOffsetX || 0) : 0,
    startOffY: f ? (f.slotOffsetY || 0) : 0,
    previewScale: getCurrentPreviewScale(),
    // Cache settings/layout for perf (avoid getSettings() every mousemove)
    cachedSettings: settings,
    cachedLayout: layout,
    // 主轴方向：单列（含报销单分段）取上下边缘，多列取左右边缘
    vertical: !!layout.reimburse || settings.cols === 1,
    perPage: perPage,
    activeLen: files.length,
    dropEl: null,
    dropIdx: -1,
    dropZone: ''
  };
  if (!pendingTail) slotEl.classList.add('dragging');

  document.addEventListener('mousemove', onSlotMouseMove);
  document.addEventListener('mouseup', onSlotMouseUp);
}

function startResize(e, idx, slotEl, corner) {
  selectSlot(idx);

  var files = getActiveFiles();
  var settings = getSettings();
  var layout = calculateLayout(settings);
  var perPage = getPerPage(settings);
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
  // 位移不够不算拖拽：避免微小手抖把纯点击误判成拖拽，吞掉浏览器合成的 click
  var totalDx = e.clientX - _slotDrag.startX;
  var totalDy = e.clientY - _slotDrag.startY;
  if (!_slotDrag.moved && Math.hypot(totalDx, totalDy) < DRAG_THRESHOLD_PX) return;
  _slotDrag.moved = true;  // Track actual mouse movement

  // 尾部空槽：确认是拖拽后才落临时占位，进入排序链路
  if (_slotDrag.pendingTail) {
    _slotDrag.pendingTail = false;
    var tailTemp = insertTempPlaceholder();
    _slotDrag.tempPlaceholder = tailTemp;
    _slotDrag.fileObj = tailTemp;
    _slotDrag.activeLen = getActiveFiles().length;
    _slotDrag.slotEl.classList.add('dragging');
  }

  var settings = _slotDrag.cachedSettings;
  var layout = _slotDrag.cachedLayout;

  if (_slotDrag.mode === 'move') {
    updateDropTarget(e);
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
    var s = _slotDrag.cachedSettings;
    var _clampDims = getObjDims(f, s);
    var imgW = _clampDims.w;
    var imgH = _clampDims.h;
    // 旋转 90°/270° 时约束范围按旋转后的视觉尺寸计算（与 renderPage 一致）
    var clampRot = getRotation(f, slot, s);
    var clampRot90 = (clampRot === 90 || clampRot === 270);
    var fitW = clampRot90 ? imgH : imgW;
    var fitH = clampRot90 ? imgW : imgH;
    var displayW, displayH;
    if (s.fitMode === 'fill') {
      displayW = (clampRot90 ? slot.h : slot.w) * (f.slotScale || 1);
      displayH = (clampRot90 ? slot.w : slot.h) * (f.slotScale || 1);
    } else if (s.fitMode === 'original') {
      displayW = fitW * (f.slotScale || 1);
      displayH = fitH * (f.slotScale || 1);
    } else {
      var fitScale = Math.min(slot.w / fitW, slot.h / fitH);
      var perScale = f.slotScale || 1;
      displayW = fitW * fitScale * perScale;
      displayH = fitH * fitScale * perScale;
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
    var newScale = Math.max(0.2, Math.min(5.0, _slotDrag.startScale * ratio));
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
  var d = _slotDrag;
  _slotDrag = null;
  document.removeEventListener('mousemove', onSlotMouseMove);
  document.removeEventListener('mouseup', onSlotMouseUp);

  // 空槽 mousedown 已 preventDefault 不会触发 click；但拖拽路径上 wrapper 等子元素
  // 仍可能派发合成 click（如占位槽 onclick=addFileToSlot），统一吞一次
  if (d.moved) {
    _slotSuppressClick = true;
    setTimeout(function() { _slotSuppressClick = false; }, 0);
  }

  if (d.mode === 'move' && d.dropIdx >= 0 && d.dropIdx !== d.idx) {
    clearDropTarget(d.dropEl);
    if (d.tempPlaceholder && d.dropZone === 'tail') {
      // 临时占位拖到尾部空槽 = 原位，视为取消，销毁占位
      var ti = S.files.indexOf(d.tempPlaceholder);
      if (ti >= 0) S.files.splice(ti, 1);
      updatePreview();
      return;
    }
    // 跨槽拖拽是排序手势，撤销拖动过程中产生的单票偏移
    d.fileObj.slotOffsetX = d.startOffX;
    d.fileObj.slotOffsetY = d.startOffY;
    if (d.dropZone === 'before' || d.dropZone === 'after') {
      moveSlotInvoice(d.idx, d.dropIdx, d.dropZone);
    } else if (d.dropZone === 'tail') {
      moveSlotInvoiceToTail(d.idx);
    } else {
      swapSlotInvoices(d.idx, d.dropIdx);
    }
    if (d.tempPlaceholder) toast('已在中间留白，点击可上传发票');
    return;
  }
  clearDropTarget(d.dropEl);
  if (d.tempPlaceholder) {
    // 未产生有效落点：临时占位用完即销毁，版面恢复原状
    var ti2 = S.files.indexOf(d.tempPlaceholder);
    if (ti2 >= 0) S.files.splice(ti2, 1);
    updatePreview();
    return;
  }
  if (d.moved) {
    updatePreview();
    updateAdjPanel();
  }
}

// 光标落在槽位间空白时，吸附到距离最近的槽位（gap 盲区兜底，80px 内有效）
function findNearestSlot(x, y) {
  var best = null, bestD = Infinity;
  var slots = document.querySelectorAll('.invoice-slot[data-slot-idx]');
  for (var i = 0; i < slots.length; i++) {
    var r = slots[i].getBoundingClientRect();
    var dx = Math.max(r.left - x, 0, x - r.right);
    var dy = Math.max(r.top - y, 0, y - r.bottom);
    var d = dx * dx + dy * dy;
    if (d < bestD) { bestD = d; best = slots[i]; }
  }
  if (!best || bestD > 6400) return { el: null, idx: -1 };
  var idx = parseInt(best.dataset.slotIdx);
  if (isNaN(idx)) idx = -1;
  return { el: best, idx: idx };
}

// 判定槽位能否作为落点，并按光标在槽内的主轴位置分区
function resolveSlotTarget(el, e) {
  var idx = el ? parseInt(el.dataset.slotIdx) : -1;
  if (isNaN(idx)) idx = -1;
  if (!el || idx < 0 || idx === _slotDrag.idx) return { el: null, idx: -1, zone: '' };
  var di = S.currentPage * _slotDrag.perPage + idx;
  var r = el.getBoundingClientRect();
  var ratio = _slotDrag.vertical
    ? (e.clientY - r.top) / (r.height || 1)
    : (e.clientX - r.left) / (r.width || 1);
  // 尾部空槽不接收；紧邻已装发票的第一个空槽整个作为「移到末尾」落点
  if (di >= _slotDrag.activeLen) {
    if (di !== _slotDrag.activeLen) return { el: null, idx: -1, zone: '' };
    return { el: el, idx: idx, zone: 'tail' };
  }
  // 落点分区：主轴两端各 25% 为顺位插入，中间 50% 为对调
  var zone = ratio < 0.25 ? 'before' : ratio > 0.75 ? 'after' : 'swap';
  return { el: el, idx: idx, zone: zone };
}

// Track the slot under cursor while dragging; cursor zone decides insert vs swap
function updateDropTarget(e) {
  var el = document.elementFromPoint(e.clientX, e.clientY);
  el = el && el.closest ? el.closest('.invoice-slot') : null;
  var target = resolveSlotTarget(el, e);
  // 槽位之间的空白间隙吸附到最近的槽位，避免插入指示时有时无
  if (!target.el) {
    var near = findNearestSlot(e.clientX, e.clientY);
    if (near.el && near.el !== el) target = resolveSlotTarget(near.el, e);
  }
  if (target.el) showDragHint();
  if (target.el === _slotDrag.dropEl && target.zone === _slotDrag.dropZone) return;
  clearDropTarget(_slotDrag.dropEl);
  _slotDrag.dropEl = target.el;
  _slotDrag.dropIdx = target.idx;
  _slotDrag.dropZone = target.zone;
  if (!target.el) return;
  if (target.zone === 'swap') {
    target.el.classList.add('drop-target');
  } else {
    target.el.classList.add('drop-insert');
    target.el.classList.add(_slotDrag.vertical ? 'drop-axis-v' : 'drop-axis-h');
    target.el.classList.add(target.zone === 'before' ? 'drop-at-start' : 'drop-at-end');
    if (target.zone === 'tail') target.el.classList.add('drop-tail');
  }
}

function clearDropTarget(el) {
  if (!el) return;
  el.classList.remove('drop-target', 'drop-insert', 'drop-axis-v', 'drop-axis-h', 'drop-at-start', 'drop-at-end', 'drop-tail');
}

// 首次跨槽拖拽时提示两种手势的区别
function showDragHint() {
  if (_dragHintShown) return;
  _dragHintShown = true;
  try {
    if (localStorage.getItem('ticketchan-drag-hint')) return;
    localStorage.setItem('ticketchan-drag-hint', '1');
  } catch (err) { return; }
  toast('拖到发票边缘 = 顺位插入，拖到中间 = 两张对调', 4000);
}

// Swap invoice positions between two slots (slot idx → active array index)
function swapSlotInvoices(fromIdx, toIdx) {
  var settings = getSettings();
  var perPage = getPerPage(settings);
  var fromDi = S.currentPage * perPage + fromIdx;
  var toDi = S.currentPage * perPage + toIdx;
  var active = getActiveFiles();
  var n = active.length;
  if (fromDi < 0 || toDi < 0 || fromDi >= n || toDi >= n || fromDi === toDi) {
    updatePreview();
    return;
  }
  var a = active[fromDi], b = active[toDi];
  if (a === b) { updatePreview(); return; }
  var ia = S.files.indexOf(a), ib = S.files.indexOf(b);
  if (ia < 0 || ib < 0) { updatePreview(); return; }
  var offA = a.slotOffsetX, offB = b.slotOffsetX;
  var offAY = a.slotOffsetY, offBY = b.slotOffsetY;
  S.files[ia] = b; S.files[ib] = a;
  a.slotOffsetX = offA; a.slotOffsetY = offAY;
  b.slotOffsetX = offB; b.slotOffsetY = offBY;
  if (_activeFileIdx === ia) { _activeFileIdx = ib; }
  else if (_activeFileIdx === ib) { _activeFileIdx = ia; }
  renderFileList();
  toast('已互换位置');
  updatePreview();
}

// Move invoice into a neighbor position (insert before/after instead of swapping)
function moveSlotInvoice(fromIdx, toIdx, zone) {
  var settings = getSettings();
  var perPage = getPerPage(settings);
  var fromDi = S.currentPage * perPage + fromIdx;
  var toDi = S.currentPage * perPage + toIdx;
  var active = getActiveFiles();
  var n = active.length;
  if (fromDi < 0 || toDi < 0 || fromDi >= n || toDi >= n || fromDi === toDi) {
    updatePreview();
    return;
  }
  var a = active[fromDi], b = active[toDi];
  if (a === b) { updatePreview(); return; }
  var ia = S.files.indexOf(a), ib = S.files.indexOf(b);
  if (ia < 0 || ib < 0) { updatePreview(); return; }
  var after = (zone === 'after');
  // 倒序渲染时 S.files 尾部先显示，显示序「插到前面」等于底层数组「插到后面」
  if (document.getElementById('pageOrder').value === 'reverse') after = !after;
  S.files.splice(ia, 1);
  var ibNew = S.files.indexOf(b);
  var insertAt = after ? ibNew + 1 : ibNew;
  S.files.splice(insertAt, 0, a);
  // 落点就是原位（相邻槽位往自己那侧插），顺序没变
  if (insertAt === ia) { updatePreview(); return; }
  _activeFileIdx = S.files.indexOf(a);
  // 选中态跟随移动后的槽位（可能跨页）
  var newDi = getActiveFiles().indexOf(a);
  if (newDi >= 0) {
    S.currentPage = Math.floor(newDi / perPage);
    S.selectedSlot = newDi % perPage;
  }
  renderFileList();
  toast('已调整顺序');
  updatePreview();
}

// Move invoice to the end of the active sequence (drop on the first blank slot after the last invoice)
function moveSlotInvoiceToTail(fromIdx) {
  var settings = getSettings();
  var perPage = getPerPage(settings);
  var fromDi = S.currentPage * perPage + fromIdx;
  var active = getActiveFiles();
  var n = active.length;
  if (fromDi < 0 || fromDi >= n - 1) { updatePreview(); return; }
  var a = active[fromDi];
  var ia = S.files.indexOf(a);
  if (ia < 0) { updatePreview(); return; }
  S.files.splice(ia, 1);
  var insertAt;
  if (document.getElementById('pageOrder').value === 'reverse') {
    // 倒序渲染：显示序末尾 = 底层数组 active 区头部，找最后一个未勾选文件之前的空隙
    insertAt = 0;
    for (var i = 0; i < S.files.length; i++) {
      var f = S.files[i];
      if (f.checked || f._placeholder) { insertAt = i; break; }
    }
    S.files.splice(insertAt, 0, a);
  } else {
    // 正序：插到最后一个 active 文件之后
    var last = active[n - 1];
    var iLast = last === a ? -1 : S.files.indexOf(last);
    insertAt = iLast >= 0 ? iLast + 1 : S.files.length;
    S.files.splice(insertAt, 0, a);
  }
  _activeFileIdx = S.files.indexOf(a);
  var newDi = getActiveFiles().indexOf(a);
  if (newDi >= 0) {
    S.currentPage = Math.floor(newDi / perPage);
    S.selectedSlot = newDi % perPage;
  }
  renderFileList();
  toast('已移到末尾');
  updatePreview();
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
