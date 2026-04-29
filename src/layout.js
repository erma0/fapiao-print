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
      if (s.fitMode === 'custom' && s.customScale !== 1) transforms += 'scale(' + s.customScale + ') ';
      if (rot) transforms += 'rotate(' + rot + 'deg) ';
      inner = '<img src="' + src + '" style="' + (s.fitMode === 'original' ? '' : 'max-width:100%;max-height:100%;') + 'object-fit:' + fit + ';' + filt + (transforms ? 'transform:' + transforms + ';' : '') + '">';
      if (s.number) inner += '<div class="slot-num">' + (pi * s.rows * s.cols + i + 1) + '</div>';
      if (s.watermark && s.watermarkText) {
        var ws = Math.min(slot.w * scale, slot.h * scale) * 0.15;
        inner += '<div class="watermark" style="color:' + s.watermarkColor + ';opacity:' + s.watermarkOpacity + ';font-size:' + ws + 'px;transform:translate(-50%,-50%) rotate(' + s.watermarkAngle + 'deg);top:50%;left:50%">' + s.watermarkText + '</div>';
      }
      html += '<div class="invoice-slot" style="position:absolute;left:' + imgX + 'px;top:' + imgY + 'px;width:' + imgW + 'px;height:' + imgH + 'px;' + bdr + '">' + inner + '</div>';
    } else {
      inner = '<div class="slot-empty">空</div>';
      html += '<div class="invoice-slot" style="position:absolute;left:' + imgX + 'px;top:' + imgY + 'px;width:' + imgW + 'px;height:' + imgH + 'px">' + inner + '</div>';
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
}

// =====================================================
// Canvas Rendering (for PDF generation / print)
// =====================================================

function renderPageToCanvas(pageFiles, s) {
  var pxPerMm = PDF_RENDER_DPI / 25.4;
  var layout = calculateLayout(s, pxPerMm);

  var canvas = document.getElementById('renderCanvas');
  canvas.width = Math.round(layout.pw);
  canvas.height = Math.round(layout.ph);
  var ctx = canvas.getContext('2d');
  ctx.fillStyle = 'white';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  for (var i = 0; i < layout.slots.length; i++) {
    var slot = layout.slots[i];
    var f = pageFiles ? pageFiles[i] : null;

    if (f && f.img) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(slot.x, slot.y, slot.w, slot.h);
      ctx.clip();

      var rot = getRotation(f, slot, s);

      var imgW = f.ow || f.img.naturalWidth;
      var imgH = f.oh || f.img.naturalHeight;
      var drawW, drawH;
      var eW = (rot === 90 || rot === 270) ? imgH : imgW;
      var eH = (rot === 90 || rot === 270) ? imgW : imgH;

      if (s.fitMode === 'fill') { drawW = slot.w; drawH = slot.h; }
      else if (s.fitMode === 'original') { drawW = eW; drawH = eH; }
      else if (s.fitMode === 'custom') {
        var sc = Math.min(slot.w / eW, slot.h / eH) * (s.customScale || 1);
        drawW = eW * sc; drawH = eH * sc;
      }
      else { var sc = Math.min(slot.w / eW, slot.h / eH); drawW = eW * sc; drawH = eH * sc; }

      var drawX = slot.x + (slot.w - drawW) / 2;
      var drawY = slot.y + (slot.h - drawH) / 2;

      if (rot) {
        var cx = slot.x + slot.w / 2, cy = slot.y + slot.h / 2;
        ctx.translate(cx, cy);
        ctx.rotate(rot * Math.PI / 180);
        ctx.translate(-cx, -cy);
      }

      ctx.drawImage(f.img, drawX, drawY, drawW, drawH);

      // Apply grayscale/bw
      if (s.colorMode === 'grayscale' || s.colorMode === 'bw') {
        var imgData = ctx.getImageData(slot.x, slot.y, Math.ceil(slot.w), Math.ceil(slot.h));
        var data = imgData.data;
        for (var j = 0; j < data.length; j += 4) {
          var gray = 0.299 * data[j] + 0.587 * data[j + 1] + 0.114 * data[j + 2];
          if (s.colorMode === 'bw') { data[j] = data[j + 1] = data[j + 2] = gray > 128 ? 255 : 0; }
          else { data[j] = data[j + 1] = data[j + 2] = gray; }
        }
        ctx.putImageData(imgData, slot.x, slot.y);
      }
      ctx.restore();
    }
    if (s.border) { ctx.strokeStyle = '#000'; ctx.lineWidth = 1; ctx.strokeRect(slot.x, slot.y, slot.w, slot.h); }
    if (s.number && f) { ctx.fillStyle = 'rgba(0,0,0,0.5)'; ctx.font = '11px sans-serif'; ctx.fillText(String(i + 1), slot.x + slot.w - 18, slot.y + 14); }
  }

  // Cut lines
  if (s.cutline && (s.cols > 1 || s.rows > 1)) {
    ctx.setLineDash([6, 4]); ctx.strokeStyle = '#94a3b8'; ctx.lineWidth = 0.5;
    for (var cl = 0; cl < layout.cutLines.length; cl++) {
      var line = layout.cutLines[cl];
      ctx.beginPath();
      if (line.type === 'horizontal') {
        ctx.moveTo(0, line.pos);
        ctx.lineTo(canvas.width, line.pos);
      } else {
        ctx.moveTo(line.pos, 0);
        ctx.lineTo(line.pos, canvas.height);
      }
      ctx.stroke();
    }
    ctx.setLineDash([]);
  }

  // Watermark
  if (s.watermark && s.watermarkText) {
    ctx.save();
    var fontSize = Math.min(canvas.width, canvas.height) * 0.08;
    ctx.font = '900 ' + fontSize + 'px sans-serif';
    ctx.fillStyle = s.watermarkColor;
    ctx.globalAlpha = s.watermarkOpacity;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate(s.watermarkAngle * Math.PI / 180);
    ctx.fillText(s.watermarkText, 0, 0);
    ctx.restore();
  }

  return canvas.toDataURL('image/png');
}
