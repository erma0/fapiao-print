// =====================================================
// OFD Client — Pure frontend OFD parser & SVG renderer
// =====================================================
// Ported from invoice-engine/src/lib.rs (Rust)
// ZIP: JSZip | XML: DOMParser | SVG: string concat (same as Rust)
//
// Produces SVG + invoice info from OFD (Open Fixed-layout Document) files.
// OFD = ZIP containing XML page descriptions + image resources (GB/T 33190-2016).

var JSZip = window.JSZip;

var OFD_SVG_SCALE = 3.5;

// =====================================================
// Helpers
// =====================================================

// _stripXmlNs / _parseXml are defined in xml-utils.js (loaded before this file)

function _attr(el, name) {
  if (!el || !el.getAttribute) return null;
  return el.getAttribute(name);
}

function _parseF2(s) {
  if (!s) return null;
  var parts = s.trim().split(/\s+/);
  if (parts.length >= 2) return [parseFloat(parts[0]), parseFloat(parts[1])];
  return null;
}

function _parseF4(s) {
  if (!s) return null;
  var parts = s.trim().split(/\s+/);
  if (parts.length >= 4) return [parseFloat(parts[0]), parseFloat(parts[1]), parseFloat(parts[2]), parseFloat(parts[3])];
  return null;
}

function _parseF6(s) {
  if (!s) return null;
  var parts = s.trim().split(/\s+/);
  if (parts.length >= 6) return [parseFloat(parts[0]), parseFloat(parts[1]), parseFloat(parts[2]), parseFloat(parts[3]), parseFloat(parts[4]), parseFloat(parts[5])];
  return null;
}

function _parseColor(s) {
  if (!s) return null;
  var parts = s.trim().split(/\s+/);
  if (parts.length >= 3) return [parseInt(parts[0]) & 0xFF, parseInt(parts[1]) & 0xFF, parseInt(parts[2]) & 0xFF];
  return null;
}

function _parseDeltaX(s) {
  if (!s) return [];
  var tokens = s.trim().split(/\s+/);
  var result = [];
  var i = 0;
  while (i < tokens.length) {
    if (tokens[i] === 'g' && i + 2 < tokens.length) {
      // Group format: g count value [extra_value...]
      var count = parseInt(tokens[i + 1]);
      var val = parseFloat(tokens[i + 2]);
      if (!isNaN(count) && !isNaN(val)) {
        for (var gi = 0; gi < count; gi++) result.push(val);
        i += 3;
        // Check if there's an extra value after the group
        if (i < tokens.length && tokens[i] !== 'g') {
          var extra = parseFloat(tokens[i]);
          if (!isNaN(extra)) {
            result.push(extra);
            i += 1;
          }
        }
      } else {
        i += 1;
      }
    } else {
      var v = parseFloat(tokens[i]);
      if (!isNaN(v)) result.push(v);
      i += 1;
    }
  }
  return result;
}

function _normalizeFontName(raw) {
  if (!raw) return 'SimSun';
  var base = raw;
  var plusIdx = base.indexOf('+');
  if (plusIdx >= 0) base = base.substring(plusIdx + 1);
  var dashIdx = base.indexOf('-');
  if (dashIdx >= 0) base = base.substring(0, dashIdx);
  switch (base) {
    case 'CourierNewPSMT': return 'Courier New';
    case 'TimesNewRomanPSMT': return 'Times New Roman';
    case 'ArialMT': case 'Arial-BoldMT': return 'Arial';
    case 'SimSun': case 'STSong': return '宋体';
    case 'KaiTi': case 'STKaiti': return '楷体';
    case 'SimHei': case 'STHeiti': return '黑体';
    case 'FangSong': case 'STFangsong': return '仿宋';
    default: return base;
  }
}

function _fontFamilyCSS(normalized) {
  switch (normalized) {
    case '楷体': case 'KaiTi': case 'STKaiti': return "楷体, KaiTi, STKaiti, serif";
    case '宋体': case 'SimSun': case 'STSong': return "宋体, SimSun, STSong, serif";
    case '黑体': case 'SimHei': case 'STHeiti': return "黑体, SimHei, STHeiti, sans-serif";
    case '仿宋': case 'FangSong': case 'STFangsong': return "仿宋, FangSong, STFangsong, serif";
    case 'Courier New': return "'Courier New', Courier, monospace";
    case 'Times New Roman': return "'Times New Roman', Times, serif";
    default: return normalized;
  }
}

function _fillAttr(color, alpha) {
  if (color) {
    var a = alpha != null ? (alpha / 255).toFixed(2) : null;
    if (a && a !== '1.00') return ' fill="rgba(' + color[0] + ',' + color[1] + ',' + color[2] + ',' + a + ')"';
    return ' fill="rgb(' + color[0] + ',' + color[1] + ',' + color[2] + ')"';
  }
  if (alpha != null) return ' fill="rgba(0,0,0,' + (alpha / 255).toFixed(2) + ')"';
  return '';
}

function _strokeAttr(color, alpha) {
  if (color) {
    var a = alpha != null ? (alpha / 255).toFixed(2) : null;
    if (a && a !== '1.00') return ' stroke="rgba(' + color[0] + ',' + color[1] + ',' + color[2] + ',' + a + ')"';
    return ' stroke="rgb(' + color[0] + ',' + color[1] + ',' + color[2] + ')"';
  }
  if (alpha != null) return ' stroke="rgba(0,0,0,' + (alpha / 255).toFixed(2) + ')"';
  return '';
}

function _escXml(s) {
  return (s == null ? '' : String(s)).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function _escXmlAttr(s) {
  return _escXml(s).replace(/"/g, '&quot;');
}

// =====================================================
// OFD AbbreviatedData → SVG path
// =====================================================

function _ofdPathToSvg(data) {
  if (!data) return '';
  var tokens = data.trim().split(/\s+/);
  var svg = '';
  var i = 0;
  while (i < tokens.length) {
    switch (tokens[i]) {
      case 'M':
        if (i + 2 < tokens.length) { svg += 'M ' + tokens[i+1] + ' ' + tokens[i+2] + ' '; i += 3; }
        else i++;
        break;
      case 'L':
        if (i + 2 < tokens.length) { svg += 'L ' + tokens[i+1] + ' ' + tokens[i+2] + ' '; i += 3; }
        else i++;
        break;
      case 'C': case 'B':
        if (i + 6 < tokens.length) { svg += 'C ' + tokens[i+1] + ' ' + tokens[i+2] + ' ' + tokens[i+3] + ' ' + tokens[i+4] + ' ' + tokens[i+5] + ' ' + tokens[i+6] + ' '; i += 7; }
        else i++;
        break;
      case 'Q':
        if (i + 4 < tokens.length) { svg += 'Q ' + tokens[i+1] + ' ' + tokens[i+2] + ' ' + tokens[i+3] + ' ' + tokens[i+4] + ' '; i += 5; }
        else i++;
        break;
      case 'A':
        if (i + 7 < tokens.length) { svg += 'A ' + tokens[i+1] + ' ' + tokens[i+2] + ' ' + tokens[i+3] + ' ' + tokens[i+4] + ' ' + tokens[i+5] + ' ' + tokens[i+6] + ' ' + tokens[i+7] + ' '; i += 8; }
        else i++;
        break;
      case 'S':
        if (i + 4 < tokens.length) { svg += 'S ' + tokens[i+1] + ' ' + tokens[i+2] + ' ' + tokens[i+3] + ' ' + tokens[i+4] + ' '; i += 5; }
        else i++;
        break;
      case 'Z': case 'z':
        svg += 'Z'; i++;
        break;
      default: i++;
    }
  }
  return svg;
}

// =====================================================
// SVG builders
// =====================================================

function _buildSvgText(t, fontMap, colorSpaces, scaleX, scaleY) {
  if (!t.text) return '';

  var font = fontMap[t.fontId];
  var fontFamilyRaw = font ? (font.familyName || font.fontName) : 'SimSun';
  var fontBase = _normalizeFontName(fontFamilyRaw);
  var fontFamily = _fontFamilyCSS(fontBase);
  var fontSize = t.size;
  var bold = t.weight >= 700 ? ' font-weight="bold"' : '';

  var text = t.text;
  var hasDelta = t.deltaX.length > 0 && text.length > 1;
  var fillStr = _fillAttr(t.fillColor || (t.layerDrawParam ? null : [0, 0, 0]), t.alpha);
  var strokeStr = t.strokeColor ? _strokeAttr(t.strokeColor, t.alpha) : '';

  // CTM transform: translate to boundary origin, apply matrix, then text at local coords
  if (t.ctm) {
    var baseX = t.textX * scaleX;
    var baseY = t.textY * scaleY;
    var content;
    if (hasDelta) {
      var chars = Array.from(text);
      content = '<tspan x="' + baseX.toFixed(4) + '">' + _escXml(chars[0]) + '</tspan>';
      var xPos = baseX;
      for (var ci = 1; ci < chars.length; ci++) {
        var dx = ci - 1 < t.deltaX.length ? t.deltaX[ci - 1] : (t.deltaX[t.deltaX.length - 1] || fontSize);
        xPos += dx * scaleX;
        content += '<tspan x="' + xPos.toFixed(4) + '">' + _escXml(chars[ci]) + '</tspan>';
      }
    } else {
      content = _escXml(text);
    }
    return '<text transform="translate(' + (t.boundary[0] * scaleX).toFixed(4) + ',' + (t.boundary[1] * scaleY).toFixed(4) + ') matrix(' +
      t.ctm[0] + ',' + t.ctm[1] + ',' + t.ctm[2] + ',' + t.ctm[3] + ',' +
      (t.ctm[4] * scaleX).toFixed(4) + ',' + (t.ctm[5] * scaleY).toFixed(4) + ')"' +
      ' x="' + baseX.toFixed(4) + '" y="' + baseY.toFixed(4) + '"' +
      ' font-family="' + _escXmlAttr(fontFamily) + '"' +
      ' font-size="' + (fontSize * scaleX).toFixed(2) + '"' + bold + fillStr + strokeStr + '>' +
      content + '</text>';
  }

  // Normal: position = Boundary + TextCode offset (absolute SVG coords, no Y-flip needed)
  var baseX = (t.boundary[0] + t.textX) * scaleX;
  var baseY = (t.boundary[1] + t.textY) * scaleY;
  var content;
  if (hasDelta) {
    var chars = Array.from(text);
    content = '<tspan x="' + baseX.toFixed(4) + '">' + _escXml(chars[0]) + '</tspan>';
    var xPos = baseX;
    for (var ci = 1; ci < chars.length; ci++) {
      var dx = ci - 1 < t.deltaX.length ? t.deltaX[ci - 1] : (t.deltaX[t.deltaX.length - 1] || fontSize);
      xPos += dx * scaleX;
      content += '<tspan x="' + xPos.toFixed(4) + '">' + _escXml(chars[ci]) + '</tspan>';
    }
  } else {
    content = _escXml(text);
  }

  return '<text x="' + baseX.toFixed(4) + '" y="' + baseY.toFixed(4) + '"' +
    ' font-family="' + _escXmlAttr(fontFamily) + '"' +
    ' font-size="' + (fontSize * scaleX).toFixed(2) + '"' + bold + fillStr + strokeStr + '>' +
    content + '</text>';
}

function _buildSvgPath(p, scale) {
  if (!p.abbreviatedData) return '';
  var svgD = _ofdPathToSvg(p.abbreviatedData);
  if (!svgD) return '';

  var tx = p.boundary[0] * scale;
  var ty = p.boundary[1] * scale;

  var attrs = ' transform="translate(' + tx.toFixed(4) + ',' + ty.toFixed(4) + ') scale(' + scale.toFixed(4) + ')"';
  attrs += ' stroke-width="' + p.lineWidth.toFixed(4) + '"';
  if (p.fill) attrs += ' fill-rule="nonzero"';
  var strokeColor = p.strokeColor || [0, 0, 0];
  attrs += _strokeAttr(strokeColor, p.alpha);
  if (p.fill) {
    if (p.fillColor) {
      attrs += _fillAttr(p.fillColor, p.alpha);
    } else {
      attrs += ' fill="none"';
    }
  } else {
    attrs += ' fill="none"';
  }

  return '<g' + attrs + '><path d="' + _escXmlAttr(svgD) + '"/></g>';
}

function _buildSvgImage(img, imageData, scale) {
  var dataUrl = imageData[img.resourceId];
  if (!dataUrl) return '';

  var ix = img.boundary[0] * scale;
  var iy = img.boundary[1] * scale;
  var iw = img.boundary[2] * scale;
  var ih = img.boundary[3] * scale;

  var attrs = ' x="' + ix.toFixed(2) + '" y="' + iy.toFixed(2) + '" width="' + iw.toFixed(2) + '" height="' + ih.toFixed(2) + '"';
  if (img.alpha != null) attrs += ' opacity="' + (img.alpha / 255).toFixed(3) + '"';

  return '<image href="' + _escXmlAttr(dataUrl) + '"' + attrs + '/>';
}

// =====================================================
// XML Parsers (using DOMParser)
// =====================================================

function _parseOfdContent(xml) {
  var textObjs = [];
  var pathObjs = [];
  var imgObjs = [];

  var doc = _parseXml(xml);
  var layers = doc.getElementsByTagName('Layer');

  for (var li = 0; li < layers.length; li++) {
    var layer = layers[li];
    var layerDp = _attr(layer, 'DrawParam');
    var layerDpId = layerDp ? parseInt(layerDp) : null;

    for (var child = layer.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      var tag = child.localName;

      if (tag === 'TextObject') {
        var t = {
          id: parseInt(_attr(child, 'ID')) || 0,
          boundary: _parseF4(_attr(child, 'Boundary')) || [0, 0, 0, 0],
          fontId: parseInt(_attr(child, 'Font')) || 0,
          size: parseFloat(_attr(child, 'Size')) || 3.175,
          ctm: _parseF6(_attr(child, 'CTM')),
          text: '',
          deltaX: [],
          textX: 0,
          textY: 0,
          fillColor: null,
          strokeColor: null,
          alpha: _attr(child, 'Alpha') ? parseInt(_attr(child, 'Alpha')) : null,
          weight: parseInt(_attr(child, 'Weight')) || 400,
          layerDrawParam: layerDpId
        };
        var tc = child.getElementsByTagName('TextCode')[0];
        if (tc) {
          t.text = (tc.textContent || '').trim();
          t.textX = parseFloat(_attr(tc, 'X')) || 0;
          t.textY = parseFloat(_attr(tc, 'Y')) || 0;
          var dxAttr = _attr(tc, 'DeltaX');
          if (dxAttr) t.deltaX = _parseDeltaX(dxAttr);
        }
        var sc = child.getElementsByTagName('StrokeColor')[0];
        if (sc) t.strokeColor = _parseColor(_attr(sc, 'Value'));
        var fc = child.getElementsByTagName('FillColor')[0];
        if (fc) t.fillColor = _parseColor(_attr(fc, 'Value'));
        textObjs.push(t);
      } else if (tag === 'PathObject') {
        var p = {
          id: parseInt(_attr(child, 'ID')) || 0,
          boundary: _parseF4(_attr(child, 'Boundary')) || [0, 0, 0, 0],
          lineWidth: parseFloat(_attr(child, 'LineWidth')) || 0.25,
          strokeColor: null,
          fillColor: null,
          fill: _attr(child, 'Fill') === 'true',
          abbreviatedData: '',
          alpha: _attr(child, 'Alpha') ? parseInt(_attr(child, 'Alpha')) : null,
          layerDrawParam: layerDpId
        };
        var ad = child.getElementsByTagName('AbbreviatedData')[0];
        if (ad) p.abbreviatedData = (ad.textContent || '').trim();
        var psc = child.getElementsByTagName('StrokeColor')[0];
        if (psc) p.strokeColor = _parseColor(_attr(psc, 'Value'));
        var pfc = child.getElementsByTagName('FillColor')[0];
        if (pfc) p.fillColor = _parseColor(_attr(pfc, 'Value'));
        pathObjs.push(p);
      } else if (tag === 'ImageObject') {
        var img = {
          id: parseInt(_attr(child, 'ID')) || 0,
          boundary: _parseF4(_attr(child, 'Boundary')) || [0, 0, 0, 0],
          resourceId: parseInt(_attr(child, 'ResourceID')) || 0,
          ctm: _parseF6(_attr(child, 'CTM')),
          alpha: _attr(child, 'Alpha') ? parseInt(_attr(child, 'Alpha')) : null,
          imageMask: _attr(child, 'ImageMask') ? parseInt(_attr(child, 'ImageMask')) : null
        };
        imgObjs.push(img);
      }
    }
  }

  return { texts: textObjs, paths: pathObjs, images: imgObjs };
}

function _parseCustomData(xml) {
  var map = {};
  var doc = _parseXml(xml);
  var items = doc.getElementsByTagName('CustomData');
  for (var i = 0; i < items.length; i++) {
    var name = _attr(items[i], 'Name');
    if (name) map[name] = (items[i].textContent || '').trim();
  }
  return map;
}

function _parseCustomTag(xml) {
  var map = {};
  var doc = _parseXml(xml);

  var fieldTags = ['InvoiceNo', 'IssueDate', 'BuyerName', 'BuyerTaxID',
    'SellerName', 'SellerTaxID', 'TaxExclusiveTotalAmount',
    'TaxTotalAmount', 'TaxInclusiveTotalAmount', 'Amount',
    'TaxAmount', 'InvoiceClerk', 'Item', 'Price', 'Quantity',
    'Note', 'TaxScheme', 'MeasurementDimension'];

  function walkField(el) {
    var field = el.localName;
    if (fieldTags.indexOf(field) < 0) return;
    var refs = el.getElementsByTagName('ObjectRef');
    for (var ri = 0; ri < refs.length; ri++) {
      var id = parseInt((refs[ri].textContent || '').trim());
      if (!isNaN(id)) {
        if (!map[field]) map[field] = [];
        map[field].push(id);
      }
    }
  }

  for (var fi = 0; fi < fieldTags.length; fi++) {
    var els = doc.getElementsByTagName(fieldTags[fi]);
    for (var ei = 0; ei < els.length; ei++) walkField(els[ei]);
  }

  return map;
}

/**
 * Parse Tags/CustomTag.xml 的 `<Item>` 分组：一个 Item 内的 ObjectRef 序列 = 一条明细。
 * 生成器会把长品名拆成多个文字对象（如「*生产生活服务*技术服务」+「费」），
 * 这些段必须按组内顺序拼成一条，不能逐段当成独立明细。
 * @returns {number[][]} 每个 Item 一组的 TextObject ID
 */
function _parseCustomTagItemGroups(xml) {
  var groups = [];
  var doc = _parseXml(xml);
  var itemEls = doc.getElementsByTagName('Item');
  for (var i = 0; i < itemEls.length; i++) {
    var refs = itemEls[i].getElementsByTagName('ObjectRef');
    var group = [];
    for (var ri = 0; ri < refs.length; ri++) {
      var id = parseInt((refs[ri].textContent || '').trim());
      if (!isNaN(id)) group.push(id);
    }
    if (group.length) groups.push(group); // 无 ObjectRef 的空组丢弃
  }
  return groups;
}

function _parseFonts(xml) {
  var fonts = {};
  var colorSpaces = {};
  var drawParams = {};

  var doc = _parseXml(xml);

  var fontEls = doc.getElementsByTagName('Font');
  for (var i = 0; i < fontEls.length; i++) {
    var id = parseInt(_attr(fontEls[i], 'ID')) || 0;
    fonts[id] = {
      id: id,
      fontName: _attr(fontEls[i], 'FontName') || '',
      familyName: _attr(fontEls[i], 'FamilyName') || ''
    };
  }

  var csEls = doc.getElementsByTagName('ColorSpace');
  for (var ci = 0; ci < csEls.length; ci++) {
    var csId = parseInt(_attr(csEls[ci], 'ID'));
    var csType = _attr(csEls[ci], 'Type');
    if (!isNaN(csId) && csType) colorSpaces[csId] = csType;
  }

  var dpEls = doc.getElementsByTagName('DrawParam');
  for (var di = 0; di < dpEls.length; di++) {
    var dpId = parseInt(_attr(dpEls[di], 'ID')) || 0;
    var dp = {
      id: dpId,
      relative: _attr(dpEls[di], 'Relative') ? parseInt(_attr(dpEls[di], 'Relative')) : null,
      lineWidth: parseFloat(_attr(dpEls[di], 'LineWidth')) || 0.25,
      strokeColor: null,
      fillColor: null
    };
    var dpSc = dpEls[di].getElementsByTagName('StrokeColor');
    if (dpSc.length > 0) dp.strokeColor = _parseColor(_attr(dpSc[0], 'Value'));
    var dpFc = dpEls[di].getElementsByTagName('FillColor');
    if (dpFc.length > 0) dp.fillColor = _parseColor(_attr(dpFc[0], 'Value'));
    drawParams[dpId] = dp;
  }

  return { fonts: fonts, colorSpaces: colorSpaces, drawParams: drawParams };
}

function _resolveDrawParam(drawParams, paramId) {
  var lw = 0.25;
  var stroke = null;
  var fill = null;
  var visited = {};
  var currentId = paramId;
  while (currentId && !visited[currentId]) {
    visited[currentId] = true;
    var dp = drawParams[currentId];
    if (!dp) break;
    if (dp.lineWidth > 0) lw = dp.lineWidth;
    if (!stroke && dp.strokeColor) stroke = dp.strokeColor;
    if (!fill && dp.fillColor) fill = dp.fillColor;
    currentId = dp.relative;
  }
  return { lineWidth: lw, strokeColor: stroke, fillColor: fill };
}

function _applyDrawParamDefaults(paths, texts, drawParams) {
  var cache = {};
  for (var pi = 0; pi < paths.length; pi++) {
    var p = paths[pi];
    if (p.layerDrawParam != null) {
      if (!cache[p.layerDrawParam]) cache[p.layerDrawParam] = _resolveDrawParam(drawParams, p.layerDrawParam);
      var dp = cache[p.layerDrawParam];
      if (!p.strokeColor) p.strokeColor = dp.strokeColor;
      if (!p.fillColor) p.fillColor = dp.fillColor;
      if (p.lineWidth === 0) p.lineWidth = dp.lineWidth;
    }
  }
  for (var ti = 0; ti < texts.length; ti++) {
    var t = texts[ti];
    if (t.layerDrawParam != null) {
      if (!cache[t.layerDrawParam]) cache[t.layerDrawParam] = _resolveDrawParam(drawParams, t.layerDrawParam);
      var dp2 = cache[t.layerDrawParam];
      if (!t.fillColor) t.fillColor = dp2.fillColor;
      if (!t.strokeColor) t.strokeColor = dp2.strokeColor;
    }
  }
}

function _parseImageResources(xml) {
  var images = {};
  var doc = _parseXml(xml);
  var mmEls = doc.getElementsByTagName('MultiMedia');
  for (var i = 0; i < mmEls.length; i++) {
    var id = parseInt(_attr(mmEls[i], 'ID'));
    var mf = mmEls[i].getElementsByTagName('MediaFile')[0];
    if (mf && !isNaN(id)) {
      images[id] = (mf.textContent || '').trim();
    }
  }
  return images;
}

function _parseAnnotations(xml) {
  var allTexts = [];
  var allImgs = [];
  var doc = _parseXml(xml);
  var annots = doc.getElementsByTagName('Annot');

  for (var ai = 0; ai < annots.length; ai++) {
    var annot = annots[ai];
    var appearances = annot.getElementsByTagName('Appearance');
    for (var apIdx = 0; apIdx < appearances.length; apIdx++) {
      var ap = appearances[apIdx];
      var apBoundary = _parseF4(_attr(ap, 'Boundary'));
      var offsetX = apBoundary ? apBoundary[0] : 0;
      var offsetY = apBoundary ? apBoundary[1] : 0;

      // Same as Rust: ensure Content>Layer wrapper exists before parsing.
      // Standard OFD Appearance already has <Content><Layer>, but non-standard
      // files may place TextObject/ImageObject directly inside Appearance.
      var apXml = new XMLSerializer().serializeToString(ap);
      var apDoc = _parseXml(apXml);
      var layers = apDoc.getElementsByTagName('Layer');
      var content;
      if (layers.length > 0) {
        content = _parseOfdContent(apXml);
      } else {
        // Wrap in Content>Layer so _parseOfdContent can find objects
        var innerXml = _stripXmlNs(apXml).replace(/<\?[^?]*\?>/g, '');
        content = _parseOfdContent('<Content><Layer>' + innerXml + '</Layer></Content>');
      }
      for (var ci = 0; ci < content.texts.length; ci++) {
        var t = content.texts[ci];
        t.boundary[0] += offsetX;
        t.boundary[1] += offsetY;
        allTexts.push(t);
      }
      for (var ii = 0; ii < content.images.length; ii++) {
        var img = content.images[ii];
        img.boundary[0] += offsetX;
        img.boundary[1] += offsetY;
        allImgs.push(img);
      }
    }
  }

  return { texts: allTexts, images: allImgs };
}

// =====================================================
// Invoice Data Extraction
// =====================================================

function _extractFromCustomData(customData) {
  var info = {};
  // Skip empty strings (self-closing <CustomData Name="xxx"/> with no value)
  function getCustom(key) {
    var v = customData[key];
    if (!v || !v.trim()) return null;
    return v.trim();
  }
  if (getCustom('发票号码')) info.invoiceNo = getCustom('发票号码');
  if (getCustom('开票日期')) info.invoiceDate = getCustom('开票日期');
  if (getCustom('购买方名称')) info.buyerName = getCustom('购买方名称');
  if (getCustom('购买方纳税人识别号')) info.buyerTaxId = getCustom('购买方纳税人识别号');
  if (getCustom('销售方名称')) info.sellerName = getCustom('销售方名称');
  if (getCustom('销售方纳税人识别号')) info.sellerTaxId = getCustom('销售方纳税人识别号');
  if (customData['合计金额']) info.amountNoTax = parseFloat(customData['合计金额']) || null;
  if (customData['合计税额']) info.taxAmount = parseFloat(customData['合计税额']) || null;
  // Compute total = no_tax + tax (same as Rust version)
  if (info.amountNoTax != null && info.taxAmount != null) {
    info.amountTax = Math.round((info.amountNoTax + info.taxAmount) * 100) / 100;
  }
  // Only use 价税合计 as fallback if we couldn't compute from 合计金额+合计税额
  if (info.amountTax == null && customData['价税合计']) {
    info.amountTax = parseFloat(customData['价税合计']) || null;
  }
  return info;
}

function _extractFromCustomTag(customTag, allTexts, itemGroups, tplTexts) {
  var info = {};
  var idToText = {};
  for (var i = 0; i < allTexts.length; i++) {
    idToText[allTexts[i].id] = allTexts[i].text;
  }
  // 项目名称 / 备注的 ID 查找与 Rust 11c 的 by_id 一致：数据层优先，模板层兜底
  var idToTextAll = {};
  var tpl = tplTexts || [];
  for (var ti = 0; ti < tpl.length; ti++) idToTextAll[tpl[ti].id] = tpl[ti].text;
  for (var pi = 0; pi < allTexts.length; pi++) idToTextAll[allTexts[pi].id] = allTexts[pi].text;

  function textById(map, id) {
    var t = map[id];
    return t == null ? null : String(t).trim();
  }

  function firstText(field) {
    var ids = customTag[field];
    if (!ids || !ids.length) return null;
    for (var j = 0; j < ids.length; j++) {
      var t = idToText[ids[j]];
      if (t) return t;
    }
    return null;
  }

  if (!info.invoiceNo) info.invoiceNo = firstText('InvoiceNo');
  if (!info.invoiceDate) info.invoiceDate = firstText('IssueDate');
  if (!info.buyerName) info.buyerName = firstText('BuyerName');
  if (!info.buyerTaxId) info.buyerTaxId = firstText('BuyerTaxID');
  if (!info.sellerName) info.sellerName = firstText('SellerName');
  if (!info.sellerTaxId) info.sellerTaxId = firstText('SellerTaxID');
  if (!info.amountNoTax) {
    var val = firstText('TaxExclusiveTotalAmount') || firstText('Amount');
    if (val) info.amountNoTax = parseFloat(val) || null;
  }
  if (!info.taxAmount) {
    var val2 = firstText('TaxTotalAmount') || firstText('TaxAmount');
    if (val2) info.taxAmount = parseFloat(val2) || null;
  }
  if (!info.amountTax) {
    var val3 = firstText('TaxInclusiveTotalAmount');
    if (val3) info.amountTax = parseFloat(val3) || null;
  }

  // 项目名称：每个 <Item> 组 = 一条明细，组内多段按顺序拼成完整品名
  if (itemGroups && itemGroups.length) {
    var items = [];
    for (var gi = 0; gi < itemGroups.length; gi++) {
      var joined = '';
      for (var gk = 0; gk < itemGroups[gi].length; gk++) {
        var seg = textById(idToTextAll, itemGroups[gi][gk]);
        if (seg) joined += seg;
      }
      items.push(_cleanItemName(joined));
    }
    var itemName = _joinItemNames(items);
    if (itemName) info.itemName = itemName;
  } else if (customTag['Item'] && customTag['Item'].length) {
    // 无分组信息时退回扁平处理（各 ObjectRef 视为独立明细）
    var flatItems = [];
    for (var fj = 0; fj < customTag['Item'].length; fj++) {
      var ft = textById(idToTextAll, customTag['Item'][fj]);
      if (ft) flatItems.push(_cleanItemName(ft));
    }
    var itemName2 = _joinItemNames(flatItems);
    if (itemName2) info.itemName = itemName2;
  }

  // 备注：Note 各段按序拼接，剔除以落款人（复核人/收款人/开票人）开头的行
  var noteIds = customTag['Note'];
  if (noteIds && noteIds.length) {
    var lines = [];
    for (var ni = 0; ni < noteIds.length; ni++) {
      var nt = textById(idToTextAll, noteIds[ni]);
      if (!nt) continue;
      if (/^(?:复核人|收款人|开票人)/.test(nt.replace(/\s+/g, ''))) continue;
      lines.push(nt);
    }
    var remark = _cleanRemark(lines.join(' '));
    if (remark) info.remark = remark;
  }

  return info;
}

/**
 * 多条项目名称去重后用「、」连接（对齐 Rust join_item_names，总长上限 ITEM_NAME_MAX）。
 * pdf-text.js 的 _joinItems 只做截断不去重，这里补上去重。
 */
function _joinItemNames(items) {
  var uniq = [];
  for (var i = 0; i < items.length; i++) {
    var s = (items[i] || '').trim();
    if (s && uniq.indexOf(s) < 0) uniq.push(s);
  }
  return uniq.length ? _joinItems(uniq) : '';
}

/**
 * 发票类型判定（OFD 文本层）。
 * 逐条文本匹配，避免跨文本拼接出来的假关键词；「普通」优先于「专用」——
 * 票面其它位置出现「专用」字样或文字层噪声时，不会被误判成专票。
 * @param {string[]} texts
 * @returns {'增值税普通发票'|'增值税专用发票'|'电子发票'|''}
 */
function _detectInvoiceTypeTexts(texts) {
  var general = false, special = false, electronic = false;
  for (var i = 0; i < texts.length; i++) {
    var s = String(texts[i] || '').replace(/\s/g, '');
    if (s.indexOf('普通发票') >= 0 || s.indexOf('增值税普通') >= 0 || s.indexOf('电子普通') >= 0) general = true;
    if (s.indexOf('专用发票') >= 0 || s.indexOf('增值税专用') >= 0) special = true;
    if (s.indexOf('电子发票') >= 0) electronic = true;
  }
  if (general) return '增值税普通发票';
  if (special) return '增值税专用发票';
  if (electronic) return '电子发票';
  return '';
}

function _extractInvoiceFromText(texts) {
  var info = {};
  var section = '';
  var nameCount = 0;
  var taxidCount = 0;
  var foundJiashui = false;
  var foundXiaoxie = false;
  var foundHeji = false;

  var compositeTexts = [];
  var charBuf = '';
  for (var i = 0; i < texts.length; i++) {
    var text = (texts[i].text || '').trim();
    if (!text) continue;
    if (text.length === 1) {
      charBuf += text;
    } else {
      if (charBuf.length >= 2) compositeTexts.push(charBuf);
      charBuf = '';
      compositeTexts.push(text);
    }
  }
  if (charBuf.length >= 2) compositeTexts.push(charBuf);

  function valueMatches(val, kind) {
    if (kind === 'taxid') return /^[A-Za-z0-9]+$/.test(val);
    if (kind === 'name') return /[\u2E80-\u9FFF]/.test(val);
    return true;
  }

  function isCommonLabel(val) {
    var labels = ['发票号码','开票日期','名称','纳税人识别号','统一社会信用代码',
      '地址','电话','开户行','账号','购买方','销售方',
      '价税合计','合计','备注','开票人','收款人','复核人',
      '货物','劳务','规格型号','单位','数量','单价','金额',
      '税率','税额','项目名称','小写','大写'];
    for (var li = 0; li < labels.length; li++) {
      if (val.indexOf(labels[li]) >= 0) return true;
    }
    return val.endsWith('：') || val.endsWith(':');
  }

  function extractValue(labelText, idx, kind) {
    for (var si = 0; si < 2; si++) {
      var sep = si === 0 ? '：' : ':';
      var pos = labelText.indexOf(sep);
      if (pos >= 0) {
        var after = labelText.substring(pos + sep.length).trim();
        if (after && valueMatches(after, kind)) return after;
      }
    }
    for (var j = idx + 1; j < Math.min(idx + 5, compositeTexts.length); j++) {
      var next = compositeTexts[j].trim();
      if (next && !isCommonLabel(next) && valueMatches(next, kind)) return next;
    }
    return null;
  }

  for (var ci = 0; ci < compositeTexts.length; ci++) {
    var t = compositeTexts[ci];
    var tns = t.replace(/\s/g, '');

    if (tns.indexOf('购买方') >= 0 || tns.indexOf('买方') >= 0) section = 'buyer';
    if (tns.indexOf('销售方') >= 0 || tns.indexOf('卖方') >= 0) section = 'seller';

    if ((t.indexOf('发票号码') >= 0 || tns.indexOf('发票号码') >= 0) && !info.invoiceNo) {
      info.invoiceNo = extractValue(t, ci, 'taxid');
    }
    if ((t.indexOf('开票日期') >= 0 || tns.indexOf('开票日期') >= 0) && !info.invoiceDate) {
      info.invoiceDate = extractValue(t, ci, 'any');
    }
    if ((t.indexOf('名称') >= 0 || tns.indexOf('名称') >= 0) && t.indexOf('货物') < 0 && t.indexOf('劳务') < 0 && t.indexOf('项目') < 0) {
      nameCount++;
      var val = extractValue(t, ci, 'name');
      var eff = section || (nameCount === 1 ? 'buyer' : 'seller');
      if (eff === 'buyer' && !info.buyerName) info.buyerName = val;
      else if (eff === 'seller' && !info.sellerName) info.sellerName = val;
    }
    if ((t.indexOf('纳税人识别号') >= 0 || t.indexOf('统一社会信用代码') >= 0) && t.indexOf('货物') < 0) {
      taxidCount++;
      var val2 = extractValue(t, ci, 'taxid');
      var eff2 = section || (taxidCount === 1 ? 'buyer' : 'seller');
      if (eff2 === 'buyer' && !info.buyerTaxId) info.buyerTaxId = val2;
      else if (eff2 === 'seller' && !info.sellerTaxId) info.sellerTaxId = val2;
    }
    if (tns.indexOf('价税合计') >= 0) {
      if (tns.indexOf('小写') >= 0) foundXiaoxie = true;
      else foundJiashui = true;
    }
    if ((t.indexOf('小写') >= 0 || tns.indexOf('小写') >= 0) && foundJiashui) foundXiaoxie = true;
    if ((tns.indexOf('合计') >= 0 || tns === '合计') && tns.indexOf('价税') < 0) foundHeji = true;

    if (t.charAt(0) === '¥' || t.charAt(0) === '￥') {
      var amtStr = t.replace(/^[¥￥]/, '').trim();
      var amt = parseFloat(amtStr);
      if (!isNaN(amt)) {
        if (foundXiaoxie) {
          if (!info.amountTax) info.amountTax = amt;
          foundXiaoxie = false;
          foundJiashui = false;
        } else if (foundHeji) {
          if (!info.amountNoTax) info.amountNoTax = amt;
          foundHeji = false;
        }
      }
    }
  }

  // 发票类型：在合并后的文本序列上判定（连续单字已拼接，拆字票同样可识别）
  if (!info.invoiceType) info.invoiceType = _detectInvoiceTypeTexts(compositeTexts);

  if (info.amountTax != null && info.amountNoTax == null) {
    info.amountNoTax = info.amountTax;
    info.taxAmount = 0;
  }
  if (info.amountNoTax != null && info.amountTax == null) {
    if (info.taxAmount != null) {
      info.amountTax = Math.round((info.amountNoTax + info.taxAmount) * 100) / 100;
    } else {
      info.amountTax = info.amountNoTax;
      info.taxAmount = 0;
    }
  }

  return info;
}

// =====================================================
// 项目名称 / 备注 坐标兜底（无 CustomTag 语义标签时）
// 对齐 Rust extract_ofd_item_name / extract_ofd_remark
// =====================================================

// 同一行判定的纵向容差（mm）
var OFD_ROW_TOL = 1.8;
// 品名列右界相对邻列表头左缘的内收量（mm）
var OFD_COL_MARGIN = 2.0;

/** 文本对象在页面坐标系里的位置：Boundary 原点 + CTM 作用于 TextCode 的 X/Y */
function _ofdPagePos(t) {
  var m = t.ctm || [1, 0, 0, 1, 0, 0];
  return [
    t.boundary[0] + m[0] * t.textX + m[2] * t.textY + m[4],
    t.boundary[1] + m[1] * t.textX + m[3] * t.textY + m[5]
  ];
}

/** 去空白后的文本，用于标签比对 */
function _ofdCompact(t) {
  return String(t.text || '').replace(/\s+/g, '');
}

/**
 * OFD 项目名称：按「项目名称」列边界取明细（不做全局星号扫描）。
 * 用「项目名称」表头 + 同排右侧最近表头的左缘定出品名列，只取列内、表头下方的文字；
 * `*` 开头为新条目、否则为换行续行并回上一条。
 */
function _extractOfdItemNameByColumn(texts, pageH) {
  // 1. 表头：最靠上的「项目名称」
  var header = null; // {x, y, w}
  for (var i = 0; i < texts.length; i++) {
    if (_ofdCompact(texts[i]) !== '项目名称') continue;
    var p = _ofdPagePos(texts[i]);
    if (!header || p[1] < header.y - 0.5 ||
        (Math.abs(p[1] - header.y) <= 0.5 && p[0] < header.x)) {
      header = { x: p[0], y: p[1], w: texts[i].boundary[2] };
    }
  }
  if (!header) return '';
  var hx = header.x, hy = header.y;

  // 2. 同一表头行里、表头右侧最近的非空对象左缘，往里收 OFD_COL_MARGIN = 品名列右界
  var colRight = null;
  for (var ci = 0; ci < texts.length; ci++) {
    if (!_ofdCompact(texts[ci])) continue;
    var cp = _ofdPagePos(texts[ci]);
    if (Math.abs(cp[1] - hy) > OFD_ROW_TOL) continue;
    if (cp[0] <= hx + 0.5) continue;
    if (colRight == null || cp[0] < colRight) colRight = cp[0];
  }
  // 找不到同排表头时退到一个保守值（表头自身右边 + 1mm）
  colRight = colRight == null ? hx + header.w + 1.0 : colRight - OFD_COL_MARGIN;

  // 3. 下界：表头下方最近的 合计/价税合计/大写/小写/备注 锚点，找不到退到表头 + 35% 页高
  var bottom = hy + pageH * 0.35;
  for (var bi = 0; bi < texts.length; bi++) {
    var s = _ofdCompact(texts[bi]);
    if (!s) continue;
    if (/^(?:单价|数量|单位|规格型号|项目名称)/.test(s)) continue;
    if (!(s.indexOf('合计') === 0 || s.indexOf('价税合计') >= 0 ||
          s.indexOf('大写') === 0 || s.indexOf('小写') === 0 || s.indexOf('备注') === 0)) continue;
    var by = _ofdPagePos(texts[bi])[1];
    if (by > hy + 0.5 && by < bottom) bottom = by;
  }

  // 4. 品名列内、表头下方的文字对象
  var cand = [];
  for (var ti = 0; ti < texts.length; ti++) {
    var raw = String(texts[ti].text || '').trim();
    if (!raw) continue;
    var tp = _ofdPagePos(texts[ti]);
    if (tp[1] <= hy + 0.5 || tp[1] >= bottom) continue;
    if (tp[0] >= colRight - 0.5) continue;
    if (_isStructuralItemRow(_ofdCompact(texts[ti]))) continue;
    // 与「规格型号」同串时裁掉尾部
    var cut = raw.indexOf('规格型号');
    var body = cut >= 0 ? raw.substring(0, cut) : raw;
    if (!body.trim()) continue;
    cand.push({ x: tp[0], y: tp[1], text: body });
  }
  if (!cand.length) return '';

  cand.sort(function(a, b) { return a.y - b.y || a.x - b.x; });

  // 5. 按 y 分行，`*` 开头为新条目、否则并回上一条
  var rows = [];
  for (var ri = 0; ri < cand.length; ri++) {
    var last = rows[rows.length - 1];
    if (last && Math.abs(cand[ri].y - last.y) < OFD_ROW_TOL) last.words.push(cand[ri]);
    else rows.push({ y: cand[ri].y, words: [cand[ri]] });
  }
  var starStyle = false;
  for (var si = 0; si < rows.length && !starStyle; si++) {
    for (var sj = 0; sj < rows[si].words.length; sj++) {
      if (/^[*＊]/.test(rows[si].words[sj].text)) { starStyle = true; break; }
    }
  }

  var items = [];
  var buf = '';
  for (var yi = 0; yi < rows.length; yi++) {
    var words = rows[yi].words;
    words.sort(function(a, b) { return a.x - b.x; });
    var line = words.map(function(w) { return w.text; }).join('');
    var isNew = /^[*＊]/.test(line);
    if (starStyle && buf && !isNew) {
      buf += line; // 换行续行：同一品名被折行的后半段
    } else {
      if (buf) {
        var n = _cleanItemName(buf);
        if (n.length >= 2) items.push(n);
      }
      buf = line;
    }
  }
  if (buf) {
    var n2 = _cleanItemName(buf);
    if (n2.length >= 2) items.push(n2);
  }
  return _joinItemNames(items);
}

/**
 * OFD 备注：左下角「备注」标签（支持竖排拆字 备/注）右侧、页脚行之上的文字（best-effort）
 */
function _extractOfdRemarkByCoords(texts, pageH) {
  // 标签：整词「备注」或竖排「备」+「注」，取最靠下的一组
  var label = null; // {x, y}
  for (var i = 0; i < texts.length; i++) {
    var c = _ofdCompact(texts[i]);
    var p = _ofdPagePos(texts[i]);
    if (c === '备注') {
      if (!label || p[1] > label.y) label = { x: p[0], y: p[1] };
      continue;
    }
    if (c !== '备') continue;
    for (var j = 0; j < texts.length; j++) {
      if (_ofdCompact(texts[j]) !== '注') continue;
      var q = _ofdPagePos(texts[j]);
      if (Math.abs(q[0] - p[0]) > 2.0 || q[1] <= p[1] || q[1] > p[1] + 14.0) continue;
      if (!label || p[1] > label.y) label = { x: p[0], y: p[1] };
      break;
    }
  }
  if (!label) return '';
  var ltop = label.y;

  // 上界：标签上方最近的 价税合计/大写/小写/合计 行之下
  // （标签常在框内垂直居中，不能只从标签顶部算，否则会漏掉比标签更高的备注行）
  var top = ltop - 6.0;
  for (var ai = 0; ai < texts.length; ai++) {
    var s = _ofdCompact(texts[ai]);
    if (!(s.indexOf('价税合计') >= 0 || s.indexOf('大写') === 0 || s.indexOf('小写') === 0 ||
          s.indexOf('合计') === 0 || s.indexOf('金额合计') === 0)) continue;
    var ay = _ofdPagePos(texts[ai])[1];
    if (ay < ltop && ay + 3.2 > top) top = ay + 3.2;
  }
  if (top < 0) top = 0;

  // 下界：标签下方最近的 开票人/收款人/复核人（没有则退到 97% 页高）；
  // 同时记录落款行的 y，用于剔除与落款标签同行的落款值（如「开票人:」旁的姓名）
  var bottom = pageH * 0.97;
  var footerRows = [];
  for (var fi = 0; fi < texts.length; fi++) {
    var fs = _ofdCompact(texts[fi]);
    if (!/^(?:开票人|收款人|复核人)/.test(fs)) continue;
    var fy = _ofdPagePos(texts[fi])[1];
    footerRows.push(fy);
    if (fy > ltop && fy < bottom) bottom = fy;
  }

  var rx = label.x + 1.6;
  var rows = [];
  for (var ti = 0; ti < texts.length; ti++) {
    var raw = String(texts[ti].text || '').trim();
    if (!raw) continue;
    var tp = _ofdPagePos(texts[ti]);
    if (tp[1] < top || tp[1] > bottom || tp[0] < rx) continue;
    var nos = _ofdCompact(texts[ti]);
    if (nos === '备注' || nos === '备' || nos === '注') continue;
    if (/^(?:开票人|收款人|复核人)/.test(nos)) continue;
    if (nos.indexOf('价税合计') >= 0 || nos.indexOf('大写') === 0 || nos.indexOf('小写') === 0) continue;
    var sameFooter = false;
    for (var fr = 0; fr < footerRows.length; fr++) {
      if (Math.abs(tp[1] - footerRows[fr]) < 3.0) { sameFooter = true; break; }
    }
    if (sameFooter) continue;
    var last = rows[rows.length - 1];
    if (last && Math.abs(tp[1] - last.y) < OFD_ROW_TOL) last.words.push({ x: tp[0], text: raw });
    else rows.push({ y: tp[1], words: [{ x: tp[0], text: raw }] });
  }
  if (!rows.length) return '';

  var joined = rows.map(function(r) {
    r.words.sort(function(a, b) { return a.x - b.x; });
    return r.words.map(function(w) { return w.text; }).join('');
  }).join(' ');
  return _cleanRemark(joined);
}

// =====================================================
// Build SVG
// =====================================================

function _buildOfdSvg(pageW, pageH, tplTexts, tplPaths, tplImgs, pageTexts, pagePaths, pageImgs, annotTexts, annotImgs, fontMap, colorSpaces, imageData) {
  var s = OFD_SVG_SCALE;
  var vw = pageW * s;
  var vh = pageH * s;

  var svg = '<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ' + vw.toFixed(1) + ' ' + vh.toFixed(1) + '" width="' + vw.toFixed(1) + '" height="' + vh.toFixed(1) + '" style="background:white">';

  svg += '<g id="template">';
  for (var i = 0; i < tplPaths.length; i++) svg += _buildSvgPath(tplPaths[i], s);
  for (var i = 0; i < tplTexts.length; i++) svg += _buildSvgText(tplTexts[i], fontMap, colorSpaces, s, s);
  for (var i = 0; i < tplImgs.length; i++) svg += _buildSvgImage(tplImgs[i], imageData, s);
  svg += '</g>';

  svg += '<g id="content">';
  for (var i = 0; i < pagePaths.length; i++) svg += _buildSvgPath(pagePaths[i], s);
  for (var i = 0; i < pageTexts.length; i++) svg += _buildSvgText(pageTexts[i], fontMap, colorSpaces, s, s);
  for (var i = 0; i < pageImgs.length; i++) svg += _buildSvgImage(pageImgs[i], imageData, s);
  svg += '</g>';

  svg += '<g id="annotations">';
  for (var i = 0; i < annotTexts.length; i++) svg += _buildSvgText(annotTexts[i], fontMap, colorSpaces, s, s);
  for (var i = 0; i < annotImgs.length; i++) svg += _buildSvgImage(annotImgs[i], imageData, s);
  svg += '</g>';

  svg += '</svg>';
  return svg;
}

// =====================================================
// Main entry: parse OFD from ArrayBuffer
// =====================================================

async function parseOfdFromArrayBuffer(arrayBuffer) {
  var zip = await JSZip.loadAsync(arrayBuffer);

  async function zipReadStr(name) {
    var f = zip.file(name);
    if (!f) return null;
    return await f.async('string');
  }

  async function zipReadBytes(name) {
    var f = zip.file(name);
    if (!f) return null;
    return await f.async('uint8array');
  }

  // 1. Read OFD.xml
  var ofdXml = await zipReadStr('OFD.xml');
  if (!ofdXml) throw new Error('OFD.xml 不存在');

  // 2. Find DocRoot (usually "Doc_0/Document.xml")
  var docRoot = 'Doc_0/Document.xml';
  var ofdDoc = _parseXml(ofdXml);
  var docRootEls = ofdDoc.getElementsByTagName('DocRoot');
  if (docRootEls.length > 0 && docRootEls[0].textContent) {
    docRoot = docRootEls[0].textContent.trim().replace(/^\//, '');
  }

  var baseDir = docRoot.indexOf('/') >= 0 ? docRoot.substring(0, docRoot.lastIndexOf('/')) : 'Doc_0';

  // 3. Parse CustomData from OFD.xml
  var customData = _parseCustomData(ofdXml);

  // 4. Read Document.xml
  var docXml = await zipReadStr(docRoot);
  if (!docXml) throw new Error(docRoot + ' 不存在');

  var docDom = _parseXml(docXml);

  // Parse template and page paths from Document.xml
  // OFD spec: <TemplatePage BaseLoc="..."> and <Page BaseLoc="...">
  // NOT textContent — BaseLoc attribute holds the path
  var tplPath = '';
  var pagePaths = [];
  var tplEls = docDom.getElementsByTagName('TemplatePage');
  if (tplEls.length > 0) {
    var baseLoc = _attr(tplEls[0], 'BaseLoc');
    if (baseLoc) tplPath = baseDir + '/' + baseLoc;
  }
  // Fallback: also check <Template> (older OFD versions)
  if (!tplPath) {
    var tEls = docDom.getElementsByTagName('Template');
    if (tEls.length > 0) {
      var tBaseLoc = _attr(tEls[0], 'BaseLoc');
      if (tBaseLoc) tplPath = baseDir + '/' + tBaseLoc;
    }
  }

  var pageEls = docDom.getElementsByTagName('Page');
  for (var pi = 0; pi < pageEls.length; pi++) {
    var pBaseLoc = _attr(pageEls[pi], 'BaseLoc');
    if (pBaseLoc) pagePaths.push(baseDir + '/' + pBaseLoc);
  }

  // 5. Parse PublicRes.xml for fonts + DrawParam
  var publicResXml = await zipReadStr(baseDir + '/PublicRes.xml');
  var fontResult = publicResXml ? _parseFonts(publicResXml) : { fonts: {}, colorSpaces: {}, drawParams: {} };

  // 6. Parse DocumentRes.xml for image resources
  var docResXml = await zipReadStr(baseDir + '/DocumentRes.xml');
  var imageResources = docResXml ? _parseImageResources(docResXml) : {};

  // Load image data as base64 data URLs
  // OFD spec: image file paths in DocumentRes.xml are relative to baseDir/Res/
  var imageData = {};
  var imageRawBytes = {};
  var imageFileNames = {};
  var imageEntries = Object.keys(imageResources);
  for (var ii = 0; ii < imageEntries.length; ii++) {
    var imgId = imageEntries[ii];
    var imgPath = baseDir + '/Res/' + imageResources[imgId];
    var imgBytes = await zipReadBytes(imgPath);
    if (imgBytes) {
      imageRawBytes[imgId] = imgBytes;
      imageFileNames[imgId] = imageResources[imgId];
    }
  }

  // Also scan per-page resources
  for (var ppi = 0; ppi < pagePaths.length; ppi++) {
    var pageXml = await zipReadStr(pagePaths[ppi]);
    if (!pageXml) continue;
    var pageDom = _parseXml(pageXml);
    var resEls = pageDom.getElementsByTagName('Resource');
    for (var ri = 0; ri < resEls.length; ri++) {
      var resBaseLoc = _attr(resEls[ri], 'BaseLoc');
      if (!resBaseLoc) continue;
      var resXml = await zipReadStr(baseDir + '/' + resBaseLoc);
      if (!resXml) continue;
      var pageImgRes = _parseImageResources(resXml);
      var pageImgKeys = Object.keys(pageImgRes);
      for (var ki = 0; ki < pageImgKeys.length; ki++) {
        var pid = pageImgKeys[ki];
        if (imageRawBytes[pid]) continue;
        var pImgPath = baseDir + '/Res/' + pageImgRes[pid];
        var pImgBytes = await zipReadBytes(pImgPath);
        if (pImgBytes) {
          imageRawBytes[pid] = pImgBytes;
          imageFileNames[pid] = pageImgRes[pid];
        }
      }
    }
  }

  // 7. Parse template content (background layer)
  var tplTexts = [], tplPaths = [], tplImgs = [];
  if (tplPath) {
    var tplXml = await zipReadStr(tplPath);
    if (tplXml) {
      var tplContent = _parseOfdContent(tplXml);
      tplTexts = tplContent.texts;
      tplPaths = tplContent.paths;
      tplImgs = tplContent.images;
    }
  }

  // 8. Parse page content (data layer)
  // Page content Layer has no DrawParam → OFD default: black (0,0,0).
  // Do NOT apply DrawParam inheritance — invoice data text and ¥ are naturally black.
  var pageTexts = [], pageObjPaths = [], pageImgs = [];
  if (pagePaths.length > 0) {
    var contentXml = await zipReadStr(pagePaths[0]);
    if (contentXml) {
      var pageContent = _parseOfdContent(contentXml);
      pageTexts = pageContent.texts;
      pageObjPaths = pageContent.paths;
      pageImgs = pageContent.images;
    }
  }

  // 9. Parse annotations (watermark layer)
  var annotTexts = [], annotImgs = [];
  var annotsPath = baseDir + '/Annots/Page_0/Annotation.xml';
  var annotXml = await zipReadStr(annotsPath);
  if (annotXml) {
    var annotContent = _parseAnnotations(annotXml);
    annotTexts = annotContent.texts;
    annotImgs = annotContent.images;
  }

  // 7b. Apply DrawParam defaults to template objects only
  // (page content objects keep default black, no DrawParam inheritance)
  _applyDrawParamDefaults(tplPaths, tplTexts, fontResult.drawParams);

  // 6b. Generate image data URLs (applying ImageMask compositing where needed)
  var maskMap = {};
  function collectMasks(imgs) {
    for (var mi = 0; mi < imgs.length; mi++) {
      if (imgs[mi].imageMask != null) {
        maskMap[imgs[mi].resourceId] = imgs[mi].imageMask;
      }
    }
  }
  collectMasks(tplImgs);
  collectMasks(pageImgs);
  collectMasks(annotImgs);

  var imageRawKeys = Object.keys(imageRawBytes);
  for (var ik = 0; ik < imageRawKeys.length; ik++) {
    var resId = imageRawKeys[ik];
    var bytes = imageRawBytes[resId];
    var fName = imageFileNames[resId] || '';
    var mime = fName.toLowerCase().endsWith('.png') || fName.toLowerCase().endsWith('.bmp')
      ? 'image/png' : 'image/jpeg';

    // Check if this image has a mask
    if (maskMap[resId] && imageRawBytes[maskMap[resId]]) {
      // Try canvas compositing for ImageMask
      try {
        var mainImg = await _loadImageFromBytes(bytes, mime);
        var maskBytes = imageRawBytes[maskMap[resId]];
        var maskMime = 'image/jpeg'; // masks are usually JPEG
        var maskImg = await _loadImageFromBytes(maskBytes, maskMime);
        if (mainImg.width === maskImg.width && mainImg.height === maskImg.height) {
          var compCanvas = document.createElement('canvas');
          compCanvas.width = mainImg.width;
          compCanvas.height = mainImg.height;
          var compCtx = compCanvas.getContext('2d');
          compCtx.drawImage(mainImg, 0, 0);
          var mainData = compCtx.getImageData(0, 0, compCanvas.width, compCanvas.height);
          var maskCanvas = document.createElement('canvas');
          maskCanvas.width = maskImg.width;
          maskCanvas.height = maskImg.height;
          var maskCtx = maskCanvas.getContext('2d');
          maskCtx.drawImage(maskImg, 0, 0);
          var maskData = maskCtx.getImageData(0, 0, maskCanvas.width, maskCanvas.height);
          for (var px = 0; px < mainData.data.length; px += 4) {
            mainData.data[px + 3] = maskData.data[px]; // Use red channel as alpha
          }
          compCtx.putImageData(mainData, 0, 0);
          imageData[resId] = compCanvas.toDataURL('image/png');
          continue;
        }
      } catch (e) {
        console.warn('ImageMask compositing failed for resource ' + resId + ', using unmasked');
      }
    }

    // Default: encode as-is
    var b64 = '';
    for (var bi = 0; bi < bytes.length; bi++) b64 += String.fromCharCode(bytes[bi]);
    imageData[resId] = 'data:' + mime + ';base64,' + btoa(b64);
  }

  // 10. Get page dimensions
  // Priority: page content XML PhysicalBox > Document.xml CommonData/PageArea > default
  var pageW = 210, pageH = 140;
  // First try Document.xml CommonData/PageArea/PhysicalBox
  var commonDataPb = docDom.getElementsByTagName('PhysicalBox');
  if (commonDataPb.length > 0) {
    var cdPbText = (commonDataPb[0].textContent || '').trim();
    var cdPb = _parseF4(cdPbText);
    if (cdPb) { pageW = cdPb[2]; pageH = cdPb[3]; }
  }
  // Then try page content XML (overrides Document.xml if present)
  if (pagePaths.length > 0) {
    var firstPageXml = await zipReadStr(pagePaths[0]);
    if (firstPageXml) {
      var fpDom = _parseXml(firstPageXml);
      var pbEls = fpDom.getElementsByTagName('PhysicalBox');
      if (pbEls.length > 0) {
        var pbText = (pbEls[0].textContent || '').trim();
        var pb = _parseF4(pbText);
        if (pb) { pageW = pb[2]; pageH = pb[3]; }
      }
    }
  }

  // 11. Parse CustomTag.xml
  var customTagPath = baseDir + '/Tags/CustomTag.xml';
  var ctXml = await zipReadStr(customTagPath);
  var customTag = ctXml ? _parseCustomTag(ctXml) : {};
  // 明细分组：长品名会被拆成多个文字对象，必须按 <Item> 组拼接
  var itemGroups = ctXml ? _parseCustomTagItemGroups(ctXml) : [];

  // 12. Extract invoice info from structured data
  var invoiceInfo = _extractFromCustomData(customData);

  // From CustomTag — only fill fields not already set
  // Rust version only looks up IDs in page_texts (not template texts)
  var ctInfo = _extractFromCustomTag(customTag, pageTexts, itemGroups, tplTexts);
  if (!invoiceInfo.invoiceNo && ctInfo.invoiceNo) invoiceInfo.invoiceNo = ctInfo.invoiceNo;
  if (!invoiceInfo.invoiceDate && ctInfo.invoiceDate) invoiceInfo.invoiceDate = ctInfo.invoiceDate;
  if (!invoiceInfo.buyerName && ctInfo.buyerName) invoiceInfo.buyerName = ctInfo.buyerName;
  if (!invoiceInfo.buyerTaxId && ctInfo.buyerTaxId) invoiceInfo.buyerTaxId = ctInfo.buyerTaxId;
  if (!invoiceInfo.sellerName && ctInfo.sellerName) invoiceInfo.sellerName = ctInfo.sellerName;
  if (!invoiceInfo.sellerTaxId && ctInfo.sellerTaxId) invoiceInfo.sellerTaxId = ctInfo.sellerTaxId;
  if (!invoiceInfo.amountNoTax && ctInfo.amountNoTax) invoiceInfo.amountNoTax = ctInfo.amountNoTax;
  if (!invoiceInfo.taxAmount && ctInfo.taxAmount) invoiceInfo.taxAmount = ctInfo.taxAmount;
  if (!invoiceInfo.amountTax && ctInfo.amountTax) invoiceInfo.amountTax = ctInfo.amountTax;
  if (!invoiceInfo.itemName && ctInfo.itemName) invoiceInfo.itemName = ctInfo.itemName;
  if (!invoiceInfo.remark && ctInfo.remark) invoiceInfo.remark = ctInfo.remark;

  // 项目名称 / 备注：语义路径（CustomTag 的 Item / Note）优先，缺失时才坐标兜底
  var layoutTexts = tplTexts.concat(pageTexts);
  if (!invoiceInfo.itemName) {
    invoiceInfo.itemName = _extractOfdItemNameByColumn(layoutTexts, pageH) || null;
  }
  if (!invoiceInfo.remark) {
    invoiceInfo.remark = _extractOfdRemarkByCoords(layoutTexts, pageH) || null;
  }

  // 发票类型：模板层 + 数据层文本合并判定（「普通」优先；逐条未命中时用拼接串兜底拆字票）
  if (!invoiceInfo.invoiceType) {
    var typeTexts = tplTexts.map(function(t) { return t.text || ''; })
      .concat(pageTexts.map(function(t) { return t.text || ''; }));
    typeTexts.push(typeTexts.join(''));
    invoiceInfo.invoiceType = _detectInvoiceTypeTexts(typeTexts);
  }

  // Text-based fallback extraction when no CustomData or CustomTag
  if (!invoiceInfo.invoiceNo && !invoiceInfo.invoiceDate &&
      !invoiceInfo.buyerName && !invoiceInfo.sellerName) {
    var allTextsForExtraction = tplTexts.concat(pageTexts);
    // Sort by ID (same as Rust version: all_texts.sort_by_key(|t| t.id))
    allTextsForExtraction.sort(function(a, b) { return a.id - b.id; });
    var textInfo = _extractInvoiceFromText(allTextsForExtraction);
    var textKeys = Object.keys(textInfo);
    for (var tk = 0; tk < textKeys.length; tk++) {
      if (!invoiceInfo[textKeys[tk]]) invoiceInfo[textKeys[tk]] = textInfo[textKeys[tk]];
    }
  }

  // 13. Build SVG
  var svg = _buildOfdSvg(pageW, pageH, tplTexts, tplPaths, tplImgs, pageTexts, pageObjPaths, pageImgs, annotTexts, annotImgs, fontResult.fonts, fontResult.colorSpaces, imageData);

  return {
    svg: svg,
    invoiceInfo: invoiceInfo,
    pageWidth: pageW,
    pageHeight: pageH
  };
}

// Helper: load image bytes into an HTML Image element
function _loadImageFromBytes(bytes, mime) {
  return new Promise(function(resolve, reject) {
    var b64 = '';
    for (var i = 0; i < bytes.length; i++) b64 += String.fromCharCode(bytes[i]);
    var dataUrl = 'data:' + mime + ';base64,' + btoa(b64);
    var img = new Image();
    img.onload = function() { resolve(img); };
    img.onerror = function() { reject(new Error('Image load failed')); };
    img.src = dataUrl;
  });
}

window.__ofdClient = {
  parseOfdFromArrayBuffer: parseOfdFromArrayBuffer
};
