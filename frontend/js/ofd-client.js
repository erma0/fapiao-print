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
  return s.trim().split(/\s+/).map(function(v) { return parseFloat(v); }).filter(function(v) { return !isNaN(v); });
}

function _normalizeFontName(raw) {
  if (!raw) return 'SimSun';
  var base = raw;
  var plusIdx = base.indexOf('+');
  if (plusIdx >= 0) base = base.substring(plusIdx + 1);
  var dashIdx = base.indexOf('-');
  if (dashIdx >= 0) base = base.substring(0, dashIdx);
  base = base.replace(/_.+$/, '');
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
  if (!color) return ' fill="none"';
  var a = alpha != null ? (alpha / 255).toFixed(3) : null;
  if (a && a !== '1.000') return ' fill="rgb(' + color[0] + ',' + color[1] + ',' + color[2] + ')" fill-opacity="' + a + '"';
  return ' fill="rgb(' + color[0] + ',' + color[1] + ',' + color[2] + ')"';
}

function _strokeAttr(color, alpha) {
  if (!color) return '';
  var a = alpha != null ? (alpha / 255).toFixed(3) : null;
  if (a && a !== '1.000') return ' stroke="rgb(' + color[0] + ',' + color[1] + ',' + color[2] + ')" stroke-opacity="' + a + '"';
  return ' stroke="rgb(' + color[0] + ',' + color[1] + ',' + color[2] + ')"';
}

function _escXml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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
  var boundary = t.boundary;
  var textX = t.textX;
  var textY = t.textY;
  var deltaX = t.deltaX;

  var tx = boundary[0] * scaleX;
  var ty = boundary[1] * scaleY;
  var bw = boundary[2] * scaleX;
  var bh = boundary[3] * scaleY;

  var baseX = textX * scaleX;
  var baseY = (boundary[3] - textY) * scaleY;

  var fsScaled = fontSize * scaleY;

  var fillStr = _fillAttr(t.fillColor || (t.layerDrawParam ? null : [0, 0, 0]), t.alpha);
  var strokeStr = t.strokeColor ? _strokeAttr(t.strokeColor, t.alpha) : '';

  if (t.ctm) {
    var a = t.ctm[0] * scaleX;
    var b = -t.ctm[1] * scaleY;
    var c = -t.ctm[2] * scaleX;
    var d = t.ctm[3] * scaleY;
    var e = t.ctm[4] * scaleX;
    var f = -t.ctm[5] * scaleY;
    return '<text transform="translate(' + tx.toFixed(4) + ',' + ty.toFixed(4) + ') matrix(' +
      a.toFixed(4) + ',' + b.toFixed(4) + ',' + c.toFixed(4) + ',' + d.toFixed(4) + ',' + e.toFixed(4) + ',' + f.toFixed(4) + ')"' +
      ' font-family="' + _escXmlAttr(fontFamily) + '"' +
      ' font-size="' + fsScaled.toFixed(2) + '"' + bold + fillStr + strokeStr + '>' +
      _escXml(text) + '</text>';
  }

  if (deltaX.length > 0 && text.length > 1) {
    var tspans = '';
    var charX = baseX;
    for (var ci = 0; ci < text.length; ci++) {
      var xp = charX.toFixed(2);
      tspans += '<tspan x="' + xp + '">' + _escXml(text[ci]) + '</tspan>';
      if (ci < deltaX.length) charX += deltaX[ci] * scaleX;
    }
    return '<text transform="translate(' + tx.toFixed(4) + ',' + ty.toFixed(4) + ')"' +
      ' font-family="' + _escXmlAttr(fontFamily) + '"' +
      ' font-size="' + fsScaled.toFixed(2) + '"' + bold + fillStr + strokeStr + '>' +
      tspans + '</text>';
  }

  return '<text transform="translate(' + tx.toFixed(4) + ',' + ty.toFixed(4) + ')"' +
    ' x="' + baseX.toFixed(2) + '" y="' + baseY.toFixed(2) + '"' +
    ' font-family="' + _escXmlAttr(fontFamily) + '"' +
    ' font-size="' + fsScaled.toFixed(2) + '"' + bold + fillStr + strokeStr + '>' +
    _escXml(text) + '</text>';
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

  return '<image' + attrs + ' xlink:href="' + _escXmlAttr(dataUrl) + '"/>';
}

// =====================================================
// XML Parsers (using DOMParser)
// =====================================================

function _parseOfdContent(xml) {
  var textObjs = [];
  var pathObjs = [];
  var imgObjs = [];

  var parser = new DOMParser();
  var doc = parser.parseFromString(xml, 'text/xml');
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
  var parser = new DOMParser();
  var doc = parser.parseFromString(xml, 'text/xml');
  var items = doc.getElementsByTagName('CustomData');
  for (var i = 0; i < items.length; i++) {
    var name = _attr(items[i], 'Name');
    if (name) map[name] = (items[i].textContent || '').trim();
  }
  return map;
}

function _parseCustomTag(xml) {
  var map = {};
  var parser = new DOMParser();
  var doc = parser.parseFromString(xml, 'text/xml');

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

function _parseFonts(xml) {
  var fonts = {};
  var colorSpaces = {};
  var drawParams = {};

  var parser = new DOMParser();
  var doc = parser.parseFromString(xml, 'text/xml');

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
  var parser = new DOMParser();
  var doc = parser.parseFromString(xml, 'text/xml');
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
  var parser = new DOMParser();
  var doc = parser.parseFromString(xml, 'text/xml');
  var annots = doc.getElementsByTagName('Annot');

  for (var ai = 0; ai < annots.length; ai++) {
    var annot = annots[ai];
    var appearances = annot.getElementsByTagName('Appearance');
    for (var apIdx = 0; apIdx < appearances.length; apIdx++) {
      var ap = appearances[apIdx];
      var apBoundary = _parseF4(_attr(ap, 'Boundary'));
      var offsetX = apBoundary ? apBoundary[0] : 0;
      var offsetY = apBoundary ? apBoundary[1] : 0;

      var content = _parseOfdContent(new XMLSerializer().serializeToString(ap));
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
  if (customData['发票号码']) info.invoiceNo = customData['发票号码'];
  if (customData['开票日期']) info.invoiceDate = customData['开票日期'];
  if (customData['购买方名称']) info.buyerName = customData['购买方名称'];
  if (customData['购买方纳税人识别号']) info.buyerTaxId = customData['购买方纳税人识别号'];
  if (customData['销售方名称']) info.sellerName = customData['销售方名称'];
  if (customData['销售方纳税人识别号']) info.sellerTaxId = customData['销售方纳税人识别号'];
  if (customData['合计金额']) info.amountNoTax = parseFloat(customData['合计金额']) || null;
  if (customData['合计税额']) info.taxAmount = parseFloat(customData['合计税额']) || null;
  if (customData['价税合计']) info.amountTax = parseFloat(customData['价税合计']) || null;
  return info;
}

function _extractFromCustomTag(customTag, allTexts) {
  var info = {};
  var idToText = {};
  for (var i = 0; i < allTexts.length; i++) {
    idToText[allTexts[i].id] = allTexts[i].text;
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

  return info;
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
    if (!info.invoiceType) {
      if (t.indexOf('增值税专用') >= 0) info.invoiceType = '增值税专用发票';
      else if (t.indexOf('增值税普通') >= 0 || t.indexOf('增值税电子普通') >= 0) info.invoiceType = '增值税普通发票';
      else if (t.indexOf('电子发票') >= 0) info.invoiceType = '电子发票';
    }
  }

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

  // 2. Find DocRoot
  var docRoot = 'Doc_0/Document.xml';
  var ofdDoc = new DOMParser().parseFromString(ofdXml, 'text/xml');
  var docRootEls = ofdDoc.getElementsByTagName('DocRoot');
  if (docRootEls.length > 0 && docRootEls[0].textContent) {
    docRoot = docRootEls[0].textContent.trim().replace(/^\//, '');
  }

  var baseDir = docRoot.indexOf('/') >= 0 ? docRoot.substring(0, docRoot.lastIndexOf('/')) : 'Doc_0';

  // 3. Parse CustomData
  var customData = _parseCustomData(ofdXml);

  // 4. Read Document.xml
  var docXml = await zipReadStr(docRoot);
  if (!docXml) throw new Error(docRoot + ' 不存在');

  var docDom = new DOMParser().parseFromString(docXml, 'text/xml');

  // Page size
  var pageDataEls = docDom.getElementsByTagName('Page');
  var pageW = 210, pageH = 297;
  if (pageDataEls.length > 0) {
    var ps = _parseF4(_attr(pageDataEls[0], 'PhysicalBox'));
    if (ps) { pageW = ps[2]; pageH = ps[3]; }
  }

  // Template
  var tplFile = null;
  var tplEls = docDom.getElementsByTagName('Template');
  if (tplEls.length > 0) {
    tplFile = (tplEls[0].textContent || '').trim();
  }

  // Page content file
  var contentFile = null;
  var contentEls = docDom.getElementsByTagName('Content');
  if (contentEls.length > 0) {
    contentFile = (contentEls[0].textContent || '').trim();
  }

  // Annotation file
  var annotFile = null;
  var annotEls = docDom.getElementsByTagName('Annotations');
  if (annotEls.length > 0) {
    annotFile = (annotEls[0].textContent || '').trim();
  }

  // 5. Parse PublicRes.xml
  var publicResXml = await zipReadStr(baseDir + '/PublicRes.xml');
  var fontResult = publicResXml ? _parseFonts(publicResXml) : { fonts: {}, colorSpaces: {}, drawParams: {} };

  // 6. Parse DocumentRes.xml for image resources
  var docResXml = await zipReadStr(baseDir + '/DocumentRes.xml');
  var imageResources = docResXml ? _parseImageResources(docResXml) : {};

  // Load image data as base64 data URLs
  var imageData = {};
  var imageEntries = Object.keys(imageResources);
  for (var ii = 0; ii < imageEntries.length; ii++) {
    var imgId = imageEntries[ii];
    var imgPath = imageResources[imgId];
    var imgBytes = await zipReadBytes(baseDir + '/' + imgPath);
    if (imgBytes) {
      var mime = imgPath.toLowerCase().endsWith('.png') ? 'image/png' : 'image/jpeg';
      var b64 = '';
      for (var bi = 0; bi < imgBytes.length; bi++) b64 += String.fromCharCode(imgBytes[bi]);
      imageData[imgId] = 'data:' + mime + ';base64,' + btoa(b64);
    }
  }

  // Also scan per-page resources
  var pagesEls = docDom.getElementsByTagName('Page');
  for (var pi = 0; pi < pagesEls.length; pi++) {
    var pageEl = pagesEls[pi];
    var pageFile = (pageEl.textContent || '').trim();
    if (!pageFile) continue;
    var pageXml = await zipReadStr(baseDir + '/' + pageFile);
    if (!pageXml) continue;
    var pageResEls = new DOMParser().parseFromString(pageXml, 'text/xml').getElementsByTagName('Resource');
    for (var ri = 0; ri < pageResEls.length; ri++) {
      var resFile = (pageResEls[ri].textContent || '').trim();
      if (!resFile) continue;
      var resXml = await zipReadStr(baseDir + '/' + resFile);
      if (!resXml) continue;
      var pageImgRes = _parseImageResources(resXml);
      var pageImgKeys = Object.keys(pageImgRes);
      for (var ki = 0; ki < pageImgKeys.length; ki++) {
        var pid = pageImgKeys[ki];
        if (imageData[pid]) continue;
        var pImgBytes = await zipReadBytes(baseDir + '/' + pageImgRes[pid]);
        if (pImgBytes) {
          var pMime = pageImgRes[pid].toLowerCase().endsWith('.png') ? 'image/png' : 'image/jpeg';
          var pB64 = '';
          for (var bj = 0; bj < pImgBytes.length; bj++) pB64 += String.fromCharCode(pImgBytes[bj]);
          imageData[pid] = 'data:' + pMime + ';base64,' + btoa(pB64);
        }
      }
    }
  }

  // 7. Parse template content
  var tplTexts = [], tplPaths = [], tplImgs = [];
  if (tplFile) {
    var tplXml = await zipReadStr(baseDir + '/' + tplFile);
    if (tplXml) {
      var tplContent = _parseOfdContent(tplXml);
      tplTexts = tplContent.texts;
      tplPaths = tplContent.paths;
      tplImgs = tplContent.images;
    }
  }

  // 8. Parse page content
  var pageTexts = [], pagePaths = [], pageImgs = [];
  if (contentFile) {
    var contentXml = await zipReadStr(baseDir + '/' + contentFile);
    if (contentXml) {
      var pageContent = _parseOfdContent(contentXml);
      pageTexts = pageContent.texts;
      pagePaths = pageContent.paths;
      pageImgs = pageContent.images;
    }
  }

  // 9. Parse annotations
  var annotTexts = [], annotImgs = [];
  if (annotFile) {
    var annotXml = await zipReadStr(baseDir + '/' + annotFile);
    if (annotXml) {
      var annotContent = _parseAnnotations(annotXml);
      annotTexts = annotContent.texts;
      annotImgs = annotContent.images;
    }
  }

  // 10. Parse CustomTag
  var customTagFile = null;
  var ctEls = docDom.getElementsByTagName('Tags');
  if (ctEls.length > 0) customTagFile = (ctEls[0].textContent || '').trim();
  var customTag = {};
  if (customTagFile) {
    var ctXml = await zipReadStr(baseDir + '/' + customTagFile);
    if (ctXml) customTag = _parseCustomTag(ctXml);
  }

  // 11. Apply DrawParam defaults
  var allTexts = tplTexts.concat(pageTexts, annotTexts);
  var allPaths = tplPaths.concat(pagePaths);
  _applyDrawParamDefaults(allPaths, allTexts, fontResult.drawParams);

  // 12. Extract invoice info
  var invoiceInfo = _extractFromCustomData(customData);

  var ctInfo = _extractFromCustomTag(customTag, allTexts);
  if (!invoiceInfo.invoiceNo && ctInfo.invoiceNo) invoiceInfo.invoiceNo = ctInfo.invoiceNo;
  if (!invoiceInfo.invoiceDate && ctInfo.invoiceDate) invoiceInfo.invoiceDate = ctInfo.invoiceDate;
  if (!invoiceInfo.buyerName && ctInfo.buyerName) invoiceInfo.buyerName = ctInfo.buyerName;
  if (!invoiceInfo.buyerTaxId && ctInfo.buyerTaxId) invoiceInfo.buyerTaxId = ctInfo.buyerTaxId;
  if (!invoiceInfo.sellerName && ctInfo.sellerName) invoiceInfo.sellerName = ctInfo.sellerName;
  if (!invoiceInfo.sellerTaxId && ctInfo.sellerTaxId) invoiceInfo.sellerTaxId = ctInfo.sellerTaxId;
  if (!invoiceInfo.amountNoTax && ctInfo.amountNoTax) invoiceInfo.amountNoTax = ctInfo.amountNoTax;
  if (!invoiceInfo.taxAmount && ctInfo.taxAmount) invoiceInfo.taxAmount = ctInfo.taxAmount;
  if (!invoiceInfo.amountTax && ctInfo.amountTax) invoiceInfo.amountTax = ctInfo.amountTax;

  // Text-based fallback
  if (!invoiceInfo.invoiceNo && !invoiceInfo.sellerName) {
    var textInfo = _extractInvoiceFromText(allTexts);
    var textKeys = Object.keys(textInfo);
    for (var tk = 0; tk < textKeys.length; tk++) {
      if (!invoiceInfo[textKeys[tk]]) invoiceInfo[textKeys[tk]] = textInfo[textKeys[tk]];
    }
  }

  // 13. Build SVG
  var svg = _buildOfdSvg(pageW, pageH, tplTexts, tplPaths, tplImgs, pageTexts, pagePaths, pageImgs, annotTexts, annotImgs, fontResult.fonts, fontResult.colorSpaces, imageData);

  return {
    svg: svg,
    invoiceInfo: invoiceInfo,
    pageWidth: pageW,
    pageHeight: pageH
  };
}

window.__ofdClient = {
  parseOfdFromArrayBuffer: parseOfdFromArrayBuffer
};
