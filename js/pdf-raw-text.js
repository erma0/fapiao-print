// =====================================================
// Content-stream 文字兜底（未嵌入字体的 PDF）
// =====================================================
// 12306 铁路电子客票这类 PDF 用 GBK-EUC-H 编码，但 KaiTi / SimHei / SimSun
// 一律未嵌入、也没有 ToUnicode。pdf.js 依赖 ToUnicode 做 CID→Unicode 映射，
// 缺了它整层文字直接消失——预览只剩背景、箭头和二维码，字段一个也识别不出来，
// 而桌面版靠 PDFium 的系统字体回退是正常的。
//
// 这里绕开 pdf.js：pdf-lib 解出 content stream，按文本操作符还原坐标与字符，
// 再用 TextDecoder('gbk') 解码，把文字补画回渲染好的 canvas，
// 同时产出与 pdf.js 同构的文本行供字段识别使用。
//
// 只在「pdf.js 判定文字层缺失」时启用，不影响正常 PDF。

(function () {
  var FONT_SERIF = '"KaiTi","Kaiti SC","STKaiti","SimSun","Songti SC","Noto Serif CJK SC",serif';
  var FONT_SANS = '"SimHei","Microsoft YaHei","PingFang SC","Noto Sans CJK SC",sans-serif';
  var MAX_RUNS = 4000;   // 兜底绘制上限，防止异常 PDF 卡死主线程

  // ---------- 基础工具 ----------
  function latin1(buf) {
    var s = '';
    for (var i = 0; i < buf.length; i += 4096) {
      s += String.fromCharCode.apply(null, buf.subarray(i, Math.min(i + 4096, buf.length)));
    }
    return s;
  }

  function nameOf(n) {
    if (!n) return '';
    var s = n.asString ? n.asString() : String(n);
    return s.replace(/^\//, '');
  }

  // PDF 矩阵 [a b c d e f]，m1 × m2 = 先应用 m1 再应用 m2
  function mul(m1, m2) {
    return [
      m1[0] * m2[0] + m1[1] * m2[2],
      m1[0] * m2[1] + m1[1] * m2[3],
      m1[2] * m2[0] + m1[3] * m2[2],
      m1[2] * m2[1] + m1[3] * m2[3],
      m1[4] * m2[0] + m1[5] * m2[2] + m2[4],
      m1[4] * m2[1] + m1[5] * m2[3] + m2[5]
    ];
  }

  // CJK 预定义 CMap 名 → 是否可用 TextDecoder 还原
  function decoderFor(encName) {
    if (!encName) return null;
    if (/^UniGB-UCS2|^UniGB-UTF16/.test(encName)) return 'utf-16be';
    if (/^(GBK|GBKp|GBK2K|GBpc|GB-EUC|GBpc-EUC|GBT-EUC|GBTpc-EUC)/.test(encName)) return 'gbk';
    return null;
  }

  function decodeBytes(bytes, decoderLabel) {
    var arr = new Uint8Array(bytes.length);
    for (var i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i) & 0xff;
    try { return new TextDecoder(decoderLabel).decode(arr); } catch (e) { return ''; }
  }

  // 无字宽表时的经验宽度（em）：ASCII 半角 0.5，CJK 全角 1
  function approxWidth(str) {
    var w = 0;
    for (var i = 0; i < str.length; i++) w += str.charCodeAt(i) < 0x2e80 ? 0.5 : 1;
    return w;
  }

  // ---------- 字体资源 ----------
  // 只收「GBK 系编码 + 未嵌入」的字体：这些正是 pdf.js 画不出来、需要兜底的那批。
  // 已嵌入的（如 DengXian）pdf.js 自己能画，再补一遍会重影。
  function readFonts(doc, page, PDFLib) {
    var ctx = doc.context;
    var map = {};
    var PN = PDFLib.PDFName;
    try {
      var resRef = page.node.Resources ? page.node.Resources() : page.node.get(PN.of('Resources'));
      var res = resRef ? ctx.lookup(resRef) : null;
      var fontsRef = res && res.get ? res.get(PN.of('Font')) : null;
      var fonts = fontsRef ? ctx.lookup(fontsRef) : null;
      if (!fonts || !fonts.entries) return map;
      fonts.entries().forEach(function (entry) {
        var key = nameOf(entry[0]);
        var fd = ctx.lookup(entry[1]);
        if (!fd || !fd.get) return;
        var enc = ctx.lookup(fd.get(PN.of('Encoding')));
        var encName = enc && enc.asString ? nameOf(enc) : '';
        var label = decoderFor(encName);
        if (!label) return;
        // 字体描述符：Type0 看 DescendantFonts[0]，否则直接看自己
        var desc = null;
        var sub = nameOf(fd.get(PN.of('Subtype')));
        if (sub === 'Type0') {
          var dfRef = fd.get(PN.of('DescendantFonts'));
          var df = dfRef ? ctx.lookup(dfRef) : null;
          var first = df && df.get ? ctx.lookup(df.get(0)) : null;
          if (first && first.get) desc = ctx.lookup(first.get(PN.of('FontDescriptor')));
        } else {
          desc = ctx.lookup(fd.get(PN.of('FontDescriptor')));
        }
        var embedded = !!(desc && desc.get && (
          desc.get(PN.of('FontFile')) || desc.get(PN.of('FontFile2')) || desc.get(PN.of('FontFile3'))));
        if (embedded) return;
        var base = nameOf(fd.get(PN.of('BaseFont')));
        map[key] = { decoder: label, sans: /Hei|YaHei|Gothic|Sans/i.test(base) };
      });
    } catch (e) {
      console.warn('[pdf-raw-text] 读取字体资源失败:', e);
    }
    return map;
  }

  // ---------- content stream ----------
  function readContent(doc, page, PDFLib) {
    var ctx = doc.context;
    var parts = [];
    function push(ref) {
      if (!ref) return;
      var s = ctx.lookup(ref);
      if (!s) return;
      try {
        var dec = PDFLib.decodePDFRawStream ? PDFLib.decodePDFRawStream(s) : s;
        var buf = dec.decode ? dec.decode() : (s.getContents ? s.getContents() : null);
        if (buf) parts.push(latin1(buf));
      } catch (e) { /* 未知过滤器，跳过 */ }
    }
    var contents = page.node.Contents();
    if (contents && contents.asArray) contents.asArray().forEach(push);
    else push(contents);
    return parts.join('\n');
  }

  // ---------- 词法分析 ----------
  function tokenize(src) {
    var toks = [];
    var i = 0;
    var n = src.length;
    while (i < n) {
      var c = src[i];
      if (c === '%') { while (i < n && src[i] !== '\n' && src[i] !== '\r') i++; continue; }
      if (c === ' ' || c === '\n' || c === '\r' || c === '\t' || c === '\f' || c === '\0') { i++; continue; }
      if (c === '(') {
        var depth = 1, j = i + 1, out = '';
        while (j < n && depth > 0) {
          var ch = src[j];
          if (ch === '\\') {
            var nx = src[j + 1];
            if (nx >= '0' && nx <= '7') {
              var oct = '', k = j + 1;
              while (k < n && oct.length < 3 && src[k] >= '0' && src[k] <= '7') { oct += src[k]; k++; }
              out += String.fromCharCode(parseInt(oct, 8) & 0xff);
              j = k;
              continue;
            }
            var esc = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f', '(': '(', ')': ')', '\\': '\\' };
            if (nx === '\n' || nx === '\r') { j += 2; continue; }
            out += esc[nx] != null ? esc[nx] : (nx || '');
            j += 2;
            continue;
          }
          if (ch === '(') depth++;
          else if (ch === ')') { depth--; if (depth === 0) { j++; break; } }
          out += ch;
          j++;
        }
        toks.push({ t: 'str', v: out });
        i = j;
        continue;
      }
      if (c === '<') {
        if (src[i + 1] === '<') { i += 2; continue; }
        var hex = '', h = i + 1;
        while (h < n && src[h] !== '>') { hex += src[h]; h++; }
        hex = hex.replace(/[^0-9A-Fa-f]/g, '');
        if (hex.length % 2) hex += '0';
        var bs = '';
        for (var q = 0; q < hex.length; q += 2) bs += String.fromCharCode(parseInt(hex.substr(q, 2), 16));
        toks.push({ t: 'str', v: bs });
        i = h + 1;
        continue;
      }
      if (c === '>') { i++; continue; }
      if (c === '[') { toks.push({ t: '[' }); i++; continue; }
      if (c === ']') { toks.push({ t: ']' }); i++; continue; }
      if (c === '/') {
        var nm = '';
        i++;
        while (i < n && !/[\s\/\[\]<>()%{}]/.test(src[i])) { nm += src[i]; i++; }
        toks.push({ t: 'name', v: nm });
        continue;
      }
      var start = i;
      while (i < n && !/[\s\/\[\]<>()%{}]/.test(src[i])) i++;
      var word = src.slice(start, i);
      if (!word) { i++; continue; }
      if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)) toks.push({ t: 'num', v: parseFloat(word) });
      else toks.push({ t: 'op', v: word });
    }
    return toks;
  }

  // ---------- 解释文本操作符 ----------
  // 只关心与文字位置/内容相关的操作符，图形状态（q/Q/cm）用于维护 CTM。
  function parseText(src, fontMap) {
    var toks = tokenize(src);
    var runs = [];
    var ctm = [1, 0, 0, 1, 0, 0];
    var stack = [];
    var st = { font: null, size: 0, charSpace: 0, wordSpace: 0, hScale: 1, leading: 0, rise: 0 };
    var tm = [1, 0, 0, 1, 0, 0];
    var tlm = [1, 0, 0, 1, 0, 0];
    var operands = [];

    function reset() { return { font: st.font, size: st.size, charSpace: st.charSpace, wordSpace: st.wordSpace, hScale: st.hScale, leading: st.leading, rise: st.rise }; }

    function show(bytes, dx) {
      if (!st.font || runs.length >= MAX_RUNS) return 0;
      var text = decodeBytes(bytes, st.font.decoder);
      if (!text) return 0;
      var m = mul(tm, ctm);
      if (dx) m = mul(mul([1, 0, 0, 1, dx, 0], tm), ctm);
      var size = st.size * Math.hypot(m[2], m[3]);
      if (size > 0.5 && Math.abs(m[1]) < 1e-4 && Math.abs(m[2]) < 1e-4) {
        runs.push({ text: text, x: m[4], y: m[5], size: size, sans: st.font.sans });
      }
      return approxWidth(text) * st.size * st.hScale;
    }

    function nextLine(tx, ty) {
      tlm = mul([1, 0, 0, 1, tx, ty], tlm);
      tm = tlm.slice();
    }

    for (var i = 0; i < toks.length; i++) {
      var tk = toks[i];
      if (tk.t !== 'op') { operands.push(tk); continue; }
      var op = tk.v;
      switch (op) {
        case 'q': stack.push({ ctm: ctm.slice(), st: reset() }); break;
        case 'Q':
          if (stack.length) { var sv = stack.pop(); ctm = sv.ctm; st = sv.st; }
          break;
        case 'cm':
          if (operands.length >= 6) {
            var a = operands.slice(-6).map(function (o) { return o.v; });
            ctm = mul(a, ctm);
          }
          break;
        case 'BT': tm = [1, 0, 0, 1, 0, 0]; tlm = tm.slice(); break;
        case 'ET': break;
        case 'Tf':
          if (operands.length >= 2) {
            st.font = fontMap[operands[operands.length - 2].v] || null;
            st.size = operands[operands.length - 1].v || 0;
          }
          break;
        case 'Tc': st.charSpace = operands[operands.length - 1].v || 0; break;
        case 'Tw': st.wordSpace = operands[operands.length - 1].v || 0; break;
        case 'Tz': st.hScale = (operands[operands.length - 1].v || 100) / 100; break;
        case 'TL': st.leading = operands[operands.length - 1].v || 0; break;
        case 'Ts': st.rise = operands[operands.length - 1].v || 0; break;
        case 'Td':
          if (operands.length >= 2) nextLine(operands[operands.length - 2].v, operands[operands.length - 1].v);
          break;
        case 'TD':
          if (operands.length >= 2) {
            st.leading = -(operands[operands.length - 1].v);
            nextLine(operands[operands.length - 2].v, operands[operands.length - 1].v);
          }
          break;
        case 'Tm':
          if (operands.length >= 6) {
            tm = operands.slice(-6).map(function (o) { return o.v; });
            tlm = tm.slice();
          }
          break;
        case 'T*': nextLine(0, -st.leading); break;
        case 'Tj':
          if (operands.length) show(operands[operands.length - 1].v);
          break;
        case "'":
          nextLine(0, -st.leading);
          if (operands.length) show(operands[operands.length - 1].v);
          break;
        case '"':
          if (operands.length >= 3) {
            st.wordSpace = operands[operands.length - 3].v || 0;
            st.charSpace = operands[operands.length - 2].v || 0;
            nextLine(0, -st.leading);
            show(operands[operands.length - 1].v);
          }
          break;
        case 'TJ': {
          var arrStart = -1;
          for (var k = operands.length - 1; k >= 0; k--) { if (operands[k].t === '[') { arrStart = k; break; } }
          if (arrStart >= 0) {
            var dx = 0;
            for (var a2 = arrStart + 1; a2 < operands.length; a2++) {
              var el = operands[a2];
              if (el.t === 'str') dx += show(el.v, dx);
              else if (el.t === 'num') dx -= (el.v / 1000) * st.size * st.hScale;
            }
          }
          break;
        }
        default: break;
      }
      operands = [];
    }
    return runs;
  }

  // ---------- 对外接口 ----------
  // 解析指定页的文字（PDF 用户空间坐标，单位 pt，y 轴向上）
  async function extract(bytes, pageIndex) {
    var PDFLib = window.PDFLib;
    if (!PDFLib) throw new Error('pdf-lib 未加载');
    var data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    var doc = await PDFLib.PDFDocument.load(data, { ignoreEncryption: true, updateMetadata: false });
    var page = doc.getPage(pageIndex || 0);
    var size = page.getSize();
    var fontMap = readFonts(doc, page, PDFLib);
    var src = readContent(doc, page, PDFLib);
    var runs = src ? parseText(src, fontMap) : [];
    return {
      widthPt: size.width,
      heightPt: size.height,
      rotation: page.getRotation ? page.getRotation().angle : 0,
      runs: runs
    };
  }

  // 把文字补画到 pdf.js 渲染好的 canvas 上（vp1 = getViewport({scale:1})）
  function draw(canvas, runs, vp1) {
    if (!runs.length) return 0;
    var scale = canvas.width / vp1.width;
    var ctx = canvas.getContext('2d');
    ctx.save();
    ctx.textBaseline = 'alphabetic';
    ctx.fillStyle = '#000';
    var drawn = 0;
    for (var i = 0; i < runs.length; i++) {
      var r = runs[i];
      var p = vp1.convertToViewportPoint(r.x, r.y);
      ctx.font = (r.size * scale).toFixed(2) + 'px ' + (r.sans ? FONT_SANS : FONT_SERIF);
      ctx.fillText(r.text, p[0] * scale, p[1] * scale);
      drawn++;
    }
    ctx.restore();
    return drawn;
  }

  // 与 pdf.js extractTextWithCoords 同构的文本行（单位：dpi 换算后的像素）
  function buildTextResult(extracted, dpi, baseResult) {
    var scale = dpi / 72;
    var pageH = extracted.heightPt;
    var lines = [];
    var current = null;
    var lastY = null;
    var texts = [];
    extracted.runs.forEach(function (r) {
      var px = r.x * scale;
      var py = (pageH - r.y - r.size) * scale;
      var word = { text: r.text, x: px, y: Math.max(0, py), w: approxWidth(r.text) * r.size * scale, h: r.size * scale };
      // 基线差 ≤ 2pt 归为同一行（与 extractTextWithCoords 的阈值一致）
      if (lastY === null || Math.abs(r.y - lastY) > 2) {
        if (current && current.length) lines.push({ words: current, confidence: 1.0 });
        current = [];
      }
      current.push(word);
      lastY = r.y;
    });
    if (current && current.length) lines.push({ words: current, confidence: 1.0 });
    lines.forEach(function (ln) { texts.push(ln.words.map(function (w) { return w.text; }).join('')); });
    var base = baseResult || {};
    return {
      text: texts.join('\n'),
      lines: lines,
      imgW: Math.round(extracted.widthPt * scale),
      imgH: Math.round(pageH * scale),
      hasTextLayer: lines.length > 0,
      textLayerDiagnostic: base.textLayerDiagnostic || null,
      rawFallback: true
    };
  }

  window.__pdfRawText = { extract: extract, draw: draw, buildTextResult: buildTextResult };
})();
