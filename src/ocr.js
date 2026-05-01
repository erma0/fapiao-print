// =====================================================
// OCR & Invoice Info Extraction
// =====================================================
// Dependencies (global): invoke, isTauri, dataUrlToUint8Array

/**
 * Parse amount string to number (2 decimal places)
 */
function parseAmt(s) {
  if (!s) return 0;
  var n = parseFloat(s.replace(/,/g, ''));
  return (!isNaN(n) && n > 0) ? Math.round(n * 100) / 100 : 0;
}

/**
 * Detect if text is a train/ride ticket (no seller info needed)
 */
function isTicketText(text) {
  var t = text.substring(0, 500);
  return /(?:车\s*次|票\s*价|座\s*位|席\s*别|检\s*票|站\s*台|进\s*站|出\s*站|铁\s*路|乘\s*车|二\s*等|一\s*等|动\s*车|高\s*铁|硬\s*座|软\s*座|卧\s*铺|铺\s*位|出\s*租|打\s*车|网\s*约|滴\s*滴)/.test(t);
}

/**
 * Get a descriptive label for ticket type (shown as sellerName for tickets)
 */
function getTicketTypeLabel(text) {
  var t = text.substring(0, 500);
  if (/(?:铁\s*路|动\s*车|高\s*铁|火\s*车|车\s*次|座\s*位|席\s*别|检\s*票|进\s*站|出\s*站|硬\s*座|软\s*座|卧\s*铺|铺\s*位)/.test(t)) return '铁路电子客票';
  if (/(?:出\s*租|打\s*车|的\s*士)/.test(t)) return '出租车票';
  if (/(?:网\s*约|滴\s*滴|专\s*车|快\s*车)/.test(t)) return '网约车票';
  return '车票';
}

/**
 * Normalize OCR currency symbol artifacts.
 * OCR commonly misreads digits and ¥ symbols because they look similar:
 *   - "1" as "¥" → "¥¥72.68" should be "¥172.68" (second ¥ is misread "1")
 *   - "¥" as "1" → "1317.00" should be "¥317.00" (handled by keyword-based rule)
 *   - Mixed full-width ￥ and half-width ¥
 */
function normalizeOcrCurrency(s) {
  if (!s) return s;
  // Double ¥ before a digit: the second ¥ is a misread "1" digit
  // "¥¥72.68" → "¥172.68", "￥¥07.00" → "¥107.00"
  s = s.replace(/[¥￥]¥(\d)/g, '¥1$1');
  // Apply again in case of triple ¥ (very rare): "¥¥¥07" → "¥1¥07" → "¥117.07"
  s = s.replace(/¥¥(\d)/g, '¥1$1');
  // "1¥" pattern before a digit: ¥ was misread as "1" and "1" as "¥" (swap)
  // "1¥72.68" → "¥172.68". Only apply when preceded by non-digit (avoid breaking real numbers)
  s = s.replace(/(\D)1¥(\d)/g, '$1¥1$2');
  // Also handle "1¥" at the start of the string
  s = s.replace(/^1¥(\d)/, '¥1$1');
  // Normalize remaining full-width ￥ to half-width ¥ for consistency
  s = s.replace(/￥/g, '¥');
  return s;
}

// =====================================================
// Coordinate-aware region analysis
// =====================================================

/**
 * Classify a word's region based on its position on the invoice.
 * Invoice layout (typical):
 *   Top-left:   购买方 (buyer)
 *   Top-right:  销售方 (seller)
 *   Bottom:     金额/合计 (amounts)
 *   Far bottom: 备注 (remarks)
 *
 * Returns: 'buyer' | 'seller' | 'amount' | 'remark' | 'unknown'
 */
function classifyRegion(wx, wy, ww, wh, imgW, imgH) {
  if (!imgW || !imgH) return 'unknown';
  var nx = wx / imgW;   // normalized 0~1
  var ny = wy / imgH;

  // Vertical zones
  // Top 55%: buyer/seller area
  // 55%~75%: amount area
  // Below 75%: remarks
  if (ny < 0.55) {
    // Top section: split left/right
    return nx < 0.5 ? 'buyer' : 'seller';
  } else if (ny < 0.75) {
    return 'amount';
  } else {
    return 'remark';
  }
}

/**
 * Build a region-annotated word list from OCR coordinates.
 * Each entry: { text, x, y, w, h, region, lineIdx, wordIdx, confidence }
 */
function buildWordMap(ocrLines, imgW, imgH) {
  if (!ocrLines || !ocrLines.length) return [];
  var map = [];
  for (var li = 0; li < ocrLines.length; li++) {
    var line = ocrLines[li];
    if (!line.words || !line.words.length) continue;
    var lineConfidence = line.confidence || 0;
    for (var wi = 0; wi < line.words.length; wi++) {
      var word = line.words[wi];
      map.push({
        text: normalizeOcrCurrency(word.text),
        x: word.x,
        y: word.y,
        w: word.w,
        h: word.h,
        region: classifyRegion(word.x, word.y, word.w, word.h, imgW, imgH),
        lineIdx: li,
        wordIdx: wi,
        confidence: lineConfidence
      });
    }
  }
  return map;
}

/**
 * Get text for a specific region from the word map.
 * Joins words in reading order (top-to-bottom, left-to-right within same line).
 */
function getRegionText(wordMap, region) {
  var words = wordMap.filter(function(w) { return w.region === region; });
  if (!words.length) return '';
  // Group by line, then join
  var byLine = {};
  words.forEach(function(w) {
    if (!byLine[w.lineIdx]) byLine[w.lineIdx] = [];
    byLine[w.lineIdx].push(w);
  });
  var lines = Object.keys(byLine).map(function(k) {
    // Sort words within line by x position
    byLine[k].sort(function(a, b) { return a.x - b.x; });
    return byLine[k].map(function(w) { return w.text; }).join('');
  });
  // Sort lines by y position (first word's y)
  var sortedKeys = Object.keys(byLine).sort(function(a, b) {
    return byLine[a][0].y - byLine[b][0].y;
  });
  return sortedKeys.map(function(k) {
    byLine[k].sort(function(a, b) { return a.x - b.x; });
    return byLine[k].map(function(w) { return w.text; }).join('');
  }).join('\n');
}

/**
 * Clean an OCR amount string: strip ¥/￥ prefix, handle "1" misread of "¥".
 * OCR often misreads "¥317.00" as "1317.00" (¥→1). We detect this by checking
 * if a leading "1" could be a misread ¥ symbol: the number after removing "1"
 * must have exactly 2 decimal places and be a reasonable amount.
 * Returns the cleaned numeric string.
 */
function cleanOcrAmtStr(raw) {
  var hadYenPrefix = /^[¥￥-]/.test(raw);
  var s = raw.replace(/^[¥￥-]+/, '').replace(/[,，]/g, '');
  // ¥→1 misread detection:
  // Only strip leading "1" if the original did NOT have a ¥/negative prefix,
  // AND the number has 4+ digits before decimal (1 + 3+ digits).
  // When ¥ is present (e.g., "¥172.68"), the "1" is a legitimate digit,
  // not a misread ¥. Without ¥ prefix (e.g., "1317.00"), the "1" is likely
  // a misread "¥" symbol (they look very similar in OCR).
  // 4+ digit check prevents stripping "1" from 3-digit amounts like "172.68"
  // which are common and legitimate (stripping would give wrong "72.68").
  // e.g., "1317.00" (4 digits, no ¥) → "317.00" ✓
  // e.g., "¥172.68" (has ¥) → keep "172.68" ✓ (NOT "72.68")
  // e.g., "172.68" (3 digits, no ¥) → keep "172.68" ✓ (NOT "72.68")
  // e.g., "1299.06" (4 digits, no ¥) → "299.06" ✓
  if (!hadYenPrefix && /^1\d{3,}\.\d{2}$/.test(s)) {
    var stripped = s.substring(1);
    var strippedVal = parseFloat(stripped);
    if (strippedVal > 0) {
      s = stripped;
    }
  }
  return s;
}

/**
 * Check if a numeric value looks like a year or date.
 * OCR can produce "2025.01" or "2025.00" from dates like "2025年01月" or "2025/01/15".
 * These should NOT be treated as monetary amounts.
 * Returns true if the value looks like a year/date, false otherwise.
 */
function isLikelyYearOrDate(val, rawText) {
  // Integer part in year range (1900-2099) and value < 2100 → almost certainly a year
  if (val >= 1900 && val < 2100) return true;
  // Check raw text for year-like pattern: "20XX.XX" where XX could be month
  if (rawText && /^-?¥?(20\d{2})\.\d{2}$/.test(rawText)) return true;
  return false;
}

/**
 * Collect all amount-like numbers from wordMap, optionally filtered by
 * region and/or normalized position ranges (0~1).
 * Returns array of { value, x, y, text, word } sorted by value descending.
 * Excludes values that look like years/dates.
 */
function collectAmountWords(wordMap, imgW, imgH, regionFilter, nxMin, nxMax, nyMin, nyMax) {
  var results = [];
  wordMap.forEach(function(w) {
    if (regionFilter && w.region !== regionFilter && regionFilter !== 'any') return;
    // Skip low-confidence OCR results (< 0.3) — likely garbage
    if (w.confidence !== undefined && w.confidence < 0.3) return;
    if (imgW > 0 && imgH > 0) {
      var nx = (w.x + w.w / 2) / imgW;
      var ny = (w.y + w.h / 2) / imgH;
      if (nxMin !== undefined && nx < nxMin) return;
      if (nxMax !== undefined && nx > nxMax) return;
      if (nyMin !== undefined && ny < nyMin) return;
      if (nyMax !== undefined && ny > nyMax) return;
    }
    var t = w.text.replace(/[,，]/g, '');
    // Match ¥-prefixed or bare amounts with exactly 2 decimal places
    var m = t.match(/^-?¥?(\d+\.\d{2})$/);
    if (m) {
      var val = parseFloat(cleanOcrAmtStr(t));
      if (val > 0 && val < 1000000 && !isLikelyYearOrDate(val, t)) {
        results.push({ value: val, x: w.x, y: w.y, text: w.text, word: w });
      }
    }
  });
  results.sort(function(a, b) { return b.value - a.value; });
  return results;
}

/**
 * Find words matching a regex in a specific region, return the matching word
 * plus nearby words (within same line or adjacent lines).
 */
function findWordsNear(wordMap, regex, region, contextWords) {
  contextWords = contextWords || 5;
  var matches = [];
  wordMap.forEach(function(w) {
    if (w.region !== region && region !== 'any') return;
    if (regex.test(w.text)) matches.push(w);
  });
  return matches;
}

// [REMOVED] extractInvoiceInfo() — legacy regex-based extraction (~1100 lines).
// Disabled since v1.7.0 in favor of coordinate-first extractByCoordinates().
// If re-enabling is needed, restore from git history (commit before this change).


/**
 * Apply an already-parsed OCR result to a file object.
 * Extracts invoice info (amounts, seller) from OCR text/coordinates.
 * Used by both applyOcr() (image files) and render_and_ocr_pdf (PDF one-pass).
 * Modifies fileObj in place.
 * @param {Object} fileObj - The file object to update
 * @param {Object} ocrResult - Parsed OCR result from Rust (with lines, imgW, imgH, text)
 */
function applyOcrResult(fileObj, ocrResult) {
  if (!ocrResult) return;
  try {
    // --- v1.7.0: Try coordinate-first extraction (PP-OCRv5 bbox) ---
    var info = null;
    if (ocrResult.lines && ocrResult.imgW > 0 && ocrResult.imgH > 0) {
      info = extractByCoordinates(ocrResult);
    }

    // Always set _ocrText for display — this is the main purpose of running OCR on all pages
    fileObj._ocrText = info._ocrText || ocrResult.text || '';
    fileObj._isTicket = info.isTicket || false;
    // Only update amounts if they are not already set (PDF.js text extraction is more reliable for text-based PDFs)
    var effAmt = info.amountTax > 0 ? info.amountTax : info.amountNoTax;
    if (effAmt > 0 && !fileObj.amountTax && !fileObj.amountNoTax) {
      fileObj.amount = effAmt;
      fileObj.amountTax = info.amountTax;
      fileObj.amountNoTax = info.amountNoTax;
      fileObj.taxAmount = info.taxAmount || 0;
    } else if (effAmt > 0 && fileObj.amountTax > 0) {
      // Amounts already set — only fill in missing taxAmount
      if (!fileObj.taxAmount && info.taxAmount > 0) {
        fileObj.taxAmount = info.taxAmount;
      }
    } else if (info.taxAmount > 0 && !fileObj.taxAmount) {
      fileObj.taxAmount = info.taxAmount;
    }
    // Set seller info — for tickets, sellerName is the ticket type label
    if (info.sellerName) fileObj.sellerName = info.sellerName;
    if (!info.isTicket) {
      if (info.sellerCreditCode) fileObj.sellerCreditCode = info.sellerCreditCode;
    }
    // Set additional extracted fields
    if (info.invoiceNo) fileObj.invoiceNo = info.invoiceNo;
    if (info.invoiceDate) fileObj.invoiceDate = info.invoiceDate;
    if (info.buyerName) fileObj.buyerName = info.buyerName;
    if (info.buyerCreditCode) fileObj.buyerCreditCode = info.buyerCreditCode;
  } catch(e) {
    console.warn('[OCR] 结果应用失败:', e);
  }
}

/**
 * Apply OCR to a file object — calls Rust OCR then applies result.
 * Used for image files (non-PDF). PDF files use render_and_ocr_pdf one-pass instead.
 * Modifies fileObj in place, adding amount/seller info if detected.
 * @param {Object} fileObj - The file object to update
 * @param {string} dataUrl - Base64 data URL of the image to OCR (fallback)
 * @param {string} [filePath] - Disk path to the image file (preferred — skips base64)
 */
async function applyOcr(fileObj, dataUrl, filePath) {
  if (!hasOcr || !isTauri || !invoke) return;
  try {
    var ocrResult = await invoke('ocr_image', {
      dataUrl: dataUrl || '',
      filePath: filePath || fileObj._filePath || null
    });
    if (!ocrResult) return;
    applyOcrResult(fileObj, ocrResult);
  } catch(e) {
    console.warn('[OCR] 识别失败:', e);
  }
}

/**
 * OCR a PDF page via ocr_pdf_page command — zero IPC round-trip.
 * Rust renders the page AND runs OCR internally, then returns just the OcrResult.
 * This avoids: Rust render → base64 → IPC → frontend downsample → base64 → IPC → Rust decode → OCR.
 * Instead: Rust render → decode in memory → OCR → return result directly.
 */
async function applyOcrPdfPage(fileObj) {
  if (!hasOcr || !isTauri || !invoke) return;
  try {
    var ocrResult = await invoke('ocr_pdf_page', {
      pdfPath: fileObj._pdfPath,
      pageIndex: fileObj._pdfPageIdx
    });
    if (!ocrResult) return;
    applyOcrResult(fileObj, ocrResult);
  } catch(e) {
    console.warn('[OCR] PDF页识别失败:', e);
  }
}


// =====================================================
// v1.7.0 — Coordinate-first invoice extraction
// =====================================================
// Designed for PP-OCRv5's high-accuracy bbox output.
// Strategy: Use real OCR coordinates to locate fields directly,
// then fall back to simple regex only when coordinates can't resolve.
//
// Invoice layout (normalized 0~1 coordinates, Y-axis: top=0, bottom=1):
//
//   VAT invoice (增值税发票):
//     ny 0.00~0.15:  标题 "电子发票(普通发票)" + 发票号码 + 开票日期
//     ny 0.15~0.35:  购买方信息 (nx 0~0.5) | 销售方信息 (nx 0.5~1.0)
//     ny 0.35~0.45:  明细表头 (项目名称/金额/税率/税额)
//     ny 0.45~0.60:  明细行
//     ny 0.60~0.70:  合计行 (不含税金额合计 + 税额合计)
//     ny 0.70~0.80:  价税合计 (大写)(小写)¥XXX.XX
//     ny 0.80~1.00:  备注 + 开票人
//
//   Train ticket (铁路电子客票):
//     ny 0.00~0.15:  标题 + 发票号码 + 开票日期
//     ny 0.15~0.35:  出发站/到达站/车次
//     ny 0.35~0.55:  票价 + 座位/等级
//     ny 0.55~0.75:  身份证号/姓名
//     ny 0.75~1.00:  客票号 + 购买方信息

/**
 * Normalize a word's text for matching (fullwidth→halfwidth, collapse CJK spaces).
 */
function _normText(s) {
  if (!s) return '';
  s = s.replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  s = s.replace(/[Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  s = s.replace(/％/g, '%').replace(/．/g, '.').replace(/，/g, ',').replace(/：/g, ':');
  s = s.replace(/￥/g, '¥');
  // Collapse spaces between CJK chars
  s = s.replace(/([\u4e00-\u9fff])\s+([\u4e00-\u9fff])/g, '$1$2');
  return s;
}

/**
 * Normalize OCR text for structured extraction.
 * Like _normText() but preserves newlines (critical for line-based regex matching).
 * The regular normText collapses CJK newlines, which merges separate "名称:" entries
 * into one line and breaks line-by-line extraction.
 */
function _normTextForExtract(text) {
  if (!text) return '';
  // Full-width → half-width
  text = text.replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  text = text.replace(/[Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  text = text.replace(/％/g, '%').replace(/．/g, '.').replace(/，/g, ',').replace(/：/g, ':');
  text = text.replace(/￥/g, '¥');
  // Collapse spaces between CJK chars ON THE SAME LINE only (preserve newlines)
  text = text.replace(/([\u4e00-\u9fff])[ \t]+([\u4e00-\u9fff])/g, '$1$2');
  // Collapse spaces between digits on the same line
  for (var _i = 0; _i < 3; _i++) {
    var _prev = '';
    while (_prev !== text) { _prev = text; text = text.replace(/(\d)[ \t]+(\d)/g, '$1$2'); }
  }
  text = text.replace(/(\d)[ \t]+\./g, '$1.');
  text = text.replace(/¥[ \t]+(\d)/g, '¥$1');
  text = text.replace(/([\u4e00-\u9fff])\s+¥/g, '$1¥');

  // Normalize OCR ¥↔1 misread artifacts (critical for amount accuracy)
  // Step 1: ¥¥ patterns — OCR misreads "1" as "¥" producing "¥¥72.68" → "¥172.68"
  text = normalizeOcrCurrency(text);

  // Step 2: Keyword-based ¥→1 misread correction (restored from v1.6.7)
  // OCR often misreads "¥" as "1" (they look very similar). After amount keywords,
  // "1XXX.XX" (4+ digits before decimal) should be "¥XXX.XX".
  // e.g., "价税合计1317.00" → "价税合计¥317.00"
  // Only apply after amount keywords to avoid corrupting legitimate numbers.
  // \d{3,} requires 3+ digits after "1" (4+ total) to avoid stripping "1" from
  // legitimate 3-digit amounts like "金额172.68" (should stay 172.68).
  text = text.replace(/(价\s*税\s*合\s*计|金\s*额|税\s*额|合\s*计|票\s*价|总\s*计|不\s*含\s*税|含\s*税|实\s*付|应\s*付|开\s*票\s*金\s*额|发\s*票\s*金\s*额|全\s*价|优\s*惠\s*价|小\s*写)([^\d¥￥]*?)1(\d{3,}\.\d{2})/g, '$1$2¥$3');

  return text;
}

/**
 * Clean a captured name string for structured extraction.
 * Removes trailing labels, punctuation, and validates format.
 */
function _cleanName(raw) {
  if (!raw) return '';
  var name = raw.trim();
  // Trim at next label keyword (when OCR merges multiple labels into one line)
  name = name.replace(/名\s*称[:：].*$/, '');
  name = name.replace(/统一社会信用代码.*$/, '');
  name = name.replace(/纳税人识别号.*$/, '');
  name = name.replace(/开户银行.*$/, '');
  name = name.replace(/银行账号.*$/, '');
  name = name.replace(/地址电话.*$/, '');
  // Remove trailing punctuation and whitespace
  name = name.replace(/[，,。.、：:；;！!？?\s]+$/, '');
  // Remove leading whitespace/colons
  name = name.replace(/^[\s:：]+/, '');
  // Skip if it's a label itself
  if (/^(?:购买方信息|销售方信息|购买方|销售方|名称|信息|纳税人|地址|电话|开户行|账号|项目名称|规格型号)$/.test(name)) return '';
  // Must contain CJK and be at least 2 chars
  if (name.length < 2 || !/[\u4e00-\u9fff]/.test(name)) return '';
  return name;
}

/**
 * Text-based extraction from OCR text (preserving line structure).
 * PRIMARY extraction method for structured fields (invoice number, date,
 * buyer/seller names and credit codes). Coordinate-based extraction is the fallback.
 *
 * Key insight: OCR output is well-formatted with clear key-value pairs:
 *   发票号码：2532200000380892372
 *   开票日期：2025年08月19日
 *   名称：无锡市天鹏食品有限公司           ← 购买方 (1st)
 *   名称：无锡市志成生化工程装备有限公司    ← 销售方 (2nd)
 *   统一社会信用代码/纳税人识别号：913202001358946118  ← 购买方 (1st)
 *   统一社会信用代码/纳税人识别号：913202057431110944  ← 销售方 (2nd)
 *
 * Returns: { invoiceNo, invoiceDate, buyerName, sellerName, buyerCreditCode, sellerCreditCode }
 */
function _extractByText(fullText) {
  var result = {
    invoiceNo: '',
    invoiceDate: '',
    buyerName: '',
    sellerName: '',
    buyerCreditCode: '',
    sellerCreditCode: ''
  };
  if (!fullText) return result;

  var text = _normTextForExtract(fullText);

  // --- Invoice number ---
  var noMatch = text.match(/发\s*票\s*号\s*码[:\s]*(\d{8,20})/);
  if (noMatch) result.invoiceNo = noMatch[1];

  // --- Invoice date ---
  var dateMatch = text.match(/开\s*票\s*日\s*期[:\s]*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
  if (dateMatch) {
    result.invoiceDate = dateMatch[1] + '-' +
      dateMatch[2].padStart(2, '0') + '-' +
      dateMatch[3].padStart(2, '0');
  }

  // --- Buyer/Seller names ---
  // Priority 1: Explicit labels "购买方名称：" / "销售方名称："
  var buyerLabelMatch = text.match(/购\s*买\s*方(?:信息)?名\s*称[:\s]*([^\n]+)/);
  if (buyerLabelMatch) {
    var bn = _cleanName(buyerLabelMatch[1]);
    if (bn) result.buyerName = bn;
  }
  var sellerLabelMatch = text.match(/销\s*售\s*方(?:信息)?名\s*称[:\s]*([^\n]+)/);
  if (sellerLabelMatch) {
    var sn = _cleanName(sellerLabelMatch[1]);
    if (sn) result.sellerName = sn;
  }

  // --- Buyer/Seller credit codes (extract FIRST — use as anchor for name matching) ---
  // First occurrence = buyer, second = seller (same left-to-right order as names)
  var ccRegex = /(?:统一社会信用代码|纳税人识别号)[^A-Z0-9]{0,30}([A-Z0-9]{15,20})/gi;
  var codes = [];
  var ccPositions = [];
  var cm;
  while ((cm = ccRegex.exec(text)) !== null) {
    var code = cm[1].toUpperCase();
    if (codes.indexOf(code) < 0) {
      codes.push(code);
      ccPositions.push(cm.index);
    }
  }
  if (codes.length >= 1) result.buyerCreditCode = codes[0];
  if (codes.length >= 2) result.sellerCreditCode = codes[1];

  // Priority 2: "销方名称" / "销方" abbreviated form (v1.6.7 strategy)
  if (!result.sellerName) {
    var shortSeller = text.match(/销\s*方(?:信息)?名\s*称[:\s]*([^\n]+)/);
    if (shortSeller) {
      var ssn = _cleanName(shortSeller[1]);
      if (ssn) result.sellerName = ssn;
    }
  }

  // Priority 3: Generic "名称：" — first = buyer, second = seller
  // Use credit code position as anchor: find "名称" before seller's credit code
  if (!result.buyerName || !result.sellerName) {
    var nameRegex = /名\s*称[:：]\s*([^\n]+)/g;
    var names = [];
    var namePositions = [];
    var nm;
    while ((nm = nameRegex.exec(text)) !== null) {
      var name = _cleanName(nm[1]);
      if (name && names.indexOf(name) < 0) {
        names.push(name);
        namePositions.push(nm.index);
      }
    }
    if (!result.buyerName && names.length >= 1) result.buyerName = names[0];
    if (!result.sellerName) {
      // Credit-code-anchored strategy (v1.6.7): find LAST "名称" before seller's credit code
      if (codes.length >= 2 && ccPositions.length >= 2) {
        var sellerCcPos = ccPositions[ccPositions.length - 1];
        var lastNameBeforeCc = '';
        for (var ni = 0; ni < namePositions.length; ni++) {
          if (namePositions[ni] < sellerCcPos) {
            var nc = _cleanName(names[ni]);
            if (nc && !/^(?:购买方|销售方|名称)/.test(nc)) {
              lastNameBeforeCc = nc;
            }
          }
        }
        if (lastNameBeforeCc) result.sellerName = lastNameBeforeCc;
      }
      // Fallback: 2nd "名称" match
      if (!result.sellerName && names.length >= 2) result.sellerName = names[1];
      else if (!result.sellerName && names.length === 1) result.sellerName = names[0];
    }
  }

  // Priority 4: "收款单位" / "销货单位" / "开票方" (non-standard invoices)
  if (!result.sellerName) {
    var altSeller = text.match(/(?:收款单位|销货单位|开票方|销售单位)[^\n]{0,30}?[:：]?\s*([^\n]{2,60}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[A-Z0-9]{15,20})|\n|$)/i);
    if (altSeller) {
      var altName = _cleanName(altSeller[1]);
      if (altName) result.sellerName = altName;
    }
  }

  // Priority 5: Company name near the last credit code (v1.6.7 Strategy 4)
  // Some OCR outputs have: "91440300xxxxxxxxx  深圳市某某科技有限公司"
  if (!result.sellerName && ccPositions.length > 0) {
    var csSuffix = '(?:公司|集团|商行|商店|厂|部|院|所|中心|店|馆|站|社|行|会|处|室|局|办|坊|铺|有限合伙|合伙企业|个体工商户|个体户|工作室|经营部|门市部|分公司|事业部|事务所|医院|学校|幼儿园|合作社|企业|商社|贸易行|服务部)';
    var lastCcPos = ccPositions[ccPositions.length - 1];
    var afterLastCc = text.substring(lastCcPos);
    var compNearCc = afterLastCc.match(new RegExp('[A-Z0-9]{15,20}\\s*[:：]?\\s*([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]+' + csSuffix + ')'));
    if (compNearCc) {
      var compName = _cleanName(compNearCc[1]);
      if (compName) result.sellerName = compName;
    }
  }

  // Priority 6: Last company name with suffix in full text (v1.6.7 Strategy 6)
  if (!result.sellerName) {
    var csSuffix2 = '(?:公司|集团|商行|商店|厂|部|院|所|中心|店|馆|站|社|行|会|处|室|局|办|坊|铺|有限合伙|合伙企业|个体工商户|个体户|工作室|经营部|门市部|分公司|事业部|事务所|医院|学校|幼儿园|合作社|企业|商社|贸易行|服务部)';
    var allCompRe = new RegExp('([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]{2,25}' + csSuffix2 + ')', 'g');
    var allCompMatches = [];
    var ccm;
    while ((ccm = allCompRe.exec(text)) !== null) {
      var cn = ccm[1].trim();
      if (cn.length > 3 && !/^(?:购买方|销售方|信息|名称|地址)/.test(cn)) {
        allCompMatches.push(cn);
      }
    }
    if (allCompMatches.length >= 2) {
      result.sellerName = allCompMatches[allCompMatches.length - 1];
    } else if (allCompMatches.length === 1 && !result.buyerName) {
      // Only one company found — could be seller if no buyer found either
      result.sellerName = allCompMatches[0];
    }
  }

  // Also try standalone credit codes (some OCR misses the label prefix)
  if (!result.buyerCreditCode || !result.sellerCreditCode) {
    var standaloneRe = /\b([0-9][A-Z0-9]{17})\b/g;
    var sm;
    var standaloneCodes = [];
    while ((sm = standaloneRe.exec(text)) !== null) {
      if (/\d{6,}/.test(sm[1]) && /[A-Z]/.test(sm[1])) {
        var sc = sm[1].toUpperCase();
        if (standaloneCodes.indexOf(sc) < 0) standaloneCodes.push(sc);
      }
    }
    if (!result.buyerCreditCode && standaloneCodes.length >= 1) result.buyerCreditCode = standaloneCodes[0];
    if (!result.sellerCreditCode && standaloneCodes.length >= 2) result.sellerCreditCode = standaloneCodes[1];
    else if (!result.sellerCreditCode && standaloneCodes.length === 1) result.sellerCreditCode = standaloneCodes[0];
  }

  return result;
}

/**
 * Text-based amount extraction from OCR text.
 * PRIMARY method for extracting amountTax, amountNoTax, and taxAmount.
 *
 * Key insight: In the "合计" row of a VAT invoice, there are two amounts:
 *   - 不含税合计 (amountNoTax) — always the LARGER value (since 税率 < 100%)
 *   - 税额合计 (taxAmount) — always the SMALLER value
 * Because 税额 = 不含税金额 × 税率, and 税率 < 100%, so 税额 < 不含税.
 *
 * The "价税合计" row has one amount:
 *   - 含税总价 (amountTax) — the ¥ amount after "（小写）" or last after "价税合计"
 *
 * Returns: { amountTax, amountNoTax, taxAmount }
 */
function _extractAmountsByText(fullText) {
  var result = { amountTax: 0, amountNoTax: 0, taxAmount: 0 };
  if (!fullText) return result;

  var text = _normTextForExtract(fullText);

  // ========== Phase 1: Extract amountTax (含税总价) ==========

  // Pattern 1: "（小写）¥XXX.XX" — most specific indicator of 含税价
  var xxMatch = text.match(/小\s*写[）\)]*[：:]*\s*¥\s*(\d[\d,]*\.\d{2})/);
  if (xxMatch) {
    result.amountTax = parseAmt(xxMatch[1]);
  }
  // Pattern 1b: "（小写）XXX.XX" — bare amount after 小写 (no ¥ prefix)
  // Handles OCR like "(小写）70000.00" where ¥ is missing
  if (!result.amountTax) {
    var xxBare = text.match(/小\s*写[）\)]*[：:]*\s*(\d[\d,]*\.\d{2})/);
    if (xxBare) {
      var v1b = parseAmt(xxBare[1]);
      if (v1b > 10 && !isLikelyYearOrDate(v1b, xxBare[1])) {
        result.amountTax = v1b;
      }
    }
  }
  // Pattern 2: Find last ¥ amount after "价税合计"
  if (!result.amountTax) {
    var jshjIdx = text.search(/价\s*税\s*合\s*计/);
    if (jshjIdx >= 0) {
      var afterJshj = text.substring(jshjIdx);
      var jshjAmtRe = /¥\s*(\d[\d,]*\.\d{2})/g;
      var jm, lastAmt = 0;
      while ((jm = jshjAmtRe.exec(afterJshj)) !== null) {
        var v2 = parseAmt(jm[1]);
        if (v2 > 0 && !isLikelyYearOrDate(v2, jm[1])) lastAmt = v2;
      }
      if (lastAmt > 0) result.amountTax = lastAmt;
    }
  }
  // Pattern 2b: Bare amount after "价税合计" (no ¥, no 小写)
  // Last resort for amountTax — handles rare formats
  if (!result.amountTax) {
    var jshjIdx2 = text.search(/价\s*税\s*合\s*计/);
    if (jshjIdx2 >= 0) {
      var afterJshj2 = text.substring(jshjIdx2);
      // Look for bare amount after closing bracket or colon
      var bareJshj = afterJshj2.match(/[）\)][：:]*\s*(\d[\d,]*\.\d{2})/);
      if (bareJshj) {
        var v2b = parseAmt(bareJshj[1]);
        if (v2b > 10 && !isLikelyYearOrDate(v2b, bareJshj[1])) {
          result.amountTax = v2b;
        }
      }
    }
  }

  // ========== Phase 2: Math-verified pair finding (PRIMARY) ==========
  // Key insight: amountNoTax + taxAmount = amountTax (always true for VAT)
  // Since 税率 ∈ (0%, 100%), amountNoTax > taxAmount.
  // So: find two amounts that sum to amountTax → larger = 不含税, smaller = 税额.
  // This is more robust than keyword/section parsing because:
  //   - Works even when "合计" keyword is garbled/split/reversed by OCR
  //   - Self-validating: the sum constraint eliminates false matches
  //   - No need to distinguish amounts by position or keyword context
  if (result.amountTax > 0 && (!result.amountNoTax || !result.taxAmount)) {
    var allAmts = [];
    var amtSeen = {};

    // Collect ¥-prefixed amounts
    var yenRe = /¥\s*(\d[\d,]*\.\d{2})/g;
    var ym;
    while ((ym = yenRe.exec(text)) !== null) {
      var yv = parseAmt(ym[1]);
      if (yv > 0 && !isLikelyYearOrDate(yv, ym[1]) && !amtSeen[yv]) {
        allAmts.push(yv);
        amtSeen[yv] = true;
      }
    }

    // Collect bare amounts with 2 decimal places (> 1.00, not years)
    // These cover cases where OCR drops the ¥ prefix
    var numRe = /(\d[\d,]*\.\d{2})/g;
    var nm;
    while ((nm = numRe.exec(text)) !== null) {
      var nv = parseAmt(nm[1]);
      if (nv > 1 && !isLikelyYearOrDate(nv, nm[1]) && !amtSeen[nv]) {
        allAmts.push(nv);
        amtSeen[nv] = true;
      }
    }

    console.log('[数学验证] 所有金额:', allAmts, '目标含税价:', result.amountTax);

    // Find pair (a, b) where a + b ≈ amountTax
    // If multiple pairs match (extremely rare), prefer the one with largest 不含税
    var bestPair = null;
    for (var pi = 0; pi < allAmts.length; pi++) {
      for (var pj = pi + 1; pj < allAmts.length; pj++) {
        var pairSum = Math.round((allAmts[pi] + allAmts[pj]) * 100) / 100;
        if (Math.abs(pairSum - result.amountTax) < 0.02) {
          var pLarger = Math.max(allAmts[pi], allAmts[pj]);
          var pSmaller = Math.min(allAmts[pi], allAmts[pj]);
          // Sanity: both must be positive and smaller than 含税价
          if (pSmaller > 0 && pLarger < result.amountTax) {
            if (!bestPair || pLarger > bestPair.larger) {
              bestPair = { larger: pLarger, smaller: pSmaller };
            }
          }
        }
      }
    }

    if (bestPair) {
      result.amountNoTax = bestPair.larger;
      result.taxAmount = bestPair.smaller;
      console.log('[数学验证] 配对成功: 不含税=' + bestPair.larger + ', 税额=' + bestPair.smaller +
                  ', 验证: ' + bestPair.larger + '+' + bestPair.smaller + '=' +
                  Math.round((bestPair.larger + bestPair.smaller) * 100) / 100);
    } else {
      console.log('[数学验证] 未找到配对');
    }
  }

  // ========== Phase 3: Fallback — section-based 合计 parsing ==========
  // Only runs if math-verified pair finding didn't work (e.g., amountTax unknown,
  // or no pair sums correctly due to OCR errors in amount digits).
  if (!result.amountNoTax || !result.taxAmount) {
    // Find standalone "合计" (not "价税合计")
    var hejiStandaloneIdx = -1;
    var hejiRegex = /合\s*计/g;
    var hm;
    while ((hm = hejiRegex.exec(text)) !== null) {
      var before = text.substring(Math.max(0, hm.index - 3), hm.index);
      if (!/价|税/.test(before)) {
        hejiStandaloneIdx = hm.index;
        break;
      }
    }

    if (hejiStandaloneIdx >= 0) {
      // Section: from standalone "合计" to "价税合计" (or end of text)
      var jshjSearchIdx = text.indexOf('价税合计', hejiStandaloneIdx);
      if (jshjSearchIdx < 0) {
        var jshjAfter = text.substring(hejiStandaloneIdx).search(/价\s*税\s*合\s*计/);
        jshjSearchIdx = jshjAfter >= 0 ? hejiStandaloneIdx + jshjAfter : text.length;
      }
      var section = text.substring(hejiStandaloneIdx, jshjSearchIdx);

      // Find all ¥-prefixed amounts in this section
      var amtRe = /¥\s*(\d[\d,]*\.\d{2})/g;
      var amts = [];
      var am;
      while ((am = amtRe.exec(section)) !== null) {
        var val = parseAmt(am[1]);
        if (val > 0 && !isLikelyYearOrDate(val, am[1])) {
          amts.push(val);
        }
      }

      // If fewer than 2 ¥ amounts, also look for bare amounts with 2 decimal places
      if (amts.length < 2) {
        var bareRe = /(\d+\.\d{2})/g;
        var bm;
        var fallbackSeen = {};
        amts.forEach(function(v) { fallbackSeen[v] = true; });
        while ((bm = bareRe.exec(section)) !== null) {
          var bval = parseFloat(bm[1]);
          if (bval > 0 && bval < 1000000 && !isLikelyYearOrDate(bval, bm[1]) && !fallbackSeen[bval]) {
            amts.push(bval);
            fallbackSeen[bval] = true;
          }
        }
      }

      if (amts.length >= 2) {
        // Sort descending: largest = 不含税合计, second = 税额合计
        amts.sort(function(a, b) { return b - a; });
        if (!result.amountNoTax) result.amountNoTax = amts[0];
        if (!result.taxAmount) result.taxAmount = amts[1];
      } else if (amts.length === 1 && !result.amountNoTax) {
        result.amountNoTax = amts[0];
      }
    }
  } // end Phase 3 fallback

  return result;
}

/**
 * Build a flat word array from OCR lines, with normalized positions.
 * Each word: { text, normText, x, y, w, h, cx, cy, nx, ny, lineIdx, wordIdx, confidence, points }
 * cx/cy = center of word; nx/ny = normalized center (0~1).
 */
function _buildWords(ocrLines, imgW, imgH) {
  var words = [];
  if (!ocrLines || !imgW || !imgH) return words;
  for (var li = 0; li < ocrLines.length; li++) {
    var line = ocrLines[li];
    if (!line.words || !line.words.length) continue;
    var lineConf = line.confidence || 0;
    for (var wi = 0; wi < line.words.length; wi++) {
      var w = line.words[wi];
      var cx = w.x + w.w / 2;
      var cy = w.y + w.h / 2;
      words.push({
        text: w.text,
        normText: _normText(w.text),
        x: w.x, y: w.y, w: w.w, h: w.h,
        cx: cx, cy: cy,
        nx: cx / imgW, ny: cy / imgH,
        lineIdx: li, wordIdx: wi,
        confidence: lineConf,
        points: line.points || null
      });
    }
  }
  return words;
}

/**
 * Find words whose normalized text matches a regex.
 * Optional: filter by normalized position ranges.
 */
function _findWords(words, regex, nxMin, nxMax, nyMin, nyMax) {
  return words.filter(function(w) {
    if (!regex.test(w.normText)) return false;
    if (nxMin !== undefined && w.nx < nxMin) return false;
    if (nxMax !== undefined && w.nx > nxMax) return false;
    if (nyMin !== undefined && w.ny < nyMin) return false;
    if (nyMax !== undefined && w.ny > nyMax) return false;
    return true;
  });
}

/**
 * Given a keyword word, find the nearest amount number.
 * Looks right on same line, then on next line below.
 * Returns { value, word } or null.
 */
function _findNearbyAmount(words, kw, opts) {
  opts = opts || {};
  var maxDx = opts.maxDx || 500;  // max horizontal distance (pixels)
  var maxDy = opts.maxDy || 60;   // max vertical distance (pixels) — same/near line
  var maxDyBelow = opts.maxDyBelow || 100; // max vertical distance for next line below
  var requireRight = opts.requireRight !== false; // default true: number must be to the right of keyword

  var candidates = [];
  for (var i = 0; i < words.length; i++) {
    var w = words[i];
    if (w === kw) continue;
    // Skip low-confidence
    if (w.confidence < 0.3) continue;

    var dx = w.cx - kw.cx;
    var dy = w.cy - kw.cy;
    var ady = Math.abs(dy);

    // Same line or near line
    if (ady <= maxDy) {
      if (requireRight && dx < -20) continue; // must be to the right
      if (Math.abs(dx) > maxDx) continue;
    }
    // Next line below
    else if (dy > 0 && dy <= maxDyBelow) {
      // For below: allow slightly left but not too far
      if (dx < -kw.w * 2) continue;
      if (dx > maxDx) continue;
    }
    // Too far
    else {
      continue;
    }

    // Parse amount
    var t = w.normText.replace(/[,，]/g, '');
    var m = t.match(/^-?¥?(\d+\.\d{2})$/);
    if (m) {
      var val = parseFloat(m[1]);
      if (val > 0 && val < 1000000 && !isLikelyYearOrDate(val, t)) {
        // Score: prefer same line, then closest
        var score = ady * 2 + Math.abs(dx) * 0.5;
        candidates.push({ value: val, word: w, score: score });
      }
    }
  }
  if (!candidates.length) return null;
  candidates.sort(function(a, b) { return a.score - b.score; });
  return candidates[0];
}

/**
 * Detect invoice type from word positions.
 * Returns: 'vat' | 'ticket' | 'ride' | 'unknown'
 */
function _detectInvoiceType(words, imgW, imgH) {
  // Check for train ticket keywords in top 60%
  var topWords = words.filter(function(w) { return w.ny < 0.6; });
  var topText = topWords.map(function(w) { return w.normText; }).join('');
  if (/(?:车\s*次|票\s*价|座\s*位|席\s*别|检\s*票|进\s*站|出\s*站|铁\s*路|乘\s*车|二\s*等|一\s*等|动\s*车|高\s*铁)/.test(topText)) {
    return 'ticket';
  }
  // Check for ride-hailing keywords
  if (/(?:出\s*租|打\s*车|网\s*约|滴\s*滴|专\s*车|客\s*运\s*服\s*务)/.test(topText)) {
    return 'ride';
  }
  // Check for VAT invoice structure: "价税合计" or "购买方"+"销售方"
  var hasJiaShui = _findWords(words, /价\s*税\s*合\s*计/).length > 0;
  var hasBuyerSeller = _findWords(words, /购买方/).length > 0 && _findWords(words, /销售方/).length > 0;
  if (hasJiaShui || hasBuyerSeller) return 'vat';

  return 'unknown';
}

/**
 * Extract seller info using coordinates.
 * Strategy: find "销售方信息" or "名称:" in right half → grab name + credit code.
 */
function _extractSeller(words, imgW, imgH) {
  var sellerName = '', sellerCreditCode = '';

  // Right-half words (nx > 0.45) in top 40% (seller region)
  var sellerWords = words.filter(function(w) {
    return w.nx > 0.45 && w.ny > 0.15 && w.ny < 0.45;
  });
  var sellerText = sellerWords.map(function(w) { return w.normText; }).join('');

  // --- Credit code in seller region ---
  // Pattern 1: "纳税人识别号:" or "统一社会信用代码:" followed by code
  var ccRe = /(?:纳税人识别号|统一社会信用代码)[\/:：\s]*([A-Z0-9]{15,20})/gi;
  var ccM;
  while ((ccM = ccRe.exec(sellerText)) !== null) {
    sellerCreditCode = ccM[1].toUpperCase();
  }
  // Pattern 2: Standalone credit code (starts with digit, has letters and digits)
  if (!sellerCreditCode) {
    var sccRe = /\b([0-9][A-Z0-9]{17})\b/g;
    var sccM;
    while ((sccM = sccRe.exec(sellerText)) !== null) {
      if (/\d{6,}/.test(sccM[1]) && /[A-Z]/.test(sccM[1])) {
        sellerCreditCode = sccM[1].toUpperCase();
      }
    }
  }
  // Pattern 3: Coordinate proximity — find "纳税人识别号" label word, then find code nearby
  if (!sellerCreditCode) {
    var ccLabels = _findWords(sellerWords, /纳税人识别号|统一社会信用代码/);
    for (var ci = 0; ci < ccLabels.length && !sellerCreditCode; ci++) {
      var nearby = _findNearbyAmount(words, ccLabels[ci], { maxDx: 400, maxDy: 30, maxDyBelow: 60, requireRight: false });
      // Not an amount — look for code word
      var codeWords = words.filter(function(w) {
        if (w === ccLabels[ci]) return false;
        if (Math.abs(w.cy - ccLabels[ci].cy) > ccLabels[ci].h * 2.5) return false;
        return /^[0-9][A-Z0-9]{14,19}$/.test(w.normText.replace(/[^A-Z0-9]/g, ''));
      });
      if (codeWords.length > 0) {
        // Pick closest
        codeWords.sort(function(a, b) {
          return Math.abs(a.cx - ccLabels[ci].cx) - Math.abs(b.cx - ccLabels[ci].cx);
        });
        sellerCreditCode = codeWords[0].normText.replace(/[^A-Z0-9]/g, '').toUpperCase();
      }
    }
  }

  // --- Seller name ---
  // Pattern 1: "销售方名称:" or "销方名称:" label
  var snLabels = _findWords(sellerWords, /销售方(?:信息)?名\s*称|销\s*方(?:信息)?名\s*称/);
  if (snLabels.length > 0) {
    // Find company name near the label
    var nearbyNames = words.filter(function(w) {
      if (w === snLabels[0]) return false;
      if (Math.abs(w.cy - snLabels[0].cy) > snLabels[0].h * 2) return false;
      if (w.cx < snLabels[0].cx - 10) return false; // must be to the right
      return /[\u4e00-\u9fff]/.test(w.text); // must contain CJK
    });
    if (nearbyNames.length > 0) {
      // Concatenate adjacent name words on same line
      nearbyNames.sort(function(a, b) { return a.x - b.x; });
      var nameParts = [];
      var lastRight = 0;
      for (var ni = 0; ni < nearbyNames.length; ni++) {
        if (nearbyNames[ni].x > lastRight + nearbyNames[ni].h * 2) {
          break; // gap too big, stop
        }
        nameParts.push(nearbyNames[ni].text);
        lastRight = nearbyNames[ni].x + nearbyNames[ni].w;
      }
      if (nameParts.length > 0) {
        sellerName = nameParts.join('');
      }
    }
  }

  // Pattern 2: "名称:" in seller region (right half) — guaranteed seller
  if (!sellerName) {
    var nameLabels = _findWords(sellerWords, /^名\s*称$/);
    if (nameLabels.length > 0) {
      // There may be 2 "名称:" — one for buyer, one for seller. Pick rightmost.
      var rightNameLabel = nameLabels[nameLabels.length - 1];
      var nearbyNames2 = words.filter(function(w) {
        if (w === rightNameLabel) return false;
        if (Math.abs(w.cy - rightNameLabel.cy) > rightNameLabel.h * 2) return false;
        if (w.cx < rightNameLabel.cx - 10) return false;
        return /[\u4e00-\u9fff]/.test(w.text) && w.text.length > 1;
      });
      if (nearbyNames2.length > 0) {
        nearbyNames2.sort(function(a, b) { return a.x - b.x; });
        var nameParts2 = [];
        var lastRight2 = 0;
        for (var ni2 = 0; ni2 < nearbyNames2.length; ni2++) {
          if (nearbyNames2[ni2].x > lastRight2 + nearbyNames2[ni2].h * 2) break;
          nameParts2.push(nearbyNames2[ni2].text);
          lastRight2 = nearbyNames2[ni2].x + nearbyNames2[ni2].w;
        }
        if (nameParts2.length > 0) sellerName = nameParts2.join('');
      }
    }
  }

  // Pattern 3: Company name with suffix in seller region
  if (!sellerName) {
    var csSuffix = '(?:公司|集团|商行|商店|厂|部|院|所|中心|店|馆|站|社|行|会|处|室|局|办|坊|铺|有限合伙|合伙企业|个体工商户|个体户|工作室|经营部|门市部|分公司|事业部|事务所|医院|学校|幼儿园|合作社|企业|商社|贸易行|服务部)';
    var companyRe = new RegExp('([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]+' + csSuffix + ')');
    var companyMatch = sellerText.match(companyRe);
    if (companyMatch) sellerName = companyMatch[1].trim();
  }

  // Cleanup
  if (sellerName) {
    sellerName = sellerName.replace(/^[\s:：]+/, '').replace(/[\s:：]+$/, '');
    sellerName = sellerName.replace(/[，,。.、：:；;！!？?]+$/, '');
    sellerName = sellerName.replace(/\d{6,}$/, '');
    sellerName = sellerName.replace(/\s+[A-Z0-9]{15,20}$/, '');
    if (/^(?:购买方信息|销售方信息|购买方|销售方|名称|信息|纳税人|地址|电话|开户行|账号)$/.test(sellerName)) {
      sellerName = '';
    }
    if (sellerName.length < 2) sellerName = '';
  }

  return { sellerName: sellerName, sellerCreditCode: sellerCreditCode };
}

/**
 * v1.7.0 — Coordinate-first invoice info extraction.
 * Uses PP-OCRv5's accurate bbox to locate fields directly by position,
 * with simple regex fallback for edge cases.
 *
 * Input: { text, lines, imgW, imgH } — OCR result with coordinates
 * Output: { amountTax, amountNoTax, taxAmount, sellerName, sellerCreditCode,
 *           invoiceNo, invoiceDate, buyerName, buyerCreditCode, _ocrText, isTicket }
 */
function extractByCoordinates(ocrResult) {
  var fullText = ocrResult.text || '';
  var imgW = ocrResult.imgW || 0;
  var imgH = ocrResult.imgH || 0;
  var words = _buildWords(ocrResult.lines, imgW, imgH);

  // Normalize full text for regex fallback
  var normText = fullText;
  normText = normText.replace(/([\u4e00-\u9fff])\s+([\u4e00-\u9fff])/g, '$1$2');
  normText = normText.replace(/([\u4e00-\u9fff])\n([\u4e00-\u9fff])/g, '$1$2');
  normText = _normText(normText);
  // Collapse digit spaces
  for (var _ni = 0; _ni < 3; _ni++) {
    var _prev = '';
    while (_prev !== normText) { _prev = normText; normText = normText.replace(/(\d)\s+(\d)/g, '$1$2'); }
  }
  normText = normText.replace(/(\d)\s+\./g, '$1.');
  normText = normText.replace(/¥\s+(\d)/g, '¥$1');

  // --- Text-based extraction (PRIMARY for structured fields) ---
  // OCR text is well-formatted with clear key-value pairs — leverage this first
  var textInfo = _extractByText(fullText);
  var invoiceNo = textInfo.invoiceNo;
  var invoiceDate = textInfo.invoiceDate;
  var buyerName = textInfo.buyerName;
  var buyerCreditCode = textInfo.buyerCreditCode;

  // Detect invoice type
  var invType = _detectInvoiceType(words, imgW, imgH);
  var isTicket = invType === 'ticket';
  var sellerName = textInfo.sellerName || '';
  var sellerCreditCode = textInfo.sellerCreditCode || '';
  var amountTax = 0, amountNoTax = 0, taxAmount = 0;

  console.log('[坐标提取] 发票类型:', invType, '字数:', fullText.length, '词数:', words.length,
    '文本提取:', { invoiceNo: invoiceNo || '(空)', invoiceDate: invoiceDate || '(空)',
    buyerName: buyerName || '(空)', sellerName: sellerName || '(空)' });

  // === Ticket extraction ===
  if (isTicket) {
    if (!sellerName) sellerName = getTicketTypeLabel(fullText);

    // Method 1: "票价:" keyword → nearby amount
    var priceLabels = _findWords(words, /票\s*价/);
    for (var pi = 0; pi < priceLabels.length && !amountTax; pi++) {
      var amt = _findNearbyAmount(words, priceLabels[pi], { maxDx: 300, maxDy: 30, maxDyBelow: 80 });
      if (amt && amt.value >= 5 && amt.value <= 5000) {
        amountTax = amt.value;
      }
    }
    // "全价"/"优惠价"/"学生价"
    if (!amountTax) {
      var discountLabels = _findWords(words, /全\s*价|优\s*惠\s*价|学\s*生\s*价/);
      for (var di = 0; di < discountLabels.length && !amountTax; di++) {
        var amt2 = _findNearbyAmount(words, discountLabels[di], { maxDx: 300, maxDy: 30, maxDyBelow: 80 });
        if (amt2 && amt2.value >= 5 && amt2.value <= 5000) {
          amountTax = amt2.value;
        }
      }
    }
    // Method 2: Positional — ¥ amount in ticket area (nx < 0.5, ny 0.35~0.65)
    if (!amountTax) {
      var ticketAmounts = words.filter(function(w) {
        if (w.confidence < 0.3) return false;
        if (w.nx > 0.55 || w.ny < 0.3 || w.ny > 0.65) return false;
        var t = w.normText.replace(/[,，]/g, '');
        var m = t.match(/^-?¥?(\d+\.\d{2})$/);
        if (!m) return false;
        var v = parseFloat(m[1]);
        return v >= 5 && v <= 5000 && !isLikelyYearOrDate(v, t);
      });
      if (ticketAmounts.length > 0) {
        // Take the largest
        ticketAmounts.sort(function(a, b) {
          var va = parseFloat(a.normText.replace(/[,，¥]/g, ''));
          var vb = parseFloat(b.normText.replace(/[,，¥]/g, ''));
          return vb - va;
        });
        amountTax = parseFloat(ticketAmounts[0].normText.replace(/[,，¥]/g, ''));
      }
    }
    if (amountTax > 0) amountNoTax = amountTax;

    console.log('[坐标提取] 车票金额:', amountTax);
    return { amountTax: amountTax, amountNoTax: amountNoTax, taxAmount: 0,
             sellerName: sellerName, sellerCreditCode: sellerCreditCode,
             invoiceNo: invoiceNo, invoiceDate: invoiceDate,
             buyerName: buyerName, buyerCreditCode: buyerCreditCode,
             _ocrText: fullText, isTicket: true };
  }

  // === VAT / Ride invoice extraction ===

  // --- Seller info (coordinate-based FALLBACK — text-based is primary) ---
  if (!sellerName || !sellerCreditCode) {
    var sellerInfo = _extractSeller(words, imgW, imgH);
    if (!sellerName) sellerName = sellerInfo.sellerName;
    if (!sellerCreditCode) sellerCreditCode = sellerInfo.sellerCreditCode;
  }

  // --- Text-based amount extraction (PRIMARY for VAT invoices) ---
  // OCR text has clear structure: "合计...¥金额...¥税额...价税合计...¥含税价"
  // This is more reliable than coordinate-based matching for the "合计" row
  // which has TWO amounts that coordinate methods can't easily distinguish.
  var textAmounts = _extractAmountsByText(fullText);
  if (textAmounts.amountTax > 0) amountTax = textAmounts.amountTax;
  if (textAmounts.amountNoTax > 0) amountNoTax = textAmounts.amountNoTax;
  if (textAmounts.taxAmount > 0) taxAmount = textAmounts.taxAmount;
  console.log('[文本提取] 金额:', textAmounts);

  // --- Amount extraction (coordinate-based FALLBACK) ---
  // Only runs when text-based extraction didn't find the amounts.

  // Step 1: 价税合计（含税总价）— coordinate-based FALLBACK
  // Location: ny ≈ 0.20~0.30 (near bottom of invoice)
  // Keywords: "价税合计", "（小写）", or just ¥ at that position
  if (!amountTax) {
  var jshjLabels = _findWords(words, /价\s*税\s*合\s*计/);
  if (jshjLabels.length > 0) {
    // Use the LOWEST "价税合计" label (bottom of invoice = 含税价, not 不含税)
    jshjLabels.sort(function(a, b) { return b.ny - a.ny; });
    var amt3 = _findNearbyAmount(words, jshjLabels[0], { maxDx: 600, maxDy: 40, maxDyBelow: 120 });
    if (amt3) {
      amountTax = amt3.value;
      // Validate: if the matched amount is on the SAME line as another amount,
      // it might be the 不含税价 row (¥amount + ¥tax on same line).
      // The 含税价 is always BELOW that row. Find amounts with LARGER y (lower on page).
      var sameLineAmts = words.filter(function(w) {
        if (w.confidence < 0.3) return false;
        if (w === amt3.word) return false;
        var dy = Math.abs(w.cy - amt3.word.cy);
        if (dy > amt3.word.h * 1.5) return false; // same line
        var t = w.normText.replace(/[,，]/g, '');
        var m = t.match(/^-?¥?(\d+\.\d{2})$/);
        if (!m) return false;
        var v = parseFloat(m[1]);
        return v > 0 && v < 1000000 && !isLikelyYearOrDate(v, t);
      });
      if (sameLineAmts.length > 0) {
        // There are other amounts on the same line → this is the 不含税+税额 row
        // The 含税价 must be BELOW. Look for amounts with larger y below the keyword.
        var belowAmts = words.filter(function(w) {
          if (w.confidence < 0.3) return false;
          // Must be below the keyword (not just below the matched amount)
          var dy = w.cy - jshjLabels[0].cy;
          if (dy <= 0) return false; // must be strictly below
          if (dy > jshjLabels[0].h * 5) return false; // not too far below
          // Must NOT be on the same line as the current match (不含税+税额 row)
          if (Math.abs(w.cy - amt3.word.cy) <= amt3.word.h * 1.5) return false;
          var t = w.normText.replace(/[,，]/g, '');
          var m = t.match(/^-?¥?(\d+\.\d{2})$/);
          if (!m) return false;
          var v = parseFloat(m[1]);
          return v > 0 && v < 1000000 && !isLikelyYearOrDate(v, t);
        });
        if (belowAmts.length > 0) {
          // Take the amount with the largest y (lowest on page) = 含税价
          belowAmts.sort(function(a, b) { return b.cy - a.cy; });
          var belowVal = parseFloat(belowAmts[0].normText.replace(/[,，¥]/g, ''));
          // Sanity: 含税价 > 不含税价
          if (belowVal > amountTax) {
            amountTax = belowVal;
            console.log('[坐标提取] 价税合计同行有多个金额，已选择下方含税价:', amountTax);
          }
        }
      }
    }
  }
  } // end if (!amountTax) for Step 1

  // Step 1.5: "（小写）" keyword — very specific to 含税价
  // Key insight: 含税价 is BELOW the 不含税+税额 row, to the right of "（小写）".
  // We must prefer amounts that are BELOW "小写", not on the same line as it.
  if (!amountTax) {
    var xiaoxieLabels = _findWords(words, /小\s*写/);
    if (xiaoxieLabels.length > 0) {
      // Strategy: look for amounts strictly BELOW "小写" first
      // The 含税价 is on a line below "（小写）", not on the same line
      var xxLabel = xiaoxieLabels[0];
      var belowXx = words.filter(function(w) {
        if (w.confidence < 0.3) return false;
        var dy = w.cy - xxLabel.cy;
        // Must be below (dy > 0) and within reasonable distance
        if (dy <= xxLabel.h * 0.5 || dy > xxLabel.h * 5) return false;
        var dx = w.cx - xxLabel.cx;
        if (dx < -xxLabel.w * 2 || dx > 400) return false;
        var t = w.normText.replace(/[,，]/g, '');
        var m = t.match(/^-?¥?(\d+\.\d{2})$/);
        if (!m) return false;
        var v = parseFloat(m[1]);
        return v > 0 && v < 1000000 && !isLikelyYearOrDate(v, t);
      });
      if (belowXx.length > 0) {
        // Pick the one closest vertically (smallest dy), then horizontally
        belowXx.sort(function(a, b) {
          var da = a.cy - xxLabel.cy;
          var db = b.cy - xxLabel.cy;
          if (da !== db) return da - db;
          return Math.abs(a.cx - xxLabel.cx) - Math.abs(b.cx - xxLabel.cx);
        });
        amountTax = parseFloat(belowXx[0].normText.replace(/[,，¥]/g, ''));
        console.log('[坐标提取] 小写→下方含税价:', amountTax);
      }
      // Fallback: if no amount found below, try right side on same line
      if (!amountTax) {
        var amt4 = _findNearbyAmount(words, xxLabel, { maxDx: 400, maxDy: 30, maxDyBelow: 60 });
        if (amt4) {
          // Same validation: check if this amount shares a line with another amount
          var sameLine4 = words.filter(function(w) {
            if (w.confidence < 0.3) return false;
            if (w === amt4.word) return false;
            if (Math.abs(w.cy - amt4.word.cy) > amt4.word.h * 1.5) return false;
            var t = w.normText.replace(/[,，]/g, '');
            var m = t.match(/^-?¥?(\d+\.\d{2})$/);
            if (!m) return false;
            var v = parseFloat(m[1]);
            return v > 0 && v < 1000000 && !isLikelyYearOrDate(v, t);
          });
          if (sameLine4.length > 0) {
            // Multiple amounts on same line = 不含税+税额 row, skip this match
            console.log('[坐标提取] 小写→同行多金额(不含税行), 跳过:', amt4.value);
          } else {
            amountTax = amt4.value;
          }
        }
      }
    }
  }

  // Step 2: 不含税合计 + 税额合计 — "合计" row
  // The "合计" row typically has TWO amounts:
  //   - 不含税合计 (larger value) and 税额合计 (smaller value)
  // Since 税额 = 不含税 × 税率 and 税率 < 100%, the 不含税 is always larger.
  // We collect ALL amounts near "合计" and assign by value.
  // Must distinguish from "价税合计" — standalone "合计" without "价" to its left.
  if (!amountNoTax || !taxAmount) {
    var hejiLabels = _findWords(words, /合\s*计/);
    // Filter: standalone "合计" (no "价" or "税" nearby to the left)
    var standaloneHeji = hejiLabels.filter(function(hw) {
      if (/税/.test(hw.normText)) return false;
      var hasJiaLeft = words.some(function(w) {
        if (w === hw) return false;
        if (!/价/.test(w.normText)) return false;
        var dx = hw.cx - w.cx;
        var dy = Math.abs(w.cy - hw.cy);
        return dx >= -20 && dx < 300 && dy < 50;
      });
      return !hasJiaLeft;
    });

    for (var hi = 0; hi < standaloneHeji.length; hi++) {
      var hejiWord = standaloneHeji[hi];
      // Collect ALL amounts near "合计" (same line or slightly below)
      var rowAmts = [];
      for (var ri = 0; ri < words.length; ri++) {
        var w = words[ri];
        if (w.confidence < 0.3 || w === hejiWord) continue;
        var dy = Math.abs(w.cy - hejiWord.cy);
        if (dy > hejiWord.h * 4) continue;
        var dx = w.cx - hejiWord.cx;
        if (dx < -hejiWord.w) continue;
        var t = w.normText.replace(/[,，]/g, '');
        var m = t.match(/^-?¥?(\d+\.\d{2})$/);
        if (!m) continue;
        var v = parseFloat(m[1]);
        if (v > 0 && v < 1000000 && !isLikelyYearOrDate(v, t)) {
          rowAmts.push(v);
        }
      }

      if (rowAmts.length >= 2) {
        // Two+ amounts: largest = amountNoTax, smallest = taxAmount
        rowAmts.sort(function(a, b) { return b - a; });
        if (!amountNoTax) amountNoTax = rowAmts[0];
        if (!taxAmount) taxAmount = rowAmts[rowAmts.length - 1];
        break;
      } else if (rowAmts.length === 1 && !amountNoTax) {
        if (amountTax > 0 && rowAmts[0] > amountTax) continue;
        if (amountTax > 0 && Math.abs(rowAmts[0] - amountTax) < 0.01) continue;
        amountNoTax = rowAmts[0];
        break;
      }
    }
  }

  // Step 2.5: "金额" keyword in amount region (secondary for 不含税价)
  if (!amountNoTax) {
    // "金额" in the lower half (amount region)
    var amtLabels = _findWords(words, /金\s*额/, undefined, undefined, 0.45, 0.70);
    // Exclude "税额" and "合计金额"
    var validAmtLabels = amtLabels.filter(function(w) {
      return !/税/.test(w.normText) && !/合/.test(w.normText);
    });
    for (var ai = 0; ai < validAmtLabels.length && !amountNoTax; ai++) {
      var amt6 = _findNearbyAmount(words, validAmtLabels[ai], { maxDx: 400, maxDy: 30, maxDyBelow: 80 });
      if (amt6) {
        if (amountTax > 0 && amt6.value > amountTax) continue;
        if (amountTax > 0 && Math.abs(amt6.value - amountTax) < 0.01) continue;
        amountNoTax = amt6.value;
      }
    }
  }

  // Step 3: 税额 — "税额" keyword in amount region (coordinate-based FALLBACK)
  if (!taxAmount) {
  var seLabels = _findWords(words, /税\s*额/, undefined, undefined, 0.40, 0.75);
  if (seLabels.length > 0) {
    // Use the bottommost "税额" (in the 合计 row)
    seLabels.sort(function(a, b) { return b.ny - a.ny; });
    var amt7 = _findNearbyAmount(words, seLabels[0], { maxDx: 300, maxDy: 30, maxDyBelow: 60 });
    if (amt7) taxAmount = amt7.value;
  }
  } // end if (!taxAmount)

  // --- Cross-derivation ---
  // VAT formula: amountTax = amountNoTax + taxAmount
  if (amountTax > 0 && amountNoTax > 0 && !taxAmount) {
    taxAmount = Math.round((amountTax - amountNoTax) * 100) / 100;
  }
  if (amountTax > 0 && taxAmount > 0 && !amountNoTax && taxAmount < amountTax) {
    amountNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
  }
  if (!amountTax && amountNoTax > 0 && taxAmount > 0) {
    amountTax = Math.round((amountNoTax + taxAmount) * 100) / 100;
  }

  // --- Positional fallback: largest ¥ in amount region ---
  if (!amountTax) {
    // Amount region: lower portion of invoice (ny 0.40~0.80)
    var regionAmounts = words.filter(function(w) {
      if (w.confidence < 0.3) return false;
      if (w.ny < 0.35 || w.ny > 0.85) return false;
      var t = w.normText.replace(/[,，]/g, '');
      var m = t.match(/^-?¥?(\d+\.\d{2})$/);
      if (!m) return false;
      var v = parseFloat(m[1]);
      return v > 0 && v < 1000000 && !isLikelyYearOrDate(v, t);
    });
    if (regionAmounts.length > 0) {
      regionAmounts.sort(function(a, b) {
        var va = parseFloat(a.normText.replace(/[,，¥]/g, ''));
        var vb = parseFloat(b.normText.replace(/[,，¥]/g, ''));
        return vb - va;
      });
      var largestVal = parseFloat(regionAmounts[0].normText.replace(/[,，¥]/g, ''));
      if (amountNoTax > 0 && largestVal < amountNoTax) {
        // The largest amount in region is smaller than amountNoTax — this means
        // we didn't find amountTax in this region. Don't overwrite amountNoTax.
        // Leave amountTax unfilled and let regex fallback handle it.
      } else {
        amountTax = largestVal;
      }
    }
  }

  // --- Simple regex fallback (only when coordinates couldn't resolve) ---
  if (!amountTax) {
    amountTax = _regexFindLast('价\\s*税\\s*合\\s*计', normText);
  }
  if (!amountNoTax && amountTax > 0) {
    // Try 合计 after removing 价税合计 text
    var workText = normText.replace(/价\s*税\s*合\s*计[\s\S]*?\d+\.\d{2}/g, '');
    var hejiNum = _regexFindFirst('合\\s*计', workText);
    if (hejiNum > 0 && Math.abs(hejiNum - amountTax) > 0.01) amountNoTax = hejiNum;
  }
  if (!amountNoTax) {
    var amtNum = _regexFindFirst('金\\s*额', normText);
    if (amtNum > 0 && (amountTax === 0 || Math.abs(amtNum - amountTax) > 0.01)) amountNoTax = amtNum;
  }
  if (!taxAmount && amountTax > 0) {
    taxAmount = _regexFindFirst('税\\s*额', normText);
  }

  // --- Cross-derivation after fallback ---
  if (amountTax > 0 && amountNoTax > 0 && !taxAmount) {
    taxAmount = Math.round((amountTax - amountNoTax) * 100) / 100;
  }
  if (amountTax > 0 && taxAmount > 0 && !amountNoTax && taxAmount < amountTax) {
    amountNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
  }
  if (!amountTax && amountNoTax > 0 && taxAmount > 0) {
    amountTax = Math.round((amountNoTax + taxAmount) * 100) / 100;
  }

  // --- Invariants ---
  // 含税价 >= 不含税价
  if (amountTax > 0 && amountNoTax > 0 && amountTax < amountNoTax) {
    var _tmp = amountTax; amountTax = amountNoTax; amountNoTax = _tmp;
  }
  // amountNoTax == amountTax → only reset for VAT invoices where they should differ
  // For non-VAT invoices (no taxAmount found), they CAN be equal — don't reset
  if (amountNoTax > 0 && amountTax > 0 && Math.abs(amountNoTax - amountTax) < 0.01 && taxAmount > 0) {
    amountNoTax = 0;
  }
  // Only amountNoTax found → for non-VAT, amountTax = amountNoTax
  if (amountNoTax > 0 && !amountTax) {
    if (taxAmount > 0 && taxAmount < amountNoTax) {
      amountTax = Math.round((amountNoTax + taxAmount) * 100) / 100;
    } else {
      amountTax = amountNoTax;
    }
  }

  // --- Credit code fallback (from full text if both text-based and coordinates missed) ---
  if (!sellerCreditCode || !buyerCreditCode) {
    var ccRe = /(?:纳税人识别号|统一社会信用代码)[^A-Z0-9]{0,30}([A-Z0-9]{15,20})/gi;
    var ccM, allCc = [];
    while ((ccM = ccRe.exec(normText)) !== null) {
      var cc = ccM[1].toUpperCase();
      if (allCc.indexOf(cc) < 0) allCc.push(cc);
    }
    if (!buyerCreditCode && allCc.length >= 1) buyerCreditCode = allCc[0];
    if (!sellerCreditCode && allCc.length >= 2) sellerCreditCode = allCc[1];
    else if (!sellerCreditCode && allCc.length === 1) sellerCreditCode = allCc[0];
  }
  if (!sellerCreditCode) {
    var sccRe = /\b([0-9][A-Z0-9]{17})\b/g;
    var sccM, lastScc = '';
    while ((sccM = sccRe.exec(normText)) !== null) {
      if (/\d{6,}/.test(sccM[1]) && /[A-Z]/.test(sccM[1])) lastScc = sccM[1];
    }
    if (lastScc) sellerCreditCode = lastScc.toUpperCase();
  }

  console.log('[坐标提取] 结果:', { amountTax: amountTax, amountNoTax: amountNoTax, taxAmount: taxAmount,
    sellerName: sellerName || '(空)', sellerCreditCode: sellerCreditCode || '(空)',
    invoiceNo: invoiceNo || '(空)', invoiceDate: invoiceDate || '(空)',
    buyerName: buyerName || '(空)', buyerCreditCode: buyerCreditCode || '(空)' });

  return { amountTax: amountTax, amountNoTax: amountNoTax, taxAmount: taxAmount,
           sellerName: sellerName, sellerCreditCode: sellerCreditCode,
           invoiceNo: invoiceNo, invoiceDate: invoiceDate,
           buyerName: buyerName, buyerCreditCode: buyerCreditCode,
           _ocrText: fullText, isTicket: false };
}

/**
 * Regex helper: find first number after keyword in text.
 */
function _regexFindFirst(keyword, text) {
  var re = new RegExp(keyword + '[\\s\\S]*?(\\d+(?:,\\d{3})*\\.\\d{2})');
  var m = text.match(re);
  if (!m) return 0;
  var v = parseAmt(m[1]);
  if (isLikelyYearOrDate(v, m[1])) return 0;
  return v;
}

/**
 * Regex helper: find LAST number after keyword in text.
 */
function _regexFindLast(keyword, text) {
  var re = new RegExp(keyword + '[\\s\\S]*?(\\d+(?:,\\d{3})*\\.\\d{2})', 'g');
  var m, lastVal = 0;
  while ((m = re.exec(text)) !== null) {
    var v = parseAmt(m[1]);
    if (v > 0 && !isLikelyYearOrDate(v, m[1])) lastVal = v;
  }
  return lastVal;
}
