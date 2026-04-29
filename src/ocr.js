// =====================================================
// OCR & Invoice Info Extraction
// =====================================================
// Dependencies (global): invoke, isTauri, pdfjsLib, dataUrlToUint8Array

/**
 * Parse amount string to number (2 decimal places)
 */
function parseAmt(s) {
  if (!s) return 0;
  var n = parseFloat(s.replace(/,/g, ''));
  return (!isNaN(n) && n > 0) ? Math.round(n * 100) / 100 : 0;
}

/**
 * Extract invoice info from OCR text or PDF.js text content.
 * Returns: { amountTax, amountNoTax, sellerName, sellerCreditCode, _ocrText }
 *
 * OCR patterns handled:
 *   VAT invoices: 价税合计(大写)…(小写) ¥317.00, 合计(小写) ¥299.06
 *   Regular invoices: 金额, 总计
 *   Train tickets: 票价, ¥ amount near train keywords
 *   Ride-hailing: 实付金额, 应付金额
 *
 * KEY RULE: All amounts have exactly 2 decimal places
 */
function extractInvoiceInfo(textContent) {
  var fullText;
  if (typeof textContent === 'string') {
    fullText = textContent;
  } else if (textContent && textContent.items && textContent.items.length) {
    fullText = textContent.items.map(function(item) { return item.str; }).join('');
  } else {
    return { amountTax: 0, amountNoTax: 0, sellerName: '', sellerCreditCode: '' };
  }

  // === Extract seller info BEFORE normalization (uses raw text with line breaks) ===
  var sellerName = '', sellerCreditCode = '';

  // Credit code: find ALL matches, take the LAST one (seller's in standard invoices where buyer comes first)
  var ccRe = /(?:统一社会信用代码|纳税人识别号)\s*[:：／/]?\s*([A-Z0-9]{15,20})/gi;
  var ccM, lastCcCode = '', lastCcPos = -1;
  while ((ccM = ccRe.exec(fullText)) !== null) {
    lastCcCode = ccM[1];
    lastCcPos = ccM.index;
  }
  if (lastCcCode) sellerCreditCode = lastCcCode.toUpperCase();

  // Seller name — three strategies in priority order:
  // Strategy 1: Direct "销售方(+信息)" + "名称:" pattern (most specific)
  var snMatch = fullText.match(/销售方(?:信息)?\s*名称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号)|\n|$)/i);
  if (snMatch) {
    sellerName = snMatch[1].trim();
  }
  // Strategy 2: Find the LAST "名称:" near the LAST credit code (seller's section)
  if (!sellerName) {
    if (lastCcPos >= 0) {
      var searchRegion = fullText.substring(0, lastCcPos);
      var nameRe = /名称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号)|\n|$)/gi;
      var nm, lastName = '';
      while ((nm = nameRe.exec(searchRegion)) !== null) {
        lastName = nm[1];
      }
      if (lastName) sellerName = lastName.trim();
    } else {
      // No credit code found at all — try any "名称:" in full text
      var fallbackMatch = fullText.match(/名称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号)|\n|$)/i);
      if (fallbackMatch) sellerName = fallbackMatch[1].trim();
    }
  }

  // === Normalize text for amount extraction ===
  fullText = fullText.replace(/([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])\s+([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])/g, '$1$2');
  fullText = fullText.replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  fullText = fullText.replace(/[Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  fullText = fullText.replace(/％/g, '%').replace(/．/g, '.').replace(/，/g, ',').replace(/：/g, ':');

  // === Normalize OCR digit/symbol artifacts ===
  fullText = fullText.replace(/(\d)\s+(\d)/g, '$1$2');
  fullText = fullText.replace(/(\d)\s+\./g, '$1.');
  fullText = fullText.replace(/¥\s+(\d)/g, '¥$1');
  fullText = fullText.replace(/([\u4e00-\u9fff])\s+¥/g, '$1¥');

  // Helper: find first number with exactly 2 decimal places after a keyword
  function findFirstNum(keyword, text) {
    var re = new RegExp(keyword + '[^\\d]*?(\\d+(?:,\\d{3})*\\.\\d{2})');
    var m = text.match(re);
    return m ? parseAmt(m[1]) : 0;
  }

  var amountTax = 0, amountNoTax = 0;

  // === Step 1: 价税合计 → 含税总价 ===
  amountTax = findFirstNum('价\\s*税\\s*合\\s*计\\s*[（(]\\s*大\\s*写\\s*[）)][^\\d]*?[（(]\\s*小\\s*写\\s*[）)]', fullText);
  if (!amountTax) amountTax = findFirstNum('价\\s*税\\s*合\\s*计\\s*[（(]\\s*小\\s*写\\s*[）)]', fullText);
  if (!amountTax) amountTax = findFirstNum('价\\s*税\\s*合\\s*计', fullText);

  // === Step 1.5: 电车票/打车电子发票 ===
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('实\\s*付\\s*金\\s*额', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('应\\s*付\\s*金\\s*额', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  if (!amountTax && !amountNoTax) {
    var amtLineMatch = fullText.match(/(?:合\s*计|总\s*计|计\s*费)[^\n]*?金\s*额[^\d]*?(\d+(?:,\d{3})*\.\d{2})/);
    if (amtLineMatch) {
      var val3 = parseAmt(amtLineMatch[1]);
      if (val3 > 0) { amountTax = val3; amountNoTax = val3; }
    }
  }

  // === Step 2: Remove ALL 价税合计 variants ===
  var workText = fullText.replace(/价\s*税\s*合\s*计\s*(?:[（(](?:\s*大\s*写\s*[^\d]*?)?\s*小\s*写\s*[）)]\s*)?[^\d]*?\d+(?:,\d{3})*\.\d{2}/g, '');

  // === Step 3: 合计 → 不含税价 ===
  if (!amountNoTax) amountNoTax = findFirstNum('合\\s*计', workText);

  // === Step 4: 金额 → 不含税价 ===
  if (!amountNoTax) amountNoTax = findFirstNum('金\\s*额', workText);

  // === Step 5: 税额反推 ===
  if (amountTax > 0 && !amountNoTax) {
    var tax = findFirstNum('税\\s*额', fullText);
    if (tax > 0 && tax < amountTax) amountNoTax = Math.round((amountTax - tax) * 100) / 100;
  }

  // === Step 6: 火车票 — robust detection ===
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('票\\s*价', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  if (!amountTax && !amountNoTax) {
    var ticketMatch = fullText.match(/票[^\d]*?¥\s*(\d+(?:,\d{3})*\.\d{2})/);
    if (ticketMatch) {
      var val = parseAmt(ticketMatch[1]);
      if (val > 0 && val < 10000) { amountTax = val; amountNoTax = val; }
    }
  }
  if (!amountTax && !amountNoTax) {
    var yMatch = fullText.match(/¥\s*(\d+(?:,\d{3})*\.\d{2})\s*[^\d]*?(?:车票|车次|座位|检票)/);
    if (!yMatch) yMatch = fullText.match(/(?:车票|车次|座位|检票)[^\d]*?¥\s*(\d+(?:,\d{3})*\.\d{2})/);
    if (yMatch) {
      var val2 = parseAmt(yMatch[1]);
      if (val2 > 0 && val2 < 10000) { amountTax = val2; amountNoTax = val2; }
    }
  }
  if (!amountTax && !amountNoTax) {
    var seatMatch = fullText.match(/(?:席\s*别|座\s*位)[^\d]*?(\d+(?:,\d{3})*\.\d{2})/);
    if (seatMatch) {
      var val3 = parseAmt(seatMatch[1]);
      if (val3 > 0 && val3 < 10000) { amountTax = val3; amountNoTax = val3; }
    }
  }
  if (!amountTax && !amountNoTax) {
    var trainKwRe = /(?:车\s*次|车\s*站|号\s*车|检\s*票|铺\s*位|卧\s*铺|二\s*等|一\s*等|动\s*车|高\s*铁|特\s*等)/;
    if (trainKwRe.test(fullText)) {
      var allYen = [];
      var yenRe = /¥\s*(\d+(?:,\d{3})*\.\d{2})/g;
      var ym;
      while ((ym = yenRe.exec(fullText)) !== null) {
        var yv = parseAmt(ym[1]);
        if (yv > 0 && yv < 5000) allYen.push(yv);
      }
      if (allYen.length === 1) {
        amountTax = allYen[0]; amountNoTax = allYen[0];
      } else if (allYen.length > 1) {
        var typicalYen = allYen.filter(function(y) { return y >= 20 && y <= 2000; });
        if (typicalYen.length === 1) { amountTax = typicalYen[0]; amountNoTax = typicalYen[0]; }
        else if (typicalYen.length > 1) { amountTax = Math.max.apply(Math, typicalYen); amountNoTax = amountTax; }
      }
    }
  }

  // === Step 6.5: Super fallback — any ¥ amount ===
  if (!amountTax && !amountNoTax) {
    var yenRe2 = /¥\s*(\d+(?:,\d{3})*\.\d{2})/g;
    var amounts = [], ym2;
    while ((ym2 = yenRe2.exec(fullText)) !== null) {
      var yv2 = parseAmt(ym2[1]);
      if (yv2 > 0 && yv2 < 50000) amounts.push(yv2);
    }
    if (amounts.length > 0) {
      var maxAmt = Math.max.apply(Math, amounts);
      amountTax = maxAmt; amountNoTax = maxAmt;
    }
  }

  // === Step 7: Generic fallback ===
  if (!amountTax && !amountNoTax) {
    var fb = findFirstNum('总\\s*计', fullText);
    if (fb > 0) { amountTax = fb; amountNoTax = fb; }
  }
  if (!amountTax && !amountNoTax) {
    var fb2 = findFirstNum('应\\s*收', fullText);
    if (fb2 > 0) { amountTax = fb2; amountNoTax = fb2; }
  }

  // AUTO-FALLBACK: if only amountNoTax found, auto-assign to amountTax
  if (amountNoTax > 0 && amountTax === 0) {
    amountTax = amountNoTax;
  }

  console.log('[OCR提取] 金额:', { amountTax: amountTax, amountNoTax: amountNoTax }, '销售方:', sellerName || '(未识别)', '信用代码:', sellerCreditCode || '(未识别)');
  if (!amountTax && !amountNoTax && !sellerName) {
    console.warn('[OCR提取] 未能识别任何信息，OCR完整文本:', fullText);
  }

  return { amountTax: amountTax, amountNoTax: amountNoTax, sellerName: sellerName, sellerCreditCode: sellerCreditCode, _ocrText: fullText };
}

/**
 * Legacy wrapper that only returns amounts
 */
function extractInvoiceAmount(textContent) {
  var info = extractInvoiceInfo(textContent);
  return { amountTax: info.amountTax, amountNoTax: info.amountNoTax };
}

/**
 * Extract invoice info from all pages of a PDF (via data URL).
 * Returns array of full info objects (amounts + seller).
 */
async function tryExtractPdfInfo(dataUrl, pageCount) {
  if (!window.pdfjsLib) return [];
  try {
    var raw = dataUrlToUint8Array(dataUrl);
    var pdf = await pdfjsLib.getDocument({
      data: raw,
      cMapUrl: CMAP_BASE_URL,
      cMapPacked: true,
      standardFontDataUrl: STD_FONT_BASE_URL,
      disableFontFace: true, useSystemFonts: false
    }).promise;
    var results = [];
    for (var p = 1; p <= pdf.numPages; p++) {
      var page = await pdf.getPage(p);
      var textContent = await page.getTextContent();
      results.push(extractInvoiceInfo(textContent));
    }
    return results;
  } catch(e) {
    console.warn('[信息提取] PDF文字提取失败:', e);
    return [];
  }
}

/**
 * Apply OCR to a file object — unified helper replacing 4 duplicate code blocks.
 * Modifies fileObj in place, adding amount/seller info if detected.
 * @param {Object} fileObj - The file object to update
 * @param {string} dataUrl - Base64 data URL of the image to OCR
 */
async function applyOcr(fileObj, dataUrl) {
  if (!isTauri || !invoke) return;
  try {
    var ocrText = await invoke('ocr_image', { dataUrl: dataUrl });
    if (!ocrText) return;
    var info = extractInvoiceInfo(ocrText);
    var effAmt = info.amountTax > 0 ? info.amountTax : info.amountNoTax;
    if (effAmt > 0) {
      fileObj.amount = effAmt;
      fileObj.amountTax = info.amountTax;
      fileObj.amountNoTax = info.amountNoTax;
    }
    if (info.sellerName) fileObj.sellerName = info.sellerName;
    if (info.sellerCreditCode) fileObj.sellerCreditCode = info.sellerCreditCode;
    fileObj._ocrText = info._ocrText || ocrText;
  } catch(e) {
    console.warn('[OCR] 识别失败:', e);
  }
}
