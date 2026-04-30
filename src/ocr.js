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

/**
 * Extract invoice info from OCR text or PDF.js text content.
 * Supports two input modes:
 *   1. Plain string (legacy PDF.js path)
 *   2. Object { text, lines, imgW, imgH } (new OCR-with-coordinates path)
 *
 * Returns: { amountTax, amountNoTax, sellerName, sellerCreditCode, _ocrText, isTicket }
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
  var wordMap = null;   // null = no coordinates available
  var imgW = 0, imgH = 0;

  if (typeof textContent === 'string') {
    fullText = textContent;
  } else if (textContent && textContent.items && textContent.items.length) {
    // PDF.js text content — build a pseudo-word-map from transform coordinates
    fullText = textContent.items.map(function(item) { return item.str; }).join('');
    if (textContent.items[0] && textContent.items[0].transform) {
      imgW = 0; imgH = 0;
      var pdfItems = textContent.items;
      // Estimate page dimensions from item positions
      var maxX = 0, maxY = 0;
      pdfItems.forEach(function(item) {
        if (item.transform) {
          var x = Math.abs(item.transform[4]);
          var y = Math.abs(item.transform[5]);
          var w = (item.width || 0);
          var h = Math.abs(item.transform[0]) || Math.abs(item.transform[3]) || 12;
          if (x + w > maxX) maxX = x + w;
          if (y + h > maxY) maxY = y + h;
        }
      });
      imgW = maxX || 1;
      imgH = maxY || 1;
      wordMap = [];
      pdfItems.forEach(function(item, idx) {
        if (!item.transform || !item.str) return;
        var x = Math.abs(item.transform[4]);
        var y = Math.abs(item.transform[5]);
        var w = item.width || (item.str.length * 12);
        var h = Math.abs(item.transform[0]) || 12;
        wordMap.push({
          text: item.str,
          x: x, y: imgH - y,  // flip Y (PDF origin is bottom-left)
          w: w, h: h,
          region: classifyRegion(x, imgH - y, w, h, imgW, imgH),
          lineIdx: idx,
          wordIdx: 0
        });
      });
    }
  } else if (textContent && typeof textContent === 'object' && textContent.text) {
    // New OCR-with-coordinates format
    fullText = textContent.text;
    imgW = textContent.imgW || 0;
    imgH = textContent.imgH || 0;
    if (textContent.lines && imgW > 0 && imgH > 0) {
      wordMap = buildWordMap(textContent.lines, imgW, imgH);
    }
  } else {
    return { amountTax: 0, amountNoTax: 0, sellerName: '', sellerCreditCode: '', isTicket: false };
  }

  // === Detect ticket type early ===
  var isTicket = isTicketText(fullText);
  var sellerName = '', sellerCreditCode = '';

  // === Skip seller extraction for tickets (train/ride tickets have no seller) ===
  if (!isTicket) {
    // === Extract seller info BEFORE normalization (uses raw text with line breaks) ===

    // === Strategy 0: Coordinate-based region analysis (highest priority) ===
    // If we have word coordinates, we can reliably identify the seller region
    // (right half of top section) and extract name + credit code from there only.
    if (wordMap && imgW > 0 && imgH > 0) {
      var sellerWords = wordMap.filter(function(w) { return w.region === 'seller'; });
      var sellerText = getRegionText(wordMap, 'seller');
      var buyerText = getRegionText(wordMap, 'buyer');

      // Extract credit code from seller region only
      if (sellerText) {
        var ccSellerRe = /(?:统一社会信用代码|纳税人识别号)\s*[:：／/]\s*([A-Z0-9]{15,20})/gi;
        var ccSellerM;
        while ((ccSellerM = ccSellerRe.exec(sellerText)) !== null) {
          sellerCreditCode = ccSellerM[1].toUpperCase();
        }
        // Standalone credit code in seller region
        if (!sellerCreditCode) {
          var sccSellerRe = /([0-9][A-Z0-9]{17})\b/g;
          var sccSellerM;
          while ((sccSellerM = sccSellerRe.exec(sellerText)) !== null) {
            if (/^[0-9]/.test(sccSellerM[1]) && /[A-Z]/.test(sccSellerM[1]) && /\d{6,}/.test(sccSellerM[1])) {
              sellerCreditCode = sccSellerM[1].toUpperCase();
            }
          }
        }

        // Extract seller name from seller region only
        // Try "销售方名称:" first
        var sn0 = sellerText.match(/销\s*售\s*方(?:信息)?\s*名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
        if (sn0) {
          sellerName = sn0[1].trim();
          sellerName = sellerName.replace(/\s*(?:购买方|销售方|信息|名称|纳税人|统一社会|地址|开户行|电话|账号).*$/i, '');
        }
        // Try "销方名称:"
        if (!sellerName) {
          var sn05 = sellerText.match(/销\s*方(?:信息)?\s*名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
          if (sn05) sellerName = sn05[1].trim();
        }
        // Try "名称:" in seller region (no ambiguity — it's guaranteed to be the seller)
        if (!sellerName) {
          var nameInSeller = sellerText.match(/名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
          if (nameInSeller) {
            var nc = nameInSeller[1].trim();
            if (nc.length > 1 && !/^(?:购买方|信息|名称)/.test(nc)) {
              sellerName = nc;
            }
          }
        }
        // Try company name with suffix in seller region
        if (!sellerName) {
          var cs0 = '(?:公司|集团|商行|商店|厂|部|院|所|中心|店|馆|站|社|行|会|处|室|局|办|坊|铺|有限合伙|合伙企业|个体工商户|个体户|工作室|经营部|门市部|分公司|事业部|事务所|医院|学校|幼儿园|合作社|企业|商社|贸易行|服务部)';
          var companyRe0 = new RegExp('([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]+' + cs0 + ')');
          var companyInSeller = sellerText.match(companyRe0);
          if (companyInSeller) {
            sellerName = companyInSeller[1].trim();
          }
        }
      }

      // Also try to extract credit code using coordinate proximity:
      // Find "纳税人识别号" / "统一社会信用代码" labels and grab the code nearby
      if (!sellerCreditCode && sellerWords.length > 0) {
        var ccLabelWords = sellerWords.filter(function(w) {
          return /(?:纳税人识别号|统一社会信用代码)/.test(w.text);
        });
        if (ccLabelWords.length > 0) {
          // Find the rightmost label (seller section is on the right)
          var rightmostLabel = ccLabelWords[ccLabelWords.length - 1];
          // Find code words near this label (same line or next, to the right)
          var nearbyCodes = sellerWords.filter(function(w) {
            return w !== rightmostLabel &&
              /[A-Z0-9]{15,20}/.test(w.text) &&
              Math.abs(w.y - rightmostLabel.y) < rightmostLabel.h * 2 &&
              w.x > rightmostLabel.x - 50;
          });
          if (nearbyCodes.length > 0) {
            sellerCreditCode = nearbyCodes[0].text.replace(/[^A-Z0-9]/g, '').toUpperCase();
          }
        }
      }

      console.log('[OCR坐标] 销售方区域文字:', sellerText ? sellerText.substring(0, 200) : '(空)');
      console.log('[OCR坐标] 购买方区域文字:', buyerText ? buyerText.substring(0, 100) : '(空)');
    }

    // Credit code: find ALL matches, take the LAST one (seller's in standard invoices where buyer comes first)
    var ccRe = /(?:统一社会信用代码|纳税人识别号)\s*[:：／/]\s*([A-Z0-9]{15,20})/gi;
    var ccM, lastCcCode = '', lastCcPos = -1;
    var firstCcPos = -1;
    var allCcPositions = []; // track all credit code positions for region analysis
    while ((ccM = ccRe.exec(fullText)) !== null) {
      if (firstCcPos < 0) firstCcPos = ccM.index;
      allCcPositions.push({ code: ccM[1], pos: ccM.index });
      lastCcCode = ccM[1];
      lastCcPos = ccM.index;
    }
    // Only override if we didn't get it from coordinates
    if (!sellerCreditCode && lastCcCode) sellerCreditCode = lastCcCode.toUpperCase();

    // Also try standalone credit codes without the prefix label (some OCR misses the label)
    if (!sellerCreditCode) {
      var standaloneCcRe = /([0-9][A-Z0-9]{17})\b/g;
      var sccM;
      while ((sccM = standaloneCcRe.exec(fullText)) !== null) {
        if (/^[0-9]/.test(sccM[1]) && /[A-Z]/.test(sccM[1]) && /\d{6,}/.test(sccM[1])) {
          lastCcCode = sccM[1];
          lastCcPos = sccM.index;
          if (firstCcPos < 0) firstCcPos = sccM.index;
          allCcPositions.push({ code: sccM[1], pos: sccM.index });
        }
      }
      if (lastCcCode) sellerCreditCode = lastCcCode.toUpperCase();
    }

    // Company name suffixes — comprehensive list for Chinese business names
    var companySuffix = '(?:公司|集团|商行|商店|厂|部|院|所|中心|店|馆|站|社|行|会|处|室|局|办|坊|铺|有限合伙|合伙企业|个体工商户|个体户|工作室|经营部|门市部|分公司|事业部|事务所|医院|学校|幼儿园|合作社|企业|商社|贸易行|服务部)';

    // Seller name extraction — anchored by credit code position (same principle as credit codes):
    // Credit codes are accurate because "纳税人识别号" prefix is unique → exactly 2 matches → last = seller.
    // For names, "名称:" is too generic — remarks can also contain "名称:".
    // Solution: use credit code position as anchor. The seller's "名称:" is always
    // BEFORE the seller's credit code in the same section. Remarks come AFTER.
    // So: find the LAST "名称:" before the seller's credit code position → seller's name.

    // Strategy 1: Direct "销售方(+信息)" + "名称:" pattern (most specific)
    var snMatch = fullText.match(/销\s*售\s*方(?:信息)?\s*名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
    if (snMatch) {
      sellerName = snMatch[1].trim();
      sellerName = sellerName.replace(/\s*(?:购买方|销售方|信息|名称|纳税人|统一社会|地址|开户行|电话|账号).*$/i, '');
    }

    // Strategy 1.5: "销方" abbreviated form
    if (!sellerName) {
      var shortSellerMatch = fullText.match(/销\s*方(?:信息)?\s*名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
      if (shortSellerMatch) {
        sellerName = shortSellerMatch[1].trim();
      }
    }

    // Strategy 2: "名称:" anchored by credit code position
    // Same principle as credit codes: use a reliable structural anchor.
    // Credit codes use "纳税人识别号" prefix as anchor → last match = seller.
    // Names use credit code POSITION as anchor → last "名称:" before seller's credit code = seller's name.
    if (!sellerName && lastCcPos >= 0) {
      var textBeforeCc = fullText.substring(0, lastCcPos);
      var nameRe2 = /名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/gi;
      var nm2, lastNameBeforeCc = '';
      while ((nm2 = nameRe2.exec(textBeforeCc)) !== null) {
        var candidate = nm2[1].trim();
        if (!/^(?:购买方|信息|名称)/.test(candidate) && candidate.length > 1) {
          lastNameBeforeCc = candidate;
        }
      }
      if (lastNameBeforeCc) sellerName = lastNameBeforeCc;
    }

    // Strategy 2.5: No credit code found — fall back to 2nd "名称:" match
    // (1st = buyer, 2nd = seller, remarks add 3+ which we skip)
    if (!sellerName) {
      var nameRe3 = /名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/gi;
      var nm3, allNames = [];
      while ((nm3 = nameRe3.exec(fullText)) !== null) {
        var c = nm3[1].trim();
        if (!/^(?:购买方|信息|名称)/.test(c) && c.length > 1) {
          allNames.push(c);
        }
      }
      if (allNames.length >= 2) sellerName = allNames[1];
      else if (allNames.length === 1) sellerName = allNames[0];
    }

    // Strategy 3: "收款单位"/"销货单位"/"开票方" pattern (non-standard invoice formats)
    if (!sellerName) {
      var altSellerMatch = fullText.match(/(?:收款单位|销货单位|开票方|销售单位|开票人|代开企业)[^\n]{0,30}?[:：]?\s*([^\n]{2,60}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
      if (altSellerMatch) {
        var altCand = altSellerMatch[1].trim();
        if (altCand.length > 1 && !/^(?:购买方|信息|名称|地址|电话)/.test(altCand)) {
          sellerName = altCand;
        }
      }
    }

    // Strategy 4: Company name near the last credit code
    // Some OCR outputs have: "91440300xxxxxxxxx  深圳市某某科技有限公司"
    if (!sellerName && lastCcPos >= 0) {
      var afterLastCc = fullText.substring(lastCcPos);
      var companyRe = new RegExp('(?:[A-Z0-9]{15,20})\\s*[:：]?\\s*([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]+' + companySuffix + ')');
      var compMatch = afterLastCc.match(companyRe);
      if (compMatch) {
        sellerName = compMatch[1].trim();
      }
    }

    // Strategy 5: Chinese company name after "销售方" keyword
    if (!sellerName) {
      var sellerRegionMatch = fullText.match(/销\s*售\s*方[^\n]{0,200}/i);
      if (sellerRegionMatch) {
        var region = sellerRegionMatch[0];
        var companyInRegion = region.match(new RegExp('([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]+' + companySuffix + ')'));
        if (companyInRegion) {
          sellerName = companyInRegion[1].trim();
        }
      }
    }

    // Strategy 5.5: Company name after "销方" keyword (short form)
    if (!sellerName) {
      var shortSellerRegion = fullText.match(/销\s*方[^\n]{0,200}/i);
      if (shortSellerRegion) {
        var region2 = shortSellerRegion[0];
        var companyInRegion2 = region2.match(new RegExp('([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]+' + companySuffix + ')'));
        if (companyInRegion2) {
          sellerName = companyInRegion2[1].trim();
        }
      }
    }

    // Strategy 6: Last resort — find company name patterns in full text
    if (!sellerName) {
      var allCompanies = [];
      var companyGlobalRe = new RegExp('([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]{2,25}' + companySuffix + ')', 'g');
      var cm2;
      while ((cm2 = companyGlobalRe.exec(fullText)) !== null) {
        var cname = cm2[1].trim();
        if (cname.length > 3 && !/^(?:购买方|销售方|信息|名称|地址)/.test(cname)) {
          allCompanies.push(cname);
        }
      }
      if (allCcPositions.length >= 1 && allCompanies.length >= 1) {
        sellerName = allCompanies[allCompanies.length - 1];
      } else if (allCompanies.length >= 2) {
        sellerName = allCompanies[allCompanies.length - 1];
      }
    }

    // Final cleanup: remove common OCR artifacts and section labels
    if (sellerName) {
      sellerName = sellerName.replace(/^[\s:：]+/, '').replace(/[\s:：]+$/, '');
      // Remove if the name is just a section label
      if (/^(?:购买方信息|销售方信息|购买方|销售方|名称|信息|纳税人|地址|电话|开户行|账号)$/.test(sellerName)) {
        sellerName = '';
      }
      // Remove trailing punctuation artifacts
      sellerName = sellerName.replace(/[，,。.、：:；;！!？?]+$/, '');
      // Remove trailing digits that are likely OCR noise
      sellerName = sellerName.replace(/\d{6,}$/, '');
      // Remove trailing credit code fragments
      sellerName = sellerName.replace(/\s+[A-Z0-9]{15,20}$/, '');
      sellerName = sellerName.trim();
      // Remove if too short after cleanup
      if (sellerName.length < 2) sellerName = '';
    }
  } else {
    // Ticket: set a descriptive type label instead of leaving seller blank
    sellerName = getTicketTypeLabel(fullText);
  }

  // === Normalize text for amount extraction ===
  // PP-OCRv5 may insert spaces between CJK characters (especially low-confidence regions)
  // and may split lines at arbitrary positions. Collapse CJK inter-character spaces first.
  fullText = fullText.replace(/([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])\s+([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])/g, '$1$2');
  // Also collapse newlines between CJK chars (PP-OCRv5 may split keywords across lines)
  fullText = fullText.replace(/([\u4e00-\u9fff])\n([\u4e00-\u9fff])/g, '$1$2');
  // Fullwidth → halfwidth normalization
  fullText = fullText.replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  fullText = fullText.replace(/[Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  fullText = fullText.replace(/％/g, '%').replace(/．/g, '.').replace(/，/g, ',').replace(/：/g, ':');
  // PP-OCRv5 may produce middle dot (· U+00B7) or bullet (• U+2022) instead of decimal point
  fullText = fullText.replace(/(\d)[·•‧∙](\d)/g, '$1.$2');
  // PP-OCRv5 may produce "O" (letter) instead of "0" (digit) in amounts
  // Only replace "O" between digits (e.g., "3O7.00" → "307.00", but not "PO Box")
  fullText = fullText.replace(/(\d)O(\d)/g, '$10$2');

  // === Normalize OCR digit/symbol artifacts ===
  // PP-OCRv5 may produce spaces inside numbers: "3 17.00" or "31 7.00" or "¥ 317.00"
  // Iteratively collapse digit spaces (multiple passes for "3 1 7.00" → "317.00")
  for (var _ni = 0; _ni < 3; _ni++) {
    var prev = '';
    while (prev !== fullText) {
      prev = fullText;
      fullText = fullText.replace(/(\d)\s+(\d)/g, '$1$2');
    }
  }
  fullText = fullText.replace(/(\d)\s+\./g, '$1.');
  fullText = fullText.replace(/¥\s+(\d)/g, '¥$1');
  fullText = fullText.replace(/￥\s+(\d)/g, '￥$1');
  fullText = fullText.replace(/([\u4e00-\u9fff])\s+¥/g, '$1¥');
  fullText = fullText.replace(/([\u4e00-\u9fff])\s+￥/g, '$1￥');

  // === Normalize OCR ¥↔1 misread — ¥¥ pattern (BEFORE keyword-based rule) ===
  // OCR misreads "1" as "¥" producing "¥¥72.68" (should be "¥172.68")
  // and "￥¥07.00" (should be "¥107.00"). Must run before the keyword-based
  // "1→¥" rule to avoid re-creating double-¥ patterns.
  fullText = normalizeOcrCurrency(fullText);

  // === Normalize OCR ¥→1 misread (keyword-based) ===
  // OCR often misreads "¥" as "1" (they look very similar in many fonts).
  // Pattern: "1XXX.XX" (4+ digits before decimal) right after amount keywords
  // → should be "¥XXX.XX"
  // e.g., "价税合计1317.00" → "价税合计¥317.00"
  // Only apply after amount-related keywords to avoid corrupting legitimate numbers.
  // [^\d¥￥\n]*? excludes ¥ and newlines so we don't re-create double-¥ or match too far
  // like "金额¥172.68" → without this exclusion, "1" in "172" would be replaced → "金额¥¥72.68"
  // \d{3,} requires 3+ digits after "1" (4+ total) to avoid stripping "1" from
  // legitimate 3-digit amounts like "金额172.68" (should stay 172.68, not become ¥72.68)
  fullText = fullText.replace(/(价\s*税\s*合\s*计|金\s*额|税\s*额|合\s*计|票\s*价|总\s*计|不\s*含\s*税|含\s*税|实\s*付|应\s*付|开\s*票\s*金\s*额|发\s*票\s*金\s*额|全\s*价|优\s*惠\s*价)([^\d¥￥\n]*?)1(\d{3,}\.\d{2})/g, '$1$2¥$3');

  // Helper: find first number with exactly 2 decimal places after a keyword
  // Uses [\\s\\S]*? instead of just .*? so it can match across newlines
  // (PP-OCRv5 may split keywords and values across multiple lines)
  function findFirstNum(keyword, text) {
    var re = new RegExp(keyword + '[\\s\\S]*?(\\d+(?:,\\d{3})*\\.\\d{2})');
    var m = text.match(re);
    if (!m) return 0;
    var v = parseAmt(m[1]);
    // Filter out year-like values (e.g., "2025.01" from dates)
    if (isLikelyYearOrDate(v, m[1])) return 0;
    return v;
  }

  // Helper: find LAST number with exactly 2 decimal places after a keyword.
  // For 含税价 (价税合计), the amount is at the BOTTOM of the invoice.
  // OCR reads top-to-bottom, so the LAST match after "价税合计" is most likely
  // the actual 含税价 number (not 不含税价 from a row above).
  // Uses [\\s\\S]*? for cross-newline matching (PP-OCRv5 line splits).
  function findLastNum(keyword, text) {
    var re = new RegExp(keyword + '[\\s\\S]*?(\\d+(?:,\\d{3})*\\.\\d{2})', 'g');
    var m, lastVal = 0, lastRaw = '';
    while ((m = re.exec(text)) !== null) {
      var v = parseAmt(m[1]);
      if (v > 0 && !isLikelyYearOrDate(v, m[1])) {
        lastVal = v;
        lastRaw = m[1];
      }
    }
    return lastVal;
  }

  /**
   * Coordinate-aware amount extraction: find a number near a keyword
   * by checking word proximity — same line (to the right) OR next line below
   * (for table layouts where keyword is a column header and value is below).
   * Returns the amount or 0.
   */
  function findAmountNearKeyword(keywordRegex, regionFilter, maxLineDist, preferBelow) {
    if (!wordMap || !imgW || !imgH) return 0;
    maxLineDist = maxLineDist || 80; // max vertical distance for "next line" (increased from 30)
    var candidates = wordMap.filter(function(w) {
      if (regionFilter && w.region !== regionFilter && regionFilter !== 'any') return false;
      // Skip low-confidence keywords
      if (w.confidence !== undefined && w.confidence < 0.3) return false;
      return keywordRegex.test(w.text);
    });
    if (!candidates.length) return 0;
    // For each keyword match, find the nearest number word
    for (var ci = 0; ci < candidates.length; ci++) {
      var kw = candidates[ci];
      // Find number words on the same line OR directly below
      var nearbyNums = wordMap.filter(function(w) {
        if (w === kw) return false;
        // Skip low-confidence number words
        if (w.confidence !== undefined && w.confidence < 0.3) return false;
        var dy = w.y - kw.y;
        var ady = Math.abs(dy);
        // Same line (within half maxLineDist): must be to the right or very close
        if (ady <= maxLineDist * 0.5) {
          if (w.x < kw.x - 20) return false;
        }
        // Next line below: must be roughly in the same column
        else if (dy > 0 && dy <= maxLineDist) {
          // Value should be roughly aligned with the keyword column
          if (w.x < kw.x - kw.w * 2 || w.x > kw.x + kw.w * 4) return false;
        }
        // Above or too far below: skip
        else {
          return false;
        }
        var t = w.text.replace(/[,，]/g, '');
        // Filter out bare "1" — OCR misreads ¥ as "1", producing a standalone "1" word
        // that is NOT an amount. Also filter other single-digit numbers (< 2) which
        // are almost never invoice amounts.
        if (/^\d$/.test(t) && parseInt(t) < 2) return false;
        // Match: bare number, ¥-refixed number, or negative amount
        return /^-?¥?\d+(\.\d{1,2})?$/.test(t);
      });
      if (nearbyNums.length > 0) {
        // When preferBelow is true, we need to check if the closest same-line amount
        // shares a line with other amounts (indicating it's the 不含税+税额 row).
        // In that case, the 含税价 is below that row.
        if (preferBelow) {
          // Separate same-line vs below-line candidates
          var sameLineNums = nearbyNums.filter(function(w) {
            return Math.abs(w.y - kw.y) <= maxLineDist * 0.5;
          });
          var belowLineNums = nearbyNums.filter(function(w) {
            return w.y - kw.y > maxLineDist * 0.5;
          });

          // Check if same-line amounts have multiple amounts on their line
          // (i.e., the 不含税+税额 row has ¥172.68 and ¥5.18 on the same line)
          if (sameLineNums.length > 0) {
            var firstSame = sameLineNums[0];
            // Count how many amounts are on the same line as firstSame
            var sameLinePeers = nearbyNums.filter(function(w) {
              return w !== firstSame && Math.abs(w.y - firstSame.y) <= firstSame.h * 1.5;
            });
            if (sameLinePeers.length > 0 && belowLineNums.length > 0) {
              // Same line has multiple amounts → this is the 不含税+税额 row
              // The 含税价 is in the below-line candidates
              belowLineNums.sort(function(a, b) {
                // Prefer closest below, then closest horizontally
                var da = a.y - kw.y;
                var db = b.y - kw.y;
                if (da !== db) return da - db;
                return Math.abs(a.x - kw.x) - Math.abs(b.x - kw.x);
              });
              var amtStr = cleanOcrAmtStr(belowLineNums[0].text);
              var val = parseFloat(amtStr);
              if (val > 0 && val < 1000000 && !isLikelyYearOrDate(val, belowLineNums[0].text)) {
                return Math.round(val * 100) / 100;
              }
            }
          }
          // If no below-line candidates or single amount on same line, fall through
        }
        // Sort: prefer same-line results, then closest horizontally
        nearbyNums.sort(function(a, b) {
          var aOnLine = Math.abs(a.y - kw.y) <= maxLineDist * 0.5 ? 0 : 1;
          var bOnLine = Math.abs(b.y - kw.y) <= maxLineDist * 0.5 ? 0 : 1;
          if (aOnLine !== bOnLine) return aOnLine - bOnLine;
          return Math.abs(a.x - kw.x) - Math.abs(b.x - kw.x);
        });
        var amtStr = cleanOcrAmtStr(nearbyNums[0].text);
        var val = parseFloat(amtStr);
        // Exclude year-like values and ensure positive reasonable amount
        if (val > 0 && val < 1000000 && !isLikelyYearOrDate(val, nearbyNums[0].text)) return Math.round(val * 100) / 100;
      }
    }
    return 0;
  }

  // Helper: find amount near a specific word object (same proximity logic as findAmountNearKeyword)
  // Used when we need context-aware matching — find the word first, then use it directly
  function findNumberNearWord(kw, maxLineDist) {
    if (!kw || !wordMap) return 0;
    maxLineDist = maxLineDist || 80;
    var nearbyNums = wordMap.filter(function(w) {
      if (w === kw) return false;
      // Skip low-confidence number words
      if (w.confidence !== undefined && w.confidence < 0.3) return false;
      var dy = w.y - kw.y;
      var ady = Math.abs(dy);
      if (ady <= maxLineDist * 0.5) {
        if (w.x < kw.x - 20) return false;
      } else if (dy > 0 && dy <= maxLineDist) {
        if (w.x < kw.x - kw.w * 2 || w.x > kw.x + kw.w * 4) return false;
      } else {
        return false;
      }
      var t = w.text.replace(/[,，]/g, '');
      if (/^\d$/.test(t) && parseInt(t) < 2) return false;
      return /^-?¥?\d+(\.\d{1,2})?$/.test(t);
    });
    if (!nearbyNums.length) return 0;
    nearbyNums.sort(function(a, b) {
      var aOnLine = Math.abs(a.y - kw.y) <= maxLineDist * 0.5 ? 0 : 1;
      var bOnLine = Math.abs(b.y - kw.y) <= maxLineDist * 0.5 ? 0 : 1;
      if (aOnLine !== bOnLine) return aOnLine - bOnLine;
      return Math.abs(a.x - kw.x) - Math.abs(b.x - kw.x);
    });
    var amtStr = cleanOcrAmtStr(nearbyNums[0].text);
    var val = parseFloat(amtStr);
    // Exclude year-like values and ensure positive reasonable amount
    if (val > 0 && val < 1000000 && !isLikelyYearOrDate(val, nearbyNums[0].text)) return Math.round(val * 100) / 100;
    return 0;
  }

  var amountTax = 0, amountNoTax = 0, taxAmount = 0;

  // === Step 0: Ticket-specific coordinate extraction (BEFORE invoice logic) ===
  // Train/ride tickets have a completely different layout from VAT invoices.
  // Ticket amounts are typically at left-half, 35-55% from top.
  // Must run BEFORE Step 0's generic invoice extraction to avoid wrong values
  // from generic region classification (55-75% = amount area doesn't apply to tickets).
  if (isTicket && wordMap && imgW > 0 && imgH > 0) {
    // 0a. Try keyword proximity with 'any' region
    amountTax = findAmountNearKeyword(/票\s*价/, 'any');
    if (!amountTax) amountTax = findAmountNearKeyword(/全\s*价/, 'any');
    if (!amountTax) amountTax = findAmountNearKeyword(/优\s*惠\s*价/, 'any');
    if (!amountTax) amountTax = findAmountNearKeyword(/学\s*生\s*价/, 'any');

    // 0b. Positional fallback: ticket amount is in left-half, ~35-55% from top
    // collectAmountWords already filters out year-like values via isLikelyYearOrDate
    if (!amountTax) {
      var ticketAmts = collectAmountWords(wordMap, imgW, imgH, null, 0, 0.55, 0.3, 0.6);
      // Further filter: ticket prices are typically ¥5-¥5000
      ticketAmts = ticketAmts.filter(function(a) { return a.value >= 5 && a.value <= 5000; });
      if (ticketAmts.length > 0) {
        // Take the largest amount in the ticket amount area
        amountTax = ticketAmts[0].value;
      }
    }

    if (amountTax > 0) {
      amountNoTax = amountTax;
      console.log('[OCR坐标] 车票金额提取:', { amountTax: amountTax });
    }
  }

  // === Step 1: Coordinate-based amount extraction for invoices (highest priority) ===
  // Use word coordinates to find amounts by keyword proximity in the amount region.
  // Only runs for non-ticket invoices.
  if (!isTicket && wordMap && imgW > 0 && imgH > 0) {
    var amountText = getRegionText(wordMap, 'amount');

    // 价税合计 — try full keyword in amount region first
    if (amountText) {
      amountTax = findAmountNearKeyword(/价\s*税\s*合\s*计/, 'amount', 80, true);
      // Also try regex on amount region text
      if (!amountTax) amountTax = findFirstNum('价\\s*税\\s*合\\s*计', amountText);
    }
    // Try partial keyword match: OCR often splits "价税合计" into "价税"+"合计" etc.
    if (!amountTax) {
      // "价税" partial — in VAT invoices, "价税" is part of "价税合计" label,
      // and the nearest number should be the 含税总价
      amountTax = findAmountNearKeyword(/价\s*税/, 'amount');
    }
    // Also try "税合计" partial (OCR split: "价" + "税合计")
    if (!amountTax) {
      amountTax = findAmountNearKeyword(/税\s*合\s*计/, 'amount');
    }
    // Try "合计" — but ONLY if "价" is nearby to the left (part of "价税合计")
    // Standalone "合计" (without "价" to the left) is the 不含税合计 row → amountNoTax
    if (!amountTax) {
      var _hejiWords = wordMap.filter(function(w) {
        if (w.region !== 'amount') return false;
        // Only "合计" words — exclude "税合计" (already handled by 税\s*合\s*计 above)
        return /合\s*计/.test(w.text) && !/税/.test(w.text);
      });
      for (var _hi = 0; _hi < _hejiWords.length && !amountTax; _hi++) {
        var _hw = _hejiWords[_hi];
        // Check if "价" is to the left (part of split "价税合计")
        var _hasJiaLeft = wordMap.some(function(w) {
          if (w === _hw) return false;
          if (!/价/.test(w.text)) return false;
          var dx = _hw.x - (w.x + w.w);
          var dy = Math.abs(_hw.y - w.y);
          return dx >= -30 && dx < 250 && dy < 50;
        });
        if (_hasJiaLeft) {
          // This "合计" is part of "价税合计" → find nearby number for 含税价
          amountTax = findNumberNearWord(_hw);
        }
      }
    }
    // Try "小写" keyword — "（小写）" is right before the 含税价 number
    // (from "价税合计（大写）...（小写）¥123.45" — very specific to 含税价)
    // preferBelow=true: 含税价 is BELOW the 不含税+税额 row, which is on the same line as "小写"
    if (!amountTax) {
      amountTax = findAmountNearKeyword(/小\s*写/, 'amount', 80, true);
    }
    // If not found in amount region, try anywhere with full keyword
    if (!amountTax) {
      amountTax = findAmountNearKeyword(/价\s*税\s*合\s*计/, 'any');
    }

    // 不含税金额 — from standalone "合计" keyword (without "价" to the left)
    // In invoice layout: "合计  100.00" is the 不含税金额合计 row
    // "合计" is closer to the number than "金额" — more reliable match
    if (!amountNoTax) {
      var _hejiWords2 = wordMap.filter(function(w) {
        if (w.region !== 'amount') return false;
        return /合\s*计/.test(w.text) && !/税/.test(w.text);
      });
      for (var _hi2 = 0; _hi2 < _hejiWords2.length && !amountNoTax; _hi2++) {
        var _hw2 = _hejiWords2[_hi2];
        // Standalone "合计" — NO "价" to the left
        var _hasJiaLeft2 = wordMap.some(function(w) {
          if (w === _hw2) return false;
          if (!/价/.test(w.text)) return false;
          var dx = _hw2.x - (w.x + w.w);
          var dy = Math.abs(_hw2.y - w.y);
          return dx >= -30 && dx < 250 && dy < 50;
        });
        if (!_hasJiaLeft2) {
          // This "合计" is standalone → 不含税价 row
          amountNoTax = findNumberNearWord(_hw2);
        }
      }
    }

    // 不含税金额 — from "金额" keyword in amount region (secondary)
    // Coordinate proximity: find "金额" label → nearby number on same or next line
    if (!amountNoTax) {
      amountNoTax = findAmountNearKeyword(/金\s*额/, 'amount');
      // Fallback: regex on amount region text, but skip if preceded by "税" or "合计"
      // (e.g., "税额" → tax amount, "合计金额" → total including tax)
      if (!amountNoTax && amountText) {
        var amtPreMatch = amountText.match(/(?:^|[^税合])金\s*额[\s\S]*?(\d+(?:,\d{3})*\.\d{2})/);
        if (amtPreMatch) amountNoTax = parseAmt(amtPreMatch[1]);
      }
      // Also try "不含税金额" explicit label
      if (!amountNoTax && amountText) {
        amountNoTax = findFirstNum('不\\s*含\\s*税\\s*金\\s*额', amountText);
      }
    }

    // Cross-check: validate amountNoTax against amountTax
    // Invariant: 含税价 >= 不含税价 (always true for any invoice)
    // If violated, "金额" keyword likely matched the wrong number
    if (amountNoTax > 0 && amountTax > 0) {
      if (Math.abs(amountNoTax - amountTax) < 0.01) {
        // Equal values — keyword matched the same number (e.g., "合计金额" near 含税价)
        console.log('[OCR坐标] 不含税价=含税价，可能匹配错误，重置不含税价让推导逻辑重算');
        amountNoTax = 0;
      } else if (amountNoTax > amountTax) {
        // 不含税价 > 含税价 — impossible, "金额" keyword matched the 含税价 number
        console.log('[OCR坐标] 不含税价>含税价，"金额"关键词可能匹配到含税价，重置');
        amountNoTax = 0;
      }
    }

    // 税额 — from "税额" keyword in amount region
    if (!taxAmount) {
      taxAmount = findAmountNearKeyword(/税\s*额/, 'amount');
      if (!taxAmount && amountText) {
        taxAmount = findFirstNum('税\\s*额', amountText);
      }
    }

    // 税率 — from "税率" keyword in amount region (for validation, not stored yet)
    var taxRate = 0;
    if (amountText) {
      var trMatch = amountText.match(/税\s*率[^\d]*?(\d+)\s*%/);
      if (trMatch) taxRate = parseInt(trMatch[1], 10);
    }

    // Cross-validation: derive missing values from found values
    // amountTax = amountNoTax + taxAmount (basic VAT formula)
    if (amountTax > 0 && amountNoTax > 0 && !taxAmount) {
      taxAmount = Math.round((amountTax - amountNoTax) * 100) / 100;
    }
    if (amountTax > 0 && taxAmount > 0 && !amountNoTax && taxAmount < amountTax) {
      amountNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
    }
    if (!amountTax && amountNoTax > 0 && taxAmount > 0) {
      amountTax = Math.round((amountNoTax + taxAmount) * 100) / 100;
    }

    // === Positional fallback: largest amount in amount region ===
    // If keyword-based extraction didn't find amountTax, use positional heuristics.
    // In the amount region of a VAT invoice, the largest ¥ amount is almost always
    // the 价税合计 (tax-inclusive total). This is robust even when keywords are
    // split by OCR into multiple words.
    if (!amountTax && amountText) {
      var regionAmounts = collectAmountWords(wordMap, imgW, imgH, 'amount');
      if (regionAmounts.length > 0) {
        // The largest amount in the amount region is most likely the 价税合计
        var largestAmt = regionAmounts[0].value;
        // Sanity check: if we already have amountNoTax, amountTax must be >= amountNoTax
        if (amountNoTax > 0 && largestAmt < amountNoTax) {
          // The largest found is actually the amountNoTax — no separate amountTax found
          // Don't use positional fallback, let regex steps handle it
        } else {
          amountTax = largestAmt;
          // If we don't have amountNoTax yet, the second-largest might be it
          // (but only if there are multiple distinct amounts)
          if (!amountNoTax && regionAmounts.length >= 2) {
            // Find amounts that are significantly smaller than amountTax
            var smallerAmts = regionAmounts.filter(function(a) {
              return a.value < amountTax * 0.95; // at least 5% smaller
            });
            if (smallerAmts.length > 0) {
              // The largest of the smaller amounts is likely amountNoTax
              // (not taxAmount which is usually much smaller)
              var noTaxCandidate = smallerAmts[0].value;
              if (noTaxCandidate > (amountTax * 0.5)) {
                // More than 50% of amountTax → this is amountNoTax, not taxAmount
                amountNoTax = noTaxCandidate;
              }
            }
          }
        }
      }
    }

    // Derive missing values after positional fallback
    if (amountTax > 0 && amountNoTax > 0 && !taxAmount) {
      taxAmount = Math.round((amountTax - amountNoTax) * 100) / 100;
    }
    if (amountTax > 0 && taxAmount > 0 && !amountNoTax && taxAmount < amountTax) {
      amountNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
    }

    // 合计金额 / 金额合计 (for non-VAT invoices)
    if (!amountTax && !amountNoTax) {
      var amtRegionVal = findFirstNum('合\\s*计\\s*金\\s*额', amountText || '');
      if (amtRegionVal > 0) { amountTax = amtRegionVal; amountNoTax = amtRegionVal; }
    }
    if (!amountTax && !amountNoTax) {
      var amtRegionVal2 = findFirstNum('金\\s*额\\s*合\\s*计', amountText || '');
      if (amtRegionVal2 > 0) { amountTax = amtRegionVal2; amountNoTax = amtRegionVal2; }
    }

    // 实付/应付金额
    if (!amountTax && !amountNoTax) {
      var payVal = findFirstNum('实\\s*付\\s*金\\s*额', amountText || '') || findFirstNum('应\\s*付\\s*金\\s*额', amountText || '');
      if (payVal > 0) { amountTax = payVal; amountNoTax = payVal; }
    }

    if (amountTax > 0 || amountNoTax > 0 || taxAmount > 0) {
      console.log('[OCR坐标] 金额区域提取:', { amountTax: amountTax, amountNoTax: amountNoTax, taxAmount: taxAmount, taxRate: taxRate || '-' });
    }

    // Unconditional invariant: 含税价 must be >= 不含税价
    // This catches cases where keyword proximity matched the wrong number
    // (e.g., "价税" matched a number near the 金额 row instead of 价税合计)
    if (amountTax > 0 && amountNoTax > 0 && amountTax < amountNoTax) {
      var _tmpAmt = amountTax;
      amountTax = amountNoTax;
      amountNoTax = _tmpAmt;
      console.log('[OCR坐标] 含税价<不含税价，已交换:', { amountTax: amountTax, amountNoTax: amountNoTax });
    }
  }

  // === Step 1: 价税合计 → 含税总价 (fallback to full-text regex) ===
  // 含税价在发票最下方，OCR 文本中后出现的数字更可能是含税价。
  // 使用 findLastNum 找最后一个匹配，避免匹配到不含税价行的数字。
  if (!amountTax) {
    amountTax = findLastNum('价\\s*税\\s*合\\s*计\\s*[（(]\\s*大\\s*写\\s*[）)][^\\d]*?[（(]\\s*小\\s*写\\s*[）)]', fullText);
    if (!amountTax) amountTax = findLastNum('价\\s*税\\s*合\\s*计\\s*[（(]\\s*小\\s*写\\s*[）)]', fullText);
    if (!amountTax) amountTax = findLastNum('价\\s*税\\s*合\\s*计', fullText);
    // Variant: 价税合计 without explicit 小写/大写, just ¥ directly
    // [\\s\\S] for cross-newline matching (PP-OCRv5 line splits)
    if (!amountTax) {
      var pthMatches = [];
      var pthRe = /价\s*税\s*合\s*计[\s\S]*?¥\s*(\d+(?:,\d{3})*\.\d{2})/g;
      var pthM;
      while ((pthM = pthRe.exec(fullText)) !== null) {
        var pthV = parseAmt(pthM[1]);
        if (pthV > 0 && !isLikelyYearOrDate(pthV, pthM[1])) pthMatches.push(pthV);
      }
      // Take the last match (含税价 is at the bottom)
      if (pthMatches.length > 0) amountTax = pthMatches[pthMatches.length - 1];
    }
  }

  // === Step 1.5: 电子发票/打车发票/网约车发票 ===
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('实\\s*付\\s*金\\s*额', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('应\\s*付\\s*金\\s*额', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('合\\s*计\\s*金\\s*额', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('金\\s*额\\s*合\\s*计', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  if (!amountTax && !amountNoTax) {
    var amtLineMatch = fullText.match(/(?:合\s*计|总\s*计|计\s*费)[\s\S]*?金\s*额[\s\S]*?(\d+(?:,\d{3})*\.\d{2})/);
    if (amtLineMatch) {
      var val3 = parseAmt(amtLineMatch[1]);
      if (val3 > 0) { amountTax = val3; amountNoTax = val3; }
    }
  }
  // 电子发票: "税价合计" (OCR sometimes swaps characters)
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('税\\s*价\\s*合\\s*计', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  // 电子发票: amount after "¥" near "合计" in any order
  // 含税价在下方 → 取最后一个匹配
  // [\\s\\S] for cross-newline matching (PP-OCRv5 may split across lines)
  if (!amountTax && !amountNoTax) {
    var amtNearTotalMatches = [];
    var amtNearTotalRe = /合\s*计[\s\S]{0,30}?¥\s*(\d+(?:,\d{3})*\.\d{2})/g;
    var amtNearTotalM;
    while ((amtNearTotalM = amtNearTotalRe.exec(fullText)) !== null) {
      var amtNearTotalV = parseAmt(amtNearTotalM[1]);
      if (amtNearTotalV > 0 && amtNearTotalV < 100000 && !isLikelyYearOrDate(amtNearTotalV, amtNearTotalM[1])) amtNearTotalMatches.push(amtNearTotalV);
    }
    var amtNearTotalRe2 = /¥\s*(\d+(?:,\d{3})*\.\d{2})[\s\S]{0,30}?合\s*计/g;
    while ((amtNearTotalM = amtNearTotalRe2.exec(fullText)) !== null) {
      var amtNearTotalV2 = parseAmt(amtNearTotalM[1]);
      if (amtNearTotalV2 > 0 && amtNearTotalV2 < 100000 && !isLikelyYearOrDate(amtNearTotalV2, amtNearTotalM[1])) amtNearTotalMatches.push(amtNearTotalV2);
    }
    if (amtNearTotalMatches.length > 0) {
      var amtVal = amtNearTotalMatches[amtNearTotalMatches.length - 1];
      amountTax = amtVal; amountNoTax = amtVal;
    }
  }

  // === Step 2: Remove ALL 价税合计 variants ===
  // [\\s\\S] instead of . to match across newlines (PP-OCRv5 may split across lines)
  var workText = fullText.replace(/价\s*税\s*合\s*计\s*(?:[（(](?:\s*大\s*写\s*[\s\S]*?)?\s*小\s*写\s*[）)]\s*)?[\s\S]*?\d+(?:,\d{3})*\.\d{2}/g, '');

  // === Step 3: 合计 → 不含税价 (after 价税合计 removed from text) ===
  // In workText, remaining "合计" should refer to the 不含税 合计 row
  if (!amountNoTax) {
    var hejiNum = findFirstNum('合\\s*计', workText);
    // But "合计" could still match 含税价 — validate against amountTax
    if (hejiNum > 0 && amountTax > 0 && Math.abs(hejiNum - amountTax) < 0.01) {
      hejiNum = 0; // Same as 含税价, probably wrong match
    }
    if (hejiNum > 0) amountNoTax = hejiNum;
  }

  // === Step 4: 金额 → 不含税价 ===
  // Exclude "合计金额"/"价税金额" patterns — those refer to 含税价, not 不含税价
  // [\\s\\S]*? for cross-newline matching (PP-OCRv5 may split keyword and value)
  if (!amountNoTax) {
    var amtNumMatch = workText.match(/(?:^|[^税合])金\s*额[\s\S]*?(\d+(?:,\d{3})*\.\d{2})/);
    if (amtNumMatch) amountNoTax = parseAmt(amtNumMatch[1]);
  }

  // Validate: if amountNoTax equals amountTax, keyword likely matched the wrong number
  if (amountNoTax > 0 && amountTax > 0 && Math.abs(amountNoTax - amountTax) < 0.01) {
    amountNoTax = 0; // Reset, let tax-based derivation handle it
  }

  // === Step 5: 税额反推 ===
  if (amountTax > 0 && !amountNoTax) {
    if (!taxAmount) {
      taxAmount = findFirstNum('税\\s*额', fullText);
    }
    if (taxAmount > 0 && taxAmount < amountTax) amountNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
  }

  // === Step 6: 火车票 — robust detection ===
  // All ticket price matches validate against isLikelyYearOrDate to avoid
  // matching "2025.01" (from dates like "2025年01月") as a price.
  // PP-OCRv5 may split keywords and values across lines — use [\s\S]*? for cross-line matching.
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('票\s*价', fullText);
    if (amountTax > 0 && amountTax <= 5000) amountNoTax = amountTax;
    else amountTax = 0;
  }
  if (!amountTax && !amountNoTax) {
    var ticketMatch = fullText.match(/票[\s\S]*?¥\s*(\d+(?:,\d{3})*\.\d{2})/);
    if (ticketMatch) {
      var val = parseAmt(ticketMatch[1]);
      if (val > 0 && val <= 5000 && !isLikelyYearOrDate(val, ticketMatch[1])) { amountTax = val; amountNoTax = val; }
    }
  }
  // 车票: "票价" keyword with amount pattern (no ¥ symbol)
  if (!amountTax && !amountNoTax) {
    var priceLineMatch = fullText.match(/票\s*价[\s\S]*?(\d+\.\d{2})/);
    if (priceLineMatch) {
      var pval = parseAmt(priceLineMatch[1]);
      if (pval > 0 && pval <= 5000 && !isLikelyYearOrDate(pval, priceLineMatch[1])) { amountTax = pval; amountNoTax = pval; }
    }
  }
  // 车票: amount near train keywords (cross-line safe)
  if (!amountTax && !amountNoTax) {
    var yMatch = fullText.match(/¥\s*(\d+(?:,\d{3})*\.\d{2})[\s\S]*?(?:车票|车次|座位|检票|进站|出站|乘车|站台)/);
    if (!yMatch) yMatch = fullText.match(/(?:车票|车次|座位|检票|进站|出站|乘车|站台)[\s\S]*?¥\s*(\d+(?:,\d{3})*\.\d{2})/);
    if (yMatch) {
      var val2 = parseAmt(yMatch[1]);
      if (val2 > 0 && val2 <= 5000 && !isLikelyYearOrDate(val2, yMatch[1])) { amountTax = val2; amountNoTax = val2; }
    }
  }
  if (!amountTax && !amountNoTax) {
    var seatMatch = fullText.match(/(?:席\s*别|座\s*位)[\s\S]*?(\d+\.\d{2})/);
    if (seatMatch) {
      var val3 = parseAmt(seatMatch[1]);
      if (val3 > 0 && val3 <= 5000 && !isLikelyYearOrDate(val3, seatMatch[1])) { amountTax = val3; amountNoTax = val3; }
    }
  }
  // 车票: "全价" or "优惠价" pattern
  if (!amountTax && !amountNoTax) {
    var discountMatch = fullText.match(/(?:全\s*价|优\s*惠\s*价|学\s*生\s*价)[\s\S]*?(\d+\.\d{2})/);
    if (discountMatch) {
      var dval = parseAmt(discountMatch[1]);
      if (dval > 0 && dval <= 5000 && !isLikelyYearOrDate(dval, discountMatch[1])) { amountTax = dval; amountNoTax = dval; }
    }
  }
  // 车票: "￥" (full-width yen sign) pattern
  if (!amountTax && !amountNoTax) {
    var fwyMatch = fullText.match(/￥\s*(\d+(?:,\d{3})*\.\d{2})/);
    if (fwyMatch) {
      var fyVal = parseAmt(fwyMatch[1]);
      if (fyVal > 0 && fyVal <= 5000 && !isLikelyYearOrDate(fyVal, fwyMatch[1])) { amountTax = fyVal; amountNoTax = fyVal; }
    }
  }
  if (!amountTax && !amountNoTax) {
    var trainKwRe = /(?:车\s*次|车\s*站|号\s*车|检\s*票|铺\s*位|卧\s*铺|二\s*等|一\s*等|动\s*车|高\s*铁|特\s*等|进\s*站|出\s*站|乘\s*车|站\s*台)/;
    if (trainKwRe.test(fullText)) {
      var allYen = [];
      var yenRe = /¥\s*(\d+(?:,\d{3})*\.\d{2})/g;
      var ym;
      while ((ym = yenRe.exec(fullText)) !== null) {
        var yv = parseAmt(ym[1]);
        if (yv > 0 && yv <= 5000 && !isLikelyYearOrDate(yv, ym[1])) allYen.push(yv);
      }
      // Also check ￥ (full-width)
      var yenRe2 = /￥\s*(\d+(?:,\d{3})*\.\d{2})/g;
      while ((ym = yenRe2.exec(fullText)) !== null) {
        var yv2 = parseAmt(ym[1]);
        if (yv2 > 0 && yv2 <= 5000 && !isLikelyYearOrDate(yv2, ym[1])) allYen.push(yv2);
      }
      // Also look for amounts without ¥ symbol
      if (allYen.length === 0) {
        var bareAmtRe = /(?:票\s*价|全\s*价|金\s*额)[^\d]*?(\d+\.\d{2})/gi;
        var bm;
        while ((bm = bareAmtRe.exec(fullText)) !== null) {
          var bv = parseAmt(bm[1]);
          if (bv > 0 && bv <= 5000 && !isLikelyYearOrDate(bv, bm[1])) allYen.push(bv);
        }
      }
      if (allYen.length === 1) {
        amountTax = allYen[0]; amountNoTax = allYen[0];
      } else if (allYen.length > 1) {
        var typicalYen = allYen.filter(function(y) { return y >= 5 && y <= 5000; });
        if (typicalYen.length === 1) { amountTax = typicalYen[0]; amountNoTax = typicalYen[0]; }
        else if (typicalYen.length > 1) { amountTax = Math.max.apply(Math, typicalYen); amountNoTax = amountTax; }
      }
    }
  }

  // === Step 6.6: 出租车票 / 乘车票 — specific patterns ===
  if (!amountTax && !amountNoTax) {
    var taxiMatch = fullText.match(/(?:乘\s*车|出\s*租|打\s*车|网\s*约\s*车)[\s\S]*?(?:金\s*额|费\s*用|价\s*格)[\s\S]*?(\d+\.\d{2})/i);
    if (taxiMatch) {
      var tval = parseAmt(taxiMatch[1]);
      if (tval > 0 && tval <= 5000 && !isLikelyYearOrDate(tval, taxiMatch[1])) { amountTax = tval; amountNoTax = tval; }
    }
  }

  // === Step 6.7: 定额发票 — amount right after "¥" in short texts ===
  if (!amountTax && !amountNoTax) {
    // 定额发票 usually very short, contains just "金额 ¥X.00"
    if (fullText.length < 500) {
      var dingEMatch = fullText.match(/金\s*额[\s\S]*?¥?\s*(\d+\.\d{2})/);
      if (dingEMatch) {
        var deVal = parseAmt(dingEMatch[1]);
        if (deVal > 0 && deVal < 100000) { amountTax = deVal; amountNoTax = deVal; }
      }
    }
  }

  // === Step 6.5: Super fallback — any ¥ amount ===
  if (!amountTax && !amountNoTax) {
    var yenRe2a = /¥\s*(\d+(?:,\d{3})*\.\d{2})/g;
    var amounts = [], ym2;
    while ((ym2 = yenRe2a.exec(fullText)) !== null) {
      var yv2a = parseAmt(ym2[1]);
      if (yv2a > 0 && yv2a < 50000 && !isLikelyYearOrDate(yv2a, ym2[1])) amounts.push(yv2a);
    }
    // Also check ￥ (full-width yen sign)
    var yenRe2b = /￥\s*(\d+(?:,\d{3})*\.\d{2})/g;
    while ((ym2 = yenRe2b.exec(fullText)) !== null) {
      var yv2b = parseAmt(ym2[1]);
      if (yv2b > 0 && yv2b < 50000 && !isLikelyYearOrDate(yv2b, ym2[1])) amounts.push(yv2b);
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
  // Fallback: try "开票金额"/"开票金额(含税)"/"发票金额"
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('开\\s*票\\s*金\\s*额', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }
  if (!amountTax && !amountNoTax) {
    amountTax = findFirstNum('发\\s*票\\s*金\\s*额', fullText);
    if (amountTax > 0) amountNoTax = amountTax;
  }

  // AUTO-FALLBACK: if only amountNoTax found, derive amountTax from taxAmount if possible
  // DO NOT blindly set amountTax = amountNoTax (that's wrong for VAT invoices where they differ)
  if (amountNoTax > 0 && amountTax === 0) {
    if (taxAmount > 0 && taxAmount < amountNoTax) {
      amountTax = Math.round((amountNoTax + taxAmount) * 100) / 100;
    } else if (wordMap && imgW > 0 && imgH > 0) {
      // Last-ditch positional: find any amount larger than amountNoTax in the image
      // This catches cases where 价税合计 is outside the standard amount region
      var allImgAmts = collectAmountWords(wordMap, imgW, imgH, null);
      var largerAmts = allImgAmts.filter(function(a) { return a.value > amountNoTax + 0.01; });
      if (largerAmts.length > 0) {
        // The smallest amount that's larger than amountNoTax is most likely amountTax
        largerAmts.sort(function(a, b) { return a.value - b.value; });
        amountTax = largerAmts[0].value;
        taxAmount = Math.round((amountTax - amountNoTax) * 100) / 100;
        console.log('[OCR提取] 全图位置反推含税价:', { amountTax: amountTax, amountNoTax: amountNoTax, taxAmount: taxAmount });
      } else {
        // No larger amount found — this is likely a non-VAT invoice
        amountTax = amountNoTax;
      }
    } else {
      // No coordinates — this is likely a non-VAT invoice, amountTax = amountNoTax is correct
      amountTax = amountNoTax;
    }
  }
  // If only amountTax found, derive amountNoTax from taxAmount
  if (amountTax > 0 && amountNoTax === 0 && taxAmount > 0 && taxAmount < amountTax) {
    amountNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
  }

  // Unconditional sanity check: 含税价 must be >= 不含税价
  // This catches all cases where keyword proximity or regex matched the wrong value,
  // regardless of whether taxAmount was found.
  if (amountTax > 0 && amountNoTax > 0 && amountTax < amountNoTax) {
    var _finalTmp = amountTax;
    amountTax = amountNoTax;
    amountNoTax = _finalTmp;
    console.log('[OCR提取] 最终校验：含税价<不含税价，已交换');
  }

  // === Final cross-validation: verify VAT formula amountTax ≈ amountNoTax + taxAmount ===
  if (amountTax > 0 && amountNoTax > 0 && taxAmount > 0) {
    var expectedTax = Math.round((amountTax - amountNoTax) * 100) / 100;
    var diff = Math.abs(expectedTax - taxAmount);
    // If the difference is more than 0.02 (rounding tolerance), values are inconsistent
    if (diff > 0.02) {
      console.warn('[OCR提取] 金额交叉验证失败: amountTax=' + amountTax + ' ≠ amountNoTax(' + amountNoTax + ') + taxAmount(' + taxAmount + ')=' + (amountNoTax + taxAmount));
      // Try to fix: if amountTax < amountNoTax, they might be swapped
      if (amountTax < amountNoTax) {
        var tmp = amountTax;
        amountTax = amountNoTax;
        amountNoTax = tmp;
        console.log('[OCR提取] 已交换含税/不含税金额');
      }
      // Re-check after swap
      expectedTax = Math.round((amountTax - amountNoTax) * 100) / 100;
      diff = Math.abs(expectedTax - taxAmount);
      if (diff > 0.02) {
        // Still inconsistent — trust amountTax (from 价税合计, most reliable) and taxAmount,
        // recompute amountNoTax
        var recomputedNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
        if (recomputedNoTax > 0) {
          console.log('[OCR提取] 重算不含税价: ' + amountNoTax + ' → ' + recomputedNoTax);
          amountNoTax = recomputedNoTax;
        }
      }
    }
  }

  console.log('[OCR提取] 金额:', { amountTax: amountTax, amountNoTax: amountNoTax, taxAmount: taxAmount }, '销售方:', sellerName || '(未识别)', '信用代码:', sellerCreditCode || '(未识别)', '车票:', isTicket);
  if (!amountTax && !amountNoTax && !sellerName && !isTicket) {
    console.warn('[OCR提取] 未能识别任何信息，OCR完整文本:', fullText);
  }

  return { amountTax: amountTax, amountNoTax: amountNoTax, taxAmount: taxAmount, sellerName: sellerName, sellerCreditCode: sellerCreditCode, _ocrText: fullText, isTicket: isTicket };
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
    var ocrResultStr = await invoke('ocr_image', { dataUrl: dataUrl });
    if (!ocrResultStr) return;
    // Parse the JSON result from Rust (new format with coordinates)
    var ocrResult;
    try {
      ocrResult = JSON.parse(ocrResultStr);
    } catch(e) {
      // Fallback: treat as plain text (shouldn't happen, but safe)
      ocrResult = { text: ocrResultStr, lines: [], imgW: 0, imgH: 0 };
    }

    // --- v1.7.0: Try coordinate-first extraction (PP-OCRv5 bbox) ---
    var info = null;
    if (ocrResult.lines && ocrResult.imgW > 0 && ocrResult.imgH > 0) {
      info = extractByCoordinates(ocrResult);
    }
    // Fallback to legacy regex-based extraction if:
    // 1) No coordinate result at all, OR
    // 2) Coordinate path didn't find amounts (seller-only result is not enough —
    //    regex path may find amounts via different patterns)
    // But: if coordinate path found amounts, don't overwrite with regex
    // (coordinate-based amounts are more reliable)
    // --- [DISABLED] Legacy regex fallback — PP-OCRv5 is accurate enough ---
    // Uncomment the block below if coordinate path needs regex supplementation.
    /*
    var coordHasAmounts = info && (info.amountTax > 0 || info.amountNoTax > 0);
    if (!info || !coordHasAmounts) {
      var regexInfo = extractInvoiceInfo(ocrResult);
      if (!info) {
        info = regexInfo;
      } else {
        // Merge: regex can fill in amounts that coordinates missed
        if (!info.amountTax && regexInfo.amountTax) info.amountTax = regexInfo.amountTax;
        if (!info.amountNoTax && regexInfo.amountNoTax) info.amountNoTax = regexInfo.amountNoTax;
        if (!info.taxAmount && regexInfo.taxAmount) info.taxAmount = regexInfo.taxAmount;
        // Also fill seller if coordinate path missed it
        if (!info.sellerName && regexInfo.sellerName) info.sellerName = regexInfo.sellerName;
        if (!info.sellerCreditCode && regexInfo.sellerCreditCode) info.sellerCreditCode = regexInfo.sellerCreditCode;
      }
    }
    */

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
  } catch(e) {
    console.warn('[OCR] 识别失败:', e);
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
 * Output: { amountTax, amountNoTax, taxAmount, sellerName, sellerCreditCode, _ocrText, isTicket }
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

  // Detect invoice type
  var invType = _detectInvoiceType(words, imgW, imgH);
  var isTicket = invType === 'ticket';
  var sellerName = '', sellerCreditCode = '';
  var amountTax = 0, amountNoTax = 0, taxAmount = 0;

  console.log('[坐标提取] 发票类型:', invType, '字数:', fullText.length, '词数:', words.length);

  // === Ticket extraction ===
  if (isTicket) {
    sellerName = getTicketTypeLabel(fullText);

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
             sellerName: sellerName, sellerCreditCode: '', _ocrText: fullText, isTicket: true };
  }

  // === VAT / Ride invoice extraction ===

  // --- Seller info ---
  var sellerInfo = _extractSeller(words, imgW, imgH);
  sellerName = sellerInfo.sellerName;
  sellerCreditCode = sellerInfo.sellerCreditCode;

  // --- Amount extraction ---

  // Step 1: 价税合计（含税总价）— most reliable
  // Location: ny ≈ 0.20~0.30 (near bottom of invoice)
  // Keywords: "价税合计", "（小写）", or just ¥ at that position
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

  // Step 2: 不含税合计 — "合计" row
  // Location: ny ≈ 0.45~0.55, just above the 价税合计 row
  // Must distinguish from "价税合计" — standalone "合计" without "价" to its left
  if (!amountNoTax) {
    var hejiLabels = _findWords(words, /合\s*计/);
    // Filter: standalone "合计" (no "价" or "税" nearby to the left)
    var standaloneHeji = hejiLabels.filter(function(hw) {
      // Exclude if "税" is in this word itself (e.g., "税合计")
      if (/税/.test(hw.normText)) return false;
      // Check if "价" is nearby to the left
      var hasJiaLeft = words.some(function(w) {
        if (w === hw) return false;
        if (!/价/.test(w.normText)) return false;
        var dx = hw.cx - w.cx;
        var dy = Math.abs(w.cy - hw.cy);
        return dx >= -20 && dx < 300 && dy < 50;
      });
      return !hasJiaLeft;
    });

    for (var hi = 0; hi < standaloneHeji.length && !amountNoTax; hi++) {
      // Use the "合计" that's ABOVE the 价税合计 (lower ny = higher on page, but wait —
      // in our coords ny=0 is top, so 合计 should have SMALLER ny than 价税合计)
      var amt5 = _findNearbyAmount(words, standaloneHeji[hi], { maxDx: 500, maxDy: 30, maxDyBelow: 80 });
      if (amt5) {
        // Validate: amountNoTax should be < amountTax (if amountTax found)
        if (amountTax > 0 && amt5.value > amountTax) continue;
        if (amountTax > 0 && Math.abs(amt5.value - amountTax) < 0.01) continue; // same = wrong match
        amountNoTax = amt5.value;
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

  // Step 3: 税额 — "税额" keyword in amount region
  var seLabels = _findWords(words, /税\s*额/, undefined, undefined, 0.40, 0.75);
  if (seLabels.length > 0) {
    // Use the bottommost "税额" (in the 合计 row)
    seLabels.sort(function(a, b) { return b.ny - a.ny; });
    var amt7 = _findNearbyAmount(words, seLabels[0], { maxDx: 300, maxDy: 30, maxDyBelow: 60 });
    if (amt7) taxAmount = amt7.value;
  }

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

  // --- Credit code fallback (from full text if coordinates missed) ---
  if (!sellerCreditCode) {
    var ccRe = /(?:纳税人识别号|统一社会信用代码)[\/:：\s]*([A-Z0-9]{15,20})/gi;
    var ccM, lastCc = '';
    while ((ccM = ccRe.exec(normText)) !== null) {
      lastCc = ccM[1];
    }
    if (lastCc) sellerCreditCode = lastCc.toUpperCase();
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
    sellerName: sellerName || '(空)', sellerCreditCode: sellerCreditCode || '(空)' });

  return { amountTax: amountTax, amountNoTax: amountNoTax, taxAmount: taxAmount,
           sellerName: sellerName, sellerCreditCode: sellerCreditCode,
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
