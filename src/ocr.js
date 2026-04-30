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
 * Detect if text is a train/ride ticket (no seller info needed)
 */
function isTicketText(text) {
  var t = text.substring(0, 500);
  return /(?:车\s*次|票\s*价|座\s*位|席\s*别|检\s*票|站\s*台|进\s*站|出\s*站|铁\s*路|乘\s*车|二\s*等|一\s*等|动\s*车|高\s*铁|硬\s*座|软\s*座|卧\s*铺|铺\s*位|出\s*租|打\s*车|网\s*约|滴\s*滴)/.test(t);
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
 * Each entry: { text, x, y, w, h, region, lineIdx, wordIdx }
 */
function buildWordMap(ocrLines, imgW, imgH) {
  if (!ocrLines || !ocrLines.length) return [];
  var map = [];
  for (var li = 0; li < ocrLines.length; li++) {
    var line = ocrLines[li];
    if (!line.words || !line.words.length) continue;
    for (var wi = 0; wi < line.words.length; wi++) {
      var word = line.words[wi];
      map.push({
        text: word.text,
        x: word.x,
        y: word.y,
        w: word.w,
        h: word.h,
        region: classifyRegion(word.x, word.y, word.w, word.h, imgW, imgH),
        lineIdx: li,
        wordIdx: wi
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

  // === Normalize OCR ¥→1 misread ===
  // OCR often misreads "¥" as "1" (they look very similar in many fonts).
  // Pattern: "1XXX.XX" right after amount keywords → should be "¥XXX.XX"
  // e.g., "价税合计1317.00" → "价税合计¥317.00"
  // Only apply after amount-related keywords to avoid corrupting legitimate numbers
  fullText = fullText.replace(/(价\s*税\s*合\s*计|金\s*额|税\s*额|合\s*计|票\s*价|总\s*计|不\s*含\s*税|含\s*税|实\s*付|应\s*付|开\s*票\s*金\s*额|发\s*票\s*金\s*额|全\s*价|优\s*惠\s*价)([^\d]*?)1(\d{2,}\.\d{2})/g, '$1$2¥$3');

  // Helper: find first number with exactly 2 decimal places after a keyword
  function findFirstNum(keyword, text) {
    var re = new RegExp(keyword + '[^\\d]*?(\\d+(?:,\\d{3})*\\.\\d{2})');
    var m = text.match(re);
    return m ? parseAmt(m[1]) : 0;
  }

  /**
   * Coordinate-aware amount extraction: find a number near a keyword
   * by checking word proximity — same line (to the right) OR next line below
   * (for table layouts where keyword is a column header and value is below).
   * Returns the amount or 0.
   */
  /**
   * Clean an OCR amount string: strip ¥/￥ prefix, handle "1" misread of "¥".
   * OCR often misreads "¥317.00" as "1317.00" (¥→1). We detect this by checking
   * if a leading "1" could be a misread ¥ symbol: the number after removing "1"
   * must have exactly 2 decimal places and be a reasonable amount.
   * Returns the cleaned numeric string.
   */
  function cleanOcrAmtStr(raw) {
    var s = raw.replace(/^[¥￥]/, '').replace(/[,，]/g, '');
    // ¥→1 misread detection:
    // If number starts with "1" followed by 2+ digits and .XX (standard amount format),
    // the "1" is likely a misread "¥" symbol (they look very similar in OCR).
    // e.g., "1317.00" → "317.00", "1299.06" → "299.06"
    // But NOT "100.00" (removing "1" gives "00.00" = 0 → rejected)
    // And NOT "12.50" (only 2 digits before decimal, likely legitimate)
    if (/^1\d{2,}\.\d{2}$/.test(s)) {
      var stripped = s.substring(1);
      var strippedVal = parseFloat(stripped);
      if (strippedVal > 0) {
        s = stripped;
      }
    }
    return s;
  }

  function findAmountNearKeyword(keywordRegex, regionFilter, maxLineDist) {
    if (!wordMap || !imgW || !imgH) return 0;
    maxLineDist = maxLineDist || 80; // max vertical distance for "next line" (increased from 30)
    var candidates = wordMap.filter(function(w) {
      if (regionFilter && w.region !== regionFilter && regionFilter !== 'any') return false;
      return keywordRegex.test(w.text);
    });
    if (!candidates.length) return 0;
    // For each keyword match, find the nearest number word
    for (var ci = 0; ci < candidates.length; ci++) {
      var kw = candidates[ci];
      // Find number words on the same line OR directly below
      var nearbyNums = wordMap.filter(function(w) {
        if (w === kw) return false;
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
        return /^¥?\d+(\.\d{1,2})?$/.test(t) || /^￥\d+(\.\d{1,2})?$/.test(t);
      });
      if (nearbyNums.length > 0) {
        // Sort: prefer same-line results, then closest horizontally
        nearbyNums.sort(function(a, b) {
          var aOnLine = Math.abs(a.y - kw.y) <= maxLineDist * 0.5 ? 0 : 1;
          var bOnLine = Math.abs(b.y - kw.y) <= maxLineDist * 0.5 ? 0 : 1;
          if (aOnLine !== bOnLine) return aOnLine - bOnLine;
          return Math.abs(a.x - kw.x) - Math.abs(b.x - kw.x);
        });
        var amtStr = cleanOcrAmtStr(nearbyNums[0].text);
        var val = parseFloat(amtStr);
        if (val > 0 && val < 1000000) return Math.round(val * 100) / 100;
      }
    }
    return 0;
  }

  var amountTax = 0, amountNoTax = 0, taxAmount = 0;

  // === Step 0: Coordinate-based amount extraction (highest priority) ===
  // Use word coordinates to find amounts by keyword proximity in the amount region
  if (wordMap && imgW > 0 && imgH > 0) {
    var amountText = getRegionText(wordMap, 'amount');

    // 价税合计 — try amount region first
    if (amountText) {
      amountTax = findAmountNearKeyword(/价\s*税\s*合\s*计/, 'amount');
      // Also try regex on amount region text
      if (!amountTax) amountTax = findFirstNum('价\\s*税\\s*合\\s*计', amountText);
    }
    // If not found in amount region, try anywhere
    if (!amountTax) {
      amountTax = findAmountNearKeyword(/价\s*税\s*合\s*计/, 'any');
    }

    // 不含税金额 — from "金额" keyword in amount region
    // Coordinate proximity: find "金额" label → nearby number on same or next line
    if (!amountNoTax) {
      amountNoTax = findAmountNearKeyword(/金\s*额/, 'amount');
      // Fallback: regex on amount region text, but skip if preceded by "税"
      if (!amountNoTax && amountText) {
        var amtPreMatch = amountText.match(/(?:^|[^税])金\s*额[^\d]*?(\d+(?:,\d{3})*\.\d{2})/);
        if (amtPreMatch) amountNoTax = parseAmt(amtPreMatch[1]);
      }
      // Also try "不含税金额" explicit label
      if (!amountNoTax && amountText) {
        amountNoTax = findFirstNum('不\\s*含\\s*税\\s*金\\s*额', amountText);
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
  }

  // === Step 1: 价税合计 → 含税总价 (fallback to full-text regex) ===
  if (!amountTax) {
    amountTax = findFirstNum('价\\s*税\\s*合\\s*计\\s*[（(]\\s*大\\s*写\\s*[）)][^\\d]*?[（(]\\s*小\\s*写\\s*[）)]', fullText);
    if (!amountTax) amountTax = findFirstNum('价\\s*税\\s*合\\s*计\\s*[（(]\\s*小\\s*写\\s*[）)]', fullText);
    if (!amountTax) amountTax = findFirstNum('价\\s*税\\s*合\\s*计', fullText);
    // Variant: 价税合计 without explicit 小写/大写, just ¥ directly
    if (!amountTax) {
      var pthMatch = fullText.match(/价\s*税\s*合\s*计[^\d]*?¥\s*(\d+(?:,\d{3})*\.\d{2})/);
      if (pthMatch) amountTax = parseAmt(pthMatch[1]);
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
    var amtLineMatch = fullText.match(/(?:合\s*计|总\s*计|计\s*费)[^\n]*?金\s*额[^\d]*?(\d+(?:,\d{3})*\.\d{2})/);
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
  if (!amountTax && !amountNoTax) {
    var amtNearTotal = fullText.match(/合\s*计[^\n]{0,30}?¥\s*(\d+(?:,\d{3})*\.\d{2})/);
    if (!amtNearTotal) amtNearTotal = fullText.match(/¥\s*(\d+(?:,\d{3})*\.\d{2})[^\n]{0,30}?合\s*计/);
    if (amtNearTotal) {
      var amtVal = parseAmt(amtNearTotal[1]);
      if (amtVal > 0 && amtVal < 100000) { amountTax = amtVal; amountNoTax = amtVal; }
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
    if (!taxAmount) {
      taxAmount = findFirstNum('税\\s*额', fullText);
    }
    if (taxAmount > 0 && taxAmount < amountTax) amountNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
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
  // 车票: "票价" keyword with amount pattern (no ¥ symbol)
  if (!amountTax && !amountNoTax) {
    var priceLineMatch = fullText.match(/票\s*价[^\d]*?(\d+\.\d{2})/);
    if (priceLineMatch) {
      var pval = parseAmt(priceLineMatch[1]);
      if (pval > 0 && pval < 10000) { amountTax = pval; amountNoTax = pval; }
    }
  }
  // 车票: amount near train keywords
  if (!amountTax && !amountNoTax) {
    var yMatch = fullText.match(/¥\s*(\d+(?:,\d{3})*\.\d{2})\s*[^\d]*?(?:车票|车次|座位|检票|进站|出站|乘车|站台)/);
    if (!yMatch) yMatch = fullText.match(/(?:车票|车次|座位|检票|进站|出站|乘车|站台)[^\d]*?¥\s*(\d+(?:,\d{3})*\.\d{2})/);
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
  // 车票: "全价" or "优惠价" pattern
  if (!amountTax && !amountNoTax) {
    var discountMatch = fullText.match(/(?:全\s*价|优\s*惠\s*价|学\s*生\s*价)[^\d]*?(\d+\.\d{2})/);
    if (discountMatch) {
      var dval = parseAmt(discountMatch[1]);
      if (dval > 0 && dval < 10000) { amountTax = dval; amountNoTax = dval; }
    }
  }
  // 车票: "￥" (full-width yen sign) pattern
  if (!amountTax && !amountNoTax) {
    var fwyMatch = fullText.match(/￥\s*(\d+(?:,\d{3})*\.\d{2})/);
    if (fwyMatch) {
      var fyVal = parseAmt(fwyMatch[1]);
      if (fyVal > 0 && fyVal < 10000) { amountTax = fyVal; amountNoTax = fyVal; }
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
        if (yv > 0 && yv < 5000) allYen.push(yv);
      }
      // Also check ￥ (full-width)
      var yenRe2 = /￥\s*(\d+(?:,\d{3})*\.\d{2})/g;
      while ((ym = yenRe2.exec(fullText)) !== null) {
        var yv2 = parseAmt(ym[1]);
        if (yv2 > 0 && yv2 < 5000) allYen.push(yv2);
      }
      // Also look for amounts without ¥ symbol — pattern: "票价 553.00" or standalone amounts after keywords
      if (allYen.length === 0) {
        var bareAmtRe = /(?:票\s*价|全\s*价|金\s*额)[^\d]*?(\d+\.\d{2})/gi;
        var bm;
        while ((bm = bareAmtRe.exec(fullText)) !== null) {
          var bv = parseAmt(bm[1]);
          if (bv > 0 && bv < 5000) allYen.push(bv);
        }
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

  // === Step 6.6: 出租车票 / 乘车票 — specific patterns ===
  if (!amountTax && !amountNoTax) {
    var taxiMatch = fullText.match(/(?:乘\s*车|出\s*租|打\s*车|网\s*约\s*车)[^\n]*?(?:金\s*额|费\s*用|价\s*格)[^\d]*?(\d+\.\d{2})/i);
    if (taxiMatch) {
      var tval = parseAmt(taxiMatch[1]);
      if (tval > 0 && tval < 5000) { amountTax = tval; amountNoTax = tval; }
    }
  }

  // === Step 6.7: 定额发票 — amount right after "¥" in short texts ===
  if (!amountTax && !amountNoTax) {
    // 定额发票 usually very short, contains just "金额 ¥X.00"
    if (fullText.length < 500) {
      var dingEMatch = fullText.match(/金\s*额[^\d]*?¥?\s*(\d+\.\d{2})/);
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
      if (yv2a > 0 && yv2a < 50000) amounts.push(yv2a);
    }
    // Also check ￥ (full-width yen sign)
    var yenRe2b = /￥\s*(\d+(?:,\d{3})*\.\d{2})/g;
    while ((ym2 = yenRe2b.exec(fullText)) !== null) {
      var yv2b = parseAmt(ym2[1]);
      if (yv2b > 0 && yv2b < 50000) amounts.push(yv2b);
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
    } else {
      // No tax info — this is likely a non-VAT invoice, amountTax = amountNoTax is correct
      amountTax = amountNoTax;
    }
  }
  // If only amountTax found, derive amountNoTax from taxAmount
  if (amountTax > 0 && amountNoTax === 0 && taxAmount > 0 && taxAmount < amountTax) {
    amountNoTax = Math.round((amountTax - taxAmount) * 100) / 100;
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
    var info = extractInvoiceInfo(ocrResult);
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
    // Skip seller info for tickets
    if (!info.isTicket) {
      if (info.sellerName) fileObj.sellerName = info.sellerName;
      if (info.sellerCreditCode) fileObj.sellerCreditCode = info.sellerCreditCode;
    }
  } catch(e) {
    console.warn('[OCR] 识别失败:', e);
  }
}
