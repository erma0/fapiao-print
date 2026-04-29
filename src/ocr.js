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

/**
 * Extract invoice info from OCR text or PDF.js text content.
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
  if (typeof textContent === 'string') {
    fullText = textContent;
  } else if (textContent && textContent.items && textContent.items.length) {
    fullText = textContent.items.map(function(item) { return item.str; }).join('');
  } else {
    return { amountTax: 0, amountNoTax: 0, sellerName: '', sellerCreditCode: '', isTicket: false };
  }

  // === Detect ticket type early ===
  var isTicket = isTicketText(fullText);
  var sellerName = '', sellerCreditCode = '';

  // === Skip seller extraction for tickets (train/ride tickets have no seller) ===
  if (!isTicket) {
    // === Extract seller info BEFORE normalization (uses raw text with line breaks) ===

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
    if (lastCcCode) sellerCreditCode = lastCcCode.toUpperCase();

    // Also try standalone credit codes without the prefix label (some OCR misses the label)
    if (!lastCcCode) {
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

    // Seller name — multiple strategies in priority order:
    // Strategy 1: Direct "销售方(+信息)" + "名称:" pattern (most specific)
    var snMatch = fullText.match(/销\s*售\s*方(?:信息)?\s*名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
    if (snMatch) {
      sellerName = snMatch[1].trim();
      // Clean up: remove trailing section labels that got captured
      sellerName = sellerName.replace(/\s*(?:购买方|销售方|信息|名称|纳税人|统一社会|地址|开户行|电话|账号).*$/i, '');
    }

    // Strategy 2: Find "名称:" AFTER "销售方" keyword (near seller's section)
    if (!sellerName) {
      var sellerKwMatch = fullText.match(/销\s*售\s*方[^\n]{0,100}?名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
      if (sellerKwMatch) {
        sellerName = sellerKwMatch[1].trim();
      }
    }

    // Strategy 2.5: "销方" abbreviated form (some invoices use short form)
    if (!sellerName) {
      var shortSellerMatch = fullText.match(/销\s*方(?:信息)?\s*名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
      if (shortSellerMatch) {
        sellerName = shortSellerMatch[1].trim();
      }
    }

    // Strategy 3: Find "名称:" in the region AFTER the buyer's credit code (i.e., seller's section)
    if (!sellerName && lastCcPos >= 0) {
      // Search from first credit code position onward — find "名称:" AFTER any credit code
      // The seller's section comes after the buyer's section
      var searchStart = firstCcPos;
      var searchRegion = fullText.substring(searchStart);
      var nameRe = /名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/gi;
      var nm, lastName = '';
      while ((nm = nameRe.exec(searchRegion)) !== null) {
        var candidate = nm[1].trim();
        // Filter out buyer section labels that may have been captured
        if (!/^(?:购买方|信息|名称)/.test(candidate) && candidate.length > 1) {
          lastName = candidate;
        }
      }
      if (lastName) sellerName = lastName;
    }

    // Strategy 4: Find the LAST "名称:" in full text (after last credit code position)
    if (!sellerName && lastCcPos >= 0) {
      var regionAfterLastCc = fullText.substring(lastCcPos);
      var nameAfter = regionAfterLastCc.match(/名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号)|\n|$)/i);
      if (nameAfter) {
        var cand = nameAfter[1].trim();
        if (cand.length > 1 && !/^(?:购买方|信息|名称)/.test(cand)) {
          sellerName = cand;
        }
      }
    }

    // Strategy 5: No credit code found at all — try any "名称:" but filter buyer keywords
    if (!sellerName) {
      var allNames = [];
      var fallbackRe = /名\s*称\s*[:：]\s*([^\n]{1,80}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号)|\n|$)/gi;
      var fm;
      while ((fm = fallbackRe.exec(fullText)) !== null) {
        var c = fm[1].trim();
        if (c.length > 1 && !/^(?:购买方|信息|名称)/.test(c)) {
          allNames.push(c);
        }
      }
      // If we found multiple names, the LAST one is likely the seller's
      if (allNames.length > 0) sellerName = allNames[allNames.length - 1];
    }

    // Strategy 6: "收款单位"/"销货单位"/"开票方" pattern (some invoice formats use these)
    if (!sellerName) {
      var altSellerMatch = fullText.match(/(?:收款单位|销货单位|开票方|销售单位|开票人|代开企业)[^\n]{0,30}?[:：]?\s*([^\n]{2,60}?)(?=\s*(?:纳税人|统一社会|地址|开户行|电话|账号|[a-zA-Z0-9]{15,20})|\n|$)/i);
      if (altSellerMatch) {
        var altCand = altSellerMatch[1].trim();
        if (altCand.length > 1 && !/^(?:购买方|信息|名称|地址|电话)/.test(altCand)) {
          sellerName = altCand;
        }
      }
    }

    // Strategy 7: Use credit code position — find company name near the last credit code
    // Some OCR outputs have: "91440300xxxxxxxxx  深圳市某某科技有限公司"
    if (!sellerName && lastCcPos >= 0) {
      var afterLastCc = fullText.substring(lastCcPos);
      // Look for Chinese company name pattern near credit code
      var companyRe = new RegExp('(?:[A-Z0-9]{15,20})\\s*[:：]?\\s*([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]+' + companySuffix + ')');
      var compMatch = afterLastCc.match(companyRe);
      if (compMatch) {
        sellerName = compMatch[1].trim();
      }
    }

    // Strategy 8: Find any Chinese company name after "销售方" keyword
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

    // Strategy 8.5: Find company name after "销方" keyword (short form)
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

    // Strategy 9: Last resort — find company name patterns in text
    // If we have credit codes, even 1 company name is likely the seller
    // If no credit codes, need ≥2 to distinguish buyer from seller
    if (!sellerName) {
      var allCompanies = [];
      var companyGlobalRe = new RegExp('([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]{2,25}' + companySuffix + ')', 'g');
      var cm2;
      while ((cm2 = companyGlobalRe.exec(fullText)) !== null) {
        var cname = cm2[1].trim();
        // Filter out common non-seller company names
        if (cname.length > 3 && !/^(?:购买方|销售方|信息|名称|地址)/.test(cname)) {
          allCompanies.push(cname);
        }
      }
      if (allCcPositions.length >= 1 && allCompanies.length >= 1) {
        // We have at least one credit code — the last company name is likely the seller
        sellerName = allCompanies[allCompanies.length - 1];
      } else if (allCompanies.length >= 2) {
        // No credit codes but multiple companies — last one is likely seller
        sellerName = allCompanies[allCompanies.length - 1];
      }
    }

    // Strategy 10: Region between buyer and seller credit codes
    // If we have ≥2 credit codes, the text between the second-to-last and last code
    // often contains the seller's company name
    if (!sellerName && allCcPositions.length >= 2) {
      var secondLastPos = allCcPositions[allCcPositions.length - 2].pos;
      var regionBetween = fullText.substring(secondLastPos, lastCcPos);
      var betweenCompany = regionBetween.match(new RegExp('([\\u4e00-\\u9fff][\\u4e00-\\u9fff\\w（）()·\\-\\.]+' + companySuffix + ')'));
      if (betweenCompany) {
        var bCand = betweenCompany[1].trim();
        if (bCand.length > 2 && !/^(?:购买方|信息|名称|地址)/.test(bCand)) {
          sellerName = bCand;
        }
      }
    }

    // Strategy 11: Find company name after last credit code (broader pattern)
    // Some OCR outputs have company name on the line AFTER the credit code
    if (!sellerName && lastCcPos >= 0) {
      var afterCc = fullText.substring(lastCcPos);
      // Match Chinese characters that look like a company name (broader than strict suffix pattern)
      var broadCompanyRe = /[\u4e00-\u9fff]{2,4}(?:[\u4e00-\u9fff\w（）()·\-\.]*)(?:公司|集团|商行|商店|厂|部|院|所|中心|店|馆|站|社|行|会|处|有限合伙|合伙企业|个体|工作室|经营部|门市|分公司|事业部|事务所|医院|学校|合作社|企业|商社)/;
      var broadMatch = afterCc.match(broadCompanyRe);
      if (broadMatch) {
        sellerName = broadMatch[0].trim();
      }
    }

    // Strategy 12: Find the last "名称:" with more relaxed content matching
    // Some OCR outputs have company name after "名称:" but with unusual characters
    if (!sellerName) {
      var allNameEntries = [];
      var relaxedNameRe = /名\s*称\s*[:：]\s*([^\n]{2,80}?)(?=\n|$)/gi;
      var rnm;
      while ((rnm = relaxedNameRe.exec(fullText)) !== null) {
        var rCand = rnm[1].trim();
        // Must contain at least 2 Chinese characters and look like a business name
        if (/[\u4e00-\u9fff]{2,}/.test(rCand) && !/^(?:购买方|销售方|信息|名称|地址|电话|纳税人)/.test(rCand)) {
          // Clean any trailing non-name content
          rCand = rCand.replace(/\s*(?:纳税人|统一社会|地址|开户行|电话|账号|复核|收款|开票).*$/i, '');
          if (rCand.length > 2) allNameEntries.push(rCand);
        }
      }
      if (allNameEntries.length > 0) {
        // Last entry is most likely the seller
        sellerName = allNameEntries[allNameEntries.length - 1];
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
  // Variant: 价税合计 without explicit 小写/大写, just ¥ directly
  if (!amountTax) {
    var pthMatch = fullText.match(/价\s*税\s*合\s*计[^\d]*?¥\s*(\d+(?:,\d{3})*\.\d{2})/);
    if (pthMatch) amountTax = parseAmt(pthMatch[1]);
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

  // AUTO-FALLBACK: if only amountNoTax found, auto-assign to amountTax
  if (amountNoTax > 0 && amountTax === 0) {
    amountTax = amountNoTax;
  }

  console.log('[OCR提取] 金额:', { amountTax: amountTax, amountNoTax: amountNoTax }, '销售方:', sellerName || '(未识别)', '信用代码:', sellerCreditCode || '(未识别)', '车票:', isTicket);
  if (!amountTax && !amountNoTax && !sellerName && !isTicket) {
    console.warn('[OCR提取] 未能识别任何信息，OCR完整文本:', fullText);
  }

  return { amountTax: amountTax, amountNoTax: amountNoTax, sellerName: sellerName, sellerCreditCode: sellerCreditCode, _ocrText: fullText, isTicket: isTicket };
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
    // Skip seller info for tickets
    if (!info.isTicket) {
      if (info.sellerName) fileObj.sellerName = info.sellerName;
      if (info.sellerCreditCode) fileObj.sellerCreditCode = info.sellerCreditCode;
    }
    fileObj._ocrText = info._ocrText || ocrText;
    fileObj._isTicket = info.isTicket || false;
  } catch(e) {
    console.warn('[OCR] 识别失败:', e);
  }
}
