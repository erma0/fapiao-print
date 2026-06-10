// =====================================================
// PDF Text Extraction — applyPdfTextResult
// =====================================================
// Extracts invoice fields from PDF.js text layer result.
// Uses regex patterns on full text (simpler than coordinate-based extraction,
// but handles the most common structured invoice PDFs).
//
// This function was originally in ocr.js (deleted in d244d68).
// Recreated here as a compact text-based extractor.

var applyPdfTextResult = (function() {

function _normText(s) {
  s = s.replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  s = s.replace(/[Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  s = s.replace(/％/g, '%').replace(/．/g, '.').replace(/，/g, ',').replace(/：/g, ':');
  s = s.replace(/￥/g, '\u00A5');
  return s;
}

function _collapseCjkSpaces(s) {
  for (var i = 0; i < 5; i++) {
    var prev = '';
    while (prev !== s) { prev = s; s = s.replace(/([\u4e00-\u9fff])\s+([\u4e00-\u9fff])/g, '$1$2'); }
  }
  s = s.replace(/([\u4e00-\u9fff])\n([\u4e00-\u9fff])/g, '$1$2');
  return s;
}

function _collapseNumberSpaces(s) {
  for (var i = 0; i < 3; i++) {
    var prev = '';
    while (prev !== s) { prev = s; s = s.replace(/(\d)\s+(\d{3}\b)/g, '$1$2'); }
  }
  s = s.replace(/(\d)\s+\./g, '$1.');
  s = s.replace(/\u00A5\s+(\d)/g, '\u00A5$1');
  return s;
}

function _regexFind(keyword, text) {
  var escaped = keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var re = new RegExp(escaped + '[：:]\\s*([^\\n\\r]+)', 'i');
  var m = text.match(re);
  return m ? m[1].trim() : null;
}

function _parseAmt(s) {
  if (!s) return null;
  s = s.replace(/[,，\s]/g, '').replace(/^[¥\u00A5]/, '');
  var v = parseFloat(s);
  return isNaN(v) ? null : v;
}

function _isCreditCode(s) {
  if (!s) return false;
  s = s.replace(/\s/g, '');
  return /^[0-9A-HJ-NP-RT-Y]{18}$/i.test(s) || /^\d{15,20}$/.test(s);
}

function _extractFields(text, normText) {
  var info = {};

  info.invoiceNo = _regexFind('发票号码', normText) || _regexFind('发票代码', normText) ||
    (normText.match(/(\d{20})\s/) || [])[1] || null;

  var dateRaw = _regexFind('开票日期', normText);
  if (dateRaw) info.invoiceDate = dateRaw.split('T')[0].replace(/[\s\r\n]/g, '');

  info.buyerName = _regexFind('购买方名称', normText) || _regexFind('名称', normText) ||
    (normText.match(/名称[：:]\s*([^\n]{2,30})/) || [])[1] || null;
  info.sellerName = _regexFind('销售方名称', normText) ||
    (normText.match(/销售方名称[：:]\s*([^\n]{2,30})/) || [])[1] || null;

  // Credit codes: find 18-digit alphanumeric strings near "信用代码" or "识别号"
  var ccPattern = /(?:统一社会信用代码|纳税人识别号)[：:]\s*([^\n]{10,30})/g;
  var ccMatches = [];
  var m;
  while ((m = ccPattern.exec(normText)) !== null) {
    var code = m[1].replace(/\s/g, '');
    if (_isCreditCode(code)) ccMatches.push(code);
  }
  if (ccMatches.length >= 1) info.buyerCreditCode = ccMatches[0];
  if (ccMatches.length >= 2) info.sellerCreditCode = ccMatches[1];

  // Amounts
  var jiashuiMatch = normText.match(/(?:价税合计|大写).*?[¥\u00A5]\s*([\d,]+\.?\d*)/);
  if (jiashuiMatch) info.amountTax = _parseAmt(jiashuiMatch[1]);

  var noTaxMatch = normText.match(/(?:不含税金额.*?|[（(]小写[）)].*?)[¥\u00A5]\s*([\d,]+\.?\d*)/);
  if (noTaxMatch) info.amountNoTax = _parseAmt(noTaxMatch[1]);

  var taxMatch = normText.match(/税额.*?[¥\u00A5]\s*([\d,]+\.?\d*)/);
  if (taxMatch) info.taxAmount = _parseAmt(taxMatch[1]);

  // Invoice type
  if (/增值税专用/.test(normText)) info.invoiceType = '增值税专用发票';
  else if (/增值税普通|增值税电子普通/.test(normText)) info.invoiceType = '增值税普通发票';
  else if (/电子发票/.test(normText)) info.invoiceType = '电子发票';
  else if (/非税/.test(normText)) info.invoiceType = '非税票据';

  // Ticket detection
  var isTicket = /铁路电子客票|车票|出发站|到达站|车次|座位/.test(normText);
  if (isTicket) {
    info.isTicket = true;
    var ticketPrice = normText.match(/票\s*价[：:]*\s*[¥\u00A5]?\s*(\d+\.\d{2})/);
    if (ticketPrice) info.amountTax = _parseAmt(ticketPrice[1]);
  }

  info.isNonTax = /非税|票据号码|票据代码|交款人/.test(normText) && !/增值税/.test(normText);

  return info;
}

function applyPdfTextResult(fileObj, pdfTextResult) {
  if (!pdfTextResult || !pdfTextResult.lines || pdfTextResult.lines.length === 0) return;
  try {
    var fullText = pdfTextResult.text || '';
    var normText = _collapseNumberSpaces(_collapseCjkSpaces(_normText(fullText)));

    var info = _extractFields(fullText, normText);

    if (info.isTicket) fileObj._isTicket = true;

    if (!info.isTicket && !info.isNonTax) {
      if (info.invoiceNo && !fileObj.invoiceNo) fileObj.invoiceNo = info.invoiceNo;
      if (info.invoiceDate && !fileObj.invoiceDate) fileObj.invoiceDate = info.invoiceDate;
      if (info.buyerName && !fileObj.buyerName) fileObj.buyerName = info.buyerName;
      if (info.buyerCreditCode && !fileObj.buyerCreditCode) fileObj.buyerCreditCode = info.buyerCreditCode;
      if (info.sellerName && !fileObj.sellerName) fileObj.sellerName = info.sellerName;
      if (info.sellerCreditCode && !fileObj.sellerCreditCode) fileObj.sellerCreditCode = info.sellerCreditCode;
    }

    // Amounts — only fill if not already set
    var effAmt = info.amountTax > 0 ? info.amountTax : info.amountNoTax;
    if (effAmt > 0 && !fileObj.amountTax && !fileObj.amountNoTax) {
      fileObj.amount = effAmt;
      fileObj.amountTax = info.amountTax || 0;
      fileObj.amountNoTax = info.amountNoTax || 0;
      fileObj.taxAmount = info.taxAmount || 0;
    }

    // Amount validation: 含税 ≈ 不含税 + 税额
    if (fileObj.amountTax > 0 && fileObj.amountNoTax > 0 && fileObj.taxAmount > 0) {
      var sum = Math.round((fileObj.amountNoTax + fileObj.taxAmount) * 100) / 100;
      if (Math.abs(sum - fileObj.amountTax) > 0.02) {
        var validRates = [0, 0.01, 0.03, 0.05, 0.06, 0.09, 0.13];
        var recalc = Math.round((fileObj.amountTax - fileObj.taxAmount) * 100) / 100;
        if (recalc > fileObj.taxAmount) {
          var rate = Math.round(fileObj.taxAmount / recalc * 10000) / 10000;
          if (validRates.some(function(r) { return Math.abs(rate - r) < 0.005; })) {
            fileObj.amountNoTax = recalc;
          }
        }
      }
    }

    // Non-tax: force amountNoTax
    if (info.isNonTax && info.amountNoTax > 0 && !fileObj.amountNoTax) {
      fileObj.amountNoTax = info.amountNoTax;
      if (!fileObj.invoiceType) fileObj.invoiceType = '非税票据';
    }

    if (info.invoiceType && !fileObj.invoiceType) fileObj.invoiceType = info.invoiceType;

    fileObj._pdfTextExtracted = true;
  } catch(e) {
    console.warn('[PDF文字提取] 结果应用失败:', e);
  }
}

return applyPdfTextResult;

})();
