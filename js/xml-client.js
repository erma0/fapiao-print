// =====================================================
// XML Client — Pure frontend XML 数电票 parser
// =====================================================
// Ported from invoice-engine/src/lib.rs parse_xml_invoice_content (Rust)
// XML: DOMParser | Field extraction via element path tracking (SAX-style)
//
// XML 数电票 is a pure data format (no layout info) — cannot be rendered as a visual page.
// Extracts structured invoice fields for file list display, summary export, and batch rename.

var __xmlClient = (function() {

  /**
   * Parse XML 数电票 content string and extract structured invoice data.
   * Replicates Rust parse_xml_invoice_content() logic exactly.
   *
   * @param {string} content - Raw XML string containing <EInvoice> root
   * @returns {Object} Invoice info with camelCase field names
   */
  function parseXmlInvoiceContent(content) {
    var info = {
      invoiceNo: null,
      invoiceDate: null,
      sellerName: null,
      sellerTaxId: null,
      buyerName: null,
      buyerTaxId: null,
      amountNoTax: null,
      taxAmount: null,
      amountTax: null,
      invoiceType: null
    };

    // Quick check: must contain <EInvoice> root element
    if (content.indexOf('<EInvoice') < 0) {
      return info;
    }

    // Strip namespace prefixes (lesson from OFD: <ofd:Page> won't match getElementsByTagName)
    var xmlStr = _stripXmlNs(content);
    var doc = new DOMParser().parseFromString(xmlStr, 'text/xml');

    if (doc.querySelector('parsererror')) {
      console.warn('XML parse error, trying raw content');
      return info;
    }

    // Track LabelName values from different parent contexts
    var einvoiceTypeLabel = null;
    var generalOrSpecialLabel = null;

    // --- InvoiceNumber ---
    var invoiceNoEl = _findFirstText(doc, 'InvoiceNumber');
    if (invoiceNoEl) info.invoiceNo = invoiceNoEl;

    // --- IssueTime ---
    var issueTimeEl = _findFirstText(doc, 'IssueTime');
    if (issueTimeEl) {
      // Only take date portion before 'T'
      info.invoiceDate = issueTimeEl.split('T')[0];
    }

    // --- Seller ---
    var sellerNameEl = _findFirstText(doc, 'SellerName');
    if (sellerNameEl) info.sellerName = sellerNameEl;

    var sellerIdEl = _findFirstText(doc, 'SellerIdNum');
    // Skip empty values (same as Rust: SellerIdNum non-empty check)
    if (sellerIdEl && sellerIdEl.trim()) info.sellerTaxId = sellerIdEl.trim();

    // --- Buyer ---
    var buyerNameEl = _findFirstText(doc, 'BuyerName');
    if (buyerNameEl) info.buyerName = buyerNameEl;

    var buyerIdEl = _findFirstText(doc, 'BuyerIdNum');
    // Skip empty values (personal invoices have empty BuyerIdNum → None in Rust)
    if (buyerIdEl && buyerIdEl.trim()) info.buyerTaxId = buyerIdEl.trim();

    // --- Amounts ---
    var amountNoTaxEl = _findFirstText(doc, 'TotalAmWithoutTax');
    if (amountNoTaxEl) info.amountNoTax = parseFloat(amountNoTaxEl);

    var taxAmountEl = _findFirstText(doc, 'TotalTaxAm');
    if (taxAmountEl) info.taxAmount = parseFloat(taxAmountEl);

    // TotalTax-includedAmount — tag name contains hyphen
    // DOMParser handles hyphens in tag names fine
    var amountTaxEl = _findFirstText(doc, 'TotalTax-includedAmount');
    if (amountTaxEl) {
      info.amountTax = parseFloat(amountTaxEl);
    }

    // Fallback: if amount_tax still empty, try alternate tag name
    // Some XML variants use TotalTaxIncludedAmount (no hyphen)
    if (info.amountTax == null || isNaN(info.amountTax)) {
      var altAmountTaxEl = _findFirstText(doc, 'TotalTaxIncludedAmount');
      if (altAmountTaxEl) info.amountTax = parseFloat(altAmountTaxEl);
    }

    // --- Invoice type: collect LabelName from different parent contexts ---
    // EInvoiceType/LabelName → einvoice_type_label
    // GeneralOrSpecialVAT/LabelName → general_or_special_label
    var einvoiceTypeEls = doc.getElementsByTagName('EInvoiceType');
    for (var i = 0; i < einvoiceTypeEls.length; i++) {
      var labelEl = _getChildText(einvoiceTypeEls[i], 'LabelName');
      if (labelEl && !einvoiceTypeLabel) {
        einvoiceTypeLabel = labelEl;
      }
    }

    var generalOrSpecialEls = doc.getElementsByTagName('GeneralOrSpecialVAT');
    for (var j = 0; j < generalOrSpecialEls.length; j++) {
      var labelEl2 = _getChildText(generalOrSpecialEls[j], 'LabelName');
      if (labelEl2 && !generalOrSpecialLabel) {
        generalOrSpecialLabel = labelEl2;
      }
    }

    // Compose invoice_type: "电子发票(普通发票)" or "电子发票(增值税专用发票)"
    if (generalOrSpecialLabel) {
      var prefix = einvoiceTypeLabel || '电子发票';
      info.invoiceType = prefix + '(' + generalOrSpecialLabel + ')';
    } else if (einvoiceTypeLabel) {
      info.invoiceType = einvoiceTypeLabel;
    }

    return info;
  }

  /**
   * Parse XML 数电票 from File object.
   * @param {File} file
   * @returns {Promise<Object>} Invoice info
   */
  async function parseXmlInvoice(file) {
    var text = await file.text();
    return parseXmlInvoiceContent(text);
  }

  /**
   * Parse XML 数电票 from text content string.
   * @param {string} content - XML string
   * @returns {Object} Invoice info
   */
  function parseXmlInvoiceFromText(content) {
    return parseXmlInvoiceContent(content);
  }

  // =====================================================
  // Internal helpers
  // =====================================================

  // _stripXmlNs / _parseXml are defined in xml-utils.js (loaded before this file)

  /**
   * Find first element by tag name and return its text content.
   * Returns null if not found or empty.
   */
  function _findFirstText(doc, tagName) {
    var els = doc.getElementsByTagName(tagName);
    if (!els || els.length === 0) return null;
    var text = els[0].textContent;
    if (!text || !text.trim()) return null;
    return text.trim();
  }

  /**
   * Get direct child element's text content by tag name.
   * Used for LabelName disambiguation (only look at direct children of a specific parent).
   */
  function _getChildText(parentEl, childTagName) {
    if (!parentEl) return null;
    var children = parentEl.getElementsByTagName(childTagName);
    if (!children || children.length === 0) return null;
    var text = children[0].textContent;
    if (!text || !text.trim()) return null;
    return text.trim();
  }

  // Public API
  return {
    parseXmlInvoice: parseXmlInvoice,
    parseXmlInvoiceFromText: parseXmlInvoiceFromText
  };

})();

window.__xmlClient = __xmlClient;
