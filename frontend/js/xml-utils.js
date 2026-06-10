// =====================================================
// XML Utilities — shared namespace stripping & parsing
// =====================================================
// Used by ofd-client.js and xml-client.js.
// Must be loaded before both.

/**
 * Strip XML namespace prefixes before DOMParser parsing.
 * OFD/EInvoice files use prefixed tags like <ofd:Page>, <ns2:InvoiceNumber>,
 * but getElementsByTagName('Page') won't match <ofd:Page> in browser DOMParser.
 * Strip all prefixes so <ofd:Page> → <Page>, <ns2:InvoiceNumber> → <InvoiceNumber>, etc.
 * Also removes xmlns:xxx declarations to keep XML clean.
 */
function _stripXmlNs(xml) {
  // Remove xmlns declarations: xmlns:ofd="..." → (removed)
  var s = xml.replace(/\s+xmlns:\w+\s*=\s*"[^"]*"/g, '');
  // Also remove bare xmlns="..."
  s = s.replace(/\s+xmlns\s*=\s*"[^"]*"/g, '');
  // Strip namespace prefixes from opening/closing/self-closing tags:
  // <ofd:XXX → <XXX, </ofd:XXX → </XXX
  s = s.replace(/<(\/?)(\w+):/g, '<$1');
  return s;
}

/**
 * Parse XML string with namespace stripping.
 * All OFD/EInvoice XML should be parsed through this function.
 */
function _parseXml(xml) {
  return new DOMParser().parseFromString(_stripXmlNs(xml), 'text/xml');
}
