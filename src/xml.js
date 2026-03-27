/**
 * xml.js -- XML utility functions for OOXML manipulation.
 *
 * Regex-based, zero external dependencies. Ported and consolidated
 * from suggest-edit-safe.js and docx-patch.js.
 */

'use strict';

const fs = require('fs');
const crypto = require('crypto');

// ============================================================================
// NAMESPACE MAP
// ============================================================================

/**
 * OOXML namespace URIs keyed by common prefix.
 * @type {Object<string, string>}
 */
const NS = {
  w:    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r:    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp:   'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  a:    'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic:  'http://schemas.openxmlformats.org/drawingml/2006/picture',
  rel:  'http://schemas.openxmlformats.org/package/2006/relationships',
  ct:   'http://schemas.openxmlformats.org/package/2006/content-types',
  w14:  'http://schemas.microsoft.com/office/word/2010/wordml',
  w15:  'http://schemas.microsoft.com/office/word/2012/wordml',
  mc:   'http://schemas.openxmlformats.org/markup-compatibility/2006',
  v:    'urn:schemas-microsoft-com:vml',
  o:    'urn:schemas-microsoft-com:office:office',
  m:    'http://schemas.openxmlformats.org/officeDocument/2006/math',
  wps:  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wne:  'http://schemas.microsoft.com/office/word/2006/wordml',
};

// ============================================================================
// ESCAPING
// ============================================================================

/**
 * Escape a string for safe inclusion in XML text or attributes.
 * @param {string} str - Raw string
 * @returns {string} XML-safe string
 */
function escapeXml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Decode XML entities back to literal characters.
 * @param {string} str - XML-encoded string
 * @returns {string} Decoded string
 */
function decodeXml(str) {
  return str
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&');
}

// ============================================================================
// TEXT EXTRACTION
// ============================================================================

/**
 * Extract visible text from a paragraph XML fragment.
 * Concatenates all <w:t> content, skipping deleted text (<w:delText>).
 * @param {string} paragraphXml - Raw XML of a <w:p> element
 * @returns {string} Plain text content
 */
function extractText(paragraphXml) {
  const texts = [];
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let m;
  while ((m = re.exec(paragraphXml)) !== null) {
    texts.push(m[1]);
  }
  return texts.join('');
}

/**
 * Extract visible text and decode XML entities.
 * @param {string} paragraphXml - Raw XML of a <w:p> element
 * @returns {string} Decoded plain text content
 */
function extractTextDecoded(paragraphXml) {
  return decodeXml(extractText(paragraphXml));
}

// ============================================================================
// ELEMENT FINDING
// ============================================================================

/**
 * Find all <w:p> elements in a document XML string.
 * Returns position info for in-place modification (no split/reassemble).
 * @param {string} docXml - Full document.xml content
 * @returns {Array<{xml: string, start: number, end: number, text: string}>}
 */
function findParagraphs(docXml) {
  const paragraphs = [];
  const pStartRe = /<w:p[\s>]/g;
  let m;
  while ((m = pStartRe.exec(docXml)) !== null) {
    const startIdx = m.index;
    const closeTag = '</w:p>';
    const closeIdx = docXml.indexOf(closeTag, startIdx);
    if (closeIdx === -1) continue;
    const endIdx = closeIdx + closeTag.length;
    const pXml = docXml.slice(startIdx, endIdx);
    const text = extractText(pXml);
    paragraphs.push({ xml: pXml, start: startIdx, end: endIdx, text });
  }
  return paragraphs;
}

/**
 * Find the first element that matches a tag name and optionally contains
 * a text substring.
 * @param {string} xml - XML string to search
 * @param {string} tag - Element tag name (e.g. 'w:p', 'w:r')
 * @param {string} [contains] - Optional text the element must contain
 * @returns {{xml: string, start: number, end: number}|null}
 */
function findElement(xml, tag, contains) {
  const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const openRe = new RegExp(`<${escaped}[\\s>]`, 'g');
  const closeTag = `</${tag}>`;
  let m;
  while ((m = openRe.exec(xml)) !== null) {
    const startIdx = m.index;
    const closeIdx = xml.indexOf(closeTag, startIdx);
    if (closeIdx === -1) continue;
    const endIdx = closeIdx + closeTag.length;
    const fragment = xml.slice(startIdx, endIdx);
    if (contains === undefined || fragment.includes(contains)) {
      return { xml: fragment, start: startIdx, end: endIdx };
    }
  }
  return null;
}

// ============================================================================
// ANCHOR SEARCH
// ============================================================================

/**
 * Find the index of the paragraph whose text contains the anchor string.
 * Tries exact match first, then substring, then case-insensitive substring.
 *
 * @param {Array<{xml?: string, text?: string}>|string[]} paragraphs - Array of paragraph objects or XML strings
 * @param {string} anchor - Text to search for
 * @returns {number} Index of the matching paragraph, or -1
 */
function findAnchorParagraph(paragraphs, anchor) {
  // Exact match
  for (let i = 0; i < paragraphs.length; i++) {
    const text = typeof paragraphs[i] === 'string' ? extractText(paragraphs[i]) : paragraphs[i].text;
    if (text === anchor) return i;
  }
  // Substring match
  for (let i = 0; i < paragraphs.length; i++) {
    const text = typeof paragraphs[i] === 'string' ? extractText(paragraphs[i]) : paragraphs[i].text;
    if (text.includes(anchor)) return i;
  }
  // Case-insensitive substring
  const lower = anchor.toLowerCase();
  for (let i = 0; i < paragraphs.length; i++) {
    const text = typeof paragraphs[i] === 'string' ? extractText(paragraphs[i]) : paragraphs[i].text;
    if (text.toLowerCase().includes(lower)) return i;
  }
  return -1;
}

// ============================================================================
// ATTRIBUTE EXTRACTION
// ============================================================================

/**
 * Extract an attribute value from an XML attribute string using regex.
 *
 * @param {string} attrs - Attribute string (e.g., 'w:id="5" w:author="Alice"')
 * @param {string} name - Attribute name (e.g., 'w:id')
 * @returns {string|null} Attribute value or null if not found
 */
function attrVal(attrs, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(escaped + '="([^"]*)"');
  const m = attrs.match(re);
  return m ? m[1] : null;
}

// ============================================================================
// ID GENERATION
// ============================================================================

/**
 * Scan document XML for all w:id="N" attributes and return max + 1.
 * Used for tracked change IDs.
 * @param {string} docXml - Document XML content
 * @returns {number} Next available change ID
 */
function nextChangeId(docXml) {
  let max = 0;
  const re = /w:id="(\d+)"/g;
  let m;
  while ((m = re.exec(docXml)) !== null) {
    const n = parseInt(m[1], 10);
    if (n > max) max = n;
  }
  return max + 1;
}

/**
 * Scan comments XML for the highest comment id attribute and return max + 1.
 * Looks for w:id="N" on w:comment elements.
 * @param {string} xml - comments.xml content
 * @returns {number} Next available comment ID
 */
function nextCommentId(xml) {
  let max = -1;
  const re = /<w:comment\b[^>]*\bw:id="(\d+)"/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const n = parseInt(m[1], 10);
    if (n > max) max = n;
  }
  return max + 1;
}

/**
 * Scan relationship XML for the highest rIdN number and return the next rId string.
 * @param {string} relsXml - document.xml.rels content
 * @returns {string} Next rId (e.g. "rId15")
 */
function nextRId(relsXml) {
  let max = 0;
  const re = /Id="rId(\d+)"/g;
  let m;
  while ((m = re.exec(relsXml)) !== null) {
    const n = parseInt(m[1], 10);
    if (n > max) max = n;
  }
  return `rId${max + 1}`;
}

/**
 * Generate an 8-character uppercase hex ID.
 * Used for w14:paraId, w14:textId, and similar attributes.
 * @returns {string} e.g. "3A4F1B2C"
 */
function randomHexId() {
  return crypto.randomBytes(4).toString('hex').toUpperCase();
}

/**
 * Generate a UUID v4 string.
 * Used for durable IDs and other globally unique identifiers.
 * @returns {string} e.g. "550e8400-e29b-41d4-a716-446655440000"
 */
function randomUUID() {
  const bytes = crypto.randomBytes(16);
  // Set version (4) and variant (RFC 4122)
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = bytes.toString('hex');
  return [
    hex.slice(0, 8),
    hex.slice(8, 12),
    hex.slice(12, 16),
    hex.slice(16, 20),
    hex.slice(20, 32),
  ].join('-');
}

// ============================================================================
// RUN PARSING
// ============================================================================

/**
 * Parse all <w:r> elements from a paragraph XML fragment.
 * Returns run metadata including formatting properties and text content.
 * @param {string} paragraphXml - Raw XML of a <w:p> element
 * @returns {Array<{fullMatch: string, rPr: string, texts: Array<{openTag: string, text: string, fullMatch: string}>, index: number, combinedText: string}>}
 */
function parseRuns(paragraphXml) {
  const runs = [];
  const runRe = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
  let m;
  while ((m = runRe.exec(paragraphXml)) !== null) {
    const runXml = m[0];
    const runIndex = m.index;

    const rPrMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    const rPr = rPrMatch ? rPrMatch[0] : '';

    const texts = [];
    const tRe = /(<w:t[^>]*>)([^<]*)<\/w:t>/g;
    let tMatch;
    while ((tMatch = tRe.exec(runXml)) !== null) {
      texts.push({
        openTag: tMatch[1],
        text: tMatch[2],
        fullMatch: tMatch[0],
      });
    }

    const combinedText = texts.map(t => t.text).join('');

    runs.push({
      fullMatch: runXml,
      rPr,
      texts,
      index: runIndex,
      combinedText,
    });
  }
  return runs;
}

/**
 * Extract the <w:rPr>...</w:rPr> block from a run XML fragment.
 * Returns the full rPr element or an empty string if none.
 * @param {string} runXml - Raw XML of a <w:r> element
 * @returns {string}
 */
function extractRpr(runXml) {
  const m = runXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

// ============================================================================
// XML FRAGMENT BUILDERS
// ============================================================================

/**
 * Build a tracked deletion element.
 * @param {number} id - Change tracking ID
 * @param {string} author - Author name
 * @param {string} date - ISO date string
 * @param {string} rPr - Run properties XML (may be empty string)
 * @param {string} text - Deleted text (will be XML-escaped)
 * @returns {string} <w:del> XML fragment
 */
function buildDel(id, author, date, rPr, text) {
  return `<w:del w:id="${id}" w:author="${escapeXml(author)}" w:date="${date}">`
    + `<w:r>${rPr}<w:delText xml:space="preserve">${escapeXml(text)}</w:delText></w:r>`
    + `</w:del>`;
}

/**
 * Build a tracked insertion element.
 * @param {number} id - Change tracking ID
 * @param {string} author - Author name
 * @param {string} date - ISO date string
 * @param {string} rPr - Run properties XML (may be empty string)
 * @param {string} text - Inserted text (will be XML-escaped)
 * @returns {string} <w:ins> XML fragment
 */
function buildIns(id, author, date, rPr, text) {
  return `<w:ins w:id="${id}" w:author="${escapeXml(author)}" w:date="${date}">`
    + `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`
    + `</w:ins>`;
}

/**
 * Build a simple run element.
 * @param {string} rPr - Run properties XML (may be empty string)
 * @param {string} text - Text content (will be XML-escaped)
 * @returns {string} <w:r> XML fragment
 */
function buildRun(rPr, text) {
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

/**
 * Build a paragraph element.
 * @param {string} pPr - Paragraph properties XML (may be empty string)
 * @param {string|string[]} runs - Run XML fragments (string or array of strings)
 * @returns {string} <w:p> XML fragment
 */
function buildParagraph(pPr, runs) {
  const paraId = randomHexId();
  const textId = randomHexId();
  const runsStr = Array.isArray(runs) ? runs.join('') : runs;
  return `<w:p w14:paraId="${paraId}" w14:textId="${textId}">`
    + (pPr ? pPr : '')
    + runsStr
    + `</w:p>`;
}

// ============================================================================
// DATE HELPERS
// ============================================================================

/**
 * Return the current time as an OOXML-compatible ISO string.
 * Format: "2026-03-27T14:30:00Z" (no milliseconds).
 * @returns {string}
 */
function isoNow() {
  return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// ============================================================================
// IMAGE HELPERS
// ============================================================================

/**
 * Read PNG dimensions from the IHDR chunk.
 * Falls back to JPEG parsing if the file is not PNG.
 * @param {string} filePath - Path to image file
 * @returns {{width: number, height: number}}
 */
function getPngDimensions(filePath) {
  const buf = fs.readFileSync(filePath);
  // PNG signature: 0x89 0x50
  if (buf[0] !== 0x89 || buf[1] !== 0x50) {
    return getJpegDimensions(buf);
  }
  const width = buf.readUInt32BE(16);
  const height = buf.readUInt32BE(20);
  return { width, height };
}

/**
 * Read JPEG dimensions from SOF0 or SOF2 markers.
 * @param {Buffer} buf - Raw file buffer
 * @returns {{width: number, height: number}}
 */
function getJpegDimensions(buf) {
  let offset = 2;
  while (offset < buf.length) {
    if (buf[offset] !== 0xff) break;
    const marker = buf[offset + 1];
    if (marker === 0xc0 || marker === 0xc2) {
      const height = buf.readUInt16BE(offset + 5);
      const width = buf.readUInt16BE(offset + 7);
      return { width, height };
    }
    const len = buf.readUInt16BE(offset + 2);
    offset += 2 + len;
  }
  return { width: 800, height: 600 }; // fallback
}

/**
 * Convert inches to English Metric Units (EMU).
 * 1 inch = 914400 EMU.
 * @param {number} inches
 * @returns {number}
 */
function emuFromInches(inches) {
  return Math.round(inches * 914400);
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = {
  // Namespace helpers
  NS,

  // Escaping
  escapeXml,
  decodeXml,

  // Text extraction
  extractText,
  extractTextDecoded,

  // Element finding
  findParagraphs,
  findElement,
  findAnchorParagraph,

  // Attribute extraction
  attrVal,

  // ID generation
  nextChangeId,
  nextCommentId,
  nextRId,
  randomHexId,
  randomUUID,

  // Run manipulation
  parseRuns,
  extractRpr,

  // XML fragment builders
  buildDel,
  buildIns,
  buildRun,
  buildParagraph,

  // Date helpers
  isoNow,

  // Image helpers
  getPngDimensions,
  getJpegDimensions,
  emuFromInches,
};
