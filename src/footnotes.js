/**
 * footnotes.js -- Footnote operations for docex
 *
 * Static methods for listing and adding footnotes in OOXML documents.
 * Manages three XML files:
 *   - word/document.xml (footnote references)
 *   - word/footnotes.xml (footnote content)
 *   - word/_rels/document.xml.rels + [Content_Types].xml (infrastructure)
 *
 * All methods operate on a Workspace object. XML manipulation is done
 * entirely with string operations and regex. Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ---------------------------------------------------------------------------
// Relationship type and content type constants
// ---------------------------------------------------------------------------
const REL_FOOTNOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes';
const CT_FOOTNOTES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml';

// ============================================================================
// FOOTNOTES
// ============================================================================

class Footnotes {

  /**
   * List all user footnotes in the document.
   *
   * Parses word/footnotes.xml and returns all footnotes except the two
   * built-in separator footnotes (id=0 and id=1).
   *
   * @param {object} ws - Workspace with ws.footnotesXml
   * @returns {Array<{id: number, text: string}>}
   */
  static list(ws) {
    const footnotesXml = ws.footnotesXml;
    if (!footnotesXml) return [];

    const results = [];
    const footnoteRe = /<w:footnote\b([^>]*)>([\s\S]*?)<\/w:footnote>/g;
    let m;

    while ((m = footnoteRe.exec(footnotesXml)) !== null) {
      const attrs = m[1];
      const body = m[2];

      const id = xml.attrVal(attrs, 'w:id');
      const idNum = id ? parseInt(id, 10) : 0;

      // Skip built-in separator footnotes (id=0 and id=1)
      if (idNum <= 1) continue;

      // Also skip if it has a w:type attribute (separator/continuationSeparator)
      const type = xml.attrVal(attrs, 'w:type');
      if (type) continue;

      const text = xml.extractText(body);
      results.push({ id: idNum, text });
    }

    return results;
  }

  /**
   * Add a new footnote anchored to specific text in the document.
   *
   * This modifies up to 4 XML files:
   *   1. document.xml: footnoteReference run inserted at anchor position
   *   2. footnotes.xml: the footnote element with text
   *   3. Relationships (if not already present)
   *   4. Content types (if not already present)
   *
   * @param {object} ws - Workspace
   * @param {string} anchor - Text to anchor the footnote to
   * @param {string} footnoteText - Footnote text
   * @param {object} [opts] - Options
   * @param {string} [opts.author] - Author (for tracked changes, not used currently)
   * @param {boolean} [opts.tracked] - Whether to track (not used currently)
   * @throws {Error} If anchor text is not found
   * @returns {{ footnoteId: number }} The new footnote's ID
   */
  static add(ws, anchor, footnoteText, opts = {}) {
    if (typeof footnoteText !== 'string' || footnoteText.length === 0) {
      throw new Error('add(): text must be a non-empty string');
    }

    // Ensure footnotes.xml exists
    _ensureFootnotesFile(ws);

    // Get next footnote ID (must be >= 2, since 0 and 1 are separators)
    const footnotesXml = ws.footnotesXml;
    const footnoteId = _nextFootnoteId(footnotesXml);

    // 1. Add footnote element to word/footnotes.xml
    const footnoteEl = '<w:footnote w:id="' + footnoteId + '">'
      + '<w:p>'
      + '<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
      + '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
      + '<w:r><w:t xml:space="preserve"> ' + xml.escapeXml(footnoteText) + '</w:t></w:r>'
      + '</w:p>'
      + '</w:footnote>';

    let fnXml = ws.footnotesXml;
    fnXml = fnXml.replace('</w:footnotes>', footnoteEl + '</w:footnotes>');
    ws.footnotesXml = fnXml;

    // 2. Insert footnote reference in document.xml at anchor position
    if (anchor) {
      let doc = ws.docXml;
      const paragraphs = xml.findParagraphs(doc);
      let anchorFound = false;

      for (const para of paragraphs) {
        const paraText = para.text;
        if (!paraText.includes(anchor)) continue;

        const paraXml = para.xml;

        // Find which runs contain the anchor text
        const runs = xml.parseRuns(paraXml);
        const textRuns = runs.filter(r => r.texts.length > 0);
        let charCount = 0;
        let endRunIdx = -1;

        const anchorStart = paraText.indexOf(anchor);
        const anchorEnd = anchorStart + anchor.length;

        for (let r = 0; r < textRuns.length; r++) {
          const runLen = textRuns[r].combinedText.length;
          const runEnd = charCount + runLen;

          if (runEnd >= anchorEnd) {
            endRunIdx = r;
            break;
          }
          charCount += runLen;
        }

        if (endRunIdx === -1) endRunIdx = textRuns.length > 0 ? textRuns.length - 1 : 0;

        const refRun = '<w:r>'
          + '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
          + '<w:footnoteReference w:id="' + footnoteId + '"/>'
          + '</w:r>';

        let modPara = paraXml;

        // Insert footnote reference run after the last matched run
        if (textRuns.length > 0 && endRunIdx < textRuns.length) {
          const endRun = textRuns[endRunIdx];
          const insertAfterPos = endRun.index + endRun.fullMatch.length;
          modPara = modPara.slice(0, insertAfterPos) + refRun + modPara.slice(insertAfterPos);
        }

        doc = doc.slice(0, para.start) + modPara + doc.slice(para.end);
        ws.docXml = doc;
        anchorFound = true;
        break;
      }

      if (!anchorFound) {
        throw new Error('add(): could not find anchor text "' + anchor + '" in document');
      }
    }

    return { footnoteId };
  }

  // --------------------------------------------------------------------------
  // INTERNAL HELPERS
  // --------------------------------------------------------------------------

}

// ---------------------------------------------------------------------------
// Internal helpers (module-level)
// ---------------------------------------------------------------------------

/**
 * Get the next available footnote ID from footnotes.xml.
 * Scans for the highest w:id on w:footnote elements and returns max + 1.
 * Always returns at least 2 (since 0 and 1 are reserved for separators).
 *
 * @param {string} footnotesXml - footnotes.xml content
 * @returns {number}
 * @private
 */
function _nextFootnoteId(footnotesXml) {
  let max = 1; // separators use 0 and 1, so start at 1
  const re = /<w:footnote\b[^>]*\bw:id="(\d+)"/g;
  let m;
  while ((m = re.exec(footnotesXml)) !== null) {
    const n = parseInt(m[1], 10);
    if (n > max) max = n;
  }
  return max + 1;
}

/**
 * Ensure footnotes.xml exists in the workspace, along with its
 * relationship and content type entries.
 *
 * @param {object} ws - The workspace
 * @private
 */
function _ensureFootnotesFile(ws) {
  // Accessing ws.footnotesXml creates it if missing (workspace handles this)
  void ws.footnotesXml;

  // Ensure relationship exists in rels file
  let relsXml = ws.relsXml;
  let relsChanged = false;

  if (relsXml && !relsXml.includes(REL_FOOTNOTES)) {
    const rId = xml.nextRId(relsXml);
    const rel = '<Relationship Id="' + rId + '" Type="' + REL_FOOTNOTES + '" Target="footnotes.xml"/>';
    relsXml = relsXml.replace('</Relationships>', rel + '</Relationships>');
    relsChanged = true;
  }

  if (relsChanged) {
    ws.relsXml = relsXml;
  }

  // Ensure content type exists
  let ctXml = ws.contentTypesXml;
  let ctChanged = false;

  if (ctXml && !ctXml.includes(CT_FOOTNOTES)) {
    ctXml = ctXml.replace('</Types>',
      '<Override PartName="/word/footnotes.xml" ContentType="' + CT_FOOTNOTES + '"/></Types>');
    ctChanged = true;
  }

  if (ctChanged) {
    ws.contentTypesXml = ctXml;
  }
}

module.exports = { Footnotes };
