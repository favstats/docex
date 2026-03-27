/**
 * crossref.js -- Cross-references and auto-numbering for docex
 *
 * Provides label/ref cross-referencing using OOXML SEQ and REF field codes,
 * and auto-numbering for figures and tables.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// CROSSREF
// ============================================================================

class CrossRef {

  /**
   * Assign a label to a paragraph (figure, table, or heading).
   * Stores the label as a w:bookmarkStart/End pair inside the paragraph.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} paraId - The w14:paraId of the paragraph to label
   * @param {string} labelName - Label name, e.g. "fig:funnel", "tab:results"
   */
  static label(ws, paraId, labelName) {
    const docXml = ws.docXml;

    // Find the paragraph by paraId
    const paraIdAttr = `w14:paraId="${paraId}"`;
    const paraStart = docXml.indexOf(paraIdAttr);
    if (paraStart === -1) {
      throw new Error(`Paragraph with paraId "${paraId}" not found`);
    }

    // Find the <w:p start
    const pStart = docXml.lastIndexOf('<w:p', paraStart);
    const pEnd = docXml.indexOf('</w:p>', pStart);
    if (pStart === -1 || pEnd === -1) {
      throw new Error(`Could not locate paragraph boundaries for paraId "${paraId}"`);
    }
    const pEndFull = pEnd + 6;

    const paraXml = docXml.slice(pStart, pEndFull);

    // Sanitize label name for bookmark (OOXML bookmarks can't have colons)
    const bookmarkName = '_docex_' + labelName.replace(/[^a-zA-Z0-9_]/g, '_');

    // Generate a bookmark ID
    const bookmarkId = xml.nextChangeId(docXml);

    // Insert bookmark at the start of the paragraph content (after pPr if present)
    const pPrEnd = paraXml.indexOf('</w:pPr>');
    let insertPos;
    if (pPrEnd !== -1) {
      insertPos = pPrEnd + 8; // after </w:pPr>
    } else {
      // Insert after the <w:p ...> opening tag
      const gtPos = paraXml.indexOf('>');
      insertPos = gtPos + 1;
    }

    const bookmarkXml =
      `<w:bookmarkStart w:id="${bookmarkId}" w:name="${bookmarkName}"/>` +
      `<w:bookmarkEnd w:id="${bookmarkId}"/>`;

    const newParaXml =
      paraXml.slice(0, insertPos) + bookmarkXml + paraXml.slice(insertPos);

    ws.docXml = docXml.slice(0, pStart) + newParaXml + docXml.slice(pEndFull);
  }

  /**
   * Insert a cross-reference at a specified location.
   * Resolves to "Figure 3" or "Table 1" based on the label's context.
   * Uses OOXML REF field code.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} labelName - Label to reference, e.g. "fig:funnel"
   * @param {object} [opts] - Options
   * @param {string} [opts.insertAt] - paraId of the paragraph to insert into
   * @param {string} [opts.after] - Insert after this text within the paragraph
   */
  static ref(ws, labelName, opts = {}) {
    const docXml = ws.docXml;
    const bookmarkName = '_docex_' + labelName.replace(/[^a-zA-Z0-9_]/g, '_');

    // Determine the type from the label prefix
    let prefix = '';
    if (labelName.startsWith('fig:')) prefix = 'Figure ';
    else if (labelName.startsWith('tab:')) prefix = 'Table ';
    else if (labelName.startsWith('sec:')) prefix = 'Section ';

    // Build the REF field code XML
    const fieldXml =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      + `<w:r><w:instrText xml:space="preserve"> REF ${bookmarkName} \\h </w:instrText></w:r>`
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + `<w:r><w:t xml:space="preserve">${xml.escapeXml(prefix)}[#]</w:t></w:r>`
      + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

    if (opts.insertAt) {
      // Insert into a specific paragraph
      const paraIdAttr = `w14:paraId="${opts.insertAt}"`;
      const paraPos = docXml.indexOf(paraIdAttr);
      if (paraPos === -1) {
        throw new Error(`Paragraph with paraId "${opts.insertAt}" not found`);
      }

      const pStart = docXml.lastIndexOf('<w:p', paraPos);
      const pEnd = docXml.indexOf('</w:p>', pStart);
      if (pStart === -1 || pEnd === -1) {
        throw new Error(`Could not locate paragraph boundaries`);
      }

      const paraXml = docXml.slice(pStart, pEnd + 6);

      if (opts.after) {
        // Find the text within the paragraph and insert after it
        const encoded = xml.escapeXml(opts.after);
        const textPos = paraXml.indexOf(encoded);
        if (textPos !== -1) {
          // Find the end of the containing </w:r>
          const runEnd = paraXml.indexOf('</w:r>', textPos);
          if (runEnd !== -1) {
            const insertAt = runEnd + 6;
            const newParaXml = paraXml.slice(0, insertAt) + fieldXml + paraXml.slice(insertAt);
            ws.docXml = docXml.slice(0, pStart) + newParaXml + docXml.slice(pEnd + 6);
            return;
          }
        }
      }

      // Default: insert before </w:p>
      const newParaXml = paraXml.slice(0, -6) + fieldXml + '</w:p>';
      ws.docXml = docXml.slice(0, pStart) + newParaXml + docXml.slice(pEnd + 6);
    } else {
      throw new Error('ref() requires opts.insertAt to specify where to insert the reference');
    }
  }

  /**
   * Scan all figures and tables and ensure captions use SEQ field codes
   * for auto-numbering.
   *
   * Figure captions: { SEQ Figure } -> "1", "2", "3"
   * Table captions:  { SEQ Table }  -> "1", "2", "3"
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {{figures: number, tables: number}} Count of captions processed
   */
  static autoNumber(ws) {
    let docXml = ws.docXml;
    let figureCount = 0;
    let tableCount = 0;

    // Find all paragraphs
    const paragraphs = xml.findParagraphs(docXml);

    // Process in reverse order to preserve offsets
    for (let i = paragraphs.length - 1; i >= 0; i--) {
      const p = paragraphs[i];
      const text = xml.extractTextDecoded(p.xml);

      // Check for figure captions: "Figure N" at start
      const figMatch = text.match(/^(Figure)\s+(\d+)/i);
      if (figMatch) {
        // Check if already has SEQ field
        if (!p.xml.includes('w:instrText') || !p.xml.includes('SEQ Figure')) {
          const newParaXml = CrossRef._injectSeqField(p.xml, 'Figure', figMatch[2]);
          docXml = docXml.slice(0, p.start) + newParaXml + docXml.slice(p.end);
          figureCount++;
        } else {
          figureCount++;
        }
        continue;
      }

      // Check for table captions: "Table N" at start
      const tabMatch = text.match(/^(Table)\s+(\d+)/i);
      if (tabMatch) {
        if (!p.xml.includes('w:instrText') || !p.xml.includes('SEQ Table')) {
          const newParaXml = CrossRef._injectSeqField(p.xml, 'Table', tabMatch[2]);
          docXml = docXml.slice(0, p.start) + newParaXml + docXml.slice(p.end);
          tableCount++;
        } else {
          tableCount++;
        }
        continue;
      }
    }

    ws.docXml = docXml;
    return { figures: figureCount, tables: tableCount };
  }

  /**
   * Return all labels found in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {Array<{name: string, type: string, number: number|null, paraId: string|null}>}
   */
  static listLabels(ws) {
    const docXml = ws.docXml;
    const labels = [];

    // Find all docex bookmarks
    const bmRe = /<w:bookmarkStart\s+w:id="(\d+)"\s+w:name="(_docex_[^"]+)"\s*\/>/g;
    let m;
    while ((m = bmRe.exec(docXml)) !== null) {
      const rawName = m[2];
      // Convert back from bookmark name to label name
      const labelName = rawName.replace(/^_docex_/, '').replace(/_/g, function(match, offset, str) {
        // Heuristic: restore colons for known prefixes
        if (offset === 3 && (str.startsWith('fig_') || str.startsWith('tab_') || str.startsWith('sec_'))) {
          return ':';
        }
        return '_';
      });

      // Determine type from prefix
      let type = 'unknown';
      if (labelName.startsWith('fig:')) type = 'figure';
      else if (labelName.startsWith('tab:')) type = 'table';
      else if (labelName.startsWith('sec:')) type = 'heading';

      // Find the paraId of the containing paragraph
      const pStart = docXml.lastIndexOf('<w:p', m.index);
      let paraId = null;
      if (pStart !== -1) {
        const pidMatch = docXml.slice(pStart, m.index).match(/w14:paraId="([^"]+)"/);
        if (pidMatch) paraId = pidMatch[1];
      }

      labels.push({ name: labelName, type, number: null, paraId });
    }

    return labels;
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Inject a SEQ field code into a caption paragraph, replacing the number.
   *
   * @param {string} paraXml - The paragraph XML
   * @param {string} seqName - "Figure" or "Table"
   * @param {string} currentNumber - The current number text to replace
   * @returns {string} Modified paragraph XML
   * @private
   */
  static _injectSeqField(paraXml, seqName, currentNumber) {
    // Build a SEQ field that displays the current number as its result
    const seqField =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      + `<w:r><w:instrText xml:space="preserve"> SEQ ${seqName} </w:instrText></w:r>`
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + `<w:r><w:t xml:space="preserve">${currentNumber}</w:t></w:r>`
      + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

    // Find and replace the number in the text
    // Pattern: look for the number after seqName + space in a w:t element
    const escapedNum = currentNumber.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const labelRe = new RegExp(
      `(<w:t[^>]*>)(${seqName}\\s+)${escapedNum}([^<]*)(<\\/w:t>)`,
      'i'
    );

    const match = paraXml.match(labelRe);
    if (match) {
      // Replace: keep the label text, replace number with SEQ field
      const replacement =
        `${match[1]}${match[2]}</w:t></w:r>`
        + seqField
        + (match[3] ? `<w:r><w:t xml:space="preserve">${match[3]}</w:t></w:r>` : '');

      // Find the enclosing <w:r and replace the whole run
      const matchPos = paraXml.indexOf(match[0]);
      const rStart = paraXml.lastIndexOf('<w:r', matchPos);
      const rEnd = paraXml.indexOf('</w:r>', matchPos);

      if (rStart !== -1 && rEnd !== -1) {
        return paraXml.slice(0, rStart) + replacement + paraXml.slice(rEnd + 6);
      }
    }

    // Fallback: return as-is if we can't find the exact pattern
    return paraXml;
  }
}

module.exports = { CrossRef };
