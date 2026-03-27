/**
 * layout.js -- Layout and structural manipulation for docex (v0.4.8)
 *
 * Static methods for page breaks, heading hierarchy fixes,
 * paragraph merging/splitting, and table replacement.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');
const { Paragraphs } = require('./paragraphs');
const { Tables } = require('./tables');

class Layout {

  /**
   * Add a page break before the paragraph that contains the given heading text.
   * Inserts <w:pageBreakBefore/> into the paragraph's pPr element.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} headingText - Text of the heading to find
   * @throws {Error} If heading is not found
   */
  static pageBreakBefore(ws, headingText) {
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);

    // Find the paragraph containing the heading text
    let target = null;
    for (const p of paragraphs) {
      const decoded = xml.decodeXml(p.text);
      if (decoded === headingText || decoded.includes(headingText)) {
        target = p;
        break;
      }
    }

    // Case-insensitive fallback
    if (!target) {
      const lower = headingText.toLowerCase();
      for (const p of paragraphs) {
        const decoded = xml.decodeXml(p.text).toLowerCase();
        if (decoded.includes(lower)) {
          target = p;
          break;
        }
      }
    }

    if (!target) {
      throw new Error('Heading not found: "' + headingText + '"');
    }

    // Check if it already has pageBreakBefore
    if (target.xml.includes('<w:pageBreakBefore') || target.xml.includes('pageBreakBefore')) {
      return; // Already has page break
    }

    // Insert pageBreakBefore into pPr
    const pPrMatch = target.xml.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/);
    const pPrSelfClose = target.xml.match(/<w:pPr\s*\/>/);

    let newParaXml;
    if (pPrMatch) {
      // Has existing pPr -- add pageBreakBefore inside it
      const newPPr = '<w:pPr>' + pPrMatch[1] + '<w:pageBreakBefore/>' + '</w:pPr>';
      newParaXml = target.xml.replace(pPrMatch[0], newPPr);
    } else if (pPrSelfClose) {
      // Self-closing pPr -- expand it
      newParaXml = target.xml.replace(pPrSelfClose[0], '<w:pPr><w:pageBreakBefore/></w:pPr>');
    } else {
      // No pPr at all -- add one at the start of the paragraph
      newParaXml = target.xml.replace('<w:p>', '<w:p><w:pPr><w:pageBreakBefore/></w:pPr>');
      // If the paragraph has attributes
      if (!target.xml.startsWith('<w:p>')) {
        const openTag = target.xml.match(/^<w:p[^>]*>/);
        if (openTag) {
          newParaXml = target.xml.replace(openTag[0], openTag[0] + '<w:pPr><w:pageBreakBefore/></w:pPr>');
        }
      }
    }

    ws.docXml = docXml.slice(0, target.start) + newParaXml + docXml.slice(target.end);
  }

  /**
   * Scan headings and fix hierarchy skips (e.g. H1 -> H3 becomes H1 -> H2).
   * Returns the count of fixes applied.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {number} Number of headings adjusted
   */
  static ensureHeadingHierarchy(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);
    const headings = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const level = Paragraphs._headingLevel(paragraphs[i].xml);
      if (level > 0) {
        headings.push({ index: i, level, para: paragraphs[i] });
      }
    }

    if (headings.length === 0) return 0;

    // Determine which headings need fixing
    // Rule: each heading's level should be at most 1 more than the previous heading's level
    let fixes = 0;
    const adjustments = []; // [{paraIndex, oldLevel, newLevel}]

    let prevLevel = headings[0].level;
    for (let i = 1; i < headings.length; i++) {
      const h = headings[i];
      if (h.level > prevLevel + 1) {
        // Skip detected: adjust this heading to prevLevel + 1
        const newLevel = prevLevel + 1;
        adjustments.push({ paraIndex: h.index, oldLevel: h.level, newLevel, para: h.para });
        prevLevel = newLevel;
        fixes++;
      } else {
        prevLevel = h.level;
      }
    }

    if (adjustments.length === 0) return 0;

    // Apply adjustments in reverse order (to preserve positions)
    let docXml = ws.docXml;
    for (let i = adjustments.length - 1; i >= 0; i--) {
      const adj = adjustments[i];
      const para = adj.para;
      let newXml = para.xml;

      // Replace the style ID
      const styleMatch = newXml.match(/<w:pStyle\s+w:val="([^"]+)"/);
      if (styleMatch) {
        const oldStyleId = styleMatch[1];
        let newStyleId;

        // Determine new style ID based on the naming pattern
        const namedMatch = oldStyleId.match(/^([Hh]eading)(\d+)$/);
        if (namedMatch) {
          newStyleId = namedMatch[1] + adj.newLevel;
        } else {
          // For numeric IDs or unknown patterns, use standard naming
          newStyleId = 'Heading' + adj.newLevel;
        }

        newXml = newXml.replace(
          '<w:pStyle w:val="' + oldStyleId + '"',
          '<w:pStyle w:val="' + newStyleId + '"'
        );
      }

      // Replace outlineLvl if present (0-based)
      const olvlMatch = newXml.match(/<w:outlineLvl\s+w:val="(\d+)"/);
      if (olvlMatch) {
        newXml = newXml.replace(
          '<w:outlineLvl w:val="' + olvlMatch[1] + '"',
          '<w:outlineLvl w:val="' + (adj.newLevel - 1) + '"'
        );
      }

      docXml = docXml.slice(0, para.start) + newXml + docXml.slice(para.end);
    }

    ws.docXml = docXml;
    return fixes;
  }

  /**
   * Merge two consecutive paragraphs into one.
   * Appends all runs from the second paragraph to the first, then removes the second.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} id1 - w14:paraId of the first paragraph
   * @param {string} id2 - w14:paraId of the second paragraph
   * @throws {Error} If either paragraph is not found
   */
  static mergeParagraphs(ws, id1, id2) {
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);

    let p1 = null, p2 = null;

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      if (Layout._getParaId(p.xml) === id1) {
        p1 = p;
      } else if (Layout._getParaId(p.xml) === id2) {
        p2 = p;
      }
    }

    if (!p1) throw new Error('Paragraph not found: ' + id1);
    if (!p2) throw new Error('Paragraph not found: ' + id2);

    // Extract runs from p2
    const runs2 = Layout._extractRuns(p2.xml);

    // Insert p2's runs into p1, before </w:p>
    const closeTag = '</w:p>';
    const p1CloseIdx = p1.xml.lastIndexOf(closeTag);
    const newP1Xml = p1.xml.slice(0, p1CloseIdx) + runs2 + closeTag;

    // Replace p1 in document, remove p2
    let result;
    if (p1.start < p2.start) {
      // p1 comes first: replace p1, then remove p2
      result = docXml.slice(0, p1.start)
        + newP1Xml
        + docXml.slice(p1.end, p2.start)
        + docXml.slice(p2.end);
    } else {
      // p2 comes first: remove p2, then replace p1
      result = docXml.slice(0, p2.start)
        + docXml.slice(p2.end, p1.start)
        + newP1Xml
        + docXml.slice(p1.end);
    }

    ws.docXml = result;
  }

  /**
   * Split a paragraph at the given text into two paragraphs.
   * The first paragraph gets content before atText, the second gets atText and after.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} paraId - w14:paraId of the paragraph to split
   * @param {string} atText - Text at which to split
   * @returns {string} New paraId for the second paragraph
   * @throws {Error} If paragraph or text is not found
   */
  static splitParagraph(ws, paraId, atText) {
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);

    let target = null;
    for (const p of paragraphs) {
      if (Layout._getParaId(p.xml) === paraId) {
        target = p;
        break;
      }
    }

    if (!target) throw new Error('Paragraph not found: ' + paraId);

    // Get the full text and find the split point
    const fullText = xml.extractTextDecoded(target.xml);
    const splitPos = fullText.indexOf(atText);
    if (splitPos === -1) {
      throw new Error('Text not found in paragraph: "' + atText + '"');
    }

    const textBefore = fullText.slice(0, splitPos);
    const textAfter = fullText.slice(splitPos);

    // Extract pPr from the original paragraph
    const pPr = Paragraphs._extractPpr(target.xml);

    // Generate a new paraId for the second paragraph
    const newParaId = xml.randomHexId();

    // Build first paragraph (content before split point)
    const p1Xml = Layout._buildSimpleParagraph(pPr, textBefore, paraId);

    // Build second paragraph (atText and after)
    const p2Xml = Layout._buildSimpleParagraph(pPr, textAfter, newParaId);

    // Replace original paragraph with two new ones
    ws.docXml = docXml.slice(0, target.start) + p1Xml + p2Xml + docXml.slice(target.end);

    return newParaId;
  }

  /**
   * Replace the content of the nth table in the document with new data.
   * Preserves table structure (borders, column widths, style).
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {number} tableNumber - 1-based table number
   * @param {Array<Array<string>>} newData - 2D array of cell values
   * @param {object} [opts] - Options { headers: bool, style: string }
   */
  static replaceTable(ws, tableNumber, newData, opts = {}) {
    if (!newData || newData.length === 0) {
      throw new Error('Table data must be a non-empty 2D array');
    }

    const docXml = ws.docXml;

    // Find all <w:tbl> elements
    const tblRe = /<w:tbl[\s>]/g;
    let match;
    let count = 0;
    let tblStart = -1;

    while ((match = tblRe.exec(docXml)) !== null) {
      count++;
      if (count === tableNumber) {
        tblStart = match.index;
        break;
      }
    }

    if (tblStart === -1) {
      throw new Error('Table ' + tableNumber + ' not found (document has ' + count + ' tables)');
    }

    // Find the closing </w:tbl>
    const closeTag = '</w:tbl>';
    const tblEnd = docXml.indexOf(closeTag, tblStart);
    if (tblEnd === -1) {
      throw new Error('Malformed table XML: missing </w:tbl>');
    }
    const tblEndFull = tblEnd + closeTag.length;

    // Extract existing table to preserve tblPr
    const existingTblXml = docXml.slice(tblStart, tblEndFull);

    // Extract tblPr
    const tblPrMatch = existingTblXml.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/);
    const existingTblPr = tblPrMatch ? tblPrMatch[0] : null;

    // Extract tblGrid
    const tblGridMatch = existingTblXml.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/);
    const existingTblGrid = tblGridMatch ? tblGridMatch[0] : null;

    // Determine style
    const style = opts.style || 'booktabs';
    const headers = opts.headers !== false;

    const numCols = Math.max(...newData.map(row => row.length));
    const colWidth = Math.floor(9360 / numCols);

    // Rebuild tblGrid if column count changed
    let tblGrid;
    if (existingTblGrid) {
      const existingCols = (existingTblGrid.match(/<w:gridCol/g) || []).length;
      if (existingCols === numCols) {
        tblGrid = existingTblGrid;
      } else {
        tblGrid = '<w:tblGrid>';
        for (let c = 0; c < numCols; c++) {
          tblGrid += '<w:gridCol w:w="' + colWidth + '"/>';
        }
        tblGrid += '</w:tblGrid>';
      }
    } else {
      tblGrid = '<w:tblGrid>';
      for (let c = 0; c < numCols; c++) {
        tblGrid += '<w:gridCol w:w="' + colWidth + '"/>';
      }
      tblGrid += '</w:tblGrid>';
    }

    // Build rows
    let rowsXml = '';
    for (let r = 0; r < newData.length; r++) {
      const row = newData[r];
      const isFirstRow = (r === 0);
      const isLastRow = (r === newData.length - 1);
      const isHeaderRow = (isFirstRow && headers);

      rowsXml += Tables._buildRowXml(
        row, numCols, colWidth, style, { isFirstRow, isLastRow, isHeaderRow }
      );
    }

    // Assemble new table
    const tblPr = existingTblPr || Tables._buildTableXml(newData, { headers, style }).match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/)[0];
    const newTblXml = '<w:tbl>' + tblPr + tblGrid + rowsXml + '</w:tbl>';

    ws.docXml = docXml.slice(0, tblStart) + newTblXml + docXml.slice(tblEndFull);
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Extract w14:paraId from paragraph XML.
   * @param {string} pXml
   * @returns {string|null}
   * @private
   */
  static _getParaId(pXml) {
    const m = pXml.match(/w14:paraId="([^"]+)"/);
    return m ? m[1] : null;
  }

  /**
   * Extract all w:r elements from paragraph XML (everything except pPr).
   * @param {string} pXml
   * @returns {string} Concatenated run XML
   * @private
   */
  static _extractRuns(pXml) {
    const runs = [];
    const runRe = /<w:r[\s>][\s\S]*?<\/w:r>/g;
    let m;
    while ((m = runRe.exec(pXml)) !== null) {
      runs.push(m[0]);
    }
    return runs.join('');
  }

  /**
   * Build a simple paragraph with given pPr, text, and optional paraId.
   * @param {string} pPr - Paragraph properties XML
   * @param {string} text - Text content
   * @param {string} [paraId] - w14:paraId value
   * @returns {string} Complete paragraph XML
   * @private
   */
  static _buildSimpleParagraph(pPr, text, paraId) {
    const paraIdAttr = paraId ? ' w14:paraId="' + paraId + '"' : '';
    const escapedText = xml.escapeXml(text);

    return '<w:p' + paraIdAttr + '>'
      + pPr
      + '<w:r>'
      + '<w:rPr>'
      + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
      + '<w:sz w:val="24"/>'
      + '</w:rPr>'
      + '<w:t xml:space="preserve">' + escapedText + '</w:t>'
      + '</w:r>'
      + '</w:p>';
  }
}

module.exports = { Layout };
