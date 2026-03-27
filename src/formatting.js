/**
 * formatting.js -- Inline formatting operations for docex
 *
 * Static methods for applying character-level formatting (bold, italic,
 * underline, highlight, color, code, etc.) to text within OOXML
 * document.xml. Supports tracked formatting changes via w:rPrChange.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// COLOR MAPS
// ============================================================================

/**
 * Named colors mapped to hex values (RRGGBB, no leading #).
 * @type {Object<string, string>}
 */
const COLORS = {
  red:     'FF0000',
  blue:    '0000FF',
  green:   '008000',
  yellow:  'FFFF00',
  orange:  'FF8C00',
  purple:  '800080',
  cyan:    '00FFFF',
  magenta: 'FF00FF',
  black:   '000000',
  gray:    '808080',
};

/**
 * Word's built-in highlight color names.
 * @type {string[]}
 */
const HIGHLIGHTS = [
  'yellow', 'green', 'cyan', 'magenta', 'blue', 'red',
  'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta',
  'darkRed', 'darkYellow', 'lightGray', 'darkGray',
];

// ============================================================================
// FORMATTING CLASS
// ============================================================================

class Formatting {

  // --------------------------------------------------------------------------
  // Public API
  // --------------------------------------------------------------------------

  /**
   * Apply bold formatting to text in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to make bold
   * @param {object} [opts] - Options
   * @param {boolean} [opts.tracked=false] - Wrap change in w:rPrChange
   * @param {string} [opts.author='Unknown'] - Author for tracked changes
   * @param {string} [opts.date] - ISO date for tracked changes
   * @throws {Error} If text is not found
   */
  static bold(ws, text, opts = {}) {
    Formatting._applyFormat(ws, text, '<w:b/>', opts);
  }

  /**
   * Apply italic formatting to text in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to make italic
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found
   */
  static italic(ws, text, opts = {}) {
    Formatting._applyFormat(ws, text, '<w:i/>', opts);
  }

  /**
   * Apply underline formatting to text in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to underline
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found
   */
  static underline(ws, text, opts = {}) {
    Formatting._applyFormat(ws, text, '<w:u w:val="single"/>', opts);
  }

  /**
   * Apply strikethrough formatting to text in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to strike through
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found
   */
  static strikethrough(ws, text, opts = {}) {
    Formatting._applyFormat(ws, text, '<w:strike/>', opts);
  }

  /**
   * Apply superscript formatting to text in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to make superscript
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found
   */
  static superscript(ws, text, opts = {}) {
    Formatting._applyFormat(ws, text, '<w:vertAlign w:val="superscript"/>', opts);
  }

  /**
   * Apply subscript formatting to text in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to make subscript
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found
   */
  static subscript(ws, text, opts = {}) {
    Formatting._applyFormat(ws, text, '<w:vertAlign w:val="subscript"/>', opts);
  }

  /**
   * Apply small caps formatting to text in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to make small caps
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found
   */
  static smallCaps(ws, text, opts = {}) {
    Formatting._applyFormat(ws, text, '<w:smallCaps/>', opts);
  }

  /**
   * Apply code/monospace formatting to text in the document.
   * Sets the font to Courier New.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to make monospace
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found
   */
  static code(ws, text, opts = {}) {
    Formatting._applyFormat(
      ws, text,
      '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>',
      opts
    );
  }

  /**
   * Apply font color to text in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to color
   * @param {string} colorName - Named color (from COLORS) or 6-char hex
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found or color is invalid
   */
  static color(ws, text, colorName, opts = {}) {
    const hex = COLORS[colorName] || colorName;
    if (!/^[0-9A-Fa-f]{6}$/.test(hex)) {
      throw new Error(`Invalid color: "${colorName}". Use a named color or 6-digit hex.`);
    }
    Formatting._applyFormat(ws, text, `<w:color w:val="${hex.toUpperCase()}"/>`, opts);
  }

  /**
   * Apply highlight to text in the document.
   * Uses Word's built-in highlight colors.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to highlight
   * @param {string} colorName - Highlight color name (from HIGHLIGHTS)
   * @param {object} [opts] - Options (same as bold)
   * @throws {Error} If text is not found or color is invalid
   */
  static highlight(ws, text, colorName, opts = {}) {
    if (!HIGHLIGHTS.includes(colorName)) {
      throw new Error(
        `Invalid highlight color: "${colorName}". Valid: ${HIGHLIGHTS.join(', ')}`
      );
    }
    Formatting._applyFormat(ws, text, `<w:highlight w:val="${colorName}"/>`, opts);
  }

  // --------------------------------------------------------------------------
  // Static properties
  // --------------------------------------------------------------------------

  /** Named color map */
  static get COLORS() { return COLORS; }

  /** Valid highlight color names */
  static get HIGHLIGHTS() { return HIGHLIGHTS; }

  // --------------------------------------------------------------------------
  // Internal: core formatting engine
  // --------------------------------------------------------------------------

  /**
   * Apply a formatting element to text in the document.
   *
   * Algorithm:
   *   1. Find the paragraph containing the text
   *   2. Find which run(s) contain the text (may span runs)
   *   3. If text is within a single run: split the run, apply formatting
   *   4. If text spans runs: split boundary runs, format middle runs
   *   5. If tracked: wrap old rPr in w:rPrChange
   *
   * @param {object} ws - Workspace
   * @param {string} text - Text to format
   * @param {string} formatElement - XML element to add to rPr (e.g. '<w:b/>')
   * @param {object} opts - Options
   * @private
   */
  static _applyFormat(ws, text, formatElement, opts) {
    const tracked = opts.tracked || false;
    const author = opts.author || 'Unknown';
    const date = opts.date || xml.isoNow();
    let docXml = ws.docXml;

    const paragraphs = xml.findParagraphs(docXml);
    let found = false;

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      const decodedText = xml.decodeXml(para.text);
      if (!decodedText.includes(text)) continue;

      const result = Formatting._formatInParagraph(
        para.xml, text, formatElement, tracked, author, date, docXml
      );
      if (result.modified) {
        docXml = docXml.slice(0, para.start) + result.xml + docXml.slice(para.end);
        found = true;
        break;
      }
    }

    if (!found) {
      throw new Error('Text not found for formatting: "' + text.slice(0, 80) + '"');
    }

    ws.docXml = docXml;
  }

  /**
   * Apply formatting to text within a single paragraph.
   *
   * @param {string} pXml - Paragraph XML
   * @param {string} searchText - Text to format
   * @param {string} formatElement - XML element to insert into rPr
   * @param {boolean} tracked - Whether to use tracked formatting change
   * @param {string} author - Author for tracked changes
   * @param {string} date - ISO date
   * @param {string} docXml - Full document XML (for nextChangeId)
   * @returns {{modified: boolean, xml: string}}
   * @private
   */
  static _formatInParagraph(pXml, searchText, formatElement, tracked, author, date, docXml) {
    // Use the same cross-run finding logic as Paragraphs
    const allRuns = xml.parseRuns(pXml);
    const textRuns = allRuns.filter(r => r.texts.length > 0);
    if (textRuns.length === 0) return { modified: false, xml: pXml };

    // Build decoded combined text for searching
    const decodedTexts = textRuns.map(r => xml.decodeXml(r.combinedText));
    const combined = decodedTexts.join('');
    const matchPos = combined.indexOf(searchText);
    if (matchPos === -1) return { modified: false, xml: pXml };

    // Map match position to runs
    let charOffset = 0;
    let matchStartRun = -1;
    let matchEndRun = -1;
    let matchStartOffset = -1;
    let matchEndOffset = -1;

    for (let i = 0; i < textRuns.length; i++) {
      const runLen = decodedTexts[i].length;
      const runStart = charOffset;
      const runEnd = charOffset + runLen;

      if (matchStartRun === -1 && matchPos < runEnd) {
        matchStartRun = i;
        matchStartOffset = matchPos - runStart;
      }
      if (matchStartRun !== -1 && matchPos + searchText.length <= runEnd) {
        matchEndRun = i;
        matchEndOffset = matchPos + searchText.length - runStart;
        break;
      }
      charOffset += runLen;
    }

    if (matchStartRun === -1 || matchEndRun === -1) return { modified: false, xml: pXml };

    // Build replacement runs
    const newRuns = [];

    for (let i = matchStartRun; i <= matchEndRun; i++) {
      const run = textRuns[i];
      const runText = decodedTexts[i];
      const runLen = runText.length;

      let sliceStart = 0;
      let sliceEnd = runLen;
      if (i === matchStartRun) sliceStart = matchStartOffset;
      if (i === matchEndRun) sliceEnd = matchEndOffset;

      // Prefix: text before the match in the start run (keep original formatting)
      if (i === matchStartRun && matchStartOffset > 0) {
        const prefixText = runText.slice(0, matchStartOffset);
        newRuns.push(Formatting._buildRun(run.rPr, prefixText));
      }

      // Matched text: apply new formatting
      const matchedText = runText.slice(sliceStart, sliceEnd);
      if (matchedText.length > 0) {
        const newRPr = Formatting._addToRPr(run.rPr, formatElement, tracked, author, date, docXml);
        newRuns.push(Formatting._buildRun(newRPr, matchedText));
      }

      // Suffix: text after the match in the end run (keep original formatting)
      if (i === matchEndRun && matchEndOffset < runLen) {
        const suffixText = runText.slice(matchEndOffset);
        newRuns.push(Formatting._buildRun(run.rPr, suffixText));
      }
    }

    // Splice the new runs into the paragraph XML
    const startPos = textRuns[matchStartRun].index;
    const lastRun = textRuns[matchEndRun];
    const endPos = lastRun.index + lastRun.fullMatch.length;
    const newPXml = pXml.slice(0, startPos) + newRuns.join('') + pXml.slice(endPos);

    return { modified: true, xml: newPXml };
  }

  /**
   * Build a <w:r> element with given rPr and text.
   *
   * @param {string} rPr - Run properties XML (full <w:rPr>...</w:rPr> or empty string)
   * @param {string} text - Decoded text content (will be XML-escaped)
   * @returns {string} Complete <w:r> XML fragment
   * @private
   */
  static _buildRun(rPr, text) {
    return '<w:r>' + rPr + '<w:t xml:space="preserve">' + xml.escapeXml(text) + '</w:t></w:r>';
  }

  /**
   * Add a formatting element to an rPr block.
   *
   * If the run already has <w:rPr>, inserts the new element inside it.
   * If no rPr exists, creates one.
   * If tracked, wraps the original rPr content in w:rPrChange.
   *
   * @param {string} rPr - Existing rPr XML (may be empty string)
   * @param {string} formatElement - XML element to add (e.g. '<w:b/>')
   * @param {boolean} tracked - Whether to use tracked formatting change
   * @param {string} author - Author name
   * @param {string} date - ISO date
   * @param {string} docXml - Full document XML (for nextChangeId)
   * @returns {string} New rPr XML
   * @private
   */
  static _addToRPr(rPr, formatElement, tracked, author, date, docXml) {
    if (!rPr || rPr.trim() === '') {
      // No existing rPr: create one
      if (tracked) {
        const id = xml.nextChangeId(docXml);
        return '<w:rPr>'
          + formatElement
          + '<w:rPrChange w:id="' + id + '" w:author="' + xml.escapeXml(author) + '" w:date="' + date + '">'
          + '<w:rPr/>'
          + '</w:rPrChange>'
          + '</w:rPr>';
      }
      return '<w:rPr>' + formatElement + '</w:rPr>';
    }

    // Has existing rPr: extract the inner content
    const innerMatch = rPr.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    if (!innerMatch) {
      // Self-closing rPr or unexpected format, create new
      if (tracked) {
        const id = xml.nextChangeId(docXml);
        return '<w:rPr>'
          + formatElement
          + '<w:rPrChange w:id="' + id + '" w:author="' + xml.escapeXml(author) + '" w:date="' + date + '">'
          + '<w:rPr/>'
          + '</w:rPrChange>'
          + '</w:rPr>';
      }
      return '<w:rPr>' + formatElement + '</w:rPr>';
    }

    const innerContent = innerMatch[1];

    if (tracked) {
      const id = xml.nextChangeId(docXml);
      // Original rPr content goes into w:rPrChange; new format element is added to current rPr
      return '<w:rPr>'
        + innerContent
        + formatElement
        + '<w:rPrChange w:id="' + id + '" w:author="' + xml.escapeXml(author) + '" w:date="' + date + '">'
        + '<w:rPr>' + innerContent + '</w:rPr>'
        + '</w:rPrChange>'
        + '</w:rPr>';
    }

    // Untracked: add the format element inside the existing rPr
    return '<w:rPr>' + innerContent + formatElement + '</w:rPr>';
  }
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = { Formatting };
