/**
 * lists.js -- Bullet and numbered list operations for docex
 *
 * Creates properly formatted OOXML lists with numbering definitions.
 * Manages word/numbering.xml for list style definitions.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const xml = require('./xml');

// ============================================================================
// EMPTY NUMBERING TEMPLATE
// ============================================================================

const EMPTY_NUMBERING_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  + '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  + ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
  + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
  + ' mc:Ignorable="w14">'
  + '</w:numbering>';

// ============================================================================
// LISTS
// ============================================================================

class Lists {

  /**
   * Insert a bullet list at a position in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} anchor - Text to position relative to
   * @param {string} mode - 'after' or 'before'
   * @param {string[]} items - Array of list item texts
   * @param {object} [opts] - Options
   * @param {string} [opts.author] - Author for tracked changes
   * @param {boolean} [opts.tracked] - Whether to use tracked changes
   * @param {string} [opts.date] - ISO date for tracked changes
   */
  static insertBulletList(ws, anchor, mode, items, opts = {}) {
    if (!items || items.length === 0) {
      throw new Error('List items must be a non-empty array');
    }

    // Ensure numbering.xml exists and get/create a bullet list definition
    const numId = Lists._ensureBulletNumbering(ws);

    // Build list paragraphs
    const listXml = items.map(item => {
      return Lists._buildListParagraph(item, numId, 0, opts);
    }).join('');

    // Insert at position
    Lists._insertAtPosition(ws, anchor, mode, listXml);
  }

  /**
   * Insert a numbered list at a position in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} anchor - Text to position relative to
   * @param {string} mode - 'after' or 'before'
   * @param {string[]} items - Array of list item texts
   * @param {object} [opts] - Options
   */
  static insertNumberedList(ws, anchor, mode, items, opts = {}) {
    if (!items || items.length === 0) {
      throw new Error('List items must be a non-empty array');
    }

    const numId = Lists._ensureNumberedNumbering(ws);

    const listXml = items.map(item => {
      return Lists._buildListParagraph(item, numId, 0, opts);
    }).join('');

    Lists._insertAtPosition(ws, anchor, mode, listXml);
  }

  /**
   * Insert a nested list (bulleted) at a position in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} anchor - Text to position relative to
   * @param {string} mode - 'after' or 'before'
   * @param {Array<{text: string, children?: Array}>} tree - Nested items
   * @param {object} [opts] - Options
   */
  static insertNestedList(ws, anchor, mode, tree, opts = {}) {
    if (!tree || tree.length === 0) {
      throw new Error('List tree must be a non-empty array');
    }

    const numId = Lists._ensureBulletNumbering(ws);
    const listXml = Lists._flattenTree(tree, numId, 0, opts);

    Lists._insertAtPosition(ws, anchor, mode, listXml);
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Flatten a nested tree into sequential paragraphs with increasing indent levels.
   *
   * @param {Array<{text: string, children?: Array}>} nodes - Tree nodes
   * @param {number} numId - Numbering definition ID
   * @param {number} level - Current indent level (0-based)
   * @param {object} opts - Options
   * @returns {string} Concatenated paragraph XML
   * @private
   */
  static _flattenTree(nodes, numId, level, opts) {
    let result = '';
    for (const node of nodes) {
      result += Lists._buildListParagraph(node.text, numId, level, opts);
      if (node.children && node.children.length > 0) {
        result += Lists._flattenTree(node.children, numId, level + 1, opts);
      }
    }
    return result;
  }

  /**
   * Build a single list item paragraph.
   *
   * @param {string} text - Item text
   * @param {number} numId - Numbering definition ID
   * @param {number} level - Indent level (0-based)
   * @param {object} opts - Options
   * @returns {string} Paragraph XML
   * @private
   */
  static _buildListParagraph(text, numId, level, opts = {}) {
    const paraId = xml.randomHexId();
    const textId = xml.randomHexId();
    const escapedText = xml.escapeXml(text);
    const tracked = opts.tracked || false;
    const author = opts.author || 'Unknown';
    const date = opts.date || xml.isoNow();

    // Indentation in twips: 720 per level (0.5 inches)
    const leftIndent = 720 * (level + 1);
    const hangingIndent = 360;

    const pPr =
      '<w:pPr>'
      + `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`
      + `<w:ind w:left="${leftIndent}" w:hanging="${hangingIndent}"/>`
      + '</w:pPr>';

    const runContent = `<w:t xml:space="preserve">${escapedText}</w:t>`;

    if (tracked) {
      const changeId = xml.nextChangeId(opts._docXml || '');
      return `<w:p w14:paraId="${paraId}" w14:textId="${textId}">`
        + pPr
        + `<w:ins w:id="${changeId}" w:author="${xml.escapeXml(author)}" w:date="${date}">`
        + `<w:r>${runContent}</w:r>`
        + '</w:ins>'
        + '</w:p>';
    }

    return `<w:p w14:paraId="${paraId}" w14:textId="${textId}">`
      + pPr
      + `<w:r>${runContent}</w:r>`
      + '</w:p>';
  }

  /**
   * Insert XML at a position relative to an anchor paragraph.
   *
   * @param {object} ws - Workspace
   * @param {string} anchor - Text to find
   * @param {string} mode - 'after' or 'before'
   * @param {string} insertXml - XML to insert
   * @private
   */
  static _insertAtPosition(ws, anchor, mode, insertXml) {
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);

    // Find the anchor paragraph
    let targetIdx = -1;
    for (let i = 0; i < paragraphs.length; i++) {
      const decoded = xml.extractTextDecoded(paragraphs[i].xml);
      if (decoded.includes(anchor)) {
        targetIdx = i;
        break;
      }
    }

    if (targetIdx === -1) {
      throw new Error(`Anchor text not found: "${anchor.slice(0, 60)}"`);
    }

    const target = paragraphs[targetIdx];

    if (mode === 'after') {
      ws.docXml = docXml.slice(0, target.end) + insertXml + docXml.slice(target.end);
    } else {
      ws.docXml = docXml.slice(0, target.start) + insertXml + docXml.slice(target.start);
    }
  }

  /**
   * Ensure numbering.xml exists and contains a bullet list definition.
   * Returns the numId for the bullet list.
   *
   * @param {object} ws - Workspace
   * @returns {number} The numId to use in paragraphs
   * @private
   */
  static _ensureBulletNumbering(ws) {
    let numberingXml = Lists._getNumberingXml(ws);

    // Check if we already have a docex bullet definition
    if (numberingXml.includes('w:abstractNumId="1000"')) {
      return 1000;
    }

    // Add abstract numbering definition for bullets
    const abstractNum =
      '<w:abstractNum w:abstractNumId="1000">'
      + '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
      + '<w:numFmt w:val="bullet"/>'
      + '<w:lvlText w:val="\u2022"/>'  // bullet character
      + '<w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
      + '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
      + '</w:lvl>'
      + '<w:lvl w:ilvl="1"><w:start w:val="1"/>'
      + '<w:numFmt w:val="bullet"/>'
      + '<w:lvlText w:val="\u25E6"/>'  // white bullet
      + '<w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
      + '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr>'
      + '</w:lvl>'
      + '<w:lvl w:ilvl="2"><w:start w:val="1"/>'
      + '<w:numFmt w:val="bullet"/>'
      + '<w:lvlText w:val="\u25AA"/>'  // black small square
      + '<w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>'
      + '</w:lvl>'
      + '</w:abstractNum>';

    const numDef = '<w:num w:numId="1000"><w:abstractNumId w:val="1000"/></w:num>';

    // Insert before closing tag
    numberingXml = numberingXml.replace('</w:numbering>', abstractNum + numDef + '</w:numbering>');

    Lists._setNumberingXml(ws, numberingXml);
    return 1000;
  }

  /**
   * Ensure numbering.xml exists and contains a numbered list definition.
   * Returns the numId for the numbered list.
   *
   * @param {object} ws - Workspace
   * @returns {number} The numId to use in paragraphs
   * @private
   */
  static _ensureNumberedNumbering(ws) {
    let numberingXml = Lists._getNumberingXml(ws);

    if (numberingXml.includes('w:abstractNumId="1001"')) {
      return 1001;
    }

    const abstractNum =
      '<w:abstractNum w:abstractNumId="1001">'
      + '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
      + '<w:numFmt w:val="decimal"/>'
      + '<w:lvlText w:val="%1."/>'
      + '<w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
      + '</w:lvl>'
      + '<w:lvl w:ilvl="1"><w:start w:val="1"/>'
      + '<w:numFmt w:val="lowerLetter"/>'
      + '<w:lvlText w:val="%2."/>'
      + '<w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
      + '</w:lvl>'
      + '<w:lvl w:ilvl="2"><w:start w:val="1"/>'
      + '<w:numFmt w:val="lowerRoman"/>'
      + '<w:lvlText w:val="%3."/>'
      + '<w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>'
      + '</w:lvl>'
      + '</w:abstractNum>';

    const numDef = '<w:num w:numId="1001"><w:abstractNumId w:val="1001"/></w:num>';

    numberingXml = numberingXml.replace('</w:numbering>', abstractNum + numDef + '</w:numbering>');

    Lists._setNumberingXml(ws, numberingXml);
    return 1001;
  }

  /**
   * Get numbering.xml content from workspace, creating if missing.
   *
   * @param {object} ws - Workspace
   * @returns {string} numbering.xml content
   * @private
   */
  static _getNumberingXml(ws) {
    if (ws._numberingXml !== undefined && ws._numberingXml !== null) {
      return ws._numberingXml;
    }

    const filePath = path.join(ws.tmpDir, 'word', 'numbering.xml');
    if (fs.existsSync(filePath)) {
      ws._numberingXml = fs.readFileSync(filePath, 'utf-8');
    } else {
      ws._numberingXml = EMPTY_NUMBERING_XML;
    }
    return ws._numberingXml;
  }

  /**
   * Set numbering.xml content in the workspace and write to disk.
   *
   * @param {object} ws - Workspace
   * @param {string} content - New numbering.xml content
   * @private
   */
  static _setNumberingXml(ws, content) {
    ws._numberingXml = content;
    const filePath = path.join(ws.tmpDir, 'word', 'numbering.xml');
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(filePath, content, 'utf-8');

    // Ensure numbering.xml is referenced in relationships
    Lists._ensureNumberingRelationship(ws);

    // Ensure content type is registered
    Lists._ensureNumberingContentType(ws);
  }

  /**
   * Ensure document.xml.rels references numbering.xml.
   *
   * @param {object} ws - Workspace
   * @private
   */
  static _ensureNumberingRelationship(ws) {
    const relsXml = ws.relsXml;
    if (relsXml.includes('numbering.xml')) return;

    const rId = xml.nextRId(relsXml);
    const rel = `<Relationship Id="${rId}" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
      + 'Target="numbering.xml"/>';

    ws.relsXml = relsXml.replace('</Relationships>', rel + '</Relationships>');
  }

  /**
   * Ensure [Content_Types].xml includes numbering.xml.
   *
   * @param {object} ws - Workspace
   * @private
   */
  static _ensureNumberingContentType(ws) {
    const ctXml = ws.contentTypesXml;
    if (ctXml.includes('numbering.xml')) return;

    const override = '<Override PartName="/word/numbering.xml" '
      + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>';

    ws.contentTypesXml = ctXml.replace('</Types>', override + '</Types>');
  }
}

module.exports = { Lists };
