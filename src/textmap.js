/**
 * textmap.js -- Cross-boundary text finding for OOXML paragraphs.
 *
 * Solves the run-splitting problem: Word can split "Hello World" across
 * multiple <w:r> elements. TextMap concatenates all visible text and maps
 * each character back to its source <w:t> element so that search results
 * can be translated into precise XML offsets.
 *
 * Zero external dependencies. Regex-based, not DOM.
 */

'use strict';

// ============================================================================
// TEXT POSITION
// ============================================================================

/**
 * Represents the position of a single character in the flattened text,
 * mapped back to its source XML element.
 */
class TextPosition {
  /**
   * @param {object} node - Regex match info for the <w:t> element
   * @param {string} node.openTag - The <w:t ...> opening tag
   * @param {string} node.text - Full text content of this <w:t>
   * @param {string} node.fullMatch - Complete <w:t>...</w:t> match
   * @param {number} node.index - Character offset of node.fullMatch in paragraph XML
   * @param {number} node.runIndex - Index into the runs array
   * @param {string} node.rPr - Run properties XML for the parent <w:r>
   * @param {number} offset - Character offset within this node's text
   * @param {boolean} insideIns - Whether this position is inside a <w:ins>
   * @param {boolean} insideDel - Whether this position is inside a <w:del>
   */
  constructor(node, offset, insideIns, insideDel) {
    this.node = node;
    this.offset = offset;
    this.insideIns = insideIns;
    this.insideDel = insideDel;
  }
}

// ============================================================================
// TEXT MAP
// ============================================================================

/**
 * Maps flattened visible text back to source XML elements.
 *
 * Constructed from a <w:p> XML fragment. Scans all <w:r> elements,
 * extracts <w:t> text (skipping <w:delText>), and builds a character-level
 * position map.
 *
 * @example
 *   const tm = new TextMap(paragraphXml);
 *   const result = tm.find("Hello World");
 *   // result.spans tells you which <w:t> nodes and offsets to modify
 */
class TextMap {

  /**
   * Build a TextMap from paragraph XML.
   * @param {string} paragraphXml - Raw XML of a <w:p> element
   */
  constructor(paragraphXml) {
    /** @type {string} */
    this._paragraphXml = paragraphXml;

    /** @type {string} Flattened visible text */
    this._text = '';

    /** @type {TextPosition[]} One entry per character in _text */
    this._positions = [];

    this._build();
  }

  // --------------------------------------------------------------------------
  // Public accessors
  // --------------------------------------------------------------------------

  /**
   * The flattened visible text of the paragraph.
   * @returns {string}
   */
  get text() {
    return this._text;
  }

  /**
   * Array of TextPosition objects, one per character in the flattened text.
   * @returns {TextPosition[]}
   */
  get positions() {
    return this._positions;
  }

  // --------------------------------------------------------------------------
  // Search
  // --------------------------------------------------------------------------

  /**
   * Find the first occurrence of searchText in the flattened visible text.
   * @param {string} searchText - Text to find
   * @returns {{found: boolean, start: number, end: number, spans: Array<{node: object, startOffset: number, endOffset: number}>}}
   */
  find(searchText) {
    const idx = this._text.indexOf(searchText);
    if (idx === -1) {
      return { found: false, start: -1, end: -1, spans: [] };
    }
    return this._buildResult(idx, idx + searchText.length);
  }

  /**
   * Find all occurrences of searchText in the flattened visible text.
   * @param {string} searchText - Text to find
   * @returns {Array<{found: boolean, start: number, end: number, spans: Array<{node: object, startOffset: number, endOffset: number}>}>}
   */
  findAll(searchText) {
    const results = [];
    let startFrom = 0;
    while (true) {
      const idx = this._text.indexOf(searchText, startFrom);
      if (idx === -1) break;
      results.push(this._buildResult(idx, idx + searchText.length));
      startFrom = idx + 1; // allow overlapping matches
    }
    return results;
  }

  // --------------------------------------------------------------------------
  // Internal
  // --------------------------------------------------------------------------

  /**
   * Build a search result object for a character range.
   * Groups consecutive characters by their source node into spans.
   * @param {number} start - Start index in flattened text (inclusive)
   * @param {number} end - End index in flattened text (exclusive)
   * @returns {{found: boolean, start: number, end: number, spans: Array<{node: object, startOffset: number, endOffset: number}>}}
   * @private
   */
  _buildResult(start, end) {
    const spans = [];
    let currentNode = null;
    let spanStart = -1;
    let spanEnd = -1;

    for (let i = start; i < end; i++) {
      const pos = this._positions[i];
      if (pos.node !== currentNode) {
        // Flush previous span
        if (currentNode !== null) {
          spans.push({ node: currentNode, startOffset: spanStart, endOffset: spanEnd });
        }
        currentNode = pos.node;
        spanStart = pos.offset;
        spanEnd = pos.offset + 1;
      } else {
        spanEnd = pos.offset + 1;
      }
    }

    // Flush last span
    if (currentNode !== null) {
      spans.push({ node: currentNode, startOffset: spanStart, endOffset: spanEnd });
    }

    return { found: true, start, end, spans };
  }

  /**
   * Parse the paragraph XML and build the position map.
   *
   * Strategy:
   * 1. Find all <w:r> elements (runs) with their positions.
   * 2. For each run, determine if it is inside <w:ins> or <w:del>.
   * 3. Extract <w:t> text (not <w:delText>) and map each character.
   *
   * We only map visible text (from <w:t>), not deleted text (<w:delText>).
   * @private
   */
  _build() {
    const pXml = this._paragraphXml;
    const textChars = [];
    const positions = [];

    // Find all runs
    const runRe = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
    let runMatch;
    let runIndex = 0;

    while ((runMatch = runRe.exec(pXml)) !== null) {
      const runXml = runMatch[0];
      const runGlobalIndex = runMatch.index;

      // Determine context: is this run inside <w:ins> or <w:del>?
      const insideIns = this._isInsideTag(pXml, runGlobalIndex, 'w:ins');
      const insideDel = this._isInsideTag(pXml, runGlobalIndex, 'w:del');

      // Skip runs inside <w:del> for visible text mapping
      if (insideDel) {
        runIndex++;
        continue;
      }

      // Extract run properties
      const rPrMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
      const rPr = rPrMatch ? rPrMatch[0] : '';

      // Find all <w:t> elements in this run (not <w:delText>)
      const tRe = /(<w:t[^>]*>)([^<]*)<\/w:t>/g;
      let tMatch;
      while ((tMatch = tRe.exec(runXml)) !== null) {
        const openTag = tMatch[1];
        const text = tMatch[2];
        const fullMatch = tMatch[0];

        // Calculate the absolute position of this <w:t> in the paragraph
        const tGlobalIndex = runGlobalIndex + tMatch.index;

        const node = {
          openTag,
          text,
          fullMatch,
          index: tGlobalIndex,
          runIndex,
          rPr,
        };

        // Map each character
        for (let ci = 0; ci < text.length; ci++) {
          textChars.push(text[ci]);
          positions.push(new TextPosition(node, ci, insideIns, insideDel));
        }
      }

      runIndex++;
    }

    this._text = textChars.join('');
    this._positions = positions;
  }

  /**
   * Check whether a position in the XML is inside a given tag.
   * Looks backward from pos for an opening tag without a matching close tag.
   * @param {string} xml - Full XML string
   * @param {number} pos - Position to check
   * @param {string} tag - Tag name (e.g. 'w:ins', 'w:del')
   * @returns {boolean}
   * @private
   */
  _isInsideTag(xml, pos, tag) {
    // Find the last opening tag before pos
    const openPattern = `<${tag}`;
    const closePattern = `</${tag}>`;

    let lastOpen = -1;
    let searchFrom = 0;
    while (true) {
      const idx = xml.indexOf(openPattern, searchFrom);
      if (idx === -1 || idx >= pos) break;
      // Verify it is the actual tag (not a prefix of something longer)
      const charAfter = xml[idx + openPattern.length];
      if (charAfter === ' ' || charAfter === '>') {
        lastOpen = idx;
      }
      searchFrom = idx + 1;
    }

    if (lastOpen === -1) return false;

    // Find the matching close tag after the last open
    const closeIdx = xml.indexOf(closePattern, lastOpen);
    // If the close is after our position, we are inside
    return closeIdx !== -1 && closeIdx >= pos;
  }
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = { TextMap, TextPosition };
