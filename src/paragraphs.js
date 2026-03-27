/**
 * paragraphs.js -- Paragraph operations for docex
 *
 * Static methods for reading, searching, and modifying paragraphs in
 * OOXML document.xml. Wraps the proven logic from suggest-edit-safe.js
 * (tracked changes) and docx-patch.js (direct replacement, insertion).
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// PARAGRAPHS
// ============================================================================

class Paragraphs {

  // --------------------------------------------------------------------------
  // READ OPERATIONS
  // --------------------------------------------------------------------------

  /**
   * List all paragraphs in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {Array<{index: number, text: string, style: string}>}
   */
  static list(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);
    return paragraphs.map((p, i) => ({
      index: i,
      text: p.text,
      style: Paragraphs._getStyleId(p.xml),
    }));
  }

  /**
   * List all headings in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {Array<{level: number, text: string, index: number}>}
   */
  static headings(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);
    const results = [];
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const level = Paragraphs._headingLevel(p.xml);
      if (level > 0) {
        results.push({ level, text: p.text, index: i });
      }
    }
    return results;
  }

  /**
   * Get the full concatenated text of all paragraphs.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {string}
   */
  static fullText(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);
    return paragraphs.map(p => p.text).join('\n');
  }

  /**
   * Find paragraphs containing the given text.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} searchText - Text to search for (substring match)
   * @returns {Array<{index: number, text: string, style: string}>}
   */
  static find(ws, searchText) {
    const paragraphs = xml.findParagraphs(ws.docXml);
    const results = [];
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      if (p.text.includes(searchText)) {
        results.push({
          index: i,
          text: p.text,
          style: Paragraphs._getStyleId(p.xml),
        });
      }
    }
    return results;
  }

  // --------------------------------------------------------------------------
  // WRITE OPERATIONS
  // --------------------------------------------------------------------------

  /**
   * Replace text in the document.
   *
   * When tracked, the old text is wrapped in w:del and the new text in
   * w:ins. When untracked, the text is replaced directly in the XML
   * across w:t elements, preserving run formatting.
   *
   * Handles cross-run text (text split across multiple w:r elements):
   *   1. Concatenate run texts to find the match position
   *   2. Map back to individual runs
   *   3. Split affected runs at match boundaries
   *   4. Build replacement XML preserving formatting
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} oldText - Text to find
   * @param {string} newText - Replacement text
   * @param {object} [opts] - Options
   * @param {boolean} [opts.tracked=true] - Use tracked changes
   * @param {string} [opts.author='Unknown'] - Author for tracked changes
   * @param {string} [opts.date] - ISO date for tracked changes
   * @throws {Error} If text is not found in the document
   */
  static replace(ws, oldText, newText, opts = {}) {
    const tracked = opts.tracked !== undefined ? opts.tracked : true;

    if (tracked) {
      Paragraphs._replaceTracked(ws, oldText, newText, opts);
    } else {
      Paragraphs._replaceDirect(ws, oldText, newText);
    }
  }

  /**
   * Insert a new paragraph before or after an anchor paragraph.
   *
   * When tracked, the new paragraph is wrapped in w:ins. The new
   * paragraph copies pPr (paragraph properties) from the anchor
   * paragraph for consistent formatting.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} anchor - Text in the anchor paragraph
   * @param {string} mode - 'after' or 'before'
   * @param {string} text - Text content for the new paragraph
   * @param {object} [opts] - Options
   * @param {boolean} [opts.tracked=true] - Use tracked changes
   * @param {string} [opts.author='Unknown'] - Author for tracked changes
   * @param {string} [opts.date] - ISO date for tracked changes
   * @throws {Error} If anchor text is not found
   */
  static insert(ws, anchor, mode, text, opts = {}) {
    const tracked = opts.tracked !== undefined ? opts.tracked : true;
    const author = opts.author || 'Unknown';
    const date = opts.date || xml.isoNow();
    const docXml = ws.docXml;

    const paragraphs = xml.findParagraphs(docXml);

    // Find the anchor paragraph
    let anchorPara = null;
    for (const p of paragraphs) {
      if (p.text.includes(anchor)) {
        anchorPara = p;
        break;
      }
    }
    if (!anchorPara) {
      throw new Error('Anchor text not found: "' + anchor.slice(0, 80) + '"');
    }

    // Extract pPr from anchor paragraph for consistent formatting
    const pPr = Paragraphs._extractPpr(anchorPara.xml);
    const escapedText = xml.escapeXml(text);

    // Build the new paragraph
    let newParaXml;
    if (tracked) {
      const id = xml.nextChangeId(docXml);
      newParaXml = '<w:p>'
        + pPr
        + '<w:ins w:id="' + id + '" w:author="' + xml.escapeXml(author) + '" w:date="' + date + '">'
        + '<w:r>'
        + '<w:rPr>'
        + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman" />'
        + '<w:sz w:val="24" />'
        + '</w:rPr>'
        + '<w:t xml:space="preserve">' + escapedText + '</w:t>'
        + '</w:r>'
        + '</w:ins>'
        + '</w:p>';
    } else {
      newParaXml = '<w:p>'
        + pPr
        + '<w:r>'
        + '<w:rPr>'
        + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman" />'
        + '<w:sz w:val="24" />'
        + '</w:rPr>'
        + '<w:t xml:space="preserve">' + escapedText + '</w:t>'
        + '</w:r>'
        + '</w:p>';
    }

    // Insert before or after the anchor paragraph using string slicing
    if (mode === 'after') {
      ws.docXml = docXml.slice(0, anchorPara.end) + newParaXml + docXml.slice(anchorPara.end);
    } else {
      // 'before'
      ws.docXml = docXml.slice(0, anchorPara.start) + newParaXml + docXml.slice(anchorPara.start);
    }
  }

  /**
   * Remove a paragraph or text from the document.
   *
   * When tracked, the content is wrapped in w:del. When untracked,
   * the text (or entire paragraph) is removed from the XML.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} text - Text to find and remove
   * @param {object} [opts] - Options
   * @param {boolean} [opts.tracked=true] - Use tracked changes
   * @param {string} [opts.author='Unknown'] - Author for tracked changes
   * @param {string} [opts.date] - ISO date for tracked changes
   * @throws {Error} If text is not found
   */
  static remove(ws, text, opts = {}) {
    const tracked = opts.tracked !== undefined ? opts.tracked : true;

    if (tracked) {
      Paragraphs._deleteTracked(ws, text, opts);
    } else {
      Paragraphs._deleteDirect(ws, text);
    }
  }

  // --------------------------------------------------------------------------
  // TRACKED CHANGE INTERNALS (ported from suggest-edit-safe.js)
  // --------------------------------------------------------------------------

  /**
   * Tracked replacement: wrap old text in w:del, new text in w:ins.
   * Handles cross-run text boundaries.
   *
   * @param {object} ws - Workspace
   * @param {string} oldText - Text to replace
   * @param {string} newText - Replacement text
   * @param {object} opts - Options with author and date
   * @private
   */
  static _replaceTracked(ws, oldText, newText, opts) {
    const author = opts.author || 'Unknown';
    const date = opts.date || xml.isoNow();
    let docXml = ws.docXml;
    let nextId = xml.nextChangeId(docXml);

    const paragraphs = xml.findParagraphs(docXml);
    let found = false;

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      // Use decoded text for the paragraph-level contains check
      // (handles XML entities like &apos; -> ')
      if (!xml.decodeXml(para.text).includes(oldText)) continue;

      const result = Paragraphs._injectReplacement(
        para.xml, oldText, newText, nextId, author, date
      );
      if (result.modified) {
        docXml = docXml.slice(0, para.start) + result.xml + docXml.slice(para.end);
        nextId = result.nextId;
        found = true;
        break;
      }
    }

    if (!found) {
      throw new Error('Text not found for replacement: "' + oldText.slice(0, 80) + '"');
    }

    ws.docXml = docXml;
  }

  /**
   * Direct (untracked) text replacement across w:t elements.
   * Ported from docx-patch.js cmdReplaceText.
   *
   * @param {object} ws - Workspace
   * @param {string} oldText - Text to replace
   * @param {string} newText - Replacement text
   * @private
   */
  static _replaceDirect(ws, oldText, newText) {
    let docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    let found = false;

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      if (!xml.decodeXml(para.text).includes(oldText)) continue;

      found = true;
      const pXml = para.xml;

      // Extract all w:t elements with their positions in the paragraph XML
      const tRegex = /(<w:t[^>]*>)([^<]*)<\/w:t>/g;
      const runs = [];
      let tMatch;
      while ((tMatch = tRegex.exec(pXml)) !== null) {
        runs.push({
          fullMatch: tMatch[0],
          openTag: tMatch[1],
          text: tMatch[2],
          index: tMatch.index,
          length: tMatch[0].length,
        });
      }

      if (runs.length === 0) continue;

      // Build combined decoded text and track boundaries (handles XML entities)
      let combined = '';
      const boundaries = [];
      for (const run of runs) {
        const decodedText = xml.decodeXml(run.text);
        boundaries.push({
          start: combined.length,
          end: combined.length + decodedText.length,
          run,
          decodedText,
        });
        combined += decodedText;
      }

      const replaceStart = combined.indexOf(oldText);
      if (replaceStart === -1) continue;
      const replaceEnd = replaceStart + oldText.length;

      // Redistribute text back to runs, preserving formatting
      const replacements = [];
      let newTextUsed = false;

      for (const b of boundaries) {
        let runNewText;
        // Use decodedText for slicing (offsets are in decoded space)
        const dt = b.decodedText;

        if (b.end <= replaceStart || b.start >= replaceEnd) {
          // Entirely outside the replacement zone (use decoded text)
          runNewText = dt;
        } else if (b.start >= replaceStart && b.end <= replaceEnd) {
          // Entirely within the replacement zone
          if (!newTextUsed) {
            runNewText = newText;
            newTextUsed = true;
          } else {
            runNewText = '';
          }
        } else if (b.start < replaceStart && b.end > replaceEnd) {
          // Replacement is entirely within this single run
          const before = dt.slice(0, replaceStart - b.start);
          const after = dt.slice(replaceEnd - b.start);
          runNewText = before + newText + after;
          newTextUsed = true;
        } else if (b.start < replaceStart) {
          // Starts before replacement, ends during it
          runNewText = dt.slice(0, replaceStart - b.start) + newText;
          newTextUsed = true;
        } else {
          // Starts during replacement, ends after it
          runNewText = dt.slice(replaceEnd - b.start);
        }

        replacements.push({ run: b.run, newText: runNewText });
      }

      // Apply replacements in reverse order to preserve character indices
      let newPXml = pXml;
      for (let j = replacements.length - 1; j >= 0; j--) {
        const { run, newText: rt } = replacements[j];
        const escapedRt = xml.escapeXml(rt);
        const newTElement = run.openTag + escapedRt + '</w:t>';
        newPXml = newPXml.slice(0, run.index) + newTElement + newPXml.slice(run.index + run.length);
      }

      // Replace paragraph in full document XML
      docXml = docXml.slice(0, para.start) + newPXml + docXml.slice(para.end);
      break; // replace first occurrence only
    }

    if (!found) {
      throw new Error('Text not found for replacement: "' + oldText.slice(0, 80) + '"');
    }

    ws.docXml = docXml;
  }

  /**
   * Tracked deletion: wrap matching text in w:del.
   * Ported from suggest-edit-safe.js injectDeletion.
   *
   * @param {object} ws - Workspace
   * @param {string} text - Text to delete
   * @param {object} opts - Options with author and date
   * @private
   */
  static _deleteTracked(ws, text, opts) {
    const author = opts.author || 'Unknown';
    const date = opts.date || xml.isoNow();
    let docXml = ws.docXml;
    let nextId = xml.nextChangeId(docXml);

    const paragraphs = xml.findParagraphs(docXml);
    let found = false;

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      if (!xml.decodeXml(para.text).includes(text)) continue;

      const result = Paragraphs._injectDeletion(
        para.xml, text, nextId, author, date
      );
      if (result.modified) {
        docXml = docXml.slice(0, para.start) + result.xml + docXml.slice(para.end);
        nextId = result.nextId;
        found = true;
        break;
      }
    }

    if (!found) {
      throw new Error('Text not found for deletion: "' + text.slice(0, 80) + '"');
    }

    ws.docXml = docXml;
  }

  /**
   * Direct (untracked) deletion: remove text from document.
   * If the text matches an entire paragraph, the paragraph element is removed.
   * Otherwise, the text is removed from within the paragraph using direct
   * replacement with an empty string.
   *
   * @param {object} ws - Workspace
   * @param {string} text - Text to delete
   * @private
   */
  static _deleteDirect(ws, text) {
    let docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    let found = false;

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      if (!xml.decodeXml(para.text).includes(text)) continue;

      found = true;

      // If the entire paragraph text matches, remove the whole paragraph
      if (xml.decodeXml(para.text).trim() === text.trim()) {
        docXml = docXml.slice(0, para.start) + docXml.slice(para.end);
        ws.docXml = docXml;
      } else {
        // Remove text within the paragraph (replace with empty string).
        // Use a proxy object so _replaceDirect can read/write the XML.
        const proxy = {
          _xml: docXml,
          get docXml() { return this._xml; },
          set docXml(v) { this._xml = v; },
        };
        Paragraphs._replaceDirect(proxy, text, '');
        ws.docXml = proxy.docXml;
      }
      break;
    }

    if (!found) {
      throw new Error('Text not found for deletion: "' + text.slice(0, 80) + '"');
    }
  }

  // --------------------------------------------------------------------------
  // CROSS-RUN TEXT HANDLING (ported from suggest-edit-safe.js)
  // --------------------------------------------------------------------------

  /**
   * Find text that may span multiple w:r elements in a paragraph.
   * Returns match info including which runs contain the text and offsets.
   *
   * @param {string} pXml - Paragraph XML
   * @param {string} searchText - Text to find
   * @returns {object} Match result with found, textRuns, offsets
   * @private
   */
  static _findTextInParagraph(pXml, searchText) {
    const allRuns = xml.parseRuns(pXml);
    const textRuns = allRuns.filter(r => r.texts.length > 0);
    if (textRuns.length === 0) return { found: false };

    // Build decoded combined text for searching (handles XML entities like &apos; -> ')
    const decodedTexts = textRuns.map(r => xml.decodeXml(r.combinedText));
    const combined = decodedTexts.join('');
    const matchPos = combined.indexOf(searchText);
    if (matchPos === -1) return { found: false };

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

    if (matchStartRun === -1 || matchEndRun === -1) return { found: false };

    return {
      found: true,
      allRuns,
      textRuns,
      matchStartRun,
      matchEndRun,
      matchStartOffset,
      matchEndOffset,
    };
  }

  /**
   * Inject a tracked replacement into a paragraph.
   *
   * Builds segments: keep-prefix + del(old) + ins(new) + keep-suffix,
   * then replaces the affected runs in the paragraph XML.
   *
   * @param {string} pXml - Paragraph XML
   * @param {string} searchText - Text to replace
   * @param {string} newText - Replacement text
   * @param {number} nextId - Next available change ID
   * @param {string} author - Author name
   * @param {string} date - ISO date string
   * @returns {{modified: boolean, xml: string, nextId: number}}
   * @private
   */
  static _injectReplacement(pXml, searchText, newText, nextId, author, date) {
    const match = Paragraphs._findTextInParagraph(pXml, searchText);
    if (!match.found) return { modified: false, xml: pXml, nextId };

    const { textRuns, matchStartRun, matchEndRun, matchStartOffset, matchEndOffset } = match;
    const rPr = textRuns[matchStartRun].rPr;

    // Build segments from affected runs.
    // NOTE: matchStartOffset and matchEndOffset are in decoded character space.
    // We work with decoded text for slicing, then escapeXml when writing.
    const segments = [];
    for (let i = matchStartRun; i <= matchEndRun; i++) {
      const run = textRuns[i];
      const runText = xml.decodeXml(run.combinedText); // work in decoded space
      const fullLen = runText.length;
      let sliceStart = 0;
      let sliceEnd = fullLen;
      if (i === matchStartRun) sliceStart = matchStartOffset;
      if (i === matchEndRun) sliceEnd = matchEndOffset;

      // Keep prefix (text before match in the start run)
      if (i === matchStartRun && matchStartOffset > 0) {
        segments.push({ type: 'keep', text: runText.slice(0, matchStartOffset), rPr: run.rPr });
      }

      // Matched text to delete
      const matchedText = runText.slice(sliceStart, sliceEnd);
      if (matchedText.length > 0) {
        segments.push({ type: 'delete', text: matchedText, rPr: run.rPr });
      }

      // Insert new text (once, at the end of the match)
      if (i === matchEndRun) {
        segments.push({ type: 'insert', text: newText, rPr: rPr });
      }

      // Keep suffix (text after match in the end run)
      if (i === matchEndRun && matchEndOffset < fullLen) {
        segments.push({ type: 'keep', text: runText.slice(matchEndOffset), rPr: run.rPr });
      }
    }

    // Collect full deleted text
    const delId = nextId;
    const insId = nextId + 1;
    let delText = '';
    for (const seg of segments) {
      if (seg.type === 'delete') delText += seg.text;
    }

    // Build replacement XML
    let replacementXml = '';
    let delEmitted = false;
    for (const seg of segments) {
      if (seg.type === 'keep') {
        replacementXml += '<w:r>' + seg.rPr + '<w:t xml:space="preserve">' + xml.escapeXml(seg.text) + '</w:t></w:r>';
      } else if (seg.type === 'delete') {
        if (!delEmitted) {
          replacementXml += xml.buildDel(delId, author, date, rPr, delText);
          delEmitted = true;
        }
      } else if (seg.type === 'insert') {
        // w:del carries the author attribution; w:ins uses empty author to avoid
        // duplicate author mentions when querying the document
        replacementXml += `<w:ins w:id="${insId}" w:author="" w:date="${date}"><w:r>${rPr}<w:t xml:space="preserve">${xml.escapeXml(seg.text)}</w:t></w:r></w:ins>`;
      }
    }

    // Splice replacement into paragraph XML
    const startPos = textRuns[matchStartRun].index;
    const lastRun = textRuns[matchEndRun];
    const endPos = lastRun.index + lastRun.fullMatch.length;
    const newPXml = pXml.slice(0, startPos) + replacementXml + pXml.slice(endPos);

    return { modified: true, xml: newPXml, nextId: insId + 1 };
  }

  /**
   * Inject a tracked deletion into a paragraph.
   *
   * Builds segments: keep-prefix + del(matched) + keep-suffix,
   * then replaces the affected runs in the paragraph XML.
   *
   * @param {string} pXml - Paragraph XML
   * @param {string} searchText - Text to delete
   * @param {number} nextId - Next available change ID
   * @param {string} author - Author name
   * @param {string} date - ISO date string
   * @returns {{modified: boolean, xml: string, nextId: number}}
   * @private
   */
  static _injectDeletion(pXml, searchText, nextId, author, date) {
    const match = Paragraphs._findTextInParagraph(pXml, searchText);
    if (!match.found) return { modified: false, xml: pXml, nextId };

    const { textRuns, matchStartRun, matchEndRun, matchStartOffset, matchEndOffset } = match;
    const rPr = textRuns[matchStartRun].rPr;

    // Build segments
    // NOTE: matchStartOffset and matchEndOffset are in decoded character space
    const segments = [];
    for (let i = matchStartRun; i <= matchEndRun; i++) {
      const run = textRuns[i];
      const runText = xml.decodeXml(run.combinedText); // work in decoded space
      const fullLen = runText.length;
      let sliceStart = 0;
      let sliceEnd = fullLen;
      if (i === matchStartRun) sliceStart = matchStartOffset;
      if (i === matchEndRun) sliceEnd = matchEndOffset;

      if (i === matchStartRun && matchStartOffset > 0) {
        segments.push({ type: 'keep', text: runText.slice(0, matchStartOffset), rPr: run.rPr });
      }

      const matchedText = runText.slice(sliceStart, sliceEnd);
      if (matchedText.length > 0) {
        segments.push({ type: 'delete', text: matchedText, rPr: run.rPr });
      }

      if (i === matchEndRun && matchEndOffset < fullLen) {
        segments.push({ type: 'keep', text: runText.slice(matchEndOffset), rPr: run.rPr });
      }
    }

    // Collect deleted text
    let delText = '';
    for (const seg of segments) {
      if (seg.type === 'delete') delText += seg.text;
    }

    // Build replacement XML
    let replacementXml = '';
    const delId = nextId;
    let delEmitted = false;
    for (const seg of segments) {
      if (seg.type === 'keep') {
        replacementXml += '<w:r>' + seg.rPr + '<w:t xml:space="preserve">' + xml.escapeXml(seg.text) + '</w:t></w:r>';
      } else if (seg.type === 'delete') {
        if (!delEmitted) {
          replacementXml += xml.buildDel(delId, author, date, rPr, delText);
          delEmitted = true;
        }
      }
    }

    // Splice into paragraph
    const startPos = textRuns[matchStartRun].index;
    const lastRun = textRuns[matchEndRun];
    const endPos = lastRun.index + lastRun.fullMatch.length;
    const newPXml = pXml.slice(0, startPos) + replacementXml + pXml.slice(endPos);

    return { modified: true, xml: newPXml, nextId: delId + 1 };
  }

  // --------------------------------------------------------------------------
  // INTERNAL HELPERS
  // --------------------------------------------------------------------------

  /**
   * Extract the paragraph style ID from paragraph XML.
   *
   * @param {string} pXml - Paragraph XML
   * @returns {string} Style ID or empty string
   * @private
   */
  static _getStyleId(pXml) {
    const m = pXml.match(/<w:pStyle\s+w:val="([^"]+)"/);
    return m ? m[1] : '';
  }

  /**
   * Determine the heading level from paragraph XML.
   * Returns 0 if the paragraph is not a heading.
   *
   * Checks both named styles (Heading1, Heading2...) and the outlineLvl
   * element. Also recognizes numeric style IDs used by OnlyOffice
   * Document Builder.
   *
   * @param {string} pXml - Paragraph XML
   * @returns {number} Heading level (1-9) or 0
   * @private
   */
  static _headingLevel(pXml) {
    // Check for outlineLvl first (most reliable)
    const olvl = pXml.match(/<w:outlineLvl\s+w:val="(\d+)"/);
    if (olvl) {
      return parseInt(olvl[1], 10) + 1; // outlineLvl is 0-based
    }

    // Check style names
    const styleId = Paragraphs._getStyleId(pXml);
    if (!styleId) return 0;

    // Standard named styles
    const namedMatch = styleId.match(/^[Hh]eading(\d+)$/);
    if (namedMatch) return parseInt(namedMatch[1], 10);

    // OnlyOffice numeric IDs: 841=H2, 842=H3, etc. (840=H1 in some docs)
    const DOCBUILDER_MAP = {
      '840': 1, '841': 2, '842': 3, '843': 4, '844': 5, '845': 6,
      '139': 1, '140': 2, '141': 3, '142': 4, '143': 5,
    };
    if (DOCBUILDER_MAP[styleId]) return DOCBUILDER_MAP[styleId];

    return 0;
  }

  /**
   * Extract w:pPr (paragraph properties) from paragraph XML.
   * Returns the full pPr element or an empty string.
   *
   * @param {string} pXml - Paragraph XML
   * @returns {string} pPr XML or ''
   * @private
   */
  static _extractPpr(pXml) {
    // Match self-closing pPr
    const selfClose = pXml.match(/<w:pPr\s*\/>/);
    if (selfClose) return selfClose[0];

    // Match regular pPr (no attributes on opening tag)
    const full = pXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
    if (full) return full[0];

    // Match pPr with attributes
    const withAttrs = pXml.match(/<w:pPr\s[^>]*>[\s\S]*?<\/w:pPr>/);
    if (withAttrs) return withAttrs[0];

    return '';
  }
}

module.exports = { Paragraphs };
