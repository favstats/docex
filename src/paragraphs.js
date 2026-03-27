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
// CLOSEST-MATCH HELPER
// ============================================================================

/**
 * Score how similar two strings are by counting the longest common
 * subsequence (character-level). Returns a value between 0 and 1.
 *
 * @param {string} a - First string
 * @param {string} b - Second string
 * @returns {number} Similarity score (0..1)
 */
function _lcsScore(a, b) {
  if (!a || !b) return 0;
  const al = a.toLowerCase();
  const bl = b.toLowerCase();
  // Fast path: substring match
  if (bl.includes(al) || al.includes(bl)) return 1;

  // LCS length via two-row DP (memory efficient)
  const m = al.length;
  const n = bl.length;
  if (m === 0 || n === 0) return 0;
  // Limit to avoid quadratic blow-up on very long paragraphs
  const cap = 300;
  const sa = m > cap ? al.slice(0, cap) : al;
  const sb = n > cap ? bl.slice(0, cap) : bl;
  const rows = 2;
  const cols = sb.length + 1;
  const dp = [new Uint16Array(cols), new Uint16Array(cols)];
  for (let i = 1; i <= sa.length; i++) {
    const cur = dp[i & 1];
    const prev = dp[(i - 1) & 1];
    for (let j = 1; j <= sb.length; j++) {
      if (sa[i - 1] === sb[j - 1]) {
        cur[j] = prev[j - 1] + 1;
      } else {
        cur[j] = Math.max(prev[j], cur[j - 1]);
      }
    }
  }
  const lcs = dp[sa.length & 1][sb.length];
  return lcs / Math.max(sa.length, sb.length);
}

/**
 * Find the closest matching paragraphs to a search string.
 *
 * @param {Array<{text: string}>} paragraphs - Paragraphs with .text
 * @param {string} searchText - Text that was not found
 * @param {number} [limit=3] - Max number of matches to return
 * @returns {string[]} Array of truncated paragraph texts, best first
 */
function findClosestMatches(paragraphs, searchText, limit = 3) {
  const scored = [];
  for (const p of paragraphs) {
    const decoded = xml.decodeXml(p.text);
    if (!decoded.trim()) continue;
    const score = _lcsScore(searchText, decoded);
    scored.push({ text: decoded, score });
  }
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, limit).map(s => {
    const t = s.text.trim();
    return t.length > 60 ? t.slice(0, 57) + '...' : t;
  });
}

// ============================================================================
// FUZZY RETRY HELPERS
// ============================================================================

/**
 * Normalize whitespace: collapse multiple spaces/newlines/tabs into single space, trim.
 * @param {string} s
 * @returns {string}
 */
function _normalizeWhitespace(s) {
  return s.replace(/\s+/g, ' ').trim();
}

/**
 * Try to find text in paragraphs using fuzzy matching strategies.
 * Attempts in order:
 *   1. Case-insensitive match
 *   2. Normalized whitespace match
 *   3. Decoded XML entities match (collapse &amp; etc.)
 *
 * Returns the actual text that matched in the paragraph (needed for replacement),
 * or null if no fuzzy match found.
 *
 * @param {Array<{text: string, xml: string}>} paragraphs
 * @param {string} searchText
 * @returns {{paraIndex: number, matchedText: string, strategy: string}|null}
 */
function fuzzyFindText(paragraphs, searchText) {
  // Strategy 1: case-insensitive
  const lowerSearch = searchText.toLowerCase();
  for (let i = 0; i < paragraphs.length; i++) {
    const decoded = xml.decodeXml(paragraphs[i].text);
    const lowerPara = decoded.toLowerCase();
    const pos = lowerPara.indexOf(lowerSearch);
    if (pos !== -1) {
      // Return the actual-cased text from the paragraph
      return {
        paraIndex: i,
        matchedText: decoded.slice(pos, pos + searchText.length),
        strategy: 'case-insensitive',
      };
    }
  }

  // Strategy 2: normalized whitespace
  const normSearch = _normalizeWhitespace(searchText);
  for (let i = 0; i < paragraphs.length; i++) {
    const decoded = xml.decodeXml(paragraphs[i].text);
    const normPara = _normalizeWhitespace(decoded);
    const pos = normPara.indexOf(normSearch);
    if (pos !== -1) {
      // We need to find the corresponding original text. We'll walk the original
      // decoded text mapping normalized positions back.
      const origText = _mapNormalizedBack(decoded, pos, normSearch.length);
      if (origText) {
        return {
          paraIndex: i,
          matchedText: origText,
          strategy: 'normalized-whitespace',
        };
      }
    }
  }

  // Strategy 3: case-insensitive + normalized whitespace
  const normSearchLower = normSearch.toLowerCase();
  for (let i = 0; i < paragraphs.length; i++) {
    const decoded = xml.decodeXml(paragraphs[i].text);
    const normPara = _normalizeWhitespace(decoded).toLowerCase();
    const pos = normPara.indexOf(normSearchLower);
    if (pos !== -1) {
      const origText = _mapNormalizedBack(decoded, pos, normSearchLower.length);
      if (origText) {
        return {
          paraIndex: i,
          matchedText: origText,
          strategy: 'case-insensitive-normalized',
        };
      }
    }
  }

  return null;
}

/**
 * Map a position in a whitespace-normalized string back to the original string.
 * Returns the slice of the original that corresponds to the normalized range.
 *
 * @param {string} original - Original (decoded) text
 * @param {number} normPos - Start position in the normalized string
 * @param {number} normLen - Length in the normalized string
 * @returns {string|null}
 */
function _mapNormalizedBack(original, normPos, normLen) {
  // Build mapping: normalized index -> original index
  let normIdx = 0;
  let origStart = -1;
  let origEnd = -1;
  let inWhitespace = false;
  let origIdx = 0;

  // Skip leading whitespace
  while (origIdx < original.length && /\s/.test(original[origIdx])) {
    origIdx++;
  }

  for (; origIdx < original.length; origIdx++) {
    const ch = original[origIdx];
    if (/\s/.test(ch)) {
      if (!inWhitespace) {
        // First whitespace char -> maps to single space in normalized
        if (normIdx === normPos) origStart = origIdx;
        normIdx++;
        if (normIdx === normPos + normLen && origStart !== -1) {
          origEnd = origIdx + 1;
          break;
        }
        inWhitespace = true;
      }
      // Additional whitespace chars are skipped
    } else {
      inWhitespace = false;
      if (normIdx === normPos) origStart = origIdx;
      normIdx++;
      if (normIdx === normPos + normLen && origStart !== -1) {
        origEnd = origIdx + 1;
        break;
      }
    }
  }

  if (origStart !== -1 && origEnd !== -1) {
    return original.slice(origStart, origEnd);
  }
  return null;
}

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

  /**
   * Count words in the document, categorized by type.
   *
   * Categories:
   *   - body: regular body text (everything not in another category)
   *   - headings: text in heading paragraphs
   *   - abstract: paragraphs between an "Abstract" heading and the next heading
   *   - captions: paragraphs starting with "Figure" or "Table" followed by a number
   *   - footnotes: text in word/footnotes.xml (if it exists)
   *
   * @param {object} ws - Workspace with ws.docXml and optionally ws.footnotesXml
   * @returns {{total: number, body: number, headings: number, abstract: number, captions: number, footnotes: number}}
   */
  static wordCount(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);

    let headingWords = 0;
    let abstractWords = 0;
    let captionWords = 0;
    let bodyWords = 0;

    let inAbstract = false;

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const text = xml.extractTextDecoded(p.xml).trim();
      if (!text) continue;

      const words = Paragraphs._countWords(text);
      const level = Paragraphs._headingLevel(p.xml);

      if (level > 0) {
        // It's a heading
        headingWords += words;

        // Check if this heading is "Abstract" (case-insensitive)
        if (/^abstract$/i.test(text.trim())) {
          inAbstract = true;
        } else {
          inAbstract = false;
        }
        continue;
      }

      // A standalone "Abstract" label paragraph (not a heading style but acts as one)
      if (/^abstract$/i.test(text)) {
        // Count the word "Abstract" as a heading, start abstract section
        headingWords += words;
        inAbstract = true;
        continue;
      }

      // Check for "Keywords:" paragraph -- ends abstract section, categorize as abstract
      if (inAbstract && /^keywords?\s*:/i.test(text)) {
        abstractWords += words;
        inAbstract = false;
        continue;
      }

      // Check for caption: starts with "Figure" or "Table" followed by a number
      if (/^(Figure|Table)\s+\d/i.test(text)) {
        captionWords += words;
        continue;
      }

      // If we're in the abstract section
      if (inAbstract) {
        abstractWords += words;
        continue;
      }

      // Everything else is body text
      bodyWords += words;
    }

    // Count footnote words
    let footnoteWords = 0;
    const footnotesXml = ws.footnotesXml;
    if (footnotesXml) {
      // Find all footnote elements (skip the separator/continuation footnotes with id 0 and 1)
      const fnRe = /<w:footnote\b([^>]*)>([\s\S]*?)<\/w:footnote>/g;
      let fnMatch;
      while ((fnMatch = fnRe.exec(footnotesXml)) !== null) {
        const attrs = fnMatch[1];
        // Skip separator and continuation footnotes (w:type="separator" or w:type="continuationSeparator")
        if (/w:type="/.test(attrs)) continue;
        const fnBody = fnMatch[2];
        const fnText = xml.extractTextDecoded(fnBody).trim();
        if (fnText) {
          footnoteWords += Paragraphs._countWords(fnText);
        }
      }
    }

    const total = bodyWords + headingWords + abstractWords + captionWords + footnoteWords;

    return {
      total,
      body: bodyWords,
      headings: headingWords,
      abstract: abstractWords,
      captions: captionWords,
      footnotes: footnoteWords,
    };
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
      const closest = findClosestMatches(paragraphs, anchor);
      let msg = 'Anchor text not found: "' + anchor.slice(0, 80) + '"';
      if (closest.length > 0) {
        msg += '\nDid you mean:\n' + closest.map(c => '  - "' + c + '"').join('\n');
      }
      throw new Error(msg);
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

  /**
   * Replace ALL occurrences of oldText in the document (not just the first).
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} oldText - Text to find
   * @param {string} newText - Replacement text
   * @param {object} [opts] - Options (same as replace)
   * @returns {number} Count of replacements made
   */
  static replaceAll(ws, oldText, newText, opts = {}) {
    let count = 0;
    // Keep replacing until no more matches are found
    while (true) {
      const paragraphs = xml.findParagraphs(ws.docXml);
      let found = false;
      for (const para of paragraphs) {
        if (xml.decodeXml(para.text).includes(oldText)) {
          found = true;
          break;
        }
      }
      if (!found) break;
      try {
        Paragraphs.replace(ws, oldText, newText, opts);
        count++;
      } catch (_) {
        break;
      }
    }
    return count;
  }

  /**
   * Replace text matching a regular expression pattern.
   *
   * Finds paragraphs whose decoded text matches the pattern, then replaces
   * each match using the standard replace() method. Pattern must have the
   * global flag to replace all occurrences.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {RegExp} pattern - Regular expression to match
   * @param {string} replacement - Replacement string (supports $1, $2, etc.)
   * @param {object} [opts] - Options (same as replace)
   * @returns {number} Count of replacements made
   */
  static replaceRegex(ws, pattern, replacement, opts = {}) {
    const flags = pattern.flags.includes('g') ? pattern.flags : pattern.flags + 'g';
    const globalRe = new RegExp(pattern.source, flags);
    let count = 0;

    // Collect all matches first by scanning paragraphs
    while (true) {
      const paragraphs = xml.findParagraphs(ws.docXml);
      let matchFound = false;

      for (const para of paragraphs) {
        const decoded = xml.decodeXml(para.text);
        // Reset regex state
        globalRe.lastIndex = 0;
        const m = globalRe.exec(decoded);
        if (m) {
          const matchedText = m[0];
          // Build replacement string with group substitutions
          let replaced = replacement;
          for (let g = 0; g < m.length; g++) {
            replaced = replaced.split('$' + g).join(m[g] || '');
          }
          // Also handle named groups would need more complex logic;
          // for now support $1..$9
          try {
            Paragraphs.replace(ws, matchedText, replaced, opts);
            count++;
            matchFound = true;
            break; // restart scan since XML changed
          } catch (_) {
            // If replace fails for this match, skip it
            matchFound = false;
            break;
          }
        }
      }
      if (!matchFound) break;
    }

    return count;
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

    // Fuzzy retry if exact match failed
    if (!found) {
      const fuzzy = fuzzyFindText(paragraphs, oldText);
      if (fuzzy) {
        // Retry with the actual matched text from the paragraph
        docXml = ws.docXml; // refresh
        const retryParagraphs = xml.findParagraphs(docXml);
        if (fuzzy.paraIndex < retryParagraphs.length) {
          const result = Paragraphs._injectReplacement(
            retryParagraphs[fuzzy.paraIndex].xml, fuzzy.matchedText, newText, nextId, author, date
          );
          if (result.modified) {
            const para = retryParagraphs[fuzzy.paraIndex];
            docXml = docXml.slice(0, para.start) + result.xml + docXml.slice(para.end);
            found = true;
          }
        }
      }
    }

    if (!found) {
      const closest = findClosestMatches(paragraphs, oldText);
      let msg = 'Text not found for replacement: "' + oldText.slice(0, 80) + '"';
      if (closest.length > 0) {
        msg += '\nDid you mean:\n' + closest.map(c => '  - "' + c + '"').join('\n');
      }
      throw new Error(msg);
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

    // Fuzzy retry if exact match failed
    if (!found) {
      const fuzzy = fuzzyFindText(paragraphs, oldText);
      if (fuzzy) {
        // Retry with the actual matched text using a proxy
        const proxy = {
          _xml: ws.docXml,
          get docXml() { return this._xml; },
          set docXml(v) { this._xml = v; },
        };
        try {
          Paragraphs._replaceDirect(proxy, fuzzy.matchedText, newText);
          docXml = proxy.docXml;
          found = true;
        } catch (_) { /* fuzzy retry also failed, fall through */ }
      }
    }

    if (!found) {
      const closest = findClosestMatches(paragraphs, oldText);
      let msg = 'Text not found for replacement: "' + oldText.slice(0, 80) + '"';
      if (closest.length > 0) {
        msg += '\nDid you mean:\n' + closest.map(c => '  - "' + c + '"').join('\n');
      }
      throw new Error(msg);
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

    // Fuzzy retry if exact match failed
    if (!found) {
      const fuzzy = fuzzyFindText(paragraphs, text);
      if (fuzzy) {
        docXml = ws.docXml;
        const retryParagraphs = xml.findParagraphs(docXml);
        if (fuzzy.paraIndex < retryParagraphs.length) {
          const result = Paragraphs._injectDeletion(
            retryParagraphs[fuzzy.paraIndex].xml, fuzzy.matchedText, nextId, author, date
          );
          if (result.modified) {
            const para = retryParagraphs[fuzzy.paraIndex];
            docXml = docXml.slice(0, para.start) + result.xml + docXml.slice(para.end);
            found = true;
          }
        }
      }
    }

    if (!found) {
      const closest = findClosestMatches(paragraphs, text);
      let msg = 'Text not found for deletion: "' + text.slice(0, 80) + '"';
      if (closest.length > 0) {
        msg += '\nDid you mean:\n' + closest.map(c => '  - "' + c + '"').join('\n');
      }
      throw new Error(msg);
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

    // Fuzzy retry if exact match failed
    if (!found) {
      const fuzzy = fuzzyFindText(paragraphs, text);
      if (fuzzy) {
        // Retry with the actual matched text
        try {
          Paragraphs._deleteDirect(ws, fuzzy.matchedText);
          found = true;
        } catch (_) { /* fuzzy retry also failed */ }
      }
    }

    if (!found) {
      const closest = findClosestMatches(paragraphs, text);
      let msg = 'Text not found for deletion: "' + text.slice(0, 80) + '"';
      if (closest.length > 0) {
        msg += '\nDid you mean:\n' + closest.map(c => '  - "' + c + '"').join('\n');
      }
      throw new Error(msg);
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
   * Count words in a plain text string (split by whitespace).
   *
   * @param {string} text - Plain text
   * @returns {number} Word count
   * @private
   */
  static _countWords(text) {
    const trimmed = text.trim();
    if (!trimmed) return 0;
    return trimmed.split(/\s+/).length;
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

module.exports = { Paragraphs, findClosestMatches, fuzzyFindText };
