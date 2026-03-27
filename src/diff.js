/**
 * diff.js -- Document comparison with tracked changes for docex
 *
 * Compares two documents paragraph by paragraph, producing tracked changes
 * (w:del + w:ins) showing what changed between them.
 *
 * Algorithm:
 *   1. Extract paragraphs from both documents
 *   2. Align paragraphs using longest common subsequence (LCS)
 *   3. For aligned paragraphs with different text: word-level diff
 *   4. For added paragraphs: wrap in w:ins
 *   5. For removed paragraphs: wrap content in w:del
 *   6. For modified paragraphs: w:del + w:ins within the paragraph
 *
 * Zero external dependencies. All XML manipulation is regex-based.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// DIFF CLASS
// ============================================================================

class Diff {

  /**
   * Compare two documents paragraph by paragraph.
   * Produces tracked changes (w:del + w:ins) showing what changed.
   * Writes the result to ws1 (modifies ws1's docXml).
   *
   * @param {object} ws1 - Workspace for the "original" document
   * @param {object} ws2 - Workspace for the "modified" document
   * @param {object} [opts] - Options
   * @param {string} [opts.author='Unknown'] - Author name for tracked changes
   * @param {string} [opts.date] - ISO date for tracked changes
   * @returns {{added: number, removed: number, modified: number, unchanged: number}}
   */
  static compare(ws1, ws2, opts = {}) {
    const author = opts.author || 'Unknown';
    const date = opts.date || xml.isoNow();

    const paras1 = xml.findParagraphs(ws1.docXml);
    const paras2 = xml.findParagraphs(ws2.docXml);

    const texts1 = paras1.map(p => xml.decodeXml(p.text));
    const texts2 = paras2.map(p => xml.decodeXml(p.text));

    // Align paragraphs using LCS
    const operations = Diff._diffParagraphs(texts1, texts2);

    // Build the output document XML by processing operations in order
    // We need to reconstruct the body content from the original document
    let nextId = xml.nextChangeId(ws1.docXml);
    const stats = { added: 0, removed: 0, modified: 0, unchanged: 0 };

    // Extract the body wrapper (everything before first <w:p and after last </w:p>)
    const bodyStart = ws1.docXml.indexOf('<w:body');
    const bodyEnd = ws1.docXml.lastIndexOf('</w:body>');
    const prefix = ws1.docXml.slice(0, bodyStart);
    const suffix = ws1.docXml.slice(bodyEnd);

    // Extract the <w:body...> opening tag
    const bodyOpenEnd = ws1.docXml.indexOf('>', bodyStart) + 1;
    const bodyOpenTag = ws1.docXml.slice(bodyStart, bodyOpenEnd);

    // Extract sectPr (section properties) if present -- it lives after the last paragraph
    // in the body, and we need to preserve it
    let sectPr = '';
    const lastPara1 = paras1.length > 0 ? paras1[paras1.length - 1] : null;
    if (lastPara1) {
      const afterLastPara = ws1.docXml.slice(lastPara1.end, bodyEnd);
      const sectPrMatch = afterLastPara.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
      if (sectPrMatch) {
        sectPr = sectPrMatch[0];
      }
    } else {
      // No paragraphs -- check if there's a sectPr in the body
      const bodyContent = ws1.docXml.slice(bodyOpenEnd, bodyEnd);
      const sectPrMatch = bodyContent.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
      if (sectPrMatch) {
        sectPr = sectPrMatch[0];
      }
    }

    // Build new body content from operations
    const bodyParts = [];

    for (const op of operations) {
      switch (op.type) {
        case 'keep': {
          // Copy paragraph as-is from ws1
          bodyParts.push(paras1[op.index1].xml);
          stats.unchanged++;
          break;
        }

        case 'add': {
          // New paragraph from ws2, wrap content in w:ins
          const p2 = paras2[op.index2];
          const pPr = Diff._extractPpr(p2.xml);
          const text2 = xml.decodeXml(p2.text);
          const insXml = Diff._buildInsertedParagraph(text2, pPr, nextId, author, date);
          nextId++;
          bodyParts.push(insXml);
          stats.added++;
          break;
        }

        case 'remove': {
          // Paragraph from ws1, wrap content in w:del
          const p1 = paras1[op.index1];
          const pPr = Diff._extractPpr(p1.xml);
          const rPr = Diff._extractFirstRpr(p1.xml);
          const text1 = xml.decodeXml(p1.text);
          const delXml = Diff._buildDeletedParagraph(text1, pPr, rPr, nextId, author, date);
          nextId++;
          bodyParts.push(delXml);
          stats.removed++;
          break;
        }

        case 'modify': {
          // Paragraphs aligned but text differs -- word-level diff
          const p1 = paras1[op.index1];
          const p2 = paras2[op.index2];
          const text1 = xml.decodeXml(p1.text);
          const text2 = xml.decodeXml(p2.text);
          const pPr = Diff._extractPpr(p1.xml);
          const rPr = Diff._extractFirstRpr(p1.xml);

          const segments = Diff._diffWords(text1, text2);
          const modXml = Diff._buildModifiedParagraph(segments, pPr, rPr, nextId, author, date);
          // Count IDs used: one per del segment, one per ins segment
          for (const seg of segments) {
            if (seg.type === 'remove' || seg.type === 'add') nextId++;
          }
          bodyParts.push(modXml);
          stats.modified++;
          break;
        }
      }
    }

    // Reassemble the document
    ws1.docXml = prefix + bodyOpenTag + bodyParts.join('') + sectPr + suffix;

    return stats;
  }

  // --------------------------------------------------------------------------
  // PARAGRAPH ALIGNMENT (LCS-based)
  // --------------------------------------------------------------------------

  /**
   * Align paragraphs using longest common subsequence (LCS).
   * Returns a list of operations: keep, add, remove, modify.
   *
   * @param {string[]} texts1 - Paragraph texts from document 1
   * @param {string[]} texts2 - Paragraph texts from document 2
   * @returns {Array<{type: string, index1?: number, index2?: number}>}
   */
  static _diffParagraphs(texts1, texts2) {
    const m = texts1.length;
    const n = texts2.length;

    // Build LCS table
    const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        if (texts1[i - 1] === texts2[j - 1]) {
          dp[i][j] = dp[i - 1][j - 1] + 1;
        } else {
          dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
        }
      }
    }

    // Backtrack to get aligned pairs
    const ops = [];
    let i = m, j = n;
    while (i > 0 || j > 0) {
      if (i > 0 && j > 0 && texts1[i - 1] === texts2[j - 1]) {
        ops.push({ type: 'keep', index1: i - 1, index2: j - 1 });
        i--;
        j--;
      } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
        ops.push({ type: 'add', index2: j - 1 });
        j--;
      } else {
        ops.push({ type: 'remove', index1: i - 1 });
        i--;
      }
    }

    ops.reverse();

    // Post-process: consecutive remove+add at same position become modify
    // if both have non-empty text
    const merged = [];
    let k = 0;
    while (k < ops.length) {
      if (k + 1 < ops.length
        && ops[k].type === 'remove'
        && ops[k + 1].type === 'add'
        && texts1[ops[k].index1].length > 0
        && texts2[ops[k + 1].index2].length > 0) {
        merged.push({
          type: 'modify',
          index1: ops[k].index1,
          index2: ops[k + 1].index2,
        });
        k += 2;
      } else {
        merged.push(ops[k]);
        k++;
      }
    }

    return merged;
  }

  // --------------------------------------------------------------------------
  // WORD-LEVEL DIFF (LCS on words)
  // --------------------------------------------------------------------------

  /**
   * Word-level diff within a modified paragraph.
   * Split into words, run LCS, generate keep/add/remove segments.
   *
   * @param {string} text1 - Original text
   * @param {string} text2 - Modified text
   * @returns {Array<{type: 'keep'|'add'|'remove', text: string}>}
   */
  static _diffWords(text1, text2) {
    const words1 = Diff._tokenize(text1);
    const words2 = Diff._tokenize(text2);

    const m = words1.length;
    const n = words2.length;

    // Build LCS table for words
    const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        if (words1[i - 1] === words2[j - 1]) {
          dp[i][j] = dp[i - 1][j - 1] + 1;
        } else {
          dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
        }
      }
    }

    // Backtrack
    const rawOps = [];
    let ii = m, jj = n;
    while (ii > 0 || jj > 0) {
      if (ii > 0 && jj > 0 && words1[ii - 1] === words2[jj - 1]) {
        rawOps.push({ type: 'keep', word: words1[ii - 1] });
        ii--;
        jj--;
      } else if (jj > 0 && (ii === 0 || dp[ii][jj - 1] >= dp[ii - 1][jj])) {
        rawOps.push({ type: 'add', word: words2[jj - 1] });
        jj--;
      } else {
        rawOps.push({ type: 'remove', word: words1[ii - 1] });
        ii--;
      }
    }

    rawOps.reverse();

    // Merge consecutive same-type operations into segments
    const segments = [];
    for (const op of rawOps) {
      if (segments.length > 0 && segments[segments.length - 1].type === op.type) {
        segments[segments.length - 1].text += op.word;
      } else {
        segments.push({ type: op.type, text: op.word });
      }
    }

    return segments;
  }

  // --------------------------------------------------------------------------
  // TOKENIZATION
  // --------------------------------------------------------------------------

  /**
   * Tokenize text into words, preserving whitespace attached to preceding words.
   * Each token includes the word plus any trailing whitespace.
   * This ensures that when tokens are joined, the original text is reconstructed.
   *
   * @param {string} text - Input text
   * @returns {string[]} Tokens
   */
  static _tokenize(text) {
    if (!text) return [];
    // Split into word+trailing-space tokens
    const tokens = [];
    const re = /\S+\s*/g;
    let m;
    while ((m = re.exec(text)) !== null) {
      tokens.push(m[0]);
    }
    return tokens;
  }

  // --------------------------------------------------------------------------
  // XML BUILDERS
  // --------------------------------------------------------------------------

  /**
   * Build a paragraph with all content wrapped in w:ins.
   *
   * @param {string} text - Paragraph text
   * @param {string} pPr - Paragraph properties XML
   * @param {number} id - Change tracking ID
   * @param {string} author - Author name
   * @param {string} date - ISO date string
   * @returns {string} <w:p> XML fragment
   * @private
   */
  static _buildInsertedParagraph(text, pPr, id, author, date) {
    const paraId = xml.randomHexId();
    const textId = xml.randomHexId();
    return `<w:p w14:paraId="${paraId}" w14:textId="${textId}">`
      + pPr
      + `<w:ins w:id="${id}" w:author="${xml.escapeXml(author)}" w:date="${date}">`
      + `<w:r><w:t xml:space="preserve">${xml.escapeXml(text)}</w:t></w:r>`
      + `</w:ins>`
      + `</w:p>`;
  }

  /**
   * Build a paragraph with all content wrapped in w:del.
   *
   * @param {string} text - Paragraph text
   * @param {string} pPr - Paragraph properties XML
   * @param {string} rPr - Run properties XML
   * @param {number} id - Change tracking ID
   * @param {string} author - Author name
   * @param {string} date - ISO date string
   * @returns {string} <w:p> XML fragment
   * @private
   */
  static _buildDeletedParagraph(text, pPr, rPr, id, author, date) {
    const paraId = xml.randomHexId();
    const textId = xml.randomHexId();
    return `<w:p w14:paraId="${paraId}" w14:textId="${textId}">`
      + pPr
      + `<w:del w:id="${id}" w:author="${xml.escapeXml(author)}" w:date="${date}">`
      + `<w:r>${rPr}<w:delText xml:space="preserve">${xml.escapeXml(text)}</w:delText></w:r>`
      + `</w:del>`
      + `</w:p>`;
  }

  /**
   * Build a paragraph with word-level tracked changes.
   *
   * @param {Array<{type: string, text: string}>} segments - Word diff segments
   * @param {string} pPr - Paragraph properties XML
   * @param {string} rPr - Run properties XML
   * @param {number} startId - Starting change tracking ID
   * @param {string} author - Author name
   * @param {string} date - ISO date string
   * @returns {string} <w:p> XML fragment
   * @private
   */
  static _buildModifiedParagraph(segments, pPr, rPr, startId, author, date) {
    const paraId = xml.randomHexId();
    const textId = xml.randomHexId();
    let currentId = startId;
    let runs = '';

    for (const seg of segments) {
      switch (seg.type) {
        case 'keep':
          runs += `<w:r>${rPr}<w:t xml:space="preserve">${xml.escapeXml(seg.text)}</w:t></w:r>`;
          break;
        case 'remove':
          runs += `<w:del w:id="${currentId}" w:author="${xml.escapeXml(author)}" w:date="${date}">`
            + `<w:r>${rPr}<w:delText xml:space="preserve">${xml.escapeXml(seg.text)}</w:delText></w:r>`
            + `</w:del>`;
          currentId++;
          break;
        case 'add':
          runs += `<w:ins w:id="${currentId}" w:author="${xml.escapeXml(author)}" w:date="${date}">`
            + `<w:r>${rPr}<w:t xml:space="preserve">${xml.escapeXml(seg.text)}</w:t></w:r>`
            + `</w:ins>`;
          currentId++;
          break;
      }
    }

    return `<w:p w14:paraId="${paraId}" w14:textId="${textId}">`
      + pPr
      + runs
      + `</w:p>`;
  }

  // --------------------------------------------------------------------------
  // HELPERS
  // --------------------------------------------------------------------------

  /**
   * Extract w:pPr (paragraph properties) from paragraph XML.
   * @param {string} pXml - Paragraph XML
   * @returns {string} pPr XML or ''
   * @private
   */
  static _extractPpr(pXml) {
    const selfClose = pXml.match(/<w:pPr\s*\/>/);
    if (selfClose) return selfClose[0];
    const full = pXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
    if (full) return full[0];
    const withAttrs = pXml.match(/<w:pPr\s[^>]*>[\s\S]*?<\/w:pPr>/);
    if (withAttrs) return withAttrs[0];
    return '';
  }

  /**
   * Extract the first run's w:rPr from paragraph XML.
   * Returns the full rPr element or empty string.
   * @param {string} pXml - Paragraph XML
   * @returns {string}
   * @private
   */
  static _extractFirstRpr(pXml) {
    const runMatch = pXml.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/);
    if (!runMatch) return '';
    const rPrMatch = runMatch[0].match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
    return rPrMatch ? rPrMatch[0] : '';
  }
}

module.exports = { Diff };
