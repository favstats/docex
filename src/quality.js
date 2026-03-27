/**
 * quality.js -- Manuscript quality and writing checks for docex
 *
 * Static methods for linting, passive voice detection, readability
 * scoring, sentence length analysis, and number verification.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const fs = require('fs');
const xml = require('./xml');

// ============================================================================
// HELPERS
// ============================================================================

/**
 * Extract all paragraphs from the workspace as an array of
 * { text, xml, paraId } objects.
 *
 * @param {object} ws - Workspace with ws.docXml
 * @returns {Array<{text: string, xml: string, paraId: string|null}>}
 */
function _extractParagraphs(ws) {
  const paragraphs = xml.findParagraphs(ws.docXml);
  return paragraphs.map(p => {
    const text = xml.extractTextDecoded(p.xml);
    // Extract paraId from w14:paraId attribute
    const idMatch = p.xml.match(/w14:paraId="([^"]+)"/);
    return {
      text,
      xml: p.xml,
      paraId: idMatch ? idMatch[1] : null,
    };
  });
}

/**
 * Count the number of figures (images) in the document.
 * @param {string} docXml
 * @returns {number}
 */
function _countFigures(docXml) {
  // Count <wp:inline> and <wp:anchor> drawing elements (each is a figure)
  const inlineMatches = docXml.match(/<wp:inline[\s>]/g) || [];
  const anchorMatches = docXml.match(/<wp:anchor[\s>]/g) || [];
  return inlineMatches.length + anchorMatches.length;
}

/**
 * Count the number of tables in the document.
 * @param {string} docXml
 * @returns {number}
 */
function _countTables(docXml) {
  return (docXml.match(/<w:tbl[\s>]/g) || []).length;
}

/**
 * Split text into sentences (approximation).
 * Splits on period, question mark, or exclamation mark followed by a space
 * or end of string. Handles common abbreviations.
 *
 * @param {string} text
 * @returns {string[]}
 */
function _splitSentences(text) {
  if (!text || !text.trim()) return [];

  // Replace common abbreviations to avoid false splits
  let processed = text
    .replace(/\b(Dr|Mr|Mrs|Ms|Prof|Inc|Ltd|Jr|Sr|vs|etc|Fig|No|Vol|pp|ed|eds)\./gi, '$1\u0000')
    .replace(/\b(e\.g|i\.e|et al)\./gi, (m) => m.replace(/\./g, '\u0000'))
    .replace(/\d+\.\d+/g, (m) => m.replace('.', '\u0000')); // decimal numbers

  // Split on sentence-ending punctuation followed by space or end
  const raw = processed.split(/(?<=[.!?])\s+/);

  // Restore abbreviation dots
  return raw
    .map(s => s.replace(/\u0000/g, '.').trim())
    .filter(s => s.length > 0);
}

/**
 * Count words in a string.
 * @param {string} text
 * @returns {number}
 */
function _countWords(text) {
  if (!text || !text.trim()) return 0;
  return text.trim().split(/\s+/).length;
}

// ============================================================================
// QUALITY CLASS
// ============================================================================

class Quality {

  /**
   * Manuscript linter. Checks for:
   * - Unclosed parentheses in paragraphs
   * - Mismatched quotation marks (odd number of quotes)
   * - "Figure N" references where N > actual figure count
   * - "Table N" references where N > actual table count
   * - Repeated words ("the the", "is is")
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {Array<{type: string, message: string, paraId: string|null, severity: string}>}
   */
  static lint(ws) {
    const issues = [];
    const paragraphs = _extractParagraphs(ws);
    const figureCount = _countFigures(ws.docXml);
    const tableCount = _countTables(ws.docXml);

    for (const p of paragraphs) {
      const text = p.text.trim();
      if (!text) continue;

      // --- Unclosed parentheses ---
      let parenDepth = 0;
      for (const ch of text) {
        if (ch === '(') parenDepth++;
        if (ch === ')') parenDepth--;
      }
      if (parenDepth !== 0) {
        issues.push({
          type: 'unclosed-paren',
          message: parenDepth > 0
            ? `Unclosed parenthesis (${parenDepth} open without close)`
            : `Extra closing parenthesis (${-parenDepth} close without open)`,
          paraId: p.paraId,
          severity: 'error',
        });
      }

      // --- Mismatched quotation marks ---
      // Count straight double quotes
      const doubleQuotes = (text.match(/"/g) || []).length;
      if (doubleQuotes % 2 !== 0) {
        issues.push({
          type: 'mismatched-quotes',
          message: `Odd number of double quotation marks (${doubleQuotes})`,
          paraId: p.paraId,
          severity: 'warning',
        });
      }

      // Count curly double quotes: left and right should match
      const leftCurly = (text.match(/\u201C/g) || []).length;
      const rightCurly = (text.match(/\u201D/g) || []).length;
      if (leftCurly !== rightCurly) {
        issues.push({
          type: 'mismatched-quotes',
          message: `Mismatched curly quotes: ${leftCurly} opening vs ${rightCurly} closing`,
          paraId: p.paraId,
          severity: 'warning',
        });
      }

      // --- Figure N references exceeding figure count ---
      const figRefs = text.matchAll(/\bFigure\s+(\d+)/gi);
      for (const match of figRefs) {
        const refNum = parseInt(match[1], 10);
        if (refNum > figureCount) {
          issues.push({
            type: 'invalid-figure-ref',
            message: `Reference to Figure ${refNum} but document only has ${figureCount} figure${figureCount !== 1 ? 's' : ''}`,
            paraId: p.paraId,
            severity: 'error',
          });
        }
      }

      // --- Table N references exceeding table count ---
      const tableRefs = text.matchAll(/\bTable\s+(\d+)/gi);
      for (const match of tableRefs) {
        const refNum = parseInt(match[1], 10);
        if (refNum > tableCount) {
          issues.push({
            type: 'invalid-table-ref',
            message: `Reference to Table ${refNum} but document only has ${tableCount} table${tableCount !== 1 ? 's' : ''}`,
            paraId: p.paraId,
            severity: 'error',
          });
        }
      }

      // --- Repeated words ---
      // Match two identical adjacent words (case-insensitive)
      const repeatedPattern = /\b(\w{2,})\s+\1\b/gi;
      let repeatMatch;
      while ((repeatMatch = repeatedPattern.exec(text)) !== null) {
        const word = repeatMatch[1].toLowerCase();
        // Skip known valid repetitions
        if (['had', 'that'].includes(word)) continue;
        issues.push({
          type: 'repeated-word',
          message: `Repeated word: "${word} ${word}"`,
          paraId: p.paraId,
          severity: 'warning',
        });
      }
    }

    return issues;
  }

  /**
   * Load a JSON stats file and verify numbers in the document match.
   * Finds all numbers in the document text and checks if key numbers
   * from the stats file appear correctly.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} statsPath - Path to a JSON file with expected numbers
   * @returns {{ matches: Array<{number: string, location: string, source: string}>, mismatches: Array<{expected: string, found: string|null, paraId: string|null}> }}
   */
  static checkNumbers(ws, statsPath) {
    const statsRaw = fs.readFileSync(statsPath, 'utf-8');
    const stats = JSON.parse(statsRaw);
    const paragraphs = _extractParagraphs(ws);

    const matches = [];
    const mismatches = [];

    // Flatten the stats object into key-value pairs where values are numbers/strings
    const expectedEntries = [];
    for (const [key, value] of Object.entries(stats)) {
      if (value !== null && value !== undefined) {
        expectedEntries.push({ source: key, value: String(value) });
      }
    }

    // For each expected number, search the document
    for (const entry of expectedEntries) {
      const numStr = entry.value;
      // Skip non-numeric values
      if (!/[\d]/.test(numStr)) continue;

      let found = false;
      for (const p of paragraphs) {
        if (p.text.includes(numStr)) {
          matches.push({
            number: numStr,
            location: p.text.slice(0, 80),
            source: entry.source,
          });
          found = true;
          break;
        }
      }

      if (!found) {
        // Check if the number appears in a different format (e.g., with/without commas)
        const plainNum = numStr.replace(/,/g, '');
        const commaNum = Number(plainNum).toLocaleString('en-US');
        let altFound = false;
        let altFoundText = null;
        let altParaId = null;

        for (const p of paragraphs) {
          if (p.text.includes(plainNum) || p.text.includes(commaNum)) {
            altFound = true;
            altFoundText = p.text.includes(plainNum) ? plainNum : commaNum;
            altParaId = p.paraId;
            break;
          }
        }

        if (altFound) {
          mismatches.push({
            expected: numStr,
            found: altFoundText,
            paraId: altParaId,
          });
        } else {
          mismatches.push({
            expected: numStr,
            found: null,
            paraId: null,
          });
        }
      }
    }

    return { matches, mismatches };
  }

  /**
   * Detect passive voice patterns in the document.
   * Looks for: "was/were/been/being/is/are + past participle"
   * Common academic patterns: "was found", "were collected", "is discussed"
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {Array<{sentence: string, paraId: string|null, suggestion: string}>}
   */
  static passiveVoice(ws) {
    const results = [];
    const paragraphs = _extractParagraphs(ws);

    // Passive voice pattern: auxiliary + optional adverb + past participle
    // Past participles commonly end in -ed, -en, -t, -wn, -ne
    const passiveRe = /\b(was|were|been|being|is|are|has been|have been|had been|will be|would be|could be|should be|might be|may be|can be)\s+(\w+ly\s+)?(\w+(ed|en|wn|ne|t))\b/gi;

    for (const p of paragraphs) {
      const text = p.text.trim();
      if (!text) continue;

      const sentences = _splitSentences(text);
      for (const sentence of sentences) {
        const matchIterator = sentence.matchAll(passiveRe);
        for (const match of matchIterator) {
          const auxiliary = match[1];
          const participle = match[3];

          results.push({
            sentence: sentence.trim(),
            paraId: p.paraId,
            suggestion: `Consider active voice: "${auxiliary} ${participle}" is passive`,
          });
          break; // One hit per sentence is enough
        }
      }
    }

    return results;
  }

  /**
   * Flag sentences over a maximum word count.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {object} [opts] - Options
   * @param {number} [opts.max=40] - Maximum words per sentence
   * @returns {Array<{sentence: string, paraId: string|null, wordCount: number}>}
   */
  static sentenceLength(ws, opts = {}) {
    const max = (opts && opts.max) || 40;
    const results = [];
    const paragraphs = _extractParagraphs(ws);

    for (const p of paragraphs) {
      const text = p.text.trim();
      if (!text) continue;

      const sentences = _splitSentences(text);
      for (const sentence of sentences) {
        const wordCount = _countWords(sentence);
        if (wordCount > max) {
          results.push({
            sentence: sentence.trim(),
            paraId: p.paraId,
            wordCount,
          });
        }
      }
    }

    return results;
  }

  /**
   * Compute Flesch-Kincaid readability metrics.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {{ fleschKincaid: number, readingEase: number, avgSentenceLength: number, avgSyllables: number }}
   */
  static readability(ws) {
    const paragraphs = _extractParagraphs(ws);

    let totalWords = 0;
    let totalSentences = 0;
    let totalSyllables = 0;

    for (const p of paragraphs) {
      const text = p.text.trim();
      if (!text) continue;

      const sentences = _splitSentences(text);
      totalSentences += sentences.length;

      const words = text.split(/\s+/).filter(w => w.length > 0);
      totalWords += words.length;

      for (const word of words) {
        totalSyllables += Quality._countSyllables(word);
      }
    }

    if (totalWords === 0 || totalSentences === 0) {
      return {
        fleschKincaid: 0,
        readingEase: 0,
        avgSentenceLength: 0,
        avgSyllables: 0,
      };
    }

    const avgSentenceLength = totalWords / totalSentences;
    const avgSyllables = totalSyllables / totalWords;

    // Flesch-Kincaid Grade Level
    const fleschKincaid = 0.39 * avgSentenceLength + 11.8 * avgSyllables - 15.59;

    // Flesch Reading Ease
    const readingEase = 206.835 - 1.015 * avgSentenceLength - 84.6 * avgSyllables;

    return {
      fleschKincaid: Math.round(fleschKincaid * 10) / 10,
      readingEase: Math.round(readingEase * 10) / 10,
      avgSentenceLength: Math.round(avgSentenceLength * 10) / 10,
      avgSyllables: Math.round(avgSyllables * 10) / 10,
    };
  }

  /**
   * Count syllables in a word (English approximation).
   * Based on the common heuristic: count vowel groups, adjust for
   * silent-e and common suffixes.
   *
   * @param {string} word
   * @returns {number}
   */
  static _countSyllables(word) {
    if (!word) return 0;

    // Strip non-alpha characters
    word = word.toLowerCase().replace(/[^a-z]/g, '');
    if (word.length === 0) return 0;
    if (word.length <= 2) return 1;

    // Count vowel groups
    const vowelGroups = word.match(/[aeiouy]+/g);
    if (!vowelGroups) return 1;

    let count = vowelGroups.length;

    // Subtract silent -e at end (but not -le which is syllabic)
    if (word.endsWith('e') && !word.endsWith('le')) {
      count--;
    }

    // -ed ending: usually silent unless preceded by d or t
    if (word.endsWith('ed') && word.length > 3) {
      const beforeEd = word[word.length - 3];
      if (beforeEd !== 'd' && beforeEd !== 't') {
        count--;
      }
    }

    // Ensure at least 1 syllable
    return Math.max(1, count);
  }
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = { Quality };
