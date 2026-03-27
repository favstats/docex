/**
 * macros.js -- Variable definition and expansion for docex
 *
 * Provides template variable substitution: define variables, find all
 * {{VAR_NAME}} patterns in the document, and expand them to their values.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// MACROS
// ============================================================================

class Macros {

  /**
   * Define a variable in the workspace's in-memory variable store.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} name - Variable name (without braces), e.g. "NUM_ADS"
   * @param {string} value - Variable value, e.g. "268,635"
   */
  static define(ws, name, value) {
    if (!ws._docexVariables) {
      ws._docexVariables = {};
    }
    ws._docexVariables[name] = value;
  }

  /**
   * Expand all {{VAR_NAME}} patterns in the document text.
   * Replaces them with the corresponding values from the variables map.
   *
   * Variables can be:
   *   1. Pre-defined via Macros.define()
   *   2. Passed directly as the variables argument
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {object} [variables] - Map of variable names to values
   * @returns {number} Count of expansions made
   */
  static expand(ws, variables = {}) {
    // Merge pre-defined variables with passed variables
    const allVars = { ...(ws._docexVariables || {}), ...variables };

    if (Object.keys(allVars).length === 0) {
      return 0;
    }

    let docXml = ws.docXml;
    let expandCount = 0;

    // Find all paragraphs and process each one
    const paragraphs = xml.findParagraphs(docXml);

    // Process in reverse order to preserve offsets
    for (let i = paragraphs.length - 1; i >= 0; i--) {
      const p = paragraphs[i];
      const text = xml.extractTextDecoded(p.xml);

      // Check if this paragraph contains any {{VAR}} patterns
      if (!text.includes('{{')) continue;

      let newParaXml = p.xml;
      let paraModified = false;

      // Process each variable
      for (const [name, value] of Object.entries(allVars)) {
        const pattern = `{{${name}}}`;
        const encodedPattern = xml.escapeXml(pattern);

        // The pattern might be in a single w:t element or split across runs
        // First try: single w:t element replacement
        if (newParaXml.includes(encodedPattern)) {
          newParaXml = newParaXml.split(encodedPattern).join(xml.escapeXml(value));
          expandCount++;
          paraModified = true;
        } else {
          // Try cross-run expansion: the pattern might be split across w:t elements
          // Reconstruct by looking at the decoded text
          const decodedText = xml.extractTextDecoded(newParaXml);
          if (decodedText.includes(pattern)) {
            // Use the TextMap approach: find the pattern in concatenated run texts
            const result = Macros._expandCrossRun(newParaXml, pattern, value);
            if (result.modified) {
              newParaXml = result.xml;
              expandCount++;
              paraModified = true;
            }
          }
        }
      }

      if (paraModified) {
        docXml = docXml.slice(0, p.start) + newParaXml + docXml.slice(p.end);
      }
    }

    ws.docXml = docXml;
    return expandCount;
  }

  /**
   * List all {{VAR}} patterns found in the document.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {Array<{name: string, paragraph: number, context: string}>}
   */
  static listVariables(ws) {
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    const results = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const text = xml.extractTextDecoded(p.xml);

      // Find all {{VAR}} patterns
      const varRe = /\{\{([A-Za-z_][A-Za-z0-9_]*)\}\}/g;
      let m;
      while ((m = varRe.exec(text)) !== null) {
        const name = m[1];
        // Get some context around the variable
        const start = Math.max(0, m.index - 20);
        const end = Math.min(text.length, m.index + m[0].length + 20);
        const context = (start > 0 ? '...' : '') + text.slice(start, end) + (end < text.length ? '...' : '');

        results.push({ name, paragraph: i, context });
      }
    }

    return results;
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Expand a variable pattern that spans multiple runs.
   * Concatenates run texts, finds the pattern, replaces it,
   * and rebuilds the runs.
   *
   * @param {string} paraXml - Paragraph XML
   * @param {string} pattern - The {{VAR_NAME}} pattern
   * @param {string} value - Replacement value
   * @returns {{xml: string, modified: boolean}}
   * @private
   */
  static _expandCrossRun(paraXml, pattern, value) {
    // Parse all runs and their text content
    const runs = xml.parseRuns(paraXml);
    if (runs.length === 0) return { xml: paraXml, modified: false };

    // Concatenate all run texts (decoded) to find the pattern position
    let concatenated = '';
    const runMap = []; // { runIdx, textIdx, charStart, charEnd }

    for (let ri = 0; ri < runs.length; ri++) {
      const run = runs[ri];
      for (let ti = 0; ti < run.texts.length; ti++) {
        const decoded = xml.decodeXml(run.texts[ti].text);
        runMap.push({
          runIdx: ri,
          textIdx: ti,
          charStart: concatenated.length,
          charEnd: concatenated.length + decoded.length,
        });
        concatenated += decoded;
      }
    }

    const patternStart = concatenated.indexOf(pattern);
    if (patternStart === -1) return { xml: paraXml, modified: false };

    const patternEnd = patternStart + pattern.length;

    // Find which runs are affected
    const affectedRuns = new Set();
    for (const rm of runMap) {
      if (rm.charEnd > patternStart && rm.charStart < patternEnd) {
        affectedRuns.add(rm.runIdx);
      }
    }

    // Simple approach: replace in the concatenated text and rebuild
    // the first affected run with the full text, empty the rest
    const newText = concatenated.slice(0, patternStart) + value + concatenated.slice(patternEnd);

    // Rebuild: put all text in the first run's first text element
    const firstRunIdx = Math.min(...affectedRuns);
    const firstRun = runs[firstRunIdx];
    const rPr = firstRun.rPr;

    // Build a single run with all the text
    const newRunXml = `<w:r>${rPr}<w:t xml:space="preserve">${xml.escapeXml(newText)}</w:t></w:r>`;

    // Replace all runs with a single run containing the full text
    // Find the first run start and last run end in the paragraph
    const firstRunStart = runs[0].index;
    const lastRun = runs[runs.length - 1];
    const lastRunEnd = lastRun.index + lastRun.fullMatch.length;

    // The run indices are relative to paraXml, so we can slice directly
    const newParaXml = paraXml.slice(0, firstRunStart) + newRunXml + paraXml.slice(lastRunEnd);

    return { xml: newParaXml, modified: true };
  }
}

module.exports = { Macros };
