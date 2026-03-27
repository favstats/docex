/**
 * redact.js -- Collaboration and privacy tools for docex
 *
 * Provides redaction and unredaction of sensitive text.
 * Stores redaction mappings in a custom XML part inside the .docx
 * so they survive file moves and renames.
 *
 * All methods operate on a Workspace object.
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const xml = require('./xml');

// ============================================================================
// CONSTANTS
// ============================================================================

/**
 * Custom XML namespace for docex redaction data.
 */
const REDACT_NS = 'http://docex.dev/redaction/2026';

/**
 * Relative path inside the .docx zip for the redaction mapping.
 */
const REDACT_XML_PATH = 'customXml/docex-redactions.xml';

/**
 * Content type for custom XML.
 */
const CUSTOM_XML_CONTENT_TYPE = 'application/xml';

// ============================================================================
// REDACT CLASS
// ============================================================================

class Redact {

  /**
   * Replace all occurrences of searchText with replacement in the document.
   * Stores the mapping in a custom XML part for later unredact().
   *
   * @param {object} ws - Workspace
   * @param {string} searchText - Text to redact
   * @param {string} [replacement='[REDACTED]'] - Replacement string
   * @param {object} [opts={}] - Options
   * @param {boolean} [opts.tracked=false] - Whether to make redactions as tracked changes
   * @returns {{ count: number, searchText: string, replacement: string }}
   */
  static redact(ws, searchText, replacement, opts = {}) {
    if (typeof replacement === 'object' && replacement !== null) {
      opts = replacement;
      replacement = '[REDACTED]';
    }
    if (!replacement) replacement = '[REDACTED]';

    const escapedSearch = xml.escapeXml(searchText);
    const escapedReplacement = xml.escapeXml(replacement);

    // Count and replace in document.xml
    let docXml = ws.docXml;
    let count = 0;

    // Replace in w:t elements only (not in XML attributes or metadata)
    const tRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    docXml = docXml.replace(tRe, (full, textContent) => {
      if (textContent.includes(escapedSearch)) {
        const occurrences = textContent.split(escapedSearch).length - 1;
        count += occurrences;
        const newText = textContent.split(escapedSearch).join(escapedReplacement);
        return full.replace(textContent, newText);
      }
      return full;
    });

    // Also check for multi-run split text (text split across runs)
    // We do a pass to catch the simple cases. Complex multi-run splits
    // are handled via paragraph-level search.
    if (count === 0) {
      // Try paragraph-level text matching
      const paragraphs = xml.findParagraphs(docXml);
      for (const p of paragraphs) {
        const pText = xml.extractText(p.xml);
        const decodedSearch = searchText;
        if (pText.includes(decodedSearch)) {
          // Replace in all w:t elements within this paragraph
          let newPXml = p.xml;
          const runs = [];
          const runRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
          let m;
          while ((m = runRe.exec(p.xml)) !== null) {
            runs.push({ full: m[0], text: m[1], index: m.index });
          }

          if (runs.length > 0) {
            const fullText = runs.map(r => r.text).join('');
            const decodedFull = xml.decodeXml(fullText);
            if (decodedFull.includes(decodedSearch)) {
              const newFull = decodedFull.split(decodedSearch).join(replacement);
              count += decodedFull.split(decodedSearch).length - 1;
              // Put all text in the first run, empty the rest
              for (let i = runs.length - 1; i >= 0; i--) {
                const rep = i === 0
                  ? `<w:t xml:space="preserve">${xml.escapeXml(newFull)}</w:t>`
                  : '<w:t xml:space="preserve"></w:t>';
                newPXml = newPXml.slice(0, runs[i].index) + rep + newPXml.slice(runs[i].index + runs[i].full.length);
              }
              docXml = docXml.slice(0, p.start) + newPXml + docXml.slice(p.end);
            }
          }
        }
      }
    }

    ws.docXml = docXml;

    // Store the mapping in customXml
    Redact._storeMapping(ws, searchText, replacement, count);

    return { count, searchText, replacement };
  }

  /**
   * Restore all redacted text from the stored mapping.
   *
   * @param {object} ws - Workspace
   * @returns {{ count: number }}
   */
  static unredact(ws) {
    const mappings = Redact._loadMappings(ws);
    if (mappings.length === 0) {
      return { count: 0 };
    }

    let docXml = ws.docXml;
    let totalCount = 0;

    for (const mapping of mappings) {
      const escapedOriginal = xml.escapeXml(mapping.original);
      const escapedReplacement = xml.escapeXml(mapping.replacement);

      // Replace the replacement text back to original in w:t elements
      const tRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
      docXml = docXml.replace(tRe, (full, textContent) => {
        if (textContent.includes(escapedReplacement)) {
          const occurrences = textContent.split(escapedReplacement).length - 1;
          totalCount += occurrences;
          const newText = textContent.split(escapedReplacement).join(escapedOriginal);
          return full.replace(textContent, newText);
        }
        return full;
      });
    }

    ws.docXml = docXml;

    // Clear the mappings
    Redact._clearMappings(ws);

    return { count: totalCount };
  }

  /**
   * List all active redactions.
   *
   * @param {object} ws - Workspace
   * @returns {Array<{original: string, replacement: string, count: number}>}
   */
  static listRedactions(ws) {
    return Redact._loadMappings(ws);
  }

  // --------------------------------------------------------------------------
  // Internal helpers: custom XML storage
  // --------------------------------------------------------------------------

  /**
   * Store a redaction mapping in the custom XML part.
   * If the custom XML part doesn't exist, create it.
   * @param {object} ws - Workspace
   * @param {string} original - Original text
   * @param {string} replacement - Replacement text
   * @param {number} count - Number of occurrences redacted
   * @private
   */
  static _storeMapping(ws, original, replacement, count) {
    const mappings = Redact._loadMappings(ws);

    // Check if this mapping already exists
    const existing = mappings.find(m => m.original === original && m.replacement === replacement);
    if (existing) {
      existing.count += count;
    } else {
      mappings.push({ original, replacement, count });
    }

    Redact._saveMappings(ws, mappings);
  }

  /**
   * Load redaction mappings from the custom XML part.
   * @param {object} ws - Workspace
   * @returns {Array<{original: string, replacement: string, count: number}>}
   * @private
   */
  static _loadMappings(ws) {
    const filePath = path.join(ws.tmpDir, REDACT_XML_PATH);
    if (!fs.existsSync(filePath)) {
      return [];
    }

    const content = fs.readFileSync(filePath, 'utf-8');
    const mappings = [];
    const entryRe = /<redaction\s+original="([^"]*?)"\s+replacement="([^"]*?)"\s+count="(\d+)"\s*\/>/g;
    let m;
    while ((m = entryRe.exec(content)) !== null) {
      mappings.push({
        original: xml.decodeXml(m[1]),
        replacement: xml.decodeXml(m[2]),
        count: parseInt(m[3], 10),
      });
    }
    return mappings;
  }

  /**
   * Save redaction mappings to the custom XML part.
   * @param {object} ws - Workspace
   * @param {Array} mappings
   * @private
   */
  static _saveMappings(ws, mappings) {
    const entries = mappings.map(m =>
      `  <redaction original="${xml.escapeXml(m.original)}" replacement="${xml.escapeXml(m.replacement)}" count="${m.count}"/>`
    ).join('\n');

    const content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + `<redactions xmlns="${REDACT_NS}">\n`
      + entries + '\n'
      + '</redactions>';

    // Ensure customXml directory exists
    const dir = path.join(ws.tmpDir, 'customXml');
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(path.join(ws.tmpDir, REDACT_XML_PATH), content, 'utf-8');

    // Ensure [Content_Types].xml includes customXml
    Redact._ensureContentType(ws);
  }

  /**
   * Clear all redaction mappings.
   * @param {object} ws - Workspace
   * @private
   */
  static _clearMappings(ws) {
    const filePath = path.join(ws.tmpDir, REDACT_XML_PATH);
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  }

  /**
   * Ensure the [Content_Types].xml includes an Override for our custom XML.
   * @param {object} ws - Workspace
   * @private
   */
  static _ensureContentType(ws) {
    let ct = ws.contentTypesXml;
    const partName = '/' + REDACT_XML_PATH;
    if (!ct.includes(partName)) {
      // Add an Override before </Types>
      const override = `<Override PartName="${partName}" ContentType="${CUSTOM_XML_CONTENT_TYPE}"/>`;
      ct = ct.replace('</Types>', override + '</Types>');
      ws.contentTypesXml = ct;
    }
  }
}

module.exports = { Redact };
