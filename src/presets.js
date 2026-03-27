/**
 * presets.js -- Journal style presets for docex
 *
 * Applies predefined journal formatting to an OOXML document.
 * Modifies: styles.xml (fonts, sizes, spacing), document defaults,
 * section properties (margins, headers, footers).
 *
 * All methods operate on a Workspace object.
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// BUILT-IN PRESETS
// ============================================================================

const PRESETS = {
  academic: {
    font: 'Times New Roman',
    size: 12,
    spacing: 'double',
    alignment: 'justified',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
    indent: 0.5,
    headingFont: 'Times New Roman',
  },
  polcomm: {
    font: 'Times New Roman',
    size: 12,
    spacing: 'double',
    alignment: 'justified',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
    indent: 0.5,
    titlePage: true,
    runningHeader: true,
    abstractWordLimit: 250,
    wordLimit: 8000,
    headingFont: 'Times New Roman',
  },
  apa7: {
    font: 'Times New Roman',
    size: 12,
    spacing: 'double',
    alignment: 'left',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
    indent: 0.5,
    titlePage: true,
    runningHeader: true,
    abstractWordLimit: 250,
    headingStyle: 'apa',
    headingFont: 'Times New Roman',
  },
  jcmc: {
    font: 'Times New Roman',
    size: 12,
    spacing: 'double',
    alignment: 'justified',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
    indent: 0.5,
    titlePage: true,
    runningHeader: true,
    abstractWordLimit: 200,
    wordLimit: 10000,
    headingFont: 'Times New Roman',
  },
  joc: {
    font: 'Times New Roman',
    size: 12,
    spacing: 'double',
    alignment: 'justified',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
    indent: 0.5,
    titlePage: true,
    runningHeader: true,
    abstractWordLimit: 250,
    wordLimit: 8000,
    headingFont: 'Times New Roman',
  },
};

// User-defined presets (accumulated at runtime)
const customPresets = {};

// ============================================================================
// PRESETS CLASS
// ============================================================================

class Presets {

  /**
   * Apply a journal style preset to the document.
   * Modifies: styles.xml (fonts, sizes, spacing), document defaults, margins.
   *
   * @param {object} ws - Workspace
   * @param {string} presetName - Name of the preset (e.g. "polcomm", "apa7")
   * @returns {{applied: string, changes: string[]}} Summary of changes made
   */
  static apply(ws, presetName) {
    const config = Presets._resolvePreset(presetName);
    if (!config) {
      const available = Presets.list().join(', ');
      throw new Error(`Unknown preset: "${presetName}". Available: ${available}`);
    }

    const changes = [];

    // 1. Apply font and size to styles.xml
    if (config.font || config.size) {
      Presets._applyDefaultFont(ws, config.font, config.size);
      changes.push(`Default font: ${config.font || 'unchanged'} ${config.size || ''}pt`);
    }

    // 2. Apply line spacing
    if (config.spacing) {
      Presets._applyLineSpacing(ws, config.spacing);
      changes.push(`Line spacing: ${config.spacing}`);
    }

    // 3. Apply paragraph alignment
    if (config.alignment) {
      Presets._applyAlignment(ws, config.alignment);
      changes.push(`Alignment: ${config.alignment}`);
    }

    // 4. Apply margins
    if (config.margins) {
      Presets._applyMargins(ws, config.margins);
      changes.push(`Margins: ${config.margins.top}" top, ${config.margins.bottom}" bottom, ${config.margins.left}" left, ${config.margins.right}" right`);
    }

    // 5. Apply paragraph indent
    if (config.indent) {
      Presets._applyFirstLineIndent(ws, config.indent);
      changes.push(`First-line indent: ${config.indent}"`);
    }

    return { applied: presetName, changes };
  }

  /**
   * Return all available preset names.
   *
   * @returns {string[]} Array of preset names
   */
  static list() {
    return [...Object.keys(PRESETS), ...Object.keys(customPresets)];
  }

  /**
   * Register a custom preset.
   *
   * @param {string} name - Preset name
   * @param {object} config - Preset configuration
   */
  static define(name, config) {
    customPresets[name] = config;
  }

  /**
   * Get the configuration for a preset.
   *
   * @param {string} name - Preset name
   * @returns {object|null} The preset config or null
   */
  static get(name) {
    return Presets._resolvePreset(name);
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Resolve a preset name to its configuration.
   *
   * @param {string} name - Preset name
   * @returns {object|null}
   * @private
   */
  static _resolvePreset(name) {
    return PRESETS[name] || customPresets[name] || null;
  }

  /**
   * Apply the default font and size to styles.xml docDefaults.
   *
   * @param {object} ws - Workspace
   * @param {string} font - Font name
   * @param {number} size - Font size in points
   * @private
   */
  static _applyDefaultFont(ws, font, size) {
    let stylesXml = ws.stylesXml;
    if (!stylesXml) return;

    const sizeHalfPt = size ? size * 2 : 24; // OOXML uses half-points

    // Update or create docDefaults rPrDefault
    if (stylesXml.includes('<w:rPrDefault>')) {
      // Update existing rPrDefault
      if (font) {
        // Replace rFonts in rPrDefault
        stylesXml = stylesXml.replace(
          /(<w:rPrDefault>[\s\S]*?<w:rFonts\b)[^/]*(\/?>)/,
          `$1 w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"$2`
        );
        // If no rFonts exists, add it
        if (!/<w:rPrDefault>[\s\S]*?<w:rFonts\b/.test(stylesXml)) {
          stylesXml = stylesXml.replace(
            /(<w:rPrDefault>\s*<w:rPr>)/,
            `$1<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`
          );
        }
      }
      if (size) {
        // Replace sz in rPrDefault
        if (stylesXml.match(/<w:rPrDefault>[\s\S]*?<w:sz\b/)) {
          stylesXml = stylesXml.replace(
            /(<w:rPrDefault>[\s\S]*?<w:sz\s+w:val=")(\d+)(")/,
            `$1${sizeHalfPt}$3`
          );
          stylesXml = stylesXml.replace(
            /(<w:rPrDefault>[\s\S]*?<w:szCs\s+w:val=")(\d+)(")/,
            `$1${sizeHalfPt}$3`
          );
        } else {
          stylesXml = stylesXml.replace(
            /(<w:rPrDefault>\s*<w:rPr>)/,
            `$1<w:sz w:val="${sizeHalfPt}"/><w:szCs w:val="${sizeHalfPt}"/>`
          );
        }
      }
    } else if (stylesXml.includes('<w:docDefaults>')) {
      // docDefaults exists but no rPrDefault
      const rPrDefault =
        '<w:rPrDefault><w:rPr>'
        + (font ? `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>` : '')
        + (size ? `<w:sz w:val="${sizeHalfPt}"/><w:szCs w:val="${sizeHalfPt}"/>` : '')
        + '</w:rPr></w:rPrDefault>';
      stylesXml = stylesXml.replace('<w:docDefaults>', '<w:docDefaults>' + rPrDefault);
    } else {
      // No docDefaults at all -- add it
      const docDefaults =
        '<w:docDefaults>'
        + '<w:rPrDefault><w:rPr>'
        + (font ? `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>` : '')
        + (size ? `<w:sz w:val="${sizeHalfPt}"/><w:szCs w:val="${sizeHalfPt}"/>` : '')
        + '</w:rPr></w:rPrDefault>'
        + '</w:docDefaults>';
      // Insert after the opening <w:styles...> tag
      stylesXml = stylesXml.replace(/(<w:styles[^>]*>)/, `$1${docDefaults}`);
    }

    ws.stylesXml = stylesXml;
  }

  /**
   * Apply line spacing to the document's Normal style.
   *
   * @param {object} ws - Workspace
   * @param {string} spacing - "single", "1.5", or "double"
   * @private
   */
  static _applyLineSpacing(ws, spacing) {
    let stylesXml = ws.stylesXml;
    if (!stylesXml) return;

    // OOXML line spacing values (in 240ths of a line)
    const spacingMap = {
      'single': 240,
      '1.15': 276,
      '1.5': 360,
      'double': 480,
    };
    const lineVal = spacingMap[spacing] || 480;

    const spacingXml = `<w:spacing w:line="${lineVal}" w:lineRule="auto"/>`;

    // Try to update the Normal style's pPr spacing
    // Find the Normal style
    const normalStyleRe = /(<w:style\s+w:type="paragraph"\s+w:default="1"[^>]*>)([\s\S]*?)(<\/w:style>)/;
    const normalMatch = stylesXml.match(normalStyleRe);

    if (normalMatch) {
      let styleContent = normalMatch[2];

      if (styleContent.includes('<w:spacing')) {
        // Update existing spacing in pPr
        styleContent = styleContent.replace(
          /<w:spacing[^/]*\/>/,
          spacingXml
        );
      } else if (styleContent.includes('<w:pPr>')) {
        // Add spacing to existing pPr
        styleContent = styleContent.replace('<w:pPr>', '<w:pPr>' + spacingXml);
      } else {
        // Add pPr with spacing
        styleContent = `<w:pPr>${spacingXml}</w:pPr>` + styleContent;
      }

      stylesXml = stylesXml.replace(normalStyleRe, normalMatch[1] + styleContent + normalMatch[3]);
    } else {
      // Also update pPrDefault if present
      if (stylesXml.includes('<w:pPrDefault>')) {
        if (stylesXml.match(/<w:pPrDefault>[\s\S]*?<w:spacing/)) {
          stylesXml = stylesXml.replace(
            /(<w:pPrDefault>[\s\S]*?)<w:spacing[^/]*\/>/,
            `$1${spacingXml}`
          );
        } else if (stylesXml.includes('<w:pPr>', stylesXml.indexOf('<w:pPrDefault>'))) {
          stylesXml = stylesXml.replace(
            /(<w:pPrDefault>\s*<w:pPr>)/,
            `$1${spacingXml}`
          );
        }
      }
    }

    ws.stylesXml = stylesXml;
  }

  /**
   * Apply paragraph alignment to the Normal style.
   *
   * @param {object} ws - Workspace
   * @param {string} alignment - "left", "center", "right", "justified"
   * @private
   */
  static _applyAlignment(ws, alignment) {
    let stylesXml = ws.stylesXml;
    if (!stylesXml) return;

    // OOXML alignment values
    const alignMap = {
      'left': 'start',
      'center': 'center',
      'right': 'end',
      'justified': 'both',
      'justify': 'both',
      'both': 'both',
    };
    const alignVal = alignMap[alignment] || 'both';
    const alignXml = `<w:jc w:val="${alignVal}"/>`;

    // Update Normal style
    const normalStyleRe = /(<w:style\s+w:type="paragraph"\s+w:default="1"[^>]*>)([\s\S]*?)(<\/w:style>)/;
    const normalMatch = stylesXml.match(normalStyleRe);

    if (normalMatch) {
      let styleContent = normalMatch[2];

      if (styleContent.includes('<w:jc')) {
        styleContent = styleContent.replace(/<w:jc[^/]*\/>/, alignXml);
      } else if (styleContent.includes('<w:pPr>')) {
        styleContent = styleContent.replace('<w:pPr>', '<w:pPr>' + alignXml);
      } else {
        styleContent = `<w:pPr>${alignXml}</w:pPr>` + styleContent;
      }

      stylesXml = stylesXml.replace(normalStyleRe, normalMatch[1] + styleContent + normalMatch[3]);
    }

    ws.stylesXml = stylesXml;
  }

  /**
   * Apply page margins via section properties in document.xml.
   *
   * @param {object} ws - Workspace
   * @param {object} margins - { top, bottom, left, right } in inches
   * @private
   */
  static _applyMargins(ws, margins) {
    let docXml = ws.docXml;

    // OOXML margins are in twips (1 inch = 1440 twips)
    const top = Math.round((margins.top || 1) * 1440);
    const bottom = Math.round((margins.bottom || 1) * 1440);
    const left = Math.round((margins.left || 1) * 1440);
    const right = Math.round((margins.right || 1) * 1440);

    const marginXml = `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="720" w:footer="720" w:gutter="0"/>`;

    // Find sectPr and update pgMar
    if (docXml.includes('<w:pgMar')) {
      docXml = docXml.replace(/<w:pgMar[^/]*\/>/, marginXml);
    } else if (docXml.includes('<w:sectPr')) {
      docXml = docXml.replace(/<w:sectPr([^>]*)>/, `<w:sectPr$1>${marginXml}`);
    }

    ws.docXml = docXml;
  }

  /**
   * Apply first-line indent to the Normal paragraph style.
   *
   * @param {object} ws - Workspace
   * @param {number} inches - Indent in inches
   * @private
   */
  static _applyFirstLineIndent(ws, inches) {
    let stylesXml = ws.stylesXml;
    if (!stylesXml) return;

    const twips = Math.round(inches * 1440);
    const indentXml = `<w:ind w:firstLine="${twips}"/>`;

    // Update Normal style
    const normalStyleRe = /(<w:style\s+w:type="paragraph"\s+w:default="1"[^>]*>)([\s\S]*?)(<\/w:style>)/;
    const normalMatch = stylesXml.match(normalStyleRe);

    if (normalMatch) {
      let styleContent = normalMatch[2];

      if (styleContent.includes('<w:ind')) {
        styleContent = styleContent.replace(/<w:ind[^/]*\/>/, indentXml);
      } else if (styleContent.includes('<w:pPr>')) {
        styleContent = styleContent.replace('<w:pPr>', '<w:pPr>' + indentXml);
      } else {
        styleContent = `<w:pPr>${indentXml}</w:pPr>` + styleContent;
      }

      stylesXml = stylesXml.replace(normalStyleRe, normalMatch[1] + styleContent + normalMatch[3]);
    }

    ws.stylesXml = stylesXml;
  }

  // --------------------------------------------------------------------------
  // Compare styles (v0.4.3)
  // --------------------------------------------------------------------------

  /**
   * Preview what a preset would change without applying it.
   * Compares the current document state against the preset config.
   *
   * @param {object} ws - Workspace
   * @param {string} presetName - Preset name to compare against
   * @returns {{ changes: string[] }} List of human-readable change descriptions
   */
  static compareStyles(ws, presetName) {
    const config = Presets._resolvePreset(presetName);
    if (!config) {
      const available = Presets.list().join(', ');
      throw new Error(`Unknown preset: "${presetName}". Available: ${available}`);
    }

    const changes = [];
    const stylesXml = ws.stylesXml || '';
    const docXml = ws.docXml || '';

    // 1. Check font
    if (config.font) {
      const currentFont = Presets._detectCurrentFont(stylesXml);
      if (currentFont && currentFont !== config.font) {
        // Count paragraphs that would be affected
        const paraCount = (docXml.match(/<w:p[\s>]/g) || []).length;
        changes.push(`${paraCount} paragraphs: font ${currentFont} -> ${config.font}`);
      } else if (!currentFont) {
        const paraCount = (docXml.match(/<w:p[\s>]/g) || []).length;
        changes.push(`${paraCount} paragraphs: font (default) -> ${config.font}`);
      }
    }

    // 2. Check size
    if (config.size) {
      const currentSize = Presets._detectCurrentSize(stylesXml);
      if (currentSize && currentSize !== config.size) {
        changes.push(`font size: ${currentSize}pt -> ${config.size}pt`);
      } else if (!currentSize) {
        changes.push(`font size: (default) -> ${config.size}pt`);
      }
    }

    // 3. Check spacing
    if (config.spacing) {
      const currentSpacing = Presets._detectCurrentSpacing(stylesXml);
      if (currentSpacing && currentSpacing !== config.spacing) {
        changes.push(`line spacing: ${currentSpacing} -> ${config.spacing}`);
      } else if (!currentSpacing) {
        changes.push(`line spacing: (default) -> ${config.spacing}`);
      }
    }

    // 4. Check alignment
    if (config.alignment) {
      const currentAlign = Presets._detectCurrentAlignment(stylesXml);
      if (currentAlign && currentAlign !== config.alignment) {
        changes.push(`alignment: ${currentAlign} -> ${config.alignment}`);
      }
    }

    // 5. Check margins
    if (config.margins) {
      const currentMargins = Presets._detectCurrentMargins(docXml);
      if (currentMargins) {
        const marginParts = [];
        if (currentMargins.top !== config.margins.top) marginParts.push(`top ${currentMargins.top}" -> ${config.margins.top}"`);
        if (currentMargins.bottom !== config.margins.bottom) marginParts.push(`bottom ${currentMargins.bottom}" -> ${config.margins.bottom}"`);
        if (currentMargins.left !== config.margins.left) marginParts.push(`left ${currentMargins.left}" -> ${config.margins.left}"`);
        if (currentMargins.right !== config.margins.right) marginParts.push(`right ${currentMargins.right}" -> ${config.margins.right}"`);
        if (marginParts.length > 0) {
          changes.push(`margins: ${marginParts.join(', ')}`);
        }
      }
    }

    // 6. Check indent
    if (config.indent) {
      const currentIndent = Presets._detectCurrentIndent(stylesXml);
      if (currentIndent !== null && currentIndent !== config.indent) {
        changes.push(`first-line indent: ${currentIndent}" -> ${config.indent}"`);
      } else if (currentIndent === null) {
        changes.push(`first-line indent: none -> ${config.indent}"`);
      }
    }

    return { changes };
  }

  // --------------------------------------------------------------------------
  // Detection helpers for compareStyles
  // --------------------------------------------------------------------------

  /**
   * Detect the current default font from styles.xml.
   * @param {string} stylesXml
   * @returns {string|null}
   * @private
   */
  static _detectCurrentFont(stylesXml) {
    // Look in rPrDefault for rFonts
    const m = stylesXml.match(/<w:rPrDefault>[\s\S]*?<w:rFonts[^>]*w:ascii="([^"]+)"/);
    return m ? m[1] : null;
  }

  /**
   * Detect the current default font size from styles.xml.
   * @param {string} stylesXml
   * @returns {number|null} Size in points
   * @private
   */
  static _detectCurrentSize(stylesXml) {
    const m = stylesXml.match(/<w:rPrDefault>[\s\S]*?<w:sz\s+w:val="(\d+)"/);
    return m ? parseInt(m[1], 10) / 2 : null;
  }

  /**
   * Detect the current line spacing from the Normal style.
   * @param {string} stylesXml
   * @returns {string|null} 'single', '1.15', '1.5', or 'double'
   * @private
   */
  static _detectCurrentSpacing(stylesXml) {
    // Look in the Normal style
    const normalStyleRe = /(<w:style\s+w:type="paragraph"\s+w:default="1"[^>]*>)([\s\S]*?)(<\/w:style>)/;
    const normalMatch = stylesXml.match(normalStyleRe);
    if (!normalMatch) return null;

    const spacingMatch = normalMatch[2].match(/<w:spacing[^>]*w:line="(\d+)"/);
    if (!spacingMatch) return null;

    const val = parseInt(spacingMatch[1], 10);
    const reverseMap = { 240: 'single', 276: '1.15', 360: '1.5', 480: 'double' };
    return reverseMap[val] || `${(val / 240).toFixed(2)}x`;
  }

  /**
   * Detect the current paragraph alignment from the Normal style.
   * @param {string} stylesXml
   * @returns {string|null}
   * @private
   */
  static _detectCurrentAlignment(stylesXml) {
    const normalStyleRe = /(<w:style\s+w:type="paragraph"\s+w:default="1"[^>]*>)([\s\S]*?)(<\/w:style>)/;
    const normalMatch = stylesXml.match(normalStyleRe);
    if (!normalMatch) return null;

    const jcMatch = normalMatch[2].match(/<w:jc\s+w:val="([^"]+)"/);
    if (!jcMatch) return null;

    const reverseMap = { 'start': 'left', 'center': 'center', 'end': 'right', 'both': 'justified' };
    return reverseMap[jcMatch[1]] || jcMatch[1];
  }

  /**
   * Detect current page margins from document.xml.
   * @param {string} docXml
   * @returns {{top: number, bottom: number, left: number, right: number}|null}
   * @private
   */
  static _detectCurrentMargins(docXml) {
    const m = docXml.match(/<w:pgMar\s+([^>]+)\/?>/);
    if (!m) return null;

    const attrs = m[1];
    const top = (attrs.match(/w:top="(\d+)"/) || [])[1];
    const bottom = (attrs.match(/w:bottom="(\d+)"/) || [])[1];
    const left = (attrs.match(/w:left="(\d+)"/) || [])[1];
    const right = (attrs.match(/w:right="(\d+)"/) || [])[1];

    if (!top || !bottom || !left || !right) return null;

    return {
      top: Math.round(parseInt(top, 10) / 1440 * 100) / 100,
      bottom: Math.round(parseInt(bottom, 10) / 1440 * 100) / 100,
      left: Math.round(parseInt(left, 10) / 1440 * 100) / 100,
      right: Math.round(parseInt(right, 10) / 1440 * 100) / 100,
    };
  }

  /**
   * Detect the current first-line indent from the Normal style.
   * @param {string} stylesXml
   * @returns {number|null} Indent in inches
   * @private
   */
  static _detectCurrentIndent(stylesXml) {
    const normalStyleRe = /(<w:style\s+w:type="paragraph"\s+w:default="1"[^>]*>)([\s\S]*?)(<\/w:style>)/;
    const normalMatch = stylesXml.match(normalStyleRe);
    if (!normalMatch) return null;

    const indentMatch = normalMatch[2].match(/<w:ind[^>]*w:firstLine="(\d+)"/);
    if (!indentMatch) return null;

    return Math.round(parseInt(indentMatch[1], 10) / 1440 * 100) / 100;
  }
}

// Export the PRESETS constant too for direct access
Presets.PRESETS = PRESETS;

module.exports = { Presets };
