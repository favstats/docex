/**
 * verify.js -- Submission validation for docex
 *
 * Validates a document against journal submission requirements.
 * Checks word count, abstract length, heading hierarchy, margins,
 * font, spacing, running header, title page, line numbering, etc.
 *
 * All methods operate on a Workspace object.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');
const { Paragraphs } = require('./paragraphs');
const { Presets } = require('./presets');

// ============================================================================
// VERIFY
// ============================================================================

class Verify {

  /**
   * Validate document against journal requirements defined by a preset.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} presetName - Preset name (e.g. "polcomm", "apa7")
   * @returns {{ pass: boolean, errors: string[], warnings: string[] }}
   */
  static check(ws, presetName) {
    const config = Presets.get(presetName);
    if (!config) {
      const available = Presets.list().join(', ');
      throw new Error(`Unknown preset: "${presetName}". Available: ${available}`);
    }

    const errors = [];
    const warnings = [];

    // 1. Word count check
    if (config.wordLimit) {
      const wc = Paragraphs.wordCount(ws);
      if (wc.body > config.wordLimit) {
        errors.push(`Word count ${wc.body.toLocaleString()} exceeds limit ${config.wordLimit.toLocaleString()}`);
      } else if (wc.body > config.wordLimit * 0.95) {
        warnings.push(`Word count ${wc.body.toLocaleString()} is close to limit ${config.wordLimit.toLocaleString()}`);
      }
    }

    // 2. Abstract word limit check
    if (config.abstractWordLimit) {
      const wc = Paragraphs.wordCount(ws);
      if (wc.abstract > config.abstractWordLimit) {
        errors.push(`Abstract ${wc.abstract} words, limit is ${config.abstractWordLimit}`);
      }
    }

    // 3. Font check
    if (config.font) {
      const fontCheck = Verify._checkFont(ws, config.font);
      if (fontCheck.nonCompliant > 0) {
        warnings.push(`${fontCheck.nonCompliant} run(s) use non-${config.font} fonts`);
      }
    }

    // 4. Line spacing check
    if (config.spacing) {
      const spacingCheck = Verify._checkSpacing(ws, config.spacing);
      if (!spacingCheck.compliant) {
        warnings.push(`Line spacing is not ${config.spacing} in document defaults`);
      }
    }

    // 5. Margin check
    if (config.margins) {
      const marginCheck = Verify._checkMargins(ws, config.margins);
      if (!marginCheck.compliant) {
        errors.push(`Margins do not match ${presetName} requirements: ${marginCheck.details}`);
      }
    }

    // 6. Title page check
    if (config.titlePage) {
      const hasTitlePage = Verify._checkTitlePage(ws);
      if (!hasTitlePage) {
        warnings.push('No title page detected (expected for ' + presetName + ')');
      }
    }

    // 7. Running header check
    if (config.runningHeader) {
      const hasHeader = Verify._checkRunningHeader(ws);
      if (!hasHeader) {
        warnings.push('No running header detected');
      }
    }

    // 8. Heading hierarchy check
    const headingCheck = Verify._checkHeadingHierarchy(ws);
    if (headingCheck.length > 0) {
      for (const issue of headingCheck) {
        warnings.push(issue);
      }
    }

    return {
      pass: errors.length === 0,
      errors,
      warnings,
    };
  }

  // --------------------------------------------------------------------------
  // Internal checks
  // --------------------------------------------------------------------------

  /**
   * Check if document body uses the expected font.
   * @private
   */
  static _checkFont(ws, expectedFont) {
    const docXml = ws.docXml;
    const fontLower = expectedFont.toLowerCase();

    let nonCompliant = 0;
    const fontRe = /<w:rFonts[^>]+w:ascii="([^"]+)"/g;
    let m;
    while ((m = fontRe.exec(docXml)) !== null) {
      if (m[1].toLowerCase() !== fontLower) {
        nonCompliant++;
      }
    }

    return { compliant: nonCompliant === 0, nonCompliant };
  }

  /**
   * Check if document uses the expected line spacing.
   * @private
   */
  static _checkSpacing(ws, expectedSpacing) {
    const stylesXml = ws.stylesXml;
    if (!stylesXml) return { compliant: false };

    const spacingMap = {
      'single': 240,
      '1.15': 276,
      '1.5': 360,
      'double': 480,
    };
    const expected = spacingMap[expectedSpacing] || 480;

    const spacingMatch = stylesXml.match(/<w:spacing[^>]+w:line="(\d+)"/);
    if (spacingMatch) {
      const actual = parseInt(spacingMatch[1], 10);
      return { compliant: actual === expected };
    }

    return { compliant: false };
  }

  /**
   * Check if document margins match expected values.
   * @private
   */
  static _checkMargins(ws, expected) {
    const docXml = ws.docXml;

    const pgMarMatch = docXml.match(/<w:pgMar\s+([^/]*)\/>/);
    if (!pgMarMatch) {
      return { compliant: false, details: 'No page margins found' };
    }

    const attrs = pgMarMatch[1];
    const getAttr = (name) => {
      const m = attrs.match(new RegExp(`w:${name}="(\\d+)"`));
      return m ? parseInt(m[1], 10) : 0;
    };

    const tolerance = 72; // 0.05 inch tolerance in twips

    const issues = [];
    if (Math.abs(getAttr('top') - expected.top * 1440) > tolerance) {
      issues.push(`top margin ${(getAttr('top') / 1440).toFixed(2)}" != ${expected.top}"`);
    }
    if (Math.abs(getAttr('bottom') - expected.bottom * 1440) > tolerance) {
      issues.push(`bottom margin ${(getAttr('bottom') / 1440).toFixed(2)}" != ${expected.bottom}"`);
    }
    if (Math.abs(getAttr('left') - expected.left * 1440) > tolerance) {
      issues.push(`left margin ${(getAttr('left') / 1440).toFixed(2)}" != ${expected.left}"`);
    }
    if (Math.abs(getAttr('right') - expected.right * 1440) > tolerance) {
      issues.push(`right margin ${(getAttr('right') / 1440).toFixed(2)}" != ${expected.right}"`);
    }

    return {
      compliant: issues.length === 0,
      details: issues.join(', ') || 'OK',
    };
  }

  /**
   * Check if a title page is present.
   * @private
   */
  static _checkTitlePage(ws) {
    const docXml = ws.docXml;

    if (docXml.includes('<w:titlePg/>') || docXml.includes('<w:titlePg />')) {
      return true;
    }

    const paragraphs = xml.findParagraphs(docXml);
    if (paragraphs.length > 0) {
      const first = paragraphs[0].xml;
      if (first.includes('w:val="Title"') || first.includes('pStyle w:val="Title"')) {
        return true;
      }
    }

    return false;
  }

  /**
   * Check if a running header is present.
   * @private
   */
  static _checkRunningHeader(ws) {
    const docXml = ws.docXml;
    if (docXml.includes('<w:headerReference')) {
      return true;
    }
    return false;
  }

  /**
   * Check heading hierarchy for skipped levels.
   * @private
   */
  static _checkHeadingHierarchy(ws) {
    const headings = Paragraphs.headings(ws);
    const issues = [];

    let prevLevel = 0;
    for (const h of headings) {
      if (h.level > prevLevel + 1 && prevLevel > 0) {
        issues.push(
          `Heading levels skip from H${prevLevel} to H${h.level}: "${xml.decodeXml(h.text).slice(0, 40)}"`
        );
      }
      prevLevel = h.level;
    }

    return issues;
  }
}

module.exports = { Verify };
