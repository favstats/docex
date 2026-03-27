/**
 * sections.js -- Writing process tools for docex
 *
 * Provides section-level operations: outline, move, split, extract,
 * duplicate, and append. Operates on headings as section boundaries.
 *
 * All methods operate on a Workspace object.
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const xml = require('./xml');

// ============================================================================
// INTERNAL HELPERS
// ============================================================================

/**
 * Get the heading level of a paragraph XML fragment.
 * Returns 0 if not a heading.
 * @param {string} pXml
 * @returns {number}
 * @private
 */
function _headingLevel(pXml) {
  // Check for outlineLvl first (most reliable)
  const olvl = pXml.match(/<w:outlineLvl\s+w:val="(\d+)"/);
  if (olvl) {
    return parseInt(olvl[1], 10) + 1;
  }

  // Check style names
  const styleMatch = pXml.match(/<w:pStyle\s+w:val="([^"]+)"/);
  if (!styleMatch) return 0;
  const styleId = styleMatch[1];

  // Standard named styles
  const namedMatch = styleId.match(/^[Hh]eading(\d+)$/);
  if (namedMatch) return parseInt(namedMatch[1], 10);

  // OnlyOffice numeric IDs
  const DOCBUILDER_MAP = {
    '840': 1, '841': 2, '842': 3, '843': 4, '844': 5, '845': 6,
    '139': 1, '140': 2, '141': 3, '142': 4, '143': 5,
  };
  if (DOCBUILDER_MAP[styleId]) return DOCBUILDER_MAP[styleId];

  return 0;
}

/**
 * Get the w14:paraId attribute from a paragraph XML fragment.
 * @param {string} pXml
 * @returns {string|null}
 * @private
 */
function _getParaId(pXml) {
  const m = pXml.match(/w14:paraId="([^"]+)"/);
  return m ? m[1] : null;
}

/**
 * Check if a paragraph contains a figure (wp:inline or wp:anchor drawing).
 * @param {string} pXml
 * @returns {boolean}
 * @private
 */
function _hasFigure(pXml) {
  return /<wp:inline[\s>]/.test(pXml) || /<wp:anchor[\s>]/.test(pXml);
}

/**
 * Find a section by heading text. Case-insensitive substring match.
 * Returns the index of the heading paragraph in the paragraphs array.
 * @param {Array} paragraphs - From xml.findParagraphs()
 * @param {string} headingText - Heading text to match
 * @returns {number} Index of matching paragraph, or -1
 * @private
 */
function _findHeadingIndex(paragraphs, headingText) {
  const needle = headingText.toLowerCase().trim();
  for (let i = 0; i < paragraphs.length; i++) {
    const level = _headingLevel(paragraphs[i].xml);
    if (level > 0) {
      const pText = xml.extractText(paragraphs[i].xml).toLowerCase().trim();
      if (pText === needle || pText.includes(needle) || needle.includes(pText)) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * Get the range of paragraph indices for a section (from heading to next
 * heading of same or higher level, or end of document).
 * @param {Array} paragraphs
 * @param {number} headingIdx - Index of the heading paragraph
 * @returns {{ start: number, end: number }} start is inclusive, end is exclusive
 * @private
 */
function _sectionRange(paragraphs, headingIdx) {
  const level = _headingLevel(paragraphs[headingIdx].xml);
  let end = paragraphs.length;
  for (let i = headingIdx + 1; i < paragraphs.length; i++) {
    const hLevel = _headingLevel(paragraphs[i].xml);
    if (hLevel > 0 && hLevel <= level) {
      end = i;
      break;
    }
  }
  return { start: headingIdx, end };
}

/**
 * Replace all w14:paraId and w14:textId attributes in XML with fresh IDs.
 * @param {string} xmlStr
 * @returns {string}
 * @private
 */
function _replaceParaIds(xmlStr) {
  return xmlStr
    .replace(/w14:paraId="[^"]+"/g, () => `w14:paraId="${xml.randomHexId()}"`)
    .replace(/w14:textId="[^"]+"/g, () => `w14:textId="${xml.randomHexId()}"`);
}

// ============================================================================
// SECTIONS CLASS
// ============================================================================

class Sections {

  /**
   * Extract headings as a flat list for restructuring overview.
   * Each entry includes: level, text, paraId, paragraphCount, figureCount.
   *
   * @param {object} ws - Workspace
   * @returns {Array<{level: number, text: string, paraId: string|null, paragraphCount: number, figureCount: number}>}
   */
  static outline(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);
    const results = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const level = _headingLevel(paragraphs[i].xml);
      if (level > 0) {
        const text = xml.extractText(paragraphs[i].xml);
        const paraId = _getParaId(paragraphs[i].xml);
        const { start, end } = _sectionRange(paragraphs, i);

        // Count paragraphs in section (excluding the heading itself)
        const paragraphCount = end - start - 1;

        // Count figures in section
        let figureCount = 0;
        for (let j = start; j < end; j++) {
          if (_hasFigure(paragraphs[j].xml)) figureCount++;
        }

        results.push({ level, text, paraId, paragraphCount, figureCount });
      }
    }

    return results;
  }

  /**
   * Move a section (heading + all content until next same-level heading)
   * to before or after another heading.
   *
   * @param {object} ws - Workspace
   * @param {string} sectionHeading - Heading text of the section to move
   * @param {object} opts - { before: "heading" } or { after: "heading" }
   * @returns {{ moved: string, position: string, paragraphsMoved: number }}
   */
  static move(ws, sectionHeading, opts = {}) {
    if (!opts.before && !opts.after) {
      throw new Error('move() requires opts.before or opts.after');
    }

    let docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);

    // Find the source section
    const srcIdx = _findHeadingIndex(paragraphs, sectionHeading);
    if (srcIdx === -1) {
      throw new Error(`Section heading not found: "${sectionHeading}"`);
    }
    const srcRange = _sectionRange(paragraphs, srcIdx);

    // Find the target heading
    const targetHeading = opts.before || opts.after;
    const targetIdx = _findHeadingIndex(paragraphs, targetHeading);
    if (targetIdx === -1) {
      throw new Error(`Target heading not found: "${targetHeading}"`);
    }

    // Extract the section XML
    const sectionStart = paragraphs[srcRange.start].start;
    const sectionEnd = srcRange.end < paragraphs.length
      ? paragraphs[srcRange.end].start
      : docXml.indexOf('</w:body>');
    const sectionXml = docXml.slice(sectionStart, sectionEnd);

    // Remove the section from its current position
    docXml = docXml.slice(0, sectionStart) + docXml.slice(sectionEnd);

    // Re-parse to find the target position after removal
    const newParagraphs = xml.findParagraphs(docXml);
    const newTargetIdx = _findHeadingIndex(newParagraphs, targetHeading);
    if (newTargetIdx === -1) {
      throw new Error(`Target heading not found after removal: "${targetHeading}"`);
    }

    // Insert at the target position
    let insertPos;
    if (opts.before) {
      insertPos = newParagraphs[newTargetIdx].start;
    } else {
      // After: insert after the target section
      const targetRange = _sectionRange(newParagraphs, newTargetIdx);
      if (targetRange.end < newParagraphs.length) {
        insertPos = newParagraphs[targetRange.end].start;
      } else {
        insertPos = docXml.indexOf('</w:body>');
      }
    }

    docXml = docXml.slice(0, insertPos) + sectionXml + docXml.slice(insertPos);
    ws.docXml = docXml;

    const posStr = opts.before ? `before "${targetHeading}"` : `after "${targetHeading}"`;
    return {
      moved: sectionHeading,
      position: posStr,
      paragraphsMoved: srcRange.end - srcRange.start,
    };
  }

  /**
   * Extract a section into a new .docx file and remove it from the original.
   *
   * @param {object} ws - Workspace
   * @param {string} sectionHeading - Heading text of the section to extract
   * @param {string} outputPath - Path for the new .docx file
   * @returns {{ outputPath: string, paragraphsExtracted: number }}
   */
  static split(ws, sectionHeading, outputPath) {
    let docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);

    const srcIdx = _findHeadingIndex(paragraphs, sectionHeading);
    if (srcIdx === -1) {
      throw new Error(`Section heading not found: "${sectionHeading}"`);
    }
    const srcRange = _sectionRange(paragraphs, srcIdx);

    // Extract the section paragraphs XML
    const sectionStart = paragraphs[srcRange.start].start;
    const sectionEnd = srcRange.end < paragraphs.length
      ? paragraphs[srcRange.end].start
      : docXml.indexOf('</w:body>');
    const sectionContentXml = docXml.slice(sectionStart, sectionEnd);

    // Remove from original
    docXml = docXml.slice(0, sectionStart) + docXml.slice(sectionEnd);
    ws.docXml = docXml;

    // Build a minimal docx containing just this section
    const resolvedOutput = path.resolve(outputPath);
    const tmpDir = path.join(os.tmpdir(), `docex-split-${crypto.randomBytes(8).toString('hex')}`);

    try {
      fs.mkdirSync(path.join(tmpDir, '_rels'), { recursive: true });
      fs.mkdirSync(path.join(tmpDir, 'word', '_rels'), { recursive: true });

      // Build the document XML
      const newDocXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + `<w:document xmlns:mc="${xml.NS.mc}" `
        + `xmlns:r="${xml.NS.r}" `
        + `xmlns:w="${xml.NS.w}" `
        + `xmlns:w14="${xml.NS.w14}" `
        + `xmlns:wp="${xml.NS.wp}" `
        + `xmlns:wp14="${xml.NS.wp14}" `
        + `mc:Ignorable="w14 wp14">`
        + '<w:body>'
        + sectionContentXml
        + '<w:sectPr>'
        + '<w:pgSz w:w="12240" w:h="15840"/>'
        + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
        + '</w:sectPr>'
        + '</w:body></w:document>';

      // Content Types
      const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + '</Types>';

      // Root rels
      const rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        + '</Relationships>';

      // Word rels
      const wordRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '</Relationships>';

      fs.writeFileSync(path.join(tmpDir, '[Content_Types].xml'), contentTypes, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, '_rels', '.rels'), rootRels, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', 'document.xml'), newDocXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', '_rels', 'document.xml.rels'), wordRels, 'utf-8');

      // Zip it using execFileSync (safe, no shell)
      const outputDir = path.dirname(resolvedOutput);
      if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
      }
      if (fs.existsSync(resolvedOutput)) {
        fs.unlinkSync(resolvedOutput);
      }
      execFileSync('zip', ['-r', '-q', resolvedOutput, '.'], {
        cwd: tmpDir,
        stdio: 'pipe',
      });

    } finally {
      try { execFileSync('rm', ['-rf', tmpDir], { stdio: 'pipe' }); } catch (_) { /* ignore */ }
    }

    return {
      outputPath: resolvedOutput,
      paragraphsExtracted: srcRange.end - srcRange.start,
    };
  }

  /**
   * Extract the abstract text as a plain string.
   * Looks for a heading with "Abstract" in the text and returns all text
   * from paragraphs in that section.
   *
   * @param {object} ws - Workspace
   * @returns {string|null} Abstract text, or null if not found
   */
  static extractAbstract(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);

    // First, look for a heading named "Abstract"
    const absIdx = _findHeadingIndex(paragraphs, 'abstract');
    if (absIdx !== -1) {
      const { start, end } = _sectionRange(paragraphs, absIdx);
      const lines = [];
      for (let i = start + 1; i < end; i++) {
        const text = xml.extractText(paragraphs[i].xml).trim();
        if (text) lines.push(text);
      }
      return lines.join('\n') || null;
    }

    // Fallback: look for a paragraph with style "Abstract" (case-insensitive)
    for (const p of paragraphs) {
      const styleMatch = p.xml.match(/<w:pStyle\s+w:val="([^"]+)"/);
      if (styleMatch && styleMatch[1].toLowerCase() === 'abstract') {
        const text = xml.extractText(p.xml).trim();
        if (text) return text;
      }
    }

    return null;
  }

  /**
   * Duplicate a section with a new heading name.
   * The duplicate gets fresh paraIds to avoid conflicts.
   *
   * @param {object} ws - Workspace
   * @param {string} sectionHeading - Heading text of the section to copy
   * @param {string} newHeading - New heading text for the duplicate
   * @returns {{ duplicated: string, newHeading: string, paragraphsCopied: number }}
   */
  static duplicate(ws, sectionHeading, newHeading) {
    let docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);

    const srcIdx = _findHeadingIndex(paragraphs, sectionHeading);
    if (srcIdx === -1) {
      throw new Error(`Section heading not found: "${sectionHeading}"`);
    }
    const srcRange = _sectionRange(paragraphs, srcIdx);

    // Extract the section XML
    const sectionStart = paragraphs[srcRange.start].start;
    const sectionEnd = srcRange.end < paragraphs.length
      ? paragraphs[srcRange.end].start
      : docXml.indexOf('</w:body>');
    let sectionXml = docXml.slice(sectionStart, sectionEnd);

    // Replace the heading text with the new heading
    const headingPXml = paragraphs[srcRange.start].xml;
    const headingText = xml.extractText(headingPXml);

    // Replace text in runs within the first paragraph of the section
    const firstPEnd = sectionXml.indexOf('</w:p>') + '</w:p>'.length;
    let firstP = sectionXml.slice(0, firstPEnd);
    const restXml = sectionXml.slice(firstPEnd);

    // Replace all w:t contents in the heading paragraph
    firstP = _replaceTextInParagraph(firstP, headingText, newHeading);

    sectionXml = firstP + restXml;

    // Assign fresh paraIds
    sectionXml = _replaceParaIds(sectionXml);

    // Insert the duplicate right after the original section
    const insertPos = sectionEnd;
    docXml = docXml.slice(0, insertPos) + sectionXml + docXml.slice(insertPos);
    ws.docXml = docXml;

    return {
      duplicated: sectionHeading,
      newHeading,
      paragraphsCopied: srcRange.end - srcRange.start,
    };
  }

  /**
   * Append a paragraph at the end of a section (before the next heading).
   *
   * @param {object} ws - Workspace
   * @param {string} sectionHeading - Heading text of the target section
   * @param {string} text - Text to add
   * @param {object} [opts] - Options (unused for now, reserved for style)
   * @returns {{ appended: boolean, section: string, paraId: string }}
   */
  static append(ws, sectionHeading, text, opts = {}) {
    let docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);

    const srcIdx = _findHeadingIndex(paragraphs, sectionHeading);
    if (srcIdx === -1) {
      throw new Error(`Section heading not found: "${sectionHeading}"`);
    }
    const { end } = _sectionRange(paragraphs, srcIdx);

    // Build a new paragraph
    const newPara = xml.buildParagraph('', xml.buildRun('', text));

    // Insert before the next heading (or end of body)
    let insertPos;
    if (end < paragraphs.length) {
      insertPos = paragraphs[end].start;
    } else {
      insertPos = docXml.indexOf('</w:body>');
    }

    docXml = docXml.slice(0, insertPos) + newPara + docXml.slice(insertPos);
    ws.docXml = docXml;

    // Extract the paraId from the newly built paragraph
    const paraIdMatch = newPara.match(/w14:paraId="([^"]+)"/);
    const paraId = paraIdMatch ? paraIdMatch[1] : null;

    return {
      appended: true,
      section: sectionHeading,
      paraId,
    };
  }
}

// ============================================================================
// INTERNAL TEXT REPLACEMENT HELPER
// ============================================================================

/**
 * Replace text content in a paragraph XML, handling multi-run text.
 * Replaces the old text with new text in the w:t elements.
 * @param {string} pXml - Paragraph XML
 * @param {string} oldText - Text to find
 * @param {string} newText - Replacement text
 * @returns {string} Modified paragraph XML
 * @private
 */
function _replaceTextInParagraph(pXml, oldText, newText) {
  // Simple case: text exists in a single run
  const escapedOld = xml.escapeXml(oldText);
  const escapedNew = xml.escapeXml(newText);

  if (pXml.includes(`>${escapedOld}<`)) {
    return pXml.replace(`>${escapedOld}<`, `>${escapedNew}<`);
  }

  // Multi-run case: concatenate all w:t texts, find and replace,
  // then put the full replacement in the first run and empty the rest
  const runs = [];
  const runRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let m;
  while ((m = runRe.exec(pXml)) !== null) {
    runs.push({ full: m[0], text: m[1], index: m.index });
  }

  if (runs.length === 0) return pXml;

  const fullText = runs.map(r => r.text).join('');
  if (!fullText.includes(escapedOld)) {
    // Text not found even decoded; try decoded comparison
    const decodedFull = xml.decodeXml(fullText);
    if (decodedFull.includes(oldText)) {
      // Replace in the concatenated text
      const newFull = decodedFull.replace(oldText, newText);
      // Put all text in the first run
      let result = pXml;
      for (let i = runs.length - 1; i >= 0; i--) {
        const replacement = i === 0
          ? `<w:t xml:space="preserve">${xml.escapeXml(newFull)}</w:t>`
          : '<w:t xml:space="preserve"></w:t>';
        result = result.slice(0, runs[i].index) + replacement + result.slice(runs[i].index + runs[i].full.length);
      }
      return result;
    }
    return pXml; // Not found at all
  }

  // Replace in the escaped text
  const newFull = fullText.replace(escapedOld, escapedNew);
  let result = pXml;
  for (let i = runs.length - 1; i >= 0; i--) {
    const replacement = i === 0
      ? `<w:t xml:space="preserve">${newFull}</w:t>`
      : '<w:t xml:space="preserve"></w:t>';
    result = result.slice(0, runs[i].index) + replacement + result.slice(runs[i].index + runs[i].full.length);
  }
  return result;
}

module.exports = { Sections };
