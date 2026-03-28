/**
 * docmap.js -- Document map and paraId injection for docex
 *
 * Provides stable addressing via w14:paraId attributes on every <w:p> element.
 * Parses document structure into sections, figures, tables, and comments.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');
const { Paragraphs } = require('./paragraphs');
const { Comments } = require('./comments');

// ============================================================================
// DOCMAP
// ============================================================================

class DocMap {

  /**
   * Generate a structured map of the document.
   *
   * Parses all paragraphs, groups them into sections by heading,
   * detects figures and tables, and returns a navigable tree.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {{sections: Array, allParagraphs: Array, allFigures: Array, allTables: Array, allComments: Array}}
   */
  static generate(ws) {
    const docXml = ws.docXml;
    const rawParagraphs = xml.findParagraphs(docXml);

    const allParagraphs = [];
    const allFigures = [];
    const allTables = [];
    const sections = [];

    let currentSection = null;
    let figureCount = 0;
    let tableCount = 0;

    for (let i = 0; i < rawParagraphs.length; i++) {
      const p = rawParagraphs[i];
      const paraId = DocMap._extractParaId(p.xml);
      const text = xml.extractTextDecoded(p.xml);
      const level = Paragraphs._headingLevel(p.xml);
      const hasFigure = /<w:drawing[\s>]/.test(p.xml) || /<w:pict[\s>]/.test(p.xml);
      const isCaption = /^(Figure|Table)\s+\d/i.test(text.trim());
      const isAbstract = /^abstract$/i.test(text.trim()) && level === 0;

      let type;
      if (level > 0) {
        type = 'heading';
      } else if (isCaption) {
        type = 'caption';
      } else if (isAbstract) {
        type = 'abstract';
      } else {
        type = 'body';
      }

      const paraInfo = {
        id: paraId,
        text,
        type,
        level: level > 0 ? level : undefined,
        index: i,
      };

      allParagraphs.push(paraInfo);

      // Start a new section when we encounter a heading
      if (level > 0) {
        currentSection = {
          heading: { id: paraId, text, level, index: i },
          paragraphs: [],
          figures: [],
          tables: [],
        };
        sections.push(currentSection);
        continue;
      }

      // Detect figures
      if (hasFigure) {
        figureCount++;
        const captionText = isCaption ? text : '';
        const figInfo = {
          id: paraId,
          number: figureCount,
          caption: captionText,
          index: i,
        };
        allFigures.push(figInfo);
        if (currentSection) {
          currentSection.figures.push(figInfo);
        }
      }

      // Add paragraph to current section
      if (currentSection) {
        currentSection.paragraphs.push(paraInfo);
      }
    }

    // Detect tables (w:tbl elements between paragraphs)
    const tblRe = /<w:tbl[\s>]/g;
    let tblMatch;
    while ((tblMatch = tblRe.exec(docXml)) !== null) {
      tableCount++;
      const tblStart = tblMatch.index;
      const tblEnd = docXml.indexOf('</w:tbl>', tblStart);
      if (tblEnd === -1) continue;

      const tblXml = docXml.slice(tblStart, tblEnd + 8);
      const tblText = xml.extractTextDecoded(tblXml);

      const tblInfo = {
        number: tableCount,
        text: tblText.slice(0, 100),
        xmlStart: tblStart,
      };
      allTables.push(tblInfo);

      // Find which section this table belongs to
      if (sections.length > 0) {
        // The table belongs to the last section whose heading starts before it
        for (let s = sections.length - 1; s >= 0; s--) {
          const sectionHeadingId = sections[s].heading.id;
          // Find the heading paragraph's position
          const headingPara = rawParagraphs.find(p =>
            DocMap._extractParaId(p.xml) === sectionHeadingId
          );
          if (headingPara && headingPara.start < tblStart) {
            sections[s].tables.push(tblInfo);
            break;
          }
        }
      }
    }

    // Get comments
    let allComments = [];
    try {
      allComments = Comments.list(ws);
    } catch (_) {
      // No comments or comments.xml doesn't exist
    }

    return {
      sections,
      allParagraphs,
      allFigures,
      allTables,
      allComments,
    };
  }

  /**
   * Inject w14:paraId attributes into all <w:p> elements that lack one.
   *
   * Scans ws.docXml for <w:p> elements without w14:paraId.
   * For each, injects w14:paraId="XXXXXXXX" (random 8-char hex, uppercased).
   * Updates ws.docXml in place.
   *
   * Also ensures the w14 namespace is declared on the root element.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {number} Count of injected IDs
   */
  static injectParaIds(ws) {
    let docXml = ws.docXml;

    // Ensure w14 namespace is declared on root element
    if (!docXml.includes('xmlns:w14=')) {
      docXml = docXml.replace(
        /<w:document\b/,
        `<w:document xmlns:w14="${xml.NS.w14}"`
      );
    }

    // Collect all existing paraIds to avoid collisions
    const existingIds = new Set();
    const existingRe = /w14:paraId="([^"]+)"/g;
    let em;
    while ((em = existingRe.exec(docXml)) !== null) {
      existingIds.add(em[1].toUpperCase());
    }

    let count = 0;

    // Find all <w:p> elements and inject paraId where missing.
    // We need to handle both <w:p> and <w:p ...> (with attrs).
    // Process from end to start to preserve offsets.
    const pOpenRe = /<w:p([\s>])/g;
    const matches = [];
    let pm;
    while ((pm = pOpenRe.exec(docXml)) !== null) {
      matches.push({ index: pm.index, after: pm[1], full: pm[0] });
    }

    // Process in reverse to preserve string positions
    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i];
      // Check if this <w:p> already has a paraId
      const closeAngle = docXml.indexOf('>', m.index);
      if (closeAngle === -1) continue;
      const tagContent = docXml.slice(m.index, closeAngle + 1);
      if (tagContent.includes('w14:paraId=')) continue;

      // Generate deterministic ID from paragraph content + position index.
      // This ensures the same paragraph always gets the same ID across
      // multiple opens (stable for .dex diffing), while the position index
      // prevents collisions between paragraphs with identical text.
      const crypto = require('crypto');
      const paraEnd = docXml.indexOf('</w:p>', m.index);
      const paraText = paraEnd !== -1 ? xml.extractText(docXml.slice(m.index, paraEnd + 6)) : String(i);
      const hashInput = i + ':' + paraText;
      let newId = crypto.createHash('md5').update(hashInput).digest('hex').slice(0, 8).toUpperCase();
      // Handle collisions (unlikely but possible)
      let collisionCounter = 0;
      while (existingIds.has(newId)) {
        collisionCounter++;
        newId = crypto.createHash('md5').update(hashInput + ':' + collisionCounter).digest('hex').slice(0, 8).toUpperCase();
      }
      existingIds.add(newId);

      // textId also deterministic
      const textHashInput = i + ':text:' + paraText;
      let textId = crypto.createHash('md5').update(textHashInput).digest('hex').slice(0, 8).toUpperCase();
      while (existingIds.has(textId)) {
        textId = crypto.createHash('md5').update(textHashInput + ':' + (++collisionCounter)).digest('hex').slice(0, 8).toUpperCase();
      }
      existingIds.add(textId);

      // Inject attributes after "<w:p"
      const insertPos = m.index + 4; // after "<w:p"
      const attrs = ` w14:paraId="${newId}" w14:textId="${textId}"`;
      docXml = docXml.slice(0, insertPos) + attrs + docXml.slice(insertPos);
      count++;
    }

    if (count > 0) {
      ws.docXml = docXml;
    }

    return count;
  }

  /**
   * Find paragraphs containing the given text.
   *
   * Returns matches with section context and surrounding text.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} searchText - Text to search for (substring match)
   * @returns {Array<{id: string, index: number, section: string, context: string}>}
   */
  static find(ws, searchText) {
    const map = DocMap.generate(ws);
    const results = [];

    for (const para of map.allParagraphs) {
      const decoded = para.text;
      if (!decoded.includes(searchText)) continue;

      // Find section this paragraph belongs to
      let sectionName = '(before first heading)';
      for (const section of map.sections) {
        if (section.heading.index <= para.index) {
          sectionName = section.heading.text;
        }
      }

      // Build context: surrounding 30 chars on each side
      const pos = decoded.indexOf(searchText);
      const ctxStart = Math.max(0, pos - 30);
      const ctxEnd = Math.min(decoded.length, pos + searchText.length + 30);
      let context = '';
      if (ctxStart > 0) context += '...';
      context += decoded.slice(ctxStart, ctxEnd);
      if (ctxEnd < decoded.length) context += '...';

      results.push({
        id: para.id,
        index: para.index,
        section: sectionName,
        context,
      });
    }

    return results;
  }

  /**
   * Return a tree-view string of the document structure.
   *
   * Format:
   *   Introduction (H1)
   *     3 paragraphs, 0 figures
   *   Methods (H1)
   *     Data Collection (H2)
   *       2 paragraphs, 1 figure
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {string}
   */
  static structure(ws) {
    const map = DocMap.generate(ws);
    const lines = [];

    for (const section of map.sections) {
      const h = section.heading;
      const indent = '  '.repeat(h.level - 1);
      lines.push(`${indent}${h.text} (H${h.level})`);

      const parts = [];
      if (section.paragraphs.length > 0) {
        parts.push(`${section.paragraphs.length} paragraph${section.paragraphs.length !== 1 ? 's' : ''}`);
      }
      if (section.figures.length > 0) {
        parts.push(`${section.figures.length} figure${section.figures.length !== 1 ? 's' : ''}`);
      }
      if (section.tables.length > 0) {
        parts.push(`${section.tables.length} table${section.tables.length !== 1 ? 's' : ''}`);
      }
      if (parts.length > 0) {
        lines.push(`${indent}  ${parts.join(', ')}`);
      }
    }

    return lines.join('\n');
  }

  /**
   * Find text and show the XML structure around it.
   *
   * Shows which run contains the text, what formatting it has.
   * Useful for debugging why an operation might fail.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {string} searchText - Text to find and explain
   * @returns {string} Human-readable explanation
   */
  static explain(ws, searchText) {
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    const lines = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      const decoded = xml.extractTextDecoded(para.xml);
      if (!decoded.includes(searchText)) continue;

      const paraId = DocMap._extractParaId(para.xml);

      // Find section
      let sectionName = '(before first heading)';
      for (let j = i - 1; j >= 0; j--) {
        if (Paragraphs._headingLevel(paragraphs[j].xml) > 0) {
          sectionName = xml.extractTextDecoded(paragraphs[j].xml);
          break;
        }
      }

      lines.push(`Paragraph ${i} (id: ${paraId}), section "${sectionName}":`);

      // Show runs
      const runs = xml.parseRuns(para.xml);
      for (let r = 0; r < runs.length; r++) {
        const run = runs[r];
        const runText = xml.decodeXml(run.combinedText);
        const isTarget = runText.includes(searchText);
        const marker = isTarget ? '  <-- HERE' : '';

        // Summarize formatting
        const fmtParts = [];
        if (run.rPr.includes('<w:b/>') || run.rPr.includes('<w:b ')) fmtParts.push('bold');
        if (run.rPr.includes('<w:i/>') || run.rPr.includes('<w:i ')) fmtParts.push('italic');
        if (run.rPr.includes('<w:u ')) fmtParts.push('underline');
        const fontMatch = run.rPr.match(/w:ascii="([^"]+)"/);
        if (fontMatch) fmtParts.push(fontMatch[1]);
        const sizeMatch = run.rPr.match(/w:sz w:val="(\d+)"/);
        if (sizeMatch) fmtParts.push(`${parseInt(sizeMatch[1], 10) / 2}pt`);

        const fmtStr = fmtParts.length > 0 ? ` (${fmtParts.join(', ')})` : '';
        let truncText;
        if (isTarget && runText.length > 80) {
          // Show the region around the search text
          const pos = runText.indexOf(searchText);
          const ctxStart = Math.max(0, pos - 20);
          const ctxEnd = Math.min(runText.length, pos + searchText.length + 20);
          truncText = (ctxStart > 0 ? '...' : '') + runText.slice(ctxStart, ctxEnd) + (ctxEnd < runText.length ? '...' : '');
        } else if (runText.length > 80) {
          truncText = runText.slice(0, 77) + '...';
        } else {
          truncText = runText;
        }
        lines.push(`  Run ${r + 1}: "${truncText}"${fmtStr}${marker}`);
      }

      lines.push('');
      break; // Only show first matching paragraph
    }

    if (lines.length === 0) {
      lines.push(`Text "${searchText}" not found in any paragraph.`);
    }

    return lines.join('\n');
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Extract w14:paraId from a paragraph XML fragment.
   *
   * @param {string} pXml - Paragraph XML
   * @returns {string} paraId or empty string
   * @private
   */
  static _extractParaId(pXml) {
    const m = pXml.match(/w14:paraId="([^"]+)"/);
    return m ? m[1] : '';
  }

  /**
   * Find a paragraph by its w14:paraId in the document XML.
   *
   * Returns the XML fragment and its start/end positions in the full document.
   *
   * @param {string} docXml - Full document XML
   * @param {string} paraId - The w14:paraId to find
   * @returns {{xml: string, start: number, end: number}|null}
   */
  static locateById(docXml, paraId) {
    // Find the paraId attribute in the document
    const searchStr = `w14:paraId="${paraId}"`;
    const attrPos = docXml.indexOf(searchStr);
    if (attrPos === -1) return null;

    // Walk backwards to find the <w:p start
    let pStart = docXml.lastIndexOf('<w:p', attrPos);
    if (pStart === -1) return null;

    // Find the closing </w:p>
    const closeTag = '</w:p>';
    const closePos = docXml.indexOf(closeTag, pStart);
    if (closePos === -1) return null;

    const pEnd = closePos + closeTag.length;
    const pXml = docXml.slice(pStart, pEnd);

    return { xml: pXml, start: pStart, end: pEnd };
  }
}

module.exports = { DocMap };
