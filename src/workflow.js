/**
 * workflow.js -- Writing workflow tools for docex (v0.4.10)
 *
 * Features for tracking writing progress, extracting TODOs,
 * previewing table of contents, and listing figures.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');
const { Paragraphs } = require('./paragraphs');
const { Comments } = require('./comments');

class Workflow {

  /**
   * Scan all comments and body text for TODO: and FIXME: patterns.
   *
   * @param {object} ws - Workspace
   * @returns {Array<{text: string, source: string, paraId: string, author: string}>}
   *   source is 'comment' or 'body'
   */
  static todo(ws) {
    const results = [];

    // 1. Scan body text for TODO: and FIXME: patterns
    const paragraphs = xml.findParagraphs(ws.docXml);
    for (const p of paragraphs) {
      const decoded = xml.extractTextDecoded(p.xml);
      const paraId = Workflow._getParaId(p.xml) || '';

      // Find TODO: and FIXME: patterns
      const todoRe = /\b(TODO|FIXME)\s*:\s*(.+?)(?:\.|$)/gi;
      let m;
      while ((m = todoRe.exec(decoded)) !== null) {
        results.push({
          text: m[0].trim(),
          source: 'body',
          paraId,
          author: '',
        });
      }
    }

    // 2. Scan comments for TODO: and FIXME: patterns
    const comments = Comments.list(ws);
    for (const c of comments) {
      const commentText = xml.decodeXml(c.text);
      // Check if the comment itself contains TODO/FIXME
      const todoRe = /\b(TODO|FIXME)\s*:\s*(.+?)(?:\.|$)/gi;
      let m;
      while ((m = todoRe.exec(commentText)) !== null) {
        results.push({
          text: m[0].trim(),
          source: 'comment',
          paraId: c.paraId || '',
          author: c.author || '',
        });
      }

      // Also include any comment that starts with or effectively is a TODO/FIXME
      if (/^\s*(TODO|FIXME)\b/i.test(commentText) && !results.some(r => r.text === commentText.trim() && r.source === 'comment')) {
        results.push({
          text: commentText.trim(),
          source: 'comment',
          paraId: c.paraId || '',
          author: c.author || '',
        });
      }
    }

    return results;
  }

  /**
   * Per-section analysis of writing progress.
   *
   * @param {object} ws - Workspace
   * @returns {Array<{section: string, status: string, wordCount: number, todoCount: number}>}
   *   status is 'done', 'draft', or 'empty'
   *   done = >200 words, 0 TODOs
   *   draft = >0 words, has TODOs or <200 words
   *   empty = 0 words
   */
  static progress(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);
    const sections = [];

    let currentSection = null;
    let currentWords = 0;
    let currentTodos = 0;

    for (const p of paragraphs) {
      const level = Paragraphs._headingLevel(p.xml);
      const decoded = xml.extractTextDecoded(p.xml).trim();

      if (level > 0) {
        // Save previous section
        if (currentSection !== null) {
          sections.push(Workflow._makeSectionStatus(currentSection, currentWords, currentTodos));
        }
        currentSection = decoded;
        currentWords = 0;
        currentTodos = 0;
        continue;
      }

      if (currentSection === null) {
        // Before any heading -- could be title/abstract area
        // We'll track it as a section called "(Preamble)" only if it has content
        if (decoded) {
          currentSection = '(Preamble)';
        }
      }

      if (decoded) {
        // Count words
        currentWords += Paragraphs._countWords(decoded);

        // Count TODOs/FIXMEs in body text
        const todoMatches = decoded.match(/\b(TODO|FIXME)\s*:/gi);
        if (todoMatches) {
          currentTodos += todoMatches.length;
        }
      }
    }

    // Save last section
    if (currentSection !== null) {
      sections.push(Workflow._makeSectionStatus(currentSection, currentWords, currentTodos));
    }

    return sections;
  }

  /**
   * Render table of contents as a string without inserting.
   * Uses heading hierarchy for indentation and numbering.
   *
   * @param {object} ws - Workspace
   * @returns {string} Multi-line TOC string
   */
  static tocPreview(ws) {
    const headings = Paragraphs.headings(ws);
    if (headings.length === 0) return '(No headings found)';

    const lines = [];
    const counters = [0, 0, 0, 0, 0, 0, 0, 0, 0]; // levels 1-9

    for (const h of headings) {
      const level = h.level;
      const decoded = xml.decodeXml(h.text).trim();

      // Increment counter at this level, reset deeper levels
      counters[level - 1]++;
      for (let i = level; i < 9; i++) {
        counters[i] = 0;
      }

      // Build number prefix
      const parts = [];
      for (let i = 0; i < level; i++) {
        parts.push(counters[i]);
      }
      const number = parts.join('.');

      // Indentation: 2 spaces per level beyond 1
      const indent = '  '.repeat(level - 1);

      lines.push(indent + number + '. ' + decoded);
    }

    return lines.join('\n');
  }

  /**
   * List all figures with captions and estimated page numbers.
   * Page estimates are based on paragraph position / estimated paragraphs per page.
   *
   * @param {object} ws - Workspace
   * @returns {string} Multi-line figure list
   */
  static figureList(ws) {
    const paragraphs = xml.findParagraphs(ws.docXml);

    // Estimate: ~25 paragraphs per page (rough heuristic)
    const PARAS_PER_PAGE = 25;

    const figures = [];
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const decoded = xml.extractTextDecoded(p.xml).trim();

      // Check for figure caption patterns
      const figMatch = decoded.match(/^(Figure\s+\d+[\.:]\s*.+)/i);
      if (figMatch) {
        const pageEst = Math.ceil((i + 1) / PARAS_PER_PAGE);
        figures.push({
          caption: figMatch[1],
          paraIndex: i,
          page: pageEst,
        });
        continue;
      }

      // Also check for embedded images (drawing elements) followed by captions
      if (p.xml.includes('<a:graphic') || p.xml.includes('<wp:inline') || p.xml.includes('<wp:anchor')) {
        // This paragraph contains an image -- check the next paragraph for a caption
        if (i + 1 < paragraphs.length) {
          const nextDecoded = xml.extractTextDecoded(paragraphs[i + 1].xml).trim();
          const nextFigMatch = nextDecoded.match(/^(Figure\s+\d+[\.:]\s*.+)/i);
          if (nextFigMatch) {
            const pageEst = Math.ceil((i + 1) / PARAS_PER_PAGE);
            figures.push({
              caption: nextFigMatch[1],
              paraIndex: i + 1,
              page: pageEst,
            });
          }
        }
      }
    }

    // Deduplicate by paraIndex
    const seen = new Set();
    const unique = [];
    for (const f of figures) {
      if (!seen.has(f.paraIndex)) {
        seen.add(f.paraIndex);
        unique.push(f);
      }
    }

    if (unique.length === 0) return '(No figures found)';

    return unique.map(f =>
      f.caption + ' (~p.' + f.page + ')'
    ).join('\n');
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Build a section status object.
   * @param {string} name
   * @param {number} wordCount
   * @param {number} todoCount
   * @returns {object}
   * @private
   */
  static _makeSectionStatus(name, wordCount, todoCount) {
    let status;
    if (wordCount === 0) {
      status = 'empty';
    } else if (wordCount > 200 && todoCount === 0) {
      status = 'done';
    } else {
      status = 'draft';
    }

    return { section: name, status, wordCount, todoCount };
  }

  /**
   * Extract w14:paraId from paragraph XML.
   * @param {string} pXml
   * @returns {string|null}
   * @private
   */
  static _getParaId(pXml) {
    const m = pXml.match(/w14:paraId="([^"]+)"/);
    return m ? m[1] : null;
  }
}

module.exports = { Workflow };
