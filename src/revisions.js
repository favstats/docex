/**
 * revisions.js -- Accept/reject tracked changes in OOXML documents.
 *
 * Static methods for listing, accepting, and rejecting w:ins and w:del
 * tracked change elements. Also provides a cleanCopy() method that
 * accepts all changes and removes all comment markup.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set)
 * and ws.commentsXml (get/set). XML manipulation is done entirely
 * with string operations and regex. Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// REVISIONS
// ============================================================================

class Revisions {

  /**
   * Scan document.xml for all w:ins and w:del elements.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {Array<{id: number, type: string, author: string, date: string, text: string, start: number, end: number}>}
   */
  static list(ws) {
    const docXml = ws.docXml;
    const results = [];

    // Match w:ins elements (insertions)
    const insRe = /<w:ins\b([^>]*)>([\s\S]*?)<\/w:ins>/g;
    let m;
    while ((m = insRe.exec(docXml)) !== null) {
      const attrs = m[1];
      const body = m[2];
      const id = Revisions._attrVal(attrs, 'w:id');
      const author = Revisions._attrVal(attrs, 'w:author') || '';
      const date = Revisions._attrVal(attrs, 'w:date') || '';
      const text = xml.extractText(body);
      results.push({
        id: id ? parseInt(id, 10) : 0,
        type: 'insertion',
        author: xml.decodeXml(author),
        date,
        text,
        start: m.index,
        end: m.index + m[0].length,
      });
    }

    // Match w:del elements (deletions)
    const delRe = /<w:del\b([^>]*)>([\s\S]*?)<\/w:del>/g;
    while ((m = delRe.exec(docXml)) !== null) {
      const attrs = m[1];
      const body = m[2];
      const id = Revisions._attrVal(attrs, 'w:id');
      const author = Revisions._attrVal(attrs, 'w:author') || '';
      const date = Revisions._attrVal(attrs, 'w:date') || '';
      // Extract text from w:delText elements
      const text = Revisions._extractDelText(body);
      results.push({
        id: id ? parseInt(id, 10) : 0,
        type: 'deletion',
        author: xml.decodeXml(author),
        date,
        text,
        start: m.index,
        end: m.index + m[0].length,
      });
    }

    // Sort by position in document (ascending)
    results.sort((a, b) => a.start - b.start);
    return results;
  }

  /**
   * Accept tracked changes.
   *
   * - For w:ins: unwrap (keep content, remove w:ins wrapper)
   * - For w:del: remove entirely (delete the w:del and its content)
   *
   * If id is provided, accept only that specific change.
   * If id is undefined/null, accept ALL changes.
   *
   * @param {object} ws - Workspace with ws.docXml get/set
   * @param {number} [id] - Specific change ID to accept, or undefined for all
   */
  static accept(ws, id) {
    let docXml = ws.docXml;

    if (id != null) {
      // Accept a specific change by ID
      docXml = Revisions._acceptInsById(docXml, id);
      docXml = Revisions._acceptDelById(docXml, id);
    } else {
      // Accept ALL changes: process from end to start to preserve positions
      // First handle all w:del (remove entirely)
      docXml = Revisions._acceptAllDel(docXml);
      // Then handle all w:ins (unwrap)
      docXml = Revisions._acceptAllIns(docXml);
    }

    ws.docXml = docXml;
  }

  /**
   * Reject tracked changes.
   *
   * - For w:ins: remove entirely (delete w:ins and content)
   * - For w:del: unwrap (keep content as w:r with w:t instead of w:delText)
   *
   * If id is provided, reject only that specific change.
   * If id is undefined/null, reject ALL changes.
   *
   * @param {object} ws - Workspace with ws.docXml get/set
   * @param {number} [id] - Specific change ID to reject, or undefined for all
   */
  static reject(ws, id) {
    let docXml = ws.docXml;

    if (id != null) {
      // Reject a specific change by ID
      docXml = Revisions._rejectInsById(docXml, id);
      docXml = Revisions._rejectDelById(docXml, id);
    } else {
      // Reject ALL changes: process from end to start
      // First handle all w:ins (remove entirely)
      docXml = Revisions._rejectAllIns(docXml);
      // Then handle all w:del (unwrap, converting delText to t)
      docXml = Revisions._rejectAllDel(docXml);
    }

    ws.docXml = docXml;
  }

  /**
   * Produce a clean copy of the document:
   *   - Accept all tracked changes
   *   - Resolve all comments (set done="1")
   *   - Remove all comment range markers from document.xml
   *   - Remove all comments from comments.xml
   *
   * @param {object} ws - Workspace with ws.docXml, ws.commentsXml, ws.commentsExtXml
   */
  static cleanCopy(ws) {
    // 1. Accept all tracked changes
    Revisions.accept(ws);

    // 2. Remove comment range markers and references from document.xml
    let docXml = ws.docXml;
    docXml = docXml.replace(/<w:commentRangeStart\s+[^>]*?\/>/g, '');
    docXml = docXml.replace(/<w:commentRangeEnd\s+[^>]*?\/>/g, '');
    // Remove commentReference runs (the run element that contains the reference)
    docXml = docXml.replace(
      /<w:r>[^<]*(?:<w:rPr>[\s\S]*?<\/w:rPr>)?[^<]*<w:commentReference\s+[^>]*?\/>[^<]*<\/w:r>/g,
      ''
    );
    ws.docXml = docXml;

    // 3. Resolve all comments in commentsExtended.xml (set done="1")
    let extXml = ws.commentsExtXml;
    if (extXml) {
      extXml = extXml.replace(/w15:done="0"/g, 'w15:done="1"');
      ws.commentsExtXml = extXml;
    }

    // 4. Remove all comments from comments.xml
    let commentsXml = ws.commentsXml;
    if (commentsXml) {
      commentsXml = commentsXml.replace(/<w:comment\b[^>]*>[\s\S]*?<\/w:comment>/g, '');
      ws.commentsXml = commentsXml;
    }
  }

  // --------------------------------------------------------------------------
  // INTERNAL: Accept helpers
  // --------------------------------------------------------------------------

  /**
   * Accept a specific w:ins by ID: unwrap (keep content).
   * @private
   */
  static _acceptInsById(docXml, id) {
    const idStr = String(id);
    const re = new RegExp(
      '<w:ins\\b[^>]*\\bw:id="' + idStr + '"[^>]*>([\\s\\S]*?)<\\/w:ins>',
      'g'
    );
    return docXml.replace(re, '$1');
  }

  /**
   * Accept a specific w:del by ID: remove entirely.
   * @private
   */
  static _acceptDelById(docXml, id) {
    const idStr = String(id);
    const re = new RegExp(
      '<w:del\\b[^>]*\\bw:id="' + idStr + '"[^>]*>[\\s\\S]*?<\\/w:del>',
      'g'
    );
    return docXml.replace(re, '');
  }

  /**
   * Accept all w:ins elements: unwrap (keep content).
   * Processes from end to start to preserve positions.
   * @private
   */
  static _acceptAllIns(docXml) {
    // Collect all matches first, then replace from end to start
    const matches = [];
    const re = /<w:ins\b[^>]*>([\s\S]*?)<\/w:ins>/g;
    let m;
    while ((m = re.exec(docXml)) !== null) {
      matches.push({ start: m.index, end: m.index + m[0].length, content: m[1] });
    }
    // Process from end to start
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      docXml = docXml.slice(0, match.start) + match.content + docXml.slice(match.end);
    }
    return docXml;
  }

  /**
   * Accept all w:del elements: remove entirely.
   * Processes from end to start to preserve positions.
   * @private
   */
  static _acceptAllDel(docXml) {
    const matches = [];
    const re = /<w:del\b[^>]*>[\s\S]*?<\/w:del>/g;
    let m;
    while ((m = re.exec(docXml)) !== null) {
      matches.push({ start: m.index, end: m.index + m[0].length });
    }
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      docXml = docXml.slice(0, match.start) + docXml.slice(match.end);
    }
    return docXml;
  }

  // --------------------------------------------------------------------------
  // INTERNAL: Reject helpers
  // --------------------------------------------------------------------------

  /**
   * Reject a specific w:ins by ID: remove entirely.
   * @private
   */
  static _rejectInsById(docXml, id) {
    const idStr = String(id);
    const re = new RegExp(
      '<w:ins\\b[^>]*\\bw:id="' + idStr + '"[^>]*>[\\s\\S]*?<\\/w:ins>',
      'g'
    );
    return docXml.replace(re, '');
  }

  /**
   * Reject a specific w:del by ID: unwrap and convert delText to t.
   * @private
   */
  static _rejectDelById(docXml, id) {
    const idStr = String(id);
    const re = new RegExp(
      '<w:del\\b[^>]*\\bw:id="' + idStr + '"[^>]*>([\\s\\S]*?)<\\/w:del>',
      'g'
    );
    return docXml.replace(re, (_, content) => {
      return Revisions._convertDelTextToText(content);
    });
  }

  /**
   * Reject all w:ins elements: remove entirely.
   * Processes from end to start.
   * @private
   */
  static _rejectAllIns(docXml) {
    const matches = [];
    const re = /<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/g;
    let m;
    while ((m = re.exec(docXml)) !== null) {
      matches.push({ start: m.index, end: m.index + m[0].length });
    }
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      docXml = docXml.slice(0, match.start) + docXml.slice(match.end);
    }
    return docXml;
  }

  /**
   * Reject all w:del elements: unwrap and convert delText to t.
   * Processes from end to start.
   * @private
   */
  static _rejectAllDel(docXml) {
    const matches = [];
    const re = /<w:del\b[^>]*>([\s\S]*?)<\/w:del>/g;
    let m;
    while ((m = re.exec(docXml)) !== null) {
      matches.push({
        start: m.index,
        end: m.index + m[0].length,
        content: m[1],
      });
    }
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      const converted = Revisions._convertDelTextToText(match.content);
      docXml = docXml.slice(0, match.start) + converted + docXml.slice(match.end);
    }
    return docXml;
  }

  // --------------------------------------------------------------------------
  // INTERNAL: Utility helpers
  // --------------------------------------------------------------------------

  /**
   * Extract text from w:delText elements within a w:del body.
   * @private
   */
  static _extractDelText(body) {
    const texts = [];
    const re = /<w:delText[^>]*>([^<]*)<\/w:delText>/g;
    let m;
    while ((m = re.exec(body)) !== null) {
      texts.push(m[1]);
    }
    return texts.join('');
  }

  /**
   * Convert w:delText elements to w:t elements within content.
   * Used when rejecting a deletion (keeping the deleted text as normal text).
   * @private
   */
  static _convertDelTextToText(content) {
    // Replace <w:delText ...> with <w:t ...> and </w:delText> with </w:t>
    let result = content.replace(/<w:delText(\b[^>]*)>/g, '<w:t$1>');
    result = result.replace(/<\/w:delText>/g, '</w:t>');
    return result;
  }

  /**
   * Extract an attribute value from an attribute string.
   * @private
   */
  static _attrVal(attrs, name) {
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(escaped + '="([^"]*)"');
    const m = attrs.match(re);
    return m ? m[1] : null;
  }

  /**
   * Scan all tracked changes and comments to find unique contributors.
   * Returns authors with counts and last-active dates.
   *
   * @param {object} ws - Workspace with ws.docXml and ws.commentsXml
   * @returns {Array<{name: string, changes: number, comments: number, lastActive: string}>}
   */
  static contributors(ws) {
    const authors = {};

    // Count tracked changes
    const revisions = Revisions.list(ws);
    for (const rev of revisions) {
      const name = rev.author || 'Unknown';
      if (!authors[name]) {
        authors[name] = { name, changes: 0, comments: 0, lastActive: '' };
      }
      authors[name].changes++;
      if (rev.date && rev.date > authors[name].lastActive) {
        authors[name].lastActive = rev.date;
      }
    }

    // Count comments (lazy require to avoid circular dependency)
    const { Comments } = require('./comments');
    const comments = Comments.list(ws);
    for (const c of comments) {
      const name = c.author || 'Unknown';
      if (!authors[name]) {
        authors[name] = { name, changes: 0, comments: 0, lastActive: '' };
      }
      authors[name].comments++;
      if (c.date && c.date > authors[name].lastActive) {
        authors[name].lastActive = c.date;
      }
    }

    // Sort by total activity descending
    return Object.values(authors).sort((a, b) =>
      (b.changes + b.comments) - (a.changes + a.comments)
    );
  }

  /**
   * Combine comment dates and revision dates into a single chronological timeline.
   *
   * @param {object} ws - Workspace with ws.docXml and ws.commentsXml
   * @returns {Array<{date: string, type: string, author: string, text: string}>}
   */
  static timeline(ws) {
    const events = [];

    // Add tracked changes
    const revisions = Revisions.list(ws);
    for (const rev of revisions) {
      events.push({
        date: rev.date || '',
        type: rev.type,
        author: rev.author,
        text: rev.text,
      });
    }

    // Add comments
    const { Comments } = require('./comments');
    const comments = Comments.list(ws);
    for (const c of comments) {
      events.push({
        date: c.date || '',
        type: 'comment',
        author: c.author,
        text: c.text,
      });
    }

    // Sort chronologically
    events.sort((a, b) => {
      if (a.date < b.date) return -1;
      if (a.date > b.date) return 1;
      return 0;
    });

    return events;
  }
}

module.exports = { Revisions };
