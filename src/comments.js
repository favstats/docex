/**
 * comments.js -- Comment operations for docex
 *
 * Static methods for listing, adding, replying to, resolving, and removing
 * comments in OOXML documents. Manages five XML files:
 *   - word/document.xml (comment ranges and references)
 *   - word/comments.xml (comment content)
 *   - word/commentsExtended.xml (paraId linkage, done state)
 *   - word/commentsIds.xml (durable IDs)
 *   - word/_rels/document.xml.rels + [Content_Types].xml (infrastructure)
 *
 * All methods operate on a Workspace object. XML manipulation is done
 * entirely with string operations and regex. Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ---------------------------------------------------------------------------
// Relationship type constants
// ---------------------------------------------------------------------------
const REL_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const REL_COMMENTS_EXT = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const REL_COMMENTS_IDS = 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds';

const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const CT_COMMENTS_EXT = 'application/vnd.ms-word.commentsExtended+xml';
const CT_COMMENTS_IDS = 'application/vnd.ms-word.commentsIds+xml';

// ============================================================================
// COMMENTS
// ============================================================================

class Comments {

  /**
   * List all comments in the document.
   *
   * @param {object} ws - Workspace with ws.commentsXml, ws.commentsExtXml
   * @returns {Array<{id: number, author: string, text: string, date: string, paraId: string}>}
   */
  static list(ws) {
    const commentsXml = ws.commentsXml;
    if (!commentsXml) return [];

    const results = [];
    const commentRe = /<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g;
    let m;

    while ((m = commentRe.exec(commentsXml)) !== null) {
      const attrs = m[1];
      const body = m[2];

      const id = Comments._attrVal(attrs, 'w:id');
      const author = Comments._attrVal(attrs, 'w:author') || '';
      const date = Comments._attrVal(attrs, 'w:date') || '';
      const text = xml.extractText(body);

      // Look up paraId from commentsExtended (parallel ordering)
      let paraId = '';
      const extXml = ws.commentsExtXml;
      if (extXml) {
        const exRe = /<w15:commentEx\s+([^>]*?)\s*\/?>/g;
        const exEntries = [];
        let exM;
        while ((exM = exRe.exec(extXml)) !== null) {
          exEntries.push(exM[1]);
        }
        // Comments and commentEx entries are in parallel order
        if (exEntries.length > results.length) {
          const exAttrs = exEntries[results.length];
          const pid = Comments._attrVal(exAttrs, 'w15:paraId');
          if (pid) paraId = pid;
        }
      }

      results.push({
        id: id ? parseInt(id, 10) : 0,
        author: xml.decodeXml(author),
        text,
        date,
        paraId,
      });
    }

    return results;
  }

  /**
   * Add a new comment anchored to specific text in the document.
   *
   * This modifies up to 5 XML files:
   *   1. document.xml: commentRangeStart, commentRangeEnd, commentReference
   *   2. comments.xml: the comment element with text
   *   3. commentsExtended.xml: commentEx with paraId
   *   4. commentsIds.xml: commentId with durableId
   *   5. Relationships and content types (if not already present)
   *
   * @param {object} ws - Workspace
   * @param {string} anchor - Text to anchor the comment to
   * @param {string} commentText - Comment text
   * @param {object} [opts] - Options
   * @param {string} [opts.author] - Comment author
   * @param {string} [opts.by] - Comment author (alias)
   * @param {string} [opts.initials] - Author initials
   * @param {string} [opts.date] - ISO date string
   * @throws {Error} If anchor text is not found
   * @returns {{ commentId: number }} The new comment's ID
   */
  static add(ws, anchor, commentText, opts = {}) {
    if (typeof commentText !== 'string' || commentText.length === 0) {
      throw new Error('add(): text must be a non-empty string');
    }

    const author = opts.by || opts.author || 'docex';
    const initials = opts.initials || _makeInitials(author);
    const date = opts.date || xml.isoNow();

    // Ensure all comment files exist
    _ensureCommentFiles(ws);

    // Get next comment ID
    const commentsXml = ws.commentsXml;
    const docXml = ws.docXml;
    const commentId = Math.max(xml.nextCommentId(commentsXml), xml.nextChangeId(docXml));
    const paraId = xml.randomHexId().toUpperCase();
    const textId = xml.randomHexId().toUpperCase();

    // 1. Add to word/comments.xml
    const commentEl = '<w:comment w:id="' + commentId
      + '" w:author="' + xml.escapeXml(author)
      + '" w:date="' + date
      + '" w:initials="' + xml.escapeXml(initials) + '">'
      + '<w:p w14:paraId="' + paraId + '" w14:textId="' + textId + '">'
      + '<w:pPr><w:pStyle w:val="CommentText"/><w:rPr/></w:pPr>'
      + '<w:r>'
      + '<w:rPr>'
      + '<w:rStyle w:val="CommentReference"/>'
      + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
      + '<w:sz w:val="20"/><w:szCs w:val="20"/>'
      + '</w:rPr>'
      + '<w:annotationRef/>'
      + '</w:r>'
      + '<w:r>'
      + '<w:rPr>'
      + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
      + '<w:sz w:val="20"/><w:szCs w:val="20"/>'
      + '</w:rPr>'
      + '<w:t xml:space="preserve">' + xml.escapeXml(commentText) + '</w:t>'
      + '</w:r>'
      + '</w:p>'
      + '</w:comment>';

    let comments = ws.commentsXml;
    comments = comments.replace('</w:comments>', commentEl + '</w:comments>');
    ws.commentsXml = comments;

    // 2. Add to word/commentsExtended.xml
    let extXml = ws.commentsExtXml;
    if (extXml) {
      const extEl = '<w15:commentEx w15:paraId="' + paraId + '" w15:done="0"/>';
      extXml = extXml.replace('</w15:commentsEx>', extEl + '</w15:commentsEx>');
      ws.commentsExtXml = extXml;
    }

    // 3. Add to word/commentsIds.xml
    let idsXml = ws.commentsIdsXml;
    if (idsXml) {
      const durableId = xml.randomHexId().toUpperCase();
      const idEl = '<w16cid:commentId w16cid:paraId="' + paraId
        + '" w16cid:durableId="' + durableId + '"/>';
      idsXml = idsXml.replace('</w16cid:commentsIds>', idEl + '</w16cid:commentsIds>');
      ws.commentsIdsXml = idsXml;
    }

    // 4. Add anchor markers in document.xml
    if (anchor) {
      let doc = ws.docXml;
      const paragraphs = xml.findParagraphs(doc);
      let anchorFound = false;

      for (const para of paragraphs) {
        const paraText = para.text;
        if (!paraText.includes(anchor)) continue;

        const paraXml = para.xml;

        // Find which runs contain the anchor text
        const runs = xml.parseRuns(paraXml);
        const textRuns = runs.filter(r => r.texts.length > 0);
        let charCount = 0;
        let startRunIdx = -1;
        let endRunIdx = -1;

        const anchorStart = paraText.indexOf(anchor);
        const anchorEnd = anchorStart + anchor.length;

        for (let r = 0; r < textRuns.length; r++) {
          const runLen = textRuns[r].combinedText.length;
          const runEnd = charCount + runLen;

          if (startRunIdx === -1 && runEnd > anchorStart) {
            startRunIdx = r;
          }
          if (startRunIdx !== -1 && runEnd >= anchorEnd) {
            endRunIdx = r;
            break;
          }
          charCount += runLen;
        }

        if (startRunIdx === -1) startRunIdx = 0;
        if (endRunIdx === -1) endRunIdx = textRuns.length > 0 ? textRuns.length - 1 : 0;

        const rangeStart = '<w:commentRangeStart w:id="' + commentId + '"/>';
        const rangeEnd = '<w:commentRangeEnd w:id="' + commentId + '"/>';
        const refRun = '<w:r>'
          + '<w:rPr>'
          + '<w:rStyle w:val="CommentReference"/>'
          + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          + '<w:sz w:val="20"/><w:szCs w:val="20"/>'
          + '</w:rPr>'
          + '<w:commentReference w:id="' + commentId + '"/>'
          + '</w:r>';

        let modPara = paraXml;

        // Insert rangeEnd and reference after the last matched run
        if (textRuns.length > 0 && endRunIdx < textRuns.length) {
          const endRun = textRuns[endRunIdx];
          const insertAfterPos = endRun.index + endRun.fullMatch.length;
          modPara = modPara.slice(0, insertAfterPos) + rangeEnd + refRun + modPara.slice(insertAfterPos);
        }

        // Insert rangeStart before the first matched run
        if (textRuns.length > 0 && startRunIdx < textRuns.length) {
          const startRun = textRuns[startRunIdx];
          const insertBeforePos = startRun.index;
          modPara = modPara.slice(0, insertBeforePos) + rangeStart + modPara.slice(insertBeforePos);
        }

        doc = doc.slice(0, para.start) + modPara + doc.slice(para.end);
        ws.docXml = doc;
        anchorFound = true;
        break;
      }

      if (!anchorFound) {
        throw new Error('add(): could not find anchor text "' + anchor + '" in document');
      }
    }

    return { commentId };
  }

  /**
   * Reply to an existing comment (threaded reply via commentsExtended.xml).
   *
   * Replies do NOT add ranges in document.xml. They are linked to the
   * parent comment via paraIdParent in commentsExtended.xml.
   *
   * @param {object} ws - Workspace
   * @param {number|string} anchorOrId - Comment ID (number) or anchor text (string)
   * @param {string} commentText - Reply text
   * @param {object} [opts] - Options
   * @param {string} [opts.author] - Reply author
   * @param {string} [opts.by] - Reply author (alias)
   * @param {string} [opts.initials] - Author initials
   * @param {string} [opts.date] - ISO date string
   * @throws {Error} If parent comment is not found
   * @returns {{ commentId: number }} The reply comment's ID
   */
  static reply(ws, anchorOrId, commentText, opts = {}) {
    if (typeof commentText !== 'string' || commentText.length === 0) {
      throw new Error('reply(): text must be a non-empty string');
    }

    const author = opts.by || opts.author || 'docex';
    const initials = opts.initials || _makeInitials(author);
    const date = opts.date || xml.isoNow();

    _ensureCommentFiles(ws);

    const commentsXml = ws.commentsXml;

    // Find parent comment paraId
    let parentParaId = null;

    if (typeof anchorOrId === 'number') {
      // Find by numeric ID
      const re = new RegExp('<w:comment\\b[^>]*\\bw:id="' + anchorOrId + '"[^>]*>');
      const m = re.exec(commentsXml);
      if (m) {
        const paraIdMatch = m[0].match(/w14:paraId="([^"]*)"/);
        parentParaId = paraIdMatch ? paraIdMatch[1] : null;
      }
      // If comment element doesn't have w14:paraId, check commentsExtXml
      if (!parentParaId) {
        parentParaId = Comments._getParaIdForComment(ws, anchorOrId);
      }
    } else {
      // Find by anchor text: look for a comment whose document anchor text matches
      const docXml = ws.docXml;
      const paragraphs = xml.findParagraphs(docXml);
      let foundCommentId = null;

      for (const p of paragraphs) {
        if (!p.text.includes(anchorOrId)) continue;
        // Look for commentRangeStart in this paragraph
        const rangeMatch = p.xml.match(/<w:commentRangeStart\s+w:id="(\d+)"/);
        if (rangeMatch) {
          foundCommentId = parseInt(rangeMatch[1], 10);
          break;
        }
        // Also check for commentReference
        const refMatch = p.xml.match(/<w:commentReference\s+w:id="(\d+)"/);
        if (refMatch) {
          foundCommentId = parseInt(refMatch[1], 10);
          break;
        }
      }

      if (foundCommentId !== null) {
        // Now find the paraId for this comment
        const re = new RegExp('<w:comment\\b[^>]*\\bw:id="' + foundCommentId + '"[^>]*>');
        const m = re.exec(commentsXml);
        if (m) {
          const paraIdMatch = m[0].match(/w14:paraId="([^"]*)"/);
          parentParaId = paraIdMatch ? paraIdMatch[1] : null;
        }
        if (!parentParaId) {
          parentParaId = Comments._getParaIdForComment(ws, foundCommentId);
        }
      }
    }

    if (!parentParaId) {
      throw new Error('reply(): parent comment not found for: ' + String(anchorOrId).slice(0, 80));
    }

    // Generate IDs for the reply
    const docXml = ws.docXml;
    const commentId = Math.max(xml.nextCommentId(commentsXml), xml.nextChangeId(docXml));
    const replyParaId = xml.randomHexId().toUpperCase();
    const replyTextId = xml.randomHexId().toUpperCase();

    // 1. Add reply to word/comments.xml (no ranges in document.xml)
    const replyEl = '<w:comment w:id="' + commentId
      + '" w:author="' + xml.escapeXml(author)
      + '" w:date="' + date
      + '" w:initials="' + xml.escapeXml(initials) + '">'
      + '<w:p w14:paraId="' + replyParaId + '" w14:textId="' + replyTextId + '">'
      + '<w:pPr><w:pStyle w:val="CommentText"/><w:rPr/></w:pPr>'
      + '<w:r>'
      + '<w:rPr>'
      + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
      + '<w:sz w:val="20"/><w:szCs w:val="20"/>'
      + '</w:rPr>'
      + '<w:t xml:space="preserve">' + xml.escapeXml(commentText) + '</w:t>'
      + '</w:r>'
      + '</w:p>'
      + '</w:comment>';

    let comments = ws.commentsXml;
    comments = comments.replace('</w:comments>', replyEl + '</w:comments>');
    ws.commentsXml = comments;

    // 2. Add threading info to commentsExtended.xml
    let extXml = ws.commentsExtXml;
    if (extXml) {
      const extEl = '<w15:commentEx w15:paraId="' + replyParaId
        + '" w15:paraIdParent="' + parentParaId + '" w15:done="0"/>';
      extXml = extXml.replace('</w15:commentsEx>', extEl + '</w15:commentsEx>');
      ws.commentsExtXml = extXml;
    }

    // 3. Add to commentsIds.xml
    let idsXml = ws.commentsIdsXml;
    if (idsXml) {
      const durableId = xml.randomHexId().toUpperCase();
      const idEl = '<w16cid:commentId w16cid:paraId="' + replyParaId
        + '" w16cid:durableId="' + durableId + '"/>';
      idsXml = idsXml.replace('</w16cid:commentsIds>', idEl + '</w16cid:commentsIds>');
      ws.commentsIdsXml = idsXml;
    }

    return { commentId };
  }

  /**
   * Resolve (mark as done) a comment.
   *
   * Sets done="1" on the commentEx element in commentsExtended.xml.
   *
   * @param {object} ws - Workspace
   * @param {number} commentId - Comment ID to resolve
   * @throws {Error} If comment is not found
   */
  static resolve(ws, commentId) {
    const extXml = ws.commentsExtXml;
    if (!extXml) {
      throw new Error('No commentsExtended.xml found');
    }

    // Find the comment's position in parallel ordering
    const commentIds = Comments._listCommentIds(ws);
    const idx = commentIds.indexOf(commentId);
    if (idx === -1) {
      throw new Error('Comment not found: ' + commentId);
    }

    // Find the idx-th commentEx entry and set done="1"
    const exRe = /<w15:commentEx\s+([^>]*?)\s*\/?>/g;
    let count = 0;
    let exMatch;
    let newExtXml = extXml;

    while ((exMatch = exRe.exec(extXml)) !== null) {
      if (count === idx) {
        const original = exMatch[0];
        let updated;
        if (original.includes('w15:done=')) {
          // Update existing done attribute
          updated = original.replace(/w15:done="[^"]*"/, 'w15:done="1"');
        } else {
          // Add done attribute before the closing
          if (original.endsWith('/>')) {
            updated = original.slice(0, -2) + ' w15:done="1"/>';
          } else {
            updated = original.slice(0, -1) + ' w15:done="1">';
          }
        }
        newExtXml = extXml.slice(0, exMatch.index) + updated + extXml.slice(exMatch.index + original.length);
        break;
      }
      count++;
    }

    ws.commentsExtXml = newExtXml;
  }

  /**
   * Remove a comment from all XML files.
   *
   * Removes from:
   *   - document.xml (commentRangeStart, commentRangeEnd, commentReference)
   *   - comments.xml (the comment element)
   *   - commentsExtended.xml (the commentEx element)
   *   - commentsIds.xml (the commentId element)
   *
   * @param {object} ws - Workspace
   * @param {number} commentId - Comment ID to remove
   */
  static remove(ws, commentId) {
    const idStr = String(commentId);

    // Determine the comment's index BEFORE modifying any files
    const commentIdx = Comments._listCommentIds(ws).indexOf(commentId);

    // 1. Remove from document.xml: ranges and reference
    let docXml = ws.docXml;
    // Remove commentRangeStart
    docXml = docXml.replace(
      new RegExp('<w:commentRangeStart\\s+w:id="' + idStr + '"\\s*/>', 'g'),
      ''
    );
    // Remove commentRangeEnd
    docXml = docXml.replace(
      new RegExp('<w:commentRangeEnd\\s+w:id="' + idStr + '"\\s*/>', 'g'),
      ''
    );
    // Remove commentReference run (with rPr)
    docXml = docXml.replace(
      new RegExp(
        '<w:r>[^<]*(?:<w:rPr>[\\s\\S]*?</w:rPr>)?[^<]*'
        + '<w:commentReference\\s+w:id="' + idStr + '"\\s*/>'
        + '[^<]*</w:r>',
        'g'
      ),
      ''
    );
    ws.docXml = docXml;

    // 2. Remove from comments.xml
    let commentsXml = ws.commentsXml;
    if (commentsXml) {
      const commentPattern = new RegExp(
        '<w:comment\\s+[^>]*w:id="' + idStr + '"[^>]*>[\\s\\S]*?</w:comment>',
        'g'
      );
      commentsXml = commentsXml.replace(commentPattern, '');
      ws.commentsXml = commentsXml;
    }

    // 3. Remove from commentsExtended.xml (by position)
    if (commentIdx !== -1) {
      let extXml = ws.commentsExtXml;
      if (extXml) {
        extXml = Comments._removeNthEntry(
          extXml,
          /<w15:commentEx\s+[^>]*?\s*\/?>/g,
          commentIdx
        );
        ws.commentsExtXml = extXml;
      }

      // 4. Remove from commentsIds.xml (by position)
      let idsXml = ws.commentsIdsXml;
      if (idsXml) {
        idsXml = Comments._removeNthEntry(
          idsXml,
          /<w16cid:commentId\s+[^>]*?\s*\/?>/g,
          commentIdx
        );
        ws.commentsIdsXml = idsXml;
      }
    }
  }

  // --------------------------------------------------------------------------
  // INTERNAL HELPERS
  // --------------------------------------------------------------------------

  /**
   * Extract an attribute value from an attribute string.
   *
   * @param {string} attrs - Attribute string (e.g., 'w:id="5" w:author="Alice"')
   * @param {string} name - Attribute name (e.g., 'w:id')
   * @returns {string|null}
   * @private
   */
  static _attrVal(attrs, name) {
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(escaped + '="([^"]*)"');
    const m = attrs.match(re);
    return m ? m[1] : null;
  }

  /**
   * List all comment IDs from comments.xml in document order.
   *
   * @param {object} ws - Workspace
   * @returns {number[]}
   * @private
   */
  static _listCommentIds(ws) {
    const commentsXml = ws.commentsXml;
    if (!commentsXml) return [];

    const ids = [];
    const re = /<w:comment\s+[^>]*w:id="(\d+)"/g;
    let m;
    while ((m = re.exec(commentsXml)) !== null) {
      ids.push(parseInt(m[1], 10));
    }
    return ids;
  }

  /**
   * Get the paraId for a comment by its position in comments.xml.
   * Comments and commentEx entries are in parallel order.
   *
   * @param {object} ws - Workspace
   * @param {number} commentId - Comment ID
   * @returns {string|null} paraId or null
   * @private
   */
  static _getParaIdForComment(ws, commentId) {
    const commentIds = Comments._listCommentIds(ws);
    const idx = commentIds.indexOf(commentId);
    if (idx === -1) return null;

    const extXml = ws.commentsExtXml;
    if (!extXml) return null;

    const exRe = /<w15:commentEx\s+([^>]*?)\s*\/?>/g;
    let count = 0;
    let exMatch;
    while ((exMatch = exRe.exec(extXml)) !== null) {
      if (count === idx) {
        return Comments._attrVal(exMatch[1], 'w15:paraId');
      }
      count++;
    }

    return null;
  }

  /**
   * Remove the nth match of a regex pattern from an XML string.
   *
   * @param {string} xmlStr - XML string to modify
   * @param {RegExp} pattern - Regex pattern (must have global flag)
   * @param {number} n - Zero-based index of the match to remove
   * @returns {string} Modified XML string
   * @private
   */
  static _removeNthEntry(xmlStr, pattern, n) {
    let count = 0;
    let match;
    // Reset the regex
    pattern.lastIndex = 0;
    while ((match = pattern.exec(xmlStr)) !== null) {
      if (count === n) {
        return xmlStr.slice(0, match.index) + xmlStr.slice(match.index + match[0].length);
      }
      count++;
    }
    return xmlStr;
  }
}

// ---------------------------------------------------------------------------
// Internal helpers (module-level)
// ---------------------------------------------------------------------------

/**
 * Ensure all comment-related XML files exist in the workspace.
 * Creates them (and their relationships) if they do not.
 *
 * @param {object} ws - The workspace
 * @private
 */
function _ensureCommentFiles(ws) {
  // Accessing ws.commentsXml creates it if missing (workspace handles this)
  // Force initialization by reading the property
  void ws.commentsXml;
  void ws.commentsExtXml;
  void ws.commentsIdsXml;

  // Ensure relationships exist in rels file
  let relsXml = ws.relsXml;
  let relsChanged = false;

  if (relsXml && !relsXml.includes(REL_COMMENTS)) {
    const rId = xml.nextRId(relsXml);
    const rel = '<Relationship Id="' + rId + '" Type="' + REL_COMMENTS + '" Target="comments.xml"/>';
    relsXml = relsXml.replace('</Relationships>', rel + '</Relationships>');
    relsChanged = true;
  }

  if (relsXml && !relsXml.includes(REL_COMMENTS_EXT)) {
    const rId = xml.nextRId(relsXml);
    const rel = '<Relationship Id="' + rId + '" Type="' + REL_COMMENTS_EXT + '" Target="commentsExtended.xml"/>';
    relsXml = relsXml.replace('</Relationships>', rel + '</Relationships>');
    relsChanged = true;
  }

  if (relsXml && !relsXml.includes(REL_COMMENTS_IDS)) {
    const rId = xml.nextRId(relsXml);
    const rel = '<Relationship Id="' + rId + '" Type="' + REL_COMMENTS_IDS + '" Target="commentsIds.xml"/>';
    relsXml = relsXml.replace('</Relationships>', rel + '</Relationships>');
    relsChanged = true;
  }

  if (relsChanged) {
    ws.relsXml = relsXml;
  }

  // Ensure content types exist
  let ctXml = ws.contentTypesXml;
  let ctChanged = false;

  if (ctXml && !ctXml.includes(CT_COMMENTS)) {
    ctXml = ctXml.replace('</Types>',
      '<Override PartName="/word/comments.xml" ContentType="' + CT_COMMENTS + '"/></Types>');
    ctChanged = true;
  }

  if (ctXml && !ctXml.includes(CT_COMMENTS_EXT)) {
    ctXml = ctXml.replace('</Types>',
      '<Override PartName="/word/commentsExtended.xml" ContentType="' + CT_COMMENTS_EXT + '"/></Types>');
    ctChanged = true;
  }

  if (ctXml && !ctXml.includes(CT_COMMENTS_IDS)) {
    ctXml = ctXml.replace('</Types>',
      '<Override PartName="/word/commentsIds.xml" ContentType="' + CT_COMMENTS_IDS + '"/></Types>');
    ctChanged = true;
  }

  if (ctChanged) {
    ws.contentTypesXml = ctXml;
  }
}

/**
 * Generate initials from an author name.
 *
 * @param {string} name
 * @returns {string}
 * @private
 */
function _makeInitials(name) {
  return name
    .split(/\s+/)
    .map(w => w.charAt(0).toUpperCase())
    .join('')
    .slice(0, 3) || 'DE';
}

module.exports = { Comments };
