/**
 * submission.js -- Submission helpers for docex
 *
 * Provides anonymization (blind review), deanonymization, and
 * highlighted changes document generation.
 *
 * All methods operate on a Workspace object.
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// SUBMISSION
// ============================================================================

class Submission {

  /**
   * Remove author names from the document for blind review.
   * Removes from: title page text, metadata, tracked changes, comments.
   * Stores removed info in ws._docexAuthorCache for deanonymize().
   *
   * @param {object} ws - Workspace
   * @returns {{ authorsRemoved: string[], locations: string[] }}
   */
  static anonymize(ws) {
    const locations = [];

    // 1. Collect all author names from comments and tracked changes
    const authorNames = new Set();

    // From comments
    const commentAuthorRe = /w:author="([^"]+)"/g;
    let m;
    while ((m = commentAuthorRe.exec(ws.commentsXml)) !== null) {
      if (m[1] !== 'Unknown') authorNames.add(m[1]);
    }

    // From tracked changes in document
    const docAuthorRe = /w:author="([^"]+)"/g;
    while ((m = docAuthorRe.exec(ws.docXml)) !== null) {
      if (m[1] !== 'Unknown') authorNames.add(m[1]);
    }

    // From core.xml metadata
    const coreXml = ws.corePropsXml;
    if (coreXml) {
      const creatorMatch = coreXml.match(/<dc:creator>([^<]+)<\/dc:creator>/);
      if (creatorMatch && creatorMatch[1] !== 'Unknown') {
        authorNames.add(creatorMatch[1]);
      }
      const lastModMatch = coreXml.match(/<cp:lastModifiedBy>([^<]+)<\/cp:lastModifiedBy>/);
      if (lastModMatch && lastModMatch[1] !== 'Unknown') {
        authorNames.add(lastModMatch[1]);
      }
    }

    if (authorNames.size === 0) {
      return { authorsRemoved: [], locations };
    }

    // Store cache for deanonymize
    ws._docexAuthorCache = {
      authors: [...authorNames],
      originalCoreXml: coreXml || null,
    };

    const authorList = [...authorNames];
    const anonymousName = 'Anonymous';

    // 2. Replace author names in tracked changes
    let docXml = ws.docXml;
    for (const author of authorList) {
      const escaped = xml.escapeXml(author);
      const pattern = `w:author="${escaped}"`;
      const replacement = `w:author="${anonymousName}"`;
      if (docXml.includes(pattern)) {
        docXml = docXml.split(pattern).join(replacement);
        locations.push('tracked changes');
      }
    }
    ws.docXml = docXml;

    // 3. Replace author names in comments
    let commentsXml = ws.commentsXml;
    for (const author of authorList) {
      const escaped = xml.escapeXml(author);
      commentsXml = commentsXml.split(`w:author="${escaped}"`).join(`w:author="${anonymousName}"`);
    }
    ws.commentsXml = commentsXml;

    // 4. Replace in commentsExtended
    let commentsExtXml = ws.commentsExtXml;
    for (const author of authorList) {
      const escaped = xml.escapeXml(author);
      commentsExtXml = commentsExtXml.split(`w15:author="${escaped}"`).join(`w15:author="${anonymousName}"`);
    }
    ws.commentsExtXml = commentsExtXml;

    // 5. Clear metadata
    if (coreXml) {
      let newCoreXml = coreXml;
      for (const author of authorList) {
        const escapedRe = xml.escapeXml(author).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        newCoreXml = newCoreXml.replace(
          new RegExp(`<dc:creator>${escapedRe}</dc:creator>`),
          `<dc:creator>${anonymousName}</dc:creator>`
        );
        newCoreXml = newCoreXml.replace(
          new RegExp(`<cp:lastModifiedBy>${escapedRe}</cp:lastModifiedBy>`),
          `<cp:lastModifiedBy>${anonymousName}</cp:lastModifiedBy>`
        );
      }
      ws.corePropsXml = newCoreXml;
      locations.push('metadata');
    }

    // 6. Replace author names in document body text
    const paragraphs = xml.findParagraphs(ws.docXml);
    docXml = ws.docXml;
    let bodyReplacements = 0;

    for (let i = paragraphs.length - 1; i >= 0; i--) {
      const p = paragraphs[i];
      const text = xml.extractTextDecoded(p.xml);

      for (const author of authorList) {
        if (text.includes(author)) {
          const escapedAuthor = xml.escapeXml(author);
          let newParaXml = p.xml;
          if (newParaXml.includes(escapedAuthor)) {
            newParaXml = newParaXml.split(escapedAuthor).join('[Author]');
            docXml = docXml.slice(0, p.start) + newParaXml + docXml.slice(p.end);
            bodyReplacements++;
          }
        }
      }
    }

    if (bodyReplacements > 0) {
      ws.docXml = docXml;
      locations.push('body text');
    }

    return {
      authorsRemoved: authorList,
      locations: [...new Set(locations)],
    };
  }

  /**
   * Restore author info from the cache created by anonymize().
   *
   * @param {object} ws - Workspace
   * @returns {{ restored: boolean, authors: string[] }}
   */
  static deanonymize(ws) {
    if (!ws._docexAuthorCache) {
      return { restored: false, authors: [] };
    }

    const cache = ws._docexAuthorCache;
    const primaryAuthor = cache.authors[0] || 'Unknown';

    // 1. Restore metadata
    if (cache.originalCoreXml) {
      ws.corePropsXml = cache.originalCoreXml;
    }

    // 2. Restore author names in tracked changes and comments
    let docXml = ws.docXml;
    docXml = docXml.split('w:author="Anonymous"').join(`w:author="${xml.escapeXml(primaryAuthor)}"`);
    ws.docXml = docXml;

    let commentsXml = ws.commentsXml;
    commentsXml = commentsXml.split('w:author="Anonymous"').join(`w:author="${xml.escapeXml(primaryAuthor)}"`);
    ws.commentsXml = commentsXml;

    // 3. Restore body text
    let bodyDocXml = ws.docXml;
    bodyDocXml = bodyDocXml.split('[Author]').join(xml.escapeXml(primaryAuthor));
    ws.docXml = bodyDocXml;

    // Clear cache
    delete ws._docexAuthorCache;

    return {
      restored: true,
      authors: cache.authors,
    };
  }

  /**
   * Highlight all tracked insertions in yellow and mark deletions in red.
   * For "highlighted changes" documents that reviewers sometimes request.
   *
   * @param {object} ws - Workspace
   * @returns {{ insertions: number, deletions: number }}
   */
  static highlightedChanges(ws) {
    let docXml = ws.docXml;
    let insertions = 0;
    let deletions = 0;

    // 1. Highlight tracked insertions (w:ins elements) in yellow
    const insRe = /<w:ins\b[^>]*>([\s\S]*?)<\/w:ins>/g;
    docXml = docXml.replace(insRe, (match, content) => {
      insertions++;
      // Add yellow highlight to all runs inside the insertion
      let highlighted = content.replace(/<w:rPr>/g, '<w:rPr><w:highlight w:val="yellow"/>');
      // If runs have no rPr, add one
      highlighted = highlighted.replace(
        /<w:r>(\s*)<w:t/g,
        '<w:r>$1<w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t'
      );
      return match.replace(content, highlighted);
    });

    // 2. Highlight tracked deletions (w:del elements) in red
    const delRe = /<w:del\b[^>]*>([\s\S]*?)<\/w:del>/g;
    docXml = docXml.replace(delRe, (match, content) => {
      deletions++;
      // Replace w:delText with w:t and add red color + strikethrough
      let visible = content.replace(/<w:delText([^>]*)>([^<]*)<\/w:delText>/g,
        '<w:t$1>$2</w:t>');
      visible = visible.replace(/<w:rPr>/g,
        '<w:rPr><w:color w:val="FF0000"/><w:strike/>');
      visible = visible.replace(
        /<w:r>(\s*)<w:t/g,
        '<w:r>$1<w:rPr><w:color w:val="FF0000"/><w:strike/></w:rPr><w:t'
      );
      // Replace w:del wrapper with plain content (no longer tracked)
      return visible;
    });

    ws.docXml = docXml;
    return { insertions, deletions };
  }
}

module.exports = { Submission };
