/**
 * response-letter.js -- Generate R&R response letters for docex
 *
 * Creates a formatted .docx response letter from extracted comments
 * and author responses. Groups comments by reviewer, formats them
 * with clear visual distinction between original comments and responses.
 *
 * Zero external dependencies beyond docex internals.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const { Workspace } = require('./workspace');
const xml = require('./xml');

// ============================================================================
// HELPERS
// ============================================================================

/**
 * Generate a random 8-char hex paraId.
 * @returns {string}
 */
function _genParaId() {
  return crypto.randomBytes(4).toString('hex').toUpperCase();
}

/**
 * Create a paragraph with specific formatting.
 * @param {string} text - Paragraph text
 * @param {object} [opts] - { bold, italic, color, size, style }
 * @returns {string}
 */
function _para(text, opts = {}) {
  let pPr = '';
  if (opts.style) {
    pPr += `<w:pStyle w:val="${opts.style}"/>`;
  }
  if (opts.spacing) {
    pPr += `<w:spacing w:before="${opts.spacing}"/>`;
  }

  let rPr = '';
  if (opts.bold) rPr += '<w:b/>';
  if (opts.italic) rPr += '<w:i/>';
  if (opts.color) rPr += `<w:color w:val="${opts.color}"/>`;
  if (opts.size) rPr += `<w:sz w:val="${opts.size}"/><w:szCs w:val="${opts.size}"/>`;
  if (opts.underline) rPr += '<w:u w:val="single"/>';

  const rPrXml = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
  const pPrXml = pPr ? `<w:pPr>${pPr}</w:pPr>` : '';

  return `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}">`
    + pPrXml
    + `<w:r>${rPrXml}<w:t xml:space="preserve">${xml.escapeXml(text)}</w:t></w:r>`
    + '</w:p>';
}

/**
 * Create an empty paragraph (spacer).
 * @returns {string}
 */
function _emptyPara() {
  return `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}"/>`;
}

// ============================================================================
// RESPONSE LETTER CLASS
// ============================================================================

class ResponseLetter {

  /**
   * Generate a response letter .docx from comments and responses.
   *
   * @param {Array<{id: number, author: string, text: string, date: string}>} comments - From doc.comments()
   * @param {object} responses - Map of comment ID to response info:
   *   { [commentId]: { action: "agree"|"partial"|"disagree", text: "...", changes: ["..."] } }
   * @param {object} [opts] - Options
   * @param {string} [opts.title] - Manuscript title
   * @param {string} [opts.journal] - Journal name
   * @param {string[]} [opts.authors] - Author names
   * @param {string} [opts.output] - Output .docx path
   * @returns {Promise<{path: string, fileSize: number, paragraphCount: number, reviewers: number, commentsAddressed: number}>}
   */
  static async generate(comments, responses, opts = {}) {
    if (!Array.isArray(comments)) {
      throw new Error('comments must be an array (from doc.comments())');
    }
    if (!responses || typeof responses !== 'object') {
      throw new Error('responses must be an object mapping comment IDs to response info');
    }

    // Group comments by author (reviewer)
    const reviewerMap = new Map();
    for (const comment of comments) {
      const author = comment.author || 'Unknown Reviewer';
      if (!reviewerMap.has(author)) {
        reviewerMap.set(author, []);
      }
      reviewerMap.get(author).push(comment);
    }

    // Build document body
    let bodyXml = '';

    // Header: Response to Reviewer Comments
    bodyXml += _para('Response to Reviewer Comments', {
      bold: true, size: '32', style: 'Heading1', spacing: '0',
    });
    bodyXml += _emptyPara();

    // Manuscript info
    if (opts.title) {
      bodyXml += _para('Manuscript: ' + opts.title, { italic: true });
    }
    if (opts.journal) {
      bodyXml += _para('Journal: ' + opts.journal, { italic: true });
    }
    if (opts.authors && opts.authors.length > 0) {
      bodyXml += _para('Authors: ' + opts.authors.join(', '), { italic: true });
    }
    bodyXml += _emptyPara();

    // Date
    const dateStr = new Date().toLocaleDateString('en-US', {
      year: 'numeric', month: 'long', day: 'numeric',
    });
    bodyXml += _para('Date: ' + dateStr);
    bodyXml += _emptyPara();

    // Opening
    bodyXml += _para(
      'We thank the reviewers for their thoughtful and constructive feedback. '
      + 'Below, we address each comment in turn, describing the changes made to the manuscript.',
    );
    bodyXml += _emptyPara();

    // For each reviewer
    let reviewerNum = 0;
    let commentsAddressed = 0;

    for (const [reviewerName, reviewerComments] of reviewerMap) {
      reviewerNum++;

      // Reviewer heading
      bodyXml += _para(`Reviewer ${reviewerNum}: ${reviewerName}`, {
        bold: true, size: '28', style: 'Heading1', spacing: '240',
      });
      bodyXml += _emptyPara();

      // For each comment by this reviewer
      let commentNum = 0;
      for (const comment of reviewerComments) {
        commentNum++;

        // Comment label
        bodyXml += _para(`Comment ${commentNum}:`, { bold: true, underline: true });

        // Original comment in italics (gray)
        bodyXml += _para(comment.text, { italic: true, color: '666666' });
        bodyXml += _emptyPara();

        // Response
        const response = responses[comment.id];
        if (response) {
          commentsAddressed++;

          // Action badge
          const actionLabel = response.action === 'agree' ? 'Agreed.'
            : response.action === 'partial' ? 'Partially agreed.'
            : response.action === 'disagree' ? 'Respectfully disagree.'
            : 'Response:';

          bodyXml += _para(actionLabel, { bold: true, color: response.action === 'agree' ? '006600' : response.action === 'partial' ? 'CC6600' : 'CC0000' });

          // Response text
          if (response.text) {
            bodyXml += _para(response.text);
          }

          // Changes made
          if (response.changes && response.changes.length > 0) {
            bodyXml += _emptyPara();
            bodyXml += _para('Changes made:', { bold: true });
            for (const change of response.changes) {
              bodyXml += _para('  - ' + change);
            }
          }
        } else {
          bodyXml += _para('[No response provided]', { italic: true, color: 'CC0000' });
        }

        bodyXml += _emptyPara();
      }
    }

    // Build the full document
    const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
      + `xmlns:mc="${xml.NS.mc}" `
      + `xmlns:r="${xml.NS.r}" `
      + `xmlns:w="${xml.NS.w}" `
      + `xmlns:w14="${xml.NS.w14}" `
      + `mc:Ignorable="w14">`
      + '<w:body>'
      + bodyXml
      + '<w:sectPr>'
      + '<w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr>'
      + '</w:body></w:document>';

    // Build styles
    const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + `<w:styles xmlns:w="${xml.NS.w}" xmlns:r="${xml.NS.r}" `
      + `xmlns:w14="${xml.NS.w14}" xmlns:mc="${xml.NS.mc}" mc:Ignorable="w14">`
      + '<w:docDefaults>'
      + '<w:rPrDefault><w:rPr>'
      + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
      + '<w:sz w:val="24"/><w:szCs w:val="24"/>'
      + '</w:rPr></w:rPrDefault>'
      + '<w:pPrDefault><w:pPr>'
      + '<w:spacing w:line="360" w:lineRule="auto"/>'
      + '</w:pPr></w:pPrDefault>'
      + '</w:docDefaults>'
      + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
      + '<w:name w:val="Normal"/>'
      + '</w:style>'
      + '<w:style w:type="paragraph" w:styleId="Heading1">'
      + '<w:name w:val="heading 1"/>'
      + '<w:basedOn w:val="Normal"/>'
      + '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
      + '<w:rPr><w:b/></w:rPr>'
      + '</w:style>'
      + '</w:styles>';

    // Content types
    const contentTypesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '</Types>';

    // Root rels
    const rootRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>';

    // Doc rels
    const docRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      + '</Relationships>';

    // Build .docx in temp dir
    const tmpId = crypto.randomBytes(8).toString('hex');
    const tmpDir = path.join(os.tmpdir(), `docex-rl-${tmpId}`);
    const outputPath = opts.output
      ? path.resolve(opts.output)
      : path.join(os.tmpdir(), `docex-response-${tmpId}.docx`);

    try {
      fs.mkdirSync(path.join(tmpDir, '_rels'), { recursive: true });
      fs.mkdirSync(path.join(tmpDir, 'word', '_rels'), { recursive: true });

      fs.writeFileSync(path.join(tmpDir, '[Content_Types].xml'), contentTypesXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, '_rels', '.rels'), rootRelsXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', '_rels', 'document.xml.rels'), docRelsXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', 'styles.xml'), stylesXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', 'document.xml'), docXml, 'utf-8');

      execFileSync('zip', ['-r', '-q', outputPath, '.'], {
        cwd: tmpDir,
        stdio: 'pipe',
      });

      const stat = fs.statSync(outputPath);

      // Count paragraphs
      const paraCount = (docXml.match(/<w:p[\s>]/g) || []).length;

      return {
        path: outputPath,
        fileSize: stat.size,
        paragraphCount: paraCount,
        reviewers: reviewerMap.size,
        commentsAddressed,
      };
    } finally {
      try { execFileSync('rm', ['-rf', tmpDir], { stdio: 'pipe' }); } catch (_) { /* ignore */ }
    }
  }
}

module.exports = { ResponseLetter };
