/**
 * template.js -- Document creation from scratch for docex
 *
 * Creates new .docx files from journal-specific templates with:
 *   - Title page with author affiliations
 *   - Abstract with word limit indicator
 *   - Keywords
 *   - Standard section headings
 *   - Running header and page numbers
 *   - Proper journal formatting applied
 *
 * All methods produce minimal valid .docx files using raw OOXML.
 * Zero external dependencies.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const { Workspace } = require('./workspace');
const { Presets } = require('./presets');
const xml = require('./xml');

// ============================================================================
// MINIMAL OOXML TEMPLATES
// ============================================================================

/**
 * Generate a minimal valid document.xml with content.
 * @param {object} metadata - { title, authors, abstract, keywords, sections }
 * @returns {string}
 */
function _buildDocumentXml(metadata = {}) {
  const title = metadata.title || '';
  const authors = metadata.authors || [];
  const abstract = metadata.abstract || '';
  const keywords = metadata.keywords || [];
  const sections = metadata.sections || [
    'Introduction', 'Literature Review', 'Methods', 'Results', 'Discussion', 'Conclusion'
  ];

  let bodyXml = '';

  // Title
  if (title) {
    bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}"><w:pPr><w:pStyle w:val="Title"/><w:jc w:val="center"/></w:pPr>`
      + `<w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>`
      + `<w:t xml:space="preserve">${xml.escapeXml(title)}</w:t></w:r></w:p>`;
  }

  // Authors
  for (const author of authors) {
    const name = typeof author === 'string' ? author : author.name;
    const affiliation = typeof author === 'object' ? (author.affiliation || '') : '';
    const email = typeof author === 'object' ? (author.email || '') : '';

    let authorLine = xml.escapeXml(name);
    if (affiliation) authorLine += ', ' + xml.escapeXml(affiliation);
    if (email) authorLine += ' (' + xml.escapeXml(email) + ')';

    bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}"><w:pPr><w:jc w:val="center"/></w:pPr>`
      + `<w:r><w:t xml:space="preserve">${authorLine}</w:t></w:r></w:p>`;
  }

  // Empty line after authors
  if (authors.length > 0) {
    bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}"/>`;
  }

  // Abstract
  if (abstract) {
    bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>`
      + `<w:r><w:t>Abstract</w:t></w:r></w:p>`;
    bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}">`
      + `<w:r><w:t xml:space="preserve">${xml.escapeXml(abstract)}</w:t></w:r></w:p>`;
  }

  // Keywords
  if (keywords.length > 0) {
    bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}">`
      + `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Keywords: </w:t></w:r>`
      + `<w:r><w:t xml:space="preserve">${xml.escapeXml(keywords.join(', '))}</w:t></w:r></w:p>`;
  }

  // Empty line before sections
  bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}"/>`;

  // Section headings
  for (const section of sections) {
    bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>`
      + `<w:r><w:t>${xml.escapeXml(section)}</w:t></w:r></w:p>`;
    // Empty body paragraph after each heading
    bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}">`
      + `<w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>`;
  }

  // References heading
  bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>`
    + `<w:r><w:t>References</w:t></w:r></w:p>`;
  bodyXml += `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}">`
    + `<w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>`;

  // Section properties (page size, margins)
  const sectPr = '<w:sectPr>'
    + '<w:pgSz w:w="12240" w:h="15840"/>'
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr>';

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    + `xmlns:mc="${xml.NS.mc}" `
    + `xmlns:o="urn:schemas-microsoft-com:office:office" `
    + `xmlns:r="${xml.NS.r}" `
    + `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" `
    + `xmlns:v="urn:schemas-microsoft-com:vml" `
    + `xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" `
    + `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" `
    + `xmlns:w10="urn:schemas-microsoft-com:office:word" `
    + `xmlns:w="${xml.NS.w}" `
    + `xmlns:w14="${xml.NS.w14}" `
    + `xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" `
    + `xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" `
    + `xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" `
    + `xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" `
    + `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" `
    + `mc:Ignorable="w14 w15 wp14">`
    + '<w:body>'
    + bodyXml
    + sectPr
    + '</w:body></w:document>';
}

/**
 * Generate a minimal styles.xml with heading styles.
 * @returns {string}
 */
function _buildStylesXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + `<w:styles xmlns:w="${xml.NS.w}" xmlns:r="${xml.NS.r}" `
    + `xmlns:w14="${xml.NS.w14}" xmlns:mc="${xml.NS.mc}" mc:Ignorable="w14">`
    + '<w:docDefaults>'
    + '<w:rPrDefault><w:rPr>'
    + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
    + '<w:sz w:val="24"/><w:szCs w:val="24"/>'
    + '</w:rPr></w:rPrDefault>'
    + '<w:pPrDefault><w:pPr>'
    + '<w:spacing w:line="480" w:lineRule="auto"/>'
    + '</w:pPr></w:pPrDefault>'
    + '</w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    + '<w:name w:val="Normal"/>'
    + '<w:pPr><w:jc w:val="both"/></w:pPr>'
    + '</w:style>'
    + '<w:style w:type="paragraph" w:styleId="Title">'
    + '<w:name w:val="Title"/>'
    + '<w:basedOn w:val="Normal"/>'
    + '<w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>'
    + '<w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
    + '</w:style>'
    + '<w:style w:type="paragraph" w:styleId="Heading1">'
    + '<w:name w:val="heading 1"/>'
    + '<w:basedOn w:val="Normal"/>'
    + '<w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="240"/></w:pPr>'
    + '<w:rPr><w:b/></w:rPr>'
    + '</w:style>'
    + '<w:style w:type="paragraph" w:styleId="Heading2">'
    + '<w:name w:val="heading 2"/>'
    + '<w:basedOn w:val="Normal"/>'
    + '<w:pPr><w:outlineLvl w:val="1"/><w:spacing w:before="120"/></w:pPr>'
    + '<w:rPr><w:b/><w:i/></w:rPr>'
    + '</w:style>'
    + '</w:styles>';
}

/**
 * Generate a minimal [Content_Types].xml.
 * @returns {string}
 */
function _buildContentTypesXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '</Types>';
}

/**
 * Generate _rels/.rels (root relationships).
 * @returns {string}
 */
function _buildRootRelsXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships>';
}

/**
 * Generate word/_rels/document.xml.rels.
 * @returns {string}
 */
function _buildDocRelsXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    + '</Relationships>';
}

/**
 * Generate a random 8-char hex paraId.
 * @returns {string}
 */
function _genParaId() {
  return crypto.randomBytes(4).toString('hex').toUpperCase();
}

// ============================================================================
// TEMPLATE CLASS
// ============================================================================

class Template {

  /**
   * Create a new .docx from scratch using a journal preset and metadata.
   *
   * @param {string} presetName - Journal preset name (e.g. "polcomm", "apa7", "academic")
   * @param {object} metadata - Document metadata
   * @param {string} metadata.title - Paper title
   * @param {Array<string|{name: string, affiliation?: string, email?: string}>} [metadata.authors] - Authors
   * @param {string} [metadata.abstract] - Abstract text
   * @param {string[]} [metadata.keywords] - Keywords
   * @param {string[]} [metadata.sections] - Section headings (default: standard set)
   * @param {string} [metadata.output] - Output .docx path
   * @returns {Promise<{path: string, fileSize: number, paragraphCount: number}>}
   */
  static async create(presetName, metadata = {}) {
    // Validate preset exists
    const config = Presets.get(presetName);
    if (!config) {
      const available = Presets.list().join(', ');
      throw new Error(`Unknown preset: "${presetName}". Available: ${available}`);
    }

    // Build the .docx in a temp directory
    const id = crypto.randomBytes(8).toString('hex');
    const tmpDir = path.join(os.tmpdir(), `docex-tpl-${id}`);
    const tmpDocx = path.join(os.tmpdir(), `docex-tpl-${id}.docx`);

    try {
      // Create directory structure
      fs.mkdirSync(path.join(tmpDir, '_rels'), { recursive: true });
      fs.mkdirSync(path.join(tmpDir, 'word', '_rels'), { recursive: true });

      // Write all XML files
      fs.writeFileSync(path.join(tmpDir, '[Content_Types].xml'), _buildContentTypesXml(), 'utf-8');
      fs.writeFileSync(path.join(tmpDir, '_rels', '.rels'), _buildRootRelsXml(), 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', '_rels', 'document.xml.rels'), _buildDocRelsXml(), 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', 'styles.xml'), _buildStylesXml(), 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', 'document.xml'), _buildDocumentXml(metadata), 'utf-8');

      // Zip into .docx
      execFileSync('zip', ['-r', '-q', tmpDocx, '.'], {
        cwd: tmpDir,
        stdio: 'pipe',
      });

      // Open with docex workspace and apply preset
      const ws = Workspace.open(tmpDocx);
      Presets.apply(ws, presetName);

      // Save to final output path
      const outputPath = metadata.output
        ? path.resolve(metadata.output)
        : tmpDocx;

      const result = ws.save(outputPath);

      return {
        path: result.path,
        fileSize: result.fileSize,
        paragraphCount: result.paragraphCount,
      };
    } finally {
      // Clean up temp dir (keep tmpDocx if it's the output)
      try { execFileSync('rm', ['-rf', tmpDir], { stdio: 'pipe' }); } catch (_) { /* ignore */ }
      if (metadata.output) {
        try { if (fs.existsSync(tmpDocx)) fs.unlinkSync(tmpDocx); } catch (_) { /* ignore */ }
      }
    }
  }
}

// ============================================================================
// DOCEX CREATE -- Minimal empty document
// ============================================================================

/**
 * Create a minimal valid .docx in memory (written to a temp file).
 * Returns a docex engine instance that can be further manipulated.
 *
 * @param {object} [opts] - Options
 * @param {string} [opts.output] - Output path. If omitted, creates a temp file.
 * @returns {Promise<{path: string, fileSize: number, paragraphCount: number}>}
 */
async function createEmpty(opts = {}) {
  const id = crypto.randomBytes(8).toString('hex');
  const tmpDir = path.join(os.tmpdir(), `docex-create-${id}`);
  const outputPath = opts.output
    ? path.resolve(opts.output)
    : path.join(os.tmpdir(), `docex-create-${id}.docx`);

  try {
    // Create directory structure
    fs.mkdirSync(path.join(tmpDir, '_rels'), { recursive: true });
    fs.mkdirSync(path.join(tmpDir, 'word', '_rels'), { recursive: true });

    // Minimal document with one empty paragraph
    const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
      + `xmlns:mc="${xml.NS.mc}" `
      + `xmlns:r="${xml.NS.r}" `
      + `xmlns:w="${xml.NS.w}" `
      + `xmlns:w14="${xml.NS.w14}" `
      + `mc:Ignorable="w14">`
      + '<w:body>'
      + `<w:p w14:paraId="${_genParaId()}" w14:textId="${_genParaId()}">`
      + '<w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>'
      + '<w:sectPr>'
      + '<w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr>'
      + '</w:body></w:document>';

    fs.writeFileSync(path.join(tmpDir, '[Content_Types].xml'), _buildContentTypesXml(), 'utf-8');
    fs.writeFileSync(path.join(tmpDir, '_rels', '.rels'), _buildRootRelsXml(), 'utf-8');
    fs.writeFileSync(path.join(tmpDir, 'word', '_rels', 'document.xml.rels'), _buildDocRelsXml(), 'utf-8');
    fs.writeFileSync(path.join(tmpDir, 'word', 'styles.xml'), _buildStylesXml(), 'utf-8');
    fs.writeFileSync(path.join(tmpDir, 'word', 'document.xml'), docXml, 'utf-8');

    // Zip
    execFileSync('zip', ['-r', '-q', outputPath, '.'], {
      cwd: tmpDir,
      stdio: 'pipe',
    });

    const stat = fs.statSync(outputPath);
    return {
      path: outputPath,
      fileSize: stat.size,
      paragraphCount: 1,
    };
  } finally {
    try { execFileSync('rm', ['-rf', tmpDir], { stdio: 'pipe' }); } catch (_) { /* ignore */ }
  }
}

module.exports = { Template, createEmpty };
