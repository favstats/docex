/**
 * production.js -- Production and presentation features for docex
 *
 * Static methods for watermarking, stamping, page count estimation,
 * and cover page generation.
 *
 * All methods operate on a Workspace object via ws.docXml (get/set).
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// HELPERS
// ============================================================================

/**
 * Get the next relationship ID by scanning existing IDs in rels XML.
 * @param {string} relsXml
 * @returns {string} e.g. "rId99"
 */
function _nextRelId(relsXml) {
  const ids = [];
  const re = /Id="rId(\d+)"/g;
  let m;
  while ((m = re.exec(relsXml)) !== null) {
    ids.push(parseInt(m[1], 10));
  }
  const max = ids.length > 0 ? Math.max(...ids) : 0;
  return 'rId' + (max + 1);
}

/**
 * Count words in the document body.
 * @param {string} docXml
 * @returns {number}
 */
function _wordCount(docXml) {
  const paragraphs = xml.findParagraphs(docXml);
  let total = 0;
  for (const p of paragraphs) {
    const text = xml.extractTextDecoded(p.xml).trim();
    if (text) {
      total += text.split(/\s+/).length;
    }
  }
  return total;
}

/**
 * Count figures in the document.
 * @param {string} docXml
 * @returns {number}
 */
function _figureCount(docXml) {
  const inlines = (docXml.match(/<wp:inline[\s>]/g) || []).length;
  const anchors = (docXml.match(/<wp:anchor[\s>]/g) || []).length;
  return inlines + anchors;
}

/**
 * Count tables in the document.
 * @param {string} docXml
 * @returns {number}
 */
function _tableCount(docXml) {
  return (docXml.match(/<w:tbl[\s>]/g) || []).length;
}

/**
 * Detect line spacing from styles or document. Returns multiplier (1=single, 2=double).
 * @param {object} ws
 * @returns {number}
 */
function _detectSpacing(ws) {
  // Check styles.xml for default spacing
  const stylesXml = ws.stylesXml || '';
  // Look for w:spacing w:line in the default paragraph style
  const lineMatch = stylesXml.match(/<w:spacing[^>]*w:line="(\d+)"/);
  if (lineMatch) {
    const lineVal = parseInt(lineMatch[1], 10);
    // Word line spacing: 240 = single, 360 = 1.5, 480 = double
    if (lineVal >= 440) return 2;
    if (lineVal >= 320) return 1.5;
    return 1;
  }
  return 1; // default to single
}

// ============================================================================
// VML WATERMARK
// ============================================================================

/**
 * Generate VML shape XML for a diagonal watermark text.
 *
 * @param {string} text - Watermark text
 * @param {object} opts - { color, size, angle }
 * @returns {string} VML shape XML
 */
function _vmlWatermark(text, opts) {
  const color = opts.color || 'C0C0C0';
  const size = opts.size || 72;
  const angle = opts.angle !== undefined ? opts.angle : -45;

  // VML shape for rotated text watermark
  return ''
    + '<w:r>'
    + '<w:rPr><w:noProof/></w:rPr>'
    + '<w:pict>'
    + `<v:shapetype id="_x0000_t136" coordsize="21600,21600" `
    + `o:spt="136" adj="10800" `
    + `path="m@7,l@8,m@5,21600l@6,21600e">`
    + '<v:formulas>'
    + '<v:f eqn="sum #0 0 10800"/>'
    + '<v:f eqn="prod #0 2 1"/>'
    + '<v:f eqn="sum 21600 0 @1"/>'
    + '<v:f eqn="sum 0 0 @2"/>'
    + '<v:f eqn="sum 21600 0 @3"/>'
    + '<v:f eqn="if @0 @3 0"/>'
    + '<v:f eqn="if @0 21600 @1"/>'
    + '<v:f eqn="if @0 0 @2"/>'
    + '<v:f eqn="if @0 @4 21600"/>'
    + '</v:formulas>'
    + '<v:path textpathok="t" o:connecttype="custom" '
    + 'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" '
    + 'o:connectangles="270,180,90,0"/>'
    + '<v:textpath on="t" fitshape="t"/>'
    + '<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>'
    + '<o:lock v:ext="edit" text="t" shapetype="t"/>'
    + '</v:shapetype>'
    + `<v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" `
    + `type="#_x0000_t136" `
    + `style="position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;`
    + `rotation:${angle};z-index:-251657216;mso-position-horizontal:center;`
    + `mso-position-horizontal-relative:margin;mso-position-vertical:center;`
    + `mso-position-vertical-relative:margin" `
    + `o:allowincell="f" fillcolor="#${color}" stroked="f">`
    + `<v:fill opacity=".5"/>`
    + `<v:textpath style="font-family:&quot;Calibri&quot;;font-size:${size}pt" `
    + `string="${xml.escapeXml(text)}"/>`
    + '</v:shape>'
    + '</w:pict>'
    + '</w:r>';
}

// ============================================================================
// PRODUCTION CLASS
// ============================================================================

class Production {

  /**
   * Add a diagonal watermark text to every page.
   * Uses Word header with VML shape (rotated text).
   *
   * @param {object} ws - Workspace with ws.docXml, ws.relsXml, ws.contentTypesXml
   * @param {string} text - Watermark text (e.g. "DRAFT")
   * @param {object} [opts] - Options
   * @param {string} [opts.color='C0C0C0'] - Hex color without #
   * @param {number} [opts.size=72] - Font size in points
   * @param {number} [opts.angle=-45] - Rotation angle in degrees
   */
  static watermark(ws, text, opts = {}) {
    const vmlShape = _vmlWatermark(text, opts);

    // Build a header XML containing the watermark
    const headerXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + `<w:hdr xmlns:w="${xml.NS.w}" xmlns:r="${xml.NS.r}" `
      + `xmlns:v="urn:schemas-microsoft-com:vml" `
      + `xmlns:o="urn:schemas-microsoft-com:office:office" `
      + `xmlns:w10="urn:schemas-microsoft-com:office:word">`
      + '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>'
      + vmlShape
      + '</w:p>'
      + '</w:hdr>';

    // Write the header file
    const headerFileName = 'headerWatermark1.xml';
    ws._writeFile('word/' + headerFileName, headerXml);

    // Add relationship
    const relId = _nextRelId(ws.relsXml);
    const newRel = `<Relationship Id="${relId}" `
      + `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" `
      + `Target="${headerFileName}"/>`;
    ws.relsXml = ws.relsXml.replace('</Relationships>', newRel + '</Relationships>');

    // Add content type override for the header
    if (!ws.contentTypesXml.includes(headerFileName)) {
      const override = `<Override PartName="/word/${headerFileName}" `
        + `ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`;
      ws.contentTypesXml = ws.contentTypesXml.replace('</Types>', override + '</Types>');
    }

    // Add header reference to the first section properties
    const headerRef = `<w:headerReference w:type="default" r:id="${relId}"/>`;
    const docXml = ws.docXml;

    // Find the last w:sectPr (main section properties)
    const sectPrMatch = docXml.match(/<w:sectPr[\s>][^]*?<\/w:sectPr>/);
    if (sectPrMatch) {
      const sectPr = sectPrMatch[0];
      // Check if there's already a default header reference
      if (sectPr.includes('w:type="default"') && sectPr.includes('headerReference')) {
        // Replace existing default header reference
        ws.docXml = docXml.replace(
          /<w:headerReference w:type="default"[^/]*\/>/,
          headerRef
        );
      } else {
        // Insert header reference at the beginning of sectPr
        const insertPoint = sectPr.indexOf('>') + 1;
        const newSectPr = sectPr.slice(0, insertPoint) + headerRef + sectPr.slice(insertPoint);
        ws.docXml = docXml.replace(sectPr, newSectPr);
      }
    }
  }

  /**
   * Add text stamp to every page header or footer.
   *
   * @param {object} ws - Workspace with ws.docXml, ws.relsXml, ws.contentTypesXml
   * @param {string} text - Stamp text (e.g. "Confidential")
   * @param {object} [opts] - Options
   * @param {string} [opts.position='header'] - 'header' or 'footer'
   * @param {string} [opts.alignment='center'] - 'left', 'center', or 'right'
   */
  static stamp(ws, text, opts = {}) {
    const position = (opts.position || 'header').toLowerCase();
    const alignment = (opts.alignment || 'center').toLowerCase();

    // Map alignment to OOXML jc values
    const jcMap = { left: 'left', center: 'center', right: 'right' };
    const jcVal = jcMap[alignment] || 'center';

    const isHeader = position === 'header';
    const partTag = isHeader ? 'w:hdr' : 'w:ftr';
    const refTag = isHeader ? 'w:headerReference' : 'w:footerReference';
    const fileName = isHeader ? 'headerStamp1.xml' : 'footerStamp1.xml';

    // Build the header/footer XML
    const partXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + `<${partTag} xmlns:w="${xml.NS.w}" xmlns:r="${xml.NS.r}">`
      + '<w:p>'
      + '<w:pPr>'
      + `<w:pStyle w:val="${isHeader ? 'Header' : 'Footer'}"/>`
      + `<w:jc w:val="${jcVal}"/>`
      + '</w:pPr>'
      + '<w:r>'
      + '<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
      + `<w:t>${xml.escapeXml(text)}</w:t>`
      + '</w:r>'
      + '</w:p>'
      + `</${partTag}>`;

    // Write the file
    ws._writeFile('word/' + fileName, partXml);

    // Add relationship
    const relId = _nextRelId(ws.relsXml);
    const relType = isHeader
      ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
    const newRel = `<Relationship Id="${relId}" Type="${relType}" Target="${fileName}"/>`;
    ws.relsXml = ws.relsXml.replace('</Relationships>', newRel + '</Relationships>');

    // Add content type override
    if (!ws.contentTypesXml.includes(fileName)) {
      const contentType = isHeader
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
      const override = `<Override PartName="/word/${fileName}" ContentType="${contentType}"/>`;
      ws.contentTypesXml = ws.contentTypesXml.replace('</Types>', override + '</Types>');
    }

    // Add reference to the main section properties
    const refXml = `<${refTag} w:type="default" r:id="${relId}"/>`;
    const docXml = ws.docXml;

    const sectPrMatch = docXml.match(/<w:sectPr[\s>][^]*?<\/w:sectPr>/);
    if (sectPrMatch) {
      const sectPr = sectPrMatch[0];
      // Check for existing default reference of same type
      const existingRefRe = new RegExp(`<${refTag}\\s+w:type="default"[^/]*/>`);
      if (existingRefRe.test(sectPr)) {
        ws.docXml = docXml.replace(existingRefRe, refXml);
      } else {
        const insertPoint = sectPr.indexOf('>') + 1;
        const newSectPr = sectPr.slice(0, insertPoint) + refXml + sectPr.slice(insertPoint);
        ws.docXml = docXml.replace(sectPr, newSectPr);
      }
    }
  }

  /**
   * Estimate page count without PDF rendering.
   * Formula: (wordCount / 250) + (figureCount * 0.5) + (tableCount * 0.3)
   * Adjusted for line spacing (double = 2x, single = 1x).
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {{ estimated: number, confidence: string, wordCount: number, figureCount: number, tableCount: number }}
   */
  static pageCount(ws) {
    const words = _wordCount(ws.docXml);
    const figures = _figureCount(ws.docXml);
    const tables = _tableCount(ws.docXml);
    const spacingMultiplier = _detectSpacing(ws);

    // Base estimate: 250 words per page (single spaced)
    const wordsPerPage = 250 / spacingMultiplier;
    const textPages = words / wordsPerPage;
    const figurePages = figures * 0.5;
    const tablePages = tables * 0.3;

    const estimated = Math.max(1, Math.round(textPages + figurePages + tablePages));

    return {
      estimated,
      confidence: 'rough',
      wordCount: words,
      figureCount: figures,
      tableCount: tables,
    };
  }

  /**
   * Insert a formatted cover page as the first page.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @param {object} opts - Cover page options
   * @param {string} [opts.title] - Document title
   * @param {string} [opts.subtitle] - Subtitle
   * @param {string} [opts.author] - Author name
   * @param {string} [opts.date] - Date string
   * @param {string} [opts.organization] - Organization name
   */
  static coverPage(ws, opts = {}) {
    const title = opts.title || 'Untitled Document';
    const subtitle = opts.subtitle || '';
    const author = opts.author || '';
    const date = opts.date || new Date().toISOString().split('T')[0];
    const organization = opts.organization || '';

    // Build cover page paragraphs
    const coverParagraphs = [];

    // Spacer paragraphs for top margin
    for (let i = 0; i < 6; i++) {
      coverParagraphs.push(
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>'
      );
    }

    // Title (large, bold, centered)
    coverParagraphs.push(
      '<w:p>'
      + '<w:pPr><w:jc w:val="center"/></w:pPr>'
      + '<w:r>'
      + '<w:rPr><w:b/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>'
      + `<w:t>${xml.escapeXml(title)}</w:t>`
      + '</w:r>'
      + '</w:p>'
    );

    // Subtitle (medium, centered)
    if (subtitle) {
      coverParagraphs.push(
        '<w:p>'
        + '<w:pPr><w:jc w:val="center"/></w:pPr>'
        + '<w:r>'
        + '<w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
        + `<w:t>${xml.escapeXml(subtitle)}</w:t>`
        + '</w:r>'
        + '</w:p>'
      );
    }

    // Spacer
    coverParagraphs.push('<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>');
    coverParagraphs.push('<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>');

    // Author
    if (author) {
      coverParagraphs.push(
        '<w:p>'
        + '<w:pPr><w:jc w:val="center"/></w:pPr>'
        + '<w:r>'
        + '<w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        + `<w:t>${xml.escapeXml(author)}</w:t>`
        + '</w:r>'
        + '</w:p>'
      );
    }

    // Organization
    if (organization) {
      coverParagraphs.push(
        '<w:p>'
        + '<w:pPr><w:jc w:val="center"/></w:pPr>'
        + '<w:r>'
        + '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/><w:i/></w:rPr>'
        + `<w:t>${xml.escapeXml(organization)}</w:t>`
        + '</w:r>'
        + '</w:p>'
      );
    }

    // Date
    if (date) {
      coverParagraphs.push(
        '<w:p>'
        + '<w:pPr><w:jc w:val="center"/></w:pPr>'
        + '<w:r>'
        + '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        + `<w:t>${xml.escapeXml(date)}</w:t>`
        + '</w:r>'
        + '</w:p>'
      );
    }

    // Page break after cover
    coverParagraphs.push(
      '<w:p>'
      + '<w:pPr><w:jc w:val="center"/></w:pPr>'
      + '<w:r><w:br w:type="page"/></w:r>'
      + '</w:p>'
    );

    const coverXml = coverParagraphs.join('');

    // Insert at the very beginning of the document body
    const docXml = ws.docXml;
    const bodyStart = docXml.indexOf('<w:body>');
    if (bodyStart === -1) {
      throw new Error('Could not find <w:body> in document');
    }
    const insertPos = bodyStart + '<w:body>'.length;
    ws.docXml = docXml.slice(0, insertPos) + coverXml + docXml.slice(insertPos);
  }
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = { Production };
