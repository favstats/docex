/**
 * figures.js -- Image and figure operations for docex
 *
 * Static methods for listing, inserting, and replacing images/figures
 * in OOXML documents. Ported from docx-patch.js and adapted for the
 * Workspace abstraction.
 *
 * Zero external dependencies. All XML is string-based.
 */

const fs = require('fs');
const path = require('path');
const xml = require('./xml');

// Maximum image width in EMU (6.5 inches for standard margins)
const MAX_WIDTH_EMU = 5943600;
const INCHES_TO_EMU = 914400;

// OOXML namespace URIs used in drawing XML
const NS = {
  w:   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp:  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a:   'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  mc:  'http://schemas.openxmlformats.org/markup-compatibility/2006',
};

class Figures {
  /**
   * List all images in the document.
   * Returns array of { rId, filename, width, height, caption, nearestHeading }
   *
   * @param {Workspace} ws - The open document workspace
   * @returns {Array<{rId: string, filename: string, width: number|null, height: number|null, caption: string, nearestHeading: string}>}
   */
  static list(ws) {
    const docXml = ws.docXml;
    const relsXml = ws.relsXml;

    // Build rId -> media filename mapping from rels
    const rIdToMedia = {};
    const relRe = /<Relationship\s+([^>]+)\/?\s*>/g;
    let rm;
    while ((rm = relRe.exec(relsXml)) !== null) {
      const attrs = rm[1];
      const id = (attrs.match(/Id="([^"]+)"/) || [])[1];
      const type = (attrs.match(/Type="([^"]+)"/) || [])[1];
      const target = (attrs.match(/Target="([^"]+)"/) || [])[1];
      if (id && type && type.endsWith('/image')) {
        rIdToMedia[id] = target.replace('media/', '');
      }
    }

    // Parse paragraphs
    const paragraphs = xml.findParagraphs(docXml);
    const results = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const pObj = paragraphs[i];
      const p = typeof pObj === 'string' ? pObj : pObj.xml;
      if (!p.includes('<w:drawing>') && !p.includes('w:drawing>')) continue;

      // Extract the r:embed rId
      const blipMatch = p.match(/a:blip[^>]+r:embed="(rId\d+)"/);
      if (!blipMatch) continue;
      const rId = blipMatch[1];

      const filename = rIdToMedia[rId] || null;

      // Extract dimensions from wp:extent
      const extentMatch = p.match(/<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/);
      const width = extentMatch ? parseInt(extentMatch[1], 10) : null;
      const height = extentMatch ? parseInt(extentMatch[2], 10) : null;

      // Look for a caption in the next paragraph
      let caption = '';
      if (i + 1 < paragraphs.length) {
        const nextPObj = paragraphs[i + 1];
        const nextText = typeof nextPObj === 'string' ? xml.extractText(nextPObj) : nextPObj.text;
        if (nextText.startsWith('Figure ') || nextText.startsWith('Fig.')) {
          caption = nextText;
        }
      }

      // Find nearest heading above
      let nearestHeading = '';
      for (let j = i - 1; j >= 0; j--) {
        const jPObj = paragraphs[j];
        const jXml = typeof jPObj === 'string' ? jPObj : jPObj.xml;
        const styleMatch = jXml.match(/<w:pStyle\s+w:val="([^"]+)"/);
        if (styleMatch) {
          const sid = styleMatch[1];
          if (/^(?:Heading\d|heading\d|\d{2,3})$/.test(sid)) {
            nearestHeading = typeof jPObj === 'string' ? xml.extractText(jPObj) : jPObj.text;
            break;
          }
        }
      }

      results.push({ rId, filename, width, height, caption, nearestHeading });
    }

    return results;
  }

  /**
   * Insert a figure at a position in the document.
   *
   * Steps:
   *   1. Copy image to word/media/imageN.ext
   *   2. Add relationship to document.xml.rels
   *   3. Add content type if needed
   *   4. Build drawing XML with proper dimensions
   *   5. Build caption paragraph
   *   6. Find anchor paragraph, insert figure + caption after/before it
   *   7. If tracked, wrap in w:ins
   *
   * @param {Workspace} ws - The open document workspace
   * @param {string} anchor - text to position relative to
   * @param {string} mode - 'after' or 'before'
   * @param {string} imagePath - path to PNG/JPEG file on disk
   * @param {string} caption - figure caption text (e.g. "Figure 1. Results")
   * @param {object} opts - { width (inches), tracked, author, date }
   */
  static insert(ws, anchor, mode, imagePath, caption, opts = {}) {
    if (!fs.existsSync(imagePath)) {
      throw new Error(`Image file not found: ${imagePath}`);
    }

    // 1. Copy image to media and add relationship
    const { mediaFilename, rId } = Figures._copyToMedia(ws, imagePath);

    // 2. Add content type for the extension if needed
    const ext = path.extname(imagePath).toLowerCase().replace('.', '');
    Figures._addContentType(ws, ext);

    // 3. Compute dimensions
    const widthInches = opts.width || 6;
    const { cx, cy } = Figures._computeEmu(imagePath, widthInches);

    // 4. Build figure name
    const name = path.basename(imagePath, path.extname(imagePath));

    // 5. Build drawing paragraph XML
    const drawingXml = Figures._buildDrawingXml(rId, cx, cy, name, name);
    const figureParagraph = '<w:p>'
      + '<w:pPr><w:pBdr/><w:spacing w:line="240" w:lineRule="auto"/>'
      + '<w:ind/><w:jc w:val="center"/><w:rPr/></w:pPr>'
      + '<w:r><w:drawing>' + drawingXml + '</w:drawing></w:r>'
      + '</w:p>';

    // 6. Build caption paragraph
    let captionParagraph = '';
    if (caption) {
      captionParagraph = Figures._buildCaptionXml(caption);
    }

    // 7. Find anchor and insert
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    const idx = Figures._findAnchorIndex(paragraphs, anchor);
    if (idx === -1) {
      throw new Error(`Anchor not found: "${anchor}"`);
    }

    // Build the elements to insert
    let newElements = figureParagraph;
    if (captionParagraph) {
      newElements += captionParagraph;
    }

    // Note: paragraph-level tracked changes (w:ins wrapping whole paragraphs)
    // are not supported via buildIns (which is for run-level text).
    // Figure insertions are always direct (untracked at paragraph level).

    // Insert at the right position
    const anchorParagraph = paragraphs[idx];
    let insertPos;
    if (mode === 'before') {
      insertPos = anchorParagraph.start;
    } else {
      // 'after' (default)
      insertPos = anchorParagraph.end;
    }

    ws.docXml = docXml.slice(0, insertPos) + newElements + docXml.slice(insertPos);
  }

  /**
   * Replace an existing figure by name prefix.
   *
   * Finds the figure whose media file starts with the given prefix,
   * overwrites the media file, and updates dimensions if needed.
   *
   * @param {Workspace} ws - The open document workspace
   * @param {string} namePrefix - e.g. "fig03" matches any media file starting with fig03
   * @param {string} newImagePath - path to replacement image on disk
   */
  static replace(ws, namePrefix, newImagePath) {
    if (!fs.existsSync(newImagePath)) {
      throw new Error(`Image file not found: ${newImagePath}`);
    }

    const docXml = ws.docXml;
    const relsXml = ws.relsXml;

    // Build rId -> media filename mapping
    const rIdToMedia = {};
    const relRe = /<Relationship\s+([^>]+)\/?\s*>/g;
    let rm;
    while ((rm = relRe.exec(relsXml)) !== null) {
      const attrs = rm[1];
      const id = (attrs.match(/Id="([^"]+)"/) || [])[1];
      const type = (attrs.match(/Type="([^"]+)"/) || [])[1];
      const target = (attrs.match(/Target="([^"]+)"/) || [])[1];
      if (id && type && type.endsWith('/image')) {
        rIdToMedia[id] = target.replace('media/', '');
      }
    }

    // Also build a caption-based image map for name matching
    const paragraphs = xml.findParagraphs(docXml);
    const imageMap = Figures._buildImageMap(paragraphs, rIdToMedia);

    // Try to find a matching figure
    let matchedRId = null;
    let matchedMedia = null;
    const lowerPrefix = namePrefix.toLowerCase();

    // 1. Match by image map key (caption-derived name)
    for (const [key, info] of Object.entries(imageMap)) {
      if (key.toLowerCase() === lowerPrefix ||
          key.toLowerCase().startsWith(lowerPrefix) ||
          key.toLowerCase().includes(lowerPrefix)) {
        matchedRId = info.rId;
        matchedMedia = info.media;
        break;
      }
    }

    // 2. Match by rId directly (e.g., "rId9")
    if (!matchedRId && namePrefix.startsWith('rId')) {
      if (rIdToMedia[namePrefix]) {
        matchedRId = namePrefix;
        matchedMedia = rIdToMedia[namePrefix];
      }
    }

    // 3. Match by media filename
    if (!matchedRId) {
      for (const [rId, media] of Object.entries(rIdToMedia)) {
        if (media.toLowerCase().includes(lowerPrefix)) {
          matchedRId = rId;
          matchedMedia = media;
          break;
        }
      }
    }

    if (!matchedRId || !matchedMedia) {
      const available = Object.entries(imageMap)
        .map(([k, v]) => `${k} -> ${v.rId} (${v.media})`)
        .join(', ');
      throw new Error(`No image found matching "${namePrefix}". Available: ${available}`);
    }

    // Overwrite the media file in word/media/
    const destPath = path.join(ws.mediaDir, matchedMedia);
    fs.copyFileSync(newImagePath, destPath);

    // Update dimensions in the document XML
    const { cx: newCx, cy: newCy } = Figures._computeEmu(newImagePath);

    let updatedXml = ws.docXml;

    // Find the drawing element that references this rId and update its dimensions
    // We scan through paragraphs to find the one containing this rId
    const escapedRId = matchedRId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const drawingRe = new RegExp(
      '(<w:p[^>]*>(?:(?!</w:p>)[\\s\\S])*?a:blip[^>]+r:embed="'
        + escapedRId
        + '"(?:(?!</w:p>)[\\s\\S])*?</w:p>)',
      'g'
    );

    updatedXml = updatedXml.replace(drawingRe, (matchedParagraph) => {
      let el = matchedParagraph;

      // Update wp:extent
      el = el.replace(
        /<wp:extent\s+cx="\d+"\s+cy="\d+"/g,
        `<wp:extent cx="${newCx}" cy="${newCy}"`
      );

      // Update a:ext in pic:spPr > a:xfrm > a:ext
      el = el.replace(
        /(<a:xfrm>[\s\S]*?<a:ext\s+)cx="\d+"\s+cy="\d+"/,
        `$1cx="${newCx}" cy="${newCy}"`
      );

      // Update VML fallback dimensions (points) if present
      const widthPt = (newCx / INCHES_TO_EMU * 72).toFixed(2);
      const heightPt = (newCy / INCHES_TO_EMU * 72).toFixed(2);
      el = el.replace(
        /width:\d+\.?\d*pt;height:\d+\.?\d*pt/,
        `width:${widthPt}pt;height:${heightPt}pt`
      );

      return el;
    });

    ws.docXml = updatedXml;
  }

  // ==========================================================================
  // INTERNAL HELPERS
  // ==========================================================================

  /**
   * Build the wp:inline drawing XML for an embedded image.
   *
   * @param {string} rId - Relationship ID (e.g. "rId7")
   * @param {number} cx - Width in EMU
   * @param {number} cy - Height in EMU
   * @param {string} name - Image name for docPr
   * @param {string} descr - Description for accessibility
   * @returns {string} Complete wp:inline XML (without the outer w:drawing wrapper)
   */
  static _buildDrawingXml(rId, cx, cy, name, descr) {
    const docPrId = Date.now() % 1000000000;
    const cNvPrId = docPrId + 1;
    const safeName = xml.escapeXml(name || '');
    const safeDescr = xml.escapeXml(descr || '');

    return '<wp:inline distT="0" distB="0" distL="0" distR="0">'
      + '<wp:extent cx="' + cx + '" cy="' + cy + '"/>'
      + '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
      + '<wp:docPr id="' + docPrId + '" name="' + safeName + '" descr="' + safeDescr + '"/>'
      + '<wp:cNvGraphicFramePr>'
      + '<a:graphicFrameLocks xmlns:a="' + NS.a + '" noChangeAspect="1"/>'
      + '</wp:cNvGraphicFramePr>'
      + '<a:graphic xmlns:a="' + NS.a + '">'
      + '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      + '<pic:pic xmlns:pic="' + NS.pic + '">'
      + '<pic:nvPicPr>'
      + '<pic:cNvPr id="' + cNvPrId + '" name="' + safeName + '"/>'
      + '<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>'
      + '<pic:nvPr/>'
      + '</pic:nvPicPr>'
      + '<pic:blipFill rotWithShape="1">'
      + '<a:blip r:embed="' + rId + '" xmlns:r="' + NS.r + '"/>'
      + '<a:stretch/>'
      + '</pic:blipFill>'
      + '<pic:spPr bwMode="auto">'
      + '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
      + '</pic:spPr>'
      + '</pic:pic>'
      + '</a:graphicData>'
      + '</a:graphic>'
      + '</wp:inline>';
  }

  /**
   * Build a caption paragraph XML string.
   *
   * If the caption matches "Figure N. description", the "Figure N." part
   * is rendered bold and the description is rendered italic, matching
   * standard academic formatting.
   *
   * @param {string} captionText - Caption text (e.g. "Figure 1. Survey results")
   * @returns {string} Complete w:p XML for the caption
   */
  static _buildCaptionXml(captionText) {
    const capMatch = captionText.match(/^(Figure\s+\d+\.\s*)(.*)/);

    if (capMatch) {
      const boldPart = capMatch[1];
      const italicPart = capMatch[2];
      return '<w:p>'
        + '<w:pPr><w:pBdr/><w:spacing w:after="120" w:line="240" w:lineRule="auto"/>'
        + '<w:ind/><w:jc w:val="center"/><w:rPr/></w:pPr>'
        + '<w:r><w:rPr>'
        + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
        + '<w:b/><w:sz w:val="24"/>'
        + '</w:rPr><w:t xml:space="preserve">' + xml.escapeXml(boldPart) + '</w:t></w:r>'
        + '<w:r><w:rPr>'
        + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
        + '<w:i/><w:sz w:val="24"/>'
        + '</w:rPr><w:t xml:space="preserve">' + xml.escapeXml(italicPart) + '</w:t></w:r>'
        + '</w:p>';
    }

    // Plain caption (italic)
    return '<w:p>'
      + '<w:pPr><w:pBdr/><w:spacing w:after="120" w:line="240" w:lineRule="auto"/>'
      + '<w:ind/><w:jc w:val="center"/><w:rPr/></w:pPr>'
      + '<w:r><w:rPr>'
      + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
      + '<w:i/><w:sz w:val="24"/>'
      + '</w:rPr><w:t xml:space="preserve">' + xml.escapeXml(captionText) + '</w:t></w:r>'
      + '</w:p>';
  }

  /**
   * Copy an image file into the workspace's word/media/ directory.
   * Assigns the next available imageN filename and adds the relationship.
   *
   * @param {Workspace} ws - The open document workspace
   * @param {string} imagePath - Absolute path to the source image
   * @returns {{ mediaFilename: string, rId: string }}
   */
  static _copyToMedia(ws, imagePath) {
    const ext = path.extname(imagePath).toLowerCase();
    const mediaDir = ws.mediaDir;

    // Ensure media directory exists
    if (!fs.existsSync(mediaDir)) {
      fs.mkdirSync(mediaDir, { recursive: true });
    }

    // Find the next available image number
    const files = fs.existsSync(mediaDir) ? fs.readdirSync(mediaDir) : [];
    let maxNum = 0;
    for (const f of files) {
      const m = f.match(/^image(\d+)\./);
      if (m) {
        const num = parseInt(m[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
    const nextNum = maxNum + 1;
    const mediaFilename = 'image' + nextNum + ext;

    // Copy the file
    fs.copyFileSync(imagePath, path.join(mediaDir, mediaFilename));

    // Add the relationship and get the rId
    const rId = Figures._addRelationship(ws, mediaFilename);

    return { mediaFilename, rId };
  }

  /**
   * Add an image relationship to document.xml.rels.
   *
   * @param {Workspace} ws - The open document workspace
   * @param {string} mediaFilename - Filename in word/media/ (e.g. "image5.png")
   * @returns {string} The assigned relationship ID (e.g. "rId12")
   */
  static _addRelationship(ws, mediaFilename) {
    const relsXml = ws.relsXml;
    const rId = xml.nextRId(relsXml);

    const relEntry = '<Relationship Id="' + rId + '" '
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
      + 'Target="media/' + mediaFilename + '"/>';

    ws.relsXml = relsXml.replace('</Relationships>', relEntry + '</Relationships>');
    return rId;
  }

  /**
   * Add a content type entry for an image extension if not already present.
   *
   * @param {Workspace} ws - The open document workspace
   * @param {string} extension - File extension without dot (e.g. "png", "jpeg")
   */
  static _addContentType(ws, extension) {
    let ctXml = ws.contentTypesXml;
    if (ctXml.includes('Extension="' + extension + '"')) return;

    const mimeMap = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      tiff: 'image/tiff',
      tif: 'image/tiff',
    };
    const mime = mimeMap[extension] || ('image/' + extension);

    ctXml = ctXml.replace(
      '</Types>',
      '<Default Extension="' + extension + '" ContentType="' + mime + '"/></Types>'
    );
    ws.contentTypesXml = ctXml;
  }

  /**
   * Compute EMU dimensions for an image, scaling to fit within the page width.
   *
   * Reads PNG/JPEG headers to get the native pixel dimensions, then scales
   * so the width matches the requested inches (default: full text width)
   * while preserving aspect ratio.
   *
   * @param {string} filePath - Path to the image file
   * @param {number} [maxWidthInches=6.5] - Maximum width in inches
   * @returns {{ cx: number, cy: number }} Dimensions in EMU
   */
  static _computeEmu(filePath, maxWidthInches) {
    const buf = fs.readFileSync(filePath);
    let dims;

    // Detect format by magic bytes
    if (buf[0] === 0x89 && buf[1] === 0x50) {
      // PNG: width at offset 16, height at offset 20 (IHDR chunk)
      dims = {
        width: buf.readUInt32BE(16),
        height: buf.readUInt32BE(20),
      };
    } else {
      // Try JPEG
      dims = Figures._getJpegDimensions(buf);
    }

    const aspectRatio = dims.height / dims.width;
    const maxCx = maxWidthInches
      ? Math.round(maxWidthInches * INCHES_TO_EMU)
      : MAX_WIDTH_EMU;
    // Clamp to page width
    const cx = Math.min(maxCx, MAX_WIDTH_EMU);
    const cy = Math.round(cx * aspectRatio);

    return { cx, cy };
  }

  /**
   * Parse JPEG dimensions from a file buffer by scanning SOF markers.
   *
   * @param {Buffer} buf - File buffer
   * @returns {{ width: number, height: number }}
   */
  static _getJpegDimensions(buf) {
    let offset = 2;
    while (offset < buf.length) {
      if (buf[offset] !== 0xFF) break;
      const marker = buf[offset + 1];
      // SOF0 (0xC0) or SOF2 (0xC2) contain dimensions
      if (marker === 0xC0 || marker === 0xC2) {
        const height = buf.readUInt16BE(offset + 5);
        const width = buf.readUInt16BE(offset + 7);
        return { width, height };
      }
      const len = buf.readUInt16BE(offset + 2);
      offset += 2 + len;
    }
    // Fallback if no SOF marker found
    return { width: 800, height: 600 };
  }

  /**
   * Find the index of the paragraph whose text contains the anchor string.
   * Tries exact match first, then substring, then case-insensitive substring.
   *
   * @param {string[]} paragraphs - Array of w:p XML strings
   * @param {string} anchor - Text to search for
   * @returns {number} Index of the matching paragraph, or -1
   */
  static _findAnchorIndex(paragraphs, anchor) {
    // Exact match
    for (let i = 0; i < paragraphs.length; i++) {
      const text = typeof paragraphs[i] === 'string' ? xml.extractText(paragraphs[i]) : paragraphs[i].text;
      if (text === anchor) return i;
    }
    // Substring match
    for (let i = 0; i < paragraphs.length; i++) {
      const text = typeof paragraphs[i] === 'string' ? xml.extractText(paragraphs[i]) : paragraphs[i].text;
      if (text.includes(anchor)) return i;
    }
    // Case-insensitive substring
    const lower = anchor.toLowerCase();
    for (let i = 0; i < paragraphs.length; i++) {
      const text = typeof paragraphs[i] === 'string' ? xml.extractText(paragraphs[i]) : paragraphs[i].text;
      if (text.toLowerCase().includes(lower)) return i;
    }
    return -1;
  }

  /**
   * Build a caption-based image map from paragraphs and rels.
   * Maps figure names (e.g. "fig01_survey_results") to { rId, media, caption }.
   *
   * @param {string[]} paragraphs - Array of w:p XML strings
   * @param {Object<string, string>} rIdToMedia - Map of rId to media filename
   * @returns {Object<string, {rId: string, media: string, caption: string}>}
   */
  static _buildImageMap(paragraphs, rIdToMedia) {
    const imageMap = {};

    for (let i = 0; i < paragraphs.length; i++) {
      const pObj = paragraphs[i];
      const p = typeof pObj === 'string' ? pObj : pObj.xml;
      if (!p.includes('<w:drawing>') && !p.includes('w:drawing>')) continue;

      const blipMatch = p.match(/a:blip[^>]+r:embed="(rId\d+)"/);
      if (!blipMatch) continue;
      const rId = blipMatch[1];

      const mediaFile = rIdToMedia[rId];
      if (!mediaFile) continue;

      // Check next paragraph for a caption
      let caption = '';
      let figName = '';
      if (i + 1 < paragraphs.length) {
        const nextPObj = paragraphs[i + 1];
        const nextText = typeof nextPObj === 'string' ? xml.extractText(nextPObj) : nextPObj.text;
        if (nextText.startsWith('Figure ')) {
          caption = nextText;
          const numMatch = nextText.match(/^Figure\s+(\d+)\.\s*(.*)/);
          if (numMatch) {
            const figNum = numMatch[1].padStart(2, '0');
            const slug = numMatch[2]
              .toLowerCase()
              .replace(/[^a-z0-9]+/g, '_')
              .replace(/_+$/, '')
              .split('_')
              .slice(0, 3)
              .join('_');
            figName = 'fig' + figNum + '_' + slug;
          }
        }
      }

      // Fall back to media filename as the key
      if (!figName) {
        figName = mediaFile.replace(/\.[^.]+$/, '');
      }

      imageMap[figName] = { rId, media: mediaFile, caption };
    }

    return imageMap;
  }
}

module.exports = { Figures };
