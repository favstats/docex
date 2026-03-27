/**
 * tables.js -- Table operations for docex
 *
 * Static methods for inserting and building OOXML tables.
 * Supports "booktabs" (academic) and "plain" (grid) styles.
 *
 * Zero external dependencies. All XML is string-based.
 *
 * Booktabs style mirrors LaTeX booktabs:
 *   - Thick top rule (sz=12)
 *   - Thin rule under the header row (sz=6)
 *   - Thick bottom rule (sz=12)
 *   - No vertical borders anywhere
 *
 * Column widths auto-fill the text area (9360 twips for standard margins).
 */

const xml = require('./xml');

// Total page text width in twips (letter paper, 1-inch margins)
const PAGE_WIDTH_TWIPS = 9360;

class Tables {
  /**
   * Insert a table at a position in the document.
   *
   * @param {Workspace} ws - The open document workspace
   * @param {string} anchor - text to position relative to
   * @param {string} mode - 'after' or 'before'
   * @param {Array<Array<string>>} data - 2D array of cell values. First row is header if opts.headers is true.
   * @param {object} opts - { caption, headers (bool), style, tracked, author, date }
   * @param {string} [opts.caption] - Optional caption paragraph above the table (e.g. "Table 1. Results")
   * @param {boolean} [opts.headers=true] - Whether the first row is a header row (bold, with bottom rule)
   * @param {string} [opts.style='booktabs'] - Table style: 'booktabs' or 'plain'
   * @param {boolean} [opts.tracked] - Whether to wrap in tracked changes
   * @param {string} [opts.author] - Author name for tracked changes
   * @param {string} [opts.date] - ISO date for tracked changes
   */
  static insert(ws, anchor, mode, data, opts = {}) {
    if (!data || data.length === 0) {
      throw new Error('Table data must be a non-empty 2D array');
    }

    const style = opts.style || 'booktabs';
    const headers = opts.headers !== false;

    // 1. Build the table XML
    const tableXml = Tables._buildTableXml(data, { headers, style });

    // 2. Build caption paragraph if provided
    let captionXml = '';
    if (opts.caption) {
      captionXml = Tables._buildCaptionXml(opts.caption);
    }

    // 3. Find the anchor paragraph
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    const idx = xml.findAnchorParagraph(paragraphs, anchor);
    if (idx === -1) {
      throw new Error('Anchor not found: "' + anchor + '"');
    }

    // 4. Build the combined elements (caption + table)
    let newElements = '';
    if (captionXml) {
      newElements += captionXml;
    }
    newElements += tableXml;

    // Note: paragraph-level tracked changes are not supported here.
    // Table insertions are always direct (untracked at paragraph level).

    // 6. Insert at the right position
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
   * Build a complete w:tbl XML string from a 2D data array.
   *
   * @param {Array<Array<string>>} data - 2D array of cell values
   * @param {object} opts - { headers (bool), style ('booktabs'|'plain') }
   * @returns {string} Complete <w:tbl>...</w:tbl> XML
   */
  static _buildTableXml(data, opts = {}) {
    const headers = opts.headers !== false;
    const style = opts.style || 'booktabs';
    const numCols = Math.max(...data.map(row => row.length));
    const colWidth = Math.floor(PAGE_WIDTH_TWIPS / numCols);

    // Table properties
    let tblPr = '<w:tblPr>';
    tblPr += '<w:tblW w:w="' + PAGE_WIDTH_TWIPS + '" w:type="dxa"/>';
    tblPr += '<w:tblLayout w:type="fixed"/>';
    tblPr += '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
      + 'w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';

    if (style === 'booktabs') {
      // Booktabs: no default borders on the table -- borders are set per cell
      tblPr += '<w:tblBorders>'
        + '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        + '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        + '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        + '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        + '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        + '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        + '</w:tblBorders>';
    } else {
      // Plain: all borders (standard grid)
      tblPr += '<w:tblBorders>'
        + '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        + '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        + '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        + '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        + '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        + '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        + '</w:tblBorders>';
    }

    tblPr += '</w:tblPr>';

    // Column grid
    let tblGrid = '<w:tblGrid>';
    for (let c = 0; c < numCols; c++) {
      tblGrid += '<w:gridCol w:w="' + colWidth + '"/>';
    }
    tblGrid += '</w:tblGrid>';

    // Build rows
    let rowsXml = '';
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      const isFirstRow = (r === 0);
      const isLastRow = (r === data.length - 1);
      const isHeaderRow = (isFirstRow && headers);

      rowsXml += Tables._buildRowXml(
        row, numCols, colWidth, style, { isFirstRow, isLastRow, isHeaderRow }
      );
    }

    return '<w:tbl>' + tblPr + tblGrid + rowsXml + '</w:tbl>';
  }

  /**
   * Build a single table row (w:tr) XML string.
   *
   * @param {Array<string>} cells - Cell values for this row
   * @param {number} numCols - Total number of columns in the table
   * @param {number} colWidth - Column width in twips
   * @param {string} style - 'booktabs' or 'plain'
   * @param {object} flags - { isFirstRow, isLastRow, isHeaderRow }
   * @returns {string} Complete <w:tr>...</w:tr> XML
   */
  static _buildRowXml(cells, numCols, colWidth, style, flags) {
    const { isHeaderRow } = flags;

    // Row properties
    let trPr = '';
    if (isHeaderRow) {
      trPr = '<w:trPr><w:tblHeader/></w:trPr>';
    }

    // Build cells
    let cellsXml = '';
    for (let c = 0; c < numCols; c++) {
      const cellText = (c < cells.length) ? String(cells[c]) : '';
      cellsXml += Tables._buildCellXml(cellText, colWidth, style, flags);
    }

    return '<w:tr>' + trPr + cellsXml + '</w:tr>';
  }

  /**
   * Build a single table cell (w:tc) XML string.
   *
   * In booktabs mode, cell borders are set individually:
   *   - First row cells: thick top border (sz=12)
   *   - Header row cells: thin bottom border (sz=6)
   *   - Last row cells: thick bottom border (sz=12)
   *   - No vertical borders ever
   *
   * @param {string} text - Cell text content
   * @param {number} colWidth - Column width in twips
   * @param {string} style - 'booktabs' or 'plain'
   * @param {object} flags - { isFirstRow, isLastRow, isHeaderRow }
   * @returns {string} Complete <w:tc>...</w:tc> XML
   */
  static _buildCellXml(text, colWidth, style, flags) {
    const { isFirstRow, isLastRow, isHeaderRow } = flags;

    let tcPr = '<w:tcPr>';
    tcPr += '<w:tcW w:w="' + colWidth + '" w:type="dxa"/>';

    if (style === 'booktabs') {
      // Build per-cell borders for booktabs
      let borders = '<w:tcBorders>';

      // Top border: thick on the first row
      if (isFirstRow) {
        borders += '<w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>';
      } else {
        borders += '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>';
      }

      // Bottom border: thin under header, thick on last row
      if (isHeaderRow) {
        borders += '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>';
      } else if (isLastRow) {
        borders += '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>';
      } else {
        borders += '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>';
      }

      // No vertical borders
      borders += '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>';
      borders += '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>';

      borders += '</w:tcBorders>';
      tcPr += borders;
    }
    // For 'plain' style, cell borders are inherited from the table-level tblBorders

    tcPr += '</w:tcPr>';

    // Run properties: Times New Roman 11pt, bold for header rows
    let rPr = '<w:rPr>'
      + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
      + 'w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
      + '<w:sz w:val="22"/><w:szCs w:val="22"/>';
    if (isHeaderRow) {
      rPr += '<w:b/><w:bCs/>';
    }
    rPr += '</w:rPr>';

    // Cell paragraph
    const pXml = '<w:p>'
      + '<w:pPr>'
      + '<w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/>'
      + '<w:rPr/>'
      + '</w:pPr>'
      + '<w:r>' + rPr
      + '<w:t xml:space="preserve">' + xml.escapeXml(String(text || '')) + '</w:t>'
      + '</w:r>'
      + '</w:p>';

    return '<w:tc>' + tcPr + pXml + '</w:tc>';
  }

  /**
   * Build a table caption paragraph XML string.
   *
   * If the caption matches "Table N. description", the "Table N." part
   * is rendered bold and the description is rendered italic, matching
   * standard academic formatting.
   *
   * @param {string} captionText - Caption text (e.g. "Table 1. Summary statistics")
   * @returns {string} Complete w:p XML for the caption
   */
  static _buildCaptionXml(captionText) {
    const capMatch = captionText.match(/^(Table\s+\d+\.\s*)(.*)/);

    if (capMatch) {
      const boldPart = capMatch[1];
      const italicPart = capMatch[2];
      return '<w:p>'
        + '<w:pPr><w:pBdr/><w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="auto"/>'
        + '<w:ind/><w:rPr/></w:pPr>'
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
      + '<w:pPr><w:pBdr/><w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="auto"/>'
      + '<w:ind/><w:rPr/></w:pPr>'
      + '<w:r><w:rPr>'
      + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
      + '<w:i/><w:sz w:val="24"/>'
      + '</w:rPr><w:t xml:space="preserve">' + xml.escapeXml(captionText) + '</w:t></w:r>'
      + '</w:p>';
  }

}

module.exports = { Tables };
