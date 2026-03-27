/**
 * latex.js -- OOXML to LaTeX conversion.
 *
 * Ported from the production docx-to-latex.js converter.
 * Works with the docex Workspace object for zero-copy integration.
 *
 * Usage:
 *   const { Latex } = require('./latex');
 *   const ws = Workspace.open('manuscript.docx');
 *   const tex = Latex.convert(ws, { documentClass: 'article' });
 *
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ============================================================================
// INTERNAL HELPERS
// ============================================================================

/**
 * Decode XML entities back to literal characters.
 * @param {string} str
 * @returns {string}
 */
function decodeEntities(str) {
  return str
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&');
}

/**
 * Escape a string for safe inclusion in LaTeX.
 * Handles XML entity decoding first, then LaTeX special chars,
 * then smart quotes and dashes.
 * @param {string} text
 * @returns {string}
 */
function escapeForLatex(text) {
  text = decodeEntities(text);
  return text
    .replace(/\\/g, '\\textbackslash{}')
    .replace(/&/g, '\\&')
    .replace(/%/g, '\\%')
    .replace(/\$/g, '\\$')
    .replace(/#/g, '\\#')
    .replace(/_/g, '\\_')
    .replace(/~/g, '\\textasciitilde{}')
    .replace(/\u201C/g, '``')
    .replace(/\u201D/g, "''")
    .replace(/\u2018/g, '`')
    .replace(/\u2019/g, "'")
    .replace(/\u2013/g, '--')
    .replace(/\u2014/g, '---')
    .replace(/\u2026/g, '\\ldots{}')
    .replace(/\u00A0/g, '~');
}

/**
 * Escape text for use inside a table cell.
 * @param {string} text
 * @returns {string}
 */
function escapeTableCell(text) {
  return text
    .replace(/\\/g, '\\textbackslash{}')
    .replace(/&/g, '\\&')
    .replace(/%/g, '\\%')
    .replace(/\$/g, '\\$')
    .replace(/#/g, '\\#')
    .replace(/_/g, '\\_')
    .replace(/~/g, '\\textasciitilde{}')
    .replace(/\u2013/g, '--')
    .replace(/\u2014/g, '---');
}

// ============================================================================
// XML PARSING
// ============================================================================

/**
 * Extract text from w:t elements within an XML fragment.
 * @param {string} xmlStr
 * @returns {string}
 */
function extractTextFromXml(xmlStr) {
  var text = '';
  var tRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  var m;
  while ((m = tRegex.exec(xmlStr)) !== null) {
    text += m[1];
  }
  return text;
}

/**
 * Find the closing tag for a given tag name, respecting nesting.
 * @param {string} xmlStr - The XML to search
 * @param {number} startIdx - Index of the opening tag
 * @param {string} tagName - Tag name (e.g. 'w:p')
 * @returns {number} End index (after closing tag), or -1
 */
function findClosingTag(xmlStr, startIdx, tagName) {
  var openTag = '<' + tagName;
  var closeTag = '</' + tagName + '>';
  var depth = 0;

  var firstTagEnd = xmlStr.indexOf('>', startIdx);
  if (firstTagEnd === -1) return -1;
  var searchIdx = firstTagEnd + 1;

  while (searchIdx < xmlStr.length) {
    var nextOpen = xmlStr.indexOf(openTag, searchIdx);
    var nextClose = xmlStr.indexOf(closeTag, searchIdx);

    if (nextClose === -1) return -1;

    if (nextOpen !== -1 && nextOpen < nextClose) {
      var charAfter = xmlStr[nextOpen + openTag.length];
      if (charAfter === '>' || charAfter === ' ' || charAfter === '/') {
        var tagEnd = xmlStr.indexOf('>', nextOpen);
        if (tagEnd !== -1 && xmlStr[tagEnd - 1] === '/') {
          searchIdx = tagEnd + 1;
          continue;
        }
        depth++;
        searchIdx = tagEnd + 1;
      } else {
        searchIdx = nextOpen + openTag.length;
      }
    } else {
      if (depth === 0) {
        return nextClose + closeTag.length;
      }
      depth--;
      searchIdx = nextClose + closeTag.length;
    }
  }
  return -1;
}

// ============================================================================
// STYLE MAP
// ============================================================================

/**
 * Build a mapping from style IDs to heading levels from styles.xml.
 * @param {string} stylesXml
 * @returns {Object<string, number>}
 */
function buildStyleMap(stylesXml) {
  var map = {};

  var styleRegex = /<w:style\b[^>]*>[\s\S]*?<\/w:style>/g;
  var match;

  while ((match = styleRegex.exec(stylesXml)) !== null) {
    var block = match[0];
    if (block.indexOf('w:type="paragraph"') === -1) continue;
    var idMatch = block.match(/w:styleId="([^"]*)"/);
    if (!idMatch) continue;
    var styleId = idMatch[1];

    var lvlMatch = block.match(/w:outlineLvl[^>]*w:val="(\d+)"/);
    if (lvlMatch) {
      map[styleId] = parseInt(lvlMatch[1]) + 1;
    }

    var nameMatch = block.match(/w:name[^>]*w:val="([^"]*)"/);
    if (nameMatch) {
      var name = nameMatch[1].toLowerCase();
      var headingMatch = name.match(/heading\s*(\d+)/);
      if (headingMatch) {
        map[styleId] = parseInt(headingMatch[1]);
      }
    }
  }

  var directIds = {
    'Heading1': 1, 'Heading2': 2, 'Heading3': 3, 'Heading4': 4,
    'heading1': 1, 'heading2': 2, 'heading3': 3, 'heading4': 4,
    'Title': 0,
  };
  for (var k in directIds) {
    if (!(k in map)) map[k] = directIds[k];
  }

  return map;
}

// ============================================================================
// PARAGRAPH PARSING
// ============================================================================

/**
 * Check if an XML fragment contains an embedded image.
 * @param {string} xmlStr
 * @returns {boolean}
 */
function hasImage(xmlStr) {
  return xmlStr.indexOf('<a:blip') !== -1 || xmlStr.indexOf('<pic:') !== -1;
}

/**
 * Extract relationship IDs for embedded images.
 * @param {string} xmlStr
 * @returns {string[]}
 */
function extractImageRIds(xmlStr) {
  var rids = [];
  var ridRegex = /r:embed="([^"]*)"/g;
  var m;
  while ((m = ridRegex.exec(xmlStr)) !== null) {
    rids.push(m[1]);
  }
  return rids;
}

/**
 * Extract formatted text segments from a paragraph's runs.
 * Each segment has { text, bold, italic, superscript, subscript, footnote }.
 * @param {string} pXml - Raw paragraph XML
 * @param {Object} footnoteMap - Map of footnote ID to text
 * @returns {Array<Object>}
 */
function extractFormattedText(pXml, footnoteMap) {
  var segments = [];
  var runRegex = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
  var runMatch;

  while ((runMatch = runRegex.exec(pXml)) !== null) {
    var runContent = runMatch[1];

    // Skip field markers and instruction text
    if (runContent.indexOf('w:fldChar') !== -1) continue;
    if (runContent.indexOf('w:instrText') !== -1) continue;

    // Footnote reference
    var fnRef = runContent.match(/w:footnoteReference\s[^>]*w:id="(\d+)"/);
    if (fnRef && footnoteMap && footnoteMap[fnRef[1]]) {
      segments.push({ text: '', footnote: footnoteMap[fnRef[1]] });
      continue;
    }

    // Formatting detection
    var bold = runContent.indexOf('<w:b/>') !== -1 || runContent.indexOf('<w:b ') !== -1;
    var italic = runContent.indexOf('<w:i/>') !== -1 || runContent.indexOf('<w:i ') !== -1;
    var superscript = runContent.indexOf('w:vertAlign') !== -1 && runContent.indexOf('"superscript"') !== -1;
    var subscript = runContent.indexOf('w:vertAlign') !== -1 && runContent.indexOf('"subscript"') !== -1;

    // Handle w:b with val="true" or val="1"
    if (!bold) {
      var bValMatch = runContent.match(/<w:b\s+w:val="([^"]*)"/);
      if (bValMatch && bValMatch[1] !== 'false' && bValMatch[1] !== '0') bold = true;
    }
    if (!italic) {
      var iValMatch = runContent.match(/<w:i\s+w:val="([^"]*)"/);
      if (iValMatch && iValMatch[1] !== 'false' && iValMatch[1] !== '0') italic = true;
    }

    var textMatch = runContent.match(/<w:t[^>]*>([^<]*)<\/w:t>/);
    if (textMatch) {
      segments.push({
        text: textMatch[1],
        bold: bold,
        italic: italic,
        superscript: superscript,
        subscript: subscript,
      });
    }
  }

  return segments;
}

/**
 * Parse a single w:p element into a structured element.
 * @param {string} pXml - Raw paragraph XML
 * @param {Object} styleMap - Style ID to heading level map
 * @param {Object} footnoteMap - Footnote ID to text map
 * @returns {Object|null}
 */
function parseParagraph(pXml, styleMap, footnoteMap) {
  var styleMatch = pXml.match(/w:pStyle\s[^>]*w:val="([^"]*)"/);
  var styleId = styleMatch ? styleMatch[1] : '';
  var headingLevel = styleMap[styleId] || 0;

  var textSegments = extractFormattedText(pXml, footnoteMap);
  var plainText = textSegments.map(function(s) { return s.text; }).join('');

  if (!plainText.trim() && !hasImage(pXml)) return null;

  // Heuristic heading detection for documents without heading styles
  if (headingLevel === 0 && !hasImage(pXml) && plainText.trim().length > 0 && plainText.trim().length < 120) {
    var allBold = textSegments.length > 0 && textSegments.every(function(s) {
      return s.bold || !s.text.trim();
    });
    var hasTextSegs = textSegments.some(function(s) { return s.text.trim().length > 0; });
    var looksLikeHeading = allBold && hasTextSegs &&
      !plainText.trim().match(/\.\s*$/) &&
      !plainText.match(/\(\d{4}\)/) &&
      !plainText.match(/^\*\*/) &&
      !plainText.match(/^Keywords?:/i);

    if (looksLikeHeading) {
      var outlineLvlMatch = pXml.match(/w:outlineLvl[^>]*w:val="(\d+)"/);
      if (outlineLvlMatch) {
        headingLevel = parseInt(outlineLvlMatch[1]) + 1;
      } else {
        headingLevel = 2;
      }
    }
  }

  var isCentered = pXml.indexOf('w:jc w:val="center"') !== -1;

  var elem = {
    type: headingLevel > 0 ? 'heading' : 'paragraph',
    level: headingLevel,
    text: plainText.trim(),
    segments: textSegments,
    styleId: styleId,
    hasImage: hasImage(pXml),
    imageRIds: extractImageRIds(pXml),
    isCentered: isCentered,
  };

  return elem;
}

/**
 * Parse a w:tbl element into rows/cells.
 * @param {string} tblXml - Raw table XML
 * @returns {Object|null}
 */
function parseTable(tblXml) {
  var rows = [];
  var rowRegex = /<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g;
  var rm;

  while ((rm = rowRegex.exec(tblXml)) !== null) {
    var cells = [];
    var cellRegex = /<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/g;
    var cm;
    var rowContent = rm[1];

    while ((cm = cellRegex.exec(rowContent)) !== null) {
      var cellText = extractTextFromXml(cm[1]);
      cells.push(cellText.trim());
    }
    if (cells.length > 0) rows.push(cells);
  }

  if (rows.length === 0) return null;

  return {
    type: 'table',
    rows: rows,
    headerRow: rows[0],
    dataRows: rows.slice(1),
  };
}

/**
 * Parse footnotes.xml content into a map of ID -> text.
 * @param {string} footnotesXml
 * @returns {Object<string, string>}
 */
function parseFootnotes(footnotesXml) {
  var footnoteMap = {};
  if (!footnotesXml) return footnoteMap;

  var fnRegex = /<w:footnote\s[^>]*w:id="(\d+)"[^>]*>([\s\S]*?)<\/w:footnote>/g;
  var fnMatch;
  while ((fnMatch = fnRegex.exec(footnotesXml)) !== null) {
    var fnId = fnMatch[1];
    if (fnId === '0' || fnId === '-1') continue;
    var fnText = extractTextFromXml(fnMatch[2]);
    if (fnText.trim()) {
      footnoteMap[fnId] = fnText.trim();
    }
  }

  return footnoteMap;
}

/**
 * Parse the document body into an array of elements (paragraphs, headings, tables).
 * @param {string} docXml - Full document.xml content
 * @param {Object} styleMap - Style ID to heading level map
 * @param {Object} footnoteMap - Footnote ID to text map
 * @returns {Array<Object>}
 */
function parseDocumentXml(docXml, styleMap, footnoteMap) {
  var elements = [];

  var bodyMatch = docXml.match(/<w:body>([\s\S]*)<\/w:body>/);
  if (!bodyMatch) return elements;
  var bodyContent = bodyMatch[1];

  var idx = 0;
  while (idx < bodyContent.length) {
    var nextP = bodyContent.indexOf('<w:p ', idx);
    var nextP2 = bodyContent.indexOf('<w:p>', idx);
    if (nextP === -1) nextP = Infinity;
    if (nextP2 === -1) nextP2 = Infinity;
    nextP = Math.min(nextP, nextP2);

    var nextTbl = bodyContent.indexOf('<w:tbl>', idx);
    if (nextTbl === -1) nextTbl = bodyContent.indexOf('<w:tbl ', idx);
    if (nextTbl === -1) nextTbl = Infinity;

    if (nextP === Infinity && nextTbl === Infinity) break;

    if (nextP <= nextTbl) {
      var pEnd = findClosingTag(bodyContent, nextP, 'w:p');
      if (pEnd === -1) { idx = nextP + 1; continue; }
      var pXml = bodyContent.substring(nextP, pEnd);
      var pElem = parseParagraph(pXml, styleMap, footnoteMap);
      if (pElem) elements.push(pElem);
      idx = pEnd;
    } else {
      var tEnd = findClosingTag(bodyContent, nextTbl, 'w:tbl');
      if (tEnd === -1) { idx = nextTbl + 1; continue; }
      var tXml = bodyContent.substring(nextTbl, tEnd);
      var tElem = parseTable(tXml);
      if (tElem) elements.push(tElem);
      idx = tEnd;
    }
  }

  return elements;
}

// ============================================================================
// FORMATTING
// ============================================================================

/**
 * Convert formatted text segments to LaTeX markup.
 * @param {Array} segments
 * @returns {string}
 */
function segmentsToLatex(segments) {
  var result = '';

  for (var i = 0; i < segments.length; i++) {
    var seg = segments[i];
    var text = seg.text;

    // Unescape XML entities
    text = decodeEntities(text);

    // Escape LaTeX special chars
    text = text
      .replace(/\\/g, '\\textbackslash{}')
      .replace(/&/g, '\\&')
      .replace(/%/g, '\\%')
      .replace(/\$/g, '\\$')
      .replace(/#/g, '\\#')
      .replace(/_/g, '\\_')
      .replace(/~/g, '\\textasciitilde{}');

    // Smart quotes and dashes
    text = text
      .replace(/\u201C/g, '``')
      .replace(/\u201D/g, "''")
      .replace(/\u2018/g, '`')
      .replace(/\u2019/g, "'")
      .replace(/\u2013/g, '--')
      .replace(/\u2014/g, '---')
      .replace(/\u2026/g, '\\ldots{}')
      .replace(/\u00A0/g, '~');

    // Apply formatting
    if (seg.bold && seg.italic) {
      text = '\\textbf{\\textit{' + text + '}}';
    } else if (seg.bold) {
      text = '\\textbf{' + text + '}';
    } else if (seg.italic) {
      text = '\\textit{' + text + '}';
    }

    if (seg.superscript) {
      text = '\\textsuperscript{' + text + '}';
    } else if (seg.subscript) {
      text = '\\textsubscript{' + text + '}';
    }

    result += text;

    if (seg.footnote) {
      var fnText = seg.footnote
        .replace(/&/g, '\\&')
        .replace(/%/g, '\\%')
        .replace(/_/g, '\\_');
      result += '\\footnote{' + fnText + '}';
    }
  }

  return result;
}

/**
 * Convert a parsed table to LaTeX booktabs format.
 * @param {Object} table
 * @returns {string}
 */
function tableToLatex(table) {
  var numCols = table.headerRow.length;
  var colSpec = 'l' + 'c'.repeat(numCols - 1);

  var lines = [];
  lines.push('\\begin{table}[htbp]');
  lines.push('\\centering');
  lines.push('\\begin{tabular}{' + colSpec + '}');
  lines.push('\\toprule');

  var headerCells = table.headerRow.map(function(c) { return escapeTableCell(c); });
  lines.push(headerCells.join(' & ') + ' \\\\');
  lines.push('\\midrule');

  for (var r = 0; r < table.dataRows.length; r++) {
    var row = table.dataRows[r];
    while (row.length < numCols) row.push('');
    var cells = row.map(function(c) { return escapeTableCell(c); });
    lines.push(cells.join(' & ') + ' \\\\');
  }

  lines.push('\\bottomrule');
  lines.push('\\end{tabular}');
  lines.push('\\end{table}');

  return lines.join('\n');
}

// ============================================================================
// UNIFORM FORMATTING STRIPPING
// ============================================================================

/**
 * Strip uniform formatting if > 80% of paragraphs share bold or italic.
 * This detects document-level formatting artifacts.
 * @param {Array} elements
 */
function stripUniformFormatting(elements) {
  var totalParas = 0;
  var allBoldParas = 0;
  var allItalicParas = 0;

  for (var i = 0; i < elements.length; i++) {
    var elem = elements[i];
    if (elem.type !== 'paragraph' || !elem.segments || elem.segments.length === 0) continue;
    var hasText = elem.segments.some(function(s) { return s.text && s.text.trim().length > 0; });
    if (!hasText) continue;

    totalParas++;
    var allBold = elem.segments.every(function(s) { return s.bold || !s.text.trim(); });
    var allItalic = elem.segments.every(function(s) { return s.italic || !s.text.trim(); });
    if (allBold) allBoldParas++;
    if (allItalic) allItalicParas++;
  }

  if (totalParas === 0) return;

  var boldRatio = allBoldParas / totalParas;
  var italicRatio = allItalicParas / totalParas;

  if (boldRatio > 0.8) {
    for (var j = 0; j < elements.length; j++) {
      if (elements[j].type === 'paragraph' && elements[j].segments) {
        elements[j].segments.forEach(function(s) { s.bold = false; });
      }
    }
  }

  if (italicRatio > 0.8) {
    for (var k = 0; k < elements.length; k++) {
      if (elements[k].type === 'paragraph' && elements[k].segments) {
        elements[k].segments.forEach(function(s) { s.italic = false; });
      }
    }
  }
}

// ============================================================================
// STRUCTURE ANALYSIS
// ============================================================================

/**
 * Analyze document elements to identify title, abstract, keywords,
 * authors, affiliations, and sections.
 * @param {Array} elements
 * @returns {Object}
 */
function analyzeStructure(elements) {
  var structure = {
    title: '',
    authors: [],
    affiliations: [],
    abstract: '',
    keywords: '',
    sections: [],
    correspondingAuthor: '',
    wordCount: '',
    targetJournal: '',
  };

  var inAbstract = false;
  var abstractParts = [];
  var inReferences = false;
  var titleFound = false;

  for (var i = 0; i < elements.length; i++) {
    var elem = elements[i];
    var text = elem.text || '';

    if (!text.trim() && elem.type !== 'table' && !elem.hasImage) continue;

    // Detect title
    if (!titleFound && i < 15 &&
        !text.match(/^(Authors?:|Affiliation|\[Author|\[Affiliation|Word count|Target journal|Corresponding|Abstract$|Keywords?:)/i)) {
      if ((elem.type === 'paragraph' && text.length > 20) ||
          (elem.type === 'heading' && text.length > 20 && !text.match(/^(Introduction|Abstract|Theoretical|Literature|Method|Data|Results|Discussion|Conclusion|References)$/i))) {
        structure.title = text;
        titleFound = true;
        continue;
      }
    }

    // Metadata lines after title, before abstract heading
    if (titleFound && !inAbstract && elem.type === 'paragraph' && !structure.sections.length) {
      if (text.match(/^(Authors?:|\[Author)/i)) continue;
      if (text.match(/^(Affiliation|\[Affiliation)/i)) continue;
      if (text.match(/^Corresponding\s+author/i)) {
        structure.correspondingAuthor = text;
        continue;
      }
      if (text.match(/^Word\s+count/i)) {
        structure.wordCount = text;
        continue;
      }
      if (text.match(/^Target\s+journal/i)) {
        structure.targetJournal = text;
        continue;
      }
      if (text.match(/^[A-Z][a-z]+\s+[A-Z]/) && text.length < 80 && !text.match(/\d{4}/)) {
        structure.authors.push(text);
        continue;
      }
      if (text.match(/^(Department|University|GESIS|Institute|School|Faculty|\d+\s)/i)) {
        structure.affiliations.push(text);
        continue;
      }
    }

    // Detect Abstract heading
    if (text.match(/^Abstract$/i) && (elem.type === 'heading' ||
        (elem.type === 'paragraph' && elem.segments && elem.segments.length > 0 && elem.segments[0].bold))) {
      inAbstract = true;
      continue;
    }

    // Abstract content
    if (inAbstract && elem.type === 'heading') {
      inAbstract = false;
      structure.abstract = abstractParts.join('\n\n');
    }

    if (inAbstract && elem.type === 'paragraph') {
      if (text.match(/^Keywords?:/i)) {
        structure.keywords = text.replace(/^Keywords?:\s*/i, '');
        inAbstract = false;
        structure.abstract = abstractParts.join('\n\n');
        continue;
      }
      abstractParts.push(text);
      continue;
    }

    // References heading
    if (elem.type === 'heading' && text.match(/^References$/i)) {
      inReferences = true;
      continue;
    }

    if (inReferences) continue;

    // Track sections
    if (elem.type === 'heading') {
      structure.sections.push({
        level: elem.level,
        title: text,
        elements: [],
      });
    }
  }

  if (inAbstract && abstractParts.length > 0) {
    structure.abstract = abstractParts.join('\n\n');
  }

  return structure;
}

// ============================================================================
// RELATIONSHIP PARSING
// ============================================================================

/**
 * Parse document.xml.rels to build a mapping of relationship IDs to targets.
 * @param {string} relsXml
 * @returns {Object<string, string>}
 */
function parseRelationships(relsXml) {
  var relationships = {};
  if (!relsXml) return relationships;
  var relRegex = /Id="([^"]+)"[^>]*Target="([^"]+)"/g;
  var rm;
  while ((rm = relRegex.exec(relsXml)) !== null) {
    relationships[rm[1]] = rm[2];
  }
  return relationships;
}

// ============================================================================
// LATEX CLASS
// ============================================================================

class Latex {

  /**
   * Convert a docx Workspace to LaTeX.
   *
   * @param {Workspace} ws - An open docex Workspace
   * @param {Object} [options] - Conversion options
   * @param {string} [options.documentClass='article'] - LaTeX document class
   * @param {string[]} [options.packages] - Additional LaTeX packages to include
   * @param {string} [options.bibFile='references'] - Bibliography file name (without .bib)
   * @returns {string} Complete LaTeX document
   */
  static convert(ws, options) {
    var opts = options || {};
    var documentClass = opts.documentClass || 'article';
    var extraPackages = opts.packages || [];
    var bibFile = opts.bibFile || 'references';

    // Read XML sources from workspace
    var docXml = ws.docXml;
    var stylesXml = ws.stylesXml || '';

    // Read footnotes if available
    var footnotesXml = '';
    try {
      footnotesXml = ws._readFile('word/footnotes.xml');
    } catch (e) {
      // No footnotes file -- that is fine
    }

    // Read relationships
    var relsXml = '';
    try {
      relsXml = ws.relsXml;
    } catch (e) {
      // No rels file
    }

    // Build maps
    var styleMap = buildStyleMap(stylesXml);
    var footnoteMap = parseFootnotes(footnotesXml);
    var relationships = parseRelationships(relsXml);

    // Parse document into elements
    var elements = parseDocumentXml(docXml, styleMap, footnoteMap);

    // Strip uniform formatting artifacts
    stripUniformFormatting(elements);

    // Analyze structure
    var structure = analyzeStructure(elements);

    // Generate LaTeX
    return Latex._generate(structure, elements, relationships, {
      documentClass: documentClass,
      extraPackages: extraPackages,
      bibFile: bibFile,
    });
  }

  /**
   * Generate the full LaTeX document string.
   * @param {Object} structure - Analyzed document structure
   * @param {Array} elements - Parsed elements
   * @param {Object} relationships - rId to target map
   * @param {Object} genOpts - Generation options
   * @returns {string}
   * @private
   */
  static _generate(structure, elements, relationships, genOpts) {
    var lines = [];

    // Preamble
    lines.push('\\documentclass[12pt]{' + genOpts.documentClass + '}');
    lines.push('\\usepackage[margin=1in]{geometry}');
    lines.push('\\usepackage{graphicx}');
    lines.push('\\usepackage{booktabs}');
    lines.push('\\usepackage{setspace}');
    lines.push('\\usepackage{hyperref}');

    // Extra packages
    for (var pi = 0; pi < genOpts.extraPackages.length; pi++) {
      lines.push('\\usepackage{' + genOpts.extraPackages[pi] + '}');
    }

    lines.push('\\doublespacing');
    lines.push('');

    // Title
    var escapedTitle = escapeForLatex(structure.title || 'Untitled');
    lines.push('\\title{' + escapedTitle + '}');

    // Authors and affiliations
    if (structure.authors.length > 0) {
      var authorLines = structure.authors.map(function(a, idx) {
        var escaped = escapeForLatex(a);
        return escaped + '\\textsuperscript{' + (idx + 1) + '}';
      });
      var affLines = structure.affiliations.map(function(a, idx) {
        var escaped = escapeForLatex(a);
        return '\\textsuperscript{' + (idx + 1) + '}' + escaped;
      });
      lines.push('\\author{' + authorLines.join(' \\and ') +
        (affLines.length > 0 ? ' \\\\\n' + affLines.join(' \\\\') : '') + '}');
    } else {
      lines.push('\\author{}');
    }

    lines.push('\\date{}');
    lines.push('');
    lines.push('\\begin{document}');
    lines.push('\\maketitle');
    lines.push('');

    // Abstract
    if (structure.abstract) {
      lines.push('\\begin{abstract}');
      lines.push(escapeForLatex(structure.abstract));
      lines.push('\\end{abstract}');
      lines.push('');
    }

    if (structure.keywords) {
      lines.push('\\noindent\\textbf{Keywords:} ' + escapeForLatex(structure.keywords));
      lines.push('');
    }

    // Build skip set for metadata elements
    var skipTexts = {};
    if (structure.title) skipTexts[structure.title] = true;
    if (structure.abstract) {
      structure.abstract.split('\n\n').forEach(function(p) { skipTexts[p.trim()] = true; });
    }
    if (structure.keywords) skipTexts[structure.keywords] = true;
    if (structure.correspondingAuthor) skipTexts[structure.correspondingAuthor] = true;
    if (structure.wordCount) skipTexts[structure.wordCount] = true;
    if (structure.targetJournal) skipTexts[structure.targetJournal] = true;
    structure.authors.forEach(function(a) { skipTexts[a] = true; });
    structure.affiliations.forEach(function(a) { skipTexts[a] = true; });

    // Determine heading level mapping
    var firstContentHeadingIdx = -1;
    var minHeadingLevel = Infinity;
    for (var ei = 0; ei < elements.length; ei++) {
      var eText = (elements[ei].text || '').trim();
      if (elements[ei].type === 'heading' && elements[ei].level > 0 &&
          !eText.match(/^(Abstract|References|Keywords?)$/i) &&
          !skipTexts[eText]) {
        if (eText.match(/^Word\s+count/i)) continue;
        if (firstContentHeadingIdx === -1) firstContentHeadingIdx = ei;
        minHeadingLevel = Math.min(minHeadingLevel, elements[ei].level);
      }
    }
    if (minHeadingLevel === Infinity) minHeadingLevel = 1;

    function levelToCmd(level) {
      var adjusted = level - minHeadingLevel;
      if (adjusted <= 0) return '\\section';
      if (adjusted === 1) return '\\subsection';
      if (adjusted === 2) return '\\subsubsection';
      return '\\paragraph';
    }

    // Body elements
    var startIdx = firstContentHeadingIdx >= 0 ? firstContentHeadingIdx : 0;
    var inReferences = false;
    var figureCounter = 0;

    for (var i = startIdx; i < elements.length; i++) {
      var elem = elements[i];
      var text = elem.text || '';

      if (!text.trim() && elem.type !== 'table' && !elem.hasImage) continue;

      if (inReferences) continue;

      // Skip metadata elements
      if (skipTexts[text.trim()]) continue;
      if (text.match(/^(Abstract|Keywords?:|Word\s+count|Authors?:|Affiliation|Target\s+journal|Corresponding\s+author|\[Author|\[Affiliation)/i)) continue;

      if (elem.type === 'heading' && text.match(/^References$/i)) {
        inReferences = true;
        continue;
      }

      if (elem.type === 'heading') {
        var cmd = levelToCmd(elem.level);
        var sectionTitle = escapeForLatex(text);
        lines.push('');
        lines.push(cmd + '{' + sectionTitle + '}');
        lines.push('');
        continue;
      }

      if (elem.type === 'table') {
        lines.push('');
        lines.push(tableToLatex(elem));
        lines.push('');
        continue;
      }

      if (elem.type === 'paragraph') {
        // Handle images
        if (elem.hasImage && elem.imageRIds.length > 0) {
          for (var ri = 0; ri < elem.imageRIds.length; ri++) {
            var rId = elem.imageRIds[ri];
            var mediaTarget = relationships[rId];
            if (mediaTarget && mediaTarget.match(/^media\//)) {
              figureCounter++;
              var figName = mediaTarget.replace(/^media\//, '');
              lines.push('');
              lines.push('\\begin{figure}[htbp]');
              lines.push('\\centering');
              lines.push('\\includegraphics[width=\\textwidth]{figures/' + figName + '}');
              lines.push('\\caption{}');
              lines.push('\\label{fig:' + figureCounter + '}');
              lines.push('\\end{figure}');
              lines.push('');
            }
          }
          if (!text.trim()) continue;
        }

        // Figure/table placeholders
        var figPlaceholder = text.match(/^\[(?:Figure|Table)\s+\d+\s+about\s+here[^\]]*\]$/i);
        if (figPlaceholder) {
          lines.push('');
          lines.push('% ' + text);
          lines.push('');
          continue;
        }

        // Convert to LaTeX
        var latex;
        if (elem.segments && elem.segments.length > 0) {
          latex = segmentsToLatex(elem.segments);
        } else {
          latex = escapeForLatex(text);
        }

        lines.push(latex);
        lines.push('');
      }
    }

    // Bibliography
    lines.push('\\bibliography{' + genOpts.bibFile + '}');
    lines.push('');
    lines.push('\\end{document}');

    return lines.join('\n');
  }
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = { Latex };
