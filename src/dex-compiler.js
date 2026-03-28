'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const {
  isXmlishPart,
  parseDex,
  parseXml,
  serializeXml,
} = require('./dex-lossless');
const { listFiles } = require('./dex-decompiler');

class DexCompiler {
  static compile(input, opts = {}) {
    // If a string is passed, try DexParser first (human-readable format), then parseDex (package format)
    let ast = input;
    if (typeof input === 'string') {
      try {
        const { DexParser } = require('./dex-markdown-parser');
        const parsed = DexParser.parse(input);
        if (parsed && parsed.body !== undefined) {
          return DexCompiler._compileHumanReadable(parsed, opts);
        }
      } catch (_) {}
      ast = parseDex(input);
    }
    // If ast has 'body' property (DexParser format), use human-readable compiler
    if (ast && ast.body !== undefined && !ast.parts) {
      return DexCompiler._compileHumanReadable(ast, opts);
    }
    if (!ast || ast.type !== 'dex') throw new Error('DexCompiler.compile expects a dex AST or string');

    const outputPath = opts.output || '/tmp/dex-compiled-' + crypto.randomBytes(8).toString('hex') + '.docx';
    const absOutput = path.resolve(outputPath);
    const tmpDir = fs.mkdtempSync('/tmp/dex-compile-');

    try {
      for (const part of ast.parts || []) {
        const absPartPath = path.join(tmpDir, part.path);
        fs.mkdirSync(path.dirname(absPartPath), { recursive: true });
        if (part.partType === 'binary') {
          fs.writeFileSync(absPartPath, Buffer.from(part.data || '', part.encoding || 'base64'));
        } else {
          fs.writeFileSync(absPartPath, serializeXml(part.nodes || []), 'utf-8');
        }
      }

      if (fs.existsSync(absOutput)) fs.unlinkSync(absOutput);
      execFileSync('zip', ['-r', '-q', absOutput, '.'], {
        cwd: tmpDir,
        stdio: 'pipe',
      });

      const result = {
        path: absOutput,
        partCount: (ast.parts || []).length,
      };

      if (opts.verifyAgainst) {
        const comparison = DexCompiler.compareDocx(opts.verifyAgainst, absOutput);
        result.verified = comparison.equal;
        result.differences = comparison.differences;
        if (!comparison.equal && opts.strictVerify !== false) {
          throw new Error(
            'round-trip verification failed: ' +
            comparison.differences.slice(0, 5).map(diff => `${diff.path}:${diff.kind}`).join(', ')
          );
        }
      }

      return result;
    } finally {
      execFileSync('rm', ['-rf', tmpDir], { stdio: 'pipe' });
    }
  }

  static compileString(dexString, opts = {}) {
    return DexCompiler.compile(dexString, opts);
  }

  /**
   * Compile a human-readable DexParser AST (from DexParser.parse()) into a .docx file.
   * @param {{ frontmatter: object, body: Array }} ast
   * @param {object} opts
   */
  static _compileHumanReadable(ast, opts = {}) {
    const xmlLib = require('./xml');
    const outputPath = opts.output || '/tmp/dex-compiled-' + crypto.randomBytes(8).toString('hex') + '.docx';
    const absOutput = path.resolve(outputPath);
    const tmpDir = fs.mkdtempSync('/tmp/dex-compile-hr-');
    const body = ast.body || [];

    function genId() { return crypto.randomBytes(4).toString('hex').toUpperCase(); }
    function esc(t) { return xmlLib.escapeXml(String(t || '')); }

    function buildRunXml(run) {
      if (typeof run === 'string' || run.type === 'text') {
        const t = typeof run === 'string' ? run : run.text;
        return '<w:r><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
      }
      if (run.type === 'bold') return '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'italic') return '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'underline') {
        const uVal = run.underlineType || 'single';
        return '<w:r><w:rPr><w:u w:val="' + esc(uVal) + '"/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      }
      if (run.type === 'superscript') return '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'subscript') return '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'strike') return '<w:r><w:rPr><w:strike/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'dstrike') return '<w:r><w:rPr><w:dstrike/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'size') return '<w:r><w:rPr><w:sz w:val="' + esc(run.size || '24') + '"/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'smallcaps') return '<w:r><w:rPr><w:smallCaps/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'caps') return '<w:r><w:rPr><w:caps/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'hidden') return '<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'highlight') return '<w:r><w:rPr><w:highlight w:val="' + esc(run.color || 'yellow') + '"/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'color') return '<w:r><w:rPr><w:color w:val="' + esc(run.color || '000000') + '"/></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'font') {
        let rFonts = '<w:rFonts w:ascii="' + esc(run.font || '') + '" w:hAnsi="' + esc(run.fontHAnsi || run.font || '') + '"';
        if (run.fontCs) rFonts += ' w:cs="' + esc(run.fontCs) + '"';
        if (run.fontEastAsia) rFonts += ' w:eastAsia="' + esc(run.fontEastAsia) + '"';
        rFonts += '/>';
        return '<w:r><w:rPr>' + rFonts + '</w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      }
      if (run.type === 'del') return '<w:del w:id="' + (run.id || 1) + '" w:author="' + esc(run.author || '') + '" w:date="' + esc(run.date || '') + '"><w:r><w:delText xml:space="preserve">' + esc(run.text) + '</w:delText></w:r></w:del>';
      if (run.type === 'ins') return '<w:ins w:id="' + (run.id || 2) + '" w:author="' + esc(run.author || '') + '" w:date="' + esc(run.date || '') + '"><w:r><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r></w:ins>';
      if (run.type === 'movefrom') return '<w:moveFrom w:id="' + (run.id || 1) + '" w:author="' + esc(run.author || '') + '" w:date="' + esc(run.date || '') + '"><w:r><w:delText xml:space="preserve">' + esc(run.text) + '</w:delText></w:r></w:moveFrom>';
      if (run.type === 'moveto') return '<w:moveTo w:id="' + (run.id || 2) + '" w:author="' + esc(run.author || '') + '" w:date="' + esc(run.date || '') + '"><w:r><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r></w:moveTo>';
      if (run.type === 'fmtchange') return '<w:r><w:rPr><w:rPrChange w:author="' + esc(run.author || '') + '" w:date="' + esc(run.date || '') + '"><w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">' + esc(run.text) + '</w:t></w:r>';
      if (run.type === 'footnote') {
        footnotes.push(run);
        return '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:r>'
          + '<w:r><w:footnoteReference w:id="' + (run.id || footnotes.length) + '"/></w:r>';
      }
      if (run.type === 'comment-start') return '<w:commentRangeStart w:id="' + (run.id || 0) + '"/>';
      if (run.type === 'comment-end') {
        return '<w:commentRangeEnd w:id="' + (run.id || 0) + '"/>'
          + '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
          + '<w:commentReference w:id="' + (run.id || 0) + '"/></w:r>';
      }
      if (run.type === 'bookmark-start') return '<w:bookmarkStart w:id="' + (run.id || 0) + '" w:name="' + esc(run.name || '') + '"/>';
      if (run.type === 'bookmark-end') return '<w:bookmarkEnd w:id="' + (run.id || 0) + '"/>';
      if (run.type === 'link') {
        if (run.anchor) return '<w:hyperlink w:anchor="' + esc(run.anchor) + '"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">' + esc(run.text || '') + '</w:t></w:r></w:hyperlink>';
        return '<w:hyperlink r:id="' + esc(run.rId || '') + '"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">' + esc(run.text || '') + '</w:t></w:r></w:hyperlink>';
      }
      if (run.type === 'linebreak') return '<w:r><w:br/></w:r>';
      if (run.type === 'colbreak') return '<w:r><w:br w:type="column"/></w:r>';
      if (run.type === 'endnote') return '<w:r><w:endnoteReference w:id="' + (run.id || 0) + '"/></w:r>';
      if (run.type === 'sym') return '<w:r><w:sym w:font="Symbol" w:char="' + esc(run.char || '') + '"/></w:r>';
      return '<w:r><w:t xml:space="preserve">' + esc(run.text || '') + '</w:t></w:r>';
    }

    function buildParagraphProperties(node) {
      const parts = [];
      // Paragraph style
      if (node.style) parts.push('<w:pStyle w:val="' + esc(node.style) + '"/>');
      // Keep with next
      if (node.keepnext) parts.push('<w:keepNext/>');
      // List numbering
      if (node.listId) {
        let numPr = '<w:numPr>';
        if (node.listLevel !== undefined && node.listLevel !== null) numPr += '<w:ilvl w:val="' + esc(node.listLevel) + '"/>';
        numPr += '<w:numId w:val="' + esc(node.listId) + '"/>';
        numPr += '</w:numPr>';
        parts.push(numPr);
      }
      // Right-to-left
      if (node.bidi) parts.push('<w:bidi/>');
      // Indentation
      if (node.indentLeft || node.indentRight || node.indentFirst || node.indentHanging) {
        let ind = '<w:ind';
        if (node.indentLeft) ind += ' w:left="' + esc(node.indentLeft) + '"';
        if (node.indentRight) ind += ' w:right="' + esc(node.indentRight) + '"';
        if (node.indentFirst) ind += ' w:firstLine="' + esc(node.indentFirst) + '"';
        if (node.indentHanging) ind += ' w:hanging="' + esc(node.indentHanging) + '"';
        ind += '/>';
        parts.push(ind);
      }
      // Spacing
      if (node.spacingLine || node.spacingBefore || node.spacingAfter || node.spacingRule) {
        let sp = '<w:spacing';
        if (node.spacingBefore) sp += ' w:before="' + esc(node.spacingBefore) + '"';
        if (node.spacingAfter) sp += ' w:after="' + esc(node.spacingAfter) + '"';
        if (node.spacingLine) sp += ' w:line="' + esc(node.spacingLine) + '"';
        if (node.spacingRule) sp += ' w:lineRule="' + esc(node.spacingRule) + '"';
        sp += '/>';
        parts.push(sp);
      }
      // Alignment (map 'justify' back to 'both')
      if (node.align) {
        const val = node.align === 'justify' ? 'both' : node.align;
        parts.push('<w:jc w:val="' + esc(val) + '"/>');
      }
      if (parts.length === 0) return '';
      return '<w:pPr>' + parts.join('') + '</w:pPr>';
    }

    const footnotes = [];
    const comments = [];
    let bodyXml = '';
    const sectPr = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

    for (const node of body) {
      if (node.type === 'heading') {
        const lvl = node.level || 1;
        const styleId = 'Heading' + lvl;
        const pid = node.id || genId();
        const runsXml = node.runs
          ? node.runs.map(buildRunXml).join('')
          : '<w:r><w:t xml:space="preserve">' + esc(node.text) + '</w:t></w:r>';
        bodyXml += '<w:p w14:paraId="' + pid + '" w14:textId="' + genId() + '">'
          + '<w:pPr><w:pStyle w:val="' + styleId + '"/></w:pPr>'
          + runsXml
          + '</w:p>';
      } else if (node.type === 'paragraph') {
        const pid = node.id || genId();
        const runs = node.runs || [{ type: 'text', text: node.text || '' }];
        const pPr = buildParagraphProperties(node);
        bodyXml += '<w:p w14:paraId="' + pid + '" w14:textId="' + genId() + '">'
          + pPr
          + runs.map(buildRunXml).join('')
          + '</w:p>';
      } else if (node.type === 'pagebreak') {
        bodyXml += '<w:p w14:paraId="' + genId() + '" w14:textId="' + genId() + '">'
          + '<w:r><w:br w:type="page"/></w:r>'
          + '</w:p>';
      } else if (node.type === 'table') {
        const rows = node.rows || [];
        const cols = node.cols || (rows[0] ? rows[0].length : 0);
        const colWidth = cols > 0 ? Math.floor(9360 / cols) : 9360;
        let tblXml = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid>';
        for (let c = 0; c < cols; c++) tblXml += '<w:gridCol w:w="' + colWidth + '"/>';
        tblXml += '</w:tblGrid>';
        for (let r = 0; r < rows.length; r++) {
          tblXml += '<w:tr>';
          for (let c = 0; c < rows[r].length; c++) {
            tblXml += '<w:tc><w:tcPr><w:tcW w:w="' + colWidth + '" w:type="dxa"/></w:tcPr>'
              + '<w:p><w:r><w:t xml:space="preserve">' + esc(rows[r][c]) + '</w:t></w:r></w:p>'
              + '</w:tc>';
          }
          tblXml += '</w:tr>';
        }
        tblXml += '</w:tbl>';
        bodyXml += tblXml;
      } else if (node.type === 'comment') {
        comments.push(node);
      } else if (node.type === 'reply') {
        comments.push({ ...node, isReply: true });
      }
    }

    // Build document XML
    const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      + 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      + 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      + 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
      + 'mc:Ignorable="w14 w15">'
      + '<w:body>' + bodyXml + sectPr + '</w:body></w:document>';

    // Build styles XML
    const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      + 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
      + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
      + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>'
      + '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:i/></w:rPr></w:style>'
      + '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>'
      + '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>'
      + '</w:styles>';

    // Build comments XML if needed
    let commentsXmlStr = null;
    let commentsExtXmlStr = null;
    if (comments.length > 0) {
      const mainComments = comments.filter(c => !c.isReply);
      const replyComments = comments.filter(c => c.isReply);
      let cBody = '';
      let extBody = '';
      for (const c of [...mainComments, ...replyComments]) {
        const pid = genId();
        const tid = genId();
        cBody += '<w:comment w:id="' + (c.id || 0) + '" w:author="' + esc(c.author || '') + '" w:date="' + esc(c.date || '') + '">'
          + '<w:p w14:paraId="' + pid + '" w14:textId="' + tid + '">'
          + '<w:r><w:t xml:space="preserve">' + esc(c.text || '') + '</w:t></w:r>'
          + '</w:p></w:comment>';
        if (c.isReply) {
          // Find parent's paraId - simplified: use parent id as reference
          const parentParaId = genId();
          extBody += '<w15:commentEx w15:paraId="' + pid + '" w15:paraIdParent="' + parentParaId + '" w15:done="0"/>';
        } else {
          extBody += '<w15:commentEx w15:paraId="' + pid + '" w15:done="0"/>';
        }
      }
      commentsXmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        + 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        + 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
        + cBody + '</w:comments>';
      commentsExtXmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
        + extBody + '</w15:commentsEx>';
    }

    // Build footnotes XML if needed
    let footnotesXmlStr = null;
    if (footnotes.length > 0) {
      let fnBody = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        + '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
      for (const fn of footnotes) {
        fnBody += '<w:footnote w:id="' + (fn.id || 1) + '">'
          + '<w:p><w:r><w:t xml:space="preserve">' + esc(fn.text || '') + '</w:t></w:r></w:p>'
          + '</w:footnote>';
      }
      footnotesXmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        + 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        + fnBody + '</w:footnotes>';
    }

    // Build rels
    let relBody = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
    if (commentsXmlStr) {
      relBody += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>';
      relBody += '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>';
    }
    if (footnotesXmlStr) {
      relBody += '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>';
    }
    const wordRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + relBody + '</Relationships>';

    // Build content types
    let ctBody = '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';
    if (commentsXmlStr) {
      ctBody += '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>';
      ctBody += '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>';
    }
    if (footnotesXmlStr) {
      ctBody += '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>';
    }
    const contentTypesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + ctBody + '</Types>';

    const rootRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>';

    try {
      fs.mkdirSync(path.join(tmpDir, 'word', '_rels'), { recursive: true });
      fs.mkdirSync(path.join(tmpDir, '_rels'), { recursive: true });
      fs.writeFileSync(path.join(tmpDir, '[Content_Types].xml'), contentTypesXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, '_rels', '.rels'), rootRelsXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', 'document.xml'), docXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', 'styles.xml'), stylesXml, 'utf-8');
      fs.writeFileSync(path.join(tmpDir, 'word', '_rels', 'document.xml.rels'), wordRelsXml, 'utf-8');
      if (commentsXmlStr) {
        fs.writeFileSync(path.join(tmpDir, 'word', 'comments.xml'), commentsXmlStr, 'utf-8');
        fs.writeFileSync(path.join(tmpDir, 'word', 'commentsExtended.xml'), commentsExtXmlStr, 'utf-8');
      }
      if (footnotesXmlStr) {
        fs.writeFileSync(path.join(tmpDir, 'word', 'footnotes.xml'), footnotesXmlStr, 'utf-8');
      }

      const outputDir = path.dirname(absOutput);
      if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
      if (fs.existsSync(absOutput)) fs.unlinkSync(absOutput);
      execFileSync('zip', ['-r', '-q', absOutput, '.'], { cwd: tmpDir, stdio: 'pipe' });

      // Count paragraphs
      const pMatches = docXml.match(/<w:p[\s>]/g) || [];
      const result = {
        path: absOutput,
        paragraphCount: pMatches.length,
        partCount: body.length,
      };

      if (opts.verifyAgainst) {
        const comparison = DexCompiler.compareDocx(opts.verifyAgainst, absOutput);
        result.verified = comparison.equal;
        result.differences = comparison.differences;
        if (!comparison.equal && opts.strictVerify !== false) {
          throw new Error('round-trip verification failed');
        }
      }
      return result;
    } finally {
      try { execFileSync('rm', ['-rf', tmpDir], { stdio: 'pipe' }); } catch (_) {}
    }
  }

  static assertRoundTrip(docxPath, opts = {}) {
    const { DexDecompiler } = require('./dex-decompiler');
    const { serializeDex } = require('./dex-lossless');
    const ast = DexDecompiler.toAst(docxPath);
    const dex = serializeDex(ast);
    const output = opts.output || '/tmp/dex-roundtrip-' + crypto.randomBytes(8).toString('hex') + '.docx';
    const compiled = DexCompiler.compile(ast, {
      ...opts,
      output,
      verifyAgainst: docxPath,
    });
    return {
      dex,
      ast,
      ...compiled,
    };
  }

  static compareDocx(leftPath, rightPath) {
    const left = unpackDocx(leftPath);
    const right = unpackDocx(rightPath);
    try {
      const leftFiles = listFiles(left.rootDir);
      const rightFiles = listFiles(right.rootDir);
      const allPaths = new Set([...leftFiles, ...rightFiles]);
      const differences = [];

      for (const relPath of Array.from(allPaths).sort()) {
        const inLeft = leftFiles.includes(relPath);
        const inRight = rightFiles.includes(relPath);
        if (!inLeft || !inRight) {
          differences.push({
            path: relPath,
            kind: 'missing-part',
            left: inLeft,
            right: inRight,
          });
          continue;
        }

        const leftAbs = path.join(left.rootDir, relPath);
        const rightAbs = path.join(right.rootDir, relPath);

        if (isXmlishPart(relPath)) {
          const leftXml = fs.readFileSync(leftAbs, 'utf-8');
          const rightXml = fs.readFileSync(rightAbs, 'utf-8');
          if (leftXml !== rightXml) {
            let normalizedEqual = false;
            try {
              normalizedEqual = serializeXml(parseXml(leftXml)) === serializeXml(parseXml(rightXml));
            } catch (_) {
              normalizedEqual = false;
            }
            differences.push({
              path: relPath,
              kind: normalizedEqual ? 'xml-lexical-diff' : 'xml-structural-diff',
            });
          }
          continue;
        }

        const leftBuf = fs.readFileSync(leftAbs);
        const rightBuf = fs.readFileSync(rightAbs);
        if (!leftBuf.equals(rightBuf)) {
          differences.push({
            path: relPath,
            kind: 'binary-diff',
          });
        }
      }

      return {
        equal: differences.length === 0,
        differences,
      };
    } finally {
      left.cleanup();
      right.cleanup();
    }
  }
}

function unpackDocx(docxPath) {
  const absPath = path.resolve(docxPath);
  if (!fs.existsSync(absPath)) throw new Error(`.docx not found: ${absPath}`);
  const tmpDir = fs.mkdtempSync('/tmp/dex-compare-');
  execFileSync('unzip', ['-q', absPath, '-d', tmpDir], { stdio: 'pipe' });
  return {
    rootDir: tmpDir,
    cleanup() {
      execFileSync('rm', ['-rf', tmpDir], { stdio: 'pipe' });
    },
  };
}

module.exports = { DexCompiler };
