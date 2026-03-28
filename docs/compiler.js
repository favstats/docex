/* ============================================================================
   compiler.js -- Browser-compatible .dex to .docx compiler
   Uses JSZip (loaded from CDN) to build .docx files in the browser.
   ============================================================================ */

var DexCompiler = (function() {
  'use strict';

  // ---------- XML helpers ----------
  function escXml(s) {
    return s
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  // Generate a random 8-char hex paraId
  function randomParaId() {
    var hex = '';
    for (var i = 0; i < 8; i++) hex += Math.floor(Math.random() * 16).toString(16).toUpperCase();
    return hex;
  }

  // ---------- .dex Parser ----------

  // Parse frontmatter: returns { meta: {}, bodyStart: number }
  function parseFrontmatter(lines) {
    var meta = {};
    if (lines[0] !== '---') return { meta: meta, bodyStart: 0 };
    var i = 1;
    while (i < lines.length && lines[i] !== '---') {
      var m = lines[i].match(/^(\w+):\s*"?([^"]*)"?\s*$/);
      if (m) meta[m[1]] = m[2];
      // Handle authors array
      var am = lines[i].match(/^\s+-\s+name:\s*"([^"]*)"/);
      if (am) {
        if (!meta.authors) meta.authors = [];
        meta.authors.push(am[1]);
      }
      i++;
    }
    return { meta: meta, bodyStart: i + 1 };
  }

  // Unescape literal braces
  function unescapeBraces(s) {
    return s.replace(/\\\{/g, '{').replace(/\\\}/g, '}');
  }

  // ---------- Inline parser ----------
  // Parses inline .dex markup into an array of "run" objects

  function parseInlineRuns(text, state) {
    var runs = [];
    var pos = 0;
    var len = text.length;
    var fmtStack = state || { bold: false, italic: false, underline: false, sup: false, sub: false, strike: false, dstrike: false, smallcaps: false, caps: false, hidden: false, size: null, font: null, color: null, highlight: null };

    while (pos < len) {
      // Check for escaped braces
      if (text[pos] === '\\' && pos + 1 < len && (text[pos + 1] === '{' || text[pos + 1] === '}')) {
        runs.push({ text: text[pos + 1], fmt: cloneFmt(fmtStack) });
        pos += 2;
        continue;
      }

      // Check for tag opening
      if (text[pos] === '{') {
        var tagEnd = findMatchingBrace(text, pos);
        if (tagEnd === -1) {
          runs.push({ text: '{', fmt: cloneFmt(fmtStack) });
          pos++;
          continue;
        }
        var tag = text.slice(pos + 1, tagEnd);

        // Closing tags
        if (tag === '/b') { fmtStack.bold = false; pos = tagEnd + 1; continue; }
        if (tag === '/i') { fmtStack.italic = false; pos = tagEnd + 1; continue; }
        if (tag === '/u') { fmtStack.underline = false; pos = tagEnd + 1; continue; }
        if (tag === '/sup') { fmtStack.sup = false; pos = tagEnd + 1; continue; }
        if (tag === '/sub') { fmtStack.sub = false; pos = tagEnd + 1; continue; }
        if (tag === '/strike') { fmtStack.strike = false; pos = tagEnd + 1; continue; }
        if (tag === '/dstrike') { fmtStack.dstrike = false; pos = tagEnd + 1; continue; }
        if (tag === '/smallcaps') { fmtStack.smallcaps = false; pos = tagEnd + 1; continue; }
        if (tag === '/caps') { fmtStack.caps = false; pos = tagEnd + 1; continue; }
        if (tag === '/hidden') { fmtStack.hidden = false; pos = tagEnd + 1; continue; }
        if (tag === '/size') { fmtStack.size = null; pos = tagEnd + 1; continue; }
        if (tag === '/font') { fmtStack.font = null; pos = tagEnd + 1; continue; }
        if (tag === '/color') { fmtStack.color = null; pos = tagEnd + 1; continue; }
        if (tag === '/highlight') { fmtStack.highlight = null; pos = tagEnd + 1; continue; }

        // Opening tags
        if (tag === 'b') { fmtStack.bold = true; pos = tagEnd + 1; continue; }
        if (tag === 'i') { fmtStack.italic = true; pos = tagEnd + 1; continue; }
        if (tag === 'u') { fmtStack.underline = true; pos = tagEnd + 1; continue; }
        if (tag === 'sup') { fmtStack.sup = true; pos = tagEnd + 1; continue; }
        if (tag === 'sub') { fmtStack.sub = true; pos = tagEnd + 1; continue; }
        if (tag === 'strike') { fmtStack.strike = true; pos = tagEnd + 1; continue; }
        if (tag === 'dstrike') { fmtStack.dstrike = true; pos = tagEnd + 1; continue; }
        if (tag === 'smallcaps') { fmtStack.smallcaps = true; pos = tagEnd + 1; continue; }
        if (tag === 'caps') { fmtStack.caps = true; pos = tagEnd + 1; continue; }
        if (tag === 'hidden') { fmtStack.hidden = true; pos = tagEnd + 1; continue; }

        // Size tag: {size 28}
        var sizeMatch = tag.match(/^size\s+(\d+)$/);
        if (sizeMatch) { fmtStack.size = sizeMatch[1]; pos = tagEnd + 1; continue; }

        // Font tag: {font "Name"}
        var fontMatch = tag.match(/^font\s+"([^"]+)"$/);
        if (fontMatch) { fmtStack.font = fontMatch[1]; pos = tagEnd + 1; continue; }

        // Color tag: {color XXXXXX}
        var colorMatch = tag.match(/^color\s+([A-Fa-f0-9]+)$/);
        if (colorMatch) { fmtStack.color = colorMatch[1]; pos = tagEnd + 1; continue; }

        // Highlight tag: {highlight name}
        var hlMatch = tag.match(/^highlight\s+(\w+)$/);
        if (hlMatch) { fmtStack.highlight = hlMatch[1]; pos = tagEnd + 1; continue; }

        // Footnote: {footnote id:N}text{/footnote}
        var fnMatch = tag.match(/^footnote\s+id:(\d+)$/);
        if (fnMatch) {
          var fnClose = text.indexOf('{/footnote}', tagEnd + 1);
          var fnText = fnClose !== -1 ? text.slice(tagEnd + 1, fnClose) : '';
          runs.push({ footnote: true, footnoteId: parseInt(fnMatch[1], 10), footnoteText: unescapeBraces(fnText), fmt: cloneFmt(fmtStack) });
          pos = fnClose !== -1 ? fnClose + '{/footnote}'.length : tagEnd + 1;
          continue;
        }

        // Del: {del id:N by:"Author" date:"..."}text{/del}
        var delMatch = tag.match(/^del\s+id:(\S+)\s+by:"([^"]*)"\s+date:"([^"]*)"$/);
        if (delMatch) {
          var delClose = text.indexOf('{/del}', tagEnd + 1);
          var delText = delClose !== -1 ? text.slice(tagEnd + 1, delClose) : '';
          runs.push({ del: true, revId: delMatch[1], revAuthor: delMatch[2], revDate: delMatch[3], text: unescapeBraces(delText), fmt: cloneFmt(fmtStack) });
          pos = delClose !== -1 ? delClose + '{/del}'.length : tagEnd + 1;
          continue;
        }

        // Ins: {ins id:N by:"Author" date:"..."}text{/ins}
        var insMatch = tag.match(/^ins\s+id:(\S+)\s+by:"([^"]*)"\s+date:"([^"]*)"$/);
        if (insMatch) {
          var insClose = text.indexOf('{/ins}', tagEnd + 1);
          var insText = insClose !== -1 ? text.slice(tagEnd + 1, insClose) : '';
          runs.push({ ins: true, revId: insMatch[1], revAuthor: insMatch[2], revDate: insMatch[3], text: unescapeBraces(insText), fmt: cloneFmt(fmtStack) });
          pos = insClose !== -1 ? insClose + '{/ins}'.length : tagEnd + 1;
          continue;
        }

        // Comment anchors
        var csMatch = tag.match(/^comment-start\s+id:(\d+)$/);
        if (csMatch) { runs.push({ commentStart: true, commentId: parseInt(csMatch[1], 10) }); pos = tagEnd + 1; continue; }
        var ceMatch = tag.match(/^comment-end\s+id:(\d+)$/);
        if (ceMatch) { runs.push({ commentEnd: true, commentId: parseInt(ceMatch[1], 10) }); pos = tagEnd + 1; continue; }

        // Bookmarks
        var bsMatch = tag.match(/^bookmark-start\s+id:(\d+)\s+name:"([^"]*)"$/);
        if (bsMatch) { runs.push({ bookmarkStart: true, bookmarkId: parseInt(bsMatch[1], 10), bookmarkName: bsMatch[2] }); pos = tagEnd + 1; continue; }
        var beMatch = tag.match(/^bookmark-end\s+id:(\d+)$/);
        if (beMatch) { runs.push({ bookmarkEnd: true, bookmarkId: parseInt(beMatch[1], 10) }); pos = tagEnd + 1; continue; }

        // Line break
        if (tag === 'br') { runs.push({ lineBreak: true }); pos = tagEnd + 1; continue; }
        if (tag === 'pagebreak') { runs.push({ pageBreak: true }); pos = tagEnd + 1; continue; }

        // Unknown tag - output as literal
        runs.push({ text: '{' + tag + '}', fmt: cloneFmt(fmtStack) });
        pos = tagEnd + 1;
        continue;
      }

      // Plain text - consume until next { or \ or end
      var nextSpecial = len;
      for (var j = pos; j < len; j++) {
        if (text[j] === '{' || (text[j] === '\\' && j + 1 < len && (text[j + 1] === '{' || text[j + 1] === '}'))) {
          nextSpecial = j;
          break;
        }
      }
      var plainText = text.slice(pos, nextSpecial);
      if (plainText) {
        runs.push({ text: plainText, fmt: cloneFmt(fmtStack) });
      }
      pos = nextSpecial;
    }

    return runs;
  }

  function cloneFmt(fmt) {
    return {
      bold: fmt.bold,
      italic: fmt.italic,
      underline: fmt.underline,
      sup: fmt.sup,
      sub: fmt.sub,
      strike: fmt.strike,
      dstrike: fmt.dstrike,
      smallcaps: fmt.smallcaps,
      caps: fmt.caps,
      hidden: fmt.hidden,
      size: fmt.size,
      font: fmt.font,
      color: fmt.color,
      highlight: fmt.highlight
    };
  }

  function findMatchingBrace(text, pos) {
    var i = pos + 1;
    var len = text.length;
    while (i < len) {
      if (text[i] === '\\' && i + 1 < len) { i += 2; continue; }
      if (text[i] === '"') {
        i++;
        while (i < len && text[i] !== '"') {
          if (text[i] === '\\') i++;
          i++;
        }
        i++;
        continue;
      }
      if (text[i] === '}') return i;
      i++;
    }
    return -1;
  }

  // ---------- Run to XML ----------

  function runToXml(run) {
    if (run.footnote) {
      return '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr>' +
        '<w:footnoteReference w:id="' + run.footnoteId + '"/></w:r>';
    }

    if (run.del) {
      return '<w:del w:id="' + escXml(run.revId) + '" w:author="' + escXml(run.revAuthor) + '" w:date="' + escXml(run.revDate) + '">' +
        '<w:r><w:rPr>' + buildRPr(run.fmt) + '</w:rPr><w:delText xml:space="preserve">' + escXml(run.text) + '</w:delText></w:r></w:del>';
    }

    if (run.ins) {
      return '<w:ins w:id="' + escXml(run.revId) + '" w:author="' + escXml(run.revAuthor) + '" w:date="' + escXml(run.revDate) + '">' +
        '<w:r><w:rPr>' + buildRPr(run.fmt) + '</w:rPr><w:t xml:space="preserve">' + escXml(run.text) + '</w:t></w:r></w:ins>';
    }

    if (run.commentStart) return '<w:commentRangeStart w:id="' + run.commentId + '"/>';
    if (run.commentEnd) return '<w:commentRangeEnd w:id="' + run.commentId + '"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="' + run.commentId + '"/></w:r>';
    if (run.bookmarkStart) return '<w:bookmarkStart w:id="' + run.bookmarkId + '" w:name="' + escXml(run.bookmarkName) + '"/>';
    if (run.bookmarkEnd) return '<w:bookmarkEnd w:id="' + run.bookmarkId + '"/>';
    if (run.lineBreak) return '<w:r><w:br/></w:r>';
    if (run.pageBreak) return '<w:r><w:br w:type="page"/></w:r>';

    var rPr = buildRPr(run.fmt);
    return '<w:r>' + (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') +
      '<w:t xml:space="preserve">' + escXml(run.text) + '</w:t></w:r>';
  }

  function buildRPr(fmt) {
    if (!fmt) return '';
    var parts = [];
    if (fmt.font) {
      parts.push('<w:rFonts w:ascii="' + escXml(fmt.font) + '" w:hAnsi="' + escXml(fmt.font) + '" w:cs="' + escXml(fmt.font) + '"/>');
    }
    if (fmt.bold) parts.push('<w:b/>');
    if (fmt.italic) parts.push('<w:i/>');
    if (fmt.strike) parts.push('<w:strike/>');
    if (fmt.dstrike) parts.push('<w:dstrike/>');
    if (fmt.smallcaps) parts.push('<w:smallCaps/>');
    if (fmt.caps) parts.push('<w:caps/>');
    if (fmt.hidden) parts.push('<w:vanish/>');
    if (fmt.underline) parts.push('<w:u w:val="single"/>');
    if (fmt.size) parts.push('<w:sz w:val="' + escXml(fmt.size) + '"/>');
    if (fmt.color) parts.push('<w:color w:val="' + escXml(fmt.color) + '"/>');
    if (fmt.highlight) parts.push('<w:highlight w:val="' + escXml(fmt.highlight) + '"/>');
    if (fmt.sup) parts.push('<w:vertAlign w:val="superscript"/>');
    if (fmt.sub) parts.push('<w:vertAlign w:val="subscript"/>');
    return parts.join('');
  }

  // ---------- Block-level parser ----------

  function parseDex(dexString) {
    var lines = dexString.split('\n');
    var fm = parseFrontmatter(lines);
    var meta = fm.meta;
    var blocks = [];
    var comments = [];
    var footnotes = [];

    var i = fm.bodyStart;

    while (i < lines.length) {
      var line = lines[i];

      // Skip empty lines
      if (line.trim() === '') { i++; continue; }

      // Pagebreak
      if (line.trim() === '{pagebreak}') {
        blocks.push({ type: 'pagebreak' });
        i++;
        continue;
      }

      // Heading: # text {id:XXXXX}
      var headingMatch = line.match(/^(#{1,6})\s+(.*?)(?:\s*\{id:([A-Fa-f0-9]+)\})?\s*$/);
      if (headingMatch) {
        blocks.push({
          type: 'heading',
          level: headingMatch[1].length,
          text: headingMatch[2],
          paraId: headingMatch[3] || randomParaId()
        });
        i++;
        continue;
      }

      // Comment block: {comment id:N by:"Author" date:"..."}
      var commentMatch = line.match(/^\{comment\s+id:(\d+)\s+by:"([^"]*)"\s+date:"([^"]*)"\}/);
      if (commentMatch) {
        var commentLines = [];
        i++;
        while (i < lines.length && lines[i].trim() !== '{/comment}') {
          commentLines.push(lines[i]);
          i++;
        }
        comments.push({
          id: parseInt(commentMatch[1], 10),
          author: commentMatch[2],
          date: commentMatch[3],
          text: commentLines.join('\n'),
          paraId: randomParaId(),
          blockIndex: blocks.length - 1  // attach to the most recent block
        });
        if (i < lines.length) i++;
        continue;
      }

      // Reply block: {reply id:N parent:M by:"Author" date:"..."}
      var replyMatch = line.match(/^\{reply\s+id:(\d+)\s+parent:(\d+)\s+by:"([^"]*)"\s+date:"([^"]*)"\}/);
      if (replyMatch) {
        var replyLines = [];
        i++;
        while (i < lines.length && lines[i].trim() !== '{/reply}') {
          replyLines.push(lines[i]);
          i++;
        }
        comments.push({
          id: parseInt(replyMatch[1], 10),
          parent: parseInt(replyMatch[2], 10),
          author: replyMatch[3],
          date: replyMatch[4],
          text: replyLines.join('\n'),
          paraId: randomParaId(),
          blockIndex: blocks.length - 1
        });
        if (i < lines.length) i++;
        continue;
      }

      // Table block: {table ...}
      var tableMatch = line.match(/^\{table\b/);
      if (tableMatch) {
        var tableRows = [];
        i++;
        while (i < lines.length && lines[i].trim() !== '{/table}') {
          var rowLine = lines[i].trim();
          if (rowLine.match(/^\|[\s-|]+\|$/)) { i++; continue; }
          if (rowLine.startsWith('|')) {
            var cells = rowLine.split('|').filter(function(c) { return c.trim() !== ''; }).map(function(c) { return c.trim(); });
            tableRows.push(cells);
          }
          i++;
        }
        if (i < lines.length) i++;
        blocks.push({ type: 'table', rows: tableRows });
        continue;
      }

      // Paragraph: {p id:XXXX}...{/p}
      var pMatch = line.match(/^\{p(?:\s+id:([A-Fa-f0-9]+))?\}/);
      if (pMatch) {
        var paraId = pMatch[1] || randomParaId();
        var contentLines = [];
        i++;
        while (i < lines.length && lines[i].trim() !== '{/p}') {
          contentLines.push(lines[i]);
          i++;
        }
        if (i < lines.length) i++;
        var content = contentLines.join('\n');

        // Collect footnotes from inline content
        var fnRe = /\{footnote\s+id:(\d+)\}([\s\S]*?)\{\/footnote\}/g;
        var fnM;
        while ((fnM = fnRe.exec(content)) !== null) {
          footnotes.push({ id: parseInt(fnM[1], 10), text: unescapeBraces(fnM[2]) });
        }

        blocks.push({
          type: 'paragraph',
          paraId: paraId,
          content: content
        });
        continue;
      }

      // Plain line (treated as a paragraph)
      var plainContent = line;

      // Collect footnotes from inline content
      var fnRe2 = /\{footnote\s+id:(\d+)\}([\s\S]*?)\{\/footnote\}/g;
      var fnM2;
      while ((fnM2 = fnRe2.exec(plainContent)) !== null) {
        footnotes.push({ id: parseInt(fnM2[1], 10), text: unescapeBraces(fnM2[2]) });
      }

      blocks.push({
        type: 'paragraph',
        paraId: randomParaId(),
        content: plainContent
      });

      i++;
    }

    return { meta: meta, blocks: blocks, comments: comments, footnotes: footnotes };
  }

  // ---------- Document XML generation ----------

  function buildDocumentXml(parsed) {
    var blocks = parsed.blocks;
    var comments = parsed.comments;

    var bodyParts = [];

    // Build map: blockIndex -> [commentId, ...] for non-reply comments
    var commentBlockMap = {};
    for (var ci = 0; ci < comments.length; ci++) {
      var c = comments[ci];
      if (c.parent !== undefined) continue;
      var bIdx = c.blockIndex >= 0 ? c.blockIndex : 0;
      if (!commentBlockMap[bIdx]) commentBlockMap[bIdx] = [];
      commentBlockMap[bIdx].push(c.id);
    }

    for (var b = 0; b < blocks.length; b++) {
      var block = blocks[b];

      if (block.type === 'pagebreak') {
        bodyParts.push('<w:p w14:paraId="' + randomParaId() + '" w14:textId="' + randomParaId() + '">' +
          '<w:r><w:br w:type="page"/></w:r></w:p>');
        continue;
      }

      if (block.type === 'heading') {
        var headingXml = '<w:p w14:paraId="' + block.paraId + '" w14:textId="' + randomParaId() + '">' +
          '<w:pPr><w:pStyle w:val="Heading' + block.level + '"/></w:pPr>';

        var headingComments = commentBlockMap[b] || [];
        for (var hci = 0; hci < headingComments.length; hci++) {
          headingXml += '<w:commentRangeStart w:id="' + headingComments[hci] + '"/>';
        }

        headingXml += '<w:r><w:t xml:space="preserve">' + escXml(unescapeBraces(block.text)) + '</w:t></w:r>';

        for (var hci2 = 0; hci2 < headingComments.length; hci2++) {
          headingXml += '<w:commentRangeEnd w:id="' + headingComments[hci2] + '"/>';
          headingXml += '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' +
            '<w:commentReference w:id="' + headingComments[hci2] + '"/></w:r>';
        }

        headingXml += '</w:p>';
        bodyParts.push(headingXml);
        continue;
      }

      if (block.type === 'table') {
        bodyParts.push(buildTableXml(block));
        continue;
      }

      if (block.type === 'paragraph') {
        var paraXml = '<w:p w14:paraId="' + block.paraId + '" w14:textId="' + randomParaId() + '">';

        var associatedComments = commentBlockMap[b] || [];
        for (var aci = 0; aci < associatedComments.length; aci++) {
          paraXml += '<w:commentRangeStart w:id="' + associatedComments[aci] + '"/>';
        }

        var runs = parseInlineRuns(block.content, null);
        for (var r = 0; r < runs.length; r++) {
          paraXml += runToXml(runs[r]);
        }

        for (var aci2 = 0; aci2 < associatedComments.length; aci2++) {
          paraXml += '<w:commentRangeEnd w:id="' + associatedComments[aci2] + '"/>';
          paraXml += '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' +
            '<w:commentReference w:id="' + associatedComments[aci2] + '"/></w:r>';
        }

        paraXml += '</w:p>';
        bodyParts.push(paraXml);
        continue;
      }
    }

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
      'xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
      'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
      'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
      'mc:Ignorable="w14 w15 wp14">' +
      '<w:body>' + bodyParts.join('') +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '</w:sectPr>' +
      '</w:body></w:document>';
  }

  function buildTableXml(block) {
    var rows = block.rows;
    if (rows.length === 0) return '';

    var colCount = rows[0].length;
    var xml = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>' +
      '<w:tblBorders>' +
      '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '</w:tblBorders></w:tblPr>' +
      '<w:tblGrid>';

    for (var c = 0; c < colCount; c++) {
      xml += '<w:gridCol w:w="' + Math.floor(9360 / colCount) + '"/>';
    }
    xml += '</w:tblGrid>';

    for (var r = 0; r < rows.length; r++) {
      xml += '<w:tr>';
      for (var cc = 0; cc < rows[r].length; cc++) {
        xml += '<w:tc><w:p w14:paraId="' + randomParaId() + '" w14:textId="' + randomParaId() + '">' +
          '<w:r><w:t xml:space="preserve">' + escXml(rows[r][cc]) + '</w:t></w:r></w:p></w:tc>';
      }
      xml += '</w:tr>';
    }

    xml += '</w:tbl>';
    return xml;
  }

  // ---------- Comments XML ----------

  function buildCommentsXml(comments) {
    if (comments.length === 0) return null;

    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
      'xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
      'mc:Ignorable="w14 w15">';

    for (var i = 0; i < comments.length; i++) {
      var c = comments[i];
      var initials = c.author ? c.author.split(/\s+/).map(function(w) { return w.charAt(0); }).join('') : '';
      xml += '<w:comment w:id="' + c.id + '" w:author="' + escXml(c.author) + '" w:date="' + escXml(c.date) + '" w:initials="' + escXml(initials) + '">' +
        '<w:p w14:paraId="' + c.paraId + '" w14:textId="' + randomParaId() + '">' +
        '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>' +
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' +
        '<w:r><w:t xml:space="preserve">' + escXml(c.text) + '</w:t></w:r></w:p></w:comment>';
    }

    xml += '</w:comments>';
    return xml;
  }

  // ---------- Comments Extended XML ----------

  function buildCommentsExtendedXml(comments) {
    var hasReplies = comments.some(function(c) { return c.parent !== undefined; });
    if (!hasReplies) return null;

    var paraIdMap = {};
    for (var i = 0; i < comments.length; i++) {
      paraIdMap[comments[i].id] = comments[i].paraId;
    }

    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'mc:Ignorable="w15">';

    for (var j = 0; j < comments.length; j++) {
      var c = comments[j];
      if (c.parent !== undefined) {
        var parentParaId = paraIdMap[c.parent] || '';
        xml += '<w15:commentEx w15:paraId="' + c.paraId + '" w15:paraIdParent="' + parentParaId + '" w15:done="0"/>';
      } else {
        xml += '<w15:commentEx w15:paraId="' + c.paraId + '" w15:done="0"/>';
      }
    }

    xml += '</w15:commentsEx>';
    return xml;
  }

  // ---------- Footnotes XML ----------

  function buildFootnotesXml(footnotes) {
    if (footnotes.length === 0) return null;

    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:footnotes xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
      'xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
      'mc:Ignorable="w14 w15">';

    // Separator and continuation separator (required by Word)
    xml += '<w:footnote w:type="separator" w:id="-1"><w:p w14:paraId="' + randomParaId() + '" w14:textId="' + randomParaId() + '">' +
      '<w:r><w:separator/></w:r></w:p></w:footnote>';
    xml += '<w:footnote w:type="continuationSeparator" w:id="0"><w:p w14:paraId="' + randomParaId() + '" w14:textId="' + randomParaId() + '">' +
      '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';

    for (var i = 0; i < footnotes.length; i++) {
      var fn = footnotes[i];
      xml += '<w:footnote w:id="' + fn.id + '"><w:p w14:paraId="' + randomParaId() + '" w14:textId="' + randomParaId() + '">' +
        '<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' +
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' +
        '<w:r><w:t xml:space="preserve"> ' + escXml(fn.text) + '</w:t></w:r></w:p></w:footnote>';
    }

    xml += '</w:footnotes>';
    return xml;
  }

  // ---------- Styles XML ----------

  function buildStylesXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
      'mc:Ignorable="w14 w15">' +

      '<w:docDefaults>' +
      '<w:rPrDefault><w:rPr>' +
      '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>' +
      '<w:sz w:val="24"/><w:szCs w:val="24"/>' +
      '</w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr>' +
      '<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>' +
      '</w:pPr></w:pPrDefault>' +
      '</w:docDefaults>' +

      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
      '<w:name w:val="Normal"/><w:qFormat/></w:style>' +

      '<w:style w:type="paragraph" w:styleId="Heading1">' +
      '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr>' +
      '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' +
      '<w:b/><w:bCs/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>' +
      '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +

      '<w:style w:type="paragraph" w:styleId="Heading2">' +
      '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr>' +
      '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' +
      '<w:b/><w:bCs/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>' +
      '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>' +

      '<w:style w:type="paragraph" w:styleId="Heading3">' +
      '<w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr>' +
      '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' +
      '<w:b/><w:bCs/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>' +
      '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' +

      '<w:style w:type="paragraph" w:styleId="Heading4">' +
      '<w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="3"/></w:pPr>' +
      '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' +
      '<w:b/><w:bCs/><w:i/><w:iCs/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/></w:rPr></w:style>' +

      '<w:style w:type="paragraph" w:styleId="Heading5">' +
      '<w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="4"/></w:pPr>' +
      '<w:rPr><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/></w:rPr></w:style>' +

      '<w:style w:type="paragraph" w:styleId="Heading6">' +
      '<w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="5"/></w:pPr>' +
      '<w:rPr><w:i/><w:iCs/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/></w:rPr></w:style>' +

      '<w:style w:type="paragraph" w:styleId="CommentText">' +
      '<w:name w:val="annotation text"/><w:basedOn w:val="Normal"/>' +
      '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +

      '<w:style w:type="character" w:styleId="CommentReference">' +
      '<w:name w:val="annotation reference"/>' +
      '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>' +

      '<w:style w:type="paragraph" w:styleId="FootnoteText">' +
      '<w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
      '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +

      '<w:style w:type="character" w:styleId="FootnoteReference">' +
      '<w:name w:val="footnote reference"/>' +
      '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +

      '<w:style w:type="table" w:styleId="TableGrid">' +
      '<w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>' +
      '<w:tblPr><w:tblBorders>' +
      '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '</w:tblBorders></w:tblPr></w:style>' +

      '<w:style w:type="table" w:default="1" w:styleId="TableNormal">' +
      '<w:name w:val="Normal Table"/>' +
      '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar>' +
      '<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
      '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>' +
      '</w:tblCellMar></w:tblPr></w:style>' +

      '</w:styles>';
  }

  // ---------- Package parts ----------

  function buildContentTypes(hasComments, hasFootnotes, hasCommentsExt) {
    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';

    if (hasComments) {
      xml += '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>';
    }
    if (hasCommentsExt) {
      xml += '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>';
    }
    if (hasFootnotes) {
      xml += '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>';
    }

    xml += '</Types>';
    return xml;
  }

  function buildRootRels() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>';
  }

  function buildDocumentRels(hasComments, hasFootnotes, hasCommentsExt) {
    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';

    var nextId = 2;
    if (hasComments) {
      xml += '<Relationship Id="rId' + nextId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>';
      nextId++;
    }
    if (hasCommentsExt) {
      xml += '<Relationship Id="rId' + nextId + '" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>';
      nextId++;
    }
    if (hasFootnotes) {
      xml += '<Relationship Id="rId' + nextId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>';
      nextId++;
    }

    xml += '</Relationships>';
    return xml;
  }

  // ---------- Main compile function ----------

  function compile(dexString) {
    return new Promise(function(resolve, reject) {
      try {
        var parsed = parseDex(dexString);
        var documentXml = buildDocumentXml(parsed);
        var stylesXml = buildStylesXml();

        var hasComments = parsed.comments.length > 0;
        var commentsXml = hasComments ? buildCommentsXml(parsed.comments) : null;
        var commentsExtXml = hasComments ? buildCommentsExtendedXml(parsed.comments) : null;
        var hasCommentsExt = commentsExtXml !== null;

        var hasFootnotes = parsed.footnotes.length > 0;
        var footnotesXml = hasFootnotes ? buildFootnotesXml(parsed.footnotes) : null;

        var contentTypesXml = buildContentTypes(hasComments, hasFootnotes, hasCommentsExt);
        var rootRels = buildRootRels();
        var docRels = buildDocumentRels(hasComments, hasFootnotes, hasCommentsExt);

        var zip = new JSZip();
        zip.file('[Content_Types].xml', contentTypesXml);
        zip.file('_rels/.rels', rootRels);
        zip.file('word/document.xml', documentXml);
        zip.file('word/styles.xml', stylesXml);
        zip.file('word/_rels/document.xml.rels', docRels);

        if (commentsXml) zip.file('word/comments.xml', commentsXml);
        if (commentsExtXml) zip.file('word/commentsExtended.xml', commentsExtXml);
        if (footnotesXml) zip.file('word/footnotes.xml', footnotesXml);

        zip.generateAsync({ type: 'arraybuffer' }).then(resolve).catch(reject);
      } catch (err) {
        reject(err);
      }
    });
  }

  return {
    compile: compile
  };

})();
