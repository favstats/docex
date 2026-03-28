/* ============================================================================
   decompiler.js -- Browser-compatible .docx to .dex decompiler
   Uses JSZip (loaded from CDN) to unzip .docx files in the browser.
   ============================================================================ */

const DocxDecompiler = {

  // ---------- Main entry point ----------
  async decompile(file) {
    const zip = await JSZip.loadAsync(file);
    const docXml = await this._readPart(zip, 'word/document.xml');
    if (!docXml) throw new Error('No word/document.xml found in .docx');

    const commentsXml = await this._readPart(zip, 'word/comments.xml');
    const footnotesXml = await this._readPart(zip, 'word/footnotes.xml');
    const stylesXml = await this._readPart(zip, 'word/styles.xml');
    const commentsExtXml = await this._readPart(zip, 'word/commentsExtended.xml') ||
                           await this._readPart(zip, 'word/commentsExtensible.xml');
    const relsXml = await this._readPart(zip, 'word/_rels/document.xml.rels');

    const parts = [];
    parts.push(this._buildFrontmatter(docXml));

    const paragraphs = this._findParagraphs(docXml);
    const commentMap = this._buildCommentMap(commentsXml, commentsExtXml);
    const footnoteMap = this._buildFootnoteMap(footnotesXml);
    const commentRanges = this._buildCommentRanges(docXml, commentMap, paragraphs);
    const tablePositions = this._findTables(docXml);
    let nextTableIdx = 0;

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const pXml = p.xml;
      const paraId = this._extractParaId(pXml);
      const level = this._headingLevel(pXml);
      const hasFigure = /<w:drawing[\s>]/.test(pXml) || /<w:pict[\s>]/.test(pXml);

      // Emit tables that appear before this paragraph
      while (nextTableIdx < tablePositions.length && tablePositions[nextTableIdx].start < p.start) {
        parts.push(this._decompileTable(tablePositions[nextTableIdx]));
        nextTableIdx++;
      }

      if (level > 0) {
        const rawText = this._extractTextDecoded(pXml);
        const text = rawText.replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}');
        const hashes = '#'.repeat(level);
        const idAttr = paraId ? ' {id:' + paraId + '}' : '';
        parts.push(hashes + ' ' + text + idAttr);
        parts.push('');
      } else if (hasFigure) {
        parts.push(this._decompileFigure(pXml, paraId, relsXml));
        parts.push('');
      } else {
        const content = this._decompileRuns(pXml, footnoteMap);
        if (content.trim() === '') {
          if (/<w:br\s+w:type="page"/.test(pXml)) {
            parts.push('{pagebreak}');
            parts.push('');
          }
          continue;
        }
        if (/<w:br\s+w:type="page"/.test(pXml)) {
          parts.push('{pagebreak}');
        }
        const idAttr = paraId ? ' id:' + paraId : '';
        parts.push('{p' + idAttr + '}');
        parts.push(content);
        parts.push('{/p}');
        parts.push('');
      }

      // Emit comments that end at this paragraph
      const commentsForPara = commentRanges.filter(cr => cr.endParaIndex === i);
      for (const cr of commentsForPara) {
        const comment = commentMap.get(cr.commentId);
        if (!comment) continue;
        parts.push(this._formatComment(comment));
        if (comment.replies) {
          for (const reply of comment.replies) {
            parts.push(this._formatReply(reply, cr.commentId));
          }
        }
        parts.push('');
      }
    }

    // Remaining tables
    while (nextTableIdx < tablePositions.length) {
      parts.push(this._decompileTable(tablePositions[nextTableIdx]));
      nextTableIdx++;
    }

    return parts.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
  },

  // ---------- ZIP helpers ----------
  async _readPart(zip, path) {
    const entry = zip.file(path);
    if (!entry) return null;
    return entry.async('string');
  },

  // ---------- XML helpers ----------
  _decodeXml(str) {
    return str
      .replace(/&apos;/g, "'")
      .replace(/&quot;/g, '"')
      .replace(/&gt;/g, '>')
      .replace(/&lt;/g, '<')
      .replace(/&amp;/g, '&');
  },

  _extractText(pXml) {
    const texts = [];
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let m;
    while ((m = re.exec(pXml)) !== null) texts.push(m[1]);
    return texts.join('');
  },

  _extractTextDecoded(pXml) {
    return this._decodeXml(this._extractText(pXml));
  },

  _attrVal(attrs, name) {
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(escaped + '="([^"]*)"');
    const m = attrs.match(re);
    return m ? m[1] : null;
  },

  // ---------- Find paragraphs ----------
  _findParagraphs(docXml) {
    const paragraphs = [];
    const re = /<w:p[\s>]/g;
    let m;
    while ((m = re.exec(docXml)) !== null) {
      const start = m.index;
      const closeTag = '</w:p>';
      const closeIdx = docXml.indexOf(closeTag, start);
      if (closeIdx === -1) continue;
      const end = closeIdx + closeTag.length;
      const xml = docXml.slice(start, end);
      paragraphs.push({ xml, start, end, text: this._extractText(xml) });
    }
    return paragraphs;
  },

  // ---------- Heading detection ----------
  _headingLevel(pXml) {
    const olvl = pXml.match(/<w:outlineLvl\s+w:val="(\d+)"/);
    if (olvl) return parseInt(olvl[1], 10) + 1;
    const styleMatch = pXml.match(/<w:pStyle\s+w:val="([^"]+)"/);
    if (!styleMatch) return 0;
    const styleId = styleMatch[1];
    const named = styleId.match(/^[Hh]eading(\d+)$/);
    if (named) return parseInt(named[1], 10);
    if (/^Title$/i.test(styleId)) return 1;
    if (/^Subtitle$/i.test(styleId)) return 2;
    return 0;
  },

  // ---------- Frontmatter ----------
  _buildFrontmatter(docXml) {
    const lines = ['---'];
    lines.push('docex: "0.4.0"');
    lines.push('---');
    lines.push('');
    return lines.join('\n');
  },

  // ---------- Footnotes ----------
  _buildFootnoteMap(footnotesXml) {
    const map = new Map();
    if (!footnotesXml) return map;
    const fnRe = /<w:footnote\b([^>]*)>([\s\S]*?)<\/w:footnote>/g;
    let m;
    while ((m = fnRe.exec(footnotesXml)) !== null) {
      const attrs = m[1]; const body = m[2];
      const id = this._attrVal(attrs, 'w:id');
      const type = this._attrVal(attrs, 'w:type');
      if (type) continue;
      if (!id) continue;
      const idNum = parseInt(id, 10);
      if (idNum <= 1) continue;
      map.set(idNum, this._extractTextDecoded(body));
    }
    return map;
  },

  // ---------- Comments ----------
  _buildCommentMap(commentsXml, commentsExtXml) {
    const map = new Map();
    if (!commentsXml) return map;
    const commentRe = /<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g;
    let m;
    while ((m = commentRe.exec(commentsXml)) !== null) {
      const attrs = m[1]; const body = m[2];
      const id = this._attrVal(attrs, 'w:id');
      const author = this._decodeXml(this._attrVal(attrs, 'w:author') || '');
      const date = this._attrVal(attrs, 'w:date') || '';
      const text = this._extractTextDecoded(body);
      const innerParaIdMatch = body.match(/w14:paraId="([^"]+)"/);
      const innerParaId = innerParaIdMatch ? innerParaIdMatch[1] : '';
      if (id) map.set(parseInt(id, 10), { id: parseInt(id, 10), author, date, text, paraId: innerParaId, replies: [] });
    }

    if (commentsExtXml) {
      const exRe = /<w15:commentEx\s+([^>]*?)\s*\/?>/g;
      let exM;
      const entries = [];
      while ((exM = exRe.exec(commentsExtXml)) !== null) entries.push(exM[1]);
      const paraIdToComment = new Map();
      for (const [cId, cData] of map) { if (cData.paraId) paraIdToComment.set(cData.paraId, cId); }
      for (const entryAttrs of entries) {
        const entryParaId = this._attrVal(entryAttrs, 'w15:paraId');
        const parentParaId = this._attrVal(entryAttrs, 'w15:paraIdParent');
        if (entryParaId && parentParaId) {
          const childCommentId = paraIdToComment.get(entryParaId);
          const parentCommentId = paraIdToComment.get(parentParaId);
          if (childCommentId !== undefined && parentCommentId !== undefined && map.has(parentCommentId) && map.has(childCommentId)) {
            map.get(parentCommentId).replies.push(map.get(childCommentId));
            map.get(childCommentId).isReply = true;
          }
        }
      }
    }
    return map;
  },

  _buildCommentRanges(docXml, commentMap, paragraphs) {
    const ranges = [];
    const rangeStarts = new Map();
    const rangeEnds = new Map();
    for (let i = 0; i < paragraphs.length; i++) {
      const pXml = paragraphs[i].xml;
      const startRe = /<w:commentRangeStart\s+w:id="(\d+)"\s*\/?>/g;
      let sm;
      while ((sm = startRe.exec(pXml)) !== null) rangeStarts.set(parseInt(sm[1], 10), i);
      const endRe = /<w:commentRangeEnd\s+w:id="(\d+)"\s*\/?>/g;
      let em;
      while ((em = endRe.exec(pXml)) !== null) rangeEnds.set(parseInt(em[1], 10), i);
      const refRe = /<w:commentReference\s+w:id="(\d+)"\s*\/?>/g;
      let rm;
      while ((rm = refRe.exec(pXml)) !== null) {
        const cid = parseInt(rm[1], 10);
        if (!rangeEnds.has(cid)) rangeEnds.set(cid, i);
        if (!rangeStarts.has(cid)) rangeStarts.set(cid, i);
      }
    }
    for (const [commentId, comment] of commentMap) {
      if (comment.isReply) continue;
      const endIdx = rangeEnds.get(commentId);
      if (endIdx !== undefined) {
        const startIdx = rangeStarts.get(commentId);
        ranges.push({ commentId, startParaIndex: startIdx !== undefined ? startIdx : endIdx, endParaIndex: endIdx });
      }
    }
    return ranges;
  },

  _formatComment(comment) {
    return '{comment id:' + comment.id + ' by:' + this._dexStr(comment.author) + ' date:' + this._dexStr(comment.date) + '}\n' + comment.text + '\n{/comment}';
  },

  _formatReply(reply, parentId) {
    return '{reply id:' + reply.id + ' parent:' + parentId + ' by:' + this._dexStr(reply.author) + ' date:' + this._dexStr(reply.date) + '}\n' + reply.text + '\n{/reply}';
  },

  // ---------- Runs ----------
  _decompileRuns(pXml, footnoteMap) {
    const parts = [];
    let bodyXml = pXml;
    const pOpenEnd = bodyXml.indexOf('>');
    bodyXml = bodyXml.slice(pOpenEnd + 1);
    if (bodyXml.endsWith('</w:p>')) bodyXml = bodyXml.slice(0, -6);
    const pPrMatch = bodyXml.match(/^(\s*<w:pPr>[\s\S]*?<\/w:pPr>)/);
    if (pPrMatch) bodyXml = bodyXml.slice(pPrMatch[0].length);
    this._walkElements(bodyXml, parts, footnoteMap);
    return parts.join('');
  },

  _walkElements(xmlStr, parts, footnoteMap) {
    let pos = 0;
    const len = xmlStr.length;
    while (pos < len) {
      if (xmlStr[pos] !== '<') { pos++; continue; }

      if (xmlStr.startsWith('<w:ins', pos)) {
        const endTag = '</w:ins>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const insXml = xmlStr.slice(pos, endIdx + endTag.length);
        const attrs = this._extractAttrs(insXml);
        const insContent = this._extractRunTexts(insXml, 'w:t');
        if (insContent) {
          const id = attrs['w:id'] || '';
          const author = this._decodeXml(attrs['w:author'] || '');
          const date = attrs['w:date'] || '';
          parts.push('{ins id:' + id + ' by:' + this._dexStr(author) + ' date:' + this._dexStr(date) + '}' + insContent + '{/ins}');
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:del', pos)) {
        const endTag = '</w:del>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const delXml = xmlStr.slice(pos, endIdx + endTag.length);
        const attrs = this._extractAttrs(delXml);
        const delContent = this._extractRunTexts(delXml, 'w:delText');
        if (delContent) {
          const id = attrs['w:id'] || '';
          const author = this._decodeXml(attrs['w:author'] || '');
          const date = attrs['w:date'] || '';
          parts.push('{del id:' + id + ' by:' + this._dexStr(author) + ' date:' + this._dexStr(date) + '}' + delContent + '{/del}');
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:r>', pos) || xmlStr.startsWith('<w:r ', pos)) {
        const endTag = '</w:r>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const runXml = xmlStr.slice(pos, endIdx + endTag.length);

        // Footnote reference
        const fnRefMatch = runXml.match(/<w:footnoteReference\s+w:id="(\d+)"/);
        if (fnRefMatch) {
          const fnId = parseInt(fnRefMatch[1], 10);
          const fnText = footnoteMap.get(fnId) || '';
          parts.push('{footnote id:' + fnId + '}' + fnText + '{/footnote}');
          pos = endIdx + endTag.length;
          continue;
        }

        // Skip comment references
        if (/<w:commentReference/.test(runXml)) { pos = endIdx + endTag.length; continue; }

        // Text
        const texts = [];
        const tRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
        let tMatch;
        while ((tMatch = tRe.exec(runXml)) !== null) texts.push(this._decodeXml(tMatch[1]));
        const text = texts.join('');
        if (text) {
          const escaped = text.replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}');
          const fmt = this._extractFormatting(runXml);
          parts.push(this._wrapFormatting(escaped, fmt));
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:hyperlink', pos)) {
        const endTag = '</w:hyperlink>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const hlXml = xmlStr.slice(pos, endIdx + endTag.length);
        this._walkElements(hlXml.slice(hlXml.indexOf('>') + 1, hlXml.lastIndexOf('<')), parts, footnoteMap);
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:commentRangeStart', pos) || xmlStr.startsWith('<w:commentRangeEnd', pos) ||
                 xmlStr.startsWith('<w:bookmarkStart', pos) || xmlStr.startsWith('<w:bookmarkEnd', pos)) {
        const closeAngle = xmlStr.indexOf('>', pos);
        pos = closeAngle + 1;
      } else {
        const closeAngle = xmlStr.indexOf('>', pos);
        if (closeAngle === -1) break;
        pos = closeAngle + 1;
      }
    }
  },

  _extractRunTexts(elXml, textTag) {
    const parts = [];
    const tagRe = new RegExp('<' + textTag + '[^>]*>([^<]*)</' + textTag + '>', 'g');
    let m;
    while ((m = tagRe.exec(elXml)) !== null) {
      let text = this._decodeXml(m[1]);
      text = text.replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}');
      parts.push(text);
    }
    return parts.join('');
  },

  _extractFormatting(runXml) {
    const rPrMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    if (!rPrMatch) return {};
    const rPr = rPrMatch[1];
    const fmt = {};
    if (/<w:b\b[^>]*\/?>/.test(rPr) && !/<w:b\s+w:val="(false|0)"/.test(rPr)) fmt.bold = true;
    if (/<w:i\b[^>]*\/?>/.test(rPr) && !/<w:i\s+w:val="(false|0)"/.test(rPr)) fmt.italic = true;
    if (/<w:u\b/.test(rPr) && !/<w:u\s+w:val="none"/.test(rPr)) fmt.underline = true;
    if (/<w:vertAlign\s+w:val="superscript"/.test(rPr)) fmt.sup = true;
    if (/<w:vertAlign\s+w:val="subscript"/.test(rPr)) fmt.sub = true;
    const hlMatch = rPr.match(/<w:highlight\s+w:val="([^"]+)"/);
    if (hlMatch) fmt.highlight = hlMatch[1];
    const colorMatch = rPr.match(/<w:color\s+w:val="([^"]+)"/);
    if (colorMatch && colorMatch[1] !== '000000' && colorMatch[1] !== 'auto') fmt.color = colorMatch[1];
    const fontMatch = rPr.match(/w:ascii="([^"]+)"/);
    if (fontMatch) fmt.font = fontMatch[1];
    return fmt;
  },

  _wrapFormatting(text, fmt) {
    if (!fmt || Object.keys(fmt).length === 0) return text;
    let result = text;
    if (fmt.bold) result = '{b}' + result + '{/b}';
    if (fmt.italic) result = '{i}' + result + '{/i}';
    if (fmt.underline) result = '{u}' + result + '{/u}';
    if (fmt.sup) result = '{sup}' + result + '{/sup}';
    if (fmt.sub) result = '{sub}' + result + '{/sub}';
    if (fmt.highlight) result = '{highlight ' + fmt.highlight + '}' + result + '{/highlight}';
    if (fmt.color) result = '{color ' + fmt.color + '}' + result + '{/color}';
    if (fmt.font) result = '{font "' + fmt.font + '"}' + result + '{/font}';
    return result;
  },

  // ---------- Figures ----------
  _decompileFigure(pXml, paraId, relsXml) {
    const drawingMatch = pXml.match(/<w:drawing[\s>][\s\S]*?<\/w:drawing>/);
    if (!drawingMatch) {
      const text = this._extractTextDecoded(pXml);
      return '{p id:' + paraId + '}\n' + text + '\n{/p}';
    }
    const drawXml = drawingMatch[0];
    const cxMatch = drawXml.match(/\bcx="(\d+)"/);
    const cyMatch = drawXml.match(/\bcy="(\d+)"/);
    const width = cxMatch ? cxMatch[1] + 'emu' : '';
    const height = cyMatch ? cyMatch[1] + 'emu' : '';
    const rIdMatch = drawXml.match(/r:embed="([^"]+)"/);
    const rId = rIdMatch ? rIdMatch[1] : '';
    let src = '';
    if (rId && relsXml) {
      const relRe = new RegExp('Id="' + rId + '"[^>]*Target="([^"]+)"');
      const relMatch = relRe.exec(relsXml);
      if (relMatch) src = relMatch[1].startsWith('media/') ? 'word/' + relMatch[1] : 'word/' + relMatch[1].replace(/^\.\.\//, '');
    }
    const altMatch = drawXml.match(/descr="([^"]*)"/);
    const alt = altMatch ? this._decodeXml(altMatch[1]) : '';
    const caption = this._extractTextDecoded(pXml);
    const figureParts = [];
    let attrs = '{figure';
    if (paraId) attrs += ' id:' + paraId;
    if (rId) attrs += ' rId:' + rId;
    if (src) attrs += ' src:' + this._dexStr(src);
    if (width) attrs += ' width:' + width;
    if (height) attrs += ' height:' + height;
    if (alt) attrs += ' alt:' + this._dexStr(alt);
    attrs += '}';
    figureParts.push(attrs);
    if (caption) figureParts.push(caption);
    figureParts.push('{/figure}');
    return figureParts.join('\n');
  },

  // ---------- Tables ----------
  _findTables(docXml) {
    const tables = [];
    const tblRe = /<w:tbl[\s>]/g;
    let m;
    while ((m = tblRe.exec(docXml)) !== null) {
      const start = m.index;
      const end = docXml.indexOf('</w:tbl>', start);
      if (end === -1) continue;
      tables.push({ start, end: end + 8, xml: docXml.slice(start, end + 8) });
    }
    return tables;
  },

  _decompileTable(tblInfo) {
    const tblXml = tblInfo.xml;
    const tableParts = [];
    const firstRow = tblXml.match(/<w:tr[\s>][\s\S]*?<\/w:tr>/);
    let colCount = 0;
    if (firstRow) colCount = (firstRow[0].match(/<w:tc[\s>]/g) || []).length;
    const styleMatch = tblXml.match(/<w:tblStyle\s+w:val="([^"]+)"/);
    const style = styleMatch ? styleMatch[1] : 'plain';
    tableParts.push('{table style:' + style + ' cols:' + colCount + '}');
    const rowRe = /<w:tr[\s>][\s\S]*?<\/w:tr>/g;
    let rm;
    const rows = [];
    while ((rm = rowRe.exec(tblXml)) !== null) {
      const rowXml = rm[0]; const cells = [];
      const cellRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
      let cm;
      while ((cm = cellRe.exec(rowXml)) !== null) cells.push(this._extractTextDecoded(cm[0]).trim());
      rows.push(cells);
    }
    if (rows.length > 0) {
      tableParts.push('| ' + rows[0].join(' | ') + ' |');
      tableParts.push('|' + rows[0].map(() => '---').join('|') + '|');
      for (let r = 1; r < rows.length; r++) tableParts.push('| ' + rows[r].join(' | ') + ' |');
    }
    tableParts.push('{/table}');
    return tableParts.join('\n');
  },

  // ---------- Helpers ----------
  _extractParaId(pXml) {
    const m = pXml.match(/w14:paraId="([^"]+)"/);
    return m ? m[1] : '';
  },

  _extractAttrs(elXml) {
    const openTag = elXml.match(/^<[^>]+>/);
    if (!openTag) return {};
    const attrs = {};
    const attrRe = /(\w[\w:]*?)="([^"]*)"/g;
    let m;
    while ((m = attrRe.exec(openTag[0])) !== null) attrs[m[1]] = m[2];
    return attrs;
  },

  _dexStr(s) {
    if (!s) return '""';
    return '"' + s.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
  }
};
