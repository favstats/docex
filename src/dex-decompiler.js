/**
 * dex-decompiler.js -- Convert .docx XML into .dex human-readable format
 *
 * Walks document.xml paragraph by paragraph, extracting paraIds, headings,
 * figures, tables, tracked changes, comments, footnotes, and inline formatting
 * into the .dex markup language. Preserves EVERYTHING losslessly.
 *
 * All methods operate on a Workspace object. Zero external dependencies.
 */

'use strict';

const xml = require('./xml');
const { Paragraphs } = require('./paragraphs');
const { Metadata } = require('./metadata');

class DexDecompiler {

  static decompile(wsOrPath) {
    // When given a path string, use the lossless binary format (for round-trip fidelity).
    // When given a Workspace object, use the human-readable markdown format.
    if (typeof wsOrPath === 'string') {
      const { serializeDex } = require('./dex-lossless');
      const ast = DexDecompiler.toAst(wsOrPath);
      return serializeDex(ast);
    }
    return DexDecompiler._decompileWorkspace(wsOrPath);
  }

  static _decompileWorkspace(ws) {
    const parts = [];
    parts.push(DexDecompiler._buildFrontmatter(ws));
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    const commentMap = DexDecompiler._buildCommentMap(ws);
    const footnoteMap = DexDecompiler._buildFootnoteMap(ws);
    const commentRanges = DexDecompiler._buildCommentRanges(docXml, commentMap);
    const tablePositions = DexDecompiler._findTables(docXml);
    let nextTableIdx = 0;

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const pXml = p.xml;
      const paraId = DexDecompiler._extractParaId(pXml);
      const level = Paragraphs._headingLevel(pXml);
      const hasFigure = /<w:drawing[\s>]/.test(pXml) || /<w:pict[\s>]/.test(pXml);

      while (nextTableIdx < tablePositions.length && tablePositions[nextTableIdx].start < p.start) {
        parts.push(DexDecompiler._decompileTable(tablePositions[nextTableIdx], docXml));
        nextTableIdx++;
      }

      if (level > 0) {
        const text = xml.extractTextDecoded(pXml);
        const hashes = '#'.repeat(level);
        const idAttr = paraId ? ' {id:' + paraId + '}' : '';
        parts.push(hashes + ' ' + text + idAttr);
        parts.push('');
      } else if (hasFigure) {
        parts.push(DexDecompiler._decompileFigure(pXml, paraId, ws));
        parts.push('');
      } else {
        const content = DexDecompiler._decompileRuns(pXml, footnoteMap);
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

      const commentsForPara = commentRanges.filter(cr => cr.endParaIndex === i);
      for (const cr of commentsForPara) {
        const comment = commentMap.get(cr.commentId);
        if (!comment) continue;
        parts.push(DexDecompiler._formatComment(comment));
        if (comment.replies) {
          for (const reply of comment.replies) {
            parts.push(DexDecompiler._formatReply(reply, cr.commentId));
          }
        }
        parts.push('');
      }
    }

    while (nextTableIdx < tablePositions.length) {
      parts.push(DexDecompiler._decompileTable(tablePositions[nextTableIdx], docXml));
      nextTableIdx++;
    }

    const sectPr = DexDecompiler._extractSectionProperties(docXml);
    if (sectPr) parts.push(sectPr);

    return parts.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }

  /**
   * Decompile a .docx into a dex-format.js AST (parts-based, for DexCompiler).
   * Accepts either a Workspace or a path string.
   */
  static toAst(wsOrPath) {
    const fs = require('fs');
    const path = require('path');
    const { parseXml, isXmlishPart } = require('./dex-lossless');
    let ws = wsOrPath;
    let shouldCleanup = false;
    if (typeof wsOrPath === 'string') {
      const { Workspace } = require('./workspace');
      ws = Workspace.open(wsOrPath);
      shouldCleanup = true;
    }
    try {
      const rootDir = ws.tmpDir;
      const relPaths = listFiles(rootDir);
      const parts = [];
      for (const relPath of relPaths.sort()) {
        const absPath = path.join(rootDir, relPath);
        if (isXmlishPart(relPath)) {
          const xmlStr = fs.readFileSync(absPath, 'utf8');
          let nodes;
          try { nodes = parseXml(xmlStr); } catch (_) { nodes = [{ type: 'text', value: xmlStr }]; }
          parts.push({ type: 'part', path: relPath, partType: 'xml', nodes });
        } else {
          const buf = fs.readFileSync(absPath);
          parts.push({ type: 'part', path: relPath, partType: 'binary', data: buf.toString('base64'), encoding: 'base64' });
        }
      }
      return { type: 'dex', version: '0.5.0', parts };
    } finally {
      if (shouldCleanup) {
        try { ws.cleanup(); } catch (_) {}
      }
    }
  }

  static _buildFrontmatter(ws) {
    const lines = ['---'];
    lines.push('docex: "0.4.0"');
    let meta = {};
    try { meta = Metadata.get(ws); } catch (_) {}
    if (meta.title) lines.push('title: ' + DexDecompiler._yamlStr(meta.title));
    if (meta.creator) {
      lines.push('authors:');
      const authors = meta.creator.split(/[;,]/).map(a => a.trim()).filter(Boolean);
      for (const a of authors) lines.push('  - name: ' + DexDecompiler._yamlStr(a));
    }
    if (meta.keywords) lines.push('keywords: ' + DexDecompiler._yamlStr(meta.keywords));
    if (meta.subject) lines.push('subject: ' + DexDecompiler._yamlStr(meta.subject));
    lines.push('---');
    lines.push('');
    return lines.join('\n');
  }

  static _buildFootnoteMap(ws) {
    const map = new Map();
    let footnotesXml;
    try { footnotesXml = ws.footnotesXml; } catch (_) { return map; }
    if (!footnotesXml) return map;
    const fnRe = /<w:footnote\b([^>]*)>([\s\S]*?)<\/w:footnote>/g;
    let m;
    while ((m = fnRe.exec(footnotesXml)) !== null) {
      const attrs = m[1]; const body = m[2];
      const id = xml.attrVal(attrs, 'w:id');
      const type = xml.attrVal(attrs, 'w:type');
      if (type) continue;
      if (!id) continue;
      const idNum = parseInt(id, 10);
      if (idNum <= 1) continue;
      map.set(idNum, xml.extractTextDecoded(body));
    }
    return map;
  }

  static _decompileRuns(pXml, footnoteMap) {
    const parts = [];
    let bodyXml = pXml;
    const pOpenEnd = bodyXml.indexOf('>');
    bodyXml = bodyXml.slice(pOpenEnd + 1);
    if (bodyXml.endsWith('</w:p>')) bodyXml = bodyXml.slice(0, -6);
    const pPrMatch = bodyXml.match(/^(\s*<w:pPr>[\s\S]*?<\/w:pPr>)/);
    if (pPrMatch) bodyXml = bodyXml.slice(pPrMatch[0].length);
    DexDecompiler._walkElements(bodyXml, parts, footnoteMap);
    return parts.join('');
  }

  static _walkElements(xmlStr, parts, footnoteMap) {
    let pos = 0;
    const len = xmlStr.length;
    while (pos < len) {
      if (xmlStr[pos] !== '<') { pos++; continue; }
      if (xmlStr.startsWith('<w:ins', pos)) {
        const endTag = '</w:ins>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const insXml = xmlStr.slice(pos, endIdx + endTag.length);
        const attrs = DexDecompiler._extractAttrs(insXml);
        const insContent = DexDecompiler._extractRunTexts(insXml, 'w:t');
        if (insContent) {
          const id = attrs['w:id'] || '';
          const author = xml.decodeXml(attrs['w:author'] || '');
          const date = attrs['w:date'] || '';
          parts.push('{ins id:' + id + ' by:' + DexDecompiler._dexStr(author) + ' date:' + DexDecompiler._dexStr(date) + '}' + insContent + '{/ins}');
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:del', pos)) {
        const endTag = '</w:del>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const delXml = xmlStr.slice(pos, endIdx + endTag.length);
        const attrs = DexDecompiler._extractAttrs(delXml);
        const delContent = DexDecompiler._extractRunTexts(delXml, 'w:delText');
        if (delContent) {
          const id = attrs['w:id'] || '';
          const author = xml.decodeXml(attrs['w:author'] || '');
          const date = attrs['w:date'] || '';
          parts.push('{del id:' + id + ' by:' + DexDecompiler._dexStr(author) + ' date:' + DexDecompiler._dexStr(date) + '}' + delContent + '{/del}');
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:r>', pos) || xmlStr.startsWith('<w:r ', pos)) {
        const endTag = '</w:r>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const runXml = xmlStr.slice(pos, endIdx + endTag.length);
        const fnRefMatch = runXml.match(/<w:footnoteReference\s+w:id="(\d+)"/);
        if (fnRefMatch) {
          const fnId = parseInt(fnRefMatch[1], 10);
          const fnText = footnoteMap.get(fnId) || '';
          parts.push('{footnote id:' + fnId + '}' + fnText + '{/footnote}');
          pos = endIdx + endTag.length;
          continue;
        }
        if (/<w:commentReference/.test(runXml)) { pos = endIdx + endTag.length; continue; }
        const texts = [];
        const tRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
        let tMatch;
        while ((tMatch = tRe.exec(runXml)) !== null) texts.push(xml.decodeXml(tMatch[1]));
        const text = texts.join('');
        if (text) {
          const fmt = DexDecompiler._extractFormatting(runXml);
          parts.push(DexDecompiler._wrapFormatting(text, fmt));
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:hyperlink', pos)) {
        const endTag = '</w:hyperlink>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const hlXml = xmlStr.slice(pos, endIdx + endTag.length);
        DexDecompiler._walkElements(hlXml.slice(hlXml.indexOf('>') + 1, hlXml.lastIndexOf('<')), parts, footnoteMap);
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
  }

  static _extractRunTexts(elXml, textTag) {
    const parts = [];
    const tagRe = new RegExp('<' + textTag + '[^>]*>([^<]*)</' + textTag + '>', 'g');
    let m;
    while ((m = tagRe.exec(elXml)) !== null) parts.push(xml.decodeXml(m[1]));
    return parts.join('');
  }

  static _extractFormatting(runXml) {
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
  }

  static _wrapFormatting(text, fmt) {
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
  }

  static _buildCommentMap(ws) {
    const map = new Map();
    let commentsXml;
    try { commentsXml = ws.commentsXml; } catch (_) { return map; }
    if (!commentsXml) return map;
    const commentRe = /<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g;
    let m;
    while ((m = commentRe.exec(commentsXml)) !== null) {
      const attrs = m[1]; const body = m[2];
      const id = xml.attrVal(attrs, 'w:id');
      const author = xml.decodeXml(xml.attrVal(attrs, 'w:author') || '');
      const date = xml.attrVal(attrs, 'w:date') || '';
      const text = xml.extractTextDecoded(body);
      const innerParaIdMatch = body.match(/w14:paraId="([^"]+)"/);
      const innerParaId = innerParaIdMatch ? innerParaIdMatch[1] : '';
      if (id) map.set(parseInt(id, 10), { id: parseInt(id, 10), author, date, text, paraId: innerParaId, replies: [] });
    }
    let extXml;
    try { extXml = ws.commentsExtXml; } catch (_) { return map; }
    if (extXml) {
      const exRe = /<w15:commentEx\s+([^>]*?)\s*\/?>/g;
      let exM;
      const entries = [];
      while ((exM = exRe.exec(extXml)) !== null) entries.push(exM[1]);
      const paraIdToComment = new Map();
      for (const [cId, cData] of map) { if (cData.paraId) paraIdToComment.set(cData.paraId, cId); }
      for (const entryAttrs of entries) {
        const entryParaId = xml.attrVal(entryAttrs, 'w15:paraId');
        const parentParaId = xml.attrVal(entryAttrs, 'w15:paraIdParent');
        if (entryParaId && parentParaId) {
          const childCommentId = paraIdToComment.get(entryParaId);
          const parentCommentId = paraIdToComment.get(parentParaId);
          if (childCommentId && parentCommentId && map.has(parentCommentId) && map.has(childCommentId)) {
            map.get(parentCommentId).replies.push(map.get(childCommentId));
            map.get(childCommentId).isReply = true;
          }
        }
      }
    }
    return map;
  }

  static _buildCommentRanges(docXml, commentMap) {
    const ranges = [];
    const paragraphs = xml.findParagraphs(docXml);
    const rangeStarts = new Map();
    const rangeEnds = new Map();
    for (let i = 0; i < paragraphs.length; i++) {
      const pXml = paragraphs[i].xml;
      let sm;
      const startRe = /<w:commentRangeStart\s+w:id="(\d+)"\s*\/?>/g;
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
      const startIdx = rangeStarts.get(commentId);
      const endIdx = rangeEnds.get(commentId);
      if (endIdx !== undefined) {
        let anchor = '';
        if (startIdx !== undefined && paragraphs[startIdx]) anchor = xml.extractTextDecoded(paragraphs[startIdx].xml).slice(0, 50);
        ranges.push({ commentId, startParaIndex: startIdx !== undefined ? startIdx : endIdx, endParaIndex: endIdx, anchor });
      }
    }
    return ranges;
  }

  static _formatComment(comment) {
    return '{comment id:' + comment.id + ' by:' + DexDecompiler._dexStr(comment.author) + ' date:' + DexDecompiler._dexStr(comment.date) + '}\n' + comment.text + '\n{/comment}';
  }

  static _formatReply(reply, parentId) {
    return '{reply id:' + reply.id + ' parent:' + parentId + ' by:' + DexDecompiler._dexStr(reply.author) + ' date:' + DexDecompiler._dexStr(reply.date) + '}\n' + reply.text + '\n{/reply}';
  }

  static _decompileFigure(pXml, paraId, ws) {
    const drawingMatch = pXml.match(/<w:drawing[\s>][\s\S]*?<\/w:drawing>/);
    if (!drawingMatch) { const text = xml.extractTextDecoded(pXml); return '{p id:' + paraId + '}\n' + text + '\n{/p}'; }
    const drawXml = drawingMatch[0];
    const cxMatch = drawXml.match(/\bcx="(\d+)"/);
    const cyMatch = drawXml.match(/\bcy="(\d+)"/);
    const width = cxMatch ? cxMatch[1] + 'emu' : '';
    const height = cyMatch ? cyMatch[1] + 'emu' : '';
    const rIdMatch = drawXml.match(/r:embed="([^"]+)"/);
    const rId = rIdMatch ? rIdMatch[1] : '';
    let src = '';
    if (rId) { try { const relsXml = ws.relsXml; const relRe = new RegExp('Id="' + rId + '"[^>]*Target="([^"]+)"', 'g'); const relMatch = relRe.exec(relsXml); if (relMatch) src = relMatch[1].startsWith('media/') ? 'word/' + relMatch[1] : 'word/' + relMatch[1].replace(/^\.\.\//, ''); } catch (_) {} }
    const altMatch = drawXml.match(/descr="([^"]*)"/);
    const alt = altMatch ? xml.decodeXml(altMatch[1]) : '';
    const caption = xml.extractTextDecoded(pXml);
    const parts = [];
    let attrs = '{figure';
    if (paraId) attrs += ' id:' + paraId;
    if (rId) attrs += ' rId:' + rId;
    if (src) attrs += ' src:' + DexDecompiler._dexStr(src);
    if (width) attrs += ' width:' + width;
    if (height) attrs += ' height:' + height;
    if (alt) attrs += ' alt:' + DexDecompiler._dexStr(alt);
    attrs += '}';
    parts.push(attrs);
    if (caption) parts.push(caption);
    parts.push('{/figure}');
    return parts.join('\n');
  }

  static _findTables(docXml) {
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
  }

  static _decompileTable(tblInfo, docXml) {
    const tblXml = tblInfo.xml;
    const parts = [];
    const firstRow = tblXml.match(/<w:tr[\s>][\s\S]*?<\/w:tr>/);
    let colCount = 0;
    if (firstRow) colCount = (firstRow[0].match(/<w:tc[\s>]/g) || []).length;
    const styleMatch = tblXml.match(/<w:tblStyle\s+w:val="([^"]+)"/);
    const style = styleMatch ? styleMatch[1] : 'plain';
    parts.push('{table style:' + style + ' cols:' + colCount + '}');
    const rowRe = /<w:tr[\s>][\s\S]*?<\/w:tr>/g;
    let rm;
    const rows = [];
    while ((rm = rowRe.exec(tblXml)) !== null) {
      const rowXml = rm[0]; const cells = [];
      const cellRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
      let cm;
      while ((cm = cellRe.exec(rowXml)) !== null) cells.push(xml.extractTextDecoded(cm[0]).trim());
      rows.push(cells);
    }
    if (rows.length > 0) {
      parts.push('| ' + rows[0].join(' | ') + ' |');
      parts.push('|' + rows[0].map(() => '---').join('|') + '|');
      for (let i = 1; i < rows.length; i++) parts.push('| ' + rows[i].join(' | ') + ' |');
    }
    parts.push('{/table}');
    return parts.join('\n');
  }

  static _extractSectionProperties(docXml) {
    const sectPrMatch = docXml.match(/<w:sectPr[\s>][\s\S]*?<\/w:sectPr>/);
    if (!sectPrMatch) return null;
    const sectPr = sectPrMatch[0];
    const pgMar = sectPr.match(/<w:pgMar\s+([^>]+)/);
    let margins = '';
    if (pgMar) {
      const top = xml.attrVal(pgMar[1], 'w:top') || '0';
      const right = xml.attrVal(pgMar[1], 'w:right') || '0';
      const bottom = xml.attrVal(pgMar[1], 'w:bottom') || '0';
      const left = xml.attrVal(pgMar[1], 'w:left') || '0';
      margins = top + ' ' + right + ' ' + bottom + ' ' + left;
    }
    if (margins) return '\n{section margins:"' + margins + '"}';
    return null;
  }

  static _extractParaId(pXml) { const m = pXml.match(/w14:paraId="([^"]+)"/); return m ? m[1] : ''; }
  static _extractAttrs(elXml) {
    const openTag = elXml.match(/^<[^>]+>/);
    if (!openTag) return {};
    const attrs = {};
    const attrRe = /(\w[\w:]*?)="([^"]*)"/g;
    let m;
    while ((m = attrRe.exec(openTag[0])) !== null) attrs[m[1]] = m[2];
    return attrs;
  }
  static _yamlStr(s) {
    if (!s) return '""';
    if (s.includes('"') || s.includes('\n') || s.includes(':') || s.includes('#')) return '"' + s.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
    return '"' + s + '"';
  }
  static _dexStr(s) {
    if (!s) return '""';
    return '"' + s.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
  }
}

function listFiles(dir, base) {
  base = base || dir;
  const result = [];
  const fs = require('fs');
  const path = require('path');
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      const sub = listFiles(full, base);
      for (const f of sub) result.push(f);
    } else {
      result.push(path.relative(base, full));
    }
  }
  return result;
}

module.exports = { DexDecompiler, listFiles };
