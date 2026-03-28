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
        // Use _decompileRuns for headings too, so comment anchors are preserved
        const result = DexDecompiler._decompileRuns(pXml, footnoteMap);
        const content = result.text;
        const hashes = '#'.repeat(level);
        const idAttr = paraId ? ' {id:' + paraId + '}' : '';
        parts.push(hashes + ' ' + content + idAttr);
        parts.push('');
      } else if (hasFigure) {
        parts.push(DexDecompiler._decompileFigure(pXml, paraId, ws));
        parts.push('');
      } else {
        const result = DexDecompiler._decompileRuns(pXml, footnoteMap);
        const content = result.text;
        const pProps = result.props;
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
        const propsAttr = DexDecompiler._formatParagraphProps(pProps);
        parts.push('{p' + idAttr + propsAttr + '}');
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
    let pProps = {};
    if (pPrMatch) {
      pProps = DexDecompiler._extractParagraphProperties(pPrMatch[1]);
      bodyXml = bodyXml.slice(pPrMatch[0].length);
    }
    DexDecompiler._walkElements(bodyXml, parts, footnoteMap);
    return { text: parts.join(''), props: pProps };
  }

  /**
   * Extract key paragraph properties from <w:pPr> XML.
   */
  static _extractParagraphProperties(pPrXml) {
    const props = {};

    // 1. Alignment
    const jcMatch = pPrXml.match(/<w:jc\s+w:val="([^"]+)"/);
    if (jcMatch) {
      props.align = jcMatch[1] === 'both' ? 'justify' : jcMatch[1];
    }

    // 2. Paragraph style (skip heading styles)
    const pStyleMatch = pPrXml.match(/<w:pStyle\s+w:val="([^"]+)"/);
    if (pStyleMatch && !/^Heading\d+$/i.test(pStyleMatch[1])) {
      props.style = pStyleMatch[1];
    }

    // 3. Indentation
    const indMatch = pPrXml.match(/<w:ind\s+([^>]+)/);
    if (indMatch) {
      const a = indMatch[1];
      const lm = a.match(/w:left="(\d+)"/);
      const rm = a.match(/w:right="(\d+)"/);
      const fm = a.match(/w:firstLine="(\d+)"/);
      const hm = a.match(/w:hanging="(\d+)"/);
      if (lm && lm[1] !== '0') props['indent-left'] = lm[1];
      if (rm && rm[1] !== '0') props['indent-right'] = rm[1];
      if (fm && fm[1] !== '0') props['indent-first'] = fm[1];
      if (hm && hm[1] !== '0') props['indent-hanging'] = hm[1];
    }

    // 4. Spacing
    const spacingMatch = pPrXml.match(/<w:spacing\s+([^>]+)/);
    if (spacingMatch) {
      const a = spacingMatch[1];
      const lm = a.match(/w:line="(\d+)"/);
      const bm = a.match(/w:before="(\d+)"/);
      const am = a.match(/w:after="(\d+)"/);
      const rlm = a.match(/w:lineRule="([^"]+)"/);
      if (lm) props['spacing-line'] = lm[1];
      if (bm && bm[1] !== '0') props['spacing-before'] = bm[1];
      if (am && am[1] !== '0') props['spacing-after'] = am[1];
      if (rlm && rlm[1] !== 'auto') props['spacing-rule'] = rlm[1];
    }

    // 5. List numbering
    const numPrMatch = pPrXml.match(/<w:numPr>([\s\S]*?)<\/w:numPr>/);
    if (numPrMatch) {
      const nb = numPrMatch[1];
      const nid = nb.match(/<w:numId\s+w:val="(\d+)"/);
      const ilv = nb.match(/<w:ilvl\s+w:val="(\d+)"/);
      if (nid) props['list-id'] = nid[1];
      if (ilv) props['list-level'] = ilv[1];
    }

    // 6. Right-to-left
    if (/<w:bidi\b/.test(pPrXml) && !/<w:bidi\s+w:val="(false|0)"/.test(pPrXml)) {
      props.bidi = true;
    }

    // 7. Keep with next
    if (/<w:keepNext\b/.test(pPrXml) && !/<w:keepNext\s+w:val="(false|0)"/.test(pPrXml)) {
      props.keepnext = true;
    }

    return props;
  }

  /**
   * Format paragraph properties into .dex attribute string.
   */
  static _formatParagraphProps(props) {
    if (!props || Object.keys(props).length === 0) return '';
    const parts = [];
    const quotedKeys = ['style'];
    const boolKeys = ['bidi', 'keepnext'];
    for (const [key, val] of Object.entries(props)) {
      if (boolKeys.includes(key)) {
        parts.push(key);
      } else if (quotedKeys.includes(key)) {
        parts.push(key + ':"' + String(val).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"');
      } else {
        parts.push(key + ':' + val);
      }
    }
    return ' ' + parts.join(' ');
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
        // Recurse into inner content to preserve formatting (bold, italic, etc.)
        const innerParts = [];
        const innerXml = insXml.slice(insXml.indexOf('>') + 1, insXml.lastIndexOf('</w:ins>'));
        DexDecompiler._walkElements(innerXml, innerParts, footnoteMap);
        const insContent = innerParts.join('');
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
        // Extract formatted deleted text from runs containing w:delText
        const delContent = DexDecompiler._extractFormattedDelTexts(delXml);
        if (delContent) {
          const id = attrs['w:id'] || '';
          const author = xml.decodeXml(attrs['w:author'] || '');
          const date = attrs['w:date'] || '';
          parts.push('{del id:' + id + ' by:' + DexDecompiler._dexStr(author) + ' date:' + DexDecompiler._dexStr(date) + '}' + delContent + '{/del}');
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:moveFrom', pos) && !xmlStr.startsWith('<w:moveFromRangeStart', pos) && !xmlStr.startsWith('<w:moveFromRangeEnd', pos)) {
        const endTag = '</w:moveFrom>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const mfXml = xmlStr.slice(pos, endIdx + endTag.length);
        const attrs = DexDecompiler._extractAttrs(mfXml);
        // Recurse: inner runs use w:delText (same as w:del)
        const mfContent = DexDecompiler._extractFormattedDelTexts(mfXml);
        if (mfContent) {
          const id = attrs['w:id'] || '';
          const author = xml.decodeXml(attrs['w:author'] || '');
          const date = attrs['w:date'] || '';
          parts.push('{movefrom id:' + id + ' by:' + DexDecompiler._dexStr(author) + ' date:' + DexDecompiler._dexStr(date) + '}' + mfContent + '{/movefrom}');
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:moveTo', pos) && !xmlStr.startsWith('<w:moveToRangeStart', pos) && !xmlStr.startsWith('<w:moveToRangeEnd', pos)) {
        const endTag = '</w:moveTo>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const mtXml = xmlStr.slice(pos, endIdx + endTag.length);
        const attrs = DexDecompiler._extractAttrs(mtXml);
        // Recurse: inner runs use w:t (same as w:ins)
        const innerParts = [];
        const innerXml = mtXml.slice(mtXml.indexOf('>') + 1, mtXml.lastIndexOf('</w:moveTo>'));
        DexDecompiler._walkElements(innerXml, innerParts, footnoteMap);
        const mtContent = innerParts.join('');
        if (mtContent) {
          const id = attrs['w:id'] || '';
          const author = xml.decodeXml(attrs['w:author'] || '');
          const date = attrs['w:date'] || '';
          parts.push('{moveto id:' + id + ' by:' + DexDecompiler._dexStr(author) + ' date:' + DexDecompiler._dexStr(date) + '}' + mtContent + '{/moveto}');
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
        // Endnote references
        const enRefMatch = runXml.match(/<w:endnoteReference\s+w:id="(\d+)"/);
        if (enRefMatch) {
          parts.push('{endnote id:' + enRefMatch[1] + '}');
          pos = endIdx + endTag.length;
          continue;
        }
        // Extract text, tabs, breaks, and symbols from the run
        const texts = [];
        const runBody = runXml.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '');
        const elemRe = /<w:t[^>]*>([^<]*)<\/w:t>|<w:tab\s*\/>|<w:br\s*\/?>|<w:br\s+w:type="([^"]*)"[^>]*\/?>|<w:sym\s+[^>]*w:char="([^"]*)"[^>]*\/?>/g;
        let tMatch;
        while ((tMatch = elemRe.exec(runBody)) !== null) {
          if (tMatch[0].startsWith('<w:t')) {
            texts.push(xml.decodeXml(tMatch[1]));
          } else if (tMatch[0].startsWith('<w:tab')) {
            texts.push('\t');
          } else if (tMatch[0].startsWith('<w:br')) {
            const brType = tMatch[2] || '';
            if (brType === 'page') texts.push('{pagebreak}');
            else if (brType === 'column') texts.push('{colbreak}');
            else texts.push('{br}');
          } else if (tMatch[0].startsWith('<w:sym')) {
            texts.push('{sym ' + (tMatch[3] || '') + '}');
          }
        }
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
        // Extract link target (r:id for relationship-based, w:anchor for bookmark-based)
        const rIdMatch = hlXml.match(/r:id="([^"]*)"/);
        const anchorMatch = hlXml.match(/w:anchor="([^"]*)"/);
        const linkParts = [];
        DexDecompiler._walkElements(hlXml.slice(hlXml.indexOf('>') + 1, hlXml.lastIndexOf('<')), linkParts, footnoteMap);
        const linkText = linkParts.join('');
        if (rIdMatch) {
          parts.push('{link rId:' + rIdMatch[1] + '}' + linkText + '{/link}');
        } else if (anchorMatch) {
          parts.push('{link anchor:' + DexDecompiler._dexStr(anchorMatch[1]) + '}' + linkText + '{/link}');
        } else {
          parts.push(linkText); // fallback: just the text
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:commentRangeStart', pos)) {
        const closeAngle = xmlStr.indexOf('>', pos);
        const tag = xmlStr.slice(pos, closeAngle + 1);
        const idMatch = tag.match(/w:id="(\d+)"/);
        if (idMatch) parts.push('{comment-start id:' + idMatch[1] + '}');
        pos = closeAngle + 1;
      } else if (xmlStr.startsWith('<w:commentRangeEnd', pos)) {
        const closeAngle = xmlStr.indexOf('>', pos);
        const tag = xmlStr.slice(pos, closeAngle + 1);
        const idMatch = tag.match(/w:id="(\d+)"/);
        if (idMatch) parts.push('{comment-end id:' + idMatch[1] + '}');
        pos = closeAngle + 1;
      } else if (xmlStr.startsWith('<w:bookmarkStart', pos)) {
        const closeAngle = xmlStr.indexOf('>', pos);
        const tag = xmlStr.slice(pos, closeAngle + 1);
        const idMatch = tag.match(/w:id="(\d+)"/);
        const nameMatch = tag.match(/w:name="([^"]*)"/);
        if (idMatch && nameMatch && nameMatch[1] !== '_GoBack') {
          parts.push('{bookmark-start id:' + idMatch[1] + ' name:' + DexDecompiler._dexStr(nameMatch[1]) + '}');
        }
        pos = closeAngle + 1;
      } else if (xmlStr.startsWith('<w:bookmarkEnd', pos)) {
        const closeAngle = xmlStr.indexOf('>', pos);
        const tag = xmlStr.slice(pos, closeAngle + 1);
        const idMatch = tag.match(/w:id="(\d+)"/);
        if (idMatch) parts.push('{bookmark-end id:' + idMatch[1] + '}');
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
    while ((m = tagRe.exec(elXml)) !== null) {
      // Decode XML entities, then escape .dex special characters
      let text = xml.decodeXml(m[1]);
      text = text.replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}');
      parts.push(text);
    }
    return parts.join('');
  }

  /**
   * Extract deleted text from runs inside a w:del (or w:moveFrom) block,
   * preserving formatting. Runs inside these blocks use <w:delText> instead
   * of <w:t>. We walk each <w:r> manually, extract formatting + delText,
   * and wrap with formatting tags.
   */
  static _extractFormattedDelTexts(elXml) {
    const parts = [];
    const runRe = /<w:r[\s>][\s\S]*?<\/w:r>/g;
    let m;
    while ((m = runRe.exec(elXml)) !== null) {
      const runXml = m[0];
      // Extract delText content
      const texts = [];
      const dtRe = /<w:delText[^>]*>([^<]*)<\/w:delText>/g;
      let dt;
      while ((dt = dtRe.exec(runXml)) !== null) {
        texts.push(xml.decodeXml(dt[1]));
      }
      const text = texts.join('');
      if (text) {
        const fmt = DexDecompiler._extractFormatting(runXml);
        parts.push(DexDecompiler._wrapFormatting(text, fmt));
      }
    }
    return parts.join('');
  }

  static _extractFormatting(runXml) {
    const rPrMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    if (!rPrMatch) return {};
    const rPr = rPrMatch[1];
    const fmt = {};
    if (/<w:b\b[^>]*\/?>/.test(rPr) && !/<w:b\s+w:val="(false|0)"/.test(rPr)) fmt.bold = true;
    if (/<w:i\b[^>]*\/?>/.test(rPr) && !/<w:i\s+w:val="(false|0)"/.test(rPr)) fmt.italic = true;
    // Underline: extract type attribute
    const uMatch = rPr.match(/<w:u\b([^>]*)\/?>/);
    if (uMatch && !/<w:u\s+w:val="none"/.test(rPr)) {
      const uValMatch = uMatch[1].match(/w:val="([^"]+)"/);
      const uType = uValMatch ? uValMatch[1] : 'single';
      if (uType === 'single') {
        fmt.underline = true;
      } else {
        fmt.underlineType = uType;
      }
    }
    if (/<w:vertAlign\s+w:val="superscript"/.test(rPr)) fmt.sup = true;
    if (/<w:vertAlign\s+w:val="subscript"/.test(rPr)) fmt.sub = true;
    // Strikethrough
    if (/<w:strike\b[^>]*\/?>/.test(rPr) && !/<w:strike\s+w:val="(false|0)"/.test(rPr)) fmt.strike = true;
    if (/<w:dstrike\b[^>]*\/?>/.test(rPr) && !/<w:dstrike\s+w:val="(false|0)"/.test(rPr)) fmt.dstrike = true;
    // Font size (half-points)
    const szMatch = rPr.match(/<w:sz\s+w:val="(\d+)"/);
    if (szMatch) fmt.size = szMatch[1];
    // Small caps
    if (/<w:smallCaps\b[^>]*\/?>/.test(rPr) && !/<w:smallCaps\s+w:val="(false|0)"/.test(rPr)) fmt.smallcaps = true;
    // All caps
    if (/<w:caps\b[^>]*\/?>/.test(rPr) && !/<w:caps\s+w:val="(false|0)"/.test(rPr)) fmt.caps = true;
    // Hidden text
    if (/<w:vanish\b[^>]*\/?>/.test(rPr) && !/<w:vanish\s+w:val="(false|0)"/.test(rPr)) fmt.hidden = true;
    const hlMatch = rPr.match(/<w:highlight\s+w:val="([^"]+)"/);
    if (hlMatch) fmt.highlight = hlMatch[1];
    const colorMatch = rPr.match(/<w:color\s+w:val="([^"]+)"/);
    if (colorMatch && colorMatch[1] !== '000000' && colorMatch[1] !== 'auto') fmt.color = colorMatch[1];
    // Font family: extract ascii plus any differing variants
    const rFontsMatch = rPr.match(/<w:rFonts\b([^>]*)\/?>/);
    if (rFontsMatch) {
      const fontsAttrs = rFontsMatch[1];
      const asciiMatch = fontsAttrs.match(/w:ascii="([^"]+)"/);
      if (asciiMatch) {
        fmt.font = asciiMatch[1];
        const hAnsiMatch = fontsAttrs.match(/w:hAnsi="([^"]+)"/);
        const csMatch = fontsAttrs.match(/w:cs="([^"]+)"/);
        const eastAsiaMatch = fontsAttrs.match(/w:eastAsia="([^"]+)"/);
        if (hAnsiMatch && hAnsiMatch[1] !== asciiMatch[1]) fmt.fontHAnsi = hAnsiMatch[1];
        if (csMatch && csMatch[1] !== asciiMatch[1]) fmt.fontCs = csMatch[1];
        if (eastAsiaMatch && eastAsiaMatch[1] !== asciiMatch[1]) fmt.fontEastAsia = eastAsiaMatch[1];
      }
    }
    // Tracked formatting change: w:rPrChange records previous formatting state
    const rPrChangeMatch = rPr.match(/<w:rPrChange\b([^>]*)>/);
    if (rPrChangeMatch) {
      const changeAttrs = rPrChangeMatch[1];
      const changeAuthor = xml.attrVal(changeAttrs, 'w:author');
      const changeDate = xml.attrVal(changeAttrs, 'w:date');
      fmt.fmtchange = {};
      if (changeAuthor) fmt.fmtchange.author = xml.decodeXml(changeAuthor);
      if (changeDate) fmt.fmtchange.date = changeDate;
    }
    return fmt;
  }

  static _wrapFormatting(text, fmt) {
    if (!fmt || Object.keys(fmt).length === 0) return text;
    let result = text;
    if (fmt.bold) result = '{b}' + result + '{/b}';
    if (fmt.italic) result = '{i}' + result + '{/i}';
    if (fmt.underline) result = '{u}' + result + '{/u}';
    if (fmt.underlineType) result = '{u ' + fmt.underlineType + '}' + result + '{/u}';
    if (fmt.sup) result = '{sup}' + result + '{/sup}';
    if (fmt.sub) result = '{sub}' + result + '{/sub}';
    if (fmt.strike) result = '{strike}' + result + '{/strike}';
    if (fmt.dstrike) result = '{dstrike}' + result + '{/dstrike}';
    if (fmt.size) result = '{size ' + fmt.size + '}' + result + '{/size}';
    if (fmt.smallcaps) result = '{smallcaps}' + result + '{/smallcaps}';
    if (fmt.caps) result = '{caps}' + result + '{/caps}';
    if (fmt.hidden) result = '{hidden}' + result + '{/hidden}';
    if (fmt.highlight) result = '{highlight ' + fmt.highlight + '}' + result + '{/highlight}';
    if (fmt.color) result = '{color ' + fmt.color + '}' + result + '{/color}';
    if (fmt.font) {
      let fontAttrs = '"' + fmt.font + '"';
      if (fmt.fontHAnsi) fontAttrs += ' hAnsi:"' + fmt.fontHAnsi + '"';
      if (fmt.fontCs) fontAttrs += ' cs:"' + fmt.fontCs + '"';
      if (fmt.fontEastAsia) fontAttrs += ' eastAsia:"' + fmt.fontEastAsia + '"';
      result = '{font ' + fontAttrs + '}' + result + '{/font}';
    }
    if (fmt.fmtchange) {
      let fmtAttrs = '';
      if (fmt.fmtchange.author) fmtAttrs += ' by:' + DexDecompiler._dexStr(fmt.fmtchange.author);
      if (fmt.fmtchange.date) fmtAttrs += ' date:' + DexDecompiler._dexStr(fmt.fmtchange.date);
      result = '{fmtchange' + fmtAttrs + '}' + result + '{/fmtchange}';
    }
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
