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
    // Detect default font from styles.xml to suppress redundant {font} wrappers
    const stylesXml = ws.stylesXml || '';
    const defaultFontMatch = stylesXml.match(/<w:rFonts[^>]*w:ascii="([^"]+)"/);
    DexDecompiler._defaultFont = defaultFontMatch ? defaultFontMatch[1] : 'Times New Roman';

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
        const result = DexDecompiler._decompileRuns(pXml, footnoteMap, commentMap);
        const content = result.text;
        const hashes = '#'.repeat(level);
        const idAttr = paraId ? ' {id:' + paraId + '}' : '';
        parts.push(hashes + ' ' + content + idAttr);
        parts.push('');
      } else if (hasFigure) {
        parts.push(DexDecompiler._decompileFigure(pXml, paraId, ws));
        parts.push('');
      } else {
        const result = DexDecompiler._decompileRuns(pXml, footnoteMap, commentMap);
        const content = result.text;
        const pProps = result.props;
        if (content.trim() === '') {
          if (/<w:br\s+w:type="page"/.test(pXml)) {
            parts.push('{pagebreak}');
            parts.push('');
          }
          // Check for inline section break even in empty paragraphs
          const pPrForSectEmpty = pXml.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/);
          if (pPrForSectEmpty) {
            const sectInEmpty = pPrForSectEmpty[1].match(/<w:sectPr[\s>][\s\S]*?<\/w:sectPr>/);
            if (sectInEmpty) {
              const sectLine = DexDecompiler._formatSectionFromXml(sectInEmpty[0]);
              if (sectLine) { parts.push(sectLine.trim()); parts.push(''); }
            }
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

      // Detect inline section breaks (<w:sectPr> inside <w:pPr>) for non-empty paragraphs
      const pPrForSect = pXml.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/);
      if (pPrForSect) {
        const sectInPara = pPrForSect[1].match(/<w:sectPr[\s>][\s\S]*?<\/w:sectPr>/);
        if (sectInPara) {
          const sectLine = DexDecompiler._formatSectionFromXml(sectInPara[0]);
          if (sectLine) parts.push(sectLine.trim());
        }
      }

      // Only emit separate blocks for comment replies (main comments are now inline)
      const commentsForPara = commentRanges.filter(cr => cr.endParaIndex === i);
      for (const cr of commentsForPara) {
        const comment = commentMap.get(cr.commentId);
        if (!comment) continue;
        if (comment.replies && comment.replies.length > 0) {
          for (const reply of comment.replies) {
            parts.push(DexDecompiler._formatReply(reply, cr.commentId));
          }
          parts.push('');
        }
      }
    }

    while (nextTableIdx < tablePositions.length) {
      parts.push(DexDecompiler._decompileTable(tablePositions[nextTableIdx], docXml));
      nextTableIdx++;
    }

    const sectPr = DexDecompiler._extractSectionProperties(docXml);
    if (sectPr) parts.push(sectPr);

    // Endnotes
    const endnotesXml = ws.endnotesXml;
    if (endnotesXml) {
      const endnoteMap = DexDecompiler._buildEndnoteMap(endnotesXml);
      if (endnoteMap.size > 0) {
        parts.push('');
        parts.push('{endnotes}');
        for (const [id, text] of endnoteMap) {
          parts.push('{endnote-def id:' + id + '}');
          parts.push(text);
          parts.push('{/endnote-def}');
        }
        parts.push('{/endnotes}');
      }
    }

    // Headers and footers
    const hdrFtrs = ws.listHeaderFooters();
    if (hdrFtrs.length > 0) {
      parts.push('');
      for (const hf of hdrFtrs) {
        const tag = hf.type; // 'header' or 'footer'
        const text = xml.extractTextDecoded(hf.xml);
        if (text.trim()) {
          parts.push('{' + tag + ' file:' + DexDecompiler._dexStr(hf.path) + '}');
          parts.push(text.trim());
          parts.push('{/' + tag + '}');
          parts.push('');
        }
      }
    }

    return parts.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }

  /**
   * Build endnote map from endnotes.xml (similar to footnotes).
   */
  static _buildEndnoteMap(endnotesXml) {
    const map = new Map();
    if (!endnotesXml) return map;
    const re = /<w:endnote\b([^>]*)>([\s\S]*?)<\/w:endnote>/g;
    let m;
    while ((m = re.exec(endnotesXml)) !== null) {
      const attrs = m[1];
      const typeMatch = attrs.match(/w:type="([^"]*)"/);
      if (typeMatch && (typeMatch[1] === 'separator' || typeMatch[1] === 'continuationSeparator')) continue;
      const idMatch = attrs.match(/w:id="(\d+)"/);
      if (!idMatch) continue;
      const id = parseInt(idMatch[1], 10);
      const text = DexDecompiler._extractFormattedText(m[2]);
      if (text.trim()) map.set(id, text.trim());
    }
    return map;
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

  /**
   * Extract formatted text from XML that contains <w:p> paragraphs with <w:r> runs.
   * Preserves inline formatting (bold, italic, underline, hyperlinks, etc.)
   * by reusing _extractFormatting and _wrapFormatting. Multiple paragraphs
   * are joined with newline characters.
   */
  static _extractFormattedText(bodyXml) {
    const paraRe = /<w:p[\s>][\s\S]*?<\/w:p>/g;
    let pm;
    const paragraphs = [];
    while ((pm = paraRe.exec(bodyXml)) !== null) {
      paragraphs.push(pm[0]);
    }
    if (paragraphs.length === 0) {
      return DexDecompiler._extractFormattedRunsFromXml(bodyXml);
    }
    const paraTexts = [];
    for (const pXml of paragraphs) {
      const text = DexDecompiler._extractFormattedRunsFromXml(pXml);
      paraTexts.push(text);
    }
    return paraTexts.join('\n');
  }

  /**
   * Extract formatted text from runs (<w:r>) within a chunk of XML.
   * Handles hyperlinks, text, formatting, tabs, breaks, and symbols.
   */
  static _extractFormattedRunsFromXml(xmlStr) {
    const parts = [];
    let pos = 0;
    const len = xmlStr.length;
    while (pos < len) {
      if (xmlStr[pos] !== '<') { pos++; continue; }
      if (xmlStr.startsWith('<w:hyperlink', pos)) {
        const endTag = '</w:hyperlink>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const hlXml = xmlStr.slice(pos, endIdx + endTag.length);
        const rIdMatch = hlXml.match(/r:id="([^"]*)"/);
        const anchorMatch = hlXml.match(/w:anchor="([^"]*)"/);
        const innerText = DexDecompiler._extractFormattedRunsFromXml(
          hlXml.slice(hlXml.indexOf('>') + 1, hlXml.lastIndexOf('<'))
        );
        if (rIdMatch) {
          parts.push('{link rId:' + rIdMatch[1] + '}' + innerText + '{/link}');
        } else if (anchorMatch) {
          parts.push('{link anchor:' + DexDecompiler._dexStr(anchorMatch[1]) + '}' + innerText + '{/link}');
        } else {
          parts.push(innerText);
        }
        pos = endIdx + endTag.length;
        continue;
      }
      if (xmlStr.startsWith('<w:r>', pos) || xmlStr.startsWith('<w:r ', pos)) {
        const endTag = '</w:r>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const runXml = xmlStr.slice(pos, endIdx + endTag.length);
        const texts = [];
        const runBody = runXml.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '');
        const elemRe = /<w:t[^>]*>([^<]*)<\/w:t>|<w:tab\s*\/>|<w:br\s*\/?>|<w:br\s+w:type="([^"]*)"[^>]*\/?>|<w:sym\s+[^>]*w:char="([^"]*)"[^>]*\/?>|<w:softHyphen\s*\/>|<w:noBreakHyphen\s*\/?>/g;
        let tMatch;
        while ((tMatch = elemRe.exec(runBody)) !== null) {
          if (tMatch[0].startsWith('<w:t')) {
            texts.push(xml.decodeXml(tMatch[1]));
          } else if (tMatch[0].startsWith('<w:tab')) {
            texts.push('\t');
          } else if (tMatch[0].startsWith('<w:br')) {
            const brType = tMatch[2] || '';
            if (brType === 'column') texts.push('{colbreak}');
            else texts.push('{br}');
          } else if (tMatch[0].startsWith('<w:sym')) {
            texts.push('{sym ' + (tMatch[3] || '') + '}');
          } else if (tMatch[0].startsWith('<w:softHyphen')) {
            texts.push('\u00AD'); // Unicode soft hyphen
          } else if (tMatch[0].startsWith('<w:noBreakHyphen')) {
            texts.push('\u2011'); // Unicode non-breaking hyphen
          }
        }
        const text = texts.join('');
        if (text) {
          const fmt = DexDecompiler._extractFormatting(runXml);
          parts.push(DexDecompiler._wrapFormatting(text, fmt));
        }
        pos = endIdx + endTag.length;
        continue;
      }
      const closeAngle = xmlStr.indexOf('>', pos);
      if (closeAngle === -1) break;
      pos = closeAngle + 1;
    }
    return parts.join('');
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
      map.set(idNum, DexDecompiler._extractFormattedText(body));
    }
    return map;
  }

  static _decompileRuns(pXml, footnoteMap, commentMap) {
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
    // Pre-process field codes: replace fldChar begin..separate..end sequences
    // with {field "INSTRUCTION"}display{/field} markers
    bodyXml = DexDecompiler._preprocessFieldCodes(bodyXml);
    DexDecompiler._walkElements(bodyXml, parts, footnoteMap, commentMap);
    return { text: DexDecompiler._mergeAdjacentRuns(parts.join('')), props: pProps };
  }

  /**
   * Merge adjacent runs that share identical formatting wrappers.
   * For example: {font "Arial"}{b}Hello {/b}{/font}{font "Arial"}{b}World{/b}{/font}
   * becomes:     {font "Arial"}{b}Hello World{/b}{/font}
   *
   * Tokenizes the text, then removes close/reopen pairs where the opening
   * tag is identical. Multiple passes handle nested formatting.
   */
  static _mergeAdjacentRuns(text) {
    if (!text) return text;
    // Tokenize into tags and text segments
    const tokenRe = /\{\/[a-z]+\}|\{[a-z]+(?:\s[^}]*)?\}/g;
    const tokens = [];
    let last = 0;
    let m;
    while ((m = tokenRe.exec(text)) !== null) {
      if (m.index > last) tokens.push({ type: 'text', value: text.slice(last, m.index) });
      const val = m[0];
      if (val.startsWith('{/')) {
        tokens.push({ type: 'close', value: val, tag: val.slice(2, -1) });
      } else {
        // Extract tag name (first word after {)
        const spaceIdx = val.indexOf(' ');
        const tag = spaceIdx > 0 ? val.slice(1, spaceIdx) : val.slice(1, -1);
        tokens.push({ type: 'open', value: val, tag: tag });
      }
      last = m.index + val.length;
    }
    if (last < text.length) tokens.push({ type: 'text', value: text.slice(last) });

    // Repeatedly scan for adjacent close/open pairs with identical full tags
    let changed = true;
    while (changed) {
      changed = false;
      for (let i = 0; i < tokens.length - 1; i++) {
        const cur = tokens[i];
        const next = tokens[i + 1];
        if (cur.type === 'close' && next.type === 'open' && cur.tag === next.tag) {
          // Find the matching opener for cur (the close tag) by scanning backwards
          // through the open-tag stack to find the full open tag value
          let depth = 0;
          let matchingOpen = null;
          for (let j = i - 1; j >= 0; j--) {
            if (tokens[j].type === 'close' && tokens[j].tag === cur.tag) depth++;
            if (tokens[j].type === 'open' && tokens[j].tag === cur.tag) {
              if (depth === 0) { matchingOpen = tokens[j]; break; }
              depth--;
            }
          }
          // Only merge if the reopened tag has the same full value as the original opener
          if (matchingOpen && matchingOpen.value === next.value) {
            tokens.splice(i, 2); // Remove the close and open pair
            changed = true;
            break; // Restart scan
          }
        }
      }
    }

    return tokens.map(t => t.value).join('');
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

  static _walkElements(xmlStr, parts, footnoteMap, commentMap) {
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
        DexDecompiler._walkElements(innerXml, innerParts, footnoteMap, commentMap);
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
        DexDecompiler._walkElements(innerXml, innerParts, footnoteMap, commentMap);
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
          if (fnId > 1) {
            const fnText = footnoteMap.get(fnId) || '';
            parts.push('{footnote id:' + fnId + '}' + fnText + '{/footnote}');
          }
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
        const elemRe = /<w:t[^>]*>([^<]*)<\/w:t>|<w:tab\s*\/>|<w:br\s*\/?>|<w:br\s+w:type="([^"]*)"[^>]*\/?>|<w:sym\s+[^>]*w:char="([^"]*)"[^>]*\/?>|<w:softHyphen\s*\/>|<w:noBreakHyphen\s*\/?>/g;
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
          } else if (tMatch[0].startsWith('<w:softHyphen')) {
            texts.push('\u00AD'); // Unicode soft hyphen
          } else if (tMatch[0].startsWith('<w:noBreakHyphen')) {
            texts.push('\u2011'); // Unicode non-breaking hyphen
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
        DexDecompiler._walkElements(hlXml.slice(hlXml.indexOf('>') + 1, hlXml.lastIndexOf('<')), linkParts, footnoteMap, commentMap);
        const linkText = linkParts.join('');
        if (rIdMatch) {
          parts.push('{link rId:' + rIdMatch[1] + '}' + linkText + '{/link}');
        } else if (anchorMatch) {
          parts.push('{link anchor:' + DexDecompiler._dexStr(anchorMatch[1]) + '}' + linkText + '{/link}');
        } else {
          parts.push(linkText); // fallback: just the text
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:fldSimple', pos)) {
        // Simple field code: <w:fldSimple w:instr="PAGE">display</w:fldSimple>
        const endTag = '</w:fldSimple>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const fldXml = xmlStr.slice(pos, endIdx + endTag.length);
        const instrMatch = fldXml.match(/w:instr="([^"]*)"/);
        const fldParts = [];
        DexDecompiler._walkElements(fldXml.slice(fldXml.indexOf('>') + 1, fldXml.lastIndexOf('<')), fldParts, footnoteMap, commentMap);
        const displayText = fldParts.join('');
        const instrText = instrMatch ? instrMatch[1].trim() : '';
        parts.push('{field ' + DexDecompiler._dexStr(instrText) + '}' + displayText + '{/field}');
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:sdt>', pos) || xmlStr.startsWith('<w:sdt ', pos)) {
        // Content control: extract inner content
        const endTag = '</w:sdt>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const sdtXml = xmlStr.slice(pos, endIdx + endTag.length);
        // Extract tag name if present
        const tagMatch = sdtXml.match(/<w:tag\s+w:val="([^"]*)"/);
        const aliasMatch = sdtXml.match(/<w:alias\s+w:val="([^"]*)"/);
        const sdtName = aliasMatch ? aliasMatch[1] : (tagMatch ? tagMatch[1] : '');
        // Extract content from sdtContent
        const contentMatch = sdtXml.match(/<w:sdtContent>([\s\S]*)<\/w:sdtContent>/);
        if (contentMatch) {
          if (sdtName) parts.push('{sdt ' + DexDecompiler._dexStr(sdtName) + '}');
          DexDecompiler._walkElements(contentMatch[1], parts, footnoteMap, commentMap);
          if (sdtName) parts.push('{/sdt}');
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:commentRangeStart', pos)) {
        const closeAngle = xmlStr.indexOf('>', pos);
        const tag = xmlStr.slice(pos, closeAngle + 1);
        const idMatch = tag.match(/w:id="(\d+)"/);
        if (idMatch) {
          const cId = parseInt(idMatch[1], 10);
          const comment = commentMap ? commentMap.get(cId) : null;
          if (comment && !comment.isReply) {
            parts.push('{comment-start id:' + idMatch[1] + ' by:' + DexDecompiler._dexStr(comment.author) + ' date:' + DexDecompiler._dexStr(comment.date) + '}');
          } else {
            parts.push('{comment-start id:' + idMatch[1] + '}');
          }
        }
        pos = closeAngle + 1;
      } else if (xmlStr.startsWith('<w:commentRangeEnd', pos)) {
        const closeAngle = xmlStr.indexOf('>', pos);
        const tag = xmlStr.slice(pos, closeAngle + 1);
        const idMatch = tag.match(/w:id="(\d+)"/);
        if (idMatch) {
          const cId = parseInt(idMatch[1], 10);
          const comment = commentMap ? commentMap.get(cId) : null;
          if (comment && !comment.isReply) {
            parts.push('{comment-end id:' + idMatch[1] + ' | ' + comment.text + '}');
          } else {
            parts.push('{comment-end id:' + idMatch[1] + '}');
          }
        }
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
      } else if (xmlStr.startsWith('<m:oMath', pos) || xmlStr.startsWith('<m:oMathPara', pos)) {
        // Math equations: preserve raw XML via base64 and extract text representation
        const isParaWrap = xmlStr.startsWith('<m:oMathPara', pos);
        const endTag = isParaWrap ? '</m:oMathPara>' : '</m:oMath>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const mathXml = xmlStr.slice(pos, endIdx + endTag.length);
        // Extract text representation from m:t elements
        const textParts = [];
        const mtRe = /<m:t[^>]*>([^<]*)<\/m:t>/g;
        let mt;
        while ((mt = mtRe.exec(mathXml)) !== null) textParts.push(mt[1]);
        const textRepr = textParts.join('');
        // Encode the raw XML for preservation (base64 to avoid .dex syntax conflicts)
        const rawB64 = Buffer.from(mathXml).toString('base64');
        parts.push('{math data:' + rawB64 + '}' + textRepr + '{/math}');
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:txbxContent', pos)) {
        // Text box content: recursively decompile inner paragraphs
        const endTag = '</w:txbxContent>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const openClose = xmlStr.indexOf('>', pos);
        const tbXml = xmlStr.slice(openClose + 1, endIdx);
        const tbParts = [];
        DexDecompiler._walkElements(tbXml, tbParts, footnoteMap, commentMap);
        parts.push('{textbox}' + tbParts.join('') + '{/textbox}');
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<mc:AlternateContent', pos)) {
        // Alternate content: prefer Choice, fallback to Fallback
        const endTag = '</mc:AlternateContent>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const acXml = xmlStr.slice(pos, endIdx + endTag.length);
        const choiceMatch = acXml.match(/<mc:Choice[^>]*>([\s\S]*?)<\/mc:Choice>/);
        if (choiceMatch) {
          DexDecompiler._walkElements(choiceMatch[1], parts, footnoteMap, commentMap);
        }
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:ruby', pos)) {
        // Ruby text (phonetic guides)
        const endTag = '</w:ruby>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const rubyXml = xmlStr.slice(pos, endIdx + endTag.length);
        const baseMatch = rubyXml.match(/<w:rubyBase>([\s\S]*?)<\/w:rubyBase>/);
        const rtMatch = rubyXml.match(/<w:rt>([\s\S]*?)<\/w:rt>/);
        const baseText = baseMatch ? xml.extractTextDecoded(baseMatch[1]) : '';
        const rtText = rtMatch ? xml.extractTextDecoded(rtMatch[1]) : '';
        parts.push('{ruby base:' + DexDecompiler._dexStr(baseText) + '}' + rtText + '{/ruby}');
        pos = endIdx + endTag.length;
      } else if (xmlStr.startsWith('<w:object', pos)) {
        // Embedded OLE objects
        const endTag = '</w:object>';
        const endIdx = xmlStr.indexOf(endTag, pos);
        if (endIdx === -1) { pos++; continue; }
        const objXml = xmlStr.slice(pos, endIdx + endTag.length);
        const progIdMatch = objXml.match(/ProgID="([^"]*)"/i) || objXml.match(/o:ProgID="([^"]*)"/i);
        const progId = progIdMatch ? progIdMatch[1] : 'unknown';
        parts.push('{object type:' + DexDecompiler._dexStr(progId) + '}');
        pos = endIdx + endTag.length;
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
      // Skip font wrapper if it matches the document's default font
      const isDefault = fmt.font === DexDecompiler._defaultFont && !fmt.fontHAnsi && !fmt.fontCs && !fmt.fontEastAsia;
      if (!isDefault) {
        let fontAttrs = '"' + fmt.font + '"';
        if (fmt.fontHAnsi) fontAttrs += ' hAnsi:"' + fmt.fontHAnsi + '"';
        if (fmt.fontCs) fontAttrs += ' cs:"' + fmt.fontCs + '"';
        if (fmt.fontEastAsia) fontAttrs += ' eastAsia:"' + fmt.fontEastAsia + '"';
        result = '{font ' + fontAttrs + '}' + result + '{/font}';
      }
    }
    if (fmt.fmtchange) {
      let fmtAttrs = '';
      if (fmt.fmtchange.author) fmtAttrs += ' by:' + DexDecompiler._dexStr(fmt.fmtchange.author);
      if (fmt.fmtchange.date) fmtAttrs += ' date:' + DexDecompiler._dexStr(fmt.fmtchange.date);
      result = '{fmtchange' + fmtAttrs + '}' + result + '{/fmtchange}';
    }
    return result;
  }

  /**
   * Pre-process field codes in paragraph XML.
   * Replaces fldChar begin..separate..end sequences with inline {field} markers.
   * Field codes span multiple <w:r> elements:
   *   <w:r><w:fldChar w:fldCharType="begin"/></w:r>
   *   <w:r><w:instrText> HYPERLINK "url" </w:instrText></w:r>
   *   <w:r><w:fldChar w:fldCharType="separate"/></w:r>
   *   <w:r><w:t>display text</w:t></w:r>
   *   <w:r><w:fldChar w:fldCharType="end"/></w:r>
   */
  static _preprocessFieldCodes(xmlStr) {
    // Quick check: if no fldChar, return unchanged
    if (!xmlStr.includes('w:fldChar')) return xmlStr;

    let result = '';
    let pos = 0;
    let inField = false;
    let fieldInstr = '';
    let fieldDepth = 0;
    const len = xmlStr.length;

    while (pos < len) {
      const fldIdx = xmlStr.indexOf('w:fldChar', pos);
      if (fldIdx === -1) {
        result += xmlStr.slice(pos);
        break;
      }
      // Find the containing run
      const runStartBefore = xmlStr.lastIndexOf('<w:r', fldIdx);
      // Copy everything up to this run
      if (runStartBefore > pos) {
        const chunk = xmlStr.slice(pos, runStartBefore);
        if (inField && fieldDepth === 1) {
          // Inside a field — extract instrText from this chunk
          const instrRe = /<w:instrText[^>]*>([^<]*)<\/w:instrText>/g;
          let im;
          while ((im = instrRe.exec(chunk)) !== null) fieldInstr += im[1];
        }
        if (!inField || fieldDepth > 1) result += chunk;
      }
      // Find the type
      const typeMatch = xmlStr.slice(fldIdx, fldIdx + 60).match(/w:fldCharType="([^"]*)"/);
      if (!typeMatch) {
        result += xmlStr.slice(pos, fldIdx + 20);
        pos = fldIdx + 20;
        continue;
      }
      // Find end of this run
      const runEnd = xmlStr.indexOf('</w:r>', fldIdx);
      if (runEnd === -1) { result += xmlStr.slice(pos); break; }
      const afterRun = runEnd + 6;

      if (typeMatch[1] === 'begin') {
        fieldDepth++;
        if (fieldDepth === 1) {
          inField = true;
          fieldInstr = '';
        }
        pos = afterRun;
      } else if (typeMatch[1] === 'separate') {
        if (fieldDepth === 1) {
          // Emit field-start marker with instruction
          const instrTrimmed = fieldInstr.trim().replace(/\s+/g, ' ');
          result += '<w:r><w:t xml:space="preserve">{field ' + DexDecompiler._dexStr(instrTrimmed) + '}</w:t></w:r>';
        }
        pos = afterRun;
      } else if (typeMatch[1] === 'end') {
        if (fieldDepth === 1) {
          result += '<w:r><w:t xml:space="preserve">{/field}</w:t></w:r>';
          inField = false;
        }
        fieldDepth = Math.max(0, fieldDepth - 1);
        pos = afterRun;
      } else {
        pos = afterRun;
      }
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
      const text = DexDecompiler._extractFormattedText(body);
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
    // Extract dimensions from wp:extent (cx/cy in EMU), convert to inches
    const cxMatch = drawXml.match(/\bcx="(\d+)"/);
    const cyMatch = drawXml.match(/\bcy="(\d+)"/);
    const EMU_PER_INCH = 914400;
    const width = cxMatch ? (parseInt(cxMatch[1], 10) / EMU_PER_INCH).toFixed(2) + 'in' : '';
    const height = cyMatch ? (parseInt(cyMatch[1], 10) / EMU_PER_INCH).toFixed(2) + 'in' : '';
    const rIdMatch = drawXml.match(/r:embed="([^"]+)"/);
    const rId = rIdMatch ? rIdMatch[1] : '';
    let src = '';
    if (rId) {
      try {
        const relsXml = ws.relsXml;
        const relRe = new RegExp('Id="' + rId + '"[^>]*Target="([^"]+)"', 'g');
        const relMatch = relRe.exec(relsXml);
        if (relMatch) {
          src = relMatch[1].startsWith('media/') ? 'word/' + relMatch[1] : 'word/' + relMatch[1].replace(/^\.\.\//, '');
        }
      } catch (_) {}
    }
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

    // Table-level properties
    const styleMatch = tblXml.match(/<w:tblStyle\s+w:val="([^"]+)"/);
    const style = styleMatch ? styleMatch[1] : 'plain';
    let tblAttrs = 'style:' + style + ' cols:' + colCount;

    // Table width
    const tblWMatch = tblXml.match(/<w:tblW\s+[^>]*w:w="(\d+)"/);
    if (tblWMatch) tblAttrs += ' width:' + tblWMatch[1];

    // Table alignment
    const tblPrMatch = tblXml.match(/<w:tblPr>([\s\S]*?)<\/w:tblPr>/);
    if (tblPrMatch) {
      const tblJcMatch = tblPrMatch[1].match(/<w:jc\s+w:val="([^"]+)"/);
      if (tblJcMatch && tblJcMatch[1] !== 'left') tblAttrs += ' align:' + tblJcMatch[1];
    }

    parts.push('{table ' + tblAttrs + '}');
    const rowRe = /<w:tr[\s>][\s\S]*?<\/w:tr>/g;
    let rm;
    const rows = [];
    while ((rm = rowRe.exec(tblXml)) !== null) {
      const rowXml = rm[0]; const cells = [];
      const cellRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
      let cm;
      while ((cm = cellRe.exec(rowXml)) !== null) {
        const cellXml = cm[0];
        // Extract cell properties
        const tcPrMatch = cellXml.match(/<w:tcPr>([\s\S]*?)<\/w:tcPr>/);
        const tcProps = DexDecompiler._extractCellProperties(tcPrMatch ? tcPrMatch[1] : '');

        // Extract cell content using _walkElements for inline formatting
        const cellContent = DexDecompiler._extractCellContent(cellXml);

        // Build {tc ...} prefix if cell has non-default properties
        const tcPrefix = DexDecompiler._formatCellProps(tcProps);
        cells.push(tcPrefix + cellContent);
      }
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

  /**
   * Extract cell-level properties from <w:tcPr> content.
   */
  static _extractCellProperties(tcPrContent) {
    const props = {};
    if (!tcPrContent) return props;

    // Grid span (column merge)
    const spanMatch = tcPrContent.match(/<w:gridSpan\s+w:val="(\d+)"/);
    if (spanMatch && spanMatch[1] !== '1') props.span = spanMatch[1];

    // Vertical merge
    const vMergeMatch = tcPrContent.match(/<w:vMerge(\s+w:val="([^"]*)")?/);
    if (vMergeMatch) {
      props.vmerge = vMergeMatch[2] === 'restart' ? 'restart' : 'continue';
    }

    // Shading (background color)
    const shdMatch = tcPrContent.match(/<w:shd\s+[^>]*w:fill="([^"]+)"/);
    if (shdMatch && shdMatch[1] !== 'auto' && shdMatch[1] !== 'FFFFFF') props.shd = shdMatch[1];

    // Cell width
    const tcWMatch = tcPrContent.match(/<w:tcW\s+[^>]*w:w="(\d+)"/);
    if (tcWMatch) props.width = tcWMatch[1];

    // Vertical alignment
    const vAlignMatch = tcPrContent.match(/<w:vAlign\s+w:val="([^"]+)"/);
    if (vAlignMatch && vAlignMatch[1] !== 'top') props.valign = vAlignMatch[1];

    // Borders
    const bordersMatch = tcPrContent.match(/<w:tcBorders>([\s\S]*?)<\/w:tcBorders>/);
    if (bordersMatch) {
      const bordersXml = bordersMatch[1];
      const borderParts = [];
      for (const side of ['top', 'bottom', 'left', 'right']) {
        const sideRe = new RegExp('<w:' + side + '\\s+([^>]+)');
        const sideMatch = sideRe.exec(bordersXml);
        if (sideMatch) {
          const valM = sideMatch[1].match(/w:val="([^"]+)"/);
          const szM = sideMatch[1].match(/w:sz="([^"]+)"/);
          const colorM = sideMatch[1].match(/w:color="([^"]+)"/);
          if (valM && valM[1] !== 'nil') {
            let b = side + ':' + valM[1];
            if (szM) b += ':' + szM[1];
            if (colorM && colorM[1] !== 'auto') b += ':' + colorM[1];
            borderParts.push(b);
          }
        }
      }
      if (borderParts.length > 0) props.borders = borderParts.join(',');
    }

    return props;
  }

  /**
   * Format cell properties as a {tc ...} prefix string.
   * Returns empty string if no non-default properties.
   */
  static _formatCellProps(props) {
    if (!props || Object.keys(props).length === 0) return '';
    const parts = [];
    if (props.shd) parts.push('shd:' + props.shd);
    if (props.span) parts.push('span:' + props.span);
    if (props.vmerge) parts.push('vmerge:' + props.vmerge);
    if (props.width) parts.push('width:' + props.width);
    if (props.valign) parts.push('valign:' + props.valign);
    if (props.borders) parts.push('borders:"' + props.borders + '"');
    return '{tc ' + parts.join(' ') + '} ';
  }

  /**
   * Extract cell content using _walkElements for inline formatting preservation.
   */
  static _extractCellContent(cellXml) {
    // Remove tcPr to get just the content paragraphs
    const withoutTcPr = cellXml.replace(/<w:tcPr>[\s\S]*?<\/w:tcPr>/g, '');
    // Find all paragraphs in the cell
    const paraRe = /<w:p[\s>][\s\S]*?<\/w:p>/g;
    let pm;
    const paraTexts = [];
    while ((pm = paraRe.exec(withoutTcPr)) !== null) {
      const pXml = pm[0];
      const result = DexDecompiler._decompileRuns(pXml, new Map());
      if (result.text.trim()) paraTexts.push(result.text.trim());
    }
    return paraTexts.join('{br}');
  }

  static _extractSectionProperties(docXml) {
    const allSectPrs = [];
    const sectRe = /<w:sectPr[\s>][\s\S]*?<\/w:sectPr>/g;
    let sm;
    while ((sm = sectRe.exec(docXml)) !== null) allSectPrs.push(sm[0]);
    if (allSectPrs.length === 0) return null;
    return DexDecompiler._formatSectionFromXml(allSectPrs[allSectPrs.length - 1]);
  }

  static _formatSectionFromXml(sectPr) {
    const attrs = [];
    const typeMatch = sectPr.match(/<w:type\s+w:val="([^"]+)"/);
    if (typeMatch) attrs.push('type:' + typeMatch[1]);
    const pgSzMatch = sectPr.match(/<w:pgSz\s+([^>]+)/);
    if (pgSzMatch) {
      const a = pgSzMatch[1];
      const wM = a.match(/w:w="(\d+)"/); const hM = a.match(/w:h="(\d+)"/);
      const oM = a.match(/w:orient="([^"]+)"/);
      if (oM) attrs.push('orient:' + oM[1]);
      if (wM) attrs.push('pgw:' + wM[1]);
      if (hM) attrs.push('pgh:' + hM[1]);
    }
    const pgMar = sectPr.match(/<w:pgMar\s+([^>]+)/);
    if (pgMar) {
      const a = pgMar[1];
      const top = xml.attrVal(a, 'w:top') || '0';
      const right = xml.attrVal(a, 'w:right') || '0';
      const bottom = xml.attrVal(a, 'w:bottom') || '0';
      const left = xml.attrVal(a, 'w:left') || '0';
      const hdr = xml.attrVal(a, 'w:header') || '';
      const ftr = xml.attrVal(a, 'w:footer') || '';
      const gut = xml.attrVal(a, 'w:gutter') || '';
      let margins = top + ',' + right + ',' + bottom + ',' + left;
      if (hdr || ftr || gut) margins += ',' + hdr + ',' + ftr + ',' + gut;
      attrs.push('margins:"' + margins + '"');
    }
    const colsMatch = sectPr.match(/<w:cols\s+([^>]+)/);
    if (colsMatch) {
      const nM = colsMatch[1].match(/w:num="(\d+)"/);
      const sM = colsMatch[1].match(/w:space="(\d+)"/);
      if (nM && nM[1] !== '1') {
        attrs.push('cols:' + nM[1]);
        if (sM) attrs.push('colspace:' + sM[1]);
      }
    }
    const pgNumMatch = sectPr.match(/<w:pgNumType\s+([^>]+)/);
    if (pgNumMatch) {
      const sM = pgNumMatch[1].match(/w:start="(\d+)"/);
      if (sM) attrs.push('pgstart:' + sM[1]);
    }
    const hfRe = /<w:(header|footer)Reference\s+([^>]+)/g;
    let hfM;
    while ((hfM = hfRe.exec(sectPr)) !== null) {
      const tag = hfM[1]; const ha = hfM[2];
      const tM = ha.match(/w:type="([^"]+)"/);
      const rM = ha.match(/r:id="([^"]+)"/);
      if (tM && rM) attrs.push(tag + '-' + tM[1] + ':' + rM[1]);
    }
    if (attrs.length === 0) return null;
    return '\n{section ' + attrs.join(' ') + '}';
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

// Default font set during decompilation; initialized to common default
DexDecompiler._defaultFont = null;

module.exports = { DexDecompiler, listFiles };
