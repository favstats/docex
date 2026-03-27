/**
 * dex-parser.js -- Parse .dex human-readable format into AST
 * Zero external dependencies (YAML parsing hand-rolled for our subset).
 */
'use strict';

class DexParser {
  static parse(dexString) {
    const trimmed = (dexString || '').trimStart();
    if (trimmed.startsWith('\\dex[') || trimmed.startsWith('\\dex{')) {
      const { parseDex } = require('./dex-lossless');
      return parseDex(dexString);
    }
    const { frontmatter, body: bodyStr } = DexParser._extractFrontmatter(dexString);
    const body = DexParser._parseBody(bodyStr);
    return { frontmatter, body };
  }

  static _extractFrontmatter(dexString) {
    const trimmed = dexString.trim();
    if (!trimmed.startsWith('---')) return { frontmatter: {}, body: trimmed };
    const endMarker = trimmed.indexOf('\n---', 3);
    if (endMarker === -1) return { frontmatter: {}, body: trimmed };
    const yamlStr = trimmed.slice(4, endMarker).trim();
    const body = trimmed.slice(endMarker + 4).trim();
    return { frontmatter: DexParser._parseYaml(yamlStr), body };
  }

  static _parseYaml(yamlStr) {
    const result = {};
    const lines = yamlStr.split('\n');
    let i = 0;
    while (i < lines.length) {
      const trimLine = lines[i].trim();
      if (!trimLine || trimLine.startsWith('#')) { i++; continue; }
      const kvMatch = trimLine.match(/^(\w[\w-]*)\s*:\s*(.*)/);
      if (!kvMatch) { i++; continue; }
      const key = kvMatch[1];
      let value = kvMatch[2].trim();
      if (value === '') {
        const items = [];
        i++;
        while (i < lines.length) {
          const nextLine = lines[i]; const nextTrimmed = nextLine.trim();
          if (!nextTrimmed) { i++; continue; }
          if (!nextLine.match(/^\s/) && !nextTrimmed.startsWith('-')) break;
          if (nextTrimmed.startsWith('- ')) {
            const itemValue = nextTrimmed.slice(2).trim();
            const itemKv = itemValue.match(/^(\w[\w-]*)\s*:\s*(.*)/);
            if (itemKv) {
              const obj = {}; obj[itemKv[1]] = DexParser._unquote(itemKv[2].trim());
              i++;
              while (i < lines.length) {
                const deepLine = lines[i]; const deepTrimmed = deepLine.trim();
                if (!deepTrimmed) { i++; continue; }
                if (!deepLine.match(/^\s{4,}/) && !deepTrimmed.match(/^\w[\w-]*\s*:/)) break;
                const deepKv = deepTrimmed.match(/^(\w[\w-]*)\s*:\s*(.*)/);
                if (deepKv) { obj[deepKv[1]] = DexParser._unquote(deepKv[2].trim()); i++; } else break;
              }
              items.push(obj);
            } else { items.push(DexParser._unquote(itemValue)); i++; }
          } else break;
        }
        result[key] = items.length > 0 ? items : '';
      } else { result[key] = DexParser._unquote(value); i++; }
    }
    return result;
  }

  static _unquote(s) {
    if ((s.startsWith('"') && s.endsWith('"')) || (s.startsWith("'") && s.endsWith("'")))
      return s.slice(1, -1).replace(/\\"/g, '"').replace(/\\\\/g, '\\');
    return s;
  }

  static _parseBody(bodyStr) {
    const nodes = []; const lines = bodyStr.split('\n'); let i = 0;
    while (i < lines.length) {
      const trimmed = lines[i].trim();
      if (trimmed === '') { i++; continue; }
      const headingMatch = trimmed.match(/^(#{1,6})\s+(.*?)(?:\s+\{id:([A-Fa-f0-9]+)\})?$/);
      if (headingMatch) { nodes.push({ type: 'heading', level: headingMatch[1].length, text: headingMatch[2].trim(), id: headingMatch[3] || null }); i++; continue; }
      if (trimmed === '{pagebreak}') { nodes.push({ type: 'pagebreak' }); i++; continue; }
      const sectionMatch = trimmed.match(/^\{section\s+(.*)\}$/);
      if (sectionMatch) { const attrs = DexParser._parseBlockAttrs(sectionMatch[1]); nodes.push({ type: 'section', margins: attrs.margins || '', header: attrs.header || '', footer: attrs.footer || '' }); i++; continue; }
      if (trimmed.startsWith('{figure')) { const r = DexParser._parseFigureBlock(lines, i); nodes.push(r.node); i = r.endLine + 1; continue; }
      if (trimmed.startsWith('{table')) { const r = DexParser._parseTableBlock(lines, i); nodes.push(r.node); i = r.endLine + 1; continue; }
      if (trimmed.startsWith('{comment')) { const r = DexParser._parseCommentBlock(lines, i); nodes.push(r.node); i = r.endLine + 1; continue; }
      if (trimmed.startsWith('{reply')) { const r = DexParser._parseReplyBlock(lines, i); nodes.push(r.node); i = r.endLine + 1; continue; }
      if (trimmed.startsWith('{p')) {
        const pMatch = trimmed.match(/^\{p(?:\s+id:([A-Fa-f0-9]+))?\}$/);
        if (pMatch) {
          const paraLines = []; const paraId = pMatch[1] || null; i++;
          while (i < lines.length && lines[i].trim() !== '{/p}') { paraLines.push(lines[i]); i++; }
          nodes.push({ type: 'paragraph', id: paraId, runs: DexParser._parseInlineContent(paraLines.join('\n')) });
          if (i < lines.length) i++; continue;
        }
      }
      const implicitLines = [];
      while (i < lines.length) {
        const l = lines[i].trim();
        if (l === '' || l.startsWith('#') || l.startsWith('{p') || l.startsWith('{figure') || l.startsWith('{table') || l.startsWith('{comment') || l.startsWith('{reply') || l === '{pagebreak}' || l.startsWith('{section')) break;
        implicitLines.push(lines[i]); i++;
      }
      if (implicitLines.length > 0) nodes.push({ type: 'paragraph', id: null, runs: DexParser._parseInlineContent(implicitLines.join('\n')) });
    }
    return nodes;
  }

  static _parseFigureBlock(lines, startLine) {
    const firstLine = lines[startLine].trim();
    const attrsStr = firstLine.slice('{figure'.length, firstLine.length - 1).trim();
    const attrs = DexParser._parseBlockAttrs(attrsStr);
    const contentLines = []; let i = startLine + 1;
    while (i < lines.length && lines[i].trim() !== '{/figure}') { contentLines.push(lines[i]); i++; }
    return { node: { type: 'figure', id: attrs.id || null, rId: attrs.rId || null, src: attrs.src || null, width: attrs.width || null, height: attrs.height || null, alt: attrs.alt || null, caption: contentLines.join('\n').trim() }, endLine: i };
  }

  static _parseTableBlock(lines, startLine) {
    const firstLine = lines[startLine].trim();
    const attrsStr = firstLine.slice('{table'.length, firstLine.length - 1).trim();
    const attrs = DexParser._parseBlockAttrs(attrsStr);
    const contentLines = []; let captionText = ''; let i = startLine + 1;
    while (i < lines.length && lines[i].trim() !== '{/table}') {
      const trimLine = lines[i].trim();
      const capMatch = trimLine.match(/^\{caption\}(.*)\{\/caption\}$/);
      if (capMatch) captionText = capMatch[1].trim(); else contentLines.push(trimLine);
      i++;
    }
    const rows = [];
    for (const line of contentLines) {
      if (!line.startsWith('|')) continue;
      if (/^\|[-|]+\|$/.test(line.replace(/\s/g, ''))) continue;
      rows.push(line.split('|').slice(1, -1).map(c => c.trim()));
    }
    return { node: { type: 'table', id: attrs.id || null, style: attrs.style || 'plain', cols: parseInt(attrs.cols, 10) || (rows.length > 0 ? rows[0].length : 0), caption: captionText, rows }, endLine: i };
  }

  static _parseCommentBlock(lines, startLine) {
    const firstLine = lines[startLine].trim();
    const attrsStr = firstLine.slice('{comment'.length, firstLine.length - 1).trim();
    const attrs = DexParser._parseBlockAttrs(attrsStr);
    const contentLines = []; let i = startLine + 1;
    while (i < lines.length && lines[i].trim() !== '{/comment}') { contentLines.push(lines[i]); i++; }
    return { node: { type: 'comment', id: parseInt(attrs.id, 10) || 0, author: attrs.by || '', date: attrs.date || '', anchor: attrs.anchor || '', text: contentLines.join('\n').trim() }, endLine: i };
  }

  static _parseReplyBlock(lines, startLine) {
    const firstLine = lines[startLine].trim();
    const attrsStr = firstLine.slice('{reply'.length, firstLine.length - 1).trim();
    const attrs = DexParser._parseBlockAttrs(attrsStr);
    const contentLines = []; let i = startLine + 1;
    while (i < lines.length && lines[i].trim() !== '{/reply}') { contentLines.push(lines[i]); i++; }
    return { node: { type: 'reply', id: parseInt(attrs.id, 10) || 0, parent: parseInt(attrs.parent, 10) || 0, author: attrs.by || '', date: attrs.date || '', text: contentLines.join('\n').trim() }, endLine: i };
  }

  static _parseInlineContent(content) {
    const runs = []; let pos = 0; const len = content.length;
    while (pos < len) {
      const braceIdx = content.indexOf('{', pos);
      if (braceIdx === -1) { const text = content.slice(pos); if (text) runs.push({ type: 'text', text }); break; }
      if (braceIdx > pos) { const text = content.slice(pos, braceIdx); if (text) runs.push({ type: 'text', text }); }
      const tagResult = DexParser._parseInlineTag(content, braceIdx);
      if (tagResult) { runs.push(tagResult.node); pos = tagResult.endPos; }
      else { runs.push({ type: 'text', text: '{' }); pos = braceIdx + 1; }
    }
    return runs;
  }

  static _parseInlineTag(content, pos) {
    const tagMap = [
      { prefix: '{b}', closeTag: '{/b}', type: 'bold' },
      { prefix: '{i}', closeTag: '{/i}', type: 'italic' },
      { prefix: '{u}', closeTag: '{/u}', type: 'underline' },
      { prefix: '{sup}', closeTag: '{/sup}', type: 'superscript' },
      { prefix: '{sub}', closeTag: '{/sub}', type: 'subscript' },
    ];
    for (const tag of tagMap) {
      if (content.startsWith(tag.prefix, pos)) {
        const closeIdx = content.indexOf(tag.closeTag, pos + tag.prefix.length);
        if (closeIdx === -1) return null;
        return { node: { type: tag.type, text: content.slice(pos + tag.prefix.length, closeIdx) }, endPos: closeIdx + tag.closeTag.length };
      }
    }
    if (content.startsWith('{del ', pos)) return DexParser._parseAttributedTag(content, pos, 'del', '{/del}', (attrs, inner) => ({ type: 'del', id: parseInt(attrs.id, 10) || 0, author: attrs.by || '', date: attrs.date || '', text: inner }));
    if (content.startsWith('{ins ', pos)) return DexParser._parseAttributedTag(content, pos, 'ins', '{/ins}', (attrs, inner) => ({ type: 'ins', id: parseInt(attrs.id, 10) || 0, author: attrs.by || '', date: attrs.date || '', text: inner }));
    if (content.startsWith('{cite ', pos)) return DexParser._parseAttributedTag(content, pos, 'cite', '{/cite}', (attrs, inner) => ({ type: 'cite', key: attrs.key || '', citeType: attrs.type || 'parenthetical', text: inner }));
    if (content.startsWith('{footnote ', pos)) return DexParser._parseAttributedTag(content, pos, 'footnote', '{/footnote}', (attrs, inner) => ({ type: 'footnote', id: parseInt(attrs.id, 10) || 0, text: inner }));
    if (content.startsWith('{highlight ', pos)) {
      const closeAngle = content.indexOf('}', pos); if (closeAngle === -1) return null;
      const attrStr = content.slice(pos + '{highlight '.length, closeAngle).trim();
      const closeIdx = content.indexOf('{/highlight}', closeAngle + 1); if (closeIdx === -1) return null;
      return { node: { type: 'highlight', color: attrStr, text: content.slice(closeAngle + 1, closeIdx) }, endPos: closeIdx + '{/highlight}'.length };
    }
    if (content.startsWith('{color ', pos)) {
      const closeAngle = content.indexOf('}', pos); if (closeAngle === -1) return null;
      const colorVal = content.slice(pos + '{color '.length, closeAngle).trim();
      const closeIdx = content.indexOf('{/color}', closeAngle + 1); if (closeIdx === -1) return null;
      return { node: { type: 'color', color: colorVal, text: content.slice(closeAngle + 1, closeIdx) }, endPos: closeIdx + '{/color}'.length };
    }
    if (content.startsWith('{font ', pos)) {
      const closeAngle = content.indexOf('}', pos); if (closeAngle === -1) return null;
      let fontName = content.slice(pos + '{font '.length, closeAngle).trim();
      fontName = DexParser._unquote(fontName);
      const closeIdx = content.indexOf('{/font}', closeAngle + 1); if (closeIdx === -1) return null;
      return { node: { type: 'font', font: fontName, text: content.slice(closeAngle + 1, closeIdx) }, endPos: closeIdx + '{/font}'.length };
    }
    return null;
  }

  static _parseAttributedTag(content, pos, tagName, closeTag, nodeBuilder) {
    const openEnd = content.indexOf('}', pos); if (openEnd === -1) return null;
    const attrStr = content.slice(pos + tagName.length + 2, openEnd);
    const attrs = DexParser._parseBlockAttrs(attrStr);
    const closeIdx = content.indexOf(closeTag, openEnd + 1); if (closeIdx === -1) return null;
    return { node: nodeBuilder(attrs, content.slice(openEnd + 1, closeIdx)), endPos: closeIdx + closeTag.length };
  }

  static _parseBlockAttrs(attrStr) {
    const attrs = {}; if (!attrStr) return attrs;
    const re = /(\w[\w-]*):(?:"((?:[^"\\]|\\.)*)"|(\S+))/g;
    let m;
    while ((m = re.exec(attrStr)) !== null) {
      const key = m[1];
      const value = m[2] !== undefined ? m[2].replace(/\\"/g, '"').replace(/\\\\/g, '\\') : m[3];
      attrs[key] = value;
    }
    return attrs;
  }
}

module.exports = { DexParser };
