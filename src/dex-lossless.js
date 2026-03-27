'use strict';

const DEX_VERSION = '0.5.0';

const CONTROL_COMMANDS = new Set([
  'dex',
  'part',
  'xml',
  'text',
  'comment',
  'cdata',
  'pi',
]);

const ROOT_XML_COMMANDS = new Set([
  'Types',
  'Default',
  'Override',
  'Relationships',
  'Relationship',
]);

function isXmlishPart(relPath) {
  return /\.(xml|rels)$/i.test(relPath);
}

function isWhitespaceOnly(value) {
  return /^[\s]*$/.test(value);
}

function isElementCommandName(name) {
  if (CONTROL_COMMANDS.has(name)) return false;
  if (ROOT_XML_COMMANDS.has(name)) return true;
  return /^(?:[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*|[A-Za-z_][A-Za-z0-9_.-]*)$/.test(name);
}

function formatAttrs(attrs) {
  if (!attrs || (attrs.length === 0 && !(attrs.trailing || ''))) return '';
  let out = '[';
  for (let i = 0; i < attrs.length; i++) {
    const attr = attrs[i];
    const prefix = attr.prefix !== undefined ? attr.prefix : (i === 0 ? '' : ' ');
    out += prefix + `${attr.name}="${escapeDexInline(String(attr.value))}"`;
  }
  out += attrs.trailing || '';
  out += ']';
  return out;
}

function escapeDexInline(value) {
  return String(value)
    .replace(/\\/g, '\\\\')
    .replace(/\{/g, '\\{')
    .replace(/\}/g, '\\}')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t')
    .replace(/"/g, '\\"');
}

function unescapeDexInline(value) {
  let out = '';
  for (let i = 0; i < value.length; i++) {
    const ch = value[i];
    if (ch !== '\\') {
      out += ch;
      continue;
    }
    i++;
    if (i >= value.length) throw new Error('dangling escape at end of inline content');
    const next = value[i];
    if (next === 'n') out += '\n';
    else if (next === 'r') out += '\r';
    else if (next === 't') out += '\t';
    else if (next === '{') out += '{';
    else if (next === '}') out += '}';
    else if (next === '\\') out += '\\';
    else if (next === '"') out += '"';
    else out += next;
  }
  return out;
}

function parseXml(xmlString) {
  const nodes = [];
  const stack = [];
  let pos = 0;

  function pushNode(node) {
    const parent = stack[stack.length - 1];
    if (parent) parent.children.push(node);
    else nodes.push(node);
  }

  while (pos < xmlString.length) {
    if (xmlString.startsWith('<?xml', pos)) {
      const end = xmlString.indexOf('?>', pos);
      if (end === -1) throw new Error('unterminated XML declaration');
      const inner = xmlString.slice(pos + 5, end).trim();
      pushNode({
        type: 'declaration',
        attrs: parseXmlAttrs(inner),
      });
      pos = end + 2;
      continue;
    }

    if (xmlString.startsWith('<?', pos)) {
      const end = xmlString.indexOf('?>', pos);
      if (end === -1) throw new Error('unterminated processing instruction');
      const inner = xmlString.slice(pos + 2, end).trim();
      const space = inner.search(/\s/);
      const target = space === -1 ? inner : inner.slice(0, space);
      const value = space === -1 ? '' : inner.slice(space + 1);
      pushNode({
        type: 'pi',
        target,
        value,
      });
      pos = end + 2;
      continue;
    }

    if (xmlString.startsWith('<!--', pos)) {
      const end = xmlString.indexOf('-->', pos);
      if (end === -1) throw new Error('unterminated XML comment');
      pushNode({
        type: 'comment',
        value: xmlString.slice(pos + 4, end),
      });
      pos = end + 3;
      continue;
    }

    if (xmlString.startsWith('<![CDATA[', pos)) {
      const end = xmlString.indexOf(']]>', pos);
      if (end === -1) throw new Error('unterminated CDATA');
      pushNode({
        type: 'cdata',
        value: xmlString.slice(pos + 9, end),
      });
      pos = end + 3;
      continue;
    }

    if (xmlString.startsWith('</', pos)) {
      const end = xmlString.indexOf('>', pos);
      if (end === -1) throw new Error('unterminated closing tag');
      const name = xmlString.slice(pos + 2, end).trim();
      const node = stack.pop();
      if (!node) throw new Error(`unexpected closing tag </${name}>`);
      if (node.name !== name) {
        throw new Error(`mismatched closing tag </${name}> for <${node.name}>`);
      }
      pos = end + 1;
      continue;
    }

    if (xmlString[pos] === '<') {
      const end = findTagEnd(xmlString, pos);
      const raw = xmlString.slice(pos + 1, end);
      const selfClosing = /\/\s*$/.test(raw);
      const inner = selfClosing ? raw.replace(/\/\s*$/, '').trim() : raw.trim();
      const nameMatch = inner.match(/^([^\s/>]+)/);
      if (!nameMatch) throw new Error('missing XML element name');
      const name = nameMatch[1];
      const attrSource = inner.slice(name.length).trim();
      const node = {
        type: 'element',
        name,
        attrs: parseXmlAttrs(attrSource),
        children: [],
        selfClosing,
      };
      pushNode(node);
      if (!selfClosing) stack.push(node);
      pos = end + 1;
      continue;
    }

    const next = xmlString.indexOf('<', pos);
    const end = next === -1 ? xmlString.length : next;
    pushNode({
      type: 'text',
      value: xmlString.slice(pos, end),
    });
    pos = end;
  }

  if (stack.length > 0) {
    throw new Error(`unclosed XML element <${stack[stack.length - 1].name}>`);
  }

  return nodes;
}

function serializeXml(nodes) {
  return nodes.map(serializeXmlNode).join('');
}

function serializeXmlNode(node) {
  switch (node.type) {
    case 'declaration':
      return '<?xml' + serializeXmlAttrs(node.attrs) + '?>';
    case 'pi':
      return node.value ? `<?${node.target} ${node.value}?>` : `<?${node.target}?>`;
    case 'comment':
      return `<!--${node.value}-->`;
    case 'cdata':
      return `<![CDATA[${node.value}]]>`;
    case 'text':
      return node.value;
    case 'element': {
      const attrs = serializeXmlAttrs(node.attrs);
      if (node.selfClosing) return `<${node.name}${attrs}/>`;
      return `<${node.name}${attrs}>${serializeXml(node.children || [])}</${node.name}>`;
    }
    default:
      throw new Error(`cannot serialize XML node type ${node.type}`);
  }
}

function serializeDex(ast) {
  if (!ast || ast.type !== 'dex') throw new Error('expected dex package AST');
  const lines = [];
  lines.push(`\\dex[version="${escapeDexInline(ast.version || DEX_VERSION)}"]{`);
  for (const part of ast.parts || []) {
    const attrs = simpleAttrs([
      { name: 'path', value: part.path },
      { name: 'type', value: part.partType || 'xml' },
    ]);
    if (part.partType === 'binary' && part.encoding) {
      attrs.push({ name: 'encoding', value: part.encoding });
    }
    if (part.partType === 'binary') {
      lines.push(`  \\part${formatAttrs(attrs)}{${escapeDexInline(part.data || '')}}`);
      continue;
    }
    lines.push(`  \\part${formatAttrs(attrs)}{`);
    for (const node of part.nodes || []) {
      appendDexNode(lines, node, '    ');
    }
    lines.push('  \\end{part}');
  }
  lines.push('\\end{dex}');
  return lines.join('\n') + '\n';
}

function appendDexNode(lines, node, indent) {
  switch (node.type) {
    case 'declaration':
      lines.push(`${indent}\\xml${formatAttrs(node.attrs)}{}`);
      return;
    case 'comment':
      lines.push(`${indent}\\comment{${escapeDexInline(node.value)}}`);
      return;
    case 'cdata':
      lines.push(`${indent}\\cdata{${escapeDexInline(node.value)}}`);
      return;
    case 'pi': {
      const attrs = simpleAttrs([{ name: 'target', value: node.target }]);
      lines.push(`${indent}\\pi${formatAttrs(attrs)}{${escapeDexInline(node.value || '')}}`);
      return;
    }
    case 'text':
      lines.push(`${indent}\\text{${escapeDexInline(node.value)}}`);
      return;
    case 'element': {
      const attrStr = formatAttrs(node.attrs);
      if ((node.children || []).length === 0) {
        if (node.selfClosing) {
          lines.push(`${indent}\\${node.name}${attrStr}{}`);
        } else {
          lines.push(`${indent}\\${node.name}${attrStr}{`);
          lines.push(`${indent}\\end{${node.name}}`);
        }
        return;
      }
      if ((node.children || []).length === 1 && node.children[0].type === 'text') {
        lines.push(`${indent}\\${node.name}${attrStr}{${escapeDexInline(node.children[0].value)}}`);
        return;
      }
      lines.push(`${indent}\\${node.name}${attrStr}{`);
      for (const child of node.children || []) appendDexNode(lines, child, indent + '  ');
      lines.push(`${indent}\\end{${node.name}}`);
      return;
    }
    default:
      throw new Error(`cannot serialize dex node type ${node.type}`);
  }
}

function parseDex(dexString) {
  const parser = new DexParserImpl(dexString);
  return parser.parse();
}

class DexParserImpl {
  constructor(source) {
    this.source = source;
    this.pos = 0;
  }

  parse() {
    this._skipWhitespace();
    const root = this._parseCommand();
    this._skipWhitespace();
    if (this.pos !== this.source.length) {
      throw this._error('unexpected trailing content');
    }
    if (root.name !== 'dex') throw this._error('root command must be \\dex');
    if (!root.block) throw this._error('\\dex must use block form');
    const version = attrValue(root.attrs, 'version') || DEX_VERSION;
    const parts = root.children.map(child => this._commandToPart(child));
    return {
      type: 'dex',
      version,
      parts,
    };
  }

  _commandToPart(command) {
    if (command.name !== 'part') {
      throw this._error(`unexpected top-level command \\${command.name}; expected \\part`);
    }
    const partType = attrValue(command.attrs, 'type') || 'xml';
    const partPath = attrValue(command.attrs, 'path');
    if (!partPath) throw this._error('\\part requires path attribute');
    if (partType === 'binary') {
      if (command.block) throw this._error('binary parts must use inline form');
      return {
        type: 'part',
        path: partPath,
        partType: 'binary',
        encoding: attrValue(command.attrs, 'encoding') || 'base64',
        data: unescapeDexInline(command.text),
      };
    }
    if (!command.block) throw this._error('XML parts must use block form');
    return {
      type: 'part',
      path: partPath,
      partType: 'xml',
      nodes: command.children.map(child => this._commandToXmlNode(child)),
    };
  }

  _commandToXmlNode(command) {
    switch (command.name) {
      case 'xml':
        if (command.block) throw this._error('\\xml declaration cannot use block form');
        return { type: 'declaration', attrs: command.attrs };
      case 'text':
        if (command.block) throw this._error('\\text cannot use block form');
        return { type: 'text', value: unescapeDexInline(command.text) };
      case 'comment':
        if (command.block) throw this._error('\\comment cannot use block form');
        return { type: 'comment', value: unescapeDexInline(command.text) };
      case 'cdata':
        if (command.block) throw this._error('\\cdata cannot use block form');
        return { type: 'cdata', value: unescapeDexInline(command.text) };
      case 'pi':
        if (command.block) throw this._error('\\pi cannot use block form');
        return {
          type: 'pi',
          target: attrValue(command.attrs, 'target') || '',
          value: unescapeDexInline(command.text),
        };
      default:
        if (CONTROL_COMMANDS.has(command.name)) {
          throw this._error(`unexpected control command \\${command.name} in XML content`);
        }
        if (!isElementCommandName(command.name)) {
          throw this._error(`unknown element command \\${command.name}`);
        }
        return {
          type: 'element',
          name: command.name,
          attrs: command.attrs,
          children: command.block
            ? command.children.map(child => this._commandToXmlNode(child))
            : (command.text === '' ? [] : [{ type: 'text', value: unescapeDexInline(command.text) }]),
          selfClosing: !command.block && command.text === '',
        };
    }
  }

  _parseCommand() {
    this._skipWhitespace();
    if (this.source[this.pos] !== '\\') throw this._error('expected command');
    this.pos++;

    const nameStart = this.pos;
    while (this.pos < this.source.length && /[A-Za-z0-9:_.?!-]/.test(this.source[this.pos])) {
      this.pos++;
    }
    if (this.pos === nameStart) throw this._error('missing command name');
    const name = this.source.slice(nameStart, this.pos);

      const attrs = this.source[this.pos] === '[' ? this._parseAttrs() : simpleAttrs([]);
    if (this.source[this.pos] !== '{') throw this._error(`expected { after \\${name}`);
    this.pos++;

    const block = this.source[this.pos] === '\n';
    if (block) {
      this.pos++;
      const children = [];
      while (true) {
        this._skipWhitespace();
        if (this.source.startsWith(`\\end{${name}}`, this.pos)) {
          this.pos += `\\end{${name}}`.length;
          break;
        }
        children.push(this._parseCommand());
      }
      return { name, attrs, block: true, children };
    }

    const text = this._readInline();
    return { name, attrs, block: false, text };
  }

  _parseAttrs() {
    const attrs = simpleAttrs([]);
    if (this.source[this.pos] !== '[') throw this._error('expected [');
    this.pos++;
    while (true) {
      let prefix = '';
      while (this.pos < this.source.length && /[\s,]/.test(this.source[this.pos])) {
        if (this.source[this.pos] !== ',') prefix += this.source[this.pos];
        this.pos++;
      }
      if (this.source[this.pos] === ']') {
        attrs.trailing = prefix;
        this.pos++;
        return attrs;
      }
      const nameStart = this.pos;
      while (this.pos < this.source.length && /[A-Za-z0-9:_.-]/.test(this.source[this.pos])) {
        this.pos++;
      }
      if (this.pos === nameStart) throw this._error('expected attribute name');
      const name = this.source.slice(nameStart, this.pos);
      this._skipInlineSpaces();
      if (this.source[this.pos] !== '=') throw this._error(`expected = after attribute ${name}`);
      this.pos++;
      this._skipInlineSpaces();
      if (this.source[this.pos] !== '"') throw this._error(`expected quoted value for attribute ${name}`);
      this.pos++;
      let raw = '';
      while (this.pos < this.source.length) {
        const ch = this.source[this.pos];
        if (ch === '"') {
          this.pos++;
          break;
        }
        if (ch === '\\') {
          raw += ch;
          this.pos++;
          if (this.pos >= this.source.length) throw this._error('unterminated escape in attribute');
          raw += this.source[this.pos];
          this.pos++;
          continue;
        }
        raw += ch;
        this.pos++;
      }
      attrs.push({ name, value: unescapeDexInline(raw), prefix });
    }
  }

  _readInline() {
    let raw = '';
    while (this.pos < this.source.length) {
      const ch = this.source[this.pos];
      if (ch === '}') {
        this.pos++;
        return raw;
      }
      if (ch === '\\') {
        raw += ch;
        this.pos++;
        if (this.pos >= this.source.length) throw this._error('unterminated escape in inline content');
        raw += this.source[this.pos];
        this.pos++;
        continue;
      }
      raw += ch;
      this.pos++;
    }
    throw this._error('unterminated inline command');
  }

  _skipWhitespace() {
    while (this.pos < this.source.length && /\s/.test(this.source[this.pos])) this.pos++;
  }

  _skipInlineSpaces() {
    while (this.pos < this.source.length && /[ \t]/.test(this.source[this.pos])) this.pos++;
  }

  _error(message) {
    return new Error(`dex parse error at offset ${this.pos}: ${message}`);
  }
}

function findTagEnd(xmlString, start) {
  let pos = start + 1;
  let quote = null;
  while (pos < xmlString.length) {
    const ch = xmlString[pos];
    if (quote) {
      if (ch === quote) quote = null;
    } else if (ch === '"' || ch === '\'') {
      quote = ch;
    } else if (ch === '>') {
      return pos;
    }
    pos++;
  }
  throw new Error('unterminated XML tag');
}

function parseXmlAttrs(source) {
  const attrs = simpleAttrs([]);
  let pos = 0;
  while (pos < source.length) {
    let prefix = '';
    while (pos < source.length && /\s/.test(source[pos])) {
      prefix += source[pos];
      pos++;
    }
    if (pos >= source.length) {
      attrs.trailing = prefix;
      break;
    }
    const nameStart = pos;
    while (pos < source.length && /[^\s=]/.test(source[pos])) pos++;
    const name = source.slice(nameStart, pos);
    while (pos < source.length && /\s/.test(source[pos])) pos++;
    if (source[pos] !== '=') {
      throw new Error(`expected = after XML attribute ${name}`);
    }
    pos++;
    while (pos < source.length && /\s/.test(source[pos])) pos++;
    const quote = source[pos];
    if (quote !== '"' && quote !== '\'') {
      throw new Error(`expected quoted value for XML attribute ${name}`);
    }
    pos++;
    const valueStart = pos;
    while (pos < source.length && source[pos] !== quote) pos++;
    if (pos >= source.length) throw new Error(`unterminated XML attribute ${name}`);
    attrs.push({
      name,
      value: source.slice(valueStart, pos),
      prefix,
    });
    pos++;
  }
  return attrs;
}

function serializeXmlAttrs(attrs) {
  if (!attrs || (attrs.length === 0 && !(attrs.trailing || ''))) return '';
  let out = '';
  for (let i = 0; i < attrs.length; i++) {
    const attr = attrs[i];
    const prefix = attr.prefix !== undefined ? attr.prefix : '';
    out += i === 0
      ? (prefix === '' ? ' ' : prefix)
      : (prefix === '' ? ' ' : prefix);
    out += `${attr.name}="${attr.value}"`;
  }
  out += attrs.trailing || '';
  return out;
}

function attrValue(attrs, name) {
  const match = (attrs || []).find(attr => attr.name === name);
  return match ? match.value : null;
}

function simpleAttrs(items) {
  const attrs = Array.isArray(items) ? items.slice() : [];
  attrs.trailing = '';
  return attrs;
}

module.exports = {
  CONTROL_COMMANDS,
  DEX_VERSION,
  escapeDexInline,
  formatAttrs,
  isElementCommandName,
  isWhitespaceOnly,
  isXmlishPart,
  parseDex,
  parseXml,
  serializeDex,
  serializeXml,
  unescapeDexInline,
};
