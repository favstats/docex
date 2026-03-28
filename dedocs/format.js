'use strict';

const crypto = require('crypto');

const DEDOCS_VERSION = '1';
const TOP_LEVEL_COMMAND = '\\dedocs';
const GUIDE_COMMAND = '\\guide';
const END_GUIDE_COMMAND = '\\end{guide}';
const PART_COMMAND = '\\part';
const END_PART_COMMAND = '\\end{part}';
const REPLACE_TEXT_COMMAND = '\\replace-text';
const END_REPLACE_TEXT_COMMAND = '\\end{replace-text}';
const END_DOC_COMMAND = '\\end{dedocs}';

function sha256(buffer) {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

function escapeAttr(value) {
  return String(value)
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t');
}

function unescapeAttr(value) {
  return String(value).replace(/\\(["\\nrt])/g, (_, ch) => {
    if (ch === 'n') return '\n';
    if (ch === 'r') return '\r';
    if (ch === 't') return '\t';
    return ch;
  });
}

function formatAttrs(attrs) {
  return Object.entries(attrs)
    .filter(([, value]) => value !== undefined && value !== null)
    .map(([key, value]) => `${key}="${escapeAttr(value)}"`)
    .join(', ');
}

function parseAttrs(source) {
  const attrs = {};
  let index = 0;

  function skipWs() {
    while (index < source.length && /\s/.test(source[index])) index += 1;
  }

  while (index < source.length) {
    skipWs();
    if (index >= source.length) break;

    const keyMatch = /^[A-Za-z_][A-Za-z0-9_-]*/.exec(source.slice(index));
    if (!keyMatch) {
      throw new Error(`Invalid attribute list near: ${source.slice(index, index + 40)}`);
    }
    const key = keyMatch[0];
    index += key.length;
    skipWs();

    if (source[index] !== '=') {
      throw new Error(`Expected '=' after attribute ${key}`);
    }
    index += 1;
    skipWs();

    if (source[index] !== '"') {
      throw new Error(`Expected quoted value for attribute ${key}`);
    }
    index += 1;

    let raw = '';
    while (index < source.length) {
      const ch = source[index];
      if (ch === '\\') {
        if (index + 1 >= source.length) {
          throw new Error(`Unterminated escape in attribute ${key}`);
        }
        raw += source.slice(index, index + 2);
        index += 2;
        continue;
      }
      if (ch === '"') break;
      raw += ch;
      index += 1;
    }

    if (source[index] !== '"') {
      throw new Error(`Unterminated quoted value for attribute ${key}`);
    }
    index += 1;
    attrs[key] = unescapeAttr(raw);

    skipWs();
    if (index >= source.length) break;
    if (source[index] !== ',') {
      throw new Error(`Expected ',' after attribute ${key}`);
    }
    index += 1;
  }

  return attrs;
}

function parseCommandLine(line, command) {
  const trimmed = line.trim();
  if (!trimmed.startsWith(command)) {
    throw new Error(`Expected ${command}, got: ${line}`);
  }

  const open = trimmed.indexOf('[');
  const close = trimmed.lastIndexOf(']');
  if (open === -1 || close === -1 || close < open) {
    throw new Error(`Malformed command line: ${line}`);
  }

  return parseAttrs(trimmed.slice(open + 1, close));
}

function splitLinesWithOffsets(text) {
  const lines = [];
  let offset = 0;

  while (offset <= text.length) {
    const newline = text.indexOf('\n', offset);
    if (newline === -1) {
      lines.push({ line: text.slice(offset), start: offset, end: text.length });
      break;
    }
    lines.push({ line: text.slice(offset, newline), start: offset, end: newline + 1 });
    offset = newline + 1;
  }

  return lines;
}

function makeBoundary(payload, label) {
  const prefix = `:::DEDOCS_${label}_${sha256(Buffer.from(payload, 'utf8')).slice(0, 12)}`;
  let attempt = 0;

  while (true) {
    const candidate = attempt === 0 ? `${prefix}:::` : `${prefix}_${attempt}:::`;
    const fullLine = new RegExp(`(^|\\n)${candidate.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\n|$)`);
    if (!fullLine.test(payload)) return candidate;
    attempt += 1;
  }
}

function isBuffer(value) {
  return Buffer.isBuffer(value);
}

function normalizePart(part) {
  if (!part || typeof part !== 'object') {
    throw new Error('Each part must be an object');
  }
  if (!part.path) {
    throw new Error('Each part must have a path');
  }

  let buffer;
  if (isBuffer(part.buffer)) {
    buffer = part.buffer;
  } else if (typeof part.text === 'string') {
    buffer = Buffer.from(part.text, 'utf8');
  } else {
    throw new Error(`Part ${part.path} must provide buffer or text`);
  }

  const encoding = part.encoding || (part.kind === 'base64' ? 'base64' : 'utf8');

  return {
    path: String(part.path),
    mediaType: part.mediaType || 'application/octet-stream',
    encoding,
    buffer,
    text: encoding === 'base64' ? buffer.toString('base64') : buffer.toString('utf8'),
  };
}

function wrapBase64(base64) {
  if (base64.length === 0) return '';
  return base64.match(/.{1,76}/g).join('\n');
}

function splitPayload(raw) {
  return raw.endsWith('\n') ? raw.slice(0, -1) : raw;
}

function readBoundaryBlock(lines, text, lineIndex, boundary, label) {
  if (lineIndex >= lines.length || lines[lineIndex].line !== boundary) {
    throw new Error(`${label} expected boundary line ${boundary}`);
  }

  const payloadStart = lines[lineIndex].end;
  lineIndex += 1;

  let closingBoundaryLine = -1;
  while (lineIndex < lines.length) {
    if (lines[lineIndex].line === boundary) {
      closingBoundaryLine = lineIndex;
      break;
    }
    lineIndex += 1;
  }

  if (closingBoundaryLine === -1) {
    throw new Error(`${label} is missing closing boundary ${boundary}`);
  }

  const payloadEnd = lines[closingBoundaryLine].start;
  return {
    payload: splitPayload(text.slice(payloadStart, payloadEnd)),
    nextLineIndex: closingBoundaryLine + 1,
  };
}

function readLabelBlock(lines, text, lineIndex, openLabel, closeLabel, label) {
  if (lineIndex >= lines.length || lines[lineIndex].line.trim() !== openLabel) {
    throw new Error(`${label} expected ${openLabel}`);
  }

  const payloadStart = lines[lineIndex].end;
  lineIndex += 1;

  let closingLine = -1;
  while (lineIndex < lines.length) {
    if (lines[lineIndex].line.trim() === closeLabel) {
      closingLine = lineIndex;
      break;
    }
    lineIndex += 1;
  }

  if (closingLine === -1) {
    throw new Error(`${label} is missing ${closeLabel}`);
  }

  const payloadEnd = lines[closingLine].start;
  return {
    payload: splitPayload(text.slice(payloadStart, payloadEnd)),
    nextLineIndex: closingLine + 1,
  };
}

function serializePackage(pkg) {
  if (!pkg || typeof pkg !== 'object') {
    throw new Error('Package must be an object');
  }
  if (!Array.isArray(pkg.parts)) {
    throw new Error('Package must include a parts array');
  }

  const header = {
    version: pkg.version || DEDOCS_VERSION,
    package: pkg.package || 'docx',
    fidelity: pkg.fidelity || 'package-exact',
    source: pkg.source || '',
  };

  const chunks = [`${TOP_LEVEL_COMMAND}[${formatAttrs(header)}]\n`];
  const guides = Array.isArray(pkg.guides) ? pkg.guides : [];
  const transforms = Array.isArray(pkg.transforms) ? pkg.transforms : [];

  guides.forEach((guide, index) => {
    const payload = typeof guide.text === 'string' ? guide.text : '';
    const boundary = makeBoundary(payload, `GUIDE_${index + 1}`);
    const attrs = {
      name: guide.name || `guide-${index + 1}`,
      part: guide.part || '',
      format: guide.format || 'text',
      boundary,
    };

    chunks.push('\n');
    chunks.push(`${GUIDE_COMMAND}[${formatAttrs(attrs)}]\n`);
    chunks.push(`${boundary}\n`);
    chunks.push(payload);
    if (!payload.endsWith('\n')) chunks.push('\n');
    chunks.push(`${boundary}\n`);
    chunks.push(`${END_GUIDE_COMMAND}\n`);
  });

  transforms.forEach(transform => {
    if (!transform || transform.type !== 'replace-text') {
      throw new Error(`Unsupported transform type: ${transform && transform.type}`);
    }

    const attrs = {
      part: transform.part,
      count: transform.count == null ? '' : String(transform.count),
    };

    chunks.push('\n');
    chunks.push(`${REPLACE_TEXT_COMMAND}[${formatAttrs(attrs)}]\n`);
    chunks.push('<<<FIND\n');
    chunks.push(transform.find || '');
    if (!(transform.find || '').endsWith('\n')) chunks.push('\n');
    chunks.push('FIND\n');
    chunks.push('<<<WITH\n');
    chunks.push(transform.replace || '');
    if (!(transform.replace || '').endsWith('\n')) chunks.push('\n');
    chunks.push('WITH\n');
    chunks.push(`${END_REPLACE_TEXT_COMMAND}\n`);
  });

  pkg.parts.forEach((rawPart, index) => {
    const part = normalizePart(rawPart);
    const payload = part.encoding === 'base64' ? wrapBase64(part.text) : part.text;
    const boundary = makeBoundary(payload, `PART_${index + 1}`);
    const attrs = {
      path: part.path,
      mediaType: part.mediaType,
      encoding: part.encoding,
      bytes: String(part.buffer.length),
      sha256: sha256(part.buffer),
      boundary,
    };

    chunks.push('\n');
    chunks.push(`${PART_COMMAND}[${formatAttrs(attrs)}]\n`);
    chunks.push(`${boundary}\n`);
    chunks.push(payload);
    if (!payload.endsWith('\n')) chunks.push('\n');
    chunks.push(`${boundary}\n`);
    chunks.push(`${END_PART_COMMAND}\n`);
  });

  chunks.push(`\n${END_DOC_COMMAND}\n`);
  return chunks.join('');
}

function parsePackage(text, opts = {}) {
  if (typeof text !== 'string') {
    throw new Error('Dedocs source must be a string');
  }

  const lines = splitLinesWithOffsets(text);
  let lineIndex = 0;

  function skipNoise() {
    while (lineIndex < lines.length) {
      const trimmed = lines[lineIndex].line.trim();
      if (trimmed === '' || trimmed.startsWith('#')) {
        lineIndex += 1;
        continue;
      }
      break;
    }
  }

  skipNoise();
  if (lineIndex >= lines.length) {
    throw new Error('Empty dedocs source');
  }

  const headerAttrs = parseCommandLine(lines[lineIndex].line, TOP_LEVEL_COMMAND);
  lineIndex += 1;

  const guides = [];
  const transforms = [];
  const parts = [];

  while (lineIndex < lines.length) {
    skipNoise();
    if (lineIndex >= lines.length) break;

    const line = lines[lineIndex].line.trim();
    if (line === END_DOC_COMMAND) {
      return {
        version: headerAttrs.version || DEDOCS_VERSION,
        package: headerAttrs.package || 'docx',
        fidelity: headerAttrs.fidelity || 'package-exact',
        source: headerAttrs.source || '',
        guides,
        transforms,
        parts,
      };
    }

    if (line.startsWith(GUIDE_COMMAND)) {
      const guideAttrs = parseCommandLine(lines[lineIndex].line, GUIDE_COMMAND);
      lineIndex += 1;

      if (!guideAttrs.boundary) {
        throw new Error(`Guide ${guideAttrs.name || ''} is missing boundary`);
      }

      const block = readBoundaryBlock(lines, text, lineIndex, guideAttrs.boundary, `Guide ${guideAttrs.name || ''}`);
      lineIndex = block.nextLineIndex;

      if (lineIndex >= lines.length || lines[lineIndex].line.trim() !== END_GUIDE_COMMAND) {
        throw new Error(`Guide ${guideAttrs.name || ''} is missing ${END_GUIDE_COMMAND}`);
      }
      lineIndex += 1;

      guides.push({
        name: guideAttrs.name || `guide-${guides.length + 1}`,
        part: guideAttrs.part || '',
        format: guideAttrs.format || 'text',
        text: block.payload,
      });
      continue;
    }

    if (line.startsWith(REPLACE_TEXT_COMMAND)) {
      const transformAttrs = parseCommandLine(lines[lineIndex].line, REPLACE_TEXT_COMMAND);
      lineIndex += 1;

      const findBlock = readLabelBlock(lines, text, lineIndex, '<<<FIND', 'FIND', 'replace-text');
      lineIndex = findBlock.nextLineIndex;
      const replaceBlock = readLabelBlock(lines, text, lineIndex, '<<<WITH', 'WITH', 'replace-text');
      lineIndex = replaceBlock.nextLineIndex;

      if (lineIndex >= lines.length || lines[lineIndex].line.trim() !== END_REPLACE_TEXT_COMMAND) {
        throw new Error(`replace-text is missing ${END_REPLACE_TEXT_COMMAND}`);
      }
      lineIndex += 1;

      transforms.push({
        type: 'replace-text',
        part: transformAttrs.part || '',
        count: transformAttrs.count === '' ? '' : transformAttrs.count,
        find: findBlock.payload,
        replace: replaceBlock.payload,
      });
      continue;
    }

    const partAttrs = parseCommandLine(lines[lineIndex].line, PART_COMMAND);
    lineIndex += 1;

    const boundary = partAttrs.boundary;
    if (!boundary) {
      throw new Error(`Part ${partAttrs.path} is missing boundary attribute`);
    }

    const block = readBoundaryBlock(lines, text, lineIndex, boundary, `Part ${partAttrs.path}`);
    const payloadRaw = block.payload;
    lineIndex = block.nextLineIndex;

    if (lineIndex >= lines.length || lines[lineIndex].line.trim() !== END_PART_COMMAND) {
      throw new Error(`Part ${partAttrs.path} is missing ${END_PART_COMMAND}`);
    }
    lineIndex += 1;

    const declaredBytes = partAttrs.bytes || '';
    const declaredSha256 = partAttrs.sha256 || '';
    const encoding = partAttrs.encoding || 'utf8';

    const candidates = [payloadRaw];
    if (payloadRaw.endsWith('\n')) {
      candidates.push(payloadRaw.slice(0, -1));
    }

    let chosenPayload = payloadRaw;
    let buffer = null;

    for (const candidate of candidates) {
      const candidateBuffer = encoding === 'base64'
        ? Buffer.from(candidate.replace(/\s+/g, ''), 'base64')
        : Buffer.from(candidate, 'utf8');
      const candidateBytes = String(candidateBuffer.length);
      const candidateSha256 = sha256(candidateBuffer);
      const bytesMatch = !declaredBytes || declaredBytes === candidateBytes;
      const hashMatch = !declaredSha256 || declaredSha256 === candidateSha256;
      if (bytesMatch && hashMatch) {
        chosenPayload = candidate;
        buffer = candidateBuffer;
        break;
      }
    }

    if (buffer === null) {
      chosenPayload = payloadRaw;
      buffer = encoding === 'base64'
        ? Buffer.from(payloadRaw.replace(/\s+/g, ''), 'base64')
        : Buffer.from(payloadRaw, 'utf8');
    }

    const actualBytes = String(buffer.length);
    const actualSha256 = sha256(buffer);

    if (opts.strictMetadata) {
      if (declaredBytes && declaredBytes !== actualBytes) {
        throw new Error(`Part ${partAttrs.path} declares ${declaredBytes} bytes but contains ${actualBytes}`);
      }
      if (declaredSha256 && declaredSha256 !== actualSha256) {
        throw new Error(`Part ${partAttrs.path} declares sha256 ${declaredSha256} but contains ${actualSha256}`);
      }
    }

    parts.push({
      path: partAttrs.path,
      mediaType: partAttrs.mediaType || 'application/octet-stream',
      encoding,
      buffer,
      text: encoding === 'base64' ? buffer.toString('base64') : chosenPayload,
      declaredBytes,
      declaredSha256,
      actualBytes,
      actualSha256,
    });
  }

  throw new Error(`Missing ${END_DOC_COMMAND}`);
}

module.exports = {
  DEDOCS_VERSION,
  END_DOC_COMMAND,
  END_GUIDE_COMMAND,
  END_PART_COMMAND,
  END_REPLACE_TEXT_COMMAND,
  GUIDE_COMMAND,
  PART_COMMAND,
  REPLACE_TEXT_COMMAND,
  TOP_LEVEL_COMMAND,
  formatAttrs,
  makeBoundary,
  parseAttrs,
  parsePackage,
  serializePackage,
  sha256,
  wrapBase64,
};
