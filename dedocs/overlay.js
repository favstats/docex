'use strict';

const { extractDocumentParagraphs, normalizeGuideText } = require('./guide');

function countLiteral(text, needle) {
  if (needle === '') {
    throw new Error('replace-text find payload cannot be empty');
  }
  return text.split(needle).length - 1;
}

function replaceLiteral(text, needle, replacement) {
  return text.split(needle).join(replacement);
}

function clonePart(part) {
  return {
    ...part,
    buffer: Buffer.from(part.buffer),
  };
}

function escapeXmlText(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeXmlAttr(value) {
  return escapeXmlText(value)
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildParagraphRuns(text) {
  const normalized = String(text == null ? '' : text);
  const lines = normalized.split('\n');
  if (lines.length === 0) {
    return '<w:r><w:t xml:space="preserve"></w:t></w:r>';
  }

  const chunks = [];
  for (let i = 0; i < lines.length; i += 1) {
    chunks.push(`<w:r><w:t xml:space="preserve">${escapeXmlText(lines[i])}</w:t></w:r>`);
    if (i < lines.length - 1) {
      chunks.push('<w:r><w:br/></w:r>');
    }
  }
  return chunks.join('');
}

function extractParagraphProperties(paragraphXml) {
  const match = /<w:pPr\b[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>/.exec(paragraphXml || '');
  return match ? match[0] : '';
}

function withParagraphStyle(pPrXml, style) {
  if (!style) return pPrXml;

  const styleXml = `<w:pStyle w:val="${escapeXmlAttr(style)}"/>`;
  if (!pPrXml) {
    return `<w:pPr>${styleXml}</w:pPr>`;
  }

  if (/^<w:pPr\b[^>]*\/>$/.test(pPrXml)) {
    const attrs = /^<w:pPr\b([^>]*)\/>$/.exec(pPrXml);
    const suffix = attrs ? attrs[1] : '';
    return `<w:pPr${suffix}>${styleXml}</w:pPr>`;
  }

  if (/<w:pStyle\b/.test(pPrXml)) {
    return pPrXml.replace(/<w:pStyle\b[^>]*\/>|<w:pStyle\b[\s\S]*?<\/w:pStyle>/, styleXml);
  }

  return pPrXml.replace(/<w:pPr\b([^>]*)>/, `<w:pPr$1>${styleXml}`);
}

function buildParagraphXml(sourceParagraphXml, text, style, opts = {}) {
  const preserveProperties = opts.preserveProperties !== false;
  const openTagMatch = /^<w:p\b[^>]*>/.exec(sourceParagraphXml || '');
  const openTag = openTagMatch ? openTagMatch[0] : '<w:p>';

  let pPrXml = preserveProperties ? extractParagraphProperties(sourceParagraphXml) : '';
  if (style) {
    pPrXml = withParagraphStyle(pPrXml, style);
  }

  return `${openTag}${pPrXml}${buildParagraphRuns(text)}</w:p>`;
}

function parseParagraphIndex(rawIndex) {
  const value = Number(rawIndex);
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`Invalid paragraph index: ${rawIndex}`);
  }
  return value;
}

function assertParagraphExpectation(transform, paragraph) {
  if (transform.expectedText != null && transform.expectedText !== '') {
    const expected = normalizeGuideText(transform.expectedText);
    if (paragraph.text !== expected) {
      throw new Error(
        `Paragraph ${transform.index} text mismatch. Expected "${expected}" but found "${paragraph.text}"`
      );
    }
  }

  if (transform.expectedStyle != null && transform.expectedStyle !== '') {
    if (paragraph.style !== transform.expectedStyle) {
      throw new Error(
        `Paragraph ${transform.index} style mismatch. Expected "${transform.expectedStyle}" but found "${paragraph.style}"`
      );
    }
  }
}

function documentPartPath(transform) {
  return transform.part || 'word/document.xml';
}

function getUtf8Part(pkg, partPath) {
  const part = pkg.parts.find(candidate => candidate.path === partPath);
  if (!part) {
    throw new Error(`Transform target part not found: ${partPath}`);
  }
  if (part.encoding !== 'utf8') {
    throw new Error(`Transform target must be utf8: ${partPath}`);
  }
  return part;
}

function findParagraph(documentXml, rawIndex) {
  const index = parseParagraphIndex(rawIndex);
  const paragraphs = extractDocumentParagraphs(documentXml);
  const paragraph = paragraphs.find(candidate => candidate.index === index);
  if (!paragraph) {
    throw new Error(`Paragraph ${rawIndex} not found in word/document.xml`);
  }
  return { paragraph, paragraphs };
}

function applyReplaceText(part, transform) {
  const currentText = part.buffer.toString('utf8');
  const matches = countLiteral(currentText, transform.find);
  const expected = transform.count === '' || transform.count == null
    ? null
    : Number(transform.count);

  if (expected !== null && (!Number.isInteger(expected) || expected < 0)) {
    throw new Error(`Invalid replace-text count for ${transform.part}: ${transform.count}`);
  }
  if (expected !== null && matches !== expected) {
    throw new Error(`replace-text expected ${expected} matches in ${transform.part} but found ${matches}`);
  }
  if (matches === 0) {
    throw new Error(`replace-text found no matches in ${transform.part}`);
  }

  const nextText = replaceLiteral(currentText, transform.find, transform.replace);
  part.buffer = Buffer.from(nextText, 'utf8');
  part.text = nextText;
}

function applyReplaceParagraph(part, transform) {
  const documentXml = part.buffer.toString('utf8');
  const { paragraph } = findParagraph(documentXml, transform.index);
  assertParagraphExpectation(transform, paragraph);

  const replacementXml = buildParagraphXml(
    paragraph.xml,
    transform.text || '',
    transform.style || '',
    { preserveProperties: true }
  );

  const nextDocumentXml = documentXml.slice(0, paragraph.start)
    + replacementXml
    + documentXml.slice(paragraph.end);

  part.buffer = Buffer.from(nextDocumentXml, 'utf8');
  part.text = nextDocumentXml;
}

function applyInsertParagraph(part, transform, where) {
  const documentXml = part.buffer.toString('utf8');
  const { paragraph } = findParagraph(documentXml, transform.index);
  assertParagraphExpectation(transform, paragraph);

  const paragraphXml = buildParagraphXml('', transform.text || '', transform.style || '', {
    preserveProperties: false,
  });

  const insertionPoint = where === 'before' ? paragraph.start : paragraph.end;
  const nextDocumentXml = documentXml.slice(0, insertionPoint)
    + paragraphXml
    + documentXml.slice(insertionPoint);

  part.buffer = Buffer.from(nextDocumentXml, 'utf8');
  part.text = nextDocumentXml;
}

function applyDeleteParagraph(part, transform) {
  const documentXml = part.buffer.toString('utf8');
  const { paragraph, paragraphs } = findParagraph(documentXml, transform.index);
  assertParagraphExpectation(transform, paragraph);

  if (paragraphs.length <= 1) {
    throw new Error('Cannot delete the only paragraph in word/document.xml');
  }

  const nextDocumentXml = documentXml.slice(0, paragraph.start) + documentXml.slice(paragraph.end);
  part.buffer = Buffer.from(nextDocumentXml, 'utf8');
  part.text = nextDocumentXml;
}

function applyTransform(nextPkg, transform) {
  if (!transform || typeof transform !== 'object') {
    throw new Error('Transform must be an object');
  }

  if (transform.type === 'replace-text') {
    applyReplaceText(getUtf8Part(nextPkg, transform.part), transform);
    return;
  }

  if (transform.type === 'replace-paragraph') {
    applyReplaceParagraph(getUtf8Part(nextPkg, documentPartPath(transform)), transform);
    return;
  }

  if (transform.type === 'insert-paragraph-after') {
    applyInsertParagraph(getUtf8Part(nextPkg, documentPartPath(transform)), transform, 'after');
    return;
  }

  if (transform.type === 'insert-paragraph-before') {
    applyInsertParagraph(getUtf8Part(nextPkg, documentPartPath(transform)), transform, 'before');
    return;
  }

  if (transform.type === 'delete-paragraph') {
    applyDeleteParagraph(getUtf8Part(nextPkg, documentPartPath(transform)), transform);
    return;
  }

  throw new Error(`Unsupported transform type: ${transform.type}`);
}

function applyTransforms(pkg) {
  const transforms = Array.isArray(pkg.transforms) ? pkg.transforms : [];
  if (transforms.length === 0) return pkg;

  const nextPkg = {
    ...pkg,
    parts: pkg.parts.map(clonePart),
  };

  for (const transform of transforms) {
    applyTransform(nextPkg, transform);
  }

  return nextPkg;
}

module.exports = {
  applyTransforms,
};
