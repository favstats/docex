'use strict';

function decodeXml(value) {
  return String(value)
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&');
}

function xmlText(fragment) {
  const texts = [];
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
  let match;
  while ((match = re.exec(fragment)) !== null) {
    texts.push(match[1]);
  }
  return decodeXml(texts.join(''));
}

function normalizeGuideText(text) {
  return text
    .replace(/\s+/g, ' ')
    .trim();
}

function extractDocumentParagraphs(documentXml) {
  const paragraphs = [];
  const paragraphRe = /<w:p\b[\s\S]*?<\/w:p>/g;
  let match;
  let index = 0;

  while ((match = paragraphRe.exec(documentXml)) !== null) {
    const paragraphXml = match[0];
    const styleMatch = /<w:pStyle\b[^>]*w:val="([^"]+)"/.exec(paragraphXml);
    const style = styleMatch ? styleMatch[1] : '';
    const text = normalizeGuideText(xmlText(paragraphXml));

    paragraphs.push({
      index,
      start: match.index,
      end: match.index + paragraphXml.length,
      xml: paragraphXml,
      style,
      text,
    });
    index += 1;
  }

  return paragraphs;
}

function documentParagraphGuide(documentXml) {
  if (typeof documentXml !== 'string' || documentXml.length === 0) return '';

  const lines = [];
  for (const paragraph of extractDocumentParagraphs(documentXml)) {
    const { index, style, text } = paragraph;
    if (text || style) {
      const padded = String(index).padStart(4, '0');
      if (style) {
        lines.push(`\\p[index="${padded}", style="${style}"] ${text}`);
      } else {
        lines.push(`\\p[index="${padded}"] ${text}`);
      }
    }
  }

  return lines.join('\n');
}

function createGuides(pkg) {
  if (!pkg || !Array.isArray(pkg.parts)) return [];

  const documentPart = pkg.parts.find(part => part.path === 'word/document.xml' && part.encoding === 'utf8');
  if (!documentPart) return [];

  const guideText = documentParagraphGuide(documentPart.buffer.toString('utf8'));
  if (!guideText) return [];

  return [{
    name: 'document-paragraphs',
    part: 'word/document.xml',
    format: 'paragraphs',
    text: guideText,
  }];
}

module.exports = {
  createGuides,
  decodeXml,
  documentParagraphGuide,
  extractDocumentParagraphs,
  normalizeGuideText,
  xmlText,
};
