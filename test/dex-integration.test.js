'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const { DexDecompiler } = require('../src/dex-decompiler');
const { DexParser } = require('../src/dex-markdown-parser');
const { DexCompiler } = require('../src/dex-compiler');

const OUTPUT_DIR = path.join(__dirname, 'output', 'dex-integration');
const VISUAL_DIR = path.join(__dirname, 'output', 'dex-visual');
const CHAOS_DOCX = '/mnt/storage/absolute_chaos.docx';
const CHAOS_DOCX_V2 = '/mnt/storage/absolute_chaos_v2.docx';
const CHAOS_DOCX_V3 = '/mnt/storage/absolute_chaos_v3.docx';

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
if (!fs.existsSync(VISUAL_DIR)) fs.mkdirSync(VISUAL_DIR, { recursive: true });

describe('.dex integration', { skip: !fs.existsSync(CHAOS_DOCX) }, () => {
  it('decompile -> parse -> compile preserves absolute_chaos exactly', () => {
    const dex = DexDecompiler.decompile(CHAOS_DOCX);
    const ast = DexParser.parse(dex);
    const outPath = path.join(OUTPUT_DIR, 'absolute-chaos-roundtrip.docx');
    const result = DexCompiler.compile(ast, {
      output: outPath,
      verifyAgainst: CHAOS_DOCX,
      strictVerify: false,
    });

    assert.equal(result.verified, true);
    assert.deepEqual(result.differences, []);
  });

  it('survives a second decompile/parse/compile cycle with no drift', () => {
    const first = DexCompiler.assertRoundTrip(CHAOS_DOCX, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-cycle1.docx'),
      strictVerify: false,
    });
    assert.equal(first.verified, true);

    const second = DexCompiler.assertRoundTrip(first.path, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-cycle2.docx'),
      strictVerify: false,
    });
    assert.equal(second.verified, true);

    const compare = DexCompiler.compareDocx(CHAOS_DOCX, second.path);
    assert.equal(compare.equal, true);
    assert.deepEqual(compare.differences, []);
  });

  it('AST-level edit produces a constrained diff against absolute_chaos', () => {
    const ast = DexDecompiler.toAst(CHAOS_DOCX);
    const docPart = ast.parts.find(part => part.path === 'word/document.xml');
    assert.ok(docPart, 'document.xml part should exist');

    const firstTextNode = findFirstTextNode(docPart.nodes);
    assert.ok(firstTextNode, 'should find a text node in document.xml');

    const originalValue = firstTextNode.value;
    firstTextNode.value = originalValue + ' [DEX_INTEGRATION_EDIT]';

    const outPath = path.join(OUTPUT_DIR, 'absolute-chaos-edited.docx');
    DexCompiler.compile(ast, { output: outPath });

    const compare = DexCompiler.compareDocx(CHAOS_DOCX, outPath);
    assert.equal(compare.equal, false);
    assert.ok(compare.differences.length >= 1);
    assert.ok(compare.differences.every(diff => diff.path === 'word/document.xml'));
  });

  it('absolute_chaos_v2 also round-trips exactly', { skip: !fs.existsSync(CHAOS_DOCX_V2) }, () => {
    const result = DexCompiler.assertRoundTrip(CHAOS_DOCX_V2, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-v2-roundtrip.docx'),
      strictVerify: false,
    });
    assert.equal(result.verified, true);
    assert.deepEqual(result.differences, []);
  });

  it('absolute_chaos_v3 also round-trips exactly', { skip: !fs.existsSync(CHAOS_DOCX_V3) }, () => {
    const result = DexCompiler.assertRoundTrip(CHAOS_DOCX_V3, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-v3-roundtrip.docx'),
      strictVerify: false,
    });
    assert.equal(result.verified, true);
    assert.deepEqual(result.differences, []);
  });

  it('absolute_chaos_v3 survives a second cycle with no package drift', {
    skip: !fs.existsSync(CHAOS_DOCX_V3),
  }, () => {
    const first = DexCompiler.assertRoundTrip(CHAOS_DOCX_V3, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-v3-cycle1.docx'),
      strictVerify: false,
    });
    assert.equal(first.verified, true);

    const second = DexCompiler.assertRoundTrip(first.path, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-v3-cycle2.docx'),
      strictVerify: false,
    });
    assert.equal(second.verified, true);

    const compare = DexCompiler.compareDocx(CHAOS_DOCX_V3, second.path);
    assert.equal(compare.equal, true);
    assert.deepEqual(compare.differences, []);
  });

  it('absolute_chaos_v3 preserves the full package signature', {
    skip: !fs.existsSync(CHAOS_DOCX_V3),
  }, () => {
    const roundTrip = DexCompiler.assertRoundTrip(CHAOS_DOCX_V3, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-v3-signature-roundtrip.docx'),
      strictVerify: false,
    });
    assert.equal(roundTrip.verified, true);

    const before = extractPackageSignature(CHAOS_DOCX_V3);
    const after = extractPackageSignature(roundTrip.path);
    assert.deepEqual(after, before);
  });
});

describe('.dex visual regression', { skip: !fs.existsSync(CHAOS_DOCX) }, () => {
  it('round-trip preserves the visual signature of absolute_chaos', () => {
    const roundTrip = DexCompiler.assertRoundTrip(CHAOS_DOCX, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-visual-roundtrip.docx'),
      strictVerify: false,
    });
    assert.equal(roundTrip.verified, true);

    const before = extractVisualSignature(CHAOS_DOCX);
    const after = extractVisualSignature(roundTrip.path);
    assert.deepEqual(after, before);
  });

  it('edited AST changes the visual signature in a targeted way', () => {
    const ast = DexDecompiler.toAst(CHAOS_DOCX);
    const docPart = ast.parts.find(part => part.path === 'word/document.xml');
    const firstTextNode = findFirstTextNode(docPart.nodes);
    const marker = 'DEX_VISUAL_SENTINEL';
    firstTextNode.value = marker + ' ' + firstTextNode.value;

    const outPath = path.join(OUTPUT_DIR, 'absolute-chaos-visual-edited.docx');
    DexCompiler.compile(ast, { output: outPath });

    const before = extractVisualSignature(CHAOS_DOCX);
    const after = extractVisualSignature(outPath);

    assert.notDeepEqual(after, before);
    assert.equal(after.textHash === before.textHash, false);
    assert.equal(after.paragraphCount, before.paragraphCount);
    assert.equal(after.tableCount, before.tableCount);
    assert.equal(after.commentCount, before.commentCount);
  });

  it('renders PDF pages when a docx-to-pdf converter is available', {
    skip: !resolveDocxToPdfCommand(),
  }, () => {
    const roundTrip = DexCompiler.assertRoundTrip(CHAOS_DOCX, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-render-roundtrip.docx'),
      strictVerify: false,
    });
    assert.equal(roundTrip.verified, true);

    const originalPdf = convertDocxToPdf(CHAOS_DOCX, path.join(VISUAL_DIR, 'absolute-chaos-original.pdf'));
    const roundTripPdf = convertDocxToPdf(roundTrip.path, path.join(VISUAL_DIR, 'absolute-chaos-roundtrip.pdf'));

    const originalPages = renderPdfPages(originalPdf, path.join(VISUAL_DIR, 'absolute-chaos-original'));
    const roundTripPages = renderPdfPages(roundTripPdf, path.join(VISUAL_DIR, 'absolute-chaos-roundtrip'));

    assert.ok(originalPages.length > 0, 'original PDF should render at least one page');
    assert.deepEqual(roundTripPages.length, originalPages.length);
  });

  it('round-trip preserves the visual signature of absolute_chaos_v3', {
    skip: !fs.existsSync(CHAOS_DOCX_V3),
  }, () => {
    const roundTrip = DexCompiler.assertRoundTrip(CHAOS_DOCX_V3, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-v3-visual-roundtrip.docx'),
      strictVerify: false,
    });
    assert.equal(roundTrip.verified, true);

    const before = extractVisualSignature(CHAOS_DOCX_V3);
    const after = extractVisualSignature(roundTrip.path);
    assert.deepEqual(after, before);
  });
});

function findFirstTextNode(nodes) {
  for (const node of nodes || []) {
    if (node.type === 'text' && /\S/.test(node.value)) return node;
    if (node.type === 'element') {
      const match = findFirstTextNode(node.children || []);
      if (match) return match;
    }
  }
  return null;
}

function extractVisualSignature(docxPath) {
  const unpacked = unpackDocx(docxPath);
  try {
    const documentXml = readIfExists(path.join(unpacked.rootDir, 'word', 'document.xml'));
    const stylesXml = readIfExists(path.join(unpacked.rootDir, 'word', 'styles.xml'));
    const commentsXml = readIfExists(path.join(unpacked.rootDir, 'word', 'comments.xml'));
    const commentsExtXml = readIfExists(path.join(unpacked.rootDir, 'word', 'commentsExtended.xml'));
    const commentsIdsXml = readIfExists(path.join(unpacked.rootDir, 'word', 'commentsIds.xml'));
    const footnotesXml = readIfExists(path.join(unpacked.rootDir, 'word', 'footnotes.xml'));
    const endnotesXml = readIfExists(path.join(unpacked.rootDir, 'word', 'endnotes.xml'));
    const numberingXml = readIfExists(path.join(unpacked.rootDir, 'word', 'numbering.xml'));
    const customXml = collectPartFamily(unpacked.rootDir, 'customXml', /\.xml$/);
    const headerXml = collectPartFamily(unpacked.rootDir, 'word', /^header\d+\.xml$/);
    const footerXml = collectPartFamily(unpacked.rootDir, 'word', /^footer\d+\.xml$/);

    const visibleText = [
      stripTags(documentXml),
      stripTags(headerXml),
      stripTags(footerXml),
      stripTags(commentsXml),
      stripTags(footnotesXml),
      stripTags(endnotesXml),
      stripTags(customXml),
    ].join('\n');

    return {
      paragraphCount: count(documentXml, /<w:p[\s>]/g),
      runCount: count(documentXml, /<w:r[\s>]/g),
      tableCount: count(documentXml, /<w:tbl[\s>]/g),
      rowCount: count(documentXml, /<w:tr[\s>]/g),
      cellCount: count(documentXml, /<w:tc[\s>]/g),
      commentCount: count(commentsXml, /<w:comment\b/g),
      threadedCommentCount: count(commentsExtXml, /<w15:commentEx\b/g),
      commentIdCount: count(commentsIdsXml, /<w16cid:commentId\b/g),
      footnoteCount: count(footnotesXml, /<w:footnote\b/g),
      endnoteCount: count(endnotesXml, /<w:endnote\b/g),
      headerPartCount: count(headerXml, /<w:hdr\b/g),
      footerPartCount: count(footerXml, /<w:ftr\b/g),
      pageBreakCount: count(documentXml, /<w:br\b[^>]*w:type="page"/g),
      sectionCount: count(documentXml, /<w:sectPr[\s>]/g),
      listLevelCount: count(numberingXml, /<w:lvl[\s>]/g),
      hyperlinkCount: count(documentXml, /<w:hyperlink\b/g),
      sdtCount: count(documentXml, /<w:sdt\b/g),
      customXmlPartCount: count(customXml, /<[^/!?][^>]*>/g),
      headingStyleCount: count(documentXml, /<w:pStyle\b[^>]*w:val="Heading\d+"/g),
      boldCount: count(documentXml, /<w:b\b/g),
      italicCount: count(documentXml, /<w:i\b/g),
      underlineCount: count(documentXml, /<w:u\b/g),
      highlightCount: count(documentXml, /<w:highlight\b/g),
      textHash: sha256(visibleText),
      layoutHash: sha256([
        documentXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)?.join('\n') || '',
        stylesXml,
        numberingXml,
      ].join('\n')),
    };
  } finally {
    unpacked.cleanup();
  }
}

function unpackDocx(docxPath) {
  const tmpDir = fs.mkdtempSync('/tmp/dex-visual-');
  execFileSync('unzip', ['-q', docxPath, '-d', tmpDir], { stdio: 'pipe' });
  return {
    rootDir: tmpDir,
    cleanup() {
      execFileSync('rm', ['-rf', tmpDir], { stdio: 'pipe' });
    },
  };
}

function extractPackageSignature(docxPath) {
  const unpacked = unpackDocx(docxPath);
  try {
    const parts = walkFiles(unpacked.rootDir);
    const hashes = {};
    for (const relPath of parts) {
      const buf = fs.readFileSync(path.join(unpacked.rootDir, relPath));
      hashes[relPath] = sha256(buf);
    }

    return {
      partCount: parts.length,
      xmlPartCount: parts.filter(relPath => /\.(xml|rels)$/i.test(relPath)).length,
      binaryPartCount: parts.filter(relPath => !/\.(xml|rels)$/i.test(relPath)).length,
      customXmlPartCount: parts.filter(relPath => relPath.startsWith('customXml/')).length,
      mediaPartCount: parts.filter(relPath => relPath.startsWith('word/media/')).length,
      relationshipCount: parts.filter(relPath => relPath.endsWith('.rels')).length,
      parts,
      hashes,
    };
  } finally {
    unpacked.cleanup();
  }
}

function readIfExists(filePath) {
  return fs.existsSync(filePath) ? fs.readFileSync(filePath, 'utf8') : '';
}

function collectPartFamily(rootDir, subdir, pattern) {
  const dir = path.join(rootDir, subdir);
  if (!fs.existsSync(dir)) return '';
  return fs.readdirSync(dir)
    .filter(name => pattern.test(name))
    .sort()
    .map(name => fs.readFileSync(path.join(dir, name), 'utf8'))
    .join('\n');
}

function stripTags(xml) {
  return xml
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function count(str, pattern) {
  return (str.match(pattern) || []).length;
}

function sha256(value) {
  return crypto.createHash('sha256').update(value).digest('hex');
}

function walkFiles(rootDir) {
  const out = [];

  function walk(currentDir) {
    const entries = fs.readdirSync(currentDir, { withFileTypes: true })
      .sort((a, b) => a.name.localeCompare(b.name));
    for (const entry of entries) {
      const absPath = path.join(currentDir, entry.name);
      if (entry.isDirectory()) {
        walk(absPath);
        continue;
      }
      out.push(path.relative(rootDir, absPath).replace(/\\/g, '/'));
    }
  }

  walk(rootDir);
  return out;
}

function convertDocxToPdf(docxPath, pdfPath) {
  const cmd = resolveDocxToPdfCommand();
  if (!cmd) throw new Error('DOCEX_DOCX_TO_PDF is not configured');
  execFileSync('bash', ['-lc', `${cmd} "${docxPath}" "${pdfPath}"`], { stdio: 'pipe' });
  assertFile(pdfPath, 'expected PDF to be created');
  return pdfPath;
}

function renderPdfPages(pdfPath, prefix) {
  execFileSync('pdftoppm', ['-png', '-r', '144', pdfPath, prefix], { stdio: 'pipe' });
  const dir = path.dirname(prefix);
  const base = path.basename(prefix);
  return fs.readdirSync(dir)
    .filter(name => name.startsWith(base + '-') && name.endsWith('.png'))
    .sort()
    .map(name => path.join(dir, name));
}

function assertFile(filePath, message) {
  assert.ok(fs.existsSync(filePath), message);
  assert.ok(fs.statSync(filePath).size > 0, message);
}

function resolveDocxToPdfCommand() {
  if (process.env.DOCEX_DOCX_TO_PDF) return process.env.DOCEX_DOCX_TO_PDF;

  const office = detectCommand(['soffice', 'libreoffice', 'lowriter']);
  if (office) {
    return `office=${shellQuote(office)}; in="$1"; out="$2"; outdir="$(dirname "$out")"; base="$(basename "$in" .docx)"; "$office" --headless --convert-to pdf --outdir "$outdir" "$in" >/dev/null && mv "$outdir/$base.pdf" "$out"`;
  }

  const pandoc = detectCommand(['pandoc']);
  const pdfEngine = detectCommand(['pdflatex', 'lualatex', 'xelatex']);
  if (pandoc && pdfEngine && pandocLatexWorks()) {
    return `${pandoc} "$1" --pdf-engine=${pdfEngine} -o "$2"`;
  }

  return null;
}

function detectCommand(names) {
  for (const name of names) {
    try {
      const resolved = execFileSync('bash', ['-lc', `command -v ${name}`], { stdio: 'pipe' })
        .toString('utf8')
        .trim();
      if (resolved) return resolved;
    } catch (_) {
      // continue probing
    }
  }
  return null;
}

function pandocLatexWorks() {
  try {
    execFileSync('bash', ['-lc', 'command -v kpsewhich >/dev/null && kpsewhich letltxmacro.sty >/dev/null'], {
      stdio: 'pipe',
    });
    return true;
  } catch (_) {
    return false;
  }
}

function shellQuote(value) {
  return `'${String(value).replace(/'/g, `'\"'\"'`)}'`;
}
