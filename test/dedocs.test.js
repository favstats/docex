'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  applyTransforms,
  compareDocxPackages,
  compileDedocsText,
  dedocsFromDocx,
  normalizeDedocsText,
  parsePackage,
  serializePackage,
} = require('../dedocs');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');

function tmpPath(name) {
  return path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'dedocs-test-')), name);
}

function readDocxPart(docxPath, partPath) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dedocs-part-'));
  try {
    require('child_process').execFileSync('unzip', ['-o', docxPath, '-d', dir], {
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    return fs.readFileSync(path.join(dir, partPath), 'utf8');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

describe('dedocs format', () => {
  it('serializes and parses a mixed package deterministically', () => {
    const pkg = {
      version: '1',
      package: 'docx',
      fidelity: 'package-exact',
      source: 'sample.docx',
      parts: [
        {
          path: 'word/document.xml',
          mediaType: 'application/xml',
          encoding: 'utf8',
          buffer: Buffer.from('<?xml version="1.0"?><doc><p>Hello</p></doc>', 'utf8'),
        },
        {
          path: 'word/media/blob.bin',
          mediaType: 'application/octet-stream',
          encoding: 'base64',
          buffer: Buffer.from([0, 1, 2, 3, 4, 250, 251, 252, 253, 254, 255]),
        },
      ],
    };

    const text = serializePackage(pkg);
    const reparsed = parsePackage(text, { strictMetadata: true });
    const reserialized = serializePackage(reparsed);

    assert.equal(reserialized, text);
    assert.equal(reparsed.parts.length, 2);
    assert.equal(reparsed.parts[0].path, 'word/document.xml');
    assert.ok(reparsed.parts[0].buffer.equals(pkg.parts[0].buffer));
    assert.ok(reparsed.parts[1].buffer.equals(pkg.parts[1].buffer));
  });

  it('round-trips a docx fixture package-exactly', () => {
    const text = dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' });
    const parsed = parsePackage(text, { strictMetadata: true });
    assert.equal(parsed.guides.length, 1);
    assert.match(parsed.guides[0].text, /\\p\[index="0000", style="Heading1"\] Introduction/);

    const outDocx = tmpPath('roundtrip.docx');
    compileDedocsText(text, outDocx, { strictMetadata: true });

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.equal(comparison.equal, true, JSON.stringify(comparison.diffs, null, 2));
  });

  it('supports direct xml edits in the single text file', () => {
    const original = dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' });
    const edited = original.replace(
      '<w:t xml:space="preserve">This is the first paragraph of the introduction. It contains some text about platform governance and political advertising.</w:t>',
      '<w:t xml:space="preserve">This is the first paragraph of the introduction. It contains some text about platform governance, political advertising, and enforcement.</w:t>'
    );
    const outDocx = tmpPath('edited.docx');
    compileDedocsText(edited, outDocx);

    const documentXml = readDocxPart(outDocx, 'word/document.xml');
    assert.match(documentXml, /platform governance, political advertising, and enforcement\./);

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.deepEqual(comparison.diffs, [
      { path: 'word/document.xml', type: 'content' },
    ]);
  });

  it('supports authoring-layer replace-text transforms', () => {
    const pkg = parsePackage(dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' }), {
      strictMetadata: true,
    });

    pkg.transforms = [{
      type: 'replace-text',
      part: 'word/document.xml',
      count: 1,
      find: 'platform governance and political advertising.',
      replace: 'platform governance, political advertising, and enforcement.',
    }];

    const transformedPkg = applyTransforms(pkg);
    const transformedXml = transformedPkg.parts.find(part => part.path === 'word/document.xml').buffer.toString('utf8');
    assert.match(transformedXml, /platform governance, political advertising, and enforcement\./);

    const outDocx = tmpPath('transformed.docx');
    compileDedocsText(serializePackage(pkg), outDocx, { strictMetadata: true });

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.deepEqual(comparison.diffs, [
      { path: 'word/document.xml', type: 'content' },
    ]);
  });

  it('serializes and applies replace-paragraph transforms', () => {
    const pkg = parsePackage(dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' }), {
      strictMetadata: true,
    });

    pkg.transforms = [{
      type: 'replace-paragraph',
      index: '0001',
      expectedText: 'This is the first paragraph of the introduction. It contains some text about platform governance and political advertising.',
      text: 'This paragraph was rewritten through the semantic dedocs paragraph layer.',
    }];

    const transformedPkg = applyTransforms(pkg);
    const transformedXml = transformedPkg.parts.find(part => part.path === 'word/document.xml').buffer.toString('utf8');
    assert.match(transformedXml, /semantic dedocs paragraph layer/);

    const outDocx = tmpPath('replace-paragraph.docx');
    compileDedocsText(serializePackage(pkg), outDocx, { strictMetadata: true });

    const documentXml = readDocxPart(outDocx, 'word/document.xml');
    assert.match(documentXml, /semantic dedocs paragraph layer/);

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.deepEqual(comparison.diffs, [
      { path: 'word/document.xml', type: 'content' },
    ]);
  });

  it('serializes and applies insert-paragraph-after transforms', () => {
    const pkg = parsePackage(dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' }), {
      strictMetadata: true,
    });

    pkg.transforms = [{
      type: 'insert-paragraph-after',
      index: '0003',
      expectedText: 'Methods',
      expectedStyle: 'Heading1',
      text: 'Inserted paragraph after the Methods heading.',
    }];

    const outDocx = tmpPath('insert-paragraph.docx');
    compileDedocsText(serializePackage(pkg), outDocx, { strictMetadata: true });

    const documentXml = readDocxPart(outDocx, 'word/document.xml');
    assert.match(documentXml, /Inserted paragraph after the Methods heading\./);

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.deepEqual(comparison.diffs, [
      { path: 'word/document.xml', type: 'content' },
    ]);
  });

  it('serializes and applies insert-paragraph-before transforms', () => {
    const pkg = parsePackage(dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' }), {
      strictMetadata: true,
    });

    pkg.transforms = [{
      type: 'insert-paragraph-before',
      index: '0005',
      expectedText: 'Results',
      expectedStyle: 'Heading1',
      text: 'Inserted paragraph before the Results heading.',
    }];

    const outDocx = tmpPath('insert-paragraph-before.docx');
    compileDedocsText(serializePackage(pkg), outDocx, { strictMetadata: true });

    const documentXml = readDocxPart(outDocx, 'word/document.xml');
    assert.match(documentXml, /Inserted paragraph before the Results heading\./);

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.deepEqual(comparison.diffs, [
      { path: 'word/document.xml', type: 'content' },
    ]);
  });

  it('serializes and applies delete-paragraph transforms', () => {
    const pkg = parsePackage(dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' }), {
      strictMetadata: true,
    });

    pkg.transforms = [{
      type: 'delete-paragraph',
      index: '0007',
      expectedText: 'Discussion',
      expectedStyle: 'Heading1',
    }];

    const outDocx = tmpPath('delete-paragraph.docx');
    compileDedocsText(serializePackage(pkg), outDocx, { strictMetadata: true });

    const documentXml = readDocxPart(outDocx, 'word/document.xml');
    assert.doesNotMatch(documentXml, />Discussion</);

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.deepEqual(comparison.diffs, [
      { path: 'word/document.xml', type: 'content' },
    ]);
  });

  it('normalizes stale metadata and regenerates guides after edits', () => {
    const original = dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' });
    const edited = original.replace(
      '<w:t xml:space="preserve">This is the first paragraph of the introduction. It contains some text about platform governance and political advertising.</w:t>',
      '<w:t xml:space="preserve">This is the first paragraph of the introduction. It contains some text about platform governance, political advertising, and enforcement.</w:t>'
    );

    const normalized = normalizeDedocsText(edited);
    const reparsed = parsePackage(normalized, { strictMetadata: true });

    assert.match(reparsed.guides[0].text, /platform governance, political advertising, and enforcement\./);

    const outDocx = tmpPath('normalized.docx');
    compileDedocsText(normalized, outDocx, { strictMetadata: true });

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.deepEqual(comparison.diffs, [
      { path: 'word/document.xml', type: 'content' },
    ]);
  });

  it('normalizes guides against the transformed preview, not the raw core only', () => {
    const pkg = parsePackage(dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' }), {
      strictMetadata: true,
    });

    pkg.transforms = [{
      type: 'insert-paragraph-after',
      index: '0003',
      expectedText: 'Methods',
      expectedStyle: 'Heading1',
      text: 'Guide preview paragraph from semantic transform.',
    }];

    const normalized = normalizeDedocsText(serializePackage(pkg));
    const reparsed = parsePackage(normalized, { strictMetadata: true });

    assert.match(reparsed.guides[0].text, /Guide preview paragraph from semantic transform\./);
  });
});
