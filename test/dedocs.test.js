'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  compareDocxPackages,
  compileDedocsText,
  dedocsFromDocx,
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
    const outDocx = tmpPath('roundtrip.docx');
    compileDedocsText(text, outDocx, { strictMetadata: true });

    const comparison = compareDocxPackages(FIXTURE, outDocx);
    assert.equal(comparison.equal, true, JSON.stringify(comparison.diffs, null, 2));
  });

  it('supports direct xml edits in the single text file', () => {
    const original = dedocsFromDocx(FIXTURE, { source: 'test-manuscript.docx' });
    const edited = original.replace(
      'platform governance and political advertising.',
      'platform governance, political advertising, and enforcement.'
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
});
