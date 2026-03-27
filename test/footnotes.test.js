/**
 * footnotes test suite
 *
 * Tests for the Footnotes module:
 *   1. Add a footnote to the test fixture
 *   2. Verify footnotes.xml is created with separator footnotes
 *   3. Verify footnote reference appears in document.xml
 *   4. List footnotes returns added footnotes
 *   5. Multiple footnotes get sequential IDs
 *   6. Verify relationship and content type added
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/footnotes.test.js
 */

const { describe, it, before } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

/** Helper: create a fresh copy of the fixture for each test */
function freshCopy(testName) {
  const out = path.join(OUTPUT_DIR, `${testName}.docx`);
  fs.copyFileSync(FIXTURE, out);
  return out;
}

/** Helper: unzip a docx and read a specific XML file */
function readDocxXml(docxPath, xmlFile) {
  const tmp = fs.mkdtempSync('/tmp/docex-test-');
  execFileSync('unzip', ['-o', docxPath, '-d', tmp], { stdio: 'pipe' });
  const filePath = path.join(tmp, xmlFile);
  let content = null;
  if (fs.existsSync(filePath)) {
    content = fs.readFileSync(filePath, 'utf8');
  }
  execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
  return content;
}

/** Helper: count occurrences of a pattern in a string */
function countMatches(str, pattern) {
  const re = pattern instanceof RegExp ? pattern : new RegExp(pattern, 'g');
  return (str.match(re) || []).length;
}

// ============================================================================
// FOOTNOTES
// ============================================================================

describe('footnotes', () => {
  let Workspace, Footnotes;

  before(() => {
    Workspace = require('../src/workspace').Workspace;
    Footnotes = require('../src/footnotes').Footnotes;
  });

  it('adds a footnote and creates footnotes.xml with separators', () => {
    const out = freshCopy('fn-add');
    const ws = Workspace.open(out);

    const result = Footnotes.add(ws, 'platform governance', 'See Smith 2020 for details.');
    assert.ok(result.footnoteId >= 2, 'footnote ID >= 2');

    ws.save(out);

    const fnXml = readDocxXml(out, 'word/footnotes.xml');
    assert.ok(fnXml, 'footnotes.xml exists');
    assert.ok(fnXml.includes('w:type="separator"'), 'has separator footnote');
    assert.ok(fnXml.includes('w:type="continuationSeparator"'), 'has continuation separator');
    assert.ok(fnXml.includes('See Smith 2020 for details.'), 'footnote text present');
    assert.ok(fnXml.includes('FootnoteText'), 'has FootnoteText paragraph style');
    assert.ok(fnXml.includes('FootnoteReference'), 'has FootnoteReference run style');
    assert.ok(fnXml.includes('w:footnoteRef'), 'has footnoteRef element');
  });

  it('inserts footnote reference in document.xml', () => {
    const out = freshCopy('fn-docref');
    const ws = Workspace.open(out);

    const result = Footnotes.add(ws, 'platform governance', 'A footnote.');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    assert.ok(docXml.includes('w:footnoteReference'), 'footnoteReference in document.xml');
    assert.ok(docXml.includes('w:id="' + result.footnoteId + '"'), 'correct footnote ID in reference');
    assert.ok(docXml.includes('FootnoteReference'), 'FootnoteReference style on ref run');
  });

  it('list returns added footnotes, skipping separators', () => {
    const out = freshCopy('fn-list');
    const ws = Workspace.open(out);

    Footnotes.add(ws, 'platform governance', 'First footnote.');
    Footnotes.add(ws, '268,635', 'Second footnote.');

    const list = Footnotes.list(ws);
    assert.equal(list.length, 2, 'two footnotes listed');
    assert.ok(list[0].text.includes('First footnote'), 'first footnote text');
    assert.ok(list[1].text.includes('Second footnote'), 'second footnote text');
    assert.ok(list[0].id >= 2, 'first ID >= 2');
    assert.ok(list[1].id > list[0].id, 'second ID > first ID');

    ws.cleanup();
  });

  it('multiple footnotes get sequential IDs', () => {
    const out = freshCopy('fn-sequential');
    const ws = Workspace.open(out);

    const r1 = Footnotes.add(ws, 'platform governance', 'Footnote A.');
    const r2 = Footnotes.add(ws, '268,635', 'Footnote B.');
    const r3 = Footnotes.add(ws, 'electoral transparency', 'Footnote C.');

    assert.equal(r1.footnoteId, 2, 'first footnote ID is 2');
    assert.equal(r2.footnoteId, 3, 'second footnote ID is 3');
    assert.equal(r3.footnoteId, 4, 'third footnote ID is 4');

    ws.save(out);

    const fnXml = readDocxXml(out, 'word/footnotes.xml');
    assert.equal(countMatches(fnXml, /<w:footnote\b(?![^>]*w:type=)/g), 3,
      'three user footnotes (no type= attribute)');
  });

  it('adds relationship to document.xml.rels', () => {
    const out = freshCopy('fn-rels');
    const ws = Workspace.open(out);

    Footnotes.add(ws, 'platform governance', 'Test footnote.');
    ws.save(out);

    const rels = readDocxXml(out, 'word/_rels/document.xml.rels');
    assert.ok(rels.includes('relationships/footnotes'), 'footnotes relationship present');
    assert.ok(rels.includes('Target="footnotes.xml"'), 'target is footnotes.xml');
  });

  it('adds content type to [Content_Types].xml', () => {
    const out = freshCopy('fn-ct');
    const ws = Workspace.open(out);

    Footnotes.add(ws, 'platform governance', 'Test footnote.');
    ws.save(out);

    const ct = readDocxXml(out, '[Content_Types].xml');
    assert.ok(ct.includes('footnotes+xml'), 'footnotes content type present');
    assert.ok(ct.includes('/word/footnotes.xml'), 'footnotes part name present');
  });

  it('throws when anchor text is not found', () => {
    const out = freshCopy('fn-noanchor');
    const ws = Workspace.open(out);

    assert.throws(
      () => Footnotes.add(ws, 'nonexistent text xyz', 'Will fail.'),
      /could not find anchor text/,
      'throws on missing anchor'
    );

    ws.cleanup();
  });

  it('throws on empty footnote text', () => {
    const out = freshCopy('fn-empty');
    const ws = Workspace.open(out);

    assert.throws(
      () => Footnotes.add(ws, 'platform governance', ''),
      /text must be a non-empty string/,
      'throws on empty text'
    );

    ws.cleanup();
  });

  it('list returns empty array when no footnotes exist', () => {
    const out = freshCopy('fn-empty-list');
    const ws = Workspace.open(out);

    // footnotesXml getter will create the file with just separators
    const list = Footnotes.list(ws);
    assert.equal(list.length, 0, 'no footnotes listed');

    ws.cleanup();
  });

  it('round-trip: save and reopen preserves footnotes', () => {
    const out = freshCopy('fn-roundtrip');
    const ws = Workspace.open(out);

    Footnotes.add(ws, 'platform governance', 'Persistent footnote.');
    ws.save(out);

    // Reopen and verify
    const ws2 = Workspace.open(out);
    const list = Footnotes.list(ws2);
    assert.equal(list.length, 1, 'one footnote after reopen');
    assert.ok(list[0].text.includes('Persistent footnote'), 'text preserved');
    assert.equal(list[0].id, 2, 'ID preserved');

    ws2.cleanup();
  });
});
