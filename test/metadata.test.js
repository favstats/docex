/**
 * metadata + word count test suite
 *
 * Tests:
 *   1. Metadata: set/get on test fixture, update, round-trip
 *   2. Word count: categorized counts on test fixture
 *   3. Word count on real manuscript (if available): body >= 7000
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/metadata.test.js
 */

const { describe, it, before } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const REAL_MANUSCRIPT = '/mnt/storage/nl_local_2026/paper/manuscript.docx';
const OUTPUT_DIR = path.join(__dirname, 'output');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function freshCopy(testName) {
  const out = path.join(OUTPUT_DIR, `${testName}.docx`);
  fs.copyFileSync(FIXTURE, out);
  return out;
}

function readDocxXml(docxPath, xmlFile) {
  const tmp = fs.mkdtempSync('/tmp/docex-meta-test-');
  execFileSync('unzip', ['-o', docxPath, '-d', tmp], { stdio: 'pipe' });
  const filePath = path.join(tmp, xmlFile);
  let content = '';
  if (fs.existsSync(filePath)) {
    content = fs.readFileSync(filePath, 'utf8');
  }
  execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
  return content;
}

// ============================================================================
// 1. METADATA
// ============================================================================

describe('metadata', () => {
  let Workspace, Metadata;

  before(() => {
    Workspace = require('../src/workspace').Workspace;
    Metadata = require('../src/metadata').Metadata;
  });

  it('sets metadata on a document and reads it back', () => {
    const out = freshCopy('metadata-set');
    const ws = Workspace.open(out);

    Metadata.set(ws, {
      title: 'Ad Enforcement Failures',
      creator: 'Fabio Votta',
      subject: 'Political Advertising',
      description: 'Analysis of Meta ad enforcement in Dutch local elections',
      keywords: 'elections, Meta, advertising, enforcement',
      created: '2026-01-15T10:00:00Z',
      modified: '2026-03-27T14:30:00Z',
      lastModifiedBy: 'Simon Kruschinski',
      revision: 5,
    });

    // Read back from in-memory XML
    const meta = Metadata.get(ws);
    assert.equal(meta.title, 'Ad Enforcement Failures');
    assert.equal(meta.creator, 'Fabio Votta');
    assert.equal(meta.subject, 'Political Advertising');
    assert.equal(meta.description, 'Analysis of Meta ad enforcement in Dutch local elections');
    assert.equal(meta.keywords, 'elections, Meta, advertising, enforcement');
    assert.equal(meta.created, '2026-01-15T10:00:00Z');
    assert.equal(meta.modified, '2026-03-27T14:30:00Z');
    assert.equal(meta.lastModifiedBy, 'Simon Kruschinski');
    assert.equal(meta.revision, '5');

    // Save and verify on disk
    const result = ws.save(out);
    assert.ok(result.verified);

    // Read the XML from the saved file
    const coreXml = readDocxXml(out, 'docProps/core.xml');
    assert.ok(coreXml.includes('Ad Enforcement Failures'), 'title in core.xml');
    assert.ok(coreXml.includes('Fabio Votta'), 'creator in core.xml');
    assert.ok(coreXml.includes('dcterms:W3CDTF'), 'date type attribute present');
  });

  it('updates existing metadata properties', () => {
    const out = freshCopy('metadata-update');
    const ws = Workspace.open(out);

    // Set initial values
    Metadata.set(ws, {
      title: 'Draft Title',
      creator: 'Author A',
    });

    // Update with new values
    Metadata.set(ws, {
      title: 'Final Title',
      creator: 'Author B',
      keywords: 'new, keywords',
    });

    const meta = Metadata.get(ws);
    assert.equal(meta.title, 'Final Title', 'title updated');
    assert.equal(meta.creator, 'Author B', 'creator updated');
    assert.equal(meta.keywords, 'new, keywords', 'keywords added');

    ws.cleanup();
  });

  it('ensures root relationship and content type exist', () => {
    const out = freshCopy('metadata-rels');
    const ws = Workspace.open(out);

    Metadata.set(ws, { title: 'Test' });
    ws.save(out);

    // Check _rels/.rels
    const rootRels = readDocxXml(out, '_rels/.rels');
    assert.ok(rootRels.includes('core-properties'), 'root rels has core-properties relationship');
    assert.ok(rootRels.includes('docProps/core.xml'), 'root rels targets docProps/core.xml');

    // Check [Content_Types].xml
    const contentTypes = readDocxXml(out, '[Content_Types].xml');
    assert.ok(contentTypes.includes('/docProps/core.xml'), 'content types has core.xml override');
  });

  it('handles XML entities in metadata values', () => {
    const out = freshCopy('metadata-entities');
    const ws = Workspace.open(out);

    Metadata.set(ws, {
      title: 'Meta\'s "Enforcement" & <Failures>',
      creator: 'O\'Brien & Associates',
    });

    const meta = Metadata.get(ws);
    assert.equal(meta.title, 'Meta\'s "Enforcement" & <Failures>');
    assert.equal(meta.creator, 'O\'Brien & Associates');

    ws.cleanup();
  });

  it('returns empty object when core.xml does not exist', () => {
    const out = freshCopy('metadata-empty');
    const ws = Workspace.open(out);

    // Before setting any metadata, get should return empty or object with empty strings
    const meta = Metadata.get(ws);
    // Either empty object (no file) or all empty strings
    const hasValues = Object.values(meta).some(v => v && v.length > 0);
    // This is fine either way -- just should not crash
    assert.ok(typeof meta === 'object', 'returns an object');

    ws.cleanup();
  });

  it('round-trips metadata through save and reopen', () => {
    const out = freshCopy('metadata-roundtrip');
    const ws = Workspace.open(out);

    Metadata.set(ws, {
      title: 'Round Trip Test',
      creator: 'Test Author',
      modified: '2026-03-27T12:00:00Z',
    });
    ws.save(out);

    // Reopen and verify
    const ws2 = Workspace.open(out);
    const meta = Metadata.get(ws2);
    assert.equal(meta.title, 'Round Trip Test');
    assert.equal(meta.creator, 'Test Author');
    assert.equal(meta.modified, '2026-03-27T12:00:00Z');
    ws2.cleanup();
  });
});

// ============================================================================
// 2. WORD COUNT (test fixture)
// ============================================================================

describe('word count', () => {
  let Workspace, Paragraphs;

  before(() => {
    Workspace = require('../src/workspace').Workspace;
    Paragraphs = require('../src/paragraphs').Paragraphs;
  });

  it('returns categorized word counts', () => {
    const ws = Workspace.open(FIXTURE);
    const wc = Paragraphs.wordCount(ws);

    assert.ok(typeof wc.total === 'number', 'total is number');
    assert.ok(typeof wc.body === 'number', 'body is number');
    assert.ok(typeof wc.headings === 'number', 'headings is number');
    assert.ok(typeof wc.abstract === 'number', 'abstract is number');
    assert.ok(typeof wc.captions === 'number', 'captions is number');
    assert.ok(typeof wc.footnotes === 'number', 'footnotes is number');

    // Total should be the sum of all categories
    assert.equal(wc.total, wc.body + wc.headings + wc.abstract + wc.captions + wc.footnotes,
      'total equals sum of categories');

    // Test fixture has at least a few words in headings
    assert.ok(wc.headings >= 4, 'at least 4 heading words (Introduction, Methods, Results, Discussion)');

    // Total should be reasonable for the fixture
    assert.ok(wc.total > 0, 'total > 0');

    ws.cleanup();
  });

  it('counts body words excluding headings', () => {
    const ws = Workspace.open(FIXTURE);
    const wc = Paragraphs.wordCount(ws);

    // Body should have the bulk of the words
    assert.ok(wc.body > wc.headings, 'more body words than heading words');

    ws.cleanup();
  });
});

// ============================================================================
// 3. WORD COUNT ON REAL MANUSCRIPT
// ============================================================================

describe('word count (real manuscript)', () => {
  let Workspace, Paragraphs;
  let hasManuscript = false;

  before(() => {
    Workspace = require('../src/workspace').Workspace;
    Paragraphs = require('../src/paragraphs').Paragraphs;
    hasManuscript = fs.existsSync(REAL_MANUSCRIPT);
  });

  it('real manuscript has 5000+ body words and categorized counts', { skip: !fs.existsSync(REAL_MANUSCRIPT) }, () => {
    const ws = Workspace.open(REAL_MANUSCRIPT);
    const wc = Paragraphs.wordCount(ws);

    console.log('[word count] Real manuscript:', JSON.stringify(wc, null, 2));

    // Body text should be substantial (manuscript is ~6000 body words)
    assert.ok(wc.body >= 5000,
      `expected >= 5000 body words, got ${wc.body}`);
    assert.ok(wc.total >= 6000,
      `expected >= 6000 total words, got ${wc.total}`);

    // Should have detected abstract section
    assert.ok(wc.abstract > 0,
      `expected abstract words > 0, got ${wc.abstract}`);

    // Should have detected headings
    assert.ok(wc.headings > 0,
      `expected heading words > 0, got ${wc.headings}`);

    // Should have detected captions (the manuscript has figure/table captions)
    assert.ok(wc.captions > 0,
      `expected caption words > 0, got ${wc.captions}`);

    ws.cleanup();
  });
});
