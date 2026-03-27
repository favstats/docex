/**
 * docex test suite
 *
 * Tests are organized in layers:
 *   1. XML utilities (pure functions, no I/O)
 *   2. Workspace lifecycle (zip/unzip/verify)
 *   3. Read operations (list, find, headings)
 *   4. Tracked changes (replace, insert, delete)
 *   5. Comments (add, reply, resolve, remove)
 *   6. Figures (insert, replace, list)
 *   7. Tables (insert with booktabs)
 *   8. Fluent API (chaining, position selectors)
 *   9. Edge cases (run splitting, Unicode, empty paragraphs)
 *  10. Round-trip integrity (open, edit, save, reopen)
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/docex.test.js
 */

const { describe, it, before, after, beforeEach } = require('node:test');
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
  const content = fs.readFileSync(path.join(tmp, xmlFile), 'utf8');
  execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
  return content;
}

/** Helper: count occurrences of a pattern in a string */
function countMatches(str, pattern) {
  const re = pattern instanceof RegExp ? pattern : new RegExp(pattern, 'g');
  return (str.match(re) || []).length;
}

// ============================================================================
// 1. XML UTILITIES
// ============================================================================

describe('xml utilities', () => {
  let xml;
  before(() => { xml = require('../src/xml'); });

  it('escapeXml handles all special characters', () => {
    assert.equal(xml.escapeXml('a & b < c > d "e" \'f\''), 'a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;');
  });

  it('decodeXml reverses escapeXml', () => {
    const original = 'Meta\'s "ban" & <enforcement>';
    assert.equal(xml.decodeXml(xml.escapeXml(original)), original);
  });

  it('extractText concatenates all w:t elements', () => {
    const pXml = '<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>';
    assert.equal(xml.extractText(pXml), 'Hello World');
  });

  it('extractText skips deleted text', () => {
    const pXml = '<w:p><w:r><w:t>Keep</w:t></w:r><w:del><w:r><w:delText>Remove</w:delText></w:r></w:del></w:p>';
    assert.equal(xml.extractText(pXml), 'Keep');
  });

  it('nextChangeId returns max+1', () => {
    const docXml = '<w:del w:id="5"/><w:ins w:id="10"/><w:del w:id="3"/>';
    assert.equal(xml.nextChangeId(docXml), 11);
  });

  it('nextChangeId returns 1 for empty document', () => {
    assert.equal(xml.nextChangeId('<w:document></w:document>'), 1);
  });

  it('emuFromInches converts correctly', () => {
    assert.equal(xml.emuFromInches(1), 914400);
    assert.equal(xml.emuFromInches(6.5), 5943600);
  });

  it('randomHexId returns 8 hex chars', () => {
    const id = xml.randomHexId();
    assert.match(id, /^[0-9a-f]{8}$/);
  });

  it('buildDel creates valid deletion markup', () => {
    const del = xml.buildDel(42, 'Fabio', '2026-03-27T00:00:00Z', '<w:rPr><w:b/></w:rPr>', 'old text');
    assert.ok(del.includes('w:id="42"'));
    assert.ok(del.includes('w:author="Fabio"'));
    assert.ok(del.includes('<w:delText'));
    assert.ok(del.includes('old text'));
  });

  it('buildIns creates valid insertion markup', () => {
    const ins = xml.buildIns(43, 'Fabio', '2026-03-27T00:00:00Z', '<w:rPr><w:b/></w:rPr>', 'new text');
    assert.ok(ins.includes('w:id="43"'));
    assert.ok(ins.includes('<w:t'));
    assert.ok(ins.includes('new text'));
  });

  it('parseRuns extracts runs with formatting', () => {
    const pXml = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t>normal</w:t></w:r></w:p>';
    const runs = xml.parseRuns(pXml);
    assert.equal(runs.length, 2);
    assert.ok(runs[0].rPr.includes('<w:b/>'));
    assert.equal(runs[0].combinedText, 'bold');
    assert.equal(runs[1].combinedText, 'normal');
  });

  it('findParagraphs returns all paragraphs with positions', () => {
    const docXml = '<w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body>';
    const paras = xml.findParagraphs(docXml);
    assert.equal(paras.length, 2);
    assert.equal(paras[0].text, 'First');
    assert.equal(paras[1].text, 'Second');
    assert.ok(paras[0].start < paras[1].start);
  });
});

// ============================================================================
// 2. WORKSPACE LIFECYCLE
// ============================================================================

describe('workspace', () => {
  let Workspace;
  before(() => { Workspace = require('../src/workspace').Workspace; });

  it('opens a valid .docx file', () => {
    const ws = Workspace.open(FIXTURE);
    assert.ok(ws.docXml.includes('<w:body>'));
    assert.ok(ws.docXml.includes('Introduction'));
    ws.cleanup();
  });

  it('counts paragraphs on open', () => {
    const ws = Workspace.open(FIXTURE);
    assert.ok(ws.originalParagraphCount >= 9); // 4 headings + 4 body + sectPr
    ws.cleanup();
  });

  it('reads styles.xml', () => {
    const ws = Workspace.open(FIXTURE);
    assert.ok(ws.stylesXml.includes('Heading1'));
    ws.cleanup();
  });

  it('reads relationships', () => {
    const ws = Workspace.open(FIXTURE);
    assert.ok(ws.relsXml.includes('styles.xml'));
    ws.cleanup();
  });

  it('saves and produces valid zip', () => {
    const out = freshCopy('workspace-save');
    const ws = Workspace.open(out);
    const result = ws.save(out);
    assert.ok(result.verified);
    assert.ok(result.fileSize > 0);
    assert.ok(result.paragraphCount >= 9);
  });

  it('save preserves document content', () => {
    const out = freshCopy('workspace-roundtrip');
    const ws = Workspace.open(out);
    ws.save(out);
    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('Introduction'));
    assert.ok(xml.includes('268,635'));
    assert.ok(xml.includes('platform self-regulation'));
  });

  it('throws on non-existent file', () => {
    assert.throws(() => Workspace.open('/tmp/nonexistent.docx'), /not found|ENOENT/i);
  });
});

// ============================================================================
// 3. READ OPERATIONS
// ============================================================================

describe('read operations', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('lists all paragraphs', async () => {
    const doc = docex(FIXTURE);
    const paras = await doc.paragraphs();
    assert.ok(paras.length >= 9);
    assert.ok(paras.some(p => p.text.includes('Introduction')));
    assert.ok(paras.some(p => p.text.includes('268,635')));
    doc.discard();
  });

  it('lists headings', async () => {
    const doc = docex(FIXTURE);
    const headings = await doc.headings();
    assert.ok(headings.length >= 4);
    const texts = headings.map(h => h.text);
    assert.ok(texts.includes('Introduction'));
    assert.ok(texts.includes('Methods'));
    assert.ok(texts.includes('Results'));
    assert.ok(texts.includes('Discussion'));
    doc.discard();
  });

  it('gets full text', async () => {
    const doc = docex(FIXTURE);
    const text = await doc.text();
    assert.ok(text.includes('platform governance'));
    assert.ok(text.includes('268,635'));
    doc.discard();
  });
});

// ============================================================================
// 4. TRACKED CHANGES
// ============================================================================

describe('tracked changes', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('replace creates w:del and w:ins', async () => {
    const out = freshCopy('tc-replace');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.replace('268,635', '300,000');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('<w:del'), 'should have w:del');
    assert.ok(xml.includes('<w:ins'), 'should have w:ins');
    assert.ok(xml.includes('268,635'), 'deleted text preserved in w:delText');
    assert.ok(xml.includes('300,000'), 'new text in w:ins');
    assert.ok(xml.includes('Fabio Votta'), 'author attributed');
  });

  it('replace preserves formatting', async () => {
    const out = freshCopy('tc-format');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.replace('268,635 advertisements', 'three hundred thousand ads');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    // The insertion should carry the same rPr as the original text
    const insMatch = xml.match(/<w:ins[^>]*>[\s\S]*?<\/w:ins>/);
    assert.ok(insMatch, 'should have insertion');
    assert.ok(insMatch[0].includes('Times New Roman'), 'insertion preserves font');
  });

  it('untracked replace modifies text directly', async () => {
    const out = freshCopy('tc-untracked');
    const doc = docex(out);
    doc.untracked();
    doc.replace('268,635', '300,000');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(!xml.includes('<w:del'), 'no w:del for untracked');
    assert.ok(!xml.includes('<w:ins'), 'no w:ins for untracked');
    assert.ok(xml.includes('300,000'), 'text replaced');
    assert.ok(!xml.includes('268,635'), 'old text gone');
  });

  it('insert after heading adds new paragraph', async () => {
    const out = freshCopy('tc-insert');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.after('Methods').insert('This is a new methodology paragraph.');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('new methodology paragraph'), 'new text present');
    assert.ok(xml.includes('<w:ins'), 'tracked as insertion');
    // Verify it comes after Methods heading
    const methodsPos = xml.indexOf('Methods');
    const newPos = xml.indexOf('new methodology paragraph');
    assert.ok(newPos > methodsPos, 'inserted after Methods');
  });

  it('delete wraps text in w:del', async () => {
    const out = freshCopy('tc-delete');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.delete('specific failure mode');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('<w:del'), 'should have deletion');
    assert.ok(xml.includes('specific failure mode'), 'deleted text in delText');
  });

  it('multiple replacements in one save', async () => {
    const out = freshCopy('tc-multi');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.replace('268,635', '300,000');
    doc.replace('1,329', '1,500');
    doc.replace('192 accounts', '200 accounts');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('300,000'));
    assert.ok(xml.includes('1,500'));
    assert.ok(xml.includes('200 accounts'));
    assert.equal(countMatches(xml, /<w:del /g), 3, 'three deletions');
    assert.equal(countMatches(xml, /<w:ins /g), 3, 'three insertions');
  });
});

// ============================================================================
// 5. COMMENTS
// ============================================================================

describe('comments', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('adds a comment anchored to text', async () => {
    const out = freshCopy('comment-add');
    const doc = docex(out);
    doc.at('platform governance').comment('Cite Gorwa 2019', { by: 'Prof. Strict' });
    await doc.save();

    const comments = readDocxXml(out, 'word/comments.xml');
    assert.ok(comments.includes('Cite Gorwa 2019'), 'comment text present');
    assert.ok(comments.includes('Prof. Strict'), 'author attributed');

    const docXml = readDocxXml(out, 'word/document.xml');
    assert.ok(docXml.includes('commentRangeStart'), 'range start in document');
    assert.ok(docXml.includes('commentRangeEnd'), 'range end in document');
    assert.ok(docXml.includes('commentReference'), 'reference in document');
  });

  it('adds multiple comments from different authors', async () => {
    const out = freshCopy('comment-multi');
    const doc = docex(out);
    doc.at('platform governance').comment('Theory needs work', { by: 'Reviewer 2' });
    doc.at('268,635').comment('Verify this number', { by: 'Dr. Numbers' });
    doc.at('electoral transparency').comment('Define this term', { by: 'Editor Chen' });
    await doc.save();

    const comments = readDocxXml(out, 'word/comments.xml');
    assert.equal(countMatches(comments, /<w:comment /g), 3, 'three comments');
    assert.ok(comments.includes('Reviewer 2'));
    assert.ok(comments.includes('Dr. Numbers'));
    assert.ok(comments.includes('Editor Chen'));
  });

  it('creates commentsExtended.xml for threading', async () => {
    const out = freshCopy('comment-extended');
    const doc = docex(out);
    doc.at('platform governance').comment('Needs citation', { by: 'Reviewer 1' });
    await doc.save();

    const ext = readDocxXml(out, 'word/commentsExtended.xml');
    assert.ok(ext.includes('commentEx'), 'commentsExtended.xml has entries');
    assert.ok(ext.includes('paraId'), 'has paraId');
  });

  it('lists existing comments', async () => {
    const out = freshCopy('comment-list');
    const doc = docex(out);
    doc.at('platform governance').comment('Test comment', { by: 'Tester' });
    await doc.save();

    const doc2 = docex(out);
    const comments = await doc2.comments();
    assert.ok(comments.length >= 1);
    assert.ok(comments.some(c => c.text.includes('Test comment')));
    doc2.discard();
  });
});

// ============================================================================
// 6. FIGURES
// ============================================================================

describe('figures', () => {
  let docex;
  before(() => {
    docex = require('../src/docex');
    // Create a tiny 1x1 red PNG for testing
    const png = Buffer.from([
      0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
      0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
      0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
      0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // 8-bit RGB
      0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
      0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01,
      0xE2, 0x21, 0xBC, 0x33,
      0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
      0xAE, 0x42, 0x60, 0x82
    ]);
    fs.writeFileSync(path.join(__dirname, 'fixtures', 'test-image.png'), png);
  });

  it('inserts a figure after a heading', async () => {
    const out = freshCopy('figure-insert');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.after('Results').figure(
      path.join(__dirname, 'fixtures', 'test-image.png'),
      'Figure 1. Test figure caption'
    );
    await doc.save();

    const docXml = readDocxXml(out, 'word/document.xml');
    assert.ok(docXml.includes('w:drawing'), 'has drawing element');
    assert.ok(docXml.includes('Test figure caption'), 'has caption');

    // Verify image was copied to media
    const rels = readDocxXml(out, 'word/_rels/document.xml.rels');
    assert.ok(rels.includes('media/'), 'has media relationship');
  });

  it('lists figures in document', async () => {
    const out = freshCopy('figure-list');
    const doc = docex(out);
    doc.after('Results').figure(
      path.join(__dirname, 'fixtures', 'test-image.png'),
      'Figure 1. Test'
    );
    await doc.save();

    const doc2 = docex(out);
    const figs = await doc2.figures();
    assert.ok(figs.length >= 1);
    doc2.discard();
  });
});

// ============================================================================
// 7. TABLES
// ============================================================================

describe('tables', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('inserts a booktabs table', async () => {
    const out = freshCopy('table-insert');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.after('Results').table(
      [['Party', 'Ads', 'Share'], ['PAX', '117', '8.8%'], ['Wakker Emmen', '82', '6.2%']],
      { caption: 'Table 1. Top advertisers', style: 'booktabs' }
    );
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('<w:tbl'), 'has table element');
    assert.ok(xml.includes('PAX'), 'has cell data');
    assert.ok(xml.includes('117'), 'has cell data');
    assert.ok(xml.includes('Top advertisers'), 'has caption');
  });

  it('header row is bold in booktabs style', async () => {
    const out = freshCopy('table-headers');
    const doc = docex(out);
    doc.after('Results').table(
      [['Col A', 'Col B'], ['val1', 'val2']],
      { style: 'booktabs' }
    );
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    // First row should have bold formatting
    const tblMatch = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/);
    assert.ok(tblMatch, 'table found');
    const firstRow = tblMatch[0].match(/<w:tr>[\s\S]*?<\/w:tr>/);
    assert.ok(firstRow[0].includes('<w:b'), 'header row has bold');
  });
});

// ============================================================================
// 8. FLUENT API
// ============================================================================

describe('fluent API', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('chains multiple operations', async () => {
    const out = freshCopy('fluent-chain');
    const doc = docex(out);
    doc.author('Fabio Votta')
       .replace('268,635', '300,000')
       .after('Methods').insert('New paragraph.')
       .at('platform governance').comment('Needs work', { by: 'Reviewer 1' });

    const result = await doc.save(out);
    assert.ok(result.verified);

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('300,000'));
    assert.ok(xml.includes('New paragraph'));
    assert.ok(xml.includes('commentRangeStart'));
  });

  it('discard clears pending operations', () => {
    const doc = docex(FIXTURE);
    doc.replace('foo', 'bar');
    doc.replace('baz', 'qux');
    doc.discard();
    assert.equal(doc._operations.length, 0);
  });

  it('author persists across operations', async () => {
    const out = freshCopy('fluent-author');
    const doc = docex(out);
    doc.author('Simon Kruschinski');
    doc.replace('268,635', '300,000');
    doc.replace('1,329', '1,500');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.equal(countMatches(xml, /Simon Kruschinski/g), 2, 'author on both changes');
  });
});

// ============================================================================
// 9. EDGE CASES
// ============================================================================

describe('edge cases', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('handles text split across runs (run splitting)', async () => {
    // The fixture has "Meta's advertising ban" split across 3 runs
    // (plain "Meta" + bold "'s advertising ban" + plain " and its consequences")
    const out = freshCopy('edge-split');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.replace("Meta's advertising ban", 'the platform advertising prohibition');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('platform advertising prohibition'), 'replacement succeeded across runs');
  });

  it('handles XML entities in text', async () => {
    const out = freshCopy('edge-entities');
    const doc = docex(out);
    doc.author('Fabio Votta');
    // The fixture has &apos; (apostrophe entity) in "Meta's"
    doc.replace("Meta's advertising ban", 'the advertising prohibition');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('advertising prohibition'));
  });

  it('replace with empty string deletes text', async () => {
    const out = freshCopy('edge-empty');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.replace('specific failure mode of ', '');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('<w:del'), 'deletion tracked');
  });

  it('warns but does not crash on text not found', async () => {
    const out = freshCopy('edge-notfound');
    const doc = docex(out);
    doc.replace('this text does not exist anywhere', 'replacement');
    // Should not throw, just warn
    const result = await doc.save();
    assert.ok(result.verified);
  });

  it('handles Unicode text', async () => {
    const out = freshCopy('edge-unicode');
    const doc = docex(out);
    doc.after('Introduction').insert('Gemeenteraadsverkiezingen 2026: EUR 15,346 spent.');
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('Gemeenteraadsverkiezingen'));
    assert.ok(xml.includes('15,346'));
  });
});

// ============================================================================
// 10. ROUND-TRIP INTEGRITY
// ============================================================================

describe('round-trip integrity', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('paragraph count never decreases', async () => {
    const out = freshCopy('integrity-paracount');
    const doc = docex(out);
    const before = (await doc.paragraphs()).length;
    doc.author('Fabio Votta');
    doc.replace('268,635', '300,000');
    doc.after('Methods').insert('Extra paragraph.');
    const result = await doc.save();
    assert.ok(result.paragraphCount >= before, 'paragraph count did not decrease');
  });

  it('file size stays reasonable after edits', async () => {
    const out = freshCopy('integrity-size');
    const originalSize = fs.statSync(out).size;
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.replace('268,635', '300,000');
    doc.at('platform governance').comment('Check this', { by: 'Reviewer 1' });
    const result = await doc.save();
    // File should not shrink to near-zero (corruption) or grow 10x (bloat)
    assert.ok(result.fileSize > originalSize * 0.5, 'not suspiciously small');
    assert.ok(result.fileSize < originalSize * 10, 'not suspiciously large');
  });

  it('output is a valid zip', async () => {
    const out = freshCopy('integrity-zip');
    const doc = docex(out);
    doc.replace('268,635', '300,000');
    await doc.save();

    // unzip -t should succeed
    const result = execFileSync('unzip', ['-t', out], { stdio: 'pipe' }).toString();
    assert.ok(result.includes('No errors'), 'valid zip archive');
  });

  it('save to different path preserves original', async () => {
    const original = freshCopy('integrity-preserve');
    const output = path.join(OUTPUT_DIR, 'integrity-preserve-v2.docx');
    const doc = docex(original);
    doc.replace('268,635', '300,000');
    await doc.save(output);

    // Original should be unchanged
    const origXml = readDocxXml(original, 'word/document.xml');
    assert.ok(origXml.includes('268,635'), 'original preserved');
    assert.ok(!origXml.includes('300,000'), 'original not modified');

    // Output should have changes
    const outXml = readDocxXml(output, 'word/document.xml');
    assert.ok(outXml.includes('300,000'), 'output has changes');
  });

  it('multiple save cycles do not corrupt', async () => {
    const out = freshCopy('integrity-multisave');

    // First edit
    let doc = docex(out);
    doc.author('Fabio Votta');
    doc.replace('268,635', '300,000');
    await doc.save();

    // Second edit on the same file
    doc = docex(out);
    doc.author('Simon Kruschinski');
    doc.replace('1,329', '1,500');
    await doc.save();

    // Third edit
    doc = docex(out);
    doc.at('platform governance').comment('Final check', { by: 'Editor' });
    await doc.save();

    const xml = readDocxXml(out, 'word/document.xml');
    assert.ok(xml.includes('300,000'), 'first edit preserved');
    assert.ok(xml.includes('1,500'), 'second edit preserved');
    assert.ok(xml.includes('commentRangeStart'), 'third edit preserved');
  });
});
