/**
 * diff.test.js -- Tests for document comparison module
 *
 * Tests:
 *   1. Identical documents: 0 changes
 *   2. One paragraph added
 *   3. One paragraph removed
 *   4. Text changed within a paragraph
 *   5. Multiple changes (add + remove + modify)
 *   6. Correct w:ins and w:del elements
 *
 * Uses the test fixture (test-manuscript.docx).
 * Modified copies are created by manipulating the workspace docXml.
 *
 * Run: node --test test/diff.test.js
 */

const { describe, it, before } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

/** Helper: create a fresh copy of the fixture */
function freshCopy(testName) {
  const out = path.join(OUTPUT_DIR, `${testName}.docx`);
  fs.copyFileSync(FIXTURE, out);
  return out;
}

/** Helper: count occurrences of a regex in a string */
function countMatches(str, pattern) {
  const re = pattern instanceof RegExp ? pattern : new RegExp(pattern, 'g');
  return (str.match(re) || []).length;
}

// ============================================================================
// 1. INTERNAL ALGORITHMS (pure functions, no I/O)
// ============================================================================

describe('diff: internal algorithms', () => {
  let Diff;
  before(() => { Diff = require('../src/diff').Diff; });

  it('_diffParagraphs: identical lists produce all keep operations', () => {
    const texts = ['Hello', 'World', 'Foo'];
    const ops = Diff._diffParagraphs(texts, texts);
    assert.equal(ops.length, 3);
    assert.ok(ops.every(op => op.type === 'keep'));
  });

  it('_diffParagraphs: added paragraph detected', () => {
    const texts1 = ['A', 'B'];
    const texts2 = ['A', 'C', 'B'];
    const ops = Diff._diffParagraphs(texts1, texts2);
    const addOps = ops.filter(op => op.type === 'add');
    assert.equal(addOps.length, 1);
    assert.equal(addOps[0].index2, 1); // 'C' at index 1 in texts2
  });

  it('_diffParagraphs: removed paragraph detected', () => {
    const texts1 = ['A', 'B', 'C'];
    const texts2 = ['A', 'C'];
    const ops = Diff._diffParagraphs(texts1, texts2);
    const removeOps = ops.filter(op => op.type === 'remove');
    assert.equal(removeOps.length, 1);
    assert.equal(removeOps[0].index1, 1); // 'B' at index 1 in texts1
  });

  it('_diffParagraphs: modified paragraph detected (remove+add merged)', () => {
    const texts1 = ['A', 'B', 'C'];
    const texts2 = ['A', 'X', 'C'];
    const ops = Diff._diffParagraphs(texts1, texts2);
    const modOps = ops.filter(op => op.type === 'modify');
    assert.equal(modOps.length, 1);
    assert.equal(modOps[0].index1, 1);
    assert.equal(modOps[0].index2, 1);
  });

  it('_diffWords: identical text produces single keep segment', () => {
    const segments = Diff._diffWords('hello world', 'hello world');
    assert.equal(segments.length, 1);
    assert.equal(segments[0].type, 'keep');
  });

  it('_diffWords: changed word produces remove+add segments', () => {
    const segments = Diff._diffWords('hello world', 'hello planet');
    const types = segments.map(s => s.type);
    assert.ok(types.includes('remove'), 'has remove segment');
    assert.ok(types.includes('add'), 'has add segment');
    assert.ok(types.includes('keep'), 'has keep segment');
  });

  it('_diffWords: added word produces add segment', () => {
    const segments = Diff._diffWords('hello world', 'hello big world');
    const addSegs = segments.filter(s => s.type === 'add');
    assert.ok(addSegs.length >= 1);
    const addedText = addSegs.map(s => s.text).join('');
    assert.ok(addedText.includes('big'));
  });

  it('_tokenize: preserves whitespace', () => {
    const Diff = require('../src/diff').Diff;
    const tokens = Diff._tokenize('hello  world foo');
    assert.equal(tokens.join(''), 'hello  world foo');
  });

  it('_tokenize: empty string returns empty array', () => {
    const Diff = require('../src/diff').Diff;
    assert.deepEqual(Diff._tokenize(''), []);
  });
});

// ============================================================================
// 2. DOCUMENT COMPARISON (workspace-level)
// ============================================================================

describe('diff: document comparison', () => {
  let Diff, Workspace, xmlMod;

  before(() => {
    Diff = require('../src/diff').Diff;
    Workspace = require('../src/workspace').Workspace;
    xmlMod = require('../src/xml');
  });

  it('identical documents produce 0 changes', () => {
    const copy1 = freshCopy('diff-identical-1');
    const copy2 = freshCopy('diff-identical-2');
    const ws1 = Workspace.open(copy1);
    const ws2 = Workspace.open(copy2);

    const stats = Diff.compare(ws1, ws2, { author: 'Test', date: '2026-03-27T00:00:00Z' });

    assert.equal(stats.added, 0);
    assert.equal(stats.removed, 0);
    assert.equal(stats.modified, 0);
    assert.ok(stats.unchanged > 0, 'should have unchanged paragraphs');
    // No tracked changes should be present from the diff
    assert.ok(!ws1.docXml.includes('<w:del '), 'no deletions in identical diff');
    assert.ok(!ws1.docXml.includes('<w:ins '), 'no insertions in identical diff');

    ws1.cleanup();
    ws2.cleanup();
  });

  it('detects added paragraph', () => {
    const copy1 = freshCopy('diff-add-1');
    const copy2 = freshCopy('diff-add-2');
    const ws1 = Workspace.open(copy1);
    const ws2 = Workspace.open(copy2);

    // Add a paragraph to ws2
    const paras2 = xmlMod.findParagraphs(ws2.docXml);
    const insertAfter = paras2[0]; // after first paragraph
    const newPara = '<w:p><w:r><w:t xml:space="preserve">This is a brand new paragraph.</w:t></w:r></w:p>';
    ws2.docXml = ws2.docXml.slice(0, insertAfter.end) + newPara + ws2.docXml.slice(insertAfter.end);

    const stats = Diff.compare(ws1, ws2, { author: 'Fabio Votta', date: '2026-03-27T00:00:00Z' });

    assert.equal(stats.added, 1, 'one paragraph added');
    assert.equal(stats.removed, 0, 'no paragraphs removed');
    assert.ok(ws1.docXml.includes('<w:ins '), 'has w:ins element');
    assert.ok(ws1.docXml.includes('brand new paragraph'), 'inserted text present');
    assert.ok(ws1.docXml.includes('Fabio Votta'), 'author attributed');

    ws1.cleanup();
    ws2.cleanup();
  });

  it('detects removed paragraph', () => {
    const copy1 = freshCopy('diff-remove-1');
    const copy2 = freshCopy('diff-remove-2');
    const ws1 = Workspace.open(copy1);
    const ws2 = Workspace.open(copy2);

    // Remove the second paragraph from ws2
    const paras2 = xmlMod.findParagraphs(ws2.docXml);
    // Find a paragraph with substantial text to remove
    let targetIdx = -1;
    for (let i = 0; i < paras2.length; i++) {
      if (paras2[i].text.length > 10) { targetIdx = i; break; }
    }
    assert.ok(targetIdx >= 0, 'found a paragraph to remove');
    const target = paras2[targetIdx];
    ws2.docXml = ws2.docXml.slice(0, target.start) + ws2.docXml.slice(target.end);

    const stats = Diff.compare(ws1, ws2, { author: 'Fabio Votta', date: '2026-03-27T00:00:00Z' });

    assert.ok(stats.removed >= 1, 'at least one paragraph removed');
    assert.ok(ws1.docXml.includes('<w:del '), 'has w:del element');
    assert.ok(ws1.docXml.includes('<w:delText'), 'has w:delText element');
    assert.ok(ws1.docXml.includes('Fabio Votta'), 'author attributed');

    ws1.cleanup();
    ws2.cleanup();
  });

  it('detects text changed within a paragraph', () => {
    const copy1 = freshCopy('diff-modify-1');
    const copy2 = freshCopy('diff-modify-2');
    const ws1 = Workspace.open(copy1);
    const ws2 = Workspace.open(copy2);

    // Find a paragraph with "268,635" and change it to "300,000" in ws2
    const paras2 = xmlMod.findParagraphs(ws2.docXml);
    let modified = false;
    for (const p of paras2) {
      if (p.text.includes('268,635')) {
        // Replace in the raw XML
        const newXml = p.xml.replace('268,635', '300,000');
        ws2.docXml = ws2.docXml.slice(0, p.start) + newXml + ws2.docXml.slice(p.end);
        modified = true;
        break;
      }
    }
    assert.ok(modified, 'found paragraph to modify');

    const stats = Diff.compare(ws1, ws2, { author: 'Simon', date: '2026-03-27T00:00:00Z' });

    assert.ok(stats.modified >= 1, 'at least one paragraph modified');
    // Modified paragraph should have both del and ins
    assert.ok(ws1.docXml.includes('<w:del '), 'has w:del for removed words');
    assert.ok(ws1.docXml.includes('<w:ins '), 'has w:ins for added words');
    assert.ok(ws1.docXml.includes('Simon'), 'author attributed');

    ws1.cleanup();
    ws2.cleanup();
  });

  it('handles multiple changes at once', () => {
    const copy1 = freshCopy('diff-multi-1');
    const copy2 = freshCopy('diff-multi-2');
    const ws1 = Workspace.open(copy1);
    const ws2 = Workspace.open(copy2);

    // 1. Modify a paragraph (change 268,635 to 300,000)
    let paras2 = xmlMod.findParagraphs(ws2.docXml);
    for (const p of paras2) {
      if (p.text.includes('268,635')) {
        const newXml = p.xml.replace('268,635', '300,000');
        ws2.docXml = ws2.docXml.slice(0, p.start) + newXml + ws2.docXml.slice(p.end);
        break;
      }
    }

    // 2. Add a paragraph after the first one
    paras2 = xmlMod.findParagraphs(ws2.docXml);
    const insertAfter = paras2[0];
    const newPara = '<w:p><w:r><w:t xml:space="preserve">Added paragraph for multi-test.</w:t></w:r></w:p>';
    ws2.docXml = ws2.docXml.slice(0, insertAfter.end) + newPara + ws2.docXml.slice(insertAfter.end);

    // 3. Remove the last content paragraph (before sectPr paragraph if any)
    paras2 = xmlMod.findParagraphs(ws2.docXml);
    // Find last paragraph with real text
    let lastContentIdx = -1;
    for (let i = paras2.length - 1; i >= 0; i--) {
      if (paras2[i].text.length > 10) { lastContentIdx = i; break; }
    }
    if (lastContentIdx >= 0) {
      const last = paras2[lastContentIdx];
      ws2.docXml = ws2.docXml.slice(0, last.start) + ws2.docXml.slice(last.end);
    }

    const stats = Diff.compare(ws1, ws2, { author: 'Test Author', date: '2026-03-27T00:00:00Z' });

    assert.ok(stats.added >= 1, 'at least one added');
    assert.ok(stats.removed >= 1, 'at least one removed');
    assert.ok(stats.modified >= 1, 'at least one modified');
    assert.ok(stats.unchanged >= 1, 'at least one unchanged');

    ws1.cleanup();
    ws2.cleanup();
  });

  it('result has correct w:ins and w:del elements with attributes', () => {
    const copy1 = freshCopy('diff-xml-1');
    const copy2 = freshCopy('diff-xml-2');
    const ws1 = Workspace.open(copy1);
    const ws2 = Workspace.open(copy2);

    // Add a paragraph
    const paras2 = xmlMod.findParagraphs(ws2.docXml);
    const insertAfter = paras2[0];
    const newPara = '<w:p><w:r><w:t xml:space="preserve">Inserted line.</w:t></w:r></w:p>';
    ws2.docXml = ws2.docXml.slice(0, insertAfter.end) + newPara + ws2.docXml.slice(insertAfter.end);

    // Remove a paragraph
    const paras2b = xmlMod.findParagraphs(ws2.docXml);
    let removeIdx = -1;
    for (let i = paras2b.length - 1; i >= 0; i--) {
      if (paras2b[i].text.length > 10) { removeIdx = i; break; }
    }
    if (removeIdx >= 0) {
      const rem = paras2b[removeIdx];
      ws2.docXml = ws2.docXml.slice(0, rem.start) + ws2.docXml.slice(rem.end);
    }

    Diff.compare(ws1, ws2, { author: 'Dr. Review', date: '2026-03-27T12:00:00Z' });

    const resultXml = ws1.docXml;

    // Verify w:ins elements have correct attributes
    const insMatches = resultXml.match(/<w:ins [^>]+>/g) || [];
    for (const ins of insMatches) {
      assert.ok(ins.includes('w:id='), 'w:ins has w:id attribute');
      assert.ok(ins.includes('w:author='), 'w:ins has w:author attribute');
      assert.ok(ins.includes('w:date='), 'w:ins has w:date attribute');
      assert.ok(ins.includes('Dr. Review'), 'w:ins has correct author');
      assert.ok(ins.includes('2026-03-27T12:00:00Z'), 'w:ins has correct date');
    }

    // Verify w:del elements have correct attributes
    const delMatches = resultXml.match(/<w:del [^>]+>/g) || [];
    for (const del of delMatches) {
      assert.ok(del.includes('w:id='), 'w:del has w:id attribute');
      assert.ok(del.includes('w:author='), 'w:del has w:author attribute');
      assert.ok(del.includes('w:date='), 'w:del has w:date attribute');
      assert.ok(del.includes('Dr. Review'), 'w:del has correct author');
    }

    // Verify insertions use w:t and deletions use w:delText
    assert.ok(insMatches.length > 0, 'should have at least one w:ins');
    assert.ok(delMatches.length > 0, 'should have at least one w:del');

    // Check that w:del contains w:delText
    const delBlocks = resultXml.match(/<w:del [^>]+>[\s\S]*?<\/w:del>/g) || [];
    for (const block of delBlocks) {
      assert.ok(block.includes('<w:delText'), 'w:del contains w:delText');
    }

    // Check that w:ins contains w:t
    const insBlocks = resultXml.match(/<w:ins [^>]+>[\s\S]*?<\/w:ins>/g) || [];
    for (const block of insBlocks) {
      assert.ok(block.includes('<w:t'), 'w:ins contains w:t');
    }

    ws1.cleanup();
    ws2.cleanup();
  });

  it('result can be saved to a valid docx file', () => {
    const copy1 = freshCopy('diff-save-1');
    const copy2 = freshCopy('diff-save-2');
    const outPath = path.join(OUTPUT_DIR, 'diff-output.docx');
    const ws1 = Workspace.open(copy1);
    const ws2 = Workspace.open(copy2);

    // Add a paragraph to ws2
    const paras2 = xmlMod.findParagraphs(ws2.docXml);
    const insertAfter = paras2[0];
    const newPara = '<w:p><w:r><w:t xml:space="preserve">Diff test paragraph.</w:t></w:r></w:p>';
    ws2.docXml = ws2.docXml.slice(0, insertAfter.end) + newPara + ws2.docXml.slice(insertAfter.end);

    Diff.compare(ws1, ws2, { author: 'Test', date: '2026-03-27T00:00:00Z' });

    const result = ws1.save(outPath);
    assert.ok(result.fileSize > 0, 'output file is not empty');
    assert.ok(fs.existsSync(outPath), 'output file exists');

    ws2.cleanup();
  });
});
