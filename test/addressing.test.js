/**
 * addressing.test.js -- Tests for v0.3 stable addressing system
 *
 * Covers:
 *   - DocMap.generate (map), DocMap.injectParaIds, DocMap.find, DocMap.structure, DocMap.explain
 *   - ParagraphHandle (doc.id()) with replace, delete, bold, italic, highlight, comment,
 *     footnote, replaceAt, insertAfter, insertBefore, remove
 *   - DocexEngine.map(), .id(), .afterHeading(), .afterText(), .find(), .structure(), .explain()
 *   - Stability: ids survive mutations, multiple operations don't interfere
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/addressing.test.js
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
// 1. DOCMAP - PARAID INJECTION
// ============================================================================

describe('paraId injection', () => {
  let DocMap, Workspace;
  before(() => {
    DocMap = require('../src/docmap').DocMap;
    Workspace = require('../src/workspace').Workspace;
  });

  it('injects paraIds into paragraphs that lack them', () => {
    const out = freshCopy('inject-ids');
    const ws = Workspace.open(out);

    // All paragraphs should now have paraIds (injected on open)
    const re = /w14:paraId="([^"]+)"/g;
    const ids = [];
    let m;
    while ((m = re.exec(ws.docXml)) !== null) {
      ids.push(m[1]);
    }

    // The fixture has 9 paragraphs
    assert.ok(ids.length >= 9, `Expected at least 9 paraIds, got ${ids.length}`);
    ws.cleanup();
  });

  it('every paragraph has a unique id', () => {
    const out = freshCopy('inject-unique');
    const ws = Workspace.open(out);

    const re = /w14:paraId="([^"]+)"/g;
    const ids = new Set();
    let m;
    while ((m = re.exec(ws.docXml)) !== null) {
      assert.ok(!ids.has(m[1]), `Duplicate paraId: ${m[1]}`);
      ids.add(m[1]);
    }
    ws.cleanup();
  });

  it('paraIds are 8-char uppercase hex', () => {
    const out = freshCopy('inject-format');
    const ws = Workspace.open(out);

    const re = /w14:paraId="([^"]+)"/g;
    let m;
    while ((m = re.exec(ws.docXml)) !== null) {
      assert.match(m[1], /^[0-9A-F]{8}$/, `paraId should be 8 uppercase hex chars: ${m[1]}`);
    }
    ws.cleanup();
  });

  it('does not duplicate ids when called twice', () => {
    const out = freshCopy('inject-twice');
    const ws = Workspace.open(out);
    const countBefore = countMatches(ws.docXml, /w14:paraId="/g);

    // Call inject again - should not add more
    const added = DocMap.injectParaIds(ws);
    assert.equal(added, 0, 'second injection should add 0');
    const countAfter = countMatches(ws.docXml, /w14:paraId="/g);
    assert.equal(countBefore, countAfter);
    ws.cleanup();
  });
});

// ============================================================================
// 2. DOCMAP - MAP GENERATION
// ============================================================================

describe('map()', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('returns structured tree with sections, paragraphs', async () => {
    const doc = docex(FIXTURE);
    const map = await doc.map();

    assert.ok(map.sections.length >= 4, 'should have at least 4 sections');
    assert.ok(map.allParagraphs.length >= 9, 'should have at least 9 paragraphs');

    // Check section structure
    const intro = map.sections.find(s => s.heading.text === 'Introduction');
    assert.ok(intro, 'should have Introduction section');
    assert.equal(intro.heading.level, 1);
    assert.ok(intro.paragraphs.length >= 1, 'Introduction should have body paragraphs');

    doc.discard();
  });

  it('every paragraph in map has an id', async () => {
    const doc = docex(FIXTURE);
    const map = await doc.map();

    for (const p of map.allParagraphs) {
      assert.ok(p.id, `Paragraph "${p.text.slice(0, 30)}..." should have an id`);
      assert.match(p.id, /^[0-9A-F]{8}$/, `paraId should be 8 uppercase hex: ${p.id}`);
    }

    doc.discard();
  });

  it('heading paragraphs have type "heading" and level', async () => {
    const doc = docex(FIXTURE);
    const map = await doc.map();

    const headings = map.allParagraphs.filter(p => p.type === 'heading');
    assert.ok(headings.length >= 4, 'should have at least 4 headings');

    for (const h of headings) {
      assert.ok(h.level >= 1 && h.level <= 9, `heading level ${h.level} should be 1-9`);
    }

    doc.discard();
  });

  it('body paragraphs have type "body"', async () => {
    const doc = docex(FIXTURE);
    const map = await doc.map();

    const bodyParas = map.allParagraphs.filter(p => p.type === 'body');
    assert.ok(bodyParas.length >= 4, 'should have at least 4 body paragraphs');

    // Body paragraphs should have text
    for (const p of bodyParas) {
      assert.ok(p.text.length > 0, 'body paragraph should have text');
    }

    doc.discard();
  });

  it('sections contain paragraphs', async () => {
    const doc = docex(FIXTURE);
    const map = await doc.map();

    const methods = map.sections.find(s => s.heading.text === 'Methods');
    assert.ok(methods, 'should have Methods section');
    assert.ok(methods.paragraphs.length >= 1, 'Methods should have paragraphs');

    // Check that paragraph text is correct
    const methodsPara = methods.paragraphs[0];
    assert.ok(methodsPara.text.includes('268,635') || methodsPara.text.includes('automated'),
      'Methods paragraph should contain expected text');

    doc.discard();
  });
});

// ============================================================================
// 3. DOC.ID() - PARAGRAPH HANDLE
// ============================================================================

describe('doc.id() - ParagraphHandle', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('finds the correct paragraph by id', async () => {
    const out = freshCopy('handle-find');
    const doc = docex(out);
    const map = await doc.map();

    // Pick the Methods body paragraph
    const methods = map.sections.find(s => s.heading.text === 'Methods');
    const targetPara = methods.paragraphs[0];

    const handle = doc.id(targetPara.id);
    assert.equal(handle.id, targetPara.id);
    assert.ok(handle.text.includes('268,635') || handle.text.includes('automated'));

    doc.discard();
  });

  it('replace() works correctly', async () => {
    const out = freshCopy('handle-replace');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    // Find the paragraph with "268,635"
    const para = map.allParagraphs.find(p => p.text.includes('268,635'));
    assert.ok(para, 'should find paragraph with 268,635');

    doc.id(para.id).replace('268,635', '300,000');
    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('300,000'), 'replacement applied');
    assert.ok(savedXml.includes('268,635'), 'old text preserved in w:del for tracking');
  });

  it('replace() survives after another insertion shifts positions', async () => {
    const out = freshCopy('handle-shift');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    // Get two paragraphs from different sections
    const introSection = map.sections.find(s => s.heading.text === 'Introduction');
    const methodsSection = map.sections.find(s => s.heading.text === 'Methods');
    const methodsPara = methodsSection.paragraphs[0];

    // First: insert a huge paragraph after Introduction heading (shifts everything)
    doc.id(introSection.heading.id).insertAfter(
      'This is a brand new paragraph that shifts all subsequent positions by a lot of characters.'
    );

    // Now: replace in Methods paragraph (should still work because id-based)
    doc.id(methodsPara.id).replace('268,635', '300,000');

    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('brand new paragraph'), 'insertion present');
    assert.ok(savedXml.includes('300,000'), 'replacement after shift succeeded');
  });

  it('insertAfter() returns a new valid id', async () => {
    const out = freshCopy('handle-insertafter');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    const methods = map.sections.find(s => s.heading.text === 'Methods');
    const newId = doc.id(methods.heading.id).insertAfter('New methodology paragraph.');

    assert.ok(newId, 'should return a paraId');
    assert.match(newId, /^[0-9A-F]{8}$/, 'new paraId should be 8 uppercase hex');

    // The new paragraph should be findable by its id
    const newHandle = doc.id(newId);
    assert.equal(newHandle.text, 'New methodology paragraph.');

    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('New methodology paragraph'), 'inserted text present');
    assert.ok(savedXml.includes(`w14:paraId="${newId}"`), 'new paraId in saved document');
  });

  it('insertBefore() returns a new valid id', async () => {
    const out = freshCopy('handle-insertbefore');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    const results = map.sections.find(s => s.heading.text === 'Results');
    const newId = doc.id(results.heading.id).insertBefore('Pre-results paragraph.');

    assert.ok(newId);
    assert.match(newId, /^[0-9A-F]{8}$/);

    const newHandle = doc.id(newId);
    assert.equal(newHandle.text, 'Pre-results paragraph.');

    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('Pre-results paragraph'));
    // Verify it comes before Results heading
    const newPos = savedXml.indexOf('Pre-results paragraph');
    const resultsPos = savedXml.indexOf('>Results<');
    assert.ok(newPos < resultsPos, 'inserted before Results');
  });

  it('comment() anchors to the specific paragraph', async () => {
    const out = freshCopy('handle-comment');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    const para = map.allParagraphs.find(p => p.text.includes('268,635'));
    doc.id(para.id).comment('Verify this number', { by: 'Reviewer 1' });

    await doc.save(out);

    const commentsXml = readDocxXml(out, 'word/comments.xml');
    assert.ok(commentsXml.includes('Verify this number'), 'comment text present');
    assert.ok(commentsXml.includes('Reviewer 1'), 'author attributed');

    const docXml = readDocxXml(out, 'word/document.xml');
    assert.ok(docXml.includes('commentRangeStart'), 'comment range in document');
  });

  it('replaceAt(start, end, text) works', async () => {
    const out = freshCopy('handle-replaceat');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    const para = map.allParagraphs.find(p => p.text.includes('268,635'));
    const text = doc.id(para.id).text;
    const start = text.indexOf('268,635');
    const end = start + '268,635'.length;

    doc.id(para.id).replaceAt(start, end, '300,000');

    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('300,000'), 'replaceAt applied');
  });

  it('handle.type returns correct type', async () => {
    const doc = docex(FIXTURE);
    const map = await doc.map();

    const heading = map.allParagraphs.find(p => p.type === 'heading');
    assert.equal(doc.id(heading.id).type, 'heading');

    const body = map.allParagraphs.find(p => p.type === 'body');
    assert.equal(doc.id(body.id).type, 'body');

    doc.discard();
  });

  it('handle.section returns section name', async () => {
    const doc = docex(FIXTURE);
    const map = await doc.map();

    const methodsSection = map.sections.find(s => s.heading.text === 'Methods');
    const methodsPara = methodsSection.paragraphs[0];

    assert.equal(doc.id(methodsPara.id).section, 'Methods');

    doc.discard();
  });

  it('bold() applies formatting within paragraph', async () => {
    const out = freshCopy('handle-bold');
    const doc = docex(out);
    doc.author('Fabio Votta').untracked();
    const map = await doc.map();

    const para = map.allParagraphs.find(p => p.text.includes('268,635'));
    doc.id(para.id).bold('268,635');

    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    // The paragraph containing 268,635 should now have bold in a run
    const paraMatch = savedXml.match(new RegExp(
      '<w:p[^>]*w14:paraId="' + para.id + '"[^>]*>[\\s\\S]*?</w:p>'
    ));
    assert.ok(paraMatch, 'should find the paragraph');
    assert.ok(paraMatch[0].includes('<w:b'), 'should have bold formatting');
  });

  it('multiple operations via ids do not interfere', async () => {
    const out = freshCopy('handle-multi');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    // Get paragraphs from different sections
    const introSection = map.sections.find(s => s.heading.text === 'Introduction');
    const methodsSection = map.sections.find(s => s.heading.text === 'Methods');
    const resultsSection = map.sections.find(s => s.heading.text === 'Results');

    const introPara = introSection.paragraphs[0];
    const methodsPara = methodsSection.paragraphs[0];
    const resultsPara = resultsSection.paragraphs[0];

    // Perform operations on each paragraph independently
    doc.id(introPara.id).replace('platform governance', 'platform regulation');
    doc.id(methodsPara.id).replace('268,635', '300,000');
    doc.id(resultsPara.id).replace('192 accounts', '200 accounts');

    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('platform regulation'), 'intro replacement applied');
    assert.ok(savedXml.includes('300,000'), 'methods replacement applied');
    assert.ok(savedXml.includes('200 accounts'), 'results replacement applied');
    assert.equal(countMatches(savedXml, /<w:del /g), 3, 'three tracked deletions');
  });

  it('remove() deletes the paragraph', async () => {
    const out = freshCopy('handle-remove');
    const doc = docex(out);
    doc.author('Fabio Votta').untracked();
    const map = await doc.map();

    const countBefore = map.allParagraphs.length;
    const discussion = map.sections.find(s => s.heading.text === 'Discussion');
    const targetPara = discussion.paragraphs[0];

    doc.id(targetPara.id).remove({ tracked: false });
    await doc.save(out);

    // Reopen and count
    const doc2 = docex(out);
    const map2 = await doc2.map();
    assert.ok(map2.allParagraphs.length < countBefore, 'paragraph removed');
    doc2.discard();
  });

  it('delete() wraps text in w:del within the paragraph', async () => {
    const out = freshCopy('handle-delete');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    const para = map.allParagraphs.find(p => p.text.includes('268,635'));
    doc.id(para.id).delete('268,635');

    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('<w:del'), 'deletion tracked');
    assert.ok(savedXml.includes('268,635'), 'deleted text in delText');
  });
});

// ============================================================================
// 4. FIND, STRUCTURE, EXPLAIN
// ============================================================================

describe('find()', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('returns matches with section context', async () => {
    const doc = docex(FIXTURE);
    const results = await doc.find('268,635');

    assert.ok(results.length >= 1, 'should find at least one match');
    const match = results[0];
    assert.ok(match.id, 'match should have id');
    assert.equal(match.section, 'Methods', 'match should be in Methods section');
    assert.ok(match.context.includes('268,635'), 'context should contain search text');

    doc.discard();
  });

  it('returns empty array for no matches', async () => {
    const doc = docex(FIXTURE);
    const results = await doc.find('xyzzy nonexistent text');
    assert.equal(results.length, 0);
    doc.discard();
  });

  it('returns multiple matches when text appears in multiple paragraphs', async () => {
    const doc = docex(FIXTURE);
    // "platform" appears in multiple paragraphs
    const results = await doc.find('platform');
    assert.ok(results.length >= 2, 'should find multiple matches');
    doc.discard();
  });
});

describe('structure()', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('returns a tree string', async () => {
    const doc = docex(FIXTURE);
    const tree = await doc.structure();

    assert.ok(typeof tree === 'string');
    assert.ok(tree.includes('Introduction (H1)'));
    assert.ok(tree.includes('Methods (H1)'));
    assert.ok(tree.includes('Results (H1)'));
    assert.ok(tree.includes('Discussion (H1)'));
    assert.ok(tree.includes('paragraph'), 'should mention paragraphs');

    doc.discard();
  });
});

describe('explain()', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('shows run info for found text', async () => {
    const doc = docex(FIXTURE);
    const info = await doc.explain('268,635');

    assert.ok(typeof info === 'string');
    assert.ok(info.includes('Paragraph'), 'should mention paragraph');
    assert.ok(info.includes('Run'), 'should show runs');
    assert.ok(info.includes('268,635'), 'should show the text');

    doc.discard();
  });

  it('reports not found for missing text', async () => {
    const doc = docex(FIXTURE);
    const info = await doc.explain('this text does not exist');
    assert.ok(info.includes('not found'));
    doc.discard();
  });
});

// ============================================================================
// 5. AFTERHEADING / AFTERTEXT
// ============================================================================

describe('afterHeading()', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('finds heading and allows insert', async () => {
    const out = freshCopy('afterheading');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.afterHeading('Methods').insert('New methodology paragraph via afterHeading.');
    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('New methodology paragraph via afterHeading'));
    const methodsPos = savedXml.indexOf('Methods');
    const newPos = savedXml.indexOf('New methodology paragraph via afterHeading');
    assert.ok(newPos > methodsPos, 'inserted after Methods heading');
  });
});

describe('afterText()', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('finds body text and allows insert', async () => {
    const out = freshCopy('aftertext');
    const doc = docex(out);
    doc.author('Fabio Votta');
    doc.afterText('268,635').insert('Additional results paragraph.');
    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('Additional results paragraph'));
    const anchorPos = savedXml.indexOf('268,635');
    const newPos = savedXml.indexOf('Additional results paragraph');
    assert.ok(newPos > anchorPos, 'inserted after body text');
  });
});

// ============================================================================
// 6. ROUND-TRIP: SAVE AND REOPEN
// ============================================================================

describe('addressing round-trip', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('paraIds are preserved across save and reopen', async () => {
    const out = freshCopy('roundtrip-ids');
    const doc = docex(out);
    const map = await doc.map();
    const originalIds = map.allParagraphs.map(p => p.id);

    await doc.save(out);

    // Reopen and check
    const doc2 = docex(out);
    const map2 = await doc2.map();
    const newIds = map2.allParagraphs.map(p => p.id);

    // All original ids should be present (there may be more if paragraphs were added)
    for (const id of originalIds) {
      assert.ok(newIds.includes(id), `Original paraId ${id} should survive save/reopen`);
    }

    doc2.discard();
  });

  it('operations on saved document work with original ids', async () => {
    const out = freshCopy('roundtrip-ops');

    // First pass: get the map
    const doc1 = docex(out);
    doc1.author('Fabio Votta');
    const map = await doc1.map();
    const para = map.allParagraphs.find(p => p.text.includes('268,635'));
    const savedId = para.id;

    // Save without changes (just to persist paraIds)
    await doc1.save(out);

    // Second pass: use the saved id
    const doc2 = docex(out);
    doc2.author('Simon Kruschinski');
    const map2 = await doc2.map();

    // The same paraId should still exist
    const para2 = map2.allParagraphs.find(p => p.id === savedId);
    assert.ok(para2, 'saved paraId should persist after reopen');
    assert.ok(para2.text.includes('268,635'), 'text should be the same');

    // Operate on it
    doc2.id(savedId).replace('268,635', '300,000');
    await doc2.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(savedXml.includes('300,000'), 'replacement applied in second pass');
  });

  it('document integrity is preserved after addressing operations', async () => {
    const out = freshCopy('roundtrip-integrity');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    // Perform multiple operations
    const intro = map.sections.find(s => s.heading.text === 'Introduction');
    const methods = map.sections.find(s => s.heading.text === 'Methods');

    doc.id(intro.heading.id).insertAfter('New intro paragraph.');
    doc.id(methods.paragraphs[0].id).replace('268,635', '300,000');

    const result = await doc.save(out);

    assert.ok(result.verified, 'document should pass verification');
    assert.ok(result.paragraphCount >= map.allParagraphs.length,
      'paragraph count should not decrease');
    assert.ok(result.fileSize > 0, 'file size should be positive');
  });
});

// ============================================================================
// 7. EDGE CASES
// ============================================================================

describe('addressing edge cases', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('throws on invalid paraId', async () => {
    const doc = docex(FIXTURE);
    await doc.map(); // ensure workspace is open

    assert.throws(() => {
      doc.id('DEADBEEF').text;
    }, /not found/i);

    doc.discard();
  });

  it('handles paragraph with XML entities', async () => {
    const out = freshCopy('edge-entities-addr');
    const doc = docex(out);
    doc.author('Fabio Votta');
    const map = await doc.map();

    // The fixture has "Meta's" with apostrophe entity
    const para = map.allParagraphs.find(p => p.text.includes("Meta's"));
    assert.ok(para, 'should find paragraph with entity-decoded text');

    // The handle should have decoded text
    const handle = doc.id(para.id);
    assert.ok(handle.text.includes("Meta's"), 'handle.text should have decoded entities');

    doc.discard();
  });

  it('italic() applies within paragraph', async () => {
    const out = freshCopy('handle-italic');
    const doc = docex(out);
    doc.author('Fabio Votta').untracked();
    const map = await doc.map();

    const para = map.allParagraphs.find(p => p.text.includes('electoral transparency'));
    doc.id(para.id).italic('electoral transparency');

    await doc.save(out);

    const savedXml = readDocxXml(out, 'word/document.xml');
    const paraMatch = savedXml.match(new RegExp(
      '<w:p[^>]*w14:paraId="' + para.id + '"[^>]*>[\\s\\S]*?</w:p>'
    ));
    assert.ok(paraMatch, 'should find the paragraph');
    assert.ok(paraMatch[0].includes('<w:i'), 'should have italic formatting');
  });
});
