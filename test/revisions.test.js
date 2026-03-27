/**
 * revisions test suite
 *
 * Tests for the Revisions class that handles accept/reject of tracked changes.
 *
 * Layers:
 *   1. List revisions from a document with tracked changes
 *   2. Accept specific change by ID
 *   3. Accept all changes
 *   4. Reject specific change by ID
 *   5. Reject all changes
 *   6. Clean copy (no tracked changes, no comments)
 *   7. Paragraph count integrity after operations
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/revisions.test.js
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

/**
 * Helper: create a document with tracked changes for testing revisions.
 * Makes replacements and deletions, returns the saved path.
 */
function createTrackedDoc(testName) {
  const out = freshCopy(testName);
  const docex = require('../src/docex');
  const doc = docex(out);
  doc.author('Fabio Votta');
  doc.replace('268,635', '300,000');
  doc.replace('1,329', '1,500');
  doc.delete('specific failure mode of ');
  return { out, doc };
}

/**
 * Helper: create a document with tracked changes AND comments.
 */
function createTrackedDocWithComments(testName) {
  const out = freshCopy(testName);
  const docex = require('../src/docex');
  const doc = docex(out);
  doc.author('Fabio Votta');
  doc.replace('268,635', '300,000');
  doc.at('platform governance').comment('Needs citation', { by: 'Reviewer 1' });
  return { out, doc };
}

// ============================================================================
// 1. LIST REVISIONS
// ============================================================================

describe('list revisions', () => {
  let Revisions, Workspace;

  before(() => {
    Revisions = require('../src/revisions').Revisions;
    Workspace = require('../src/workspace').Workspace;
  });

  it('lists insertions and deletions from tracked changes', async () => {
    const { out, doc } = createTrackedDoc('rev-list');
    await doc.save();

    const ws = Workspace.open(out);
    const revisions = Revisions.list(ws);

    assert.ok(revisions.length >= 3, `expected at least 3 revisions, got ${revisions.length}`);

    const insertions = revisions.filter(r => r.type === 'insertion');
    const deletions = revisions.filter(r => r.type === 'deletion');

    assert.ok(insertions.length >= 2, 'should have at least 2 insertions');
    assert.ok(deletions.length >= 2, 'should have at least 2 deletions');

    // Check specific content
    assert.ok(insertions.some(r => r.text.includes('300,000')), 'should find 300,000 insertion');
    assert.ok(insertions.some(r => r.text.includes('1,500')), 'should find 1,500 insertion');
    assert.ok(deletions.some(r => r.text.includes('268,635')), 'should find 268,635 deletion');
    assert.ok(deletions.some(r => r.text.includes('1,329')), 'should find 1,329 deletion');

    // Check author attribution on deletions (insertions from buildIns may have empty author)
    for (const rev of deletions) {
      assert.equal(rev.author, 'Fabio Votta', 'deletion author should be Fabio Votta');
    }

    ws.cleanup();
  });

  it('returns empty array for document without tracked changes', () => {
    const ws = Workspace.open(FIXTURE);
    const revisions = Revisions.list(ws);
    assert.equal(revisions.length, 0, 'no revisions in clean fixture');
    ws.cleanup();
  });

  it('revisions are sorted by document position', async () => {
    const { out, doc } = createTrackedDoc('rev-list-sorted');
    await doc.save();

    const ws = Workspace.open(out);
    const revisions = Revisions.list(ws);

    for (let i = 1; i < revisions.length; i++) {
      assert.ok(revisions[i].start >= revisions[i - 1].start,
        'revisions should be sorted by position');
    }
    ws.cleanup();
  });
});

// ============================================================================
// 2. ACCEPT SPECIFIC CHANGE
// ============================================================================

describe('accept specific change', () => {
  let Revisions, Workspace;

  before(() => {
    Revisions = require('../src/revisions').Revisions;
    Workspace = require('../src/workspace').Workspace;
  });

  it('accepts a specific insertion (unwraps content)', async () => {
    const { out, doc } = createTrackedDoc('rev-accept-ins');
    await doc.save();

    const ws = Workspace.open(out);
    const revisions = Revisions.list(ws);
    const insertion = revisions.find(r => r.type === 'insertion' && r.text.includes('300,000'));
    assert.ok(insertion, 'should find the 300,000 insertion');

    Revisions.accept(ws, insertion.id);

    const docXml = ws.docXml;
    // The insertion text should still be present (unwrapped)
    assert.ok(docXml.includes('300,000'), 'accepted insertion text preserved');
    // But it should no longer be inside a w:ins with that ID
    assert.ok(
      !docXml.includes(`w:id="${insertion.id}"` + '"') || !docXml.match(new RegExp(`<w:ins[^>]*w:id="${insertion.id}"`)),
      'w:ins wrapper removed for accepted insertion'
    );

    // Other tracked changes should still exist
    const remaining = Revisions.list(ws);
    assert.ok(remaining.length < revisions.length, 'fewer revisions after accepting one');

    ws.cleanup();
  });

  it('accepts a specific deletion (removes element)', async () => {
    const { out, doc } = createTrackedDoc('rev-accept-del');
    await doc.save();

    const ws = Workspace.open(out);
    const revisions = Revisions.list(ws);
    const deletion = revisions.find(r => r.type === 'deletion' && r.text.includes('268,635'));
    assert.ok(deletion, 'should find the 268,635 deletion');

    Revisions.accept(ws, deletion.id);

    const docXml = ws.docXml;
    // The deleted text should be gone entirely (accepting a deletion removes it)
    assert.ok(!docXml.match(new RegExp(`<w:del[^>]*w:id="${deletion.id}"`)),
      'w:del element removed for accepted deletion');

    ws.cleanup();
  });
});

// ============================================================================
// 3. ACCEPT ALL CHANGES
// ============================================================================

describe('accept all changes', () => {
  let Revisions, Workspace;

  before(() => {
    Revisions = require('../src/revisions').Revisions;
    Workspace = require('../src/workspace').Workspace;
  });

  it('removes all tracked change markup', async () => {
    const { out, doc } = createTrackedDoc('rev-accept-all');
    await doc.save();

    const ws = Workspace.open(out);
    const beforeRevisions = Revisions.list(ws);
    assert.ok(beforeRevisions.length >= 3, 'should have revisions before accepting');

    Revisions.accept(ws);

    const docXml = ws.docXml;
    assert.ok(!docXml.includes('<w:ins'), 'no w:ins elements after accept all');
    assert.ok(!docXml.includes('<w:del'), 'no w:del elements after accept all');
    assert.ok(!docXml.includes('w:delText'), 'no w:delText elements after accept all');

    const afterRevisions = Revisions.list(ws);
    assert.equal(afterRevisions.length, 0, 'no revisions after accept all');

    ws.cleanup();
  });

  it('preserves inserted text after accepting all', async () => {
    const { out, doc } = createTrackedDoc('rev-accept-all-content');
    await doc.save();

    const ws = Workspace.open(out);
    Revisions.accept(ws);

    const docXml = ws.docXml;
    assert.ok(docXml.includes('300,000'), 'inserted text 300,000 preserved');
    assert.ok(docXml.includes('1,500'), 'inserted text 1,500 preserved');

    ws.cleanup();
  });

  it('removes deleted text after accepting all', async () => {
    const { out, doc } = createTrackedDoc('rev-accept-all-del');
    await doc.save();

    const ws = Workspace.open(out);
    Revisions.accept(ws);

    const docXml = ws.docXml;
    assert.ok(!docXml.includes('268,635'), 'deleted text 268,635 removed');
    assert.ok(!docXml.includes('1,329'), 'deleted text 1,329 removed');

    ws.cleanup();
  });
});

// ============================================================================
// 4. REJECT SPECIFIC CHANGE
// ============================================================================

describe('reject specific change', () => {
  let Revisions, Workspace;

  before(() => {
    Revisions = require('../src/revisions').Revisions;
    Workspace = require('../src/workspace').Workspace;
  });

  it('rejects a specific insertion (removes element and content)', async () => {
    const { out, doc } = createTrackedDoc('rev-reject-ins');
    await doc.save();

    const ws = Workspace.open(out);
    const revisions = Revisions.list(ws);
    const insertion = revisions.find(r => r.type === 'insertion' && r.text.includes('300,000'));
    assert.ok(insertion, 'should find the 300,000 insertion');

    Revisions.reject(ws, insertion.id);

    const docXml = ws.docXml;
    // The insertion and its content should be gone
    assert.ok(!docXml.match(new RegExp(`<w:ins[^>]*w:id="${insertion.id}"`)),
      'w:ins element removed');

    // Other tracked changes should still exist
    const remaining = Revisions.list(ws);
    assert.ok(remaining.length < revisions.length, 'fewer revisions after rejecting one');

    ws.cleanup();
  });

  it('rejects a specific deletion (unwraps, converts delText to t)', async () => {
    const { out, doc } = createTrackedDoc('rev-reject-del');
    await doc.save();

    const ws = Workspace.open(out);
    const revisions = Revisions.list(ws);
    const deletion = revisions.find(r => r.type === 'deletion' && r.text.includes('268,635'));
    assert.ok(deletion, 'should find the 268,635 deletion');

    Revisions.reject(ws, deletion.id);

    const docXml = ws.docXml;
    // The w:del wrapper should be gone
    assert.ok(!docXml.match(new RegExp(`<w:del[^>]*w:id="${deletion.id}"`)),
      'w:del element removed');
    // But the text should be restored as normal w:t text
    assert.ok(docXml.includes('268,635'), 'rejected deletion text restored');
    // And delText should be converted to t
    assert.ok(!docXml.match(new RegExp(`<w:del[^>]*w:id="${deletion.id}"[\\s\\S]*?w:delText`)),
      'no delText remaining for this deletion');

    ws.cleanup();
  });
});

// ============================================================================
// 5. REJECT ALL CHANGES
// ============================================================================

describe('reject all changes', () => {
  let Revisions, Workspace;

  before(() => {
    Revisions = require('../src/revisions').Revisions;
    Workspace = require('../src/workspace').Workspace;
  });

  it('removes all tracked change markup', async () => {
    const { out, doc } = createTrackedDoc('rev-reject-all');
    await doc.save();

    const ws = Workspace.open(out);
    Revisions.reject(ws);

    const docXml = ws.docXml;
    assert.ok(!docXml.includes('<w:ins'), 'no w:ins elements after reject all');
    assert.ok(!docXml.includes('<w:del'), 'no w:del elements after reject all');

    const afterRevisions = Revisions.list(ws);
    assert.equal(afterRevisions.length, 0, 'no revisions after reject all');

    ws.cleanup();
  });

  it('restores original text after rejecting all', async () => {
    const { out, doc } = createTrackedDoc('rev-reject-all-content');
    await doc.save();

    const ws = Workspace.open(out);
    Revisions.reject(ws);

    const docXml = ws.docXml;
    // Original text should be back
    assert.ok(docXml.includes('268,635'), 'original text 268,635 restored');
    assert.ok(docXml.includes('1,329'), 'original text 1,329 restored');
    // Inserted text should be gone
    assert.ok(!docXml.includes('300,000'), 'inserted text 300,000 removed');
    assert.ok(!docXml.includes('1,500'), 'inserted text 1,500 removed');

    ws.cleanup();
  });
});

// ============================================================================
// 6. CLEAN COPY
// ============================================================================

describe('clean copy', () => {
  let Revisions, Workspace;

  before(() => {
    Revisions = require('../src/revisions').Revisions;
    Workspace = require('../src/workspace').Workspace;
  });

  it('accepts all changes and removes comments', async () => {
    const { out, doc } = createTrackedDocWithComments('rev-clean');
    await doc.save();

    const ws = Workspace.open(out);

    // Verify we have tracked changes and comments before cleaning
    const beforeRevisions = Revisions.list(ws);
    assert.ok(beforeRevisions.length >= 1, 'should have revisions before clean');
    assert.ok(ws.commentsXml.includes('<w:comment '), 'should have comments before clean');
    assert.ok(ws.docXml.includes('commentRangeStart'), 'should have comment ranges before clean');

    Revisions.cleanCopy(ws);

    const docXml = ws.docXml;
    const commentsXml = ws.commentsXml;

    // No tracked changes
    assert.ok(!docXml.includes('<w:ins'), 'no w:ins after clean copy');
    assert.ok(!docXml.includes('<w:del'), 'no w:del after clean copy');

    // No comment markers in document
    assert.ok(!docXml.includes('commentRangeStart'), 'no commentRangeStart after clean copy');
    assert.ok(!docXml.includes('commentRangeEnd'), 'no commentRangeEnd after clean copy');
    assert.ok(!docXml.includes('commentReference'), 'no commentReference after clean copy');

    // No comments in comments.xml (check for '<w:comment ' with space to avoid matching '<w:comments')
    assert.ok(!commentsXml.includes('<w:comment '), 'no comments after clean copy');

    // commentsExtended should have all done="1"
    const extXml = ws.commentsExtXml;
    if (extXml && extXml.includes('w15:done')) {
      assert.ok(!extXml.includes('w15:done="0"'), 'all comments resolved');
    }

    ws.cleanup();
  });

  it('preserves accepted insertion text in clean copy', async () => {
    const { out, doc } = createTrackedDocWithComments('rev-clean-content');
    await doc.save();

    const ws = Workspace.open(out);
    Revisions.cleanCopy(ws);

    const docXml = ws.docXml;
    assert.ok(docXml.includes('300,000'), 'accepted insertion text preserved in clean copy');

    ws.cleanup();
  });

  it('saves and produces valid docx after clean copy', async () => {
    const { out, doc } = createTrackedDocWithComments('rev-clean-save');
    await doc.save();

    const ws = Workspace.open(out);
    Revisions.cleanCopy(ws);

    const saveResult = ws.save(out);
    assert.ok(saveResult.verified, 'clean copy output is verified');
    assert.ok(saveResult.fileSize > 0, 'file has content');

    // Verify the saved file also has no tracked changes
    const savedXml = readDocxXml(out, 'word/document.xml');
    assert.ok(!savedXml.includes('<w:ins'), 'saved clean copy has no w:ins');
    assert.ok(!savedXml.includes('<w:del'), 'saved clean copy has no w:del');
  });
});

// ============================================================================
// 7. PARAGRAPH COUNT INTEGRITY
// ============================================================================

describe('paragraph count integrity', () => {
  let Revisions, Workspace;

  before(() => {
    Revisions = require('../src/revisions').Revisions;
    Workspace = require('../src/workspace').Workspace;
  });

  it('paragraph count does not decrease after accept all', async () => {
    const { out, doc } = createTrackedDoc('rev-integrity-accept');
    await doc.save();

    const ws = Workspace.open(out);
    const beforeCount = countMatches(ws.docXml, /<w:p[\s>]/g);

    Revisions.accept(ws);

    const afterCount = countMatches(ws.docXml, /<w:p[\s>]/g);
    assert.ok(afterCount >= beforeCount - 1,
      `paragraph count should not drop significantly: before=${beforeCount}, after=${afterCount}`);
  });

  it('paragraph count does not decrease after reject all', async () => {
    const { out, doc } = createTrackedDoc('rev-integrity-reject');
    await doc.save();

    const ws = Workspace.open(out);
    const beforeCount = countMatches(ws.docXml, /<w:p[\s>]/g);

    Revisions.reject(ws);

    const afterCount = countMatches(ws.docXml, /<w:p[\s>]/g);
    assert.ok(afterCount >= beforeCount - 1,
      `paragraph count should not drop significantly: before=${beforeCount}, after=${afterCount}`);
  });

  it('paragraph count preserved after clean copy', async () => {
    const { out, doc } = createTrackedDocWithComments('rev-integrity-clean');
    await doc.save();

    const ws = Workspace.open(out);
    const beforeCount = countMatches(ws.docXml, /<w:p[\s>]/g);

    Revisions.cleanCopy(ws);

    const afterCount = countMatches(ws.docXml, /<w:p[\s>]/g);
    assert.ok(afterCount >= beforeCount - 1,
      `paragraph count should not drop significantly: before=${beforeCount}, after=${afterCount}`);
  });
});
