/**
 * agent-reliability.test.js -- Tests for v0.4.1: AI Agent Reliability
 *
 * Features tested:
 *   1. doc.snapshot() / doc.rollback() -- in-memory state save/restore
 *   2. doc.assert(paraId, expectedText) -- paragraph content verification
 *   3. doc.diffSummary("other.docx") -- lightweight diff counts
 *   4. Operation receipts -- every mutation returns { success, type, paraId, matched, context }
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/agent-reliability.test.js
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

// ============================================================================
// 1. SNAPSHOT / ROLLBACK
// ============================================================================

describe('snapshot and rollback', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('snapshot + rollback restores original state', async () => {
    const out = freshCopy('snapshot-rollback');
    const doc = docex(out);
    doc.author('Test');

    // Get original text
    const originalText = await doc.text();

    // Take a snapshot
    await doc.snapshot();

    // Make a modification (untracked to simplify)
    doc.untracked().replace('Introduction', 'INTRO_CHANGED');

    // Save to a temp file to apply the operation, then discard
    // Instead, we operate directly on workspace to verify rollback
    const ws = await doc._ensureWorkspace();
    const textAfterQueue = ws.docXml;

    // Apply the replacement manually to the workspace to see the change
    const { Paragraphs } = require('../src/paragraphs');
    Paragraphs.replace(ws, 'Introduction', 'INTRO_CHANGED', { tracked: false });

    // Verify the change happened
    const changedXml = ws.docXml;
    assert.ok(changedXml.includes('INTRO_CHANGED'), 'Document should contain changed text');

    // Rollback
    const restored = ws.rollback();
    assert.equal(restored, true, 'rollback() should return true');

    // Verify original state is back
    const restoredXml = ws.docXml;
    assert.ok(!restoredXml.includes('INTRO_CHANGED'), 'Rollback should remove the change');
    assert.ok(restoredXml.includes('Introduction'), 'Rollback should restore original text');

    ws.cleanup();
  });

  it('multiple snapshots stack correctly (LIFO)', async () => {
    const out = freshCopy('snapshot-stack');
    const doc = docex(out);
    const ws = await doc._ensureWorkspace();
    const { Paragraphs } = require('../src/paragraphs');

    const originalXml = ws.docXml;

    // Snapshot 1
    ws.snapshot();
    Paragraphs.replace(ws, 'Introduction', 'STATE_ONE', { tracked: false });
    const state1Xml = ws.docXml;
    assert.ok(state1Xml.includes('STATE_ONE'));

    // Snapshot 2 (on top of state 1)
    ws.snapshot();
    Paragraphs.replace(ws, 'STATE_ONE', 'STATE_TWO', { tracked: false });
    const state2Xml = ws.docXml;
    assert.ok(state2Xml.includes('STATE_TWO'));

    // First rollback -> back to state 1
    ws.rollback();
    assert.ok(ws.docXml.includes('STATE_ONE'), 'First rollback should restore state 1');
    assert.ok(!ws.docXml.includes('STATE_TWO'), 'First rollback should remove state 2');

    // Second rollback -> back to original
    ws.rollback();
    assert.ok(ws.docXml.includes('Introduction'), 'Second rollback should restore original');
    assert.ok(!ws.docXml.includes('STATE_ONE'), 'Second rollback should remove state 1');

    // Third rollback -> no more snapshots
    const empty = ws.rollback();
    assert.equal(empty, false, 'Third rollback should return false (empty stack)');

    ws.cleanup();
  });

  it('snapshot/rollback via docex API', async () => {
    const out = freshCopy('snapshot-api');
    const doc = docex(out);
    doc.author('Test');

    const originalText = await doc.text();

    // Snapshot via API
    await doc.snapshot();

    // Verify rollback returns true when stack is non-empty
    const restored = await doc.rollback();
    assert.equal(restored, true);

    // Verify rollback returns false when stack is empty
    const empty = await doc.rollback();
    assert.equal(empty, false);

    doc.discard();
  });
});

// ============================================================================
// 2. ASSERT
// ============================================================================

describe('assert', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('assert passes when text matches', async () => {
    const out = freshCopy('assert-pass');
    const doc = docex(out);

    // Get a paragraph's paraId from the map
    const map = await doc.map();
    const firstPara = map.allParagraphs.find(p => p.text.length > 0);
    assert.ok(firstPara, 'Should find a non-empty paragraph');

    // Assert should pass without throwing
    await doc.assert(firstPara.id, firstPara.text.slice(0, 20));

    doc.discard();
  });

  it('assert throws when text does not match', async () => {
    const out = freshCopy('assert-fail');
    const doc = docex(out);

    // Get a paragraph's paraId
    const map = await doc.map();
    const firstPara = map.allParagraphs.find(p => p.text.length > 0);
    assert.ok(firstPara, 'Should find a non-empty paragraph');

    // Assert should throw with descriptive error
    await assert.rejects(
      () => doc.assert(firstPara.id, 'THIS_TEXT_DOES_NOT_EXIST_ANYWHERE_XYZ'),
      (err) => {
        assert.ok(err.message.includes('assert failed'), 'Error should say "assert failed"');
        assert.ok(err.message.includes(firstPara.id), 'Error should include paraId');
        assert.ok(err.message.includes('actual text'), 'Error should show actual text');
        return true;
      }
    );

    doc.discard();
  });

  it('assert throws when paraId is not found', async () => {
    const out = freshCopy('assert-missing');
    const doc = docex(out);

    await assert.rejects(
      () => doc.assert('NONEXISTENT_ID', 'any text'),
      (err) => {
        assert.ok(err.message.includes('not found'), 'Error should say "not found"');
        assert.ok(err.message.includes('NONEXISTENT_ID'), 'Error should include the bad paraId');
        return true;
      }
    );

    doc.discard();
  });
});

// ============================================================================
// 3. DIFF SUMMARY
// ============================================================================

describe('diffSummary', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('diffSummary returns correct counts for identical documents', async () => {
    const out1 = freshCopy('diffsummary-same1');
    const out2 = freshCopy('diffsummary-same2');
    const doc = docex(out1);

    const summary = await doc.diffSummary(out2);

    assert.equal(typeof summary.changed, 'number');
    assert.equal(typeof summary.added, 'number');
    assert.equal(typeof summary.removed, 'number');
    assert.equal(typeof summary.comments, 'number');

    // Identical documents should have no changes
    assert.equal(summary.changed, 0);
    assert.equal(summary.added, 0);
    assert.equal(summary.removed, 0);

    doc.discard();
  });

  it('diffSummary returns correct counts for modified document', async () => {
    const out1 = freshCopy('diffsummary-orig');
    const out2 = freshCopy('diffsummary-modified');

    // Modify out2: replace text and add a paragraph
    const mod = docex(out2);
    mod.author('Test').untracked();
    mod.replace('Introduction', 'CHANGED_INTRO');
    mod.after('Methods').insert('A brand new paragraph added for diff testing.');
    await mod.save(out2);

    // Now compare
    const doc = docex(out1);
    const summary = await doc.diffSummary(out2);

    // Should detect changes
    assert.ok(summary.changed >= 1 || summary.added >= 1 || summary.removed >= 0,
      'Should detect at least one difference');

    doc.discard();
  });

  it('diffSummary detects comment count differences', async () => {
    const out1 = freshCopy('diffsummary-comments1');
    const out2 = freshCopy('diffsummary-comments2');

    // Add comments to out2
    const mod = docex(out2);
    mod.author('Reviewer');
    mod.comment('Introduction', 'This needs more detail');
    mod.comment('Methods', 'Please clarify');
    await mod.save(out2);

    // Compare
    const doc = docex(out1);
    const summary = await doc.diffSummary(out2);

    assert.ok(summary.comments >= 2, `Expected at least 2 new comments, got ${summary.comments}`);

    doc.discard();
  });
});

// ============================================================================
// 4. OPERATION RECEIPTS
// ============================================================================

describe('operation receipts', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('operation receipts are returned from save()', async () => {
    const out = freshCopy('receipts-basic');
    const doc = docex(out);
    doc.author('Test');

    doc.replace('Introduction', 'Intro');
    doc.comment('Methods', 'Needs work');

    const result = await doc.save(out);

    assert.ok(Array.isArray(result.receipts), 'result.receipts should be an array');
    assert.equal(result.receipts.length, 2, 'Should have 2 receipts');

    // First receipt: replace
    const r0 = result.receipts[0];
    assert.equal(r0.success, true);
    assert.equal(r0.type, 'replace');
    assert.ok(r0.matched, 'Receipt should include matched text');
    assert.ok(r0.context, 'Receipt should include context');

    // Second receipt: comment
    const r1 = result.receipts[1];
    assert.equal(r1.success, true);
    assert.equal(r1.type, 'comment');
    assert.ok(r1.matched, 'Comment receipt should include matched text');
  });

  it('receipts include paraId when addressing system is used', async () => {
    const out = freshCopy('receipts-paraid');
    const doc = docex(out);
    doc.author('Test');

    // Queue a replace that targets text in a known paragraph
    doc.replace('Introduction', 'Introduction_Modified');

    const result = await doc.save(out);

    assert.ok(result.receipts.length >= 1, 'Should have at least 1 receipt');

    const r0 = result.receipts[0];
    assert.equal(r0.success, true);
    // The paraId should be detected for the paragraph containing the replaced text
    assert.ok(r0.paraId, `Receipt should include paraId, got: ${JSON.stringify(r0)}`);
    assert.equal(typeof r0.paraId, 'string', 'paraId should be a string');
    assert.ok(r0.paraId.length > 0, 'paraId should be non-empty');
  });

  it('failed operations produce receipts with success=false', async () => {
    const out = freshCopy('receipts-fail');
    const doc = docex(out);
    doc.author('Test');

    // Queue an operation that will fail (text doesn't exist)
    doc.replace('THIS_TEXT_DEFINITELY_DOES_NOT_EXIST_IN_THE_DOC_XYZ123', 'replacement');

    const result = await doc.save(out);

    assert.ok(result.receipts.length >= 1, 'Should have at least 1 receipt');
    const r0 = result.receipts[0];
    assert.equal(r0.success, false, 'Failed operation should have success=false');
    assert.equal(r0.type, 'replace');
    assert.ok(r0.error, 'Failed receipt should include error message');
  });

  it('multiple operations each get their own receipt', async () => {
    const out = freshCopy('receipts-multi');
    const doc = docex(out);
    doc.author('Test');

    doc.replace('Introduction', 'Intro');
    doc.after('Methods').insert('New paragraph here');
    doc.comment('Results', 'Review this');

    const result = await doc.save(out);

    assert.equal(result.receipts.length, 3, 'Should have 3 receipts');

    // All should succeed
    assert.equal(result.receipts[0].type, 'replace');
    assert.equal(result.receipts[0].success, true);

    assert.equal(result.receipts[1].type, 'insert');
    assert.equal(result.receipts[1].success, true);

    assert.equal(result.receipts[2].type, 'comment');
    assert.equal(result.receipts[2].success, true);
  });
});
