/**
 * transactions.test.js -- Tests for transactions, conditionals, and verification chaining
 *
 * Features tested:
 *   1. Transaction commit/abort/rollback
 *   2. Conditional if/unless/ifContains/ifEmpty
 *   3. Verification chaining with verify()
 *   4. Combined: transactions + conditionals + verify
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/transactions.test.js
 */

const { describe, it, before } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function freshCopy(testName) {
  const out = path.join(OUTPUT_DIR, testName + '.docx');
  fs.copyFileSync(FIXTURE, out);
  return out;
}

// Load docex first, then extensions to monkey-patch
let docex, extensions, VerificationError;

// ============================================================================
// 1. TRANSACTIONS
// ============================================================================

describe('transactions', () => {
  before(() => {
    docex = require('../src/docex');
    extensions = require('../src/extensions');
    VerificationError = extensions.VerificationError;
  });

  it('transaction commit applies all operations atomically', async () => {
    const out = freshCopy('tx-commit');
    const doc = docex(out);
    doc.author('Test').untracked();
    await doc.map();

    const tx = doc.transaction();
    tx.replace('Introduction', 'INTRO_NEW', { tracked: false });
    tx.replace('political advertising', 'political ads', { tracked: false });

    const result = await tx.commit({ backup: false });
    assert.ok(result.verified !== false, 'Save should verify');
    assert.equal(result.operations, 2, 'Should report 2 operations');

    const doc2 = docex(out);
    const text = await doc2.text();
    assert.ok(text.includes('INTRO_NEW'), 'First replacement applied');
    assert.ok(text.includes('political ads'), 'Second replacement applied');
    doc2.discard();
  });

  it('transaction abort discards all operations', async () => {
    const out = freshCopy('tx-abort');
    const doc = docex(out);
    doc.author('Test').untracked();
    await doc.text(); // ensure workspace

    const tx = doc.transaction();
    tx.replace('Introduction', 'SHOULD_NOT_APPEAR', { tracked: false });
    await tx.abort();

    const ws = await doc._ensureWorkspace();
    assert.ok(!ws.docXml.includes('SHOULD_NOT_APPEAR'), 'Aborted changes should not appear');
    assert.ok(ws.docXml.includes('Introduction'), 'Original text preserved');
    ws.cleanup();
  });

  it('transaction auto-rollback on failure', async () => {
    const out = freshCopy('tx-rollback');
    const doc = docex(out);
    doc.author('Test').untracked();
    await doc.map();

    const tx = doc.transaction();
    tx.replace('Introduction', 'CHANGED', { tracked: false });
    tx.replace('NONEXISTENT_TEXT_XYZZY_12345', 'anything', { tracked: false });

    try {
      await tx.commit({ backup: false });
      assert.fail('Should have thrown');
    } catch (err) {
      assert.ok(err.message.includes('NONEXISTENT_TEXT_XYZZY_12345') || err.message.includes('not found'),
        'Error should mention the failed text');
    }

    const ws = await doc._ensureWorkspace();
    assert.ok(!ws.docXml.includes('CHANGED'), 'First op rolled back');
    assert.ok(ws.docXml.includes('Introduction'), 'Original restored');
    ws.cleanup();
  });

  it('transaction preview shows pending operations', async () => {
    const out = freshCopy('tx-preview');
    const doc = docex(out);
    doc.author('Test');

    const tx = doc.transaction();
    tx.replace('old', 'new');
    tx.comment('anchor', 'note');

    const preview = tx.preview();
    assert.ok(preview.includes('2 pending operations'), 'Should show count');
    assert.ok(preview.includes('replace'), 'Should show replace');
    assert.ok(preview.includes('comment'), 'Should show comment');
    doc.discard();
  });

  it('multiple transactions do not interfere', async () => {
    const out1 = freshCopy('tx-multi-1');
    const out2 = freshCopy('tx-multi-2');

    const doc1 = docex(out1);
    doc1.author('Test').untracked();
    await doc1.map();

    const doc2 = docex(out2);
    doc2.author('Test').untracked();
    await doc2.map();

    const tx1 = doc1.transaction();
    tx1.replace('Introduction', 'INTRO_A', { tracked: false });

    const tx2 = doc2.transaction();
    tx2.replace('Introduction', 'INTRO_B', { tracked: false });

    await tx1.commit({ backup: false });
    await tx2.commit({ backup: false });

    const text1 = await docex(out1).text();
    const text2 = await docex(out2).text();
    assert.ok(text1.includes('INTRO_A'), 'tx1 has INTRO_A');
    assert.ok(!text1.includes('INTRO_B'), 'tx1 not INTRO_B');
    assert.ok(text2.includes('INTRO_B'), 'tx2 has INTRO_B');
    assert.ok(!text2.includes('INTRO_A'), 'tx2 not INTRO_A');
  });
});

// ============================================================================
// 2. CONDITIONAL OPERATIONS
// ============================================================================

describe('conditional operations', () => {
  before(() => {
    docex = require('../src/docex');
    require('../src/extensions');
  });

  it('if: executes when condition true', async () => {
    const out = freshCopy('cond-if-true');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));
    assert.ok(para, 'Should find Introduction paragraph');

    let executed = false;
    doc.id(para.id).if(h => h.text.includes('Introduction'), h => { executed = true; });
    assert.ok(executed, 'thenFn should execute when condition is true');
    doc.discard();
  });

  it('if: skips when condition false', async () => {
    const out = freshCopy('cond-if-false');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    let executed = false;
    doc.id(para.id).if(h => h.text.includes('NONEXISTENT_XYZZY'), h => { executed = true; });
    assert.ok(!executed, 'thenFn should NOT execute when condition is false');
    doc.discard();
  });

  it('ifContains: works with text check', async () => {
    const out = freshCopy('cond-ifcontains');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    let executed = false;
    doc.id(para.id).ifContains('Introduction', h => { executed = true; });
    assert.ok(executed, 'ifContains should execute when text present');

    let notExecuted = true;
    doc.id(para.id).ifContains('NONEXISTENT_TEXT', h => { notExecuted = false; });
    assert.ok(notExecuted, 'ifContains should NOT execute when text absent');
    doc.discard();
  });

  it('unless: inverse logic', async () => {
    const out = freshCopy('cond-unless');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    let executedOnAbsent = false;
    doc.id(para.id).unless(h => h.text.includes('NONEXISTENT_XYZZY'), h => { executedOnAbsent = true; });
    assert.ok(executedOnAbsent, 'unless should execute when condition false');

    let executedOnPresent = false;
    doc.id(para.id).unless(h => h.text.includes('Introduction'), h => { executedOnPresent = true; });
    assert.ok(!executedOnPresent, 'unless should NOT execute when condition true');
    doc.discard();
  });
});

// ============================================================================
// 3. VERIFICATION CHAINING
// ============================================================================

describe('verification chaining', () => {
  before(() => {
    docex = require('../src/docex');
    extensions = require('../src/extensions');
    VerificationError = extensions.VerificationError;
  });

  it('verify passes when check succeeds', async () => {
    const out = freshCopy('verify-pass');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    const handle = doc.id(para.id).verify(h => h.text.includes('Introduction'));
    assert.ok(handle, 'verify should return handle for chaining');
    doc.discard();
  });

  it('verify throws when check fails', async () => {
    const out = freshCopy('verify-fail');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    assert.throws(
      () => doc.id(para.id).verify(h => h.text.includes('NONEXISTENT')),
      (err) => {
        assert.equal(err.name, 'VerificationError');
        assert.ok(err.paraId === para.id, 'Should include paraId');
        assert.ok(typeof err.currentText === 'string', 'Should include currentText');
        return true;
      }
    );
    doc.discard();
  });

  it('verify stops subsequent chained operations on failure', async () => {
    const out = freshCopy('verify-chain-stop');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    let commentAdded = false;
    try {
      doc.id(para.id)
        .verify(h => h.text.includes('NONEXISTENT'))
        .comment('This should not be added');
      commentAdded = true;
    } catch (err) {
      assert.equal(err.name, 'VerificationError');
    }
    assert.ok(!commentAdded, 'Comment should not be added after verify fails');
    doc.discard();
  });
});

// ============================================================================
// 4. TRANSACTION + CONDITIONAL COMBINED
// ============================================================================

describe('transaction + conditional combined', () => {
  before(() => {
    docex = require('../src/docex');
    require('../src/extensions');
  });

  it('conditional operations work within transactions', async () => {
    const out = freshCopy('tx-cond');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    let shouldReplace = false;
    doc.id(para.id).ifContains('Introduction', h => { shouldReplace = true; });
    assert.ok(shouldReplace);

    const tx = doc.transaction();
    if (shouldReplace) tx.replace('Introduction', 'INTRO_CONDITIONAL', { tracked: false });
    await tx.commit({ backup: false });

    const text = await docex(out).text();
    assert.ok(text.includes('INTRO_CONDITIONAL'), 'Conditional transaction applied');
  });
});

// ============================================================================
// 5. TRANSACTION + VERIFY COMBINED
// ============================================================================

describe('transaction + verify combined', () => {
  before(() => {
    docex = require('../src/docex');
    extensions = require('../src/extensions');
    VerificationError = extensions.VerificationError;
  });

  it('verify before transaction commit prevents bad operations', async () => {
    const out = freshCopy('tx-verify');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    doc.id(para.id).verify(h => h.text.includes('Introduction'));

    const tx = doc.transaction();
    tx.replace('Introduction', 'VERIFIED_REPLACE', { tracked: false });
    await tx.commit({ backup: false });

    const text = await docex(out).text();
    assert.ok(text.includes('VERIFIED_REPLACE'), 'Verified replace applied');
  });

  it('verify failure prevents transaction from executing', async () => {
    const out = freshCopy('tx-verify-fail');
    const doc = docex(out);
    doc.author('Test').untracked();
    const map = await doc.map();
    const para = map.allParagraphs.find(p => p.text.includes('Introduction'));

    let txCreated = false;
    try {
      doc.id(para.id).verify(h => h.text.includes('DOES_NOT_EXIST'));
      txCreated = true;
    } catch (err) {
      assert.equal(err.name, 'VerificationError');
    }
    assert.ok(!txCreated, 'Transaction should not be created after verify fails');
    doc.discard();
  });
});
