/**
 * fuzz.test.js -- Property-based fuzz testing for docex
 *
 * Generates random operations and verifies invariants survive.
 * Tests document integrity under random sequences of:
 *   - replace, insert, comment, bold, footnote
 *   - replaceAll cycles
 *   - alternating accept/reject of tracked changes
 *
 * Uses the test fixture. Verifies: valid zip, paragraph count >= original,
 * file size > 0 after each mutation sequence.
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/fuzz.test.js
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
  const tmp = fs.mkdtempSync('/tmp/docex-fuzz-');
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

/** Helper: verify a .docx file is valid */
function verifyDocx(docxPath) {
  // Must exist and be non-empty
  assert.ok(fs.existsSync(docxPath), 'output file exists');
  const stat = fs.statSync(docxPath);
  assert.ok(stat.size > 0, 'file size > 0');

  // Must be a valid zip
  const zipResult = execFileSync('unzip', ['-t', docxPath], {
    encoding: 'utf8',
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  assert.ok(zipResult.includes('No errors'), 'output is valid zip');

  // Must have document.xml with paragraphs
  const docXml = readDocxXml(docxPath, 'word/document.xml');
  const paraCount = countMatches(docXml, /<w:p[\s>]/g);
  assert.ok(paraCount > 0, 'document has paragraphs');

  return { fileSize: stat.size, paraCount };
}

/** Simple seeded PRNG for reproducibility */
function seededRandom(seed) {
  let s = seed;
  return function () {
    s = (s * 1664525 + 1013904223) & 0xffffffff;
    return (s >>> 0) / 0xffffffff;
  };
}

/** Pick a random element from an array */
function pick(arr, rng) {
  return arr[Math.floor(rng() * arr.length)];
}

/** Generate a random short string */
function randomString(rng, maxLen) {
  maxLen = maxLen || 12;
  const len = Math.floor(rng() * maxLen) + 1;
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789 ';
  let out = '';
  for (let i = 0; i < len; i++) {
    out += chars[Math.floor(rng() * chars.length)];
  }
  return out;
}

// ============================================================================
// 1. RANDOM OPERATIONS FUZZ
// ============================================================================

describe('fuzz testing', () => {
  let docex, Formatting, Workspace, Footnotes;

  before(() => {
    docex = require('../src/docex');
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
    Footnotes = require('../src/footnotes').Footnotes;
  });

  it('survives 50 random operations', async () => {
    const out = freshCopy('fuzz-50-random');
    const rng = seededRandom(42);

    // Get text anchors from the fixture to use as targets
    const doc = docex(out);
    const paras = await doc.paragraphs();
    doc.discard();

    // Collect non-empty paragraph texts for use as anchors
    const anchors = paras
      .filter(p => p.text.length > 10)
      .map(p => p.text.slice(0, 30));

    assert.ok(anchors.length > 5, 'need at least 5 anchors');

    // Get the original paragraph count
    const ws0 = Workspace.open(out);
    const origParaCount = countMatches(ws0.docXml, /<w:p[\s>]/g);
    ws0.cleanup();

    const operations = ['replace', 'insert', 'comment', 'bold', 'footnote'];

    for (let i = 0; i < 50; i++) {
      const op = pick(operations, rng);
      const anchor = pick(anchors, rng);
      const replacement = randomString(rng, 8);

      try {
        switch (op) {
          case 'replace': {
            const doc2 = docex(out);
            doc2.author('Fuzz Tester');
            // Take a small portion of the anchor as old text
            const oldText = anchor.slice(0, Math.min(8, anchor.length));
            doc2.replace(oldText, replacement);
            await doc2.save();
            break;
          }
          case 'insert': {
            const doc2 = docex(out);
            doc2.author('Fuzz Tester');
            doc2.after(anchor).insert('Fuzz inserted: ' + replacement);
            await doc2.save();
            break;
          }
          case 'comment': {
            const doc2 = docex(out);
            doc2.author('Fuzz Tester');
            doc2.at(anchor).comment('Fuzz comment: ' + replacement, { by: 'Fuzz Bot' });
            await doc2.save();
            break;
          }
          case 'bold': {
            const ws = Workspace.open(out);
            try {
              Formatting.bold(ws, anchor.slice(0, 10));
              ws.save(out);
            } catch (e) {
              // Text may not be found after previous mutations -- that is OK
              ws.cleanup();
            }
            break;
          }
          case 'footnote': {
            const ws = Workspace.open(out);
            try {
              Footnotes.add(ws, anchor.slice(0, 10), 'Fuzz fn: ' + replacement);
              ws.save(out);
            } catch (e) {
              // Anchor may not exist after mutations -- that is OK
              ws.cleanup();
            }
            break;
          }
        }
      } catch (e) {
        // Operation may fail if text was already mutated away -- that is expected
        // in fuzz testing. Just continue.
      }
    }

    // Final verification
    const result = verifyDocx(out);
    assert.ok(result.paraCount >= origParaCount,
      `paragraph count should not decrease: got ${result.paraCount}, expected >= ${origParaCount}`);
  });

  it('survives 10 random replace-all cycles', async () => {
    const out = freshCopy('fuzz-10-replace-all');
    const rng = seededRandom(137);

    // Get original paragraph count
    const ws0 = Workspace.open(out);
    const origParaCount = countMatches(ws0.docXml, /<w:p[\s>]/g);
    ws0.cleanup();

    // Short strings to find and replace
    const targets = ['the', 'and', 'of', 'to', 'in', 'for', 'a', 'is', 'that', 'on'];

    for (let i = 0; i < 10; i++) {
      const target = pick(targets, rng);
      const replacement = target + randomString(rng, 3).trim();

      const doc2 = docex(out);
      doc2.author('Fuzz Tester');
      doc2.replace(target, replacement);
      const result = await doc2.save();

      // Verify after each cycle
      assert.ok(result.verified, `cycle ${i + 1}: document should be verified`);
      assert.ok(result.fileSize > 0, `cycle ${i + 1}: file size > 0`);
      assert.ok(result.paragraphCount >= origParaCount,
        `cycle ${i + 1}: paragraph count ${result.paragraphCount} >= ${origParaCount}`);
    }

    // Final zip validity check
    verifyDocx(out);
  });

  it('survives alternating accept/reject cycles', async () => {
    const out = freshCopy('fuzz-accept-reject');

    const Revisions = require('../src/revisions').Revisions;

    // Step 1: Add 5 tracked changes
    const doc2 = docex(out);
    doc2.author('Fuzz Tester');
    doc2.replace('268,635', '300,000');
    doc2.replace('1,329', '1,500');
    doc2.replace('platform governance', 'platform regulation');
    doc2.after('Introduction').insert('Fuzz-inserted paragraph for testing.');
    doc2.after('Discussion').insert('Another fuzz-inserted paragraph.');
    await doc2.save();

    // Step 2: Open workspace and list revisions
    const ws = Workspace.open(out);
    const revisions = Revisions.list(ws);
    assert.ok(revisions.length >= 4, `expected at least 4 revisions, got ${revisions.length}`);

    // Step 3: Accept 2 revisions
    const toAccept = revisions.slice(0, 2);
    for (const rev of toAccept) {
      Revisions.accept(ws, rev.id);
    }

    // Step 4: Reject 2 revisions
    const remaining = Revisions.list(ws);
    const toReject = remaining.slice(0, Math.min(2, remaining.length));
    for (const rev of toReject) {
      Revisions.reject(ws, rev.id);
    }

    // Step 5: Save and verify
    const saveResult = ws.save(out);
    assert.ok(saveResult.verified, 'document verified after accept/reject');
    assert.ok(saveResult.fileSize > 0, 'file size > 0');

    // Final zip validity
    verifyDocx(out);
  });
});
