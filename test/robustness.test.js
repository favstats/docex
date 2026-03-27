/**
 * robustness test suite
 *
 * Tests for v0.2 robustness improvements:
 *   1. Better error messages with "Did you mean" closest matches
 *   2. Comment anchoring to exact phrase (run-level splitting)
 *   3. replaceAll replaces every occurrence
 *   4. replaceRegex works with patterns
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/robustness.test.js
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
  const tmp = fs.mkdtempSync('/tmp/docex-robust-');
  execFileSync('unzip', ['-o', docxPath, '-d', tmp], { stdio: 'pipe' });
  const content = fs.readFileSync(path.join(tmp, xmlFile), 'utf8');
  execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
  return content;
}

// ============================================================================
// 1. BETTER ERROR MESSAGES
// ============================================================================

describe('error messages with closest matches', () => {
  let Paragraphs, Workspace, findClosestMatches;

  before(() => {
    Paragraphs = require('../src/paragraphs').Paragraphs;
    findClosestMatches = require('../src/paragraphs').findClosestMatches;
    Workspace = require('../src/workspace').Workspace;
  });

  it('findClosestMatches returns similar paragraphs', () => {
    const paragraphs = [
      { text: 'The quick brown fox jumps over the lazy dog' },
      { text: 'A completely different paragraph about nothing' },
      { text: 'The quick brown cat sleeps on the mat' },
      { text: '' },
    ];
    const matches = findClosestMatches(paragraphs, 'quick brown fox');
    assert.ok(matches.length <= 3, 'returns at most 3 matches');
    assert.ok(matches.length >= 1, 'returns at least 1 match');
    // The fox paragraph should rank highest since it contains the search text
    assert.ok(matches[0].includes('quick brown fox'), 'best match contains search text');
  });

  it('findClosestMatches truncates long paragraphs to 60 chars', () => {
    const longText = 'A'.repeat(200);
    const paragraphs = [{ text: longText }];
    const matches = findClosestMatches(paragraphs, 'AAA');
    assert.ok(matches[0].length <= 60, 'truncated to 60 chars');
    assert.ok(matches[0].endsWith('...'), 'ends with ellipsis');
  });

  it('replace error includes "Did you mean" for text not found', () => {
    const ws = Workspace.open(FIXTURE);
    try {
      Paragraphs.replace(ws, 'xyzzy_nonexistent_text_foobar', 'replacement', { tracked: false });
      assert.fail('should have thrown');
    } catch (err) {
      assert.ok(err.message.includes('Did you mean'), 'error contains "Did you mean"');
      assert.ok(err.message.includes('Text not found'), 'error mentions "Text not found"');
    }
    ws.cleanup();
  });

  it('tracked replace error includes "Did you mean"', () => {
    const ws = Workspace.open(FIXTURE);
    try {
      Paragraphs.replace(ws, 'xyzzy_nonexistent_text_foobar', 'replacement', { tracked: true });
      assert.fail('should have thrown');
    } catch (err) {
      assert.ok(err.message.includes('Did you mean'), 'error contains "Did you mean"');
    }
    ws.cleanup();
  });

  it('insert anchor error includes "Did you mean"', () => {
    const ws = Workspace.open(FIXTURE);
    try {
      Paragraphs.insert(ws, 'xyzzy_nonexistent_anchor', 'after', 'New text', { tracked: false });
      assert.fail('should have thrown');
    } catch (err) {
      assert.ok(err.message.includes('Did you mean'), 'error contains "Did you mean"');
    }
    ws.cleanup();
  });

  it('delete error includes "Did you mean"', () => {
    const ws = Workspace.open(FIXTURE);
    try {
      Paragraphs.remove(ws, 'xyzzy_nonexistent_text_foobar', { tracked: true });
      assert.fail('should have thrown');
    } catch (err) {
      assert.ok(err.message.includes('Did you mean'), 'error contains "Did you mean"');
    }
    ws.cleanup();
  });

  it('untracked delete error includes "Did you mean"', () => {
    const ws = Workspace.open(FIXTURE);
    try {
      Paragraphs.remove(ws, 'xyzzy_nonexistent_text_foobar', { tracked: false });
      assert.fail('should have thrown');
    } catch (err) {
      assert.ok(err.message.includes('Did you mean'), 'error contains "Did you mean"');
    }
    ws.cleanup();
  });
});

describe('comment error messages with closest matches', () => {
  let Comments, Workspace;

  before(() => {
    Comments = require('../src/comments').Comments;
    Workspace = require('../src/workspace').Workspace;
  });

  it('comment add error includes "Did you mean" for anchor not found', () => {
    const out = freshCopy('comment-err-didyoumean');
    const ws = Workspace.open(out);
    try {
      Comments.add(ws, 'xyzzy_nonexistent_anchor_text', 'Test comment', { by: 'Tester' });
      assert.fail('should have thrown');
    } catch (err) {
      assert.ok(err.message.includes('Did you mean'), 'error contains "Did you mean"');
      assert.ok(err.message.includes('could not find anchor'), 'error mentions anchor not found');
    }
    ws.cleanup();
  });
});

// ============================================================================
// 2. COMMENT ANCHORING TO EXACT PHRASE
// ============================================================================

describe('exact phrase comment anchoring', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('commentRangeStart is placed before the run containing anchor text', async () => {
    const out = freshCopy('comment-exact-anchor');
    const doc = docex(out);
    // "platform governance" appears in the middle of a paragraph
    doc.at('platform governance').comment('Test exact anchoring', { by: 'Tester' });
    await doc.save();

    const docXml = readDocxXml(out, 'word/document.xml');

    // commentRangeStart should appear BEFORE a <w:r> that contains "platform governance"
    assert.ok(docXml.includes('commentRangeStart'), 'has commentRangeStart');
    assert.ok(docXml.includes('commentRangeEnd'), 'has commentRangeEnd');

    // The rangeStart should be near the anchor text, not at paragraph start
    const rangeStartPos = docXml.indexOf('commentRangeStart');
    const anchorTextPos = docXml.indexOf('platform governance', rangeStartPos);

    // The rangeStart should be close to (and before) the anchor text
    assert.ok(rangeStartPos < anchorTextPos, 'rangeStart before anchor text');
    // Check that rangeStart is within the same paragraph region as the anchor text,
    // not at the very beginning. The distance should be small (a run tag + rPr + t tag).
    const distance = anchorTextPos - rangeStartPos;
    assert.ok(distance < 500, 'rangeStart is close to anchor text (within 500 chars), actual: ' + distance);
  });

  it('comment anchors only the target phrase, not entire paragraph', async () => {
    const out = freshCopy('comment-exact-narrow');
    const doc = docex(out);
    doc.at('268,635').comment('Verify this number', { by: 'Dr. Check' });
    await doc.save();

    const docXml = readDocxXml(out, 'word/document.xml');

    // Find the commentRangeStart and commentRangeEnd
    const startMatch = docXml.match(/<w:commentRangeStart[^/]*\/>/);
    const endMatch = docXml.match(/<w:commentRangeEnd[^/]*\/>/);
    assert.ok(startMatch, 'has commentRangeStart');
    assert.ok(endMatch, 'has commentRangeEnd');

    // Extract the text between commentRangeStart and commentRangeEnd
    const startIdx = docXml.indexOf(startMatch[0]) + startMatch[0].length;
    const endIdx = docXml.indexOf(endMatch[0]);
    const between = docXml.slice(startIdx, endIdx);

    // Extract w:t text from between the markers
    const textParts = [];
    const tRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let m;
    while ((m = tRe.exec(between)) !== null) {
      textParts.push(m[1]);
    }
    const highlightedText = textParts.join('');

    // The highlighted text should contain the anchor phrase but not the entire paragraph
    assert.ok(highlightedText.includes('268,635'), 'highlighted text contains anchor: ' + highlightedText);
    // Should NOT include text far from the anchor
    assert.ok(!highlightedText.includes('Introduction'), 'highlighted text does not include distant text');
  });

  it('split runs are valid XML after exact anchoring', async () => {
    const out = freshCopy('comment-exact-valid');
    const doc = docex(out);
    doc.at('electoral transparency').comment('Define this', { by: 'Editor' });
    await doc.save();

    // Verify the output is a valid zip (if runs were malformed, it would fail)
    const result = execFileSync('unzip', ['-t', out], { encoding: 'utf8' });
    assert.ok(result.includes('No errors'), 'output is valid zip');

    // Verify the document still has all paragraphs
    const docXml = readDocxXml(out, 'word/document.xml');
    assert.ok(docXml.includes('Introduction'), 'document content preserved');
    assert.ok(docXml.includes('electoral transparency'), 'anchor text preserved');
  });
});

// ============================================================================
// 3. REPLACEALL
// ============================================================================

describe('replaceAll', () => {
  let Paragraphs, Workspace;

  before(() => {
    Paragraphs = require('../src/paragraphs').Paragraphs;
    Workspace = require('../src/workspace').Workspace;
  });

  it('replaces all occurrences of text', () => {
    const out = freshCopy('replaceall-basic');
    const ws = Workspace.open(out);

    // "the" appears many times in the test manuscript
    const count = Paragraphs.replaceAll(ws, 'the', 'THE', { tracked: false });
    assert.ok(count > 1, 'replaced more than one occurrence, got ' + count);

    // Verify replacement text is present
    const text = Paragraphs.fullText(ws);
    assert.ok(text.includes('THE'), 'replacement text present');

    ws.cleanup();
  });

  it('returns 0 when text not found', () => {
    const ws = Workspace.open(FIXTURE);
    const count = Paragraphs.replaceAll(ws, 'xyzzy_nonexistent', 'replacement', { tracked: false });
    assert.equal(count, 0, 'returns 0 for no matches');
    ws.cleanup();
  });

  it('works with tracked changes', () => {
    const out = freshCopy('replaceall-tracked');
    const ws = Workspace.open(out);

    // The fixture has "platform" appearing multiple times
    const count = Paragraphs.replaceAll(ws, 'platform', 'PLATFORM', {
      tracked: true,
      author: 'Tester',
    });
    assert.ok(count >= 2, 'replaced multiple occurrences with tracking, got ' + count);

    // Verify tracked changes were created
    const docXml = ws.docXml;
    assert.ok(docXml.includes('<w:del'), 'tracked deletions present');
    assert.ok(docXml.includes('<w:ins'), 'tracked insertions present');

    ws.cleanup();
  });
});

// ============================================================================
// 4. REPLACEREGEX
// ============================================================================

describe('replaceRegex', () => {
  let Paragraphs, Workspace;

  before(() => {
    Paragraphs = require('../src/paragraphs').Paragraphs;
    Workspace = require('../src/workspace').Workspace;
  });

  it('replaces text matching a regex pattern', () => {
    const out = freshCopy('replaceregex-basic');
    const ws = Workspace.open(out);

    // Replace numbers like "268,635" with "NUM"
    const count = Paragraphs.replaceRegex(ws, /\d{1,3},\d{3}/, 'NUM', { tracked: false });
    assert.ok(count >= 1, 'replaced at least one match, got ' + count);

    const text = Paragraphs.fullText(ws);
    assert.ok(text.includes('NUM'), 'regex replacement text present');

    ws.cleanup();
  });

  it('returns 0 for no regex matches', () => {
    const ws = Workspace.open(FIXTURE);
    const count = Paragraphs.replaceRegex(ws, /ZZZZNOTFOUND\d{10}/, 'replacement', { tracked: false });
    assert.equal(count, 0, 'returns 0 for no matches');
    ws.cleanup();
  });

  it('replaces all matches when pattern has global flag', () => {
    const out = freshCopy('replaceregex-global');
    const ws = Workspace.open(out);

    // Count how many comma-separated numbers exist before
    const textBefore = Paragraphs.fullText(ws);
    const matchesBefore = textBefore.match(/\d{1,3},\d{3}/g) || [];

    const count = Paragraphs.replaceRegex(ws, /\d{1,3},\d{3}/g, 'NUM', { tracked: false });
    assert.ok(count >= 2, 'replaced multiple regex matches, got ' + count);
    assert.equal(count, matchesBefore.length, 'replaced all regex matches');

    ws.cleanup();
  });

  it('works with tracked changes', () => {
    const out = freshCopy('replaceregex-tracked');
    const ws = Workspace.open(out);

    const count = Paragraphs.replaceRegex(ws, /\d{1,3},\d{3}/, 'NUM', {
      tracked: true,
      author: 'Regex Tester',
    });
    assert.ok(count >= 1, 'replaced at least one regex match with tracking');

    const docXml = ws.docXml;
    assert.ok(docXml.includes('<w:del'), 'tracked deletion present');
    assert.ok(docXml.includes('<w:ins'), 'tracked insertion present');

    ws.cleanup();
  });
});
