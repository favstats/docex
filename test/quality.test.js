/**
 * quality.test.js -- Tests for the Quality module (v0.4.4)
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/quality.test.js
 */
const { describe, it, before } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
const STATS_FILE = path.join(__dirname, 'output', 'test-stats.json');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function freshCopy(n) {
  const out = path.join(OUTPUT_DIR, n + '.docx');
  fs.copyFileSync(FIXTURE, out);
  return out;
}

describe('quality - lint', () => {
  let Q, W;
  before(() => { Q = require('../src/quality').Quality; W = require('../src/workspace').Workspace; });

  it('returns an array', () => { const ws = W.open(freshCopy('ql1')); assert.ok(Array.isArray(Q.lint(ws))); ws.cleanup(); });

  it('catches repeated words', () => {
    const ws = W.open(freshCopy('ql2'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>The the quick fox.</w:t></w:r></w:p></w:body>');
    const issues = Q.lint(ws).filter(i => i.type === 'repeated-word');
    assert.ok(issues.length > 0); assert.ok(issues.some(r => r.message.includes('the the'))); ws.cleanup();
  });

  it('catches Figure 5 when few figures exist', () => {
    const ws = W.open(freshCopy('ql3'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>See Figure 5 here.</w:t></w:r></w:p></w:body>');
    const issues = Q.lint(ws).filter(i => i.type === 'invalid-figure-ref');
    assert.ok(issues.length > 0); assert.ok(issues[0].severity === 'error'); ws.cleanup();
  });

  it('catches unclosed parentheses', () => {
    const ws = W.open(freshCopy('ql4'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>Unclosed paren (here.</w:t></w:r></w:p></w:body>');
    assert.ok(Q.lint(ws).filter(i => i.type === 'unclosed-paren').length > 0); ws.cleanup();
  });

  it('catches mismatched quotes', () => {
    const ws = W.open(freshCopy('ql5'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>He said "hello.</w:t></w:r></w:p></w:body>');
    assert.ok(Q.lint(ws).filter(i => i.type === 'mismatched-quotes').length > 0); ws.cleanup();
  });

  it('catches invalid table refs', () => {
    const ws = W.open(freshCopy('ql6'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>See Table 99.</w:t></w:r></w:p></w:body>');
    assert.ok(Q.lint(ws).filter(i => i.type === 'invalid-table-ref').length > 0); ws.cleanup();
  });

  it('includes paraId', () => {
    const ws = W.open(freshCopy('ql7'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p w14:paraId="DEAD01"><w:r><w:t>The the fox.</w:t></w:r></w:p></w:body>');
    assert.ok(Q.lint(ws).filter(i => i.type === 'repeated-word').some(r => r.paraId === 'DEAD01')); ws.cleanup();
  });
});

describe('quality - passiveVoice', () => {
  let Q, W;
  before(() => { Q = require('../src/quality').Quality; W = require('../src/workspace').Workspace; });

  it('detects were collected', () => {
    const ws = W.open(freshCopy('qp1'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>The samples were collected from sites.</w:t></w:r></w:p></w:body>');
    const r = Q.passiveVoice(ws);
    assert.ok(r.length > 0, 'should detect passive voice');
    assert.ok(r.some(x => x.suggestion.includes('collected'))); ws.cleanup();
  });

  it('detects is discussed', () => {
    const ws = W.open(freshCopy('qp2'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>This topic is discussed in detail.</w:t></w:r></w:p></w:body>');
    assert.ok(Q.passiveVoice(ws).length > 0); ws.cleanup();
  });

  it('returns sentence and paraId', () => {
    const ws = W.open(freshCopy('qp3'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p w14:paraId="PAS01"><w:r><w:t>Data has been collected.</w:t></w:r></w:p></w:body>');
    const m = Q.passiveVoice(ws).find(r => r.paraId === 'PAS01');
    assert.ok(m); assert.ok(typeof m.suggestion === 'string'); ws.cleanup();
  });
});

describe('quality - sentenceLength', () => {
  let Q, W;
  before(() => { Q = require('../src/quality').Quality; W = require('../src/workspace').Workspace; });

  it('flags sentences over 40 words', () => {
    const ws = W.open(freshCopy('qs1'));
    const words = Array(45).fill('word').join(' ') + '.';
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>' + words + '</w:t></w:r></w:p></w:body>');
    assert.ok(Q.sentenceLength(ws).some(r => r.wordCount >= 45)); ws.cleanup();
  });

  it('does not flag short sentences', () => {
    const ws = W.open(freshCopy('qs2'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>Short sentence.</w:t></w:r></w:p></w:body>');
    assert.strictEqual(Q.sentenceLength(ws).filter(r => r.sentence === 'Short sentence.').length, 0); ws.cleanup();
  });

  it('respects custom max', () => {
    const ws = W.open(freshCopy('qs3'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>One two three four five six seven eight nine ten eleven.</w:t></w:r></w:p></w:body>');
    assert.ok(Q.sentenceLength(ws, { max: 5 }).some(r => r.wordCount >= 10)); ws.cleanup();
  });
});

describe('quality - readability', () => {
  let Q, W;
  before(() => { Q = require('../src/quality').Quality; W = require('../src/workspace').Workspace; });

  it('returns numeric scores', () => {
    const r = Q.readability(W.open(freshCopy('qr1')));
    assert.ok(typeof r.fleschKincaid === 'number'); assert.ok(typeof r.readingEase === 'number');
    assert.ok(typeof r.avgSentenceLength === 'number'); assert.ok(typeof r.avgSyllables === 'number');
  });

  it('readingEase in range', () => { const r = Q.readability(W.open(freshCopy('qr2'))); assert.ok(r.readingEase > -100 && r.readingEase < 150); });
  it('fleschKincaid in range', () => { const r = Q.readability(W.open(freshCopy('qr3'))); assert.ok(r.fleschKincaid > 0 && r.fleschKincaid < 30); });
  it('avgSentenceLength positive', () => { assert.ok(Q.readability(W.open(freshCopy('qr4'))).avgSentenceLength > 0); });
  it('avgSyllables 1-3', () => { const r = Q.readability(W.open(freshCopy('qr5'))); assert.ok(r.avgSyllables >= 1 && r.avgSyllables <= 3); });
});

describe('quality - _countSyllables', () => {
  let Q;
  before(() => { Q = require('../src/quality').Quality; });
  it('counts simple words', () => { assert.strictEqual(Q._countSyllables('cat'), 1); assert.strictEqual(Q._countSyllables('water'), 2); assert.strictEqual(Q._countSyllables('beautiful'), 3); });
  it('handles empty', () => { assert.strictEqual(Q._countSyllables(''), 0); assert.strictEqual(Q._countSyllables(null), 0); });
  it('at least 1 for real words', () => { assert.ok(Q._countSyllables('the') >= 1); });
});

describe('quality - checkNumbers', () => {
  let Q, W;
  before(() => { Q = require('../src/quality').Quality; W = require('../src/workspace').Workspace;
    fs.writeFileSync(STATS_FILE, JSON.stringify({ total_ads: '268,635', missing: '99,999', sample: '500', txt: 'hello' }));
  });
  it('returns matches and mismatches', () => { const r = Q.checkNumbers(W.open(freshCopy('qn1')), STATS_FILE); assert.ok(Array.isArray(r.matches)); assert.ok(Array.isArray(r.mismatches)); });
  it('finds present numbers', () => {
    const ws = W.open(freshCopy('qn2'));
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>We surveyed 500 people.</w:t></w:r></w:p></w:body>');
    assert.ok(Q.checkNumbers(ws, STATS_FILE).matches.find(m => m.source === 'sample')); ws.cleanup();
  });
  it('reports missing numbers', () => { assert.ok(Q.checkNumbers(W.open(freshCopy('qn3')), STATS_FILE).mismatches.find(m => m.expected === '99,999')); });
});

describe('quality - API', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });
  it('doc.lint()', async () => { const d = docex(freshCopy('qa1')); assert.ok(Array.isArray(await d.lint())); d.discard(); });
  it('doc.passiveVoice()', async () => { const d = docex(freshCopy('qa2')); assert.ok(Array.isArray(await d.passiveVoice())); d.discard(); });
  it('doc.sentenceLength()', async () => { const d = docex(freshCopy('qa3')); assert.ok(Array.isArray(await d.sentenceLength())); d.discard(); });
  it('doc.readability()', async () => { const d = docex(freshCopy('qa4')); const r = await d.readability(); assert.ok(typeof r.fleschKincaid === 'number'); d.discard(); });
  it('doc.checkNumbers()', async () => { const d = docex(freshCopy('qa5')); const r = await d.checkNumbers(STATS_FILE); assert.ok(r.matches); d.discard(); });
});
