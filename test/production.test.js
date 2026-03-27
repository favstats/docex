/**
 * production.test.js -- Tests for the Production module (v0.4.5)
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/production.test.js
 */
const { describe, it, before } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function freshCopy(n) { const o = path.join(OUTPUT_DIR, n + '.docx'); fs.copyFileSync(FIXTURE, o); return o; }

function readDocxXml(dp, xf) {
  const t = fs.mkdtempSync('/tmp/docex-pt-');
  execFileSync('unzip', ['-o', dp, '-d', t], { stdio: 'pipe' });
  const fp = path.join(t, xf);
  const c = fs.existsSync(fp) ? fs.readFileSync(fp, 'utf8') : '';
  execFileSync('rm', ['-rf', t], { stdio: 'pipe' });
  return c;
}

function fileInDocx(dp, inner) {
  const t = fs.mkdtempSync('/tmp/docex-pt-');
  execFileSync('unzip', ['-o', dp, '-d', t], { stdio: 'pipe' });
  const e = fs.existsSync(path.join(t, inner));
  execFileSync('rm', ['-rf', t], { stdio: 'pipe' });
  return e;
}

describe('production - watermark', () => {
  let P, W;
  before(() => { P = require('../src/production').Production; W = require('../src/workspace').Workspace; });

  it('adds VML shape to header', () => {
    const o = freshCopy('pw1'); const ws = W.open(o); P.watermark(ws, 'DRAFT'); ws.save(o);
    assert.ok(fileInDocx(o, 'word/headerWatermark1.xml'));
    const h = readDocxXml(o, 'word/headerWatermark1.xml');
    assert.ok(h.includes('v:shape')); assert.ok(h.includes('DRAFT'));
  });

  it('uses custom color', () => {
    const o = freshCopy('pw2'); const ws = W.open(o); P.watermark(ws, 'X', { color: 'FF0000' }); ws.save(o);
    assert.ok(readDocxXml(o, 'word/headerWatermark1.xml').includes('#FF0000'));
  });

  it('uses custom size', () => {
    const o = freshCopy('pw3'); const ws = W.open(o); P.watermark(ws, 'X', { size: 100 }); ws.save(o);
    assert.ok(readDocxXml(o, 'word/headerWatermark1.xml').includes('font-size:100pt'));
  });

  it('adds relationship', () => {
    const o = freshCopy('pw4'); const ws = W.open(o); P.watermark(ws, 'X'); ws.save(o);
    assert.ok(readDocxXml(o, 'word/_rels/document.xml.rels').includes('headerWatermark1.xml'));
  });

  it('adds content type', () => {
    const o = freshCopy('pw5'); const ws = W.open(o); P.watermark(ws, 'X'); ws.save(o);
    assert.ok(readDocxXml(o, '[Content_Types].xml').includes('headerWatermark1.xml'));
  });
});

describe('production - stamp', () => {
  let P, W;
  before(() => { P = require('../src/production').Production; W = require('../src/workspace').Workspace; });

  it('adds header stamp', () => {
    const o = freshCopy('ps1'); const ws = W.open(o); P.stamp(ws, 'Confidential'); ws.save(o);
    assert.ok(fileInDocx(o, 'word/headerStamp1.xml'));
    const h = readDocxXml(o, 'word/headerStamp1.xml');
    assert.ok(h.includes('Confidential')); assert.ok(h.includes('w:hdr'));
  });

  it('adds footer stamp', () => {
    const o = freshCopy('ps2'); const ws = W.open(o); P.stamp(ws, 'Footer', { position: 'footer' }); ws.save(o);
    assert.ok(fileInDocx(o, 'word/footerStamp1.xml'));
    const f = readDocxXml(o, 'word/footerStamp1.xml');
    assert.ok(f.includes('Footer')); assert.ok(f.includes('w:ftr'));
  });

  it('uses right alignment', () => {
    const o = freshCopy('ps3'); const ws = W.open(o); P.stamp(ws, 'Right', { alignment: 'right' }); ws.save(o);
    assert.ok(readDocxXml(o, 'word/headerStamp1.xml').includes('w:val="right"'));
  });
});

describe('production - pageCount', () => {
  let P, W;
  before(() => { P = require('../src/production').Production; W = require('../src/workspace').Workspace; });

  it('returns reasonable estimate', () => {
    const ws = W.open(freshCopy('pp1')); const r = P.pageCount(ws);
    assert.ok(typeof r.estimated === 'number'); assert.ok(r.estimated >= 1); assert.strictEqual(r.confidence, 'rough'); ws.cleanup();
  });

  it('includes counts', () => {
    const ws = W.open(freshCopy('pp2')); const r = P.pageCount(ws);
    assert.ok(typeof r.wordCount === 'number'); assert.ok(typeof r.figureCount === 'number');
    assert.ok(typeof r.tableCount === 'number'); assert.ok(r.wordCount > 0); ws.cleanup();
  });

  it('more content means more pages', () => {
    const ws1 = W.open(freshCopy('pp3a')); const base = P.pageCount(ws1); ws1.cleanup();
    const ws2 = W.open(freshCopy('pp3b'));
    let extra = '';
    for (let i = 0; i < 100; i++) extra += '<w:p><w:r><w:t>Extra paragraph with significant text content for testing page count estimation.</w:t></w:r></w:p>';
    ws2.docXml = ws2.docXml.replace('</w:body>', extra + '</w:body>');
    assert.ok(P.pageCount(ws2).estimated >= base.estimated); ws2.cleanup();
  });
});

describe('production - coverPage', () => {
  let P, W;
  before(() => { P = require('../src/production').Production; W = require('../src/workspace').Workspace; });

  it('inserts title', () => {
    const o = freshCopy('pc1'); const ws = W.open(o); P.coverPage(ws, { title: 'My Paper', author: 'Fabio', date: '2026-03-27' }); ws.save(o);
    const d = readDocxXml(o, 'word/document.xml');
    assert.ok(d.includes('My Paper')); assert.ok(d.includes('Fabio')); assert.ok(d.includes('2026-03-27'));
  });

  it('includes page break', () => {
    const o = freshCopy('pc2'); const ws = W.open(o); P.coverPage(ws, { title: 'T' }); ws.save(o);
    assert.ok(readDocxXml(o, 'word/document.xml').includes('w:type="page"'));
  });

  it('includes subtitle', () => {
    const o = freshCopy('pc3'); const ws = W.open(o); P.coverPage(ws, { title: 'T', subtitle: 'Sub' }); ws.save(o);
    assert.ok(readDocxXml(o, 'word/document.xml').includes('Sub'));
  });

  it('includes organization', () => {
    const o = freshCopy('pc4'); const ws = W.open(o); P.coverPage(ws, { title: 'T', organization: 'UvA' }); ws.save(o);
    assert.ok(readDocxXml(o, 'word/document.xml').includes('UvA'));
  });

  it('title near body start', () => {
    const o = freshCopy('pc5'); const ws = W.open(o); P.coverPage(ws, { title: 'MARKER_TITLE' }); ws.save(o);
    const d = readDocxXml(o, 'word/document.xml');
    assert.ok(d.indexOf('MARKER_TITLE') > d.indexOf('<w:body>'));
    assert.ok(d.indexOf('MARKER_TITLE') < d.indexOf('<w:body>') + 2000);
  });

  it('default title', () => {
    const o = freshCopy('pc6'); const ws = W.open(o); P.coverPage(ws, {}); ws.save(o);
    assert.ok(readDocxXml(o, 'word/document.xml').includes('Untitled Document'));
  });
});

describe('production - API', () => {
  let docex;
  before(() => { docex = require('../src/docex'); });

  it('doc.watermark()', async () => {
    const o = freshCopy('pa1'); const d = docex(o); await d.watermark('DRAFT'); const r = await d.save(o);
    assert.ok(r.verified); assert.ok(fileInDocx(o, 'word/headerWatermark1.xml'));
  });

  it('doc.stamp()', async () => {
    const o = freshCopy('pa2'); const d = docex(o); await d.stamp('Conf'); const r = await d.save(o);
    assert.ok(r.verified); assert.ok(fileInDocx(o, 'word/headerStamp1.xml'));
  });

  it('doc.pageCount()', async () => {
    const o = freshCopy('pa3'); const d = docex(o); const r = await d.pageCount();
    assert.ok(typeof r.estimated === 'number'); assert.ok(r.estimated >= 1); d.discard();
  });

  it('doc.coverPage()', async () => {
    const o = freshCopy('pa4'); const d = docex(o); await d.coverPage({ title: 'API Test' }); const r = await d.save(o);
    assert.ok(r.verified); assert.ok(readDocxXml(o, 'word/document.xml').includes('API Test'));
  });
});
