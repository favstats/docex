
const { describe, it, before } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
function freshCopy(n) { var o = path.join(OUTPUT_DIR, 'adv-' + n + '.docx'); fs.copyFileSync(FIXTURE, o); return o; }

describe('RangeHandle', () => {
  var docex, DocMap, xmlLib;
  before(() => { docex = require('../src/docex'); DocMap = require('../src/docmap').DocMap; xmlLib = require('../src/xml'); });
  it('text() extracts text across paragraphs', async () => {
    var out = freshCopy('range-text'); var doc = docex(out); doc.author('Test').untracked();
    var map = await doc.map(); var p1 = map.allParagraphs[1]; var p2 = map.allParagraphs[2];
    var rng = doc.range(p1.id, 0, p2.id, 20); var text = rng.text();
    assert.ok(text.includes('This is the first'), 'Should contain first para text');
    assert.ok(text.includes('The second paragrap'), 'Should contain beginning of second para');
    doc.discard();
  });
  it('bold() applies formatting across range', async () => {
    var out = freshCopy('range-bold'); var doc = docex(out); doc.author('Test').untracked();
    var map = await doc.map(); var p1 = map.allParagraphs[1];
    var rng = doc.range(p1.id, 0, p1.id, 10); rng.bold();
    var ws = doc._workspace; var loc = DocMap.locateById(ws.docXml, p1.id);
    assert.ok(loc.xml.includes('<w:b/>'), 'Should contain bold element');
    doc.discard();
  });
  it('delete() with tracked changes wraps in w:del', async () => {
    var out = freshCopy('range-del'); var doc = docex(out); doc.author('Test').tracked();
    var map = await doc.map(); var p1 = map.allParagraphs[1];
    var rng = doc.range(p1.id, 0, p1.id, 20);
    rng.delete({ tracked: true, author: 'Reviewer' });
    var ws = doc._workspace; var loc = DocMap.locateById(ws.docXml, p1.id);
    assert.ok(loc.xml.includes('<w:del'), 'Should contain tracked deletion');
    assert.ok(loc.xml.includes('w:author="Reviewer"'), 'Should attribute to Reviewer');
    doc.discard();
  });
  it('comment() anchors comment to range text', async () => {
    var out = freshCopy('range-comment'); var doc = docex(out); doc.author('Test');
    var map = await doc.map(); var p1 = map.allParagraphs[1];
    doc.range(p1.id, 0, p1.id, 30).comment('Needs citation', { by: 'Reviewer 2' });
    var ws = doc._workspace;
    assert.ok(ws.commentsXml.includes('Needs citation'));
    assert.ok(ws.commentsXml.includes('Reviewer 2'));
    doc.discard();
  });
  it('cut/paste moves content', async () => {
    var out = freshCopy('range-cut'); var doc = docex(out); doc.author('Test').untracked();
    var map = await doc.map(); var p1 = map.allParagraphs[1];
    var origText = p1.text;
    var cutText = doc.range(p1.id, 0, p1.id, 10).cut();
    assert.equal(cutText, origText.slice(0, 10));
    doc.discard();
  });
});

describe('ParagraphHandle formatting', () => {
  var docex;
  before(() => { docex = require('../src/docex'); });
  it('formatting() reads paragraph properties', async () => {
    var out = freshCopy('para-fmt'); var doc = docex(out); doc.author('Test');
    var map = await doc.map(); var handle = doc.id(map.allParagraphs[1].id);
    var fmt = handle.formatting();
    assert.ok('style' in fmt); assert.ok('font' in fmt); assert.ok('size' in fmt);
    assert.ok('keepWithNext' in fmt); assert.ok('pageBreakBefore' in fmt);
    doc.discard();
  });
  it('formattingAt() reads character properties', async () => {
    var out = freshCopy('para-fmtAt'); var doc = docex(out); doc.author('Test').untracked();
    var map = await doc.map(); var handle = doc.id(map.allParagraphs[1].id);
    var charFmt = handle.formattingAt(0);
    assert.ok('bold' in charFmt); assert.ok('italic' in charFmt);
    assert.equal(charFmt.font, 'Times New Roman'); assert.equal(charFmt.size, 12);
    assert.equal(charFmt.bold, false);
    doc.discard();
  });
});

describe('findFormatted', () => {
  var docex, Paragraphs, xmlLib;
  before(() => { docex = require('../src/docex'); Paragraphs = require('../src/paragraphs').Paragraphs; xmlLib = require('../src/xml'); });
  it('findFormatted finds bold text', async () => {
    var out = freshCopy('findFmt'); var doc = docex(out); doc.author('Test').untracked();
    await doc.map();
    doc.id((await doc.map()).allParagraphs[1].id).bold('platform');
    var results = Paragraphs.findFormatted(doc._workspace, { text: 'platform', bold: true });
    assert.ok(results.length >= 1); assert.ok(results[0].text.includes('platform'));
    doc.discard();
  });
  it('replaceFormatted changes text + formatting', async () => {
    var out = freshCopy('replaceFmt'); var doc = docex(out); doc.author('Test').untracked();
    await doc.map();
    doc.id((await doc.map()).allParagraphs[2].id).bold("Meta");
    var count = Paragraphs.replaceFormatted(doc._workspace, { text: 'Meta', bold: true }, { text: 'Meta Platforms', bold: true, italic: true });
    assert.ok(count >= 1);
    assert.ok(xmlLib.extractTextDecoded(doc._workspace.docXml).includes('Meta Platforms'));
    doc.discard();
  });
  it('findByFormatting finds by font/size', async () => {
    var out = freshCopy('findByFmt'); var doc = docex(out); doc.author('Test').untracked();
    await doc.map();
    var results = Paragraphs.findByFormatting(doc._workspace, { font: 'Times New Roman' });
    assert.ok(results.length >= 1);
    doc.discard();
  });
});

describe('Named Checkpoints', () => {
  var docex;
  before(() => { docex = require('../src/docex'); });
  it('checkpoint saves and restores state', async () => {
    var out = freshCopy('ckpt-sr'); var doc = docex(out); doc.author('Test').untracked();
    var map = await doc.map(); var p1 = map.allParagraphs[1];
    await doc.checkpoint('before-edit');
    doc.id(p1.id).replace('first paragraph', 'MODIFIED paragraph');
    assert.ok(doc.id(p1.id).text.includes('MODIFIED'));
    await doc.restoreTo('before-edit');
    assert.ok(doc.id(p1.id).text.includes('first paragraph'));
    assert.ok(!doc.id(p1.id).text.includes('MODIFIED'));
    doc.discard();
  });
  it('listCheckpoints returns saved checkpoints', async () => {
    var out = freshCopy('ckpt-list'); var doc = docex(out); doc.author('Test'); await doc.map();
    await doc.checkpoint('alpha'); await doc.checkpoint('beta');
    var checkpoints = await doc.listCheckpoints();
    assert.equal(checkpoints.length, 2); assert.equal(checkpoints[0].name, 'alpha');
    doc.discard();
  });
  it('restoreTo throws for unknown checkpoint', async () => {
    var out = freshCopy('ckpt-unknown'); var doc = docex(out); doc.author('Test'); await doc.map();
    await assert.rejects(() => doc.restoreTo('nonexistent'), /does not exist/);
    doc.discard();
  });
});

describe('Fields', () => {
  var Fields, Workspace, xmlLib;
  before(() => { Fields = require('../src/fields').Fields; Workspace = require('../src/workspace').Workspace; xmlLib = require('../src/xml'); });
  it('Fields.list() finds field codes', () => {
    var out = freshCopy('fields-list'); var ws = Workspace.open(out);
    var fields = Fields.list(ws); assert.ok(Array.isArray(fields));
    ws.cleanup();
  });
  it('Fields.insert() adds a field and list finds it', () => {
    var out = freshCopy('fields-ins'); var ws = Workspace.open(out);
    var paras = xmlLib.findParagraphs(ws.docXml);
    var paraId = paras[1].xml.match(/w14:paraId="([^"]+)"/)[1];
    Fields.insert(ws, 'PAGE', 'PAGE', { paraId: paraId, result: '1' });
    var fields = Fields.list(ws);
    assert.ok(fields.length >= 1);
    var pf = fields.find(function(f) { return f.type === 'PAGE'; });
    assert.ok(pf); assert.equal(pf.result, '1');
    ws.cleanup();
  });
});

describe('Headers', () => {
  var Headers, Workspace;
  before(() => { Headers = require('../src/headers').Headers; Workspace = require('../src/workspace').Workspace; });
  it('Headers.get/set round-trip', () => {
    var out = freshCopy('hdr-rt'); var ws = Workspace.open(out);
    Headers.set(ws, 'My Title', { type: 'default' });
    var result = Headers.get(ws, 'default');
    assert.ok(result); assert.ok(result.text.includes('My Title'));
    ws.cleanup();
  });
});
