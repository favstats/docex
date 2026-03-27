const { describe, it, before } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
function freshCopy(n) { var o = path.join(OUTPUT_DIR, 'surgery-' + n + '.docx'); fs.copyFileSync(FIXTURE, o); return o; }

describe('ParagraphHandle.getXml()', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('returns valid OOXML string', async () => {
    var out = freshCopy('xml-get'); var doc = docex(out); var map = await doc.map();
    var fb = map.allParagraphs.find(function(p) { return p.type === 'body'; });
    assert.ok(fb); var h = doc.id(fb.id); var x = h.getXml();
    assert.ok(x.startsWith('<w:p')); assert.ok(x.endsWith('</w:p>'));
    assert.ok(x.includes('w14:paraId="' + fb.id + '"')); doc.discard();
  });
});

describe('ParagraphHandle.setXml()', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('replaces paragraph XML', async () => {
    var out = freshCopy('xml-set'); var doc = docex(out); var map = await doc.map();
    var fb = map.allParagraphs.find(function(p) { return p.type === 'body'; }); assert.ok(fb);
    var h = doc.id(fb.id);
    var nx = '<w:p w14:paraId="' + fb.id + '" w14:textId="AABBCCDD"><w:r><w:t xml:space="preserve">Replaced content.</w:t></w:r></w:p>';
    h.setXml(nx); assert.equal(h.text, 'Replaced content.'); doc.discard();
  });
});

describe('ParagraphHandle.runs()', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('lists runs with formatting', async () => {
    var out = freshCopy('runs-list'); var doc = docex(out); var map = await doc.map();
    var fb = map.allParagraphs.find(function(p) { return p.type === 'body' && p.text.length > 10; }); assert.ok(fb);
    var runs = doc.id(fb.id).runs(); assert.ok(runs.length >= 1);
    assert.ok('bold' in runs[0].formatting); assert.ok('italic' in runs[0].formatting);
    assert.ok('font' in runs[0].formatting); assert.ok('size' in runs[0].formatting);
    doc.discard();
  });
});

describe('RunHandle.formatting()', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('returns formatting properties', async () => {
    var out = freshCopy('run-fmt'); var doc = docex(out); var map = await doc.map();
    var fb = map.allParagraphs.find(function(p) { return p.type === 'body' && p.text.length > 10; }); assert.ok(fb);
    var runs = doc.id(fb.id).runs(); assert.ok(runs.length > 0);
    var rw = runs.find(function(r) { return r.textId; });
    if (rw) { var fmt = doc.id(fb.id).run(rw.textId).formatting(); assert.ok(typeof fmt.bold === 'boolean'); }
    doc.discard();
  });
});

describe('RunHandle.splitAt()', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('creates two runs', async () => {
    var out = freshCopy('run-split'); var doc = docex(out); var map = await doc.map();
    var fb = map.allParagraphs.find(function(p) { return p.type === 'body' && p.text.length > 20; }); assert.ok(fb);
    var runs = doc.id(fb.id).runs(); var rw = runs.find(function(r) { return r.textId && r.text.length > 5; });
    if (rw) {
      var rh = doc.id(fb.id).run(rw.textId); var orig = rh.text(); var ids = rh.splitAt(3);
      assert.ok(ids[0]); assert.ok(ids[1]); assert.notEqual(ids[0], ids[1]);
      var t1 = doc.id(fb.id).run(ids[0]).text(); var t2 = doc.id(fb.id).run(ids[1]).text();
      assert.equal(t1 + t2, orig);
    }
    doc.discard();
  });
});

describe('ParagraphHandle.mergeRuns()', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('reduces run count', async () => {
    var out = freshCopy('runs-merge'); var doc = docex(out); var map = await doc.map();
    var fb = map.allParagraphs.find(function(p) { return p.type === 'body' && p.text.length > 20; }); assert.ok(fb);
    var h = doc.id(fb.id); var runs = h.runs();
    if (runs.length === 1 && runs[0].textId) {
      doc.id(fb.id).run(runs[0].textId).splitAt(5); var after = h.runs(); assert.ok(after.length > 1);
      var m = h.mergeRuns(); assert.ok(m >= 1);
    } else {
      var m = h.mergeRuns(); assert.ok(typeof m === 'number');
    }
    doc.discard();
  });
});

describe('TableHandle', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('dimensions() returns correct counts', async () => {
    var out = freshCopy('tbl-dims'); var doc = docex(out); doc.untracked();
    doc.after('Introduction').table([['H1', 'H2', 'H3'], ['A', 'B', 'C'], ['D', 'E', 'F']], { caption: 'Table 1. Test' });
    await doc.save(out);
    var d2 = docex(out); await d2.map(); var t = d2.table(0); var dm = t.dimensions();
    assert.equal(dm.rows, 3); assert.equal(dm.cols, 3); d2.discard();
  });
  it('cell().text() returns content', async () => {
    var out = freshCopy('tbl-cell'); var doc = docex(out); doc.untracked();
    doc.after('Introduction').table([['Name', 'Val'], ['Alpha', '42']], { caption: 'Table 1. V' });
    await doc.save(out);
    var d2 = docex(out); await d2.map(); var t = d2.table(0);
    assert.equal(t.cell(0, 0).text(), 'Name'); assert.equal(t.cell(1, 1).text(), '42'); d2.discard();
  });
  it('cell().setText() changes content', async () => {
    var out = freshCopy('tbl-set'); var doc = docex(out); doc.untracked();
    doc.after('Introduction').table([['C1', 'C2'], ['old', 'val']], { caption: 'Table 1. T' });
    await doc.save(out);
    var d2 = docex(out); await d2.map(); var t = d2.table(0);
    t.cell(1, 0).setText('new'); assert.equal(t.cell(1, 0).text(), 'new'); d2.discard();
  });
  it('addRow() adds a row', async () => {
    var out = freshCopy('tbl-add'); var doc = docex(out); doc.untracked();
    doc.after('Introduction').table([['A', 'B'], ['1', '2']], { caption: 'Table 1. S' });
    await doc.save(out);
    var d2 = docex(out); await d2.map(); var t = d2.table(0);
    assert.equal(t.dimensions().rows, 2); t.addRow(['3', '4']); assert.equal(t.dimensions().rows, 3);
    assert.equal(t.cell(2, 0).text(), '3'); d2.discard();
  });
  it('data() returns 2D array', async () => {
    var out = freshCopy('tbl-data'); var doc = docex(out); doc.untracked();
    doc.after('Introduction').table([['X', 'Y'], ['10', '20'], ['30', '40']], { caption: 'Table 1. N' });
    await doc.save(out);
    var d2 = docex(out); await d2.map(); var t = d2.table(0); var data = t.data();
    assert.equal(data.length, 3); assert.equal(data[0][0], 'X'); assert.equal(data[2][1], '40'); d2.discard();
  });
});

describe('FigureHandle', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('dimensions() returns inches', async () => {
    var out = freshCopy('fig-dims'); var doc = docex(out); doc.untracked();
    var img = path.join(__dirname, 'fixtures', 'test-image.png');
    if (!fs.existsSync(img)) { doc.discard(); return; }
    doc.after('Introduction').figure(img, 'Figure 1. Test'); await doc.save(out);
    var d2 = docex(out); await d2.map(); var f = d2.figure(0); var dm = f.dimensions();
    assert.ok(dm.width > 0); assert.ok(dm.height > 0); d2.discard();
  });
  it('altText() and setAltText() work', async () => {
    var out = freshCopy('fig-alt'); var doc = docex(out); doc.untracked();
    var img = path.join(__dirname, 'fixtures', 'test-image.png');
    if (!fs.existsSync(img)) { doc.discard(); return; }
    doc.after('Introduction').figure(img, 'Figure 1. Alt'); await doc.save(out);
    var d2 = docex(out); await d2.map(); var f = d2.figure(0);
    f.setAltText('Descriptive text'); assert.equal(f.altText(), 'Descriptive text');
    f.setAltText('Updated'); assert.equal(f.altText(), 'Updated'); d2.discard();
  });
  it('caption() and setCaption() work', async () => {
    var out = freshCopy('fig-cap'); var doc = docex(out); doc.untracked();
    var img = path.join(__dirname, 'fixtures', 'test-image.png');
    if (!fs.existsSync(img)) { doc.discard(); return; }
    doc.after('Introduction').figure(img, 'Figure 1. Original cap'); await doc.save(out);
    var d2 = docex(out); await d2.map(); var f = d2.figure(0);
    var cap = f.caption(); assert.ok(cap.includes('Figure 1'));
    f.setCaption('Figure 1. New caption'); assert.ok(f.caption().includes('Figure 1. New caption')); d2.discard();
  });
});

describe('Comments.anchor()', () => {
  let docex, Comments; before(() => { docex = require('../src/docex'); Comments = require('../src/comments').Comments; });
  it('returns anchor info', async () => {
    var out = freshCopy('cmt-anchor'); var doc = docex(out); doc.untracked();
    doc.comment('automated', 'Test comment', { by: 'Test' }); await doc.save(out);
    var d2 = docex(out); await d2.map(); var ws = d2._workspace;
    var cmts = Comments.list(ws); assert.ok(cmts.length > 0);
    var a = Comments.anchor(ws, cmts[0].id); assert.ok(a.text.length > 0); d2.discard();
  });
});

describe('Comments.edit()', () => {
  let docex, Comments; before(() => { docex = require('../src/docex'); Comments = require('../src/comments').Comments; });
  it('changes comment text', async () => {
    var out = freshCopy('cmt-edit'); var doc = docex(out); doc.untracked();
    doc.comment('automated', 'Original text', { by: 'Ed' }); await doc.save(out);
    var d2 = docex(out); await d2.map(); var ws = d2._workspace;
    var cmts = Comments.list(ws); assert.ok(cmts.length > 0);
    Comments.edit(ws, cmts[0].id, 'Edited text');
    var updated = Comments.list(ws); var u = updated.find(function(c) { return c.id === cmts[0].id; });
    assert.ok(u); assert.ok(u.text.includes('Edited text')); d2.discard();
  });
});

describe('Comments.replies()', () => {
  let docex, Comments; before(() => { docex = require('../src/docex'); Comments = require('../src/comments').Comments; });
  it('returns threaded replies', async () => {
    var out = freshCopy('cmt-reply'); var doc = docex(out); doc.untracked();
    doc.comment('automated', 'Parent', { by: 'Author' }); await doc.save(out);
    var d2 = docex(out); await d2.map(); var ws = d2._workspace;
    var cmts = Comments.list(ws); assert.ok(cmts.length > 0);
    Comments.reply(ws, cmts[0].id, 'Reply text', { by: 'Reviewer' });
    var replies = Comments.replies(ws, cmts[0].id);
    assert.ok(replies.length >= 1); assert.ok(replies[0].text.includes('Reply text')); d2.discard();
  });
});

describe('ParagraphHandle.injectXml()', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('injects XML before text', async () => {
    var out = freshCopy('inject-xml'); var doc = docex(out); var map = await doc.map();
    var fb = map.allParagraphs.find(function(p) { return p.type === 'body' && p.text.length > 20; }); assert.ok(fb);
    var h = doc.id(fb.id); var ot = h.text; var tw = ot.slice(5, 15);
    var bk = '<w:bookmarkStart w:id="999" w:name="test"/><w:bookmarkEnd w:id="999"/>';
    h.injectXml(bk, { before: tw }); assert.ok(h.getXml().includes('w:bookmarkStart')); doc.discard();
  });
});

describe('doc.table() integration', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('returns TableHandle', async () => {
    var out = freshCopy('doc-tbl'); var doc = docex(out); doc.untracked();
    doc.after('Introduction').table([['H1', 'H2'], ['V1', 'V2']], { caption: 'Table 1. Int' });
    await doc.save(out);
    var d2 = docex(out); await d2.map(); var t = d2.table(0);
    assert.ok(t); assert.equal(t.dimensions().rows, 2); d2.discard();
  });
});

describe('doc.figure() integration', () => {
  let docex; before(() => { docex = require('../src/docex'); });
  it('returns FigureHandle', async () => {
    var out = freshCopy('doc-fig'); var doc = docex(out); doc.untracked();
    var img = path.join(__dirname, 'fixtures', 'test-image.png');
    if (!fs.existsSync(img)) { doc.discard(); return; }
    doc.after('Introduction').figure(img, 'Figure 1. Int'); await doc.save(out);
    var d2 = docex(out); await d2.map(); var f = d2.figure(0);
    assert.ok(f); assert.ok(f.dimensions().width > 0); d2.discard();
  });
});
