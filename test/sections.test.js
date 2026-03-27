const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
function freshCopy(n) { const o = path.join(OUTPUT_DIR, n + '.docx'); fs.copyFileSync(FIXTURE, o); return o; }

describe('Sections.outline', () => {
  it('extracts headings with structure', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-outline'));
    const outline = await doc.outline();
    assert.ok(outline.length >= 4);
    assert.equal(outline[0].text, 'Introduction');
    assert.equal(outline[0].level, 1);
    assert.ok(typeof outline[0].paragraphCount === 'number');
    assert.ok(typeof outline[0].figureCount === 'number');
  });
  it('includes paragraphCount', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-outline-pc'));
    const outline = await doc.outline();
    const intro = outline.find(h => h.text === 'Introduction');
    assert.ok(intro);
    assert.equal(intro.paragraphCount, 2);
  });
  it('includes paraId', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-outline-id'));
    const outline = await doc.outline();
    for (const h of outline) assert.ok(h.paraId);
  });
  it('returns empty for no headings', async () => {
    const docex = require('../src/docex');
    const r = await docex.create({ output: path.join(OUTPUT_DIR, 'no-h.docx') });
    const doc = docex(r.path);
    assert.deepEqual(await doc.outline(), []);
  });
});

describe('Sections.move', () => {
  it('moves section before another', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-move-b'));
    const r = await doc.moveSection('Discussion', { before: 'Results' });
    assert.equal(r.moved, 'Discussion');
    assert.ok(r.paragraphsMoved > 0);
    const o = await doc.outline();
    const h = o.map(x => x.text);
    assert.ok(h.indexOf('Discussion') < h.indexOf('Results'));
  });
  it('moves section after another', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-move-a'));
    await doc.moveSection('Introduction', { after: 'Methods' });
    const o = await doc.outline();
    const h = o.map(x => x.text);
    assert.ok(h.indexOf('Introduction') > h.indexOf('Methods'));
  });
  it('throws for missing source', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-move-e1'));
    await assert.rejects(() => doc.moveSection('Nonexistent', { before: 'Methods' }), /not found/i);
  });
  it('throws for missing target', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-move-e2'));
    await assert.rejects(() => doc.moveSection('Discussion', { before: 'Nonexistent' }), /not found/i);
  });
  it('preserves paragraph count', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-move-pc'));
    const before = (await doc.paragraphs()).length;
    await doc.moveSection('Discussion', { before: 'Methods' });
    assert.equal((await doc.paragraphs()).length, before);
  });
});

describe('Sections.split', () => {
  it('extracts section into new file', async () => {
    const docex = require('../src/docex');
    const f = freshCopy('sec-split');
    const out = path.join(OUTPUT_DIR, 'split-methods.docx');
    const doc = docex(f);
    const r = await doc.splitDocument('Methods', out);
    assert.ok(r.outputPath);
    assert.ok(r.paragraphsExtracted > 0);
    assert.ok(fs.existsSync(r.outputPath));
    execFileSync('unzip', ['-t', r.outputPath], { stdio: 'pipe' });
  });
  it('removes section from original', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-split-rm'));
    await doc.splitDocument('Methods', path.join(OUTPUT_DIR, 'split-rm.docx'));
    const o = await doc.outline();
    assert.ok(!o.some(h => h.text === 'Methods'));
  });
  it('throws for missing section', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-split-e'));
    await assert.rejects(() => doc.splitDocument('Nonexistent', path.join(OUTPUT_DIR, 'split-e.docx')), /not found/i);
  });
});

describe('Sections.extractAbstract', () => {
  it('returns null when no abstract', async () => {
    const docex = require('../src/docex');
    assert.equal(await docex(freshCopy('sec-abs-n')).extractAbstract(), null);
  });
  it('extracts abstract from heading section', () => {
    const { Workspace } = require('../src/workspace');
    const { Sections } = require('../src/sections');
    const xml = require('../src/xml');
    const ws = Workspace.open(freshCopy('sec-abs-y'));
    let d = ws.docXml;
    const p = xml.findParagraphs(d);
    const pPr = '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>';
    const hd = xml.buildParagraph(pPr, xml.buildRun('', 'Abstract'));
    const bd = xml.buildParagraph('', xml.buildRun('', 'Test abstract text.'));
    d = d.slice(0, p[0].start) + hd + bd + d.slice(p[0].start);
    ws.docXml = d;
    const r = Sections.extractAbstract(ws);
    assert.ok(r);
    assert.ok(r.includes('Test abstract text'));
    ws.cleanup();
  });
});

describe('Sections.duplicate', () => {
  it('copies section with new heading', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-dup'));
    const r = await doc.duplicateSection('Results', 'Results (Appendix)');
    assert.equal(r.duplicated, 'Results');
    assert.equal(r.newHeading, 'Results (Appendix)');
    const o = await doc.outline();
    const h = o.map(x => x.text);
    assert.ok(h.includes('Results'));
    assert.ok(h.includes('Results (Appendix)'));
  });
  it('assigns fresh paraIds', () => {
    const { Workspace } = require('../src/workspace');
    const { Sections } = require('../src/sections');
    const xml = require('../src/xml');
    const ws = Workspace.open(freshCopy('sec-dup-id'));
    Sections.duplicate(ws, 'Methods', 'Methods Copy');
    const ps = xml.findParagraphs(ws.docXml);
    const ids = ps.map(p => (p.xml.match(/w14:paraId="([^"]+)"/) || [])[1]).filter(Boolean);
    assert.equal(new Set(ids).size, ids.length, 'All paraIds unique');
    ws.cleanup();
  });
  it('throws for missing section', async () => {
    const docex = require('../src/docex');
    await assert.rejects(() => docex(freshCopy('sec-dup-e')).duplicateSection('X', 'Y'), /not found/i);
  });
});

describe('Sections.append', () => {
  it('adds paragraph at end of section', () => {
    const { Workspace } = require('../src/workspace');
    const { Sections } = require('../src/sections');
    const xml = require('../src/xml');
    const ws = Workspace.open(freshCopy('sec-append'));
    const before = xml.findParagraphs(ws.docXml).length;
    const r = Sections.append(ws, 'Methods', 'Appended text.');
    assert.ok(r.appended);
    assert.equal(r.section, 'Methods');
    assert.ok(r.paraId);
    assert.equal(xml.findParagraphs(ws.docXml).length, before + 1);
    ws.cleanup();
  });
  it('throws for missing section', () => {
    const { Workspace } = require('../src/workspace');
    const { Sections } = require('../src/sections');
    const ws = Workspace.open(freshCopy('sec-append-e'));
    assert.throws(() => Sections.append(ws, 'X', 'text'), /not found/i);
    ws.cleanup();
  });
});

describe('Sections round-trip', () => {
  it('outline has same headings after move', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('sec-rt'));
    const before = (await doc.outline()).map(h => h.text).sort();
    await doc.moveSection('Discussion', { before: 'Methods' });
    const after = (await doc.outline()).map(h => h.text).sort();
    assert.deepEqual(after, before);
  });
});
