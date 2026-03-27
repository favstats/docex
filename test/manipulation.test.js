/**
 * manipulation.test.js -- Tests for v0.4.8 document manipulation
 */
const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
function freshCopy(n) { const o = path.join(OUTPUT_DIR, n + '.docx'); fs.copyFileSync(FIXTURE, o); return o; }
function readDocxXml(p, f) { const t = fs.mkdtempSync('/tmp/docex-test-'); execFileSync('unzip', ['-o', p, '-d', t], { stdio: 'pipe' }); const fp = path.join(t, f); let c = ''; if (fs.existsSync(fp)) c = fs.readFileSync(fp, 'utf8'); execFileSync('rm', ['-rf', t], { stdio: 'pipe' }); return c; }

describe('replaceTable', () => {
  it('replaces table data by number', async () => {
    const file = freshCopy('replaceTable-basic');
    const docex = require('../src/docex');
    const doc = docex(file); doc.author('Test').untracked();
    doc.after('Introduction').table([['H1','H2'],['a','b'],['c','d']], { caption: 'Table 1. Test' });
    await doc.save();
    const doc2 = docex(file); doc2.author('Test').untracked();
    const ws = await doc2._ensureWorkspace();
    const { Layout } = require('../src/layout');
    Layout.replaceTable(ws, 1, [['X','Y','Z'],['1','2','3']]);
    const result = ws.save();
    assert.ok(result.verified);
    const xml = readDocxXml(file, 'word/document.xml');
    assert.ok(xml.includes('>X<'));
    assert.ok(xml.includes('>Y<'));
  });
  it('throws for non-existent table number', () => {
    const file = freshCopy('replaceTable-missing');
    const { Workspace } = require('../src/workspace');
    const { Layout } = require('../src/layout');
    const ws = Workspace.open(file);
    assert.throws(() => Layout.replaceTable(ws, 99, [['a','b']]), /Table 99 not found/);
    ws.cleanup();
  });
});

describe('pageBreakBefore', () => {
  it('adds pageBreakBefore to a heading', () => {
    const file = freshCopy('pageBreak-basic');
    const { Workspace } = require('../src/workspace');
    const { Layout } = require('../src/layout');
    const { Paragraphs } = require('../src/paragraphs');
    const xml = require('../src/xml');
    const ws = Workspace.open(file);
    const headings = Paragraphs.headings(ws);
    assert.ok(headings.length > 0);
    Layout.pageBreakBefore(ws, headings[0].text);
    const paragraphs = xml.findParagraphs(ws.docXml);
    let found = false;
    for (const p of paragraphs) { if (p.text.includes(headings[0].text) && p.xml.includes('<w:pageBreakBefore')) { found = true; break; } }
    assert.ok(found, 'Heading should have pageBreakBefore');
    ws.cleanup();
  });
  it('throws for non-existent heading', () => {
    const file = freshCopy('pageBreak-missing');
    const { Workspace } = require('../src/workspace');
    const { Layout } = require('../src/layout');
    const ws = Workspace.open(file);
    assert.throws(() => Layout.pageBreakBefore(ws, 'NONEXISTENT HEADING XYZ'), /Heading not found/);
    ws.cleanup();
  });
});

describe('ensureHeadingHierarchy', () => {
  it('fixes heading level skips', () => {
    const file = freshCopy('hierarchy-fix');
    const { Workspace } = require('../src/workspace');
    const { Layout } = require('../src/layout');
    const { Paragraphs } = require('../src/paragraphs');
    const ws = Workspace.open(file);
    let docXml = ws.docXml;
    const bodyEnd = docXml.indexOf('</w:body>');
    ws.docXml = docXml.slice(0, bodyEnd) + '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Test H1</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading3"/><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Skipped H3</w:t></w:r></w:p>' + docXml.slice(bodyEnd);
    const fixes = Layout.ensureHeadingHierarchy(ws);
    assert.ok(fixes >= 1, 'Should fix at least one skip');
    ws.cleanup();
  });
  it('returns 0 for valid hierarchy', () => {
    const file = freshCopy('hierarchy-ok');
    const { Workspace } = require('../src/workspace');
    const { Layout } = require('../src/layout');
    const ws = Workspace.open(file);
    const fixes = Layout.ensureHeadingHierarchy(ws);
    assert.ok(typeof fixes === 'number');
    ws.cleanup();
  });
});

describe('mergeParagraphs', () => {
  it('combines two paragraphs into one', () => {
    const file = freshCopy('merge-basic');
    const { Workspace } = require('../src/workspace');
    const { Layout } = require('../src/layout');
    const xml = require('../src/xml');
    const ws = Workspace.open(file);
    const paragraphs = xml.findParagraphs(ws.docXml);
    let p1 = null, p2 = null;
    for (let i = 0; i < paragraphs.length - 1; i++) {
      const t1 = xml.extractTextDecoded(paragraphs[i].xml).trim();
      const t2 = xml.extractTextDecoded(paragraphs[i+1].xml).trim();
      const id1 = paragraphs[i].xml.match(/w14:paraId="([^"]+)"/);
      const id2 = paragraphs[i+1].xml.match(/w14:paraId="([^"]+)"/);
      if (t1 && t2 && id1 && id2 && t1.length > 5 && t2.length > 5) { p1 = { id: id1[1] }; p2 = { id: id2[1] }; break; }
    }
    assert.ok(p1 && p2);
    Layout.mergeParagraphs(ws, p1.id, p2.id);
    const newParas = xml.findParagraphs(ws.docXml);
    let p2Found = false;
    for (const p of newParas) { const pid = p.xml.match(/w14:paraId="([^"]+)"/); if (pid && pid[1] === p2.id) p2Found = true; }
    assert.ok(!p2Found, 'Second paragraph should be removed');
    ws.cleanup();
  });
});

describe('splitParagraph', () => {
  it('creates two paragraphs from one', () => {
    const file = freshCopy('split-basic');
    const { Workspace } = require('../src/workspace');
    const { Layout } = require('../src/layout');
    const xml = require('../src/xml');
    const ws = Workspace.open(file);
    const paragraphs = xml.findParagraphs(ws.docXml);
    let target = null;
    for (const p of paragraphs) {
      const text = xml.extractTextDecoded(p.xml).trim();
      const pid = p.xml.match(/w14:paraId="([^"]+)"/);
      if (text && pid && text.length > 20 && text.includes(' ')) { target = { text, id: pid[1] }; break; }
    }
    assert.ok(target);
    const words = target.text.split(' ');
    const splitWord = words[Math.floor(words.length / 2)];
    const newParaId = Layout.splitParagraph(ws, target.id, splitWord);
    assert.ok(newParaId && typeof newParaId === 'string');
    const newParas = xml.findParagraphs(ws.docXml);
    let fo = false, fn = false;
    for (const p of newParas) { const pid = p.xml.match(/w14:paraId="([^"]+)"/); if (pid) { if (pid[1] === target.id) fo = true; if (pid[1] === newParaId) fn = true; } }
    assert.ok(fo, 'Original should exist');
    assert.ok(fn, 'New should exist');
    ws.cleanup();
  });
  it('throws when text not found', () => {
    const file = freshCopy('split-notfound');
    const { Workspace } = require('../src/workspace');
    const { Layout } = require('../src/layout');
    const xml = require('../src/xml');
    const ws = Workspace.open(file);
    const paragraphs = xml.findParagraphs(ws.docXml);
    let tid = null;
    for (const p of paragraphs) { const pid = p.xml.match(/w14:paraId="([^"]+)"/); if (pid && xml.extractTextDecoded(p.xml).trim()) { tid = pid[1]; break; } }
    assert.throws(() => Layout.splitParagraph(ws, tid, 'XYZNONEXISTENT123'), /Text not found/);
    ws.cleanup();
  });
});

describe('manipulation via fluent API', () => {
  it('doc.pageBreakBefore() works', async () => {
    const file = freshCopy('api-pageBreak');
    const docex = require('../src/docex');
    const doc = docex(file); doc.author('Test').untracked();
    const headings = await doc.headings();
    assert.ok(headings.length > 0);
    await doc.pageBreakBefore(headings[0].text);
    const result = await doc.save();
    assert.ok(result.verified);
    const docXml = readDocxXml(file, 'word/document.xml');
    assert.ok(docXml.includes('pageBreakBefore'));
  });
  it('doc.ensureHeadingHierarchy() returns a number', async () => {
    const file = freshCopy('api-hierarchy');
    const docex = require('../src/docex');
    const doc = docex(file); doc.author('Test').untracked();
    const fixes = await doc.ensureHeadingHierarchy();
    assert.ok(typeof fixes === 'number');
  });
});
