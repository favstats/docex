const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
function freshCopy(n) { const o = path.join(OUTPUT_DIR, n + '.docx'); fs.copyFileSync(FIXTURE, o); return o; }

describe('Redact.redact', () => {
  it('replaces text with [REDACTED] by default', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const xml = require('../src/xml');
    const ws = Workspace.open(freshCopy('col-red-def'));
    const r = Redact.redact(ws, '268,635');
    assert.ok(r.count > 0);
    assert.equal(r.replacement, '[REDACTED]');
    const texts = xml.findParagraphs(ws.docXml).map(p => p.text).join(' ');
    assert.ok(!texts.includes('268,635'));
    assert.ok(texts.includes('[REDACTED]'));
    ws.cleanup();
  });
  it('uses custom replacement', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const ws = Workspace.open(freshCopy('col-red-cust'));
    const r = Redact.redact(ws, 'Meta', '[CO]');
    assert.ok(r.count > 0);
    assert.equal(r.replacement, '[CO]');
    ws.cleanup();
  });
  it('returns count 0 when not found', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const ws = Workspace.open(freshCopy('col-red-nf'));
    assert.equal(Redact.redact(ws, 'XYZNONEXISTENT').count, 0);
    ws.cleanup();
  });
  it('stores mapping', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const ws = Workspace.open(freshCopy('col-red-map'));
    Redact.redact(ws, '268,635', '[NUM]');
    const m = Redact.listRedactions(ws);
    assert.ok(m.length > 0);
    assert.equal(m[0].original, '268,635');
    ws.cleanup();
  });
});

describe('Redact.unredact', () => {
  it('restores redacted text', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const xml = require('../src/xml');
    const ws = Workspace.open(freshCopy('col-unred'));
    Redact.redact(ws, '268,635', '[N]');
    const r = Redact.unredact(ws);
    assert.ok(r.count > 0);
    const texts = xml.findParagraphs(ws.docXml).map(p => p.text).join(' ');
    assert.ok(texts.includes('268,635'));
    ws.cleanup();
  });
  it('returns 0 when no redactions', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const ws = Workspace.open(freshCopy('col-unred-0'));
    assert.equal(Redact.unredact(ws).count, 0);
    ws.cleanup();
  });
  it('clears mappings', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const ws = Workspace.open(freshCopy('col-unred-clr'));
    Redact.redact(ws, '268,635');
    Redact.unredact(ws);
    assert.deepEqual(Redact.listRedactions(ws), []);
    ws.cleanup();
  });
});

describe('Redact.listRedactions', () => {
  it('returns empty when no redactions', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const ws = Workspace.open(freshCopy('col-list-0'));
    assert.deepEqual(Redact.listRedactions(ws), []);
    ws.cleanup();
  });
  it('lists multiple', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const ws = Workspace.open(freshCopy('col-list-m'));
    Redact.redact(ws, '268,635', '[A]');
    Redact.redact(ws, '1,329', '[B]');
    const m = Redact.listRedactions(ws);
    assert.ok(m.length >= 2);
    ws.cleanup();
  });
});

describe('Redact round-trip', () => {
  it('text matches after redact/unredact', () => {
    const { Workspace } = require('../src/workspace');
    const { Redact } = require('../src/redact');
    const xml = require('../src/xml');
    const ws = Workspace.open(freshCopy('col-rt'));
    const orig = xml.findParagraphs(ws.docXml).map(p => p.text);
    Redact.redact(ws, '268,635', '[R]');
    Redact.unredact(ws);
    assert.deepEqual(xml.findParagraphs(ws.docXml).map(p => p.text), orig);
    ws.cleanup();
  });
});

describe('Presets.compareStyles', () => {
  it('returns changes array', () => {
    const { Workspace } = require('../src/workspace');
    const { Presets } = require('../src/presets');
    const ws = Workspace.open(freshCopy('col-cmp'));
    const r = Presets.compareStyles(ws, 'academic');
    assert.ok(Array.isArray(r.changes));
    ws.cleanup();
  });
  it('throws for unknown preset', () => {
    const { Workspace } = require('../src/workspace');
    const { Presets } = require('../src/presets');
    const ws = Workspace.open(freshCopy('col-cmp-e'));
    assert.throws(() => Presets.compareStyles(ws, 'nonexistent'), /Unknown preset/);
    ws.cleanup();
  });
});

describe('API integration', () => {
  it('doc.redact works', async () => {
    const docex = require('../src/docex');
    const r = await docex(freshCopy('col-api-r')).redact('268,635', '[X]');
    assert.ok(r.count > 0);
  });
  it('doc.unredact works', async () => {
    const docex = require('../src/docex');
    const doc = docex(freshCopy('col-api-u'));
    await doc.redact('268,635');
    assert.ok((await doc.unredact()).count > 0);
  });
  it('doc.compareStyles works', async () => {
    const docex = require('../src/docex');
    const r = await docex(freshCopy('col-api-cs')).compareStyles('polcomm');
    assert.ok(Array.isArray(r.changes));
  });
  it('redact survives save/reopen', async () => {
    const docex = require('../src/docex');
    const f = freshCopy('col-api-sv');
    const d1 = docex(f);
    await d1.redact('268,635', '[H]');
    await d1.save();
    const d2 = docex(f);
    const ws = await d2._ensureWorkspace();
    const { Redact } = require('../src/redact');
    const m = Redact.listRedactions(ws);
    assert.ok(m.length > 0);
    assert.equal(m[0].original, '268,635');
  });
});
