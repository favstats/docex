const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function freshCopy(n) { const o = path.join(OUTPUT_DIR, `${n}.docx`); fs.copyFileSync(FIXTURE, o); return o; }
function readDocxXml(p, x) { const t = fs.mkdtempSync('/tmp/docex-test-'); execFileSync('unzip', ['-o', p, '-d', t], { stdio: 'pipe' }); const f = path.join(t, x); let c = ''; if (fs.existsSync(f)) c = fs.readFileSync(f, 'utf8'); execFileSync('rm', ['-rf', t], { stdio: 'pipe' }); return c; }

describe('Presets.setFont', () => {
  it('changes the default body font in docDefaults', () => { const f = freshCopy('font-set-body'); const { Workspace } = require('../src/workspace'); const { Presets } = require('../src/presets'); const ws = Workspace.open(f); Presets.setFont(ws, 'CMU Serif'); assert.ok(ws.stylesXml.includes('w:ascii="CMU Serif"')); assert.ok(ws.stylesXml.includes('w:hAnsi="CMU Serif"')); ws.cleanup(); });
  it('updates docDefaults rPrDefault', () => { const f = freshCopy('font-set-body2'); const { Workspace } = require('../src/workspace'); const { Presets } = require('../src/presets'); const ws = Workspace.open(f); Presets.setFont(ws, 'Arial'); const m = ws.stylesXml.match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/); if (m) assert.ok(m[0].includes('Arial')); ws.cleanup(); });
  it('persists through save and reload', async () => { const f = freshCopy('font-persist'); const docex = require('../src/docex'); const doc = docex(f); await doc.font('Garamond'); await doc.save(); const s = readDocxXml(f, 'word/styles.xml'); assert.ok(s.includes('Garamond')); });
});

describe('Presets.setFontSize', () => {
  it('sets 11pt as 22 half-points', () => { const f = freshCopy('font-size-11'); const { Workspace } = require('../src/workspace'); const { Presets } = require('../src/presets'); const ws = Workspace.open(f); Presets.setFontSize(ws, 11); assert.ok(ws.stylesXml.includes('w:val="22"')); ws.cleanup(); });
  it('sets 12pt as 24 half-points', () => { const f = freshCopy('font-size-12'); const { Workspace } = require('../src/workspace'); const { Presets } = require('../src/presets'); const ws = Workspace.open(f); Presets.setFontSize(ws, 12); assert.ok(ws.stylesXml.includes('w:val="24"')); ws.cleanup(); });
});

describe('Presets.setHeadingFont', () => {
  it('does not throw', () => { const f = freshCopy('font-heading'); const { Workspace } = require('../src/workspace'); const { Presets } = require('../src/presets'); const ws = Workspace.open(f); Presets.setHeadingFont(ws, 'Helvetica'); assert.ok(typeof ws.stylesXml === 'string'); ws.cleanup(); });
});

describe('Presets.setLinkColor', () => {
  it('sets hyperlink color', () => { const f = freshCopy('font-link'); const { Workspace } = require('../src/workspace'); const { Presets } = require('../src/presets'); const ws = Workspace.open(f); Presets.setLinkColor(ws, '3498DB'); assert.ok(ws.stylesXml.includes('3498DB')); ws.cleanup(); });
});

describe('Presets.setParagraphSpacing', () => {
  it('sets spacing with before/after/line', () => { const f = freshCopy('font-spacing'); const { Workspace } = require('../src/workspace'); const { Presets } = require('../src/presets'); const ws = Workspace.open(f); Presets.setParagraphSpacing(ws, { before: 6, after: 12, line: 480 }); assert.ok(ws.stylesXml.includes('w:before="120"')); assert.ok(ws.stylesXml.includes('w:after="240"')); assert.ok(ws.stylesXml.includes('w:line="480"')); ws.cleanup(); });
  it('handles line-only spacing', () => { const f = freshCopy('font-spacing-line'); const { Workspace } = require('../src/workspace'); const { Presets } = require('../src/presets'); const ws = Workspace.open(f); Presets.setParagraphSpacing(ws, { line: 360 }); assert.ok(ws.stylesXml.includes('w:line="360"')); ws.cleanup(); });
});

describe('Fluent API', () => {
  it('font() returns engine', async () => { const f = freshCopy('font-fluent'); const docex = require('../src/docex'); const doc = docex(f); const r = await doc.font('Georgia'); assert.ok(r === doc); });
  it('fontSize() returns engine', async () => { const f = freshCopy('font-fluent-sz'); const docex = require('../src/docex'); const doc = docex(f); const r = await doc.fontSize(14); assert.ok(r === doc); });
});
