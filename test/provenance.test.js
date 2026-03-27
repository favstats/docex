const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
function freshCopy(n) { const o = path.join(OUTPUT_DIR, `${n}.docx`); fs.copyFileSync(FIXTURE, o); return o; }

describe('changelog', () => {
  it('records operations', () => { const f = freshCopy('cl-basic'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); assert.equal(Provenance.getChangelog(ws).length, 0); Provenance.appendChangelog(ws, [{ timestamp: '2026-03-27T10:00:00Z', operation: 'replace', author: 'Test', description: 'Fixed typo' }]); assert.equal(Provenance.getChangelog(ws).length, 1); ws.cleanup(); });
  it('since filters by date', () => { const f = freshCopy('cl-since'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); Provenance.appendChangelog(ws, [{ timestamp: '2026-03-01T10:00:00Z', operation: 'old', author: 'A', description: '' }, { timestamp: '2026-03-20T10:00:00Z', operation: 'new', author: 'B', description: '' }]); const s = Provenance.changelogSince(ws, '2026-03-15'); assert.equal(s.length, 1); assert.equal(s[0].operation, 'new'); ws.cleanup(); });
  it('persists through save/reload', () => { const f = freshCopy('cl-persist'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); Provenance.appendChangelog(ws, [{ timestamp: '2026-03-27T12:00:00Z', operation: 'test', author: 'Bot', description: 'x' }]); ws.save(); const ws2 = Workspace.open(f); assert.ok(Provenance.getChangelog(ws2).length >= 1); ws2.cleanup(); });
});

describe('certify', () => {
  it('stores hash and label', () => { const f = freshCopy('cert-basic'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); Provenance.certify(ws, 'test label'); const c = Provenance.certifications(ws); assert.equal(c.length, 1); assert.equal(c[0].label, 'test label'); assert.equal(c[0].hash.length, 64); ws.cleanup(); });
  it('verify confirms match', () => { const f = freshCopy('cert-verify'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); Provenance.certify(ws, 'x'); assert.equal(Provenance.verifyCertification(ws).certified, true); ws.cleanup(); });
  it('detects changes', () => { const f = freshCopy('cert-changed'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); Provenance.certify(ws, 'pre'); ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>Added</w:t></w:r></w:p></w:body>'); assert.equal(Provenance.verifyCertification(ws).certified, false); ws.cleanup(); });
  it('supports multiple certs', () => { const f = freshCopy('cert-multi'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); Provenance.certify(ws, 'first'); Provenance.certify(ws, 'second'); assert.equal(Provenance.certifications(ws).length, 2); ws.cleanup(); });
});

describe('origin', () => {
  it('returns default message when none set', () => { const f = freshCopy('origin-def'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); assert.ok(Provenance.origin(ws).includes('No origin')); ws.cleanup(); });
  it('setOrigin and origin read', () => { const f = freshCopy('origin-set'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); Provenance.setOrigin(ws, { version: 'v0.4.9', date: '2026-03-27', template: 'polcomm', tool: 'docex' }); const r = Provenance.origin(ws); assert.ok(r.includes('docex')); assert.ok(r.includes('v0.4.9')); ws.cleanup(); });
  it('includes op count from changelog', () => { const f = freshCopy('origin-ops'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); Provenance.setOrigin(ws, { version: 'v0.4.9', date: '2026-03-27', tool: 'docex' }); Provenance.appendChangelog(ws, [{ timestamp: '2026-03-27T10:00:00Z', operation: 'r', author: 'T', description: '' }, { timestamp: '2026-03-27T11:00:00Z', operation: 'i', author: 'T', description: '' }]); assert.ok(Provenance.origin(ws).includes('2 operations')); ws.cleanup(); });
  it('uncertified with no certs', () => { const f = freshCopy('verify-no'); const { Workspace } = require('../src/workspace'); const { Provenance } = require('../src/provenance'); const ws = Workspace.open(f); assert.equal(Provenance.verifyCertification(ws).certified, false); ws.cleanup(); });
});
