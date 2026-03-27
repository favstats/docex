/**
 * workflow.test.js -- Tests for v0.4.10 workflow tools
 */
const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
function freshCopy(n) { const o = path.join(OUTPUT_DIR, n + '.docx'); fs.copyFileSync(FIXTURE, o); return o; }

describe('Workflow.todo', () => {
  it('returns array', () => {
    const file = freshCopy('wf-todo');
    const { Workspace } = require('../src/workspace');
    const { Workflow } = require('../src/workflow');
    const ws = Workspace.open(file);
    assert.ok(Array.isArray(Workflow.todo(ws)));
    ws.cleanup();
  });
  it('detects TODO patterns', () => {
    const file = freshCopy('wf-todo-detect');
    const { Workspace } = require('../src/workspace');
    const { Workflow } = require('../src/workflow');
    const ws = Workspace.open(file);
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>TODO: Write conclusion.</w:t></w:r></w:p></w:body>');
    const todos = Workflow.todo(ws);
    assert.ok(todos.filter(t => t.source === 'body').length >= 1);
    ws.cleanup();
  });
});

describe('Workflow.progress', () => {
  it('returns per-section status', () => {
    const file = freshCopy('wf-progress');
    const { Workspace } = require('../src/workspace');
    const { Workflow } = require('../src/workflow');
    const ws = Workspace.open(file);
    const sections = Workflow.progress(ws);
    assert.ok(sections.length > 0);
    for (const s of sections) {
      assert.ok(['done','draft','empty'].includes(s.status));
      assert.ok(typeof s.wordCount === 'number');
    }
    ws.cleanup();
  });
});

describe('Workflow.tocPreview', () => {
  it('renders headings as string', () => {
    const file = freshCopy('wf-toc');
    const { Workspace } = require('../src/workspace');
    const { Workflow } = require('../src/workflow');
    const ws = Workspace.open(file);
    const toc = Workflow.tocPreview(ws);
    assert.ok(typeof toc === 'string' && toc.length > 0);
    assert.ok(toc.includes('Introduction'));
    ws.cleanup();
  });
  it('returns message for no headings', () => {
    const file = freshCopy('wf-toc-empty');
    const { Workspace } = require('../src/workspace');
    const { Workflow } = require('../src/workflow');
    const ws = Workspace.open(file);
    ws.docXml = ws.docXml.replace(/<w:body>[\s\S]*<\/w:body>/, '<w:body><w:p><w:r><w:t>Just text.</w:t></w:r></w:p><w:sectPr/></w:body>');
    assert.ok(Workflow.tocPreview(ws).includes('No headings'));
    ws.cleanup();
  });
});

describe('Workflow.figureList', () => {
  it('shows figure captions with page estimates', () => {
    const file = freshCopy('wf-figlist');
    const { Workspace } = require('../src/workspace');
    const { Workflow } = require('../src/workflow');
    const ws = Workspace.open(file);
    ws.docXml = ws.docXml.replace('</w:body>', '<w:p><w:r><w:t>Figure 1. Classification funnel</w:t></w:r></w:p></w:body>');
    const r = Workflow.figureList(ws);
    assert.ok(r.includes('Figure 1'));
    assert.ok(r.includes('~p.'));
    ws.cleanup();
  });
  it('returns no-figures message', () => {
    const file = freshCopy('wf-figlist-none');
    const { Workspace } = require('../src/workspace');
    const { Workflow } = require('../src/workflow');
    const ws = Workspace.open(file);
    ws.docXml = ws.docXml.replace(/<w:body>[\s\S]*<\/w:body>/, '<w:body><w:p><w:r><w:t>No figures.</w:t></w:r></w:p><w:sectPr/></w:body>');
    assert.ok(Workflow.figureList(ws).includes('No figures'));
    ws.cleanup();
  });
});

describe('workflow fluent API', () => {
  it('doc.todo() returns array', async () => {
    const file = freshCopy('api-todo');
    const docex = require('../src/docex');
    assert.ok(Array.isArray(await docex(file).todo()));
  });
  it('doc.progress() returns sections', async () => {
    const file = freshCopy('api-progress');
    const docex = require('../src/docex');
    const s = await docex(file).progress();
    assert.ok(Array.isArray(s) && s.length > 0);
  });
  it('doc.tocPreview() returns string', async () => {
    const file = freshCopy('api-toc');
    const docex = require('../src/docex');
    assert.ok(typeof (await docex(file).tocPreview()) === 'string');
  });
  it('doc.figureList() returns string', async () => {
    const file = freshCopy('api-figlist');
    const docex = require('../src/docex');
    assert.ok(typeof (await docex(file).figureList()) === 'string');
  });
});
