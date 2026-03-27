'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const { Workspace } = require('../src/workspace');
const { Comments } = require('../src/comments');
const { DexDecompiler } = require('../src/dex-decompiler');
const { DexParser } = require('../src/dex-markdown-parser');
const { DexCompiler } = require('../src/dex-compiler');

const ABSOLUTE_CHAOS_V2 = '/mnt/storage/absolute_chaos_v2.docx';
const ABSOLUTE_CHAOS_V3 = '/mnt/storage/absolute_chaos_v3.docx';
const OUTPUT_DIR = path.join(__dirname, 'output', 'dex-package-format');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

describe('.dex package grammar', () => {
  it('rejects reserved control commands as XML element names', () => {
    const dex = [
      '\\dex[version="0.5.0"]{',
      '  \\part[path="word/document.xml" type="xml"]{',
      '    \\xml[version="1.0" encoding="UTF-8"]{}',
      '    \\dex{}',
      '  \\end{part}',
      '\\end{dex}',
      '',
    ].join('\n');

    assert.throws(() => DexParser.parse(dex), /unexpected control command/);
  });

  it('parses lowercase bare XML element names for custom XML parts', () => {
    const dex = [
      '\\dex[version="0.5.0"]{',
      '  \\part[path="customXml/docex-origin.xml" type="xml"]{',
      '    \\xml[version="1.0" encoding="UTF-8"]{}',
      '    \\docex-origin[xmlns="https://docex.dev/origin"]{',
      '      \\version{v3}',
      '      \\date{2026-03-27}',
      '      \\template{absolute_chaos_v2}',
      '      \\tool{docex}',
      '    \\end{docex-origin}',
      '  \\end{part}',
      '\\end{dex}',
      '',
    ].join('\n');

    const ast = DexParser.parse(dex);
    assert.equal(ast.parts[0].path, 'customXml/docex-origin.xml');
    assert.equal(ast.parts[0].nodes[1].name, 'docex-origin');
    assert.equal(ast.parts[0].nodes[1].children[0].name, 'version');
  });

  it('decompiles lowercase-root provenance parts from absolute_chaos_v3', {
    skip: !fs.existsSync(ABSOLUTE_CHAOS_V3),
  }, () => {
    const result = DexCompiler.assertRoundTrip(ABSOLUTE_CHAOS_V3, {
      output: path.join(OUTPUT_DIR, 'absolute-chaos-v3-provenance-roundtrip.docx'),
      strictVerify: false,
    });
    assert.equal(result.verified, true);
    assert.ok(result.dex.includes('\\part[path="word/endnotes.xml" type="xml"]{'));
    assert.ok(result.dex.includes('\\part[path="word/commentsExtended.xml" type="xml"]{'));
    assert.ok(result.dex.includes('\\part[path="word/commentsIds.xml" type="xml"]{'));
    assert.ok(result.dex.includes('\\part[path="customXml/docex-changelog.xml" type="xml"]{'));
    assert.ok(result.dex.includes('\\part[path="customXml/docex-origin.xml" type="xml"]{'));
    assert.ok(result.dex.includes('\\part[path="customXml/docex-certifications.xml" type="xml"]{'));
    assert.ok(result.dex.includes('\\docex-origin['));
    assert.ok(result.dex.includes('\\docex-changelog['));
    assert.ok(result.dex.includes('\\docex-certifications['));
  });
});

describe('Comments.reply legacy-thread support', () => {
  it('replies cleanly on a document that started without commentsExtended.xml', {
    skip: !fs.existsSync(ABSOLUTE_CHAOS_V2),
  }, () => {
    const outPath = path.join(OUTPUT_DIR, 'comments-reply-legacy.docx');
    fs.copyFileSync(ABSOLUTE_CHAOS_V2, outPath);

    const ws = Workspace.open(outPath);
    try {
      const added = Comments.add(ws, 'CONFIDENTIAL', 'Fresh top-level note', { by: 'Regression' });
      const reply = Comments.reply(ws, added.commentId, 'Fresh threaded reply', { by: 'Regression' });
      assert.ok(added.commentId >= 0);
      assert.ok(reply.commentId > added.commentId);
      assert.ok(ws.commentsExtXml.includes('w15:paraIdParent='));
      ws.save({ outputPath: outPath, backup: false });
    } finally {
      ws.cleanup();
    }

    const dex = DexDecompiler.decompile(outPath);
    assert.ok(dex.includes('Fresh top-level note'));
    assert.ok(dex.includes('Fresh threaded reply'));
  });
});
