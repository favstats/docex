/**
 * dex-format.test.js -- Tests for the .dex lossless document format
 *
 * Tests decompile, parse, compile, and round-trip operations.
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/dex-format.test.js
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function freshCopy(testName) {
  const out = path.join(OUTPUT_DIR, `dex-${testName}.docx`);
  fs.copyFileSync(FIXTURE, out);
  return out;
}

// Import modules directly (avoiding docex.js which has a missing table-handle dep)
const { Workspace } = require('../src/workspace');
const { DexDecompiler } = require('../src/dex-decompiler');
const { DexParser } = require('../src/dex-markdown-parser');
const { DexCompiler } = require('../src/dex-compiler');
const { Comments } = require('../src/comments');
const xmlLib = require('../src/xml');

function createDocxWithTrackedChanges(testName) {
  const outPath = freshCopy(testName);
  const ws = Workspace.open(outPath);
  let docXml = ws.docXml;
  const author = 'Fabio Votta';
  const date = '2026-03-27T10:00:00Z';
  const nextId = xmlLib.nextChangeId(docXml);
  const searchText = '268,635';
  const tIdx = docXml.indexOf(searchText);
  if (tIdx !== -1) {
    const rStart = docXml.lastIndexOf('<w:r', tIdx);
    const rEnd = docXml.indexOf('</w:r>', tIdx) + 6;
    const rEl = docXml.slice(rStart, rEnd);
    const rPrMatch = rEl.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
    const rPr = rPrMatch ? rPrMatch[0] : '';
    const tMatch = rEl.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/);
    const tContent = tMatch ? tMatch[1] : '';
    const textPos = tContent.indexOf(searchText);
    const beforeText = tContent.slice(0, textPos);
    const afterText = tContent.slice(textPos + searchText.length);
    let replacement = '';
    if (beforeText) {
      replacement += '<w:r>' + rPr + '<w:t xml:space="preserve">' + beforeText + '</w:t></w:r>';
    }
    replacement += xmlLib.buildDel(nextId, author, date, rPr, searchText);
    replacement += xmlLib.buildIns(nextId + 1, author, date, rPr, '300,000');
    if (afterText) {
      replacement += '<w:r>' + rPr + '<w:t xml:space="preserve">' + afterText + '</w:t></w:r>';
    }
    docXml = docXml.slice(0, rStart) + replacement + docXml.slice(rEnd);
    ws.docXml = docXml;
  }
  ws.save({ outputPath: outPath, backup: false });
  return outPath;
}

function createDocxWithComments(testName) {
  const outPath = freshCopy(testName);
  const ws = Workspace.open(outPath);
  Comments.add(ws, 'platform governance', 'Needs citation here', {
    by: 'Reviewer 1',
    date: '2026-03-24T10:00:00Z',
  });
  ws.save({ outputPath: outPath, backup: false });
  return outPath;
}

// ============================================================================
// 1. DECOMPILER TESTS
// ============================================================================

describe('DexDecompiler', () => {
  it('decompiles fixture to .dex string with headings', () => {
    const ws = Workspace.open(FIXTURE);
    const dex = DexDecompiler.decompile(ws);
    ws.cleanup();
    assert.ok(dex.includes('# Introduction'), 'should contain H1 Introduction');
    assert.ok(dex.includes('# Methods'), 'should contain H1 Methods');
    assert.ok(dex.includes('# Results'), 'should contain H1 Results');
  });

  it('decompiles fixture with paragraphs in {p} blocks', () => {
    const ws = Workspace.open(FIXTURE);
    const dex = DexDecompiler.decompile(ws);
    ws.cleanup();
    assert.ok(dex.includes('{p'), 'should contain paragraph blocks');
    assert.ok(dex.includes('{/p}'), 'should contain closing paragraph blocks');
  });

  it('preserves paraIds in output', () => {
    const ws = Workspace.open(FIXTURE);
    const dex = DexDecompiler.decompile(ws);
    ws.cleanup();
    const paraIdRe = /id:[A-Fa-f0-9]{8}/g;
    const matches = dex.match(paraIdRe);
    assert.ok(matches && matches.length > 0, 'should contain paraIds');
  });

  it('includes YAML frontmatter', () => {
    const ws = Workspace.open(FIXTURE);
    const dex = DexDecompiler.decompile(ws);
    ws.cleanup();
    assert.ok(dex.startsWith('---'), 'should start with YAML frontmatter');
    assert.ok(dex.includes('docex: "0.4.0"'), 'should contain version');
  });

  it('decompiles tracked changes with author and date', () => {
    const outPath = createDocxWithTrackedChanges('decompile-tracked');
    const ws = Workspace.open(outPath);
    const dex = DexDecompiler.decompile(ws);
    ws.cleanup();
    assert.ok(dex.includes('{del'), 'should contain tracked deletion');
    assert.ok(dex.includes('{ins'), 'should contain tracked insertion');
    assert.ok(dex.includes('Fabio Votta'), 'should preserve author name');
    assert.ok(dex.includes('{/del}'), 'should close deletion');
    assert.ok(dex.includes('{/ins}'), 'should close insertion');
  });

  it('decompiles comments', () => {
    const outPath = createDocxWithComments('decompile-comments');
    const ws = Workspace.open(outPath);
    const dex = DexDecompiler.decompile(ws);
    ws.cleanup();
    // Comments are now inline: {comment-start id:N by:"Author"} ... {comment-end id:N | text}
    assert.ok(dex.includes('{comment-start'), 'should contain comment-start marker');
    assert.ok(dex.includes('Reviewer 1'), 'should preserve comment author');
    assert.ok(dex.includes('Needs citation here'), 'should preserve comment text');
    assert.ok(dex.includes('{comment-end'), 'should contain comment-end marker');
  });
});

// ============================================================================
// 2. PARSER TESTS
// ============================================================================

describe('DexParser', () => {
  it('parses YAML frontmatter', () => {
    const dex = '---\ndocex: "0.4.0"\ntitle: "Test Document"\nauthors:\n  - name: "Author One"\n  - name: "Author Two"\n---\n\n# Introduction {id:3A2F001B}\n';
    const ast = DexParser.parse(dex);
    assert.equal(ast.frontmatter.docex, '0.4.0');
    assert.equal(ast.frontmatter.title, 'Test Document');
    assert.ok(Array.isArray(ast.frontmatter.authors));
    assert.equal(ast.frontmatter.authors.length, 2);
    assert.equal(ast.frontmatter.authors[0].name, 'Author One');
  });

  it('parses headings with levels and IDs', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n# Introduction {id:3A2F001B}\n\n## Data Collection {id:3A2F0020}\n';
    const ast = DexParser.parse(dex);
    const headings = ast.body.filter(n => n.type === 'heading');
    assert.equal(headings.length, 2);
    assert.equal(headings[0].level, 1);
    assert.equal(headings[0].text, 'Introduction');
    assert.equal(headings[0].id, '3A2F001B');
    assert.equal(headings[1].level, 2);
    assert.equal(headings[1].text, 'Data Collection');
    assert.equal(headings[1].id, '3A2F0020');
  });

  it('parses paragraphs with inline formatting', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:3A2F001C}\nThis is text with {b}bold{/b} and {i}italic{/i} words.\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    assert.ok(para, 'should have a paragraph');
    assert.equal(para.id, '3A2F001C');
    const boldRun = para.runs.find(r => r.type === 'bold');
    assert.ok(boldRun, 'should have bold run');
    assert.equal(boldRun.text, 'bold');
    const italicRun = para.runs.find(r => r.type === 'italic');
    assert.ok(italicRun, 'should have italic run');
    assert.equal(italicRun.text, 'italic');
  });

  it('parses tracked changes', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:3A2F001C}\nWe collected {del id:5 by:"Fabio Votta" date:"2026-03-27T10:00:00Z"}268,635{/del}{ins id:6 by:"Fabio Votta" date:"2026-03-27T10:00:00Z"}300,000{/ins} advertisements.\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const delRun = para.runs.find(r => r.type === 'del');
    assert.ok(delRun, 'should have del run');
    assert.equal(delRun.id, 5);
    assert.equal(delRun.author, 'Fabio Votta');
    assert.equal(delRun.text, '268,635');
    const insRun = para.runs.find(r => r.type === 'ins');
    assert.ok(insRun, 'should have ins run');
    assert.equal(insRun.id, 6);
    assert.equal(insRun.text, '300,000');
  });

  it('parses comments and replies', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{comment id:17 by:"Prof. Strict" date:"2026-03-24T10:00:00Z"}\nCite Gorwa 2019 here\n{/comment}\n\n{reply id:18 parent:17 by:"Fabio Votta" date:"2026-03-25T09:00:00Z"}\nAdded, thanks.\n{/reply}\n';
    const ast = DexParser.parse(dex);
    const comment = ast.body.find(n => n.type === 'comment');
    assert.ok(comment, 'should have a comment');
    assert.equal(comment.id, 17);
    assert.equal(comment.author, 'Prof. Strict');
    assert.equal(comment.text, 'Cite Gorwa 2019 here');
    const reply = ast.body.find(n => n.type === 'reply');
    assert.ok(reply, 'should have a reply');
    assert.equal(reply.id, 18);
    assert.equal(reply.parent, 17);
    assert.equal(reply.author, 'Fabio Votta');
    assert.equal(reply.text, 'Added, thanks.');
  });

  it('parses figure blocks', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{figure id:3A2F0025 rId:rId9 src:"word/media/image1.png" width:5943600emu height:3714750emu alt:"Funnel"}\nFigure 1. Classification funnel\n{/figure}\n';
    const ast = DexParser.parse(dex);
    const fig = ast.body.find(n => n.type === 'figure');
    assert.ok(fig, 'should have a figure');
    assert.equal(fig.id, '3A2F0025');
    assert.equal(fig.rId, 'rId9');
    assert.equal(fig.src, 'word/media/image1.png');
    assert.equal(fig.width, '5943600emu');
    assert.equal(fig.height, '3714750emu');
    assert.equal(fig.alt, 'Funnel');
    assert.ok(fig.caption.includes('Figure 1'));
  });

  it('parses table blocks', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{table style:booktabs cols:3}\n| Party | Ads | Share |\n|---|---|---|\n| PAX | 117 | 8.8% |\n| Wakker | 82 | 6.2% |\n{/table}\n';
    const ast = DexParser.parse(dex);
    const tbl = ast.body.find(n => n.type === 'table');
    assert.ok(tbl, 'should have a table');
    assert.equal(tbl.style, 'booktabs');
    assert.equal(tbl.cols, 3);
    assert.equal(tbl.rows.length, 3);
    assert.equal(tbl.rows[0][0].text, 'Party');
    assert.equal(tbl.rows[1][0].text, 'PAX');
  });

  it('parses footnotes', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:3A2F001C}\nThis is text{footnote id:2}Footnote content here.{/footnote}.\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const fnRun = para.runs.find(r => r.type === 'footnote');
    assert.ok(fnRun, 'should have a footnote');
    assert.equal(fnRun.id, 2);
    assert.equal(fnRun.text, 'Footnote content here.');
  });

  it('parses highlight and color formatting', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:3A2F001C}\n{highlight yellow}highlighted{/highlight} and {color FF0000}red text{/color}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const hlRun = para.runs.find(r => r.type === 'highlight');
    assert.ok(hlRun, 'should have highlight');
    assert.equal(hlRun.color, 'yellow');
    assert.equal(hlRun.text, 'highlighted');
    const colorRun = para.runs.find(r => r.type === 'color');
    assert.ok(colorRun, 'should have color');
    assert.equal(colorRun.color, 'FF0000');
    assert.equal(colorRun.text, 'red text');
  });

  it('parses page breaks', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n# Introduction {id:AA}\n\n{pagebreak}\n\n# Methods {id:BB}\n';
    const ast = DexParser.parse(dex);
    const pb = ast.body.find(n => n.type === 'pagebreak');
    assert.ok(pb, 'should have a pagebreak');
  });

  it('parses section properties', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n# Introduction {id:AA}\n\n{section margins:"1440 1440 1440 1440"}\n';
    const ast = DexParser.parse(dex);
    const sect = ast.body.find(n => n.type === 'section');
    assert.ok(sect, 'should have section');
    assert.equal(sect.margins, '1440 1440 1440 1440');
  });
});

// ============================================================================
// 2b. PARSER TESTS — New .dex format features
// ============================================================================

describe('DexParser — Comment anchors', () => {
  it('parses {comment-start id:N} and {comment-end id:N} inline markers', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{comment-start id:0}annotated text{comment-end id:0}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const cStart = para.runs.find(r => r.type === 'comment-start');
    assert.ok(cStart, 'should have comment-start node');
    assert.equal(cStart.id, 0);
    const cEnd = para.runs.find(r => r.type === 'comment-end');
    assert.ok(cEnd, 'should have comment-end node');
    assert.equal(cEnd.id, 0);
  });
});

describe('DexParser — Bookmarks', () => {
  it('parses {bookmark-start id:N name:"ref"} and {bookmark-end id:N}', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{bookmark-start id:1 name:"ref"}bookmarked text{bookmark-end id:1}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const bStart = para.runs.find(r => r.type === 'bookmark-start');
    assert.ok(bStart, 'should have bookmark-start node');
    assert.equal(bStart.id, 1);
    assert.equal(bStart.name, 'ref');
    const bEnd = para.runs.find(r => r.type === 'bookmark-end');
    assert.ok(bEnd, 'should have bookmark-end node');
    assert.equal(bEnd.id, 1);
  });
});

describe('DexParser — Hyperlinks', () => {
  it('parses {link rId:rId5}text{/link} (external hyperlink)', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{link rId:rId5}click here{/link}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const linkRun = para.runs.find(r => r.type === 'link');
    assert.ok(linkRun, 'should have link node');
    assert.equal(linkRun.rId, 'rId5');
    assert.equal(linkRun.text, 'click here');
  });

  it('parses {link anchor:"_Ref1"}text{/link} (internal anchor link)', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{link anchor:"_Ref1"}see above{/link}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const linkRun = para.runs.find(r => r.type === 'link');
    assert.ok(linkRun, 'should have link node');
    assert.equal(linkRun.anchor, '_Ref1');
    assert.equal(linkRun.text, 'see above');
  });
});

describe('DexParser — Line breaks', () => {
  it('parses {br} as a linebreak node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\nfirst line{br}second line\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const brNode = para.runs.find(r => r.type === 'linebreak');
    assert.ok(brNode, 'should have linebreak node');
  });

  it('parses {colbreak} as a colbreak node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\nfirst column{colbreak}second column\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const colNode = para.runs.find(r => r.type === 'colbreak');
    assert.ok(colNode, 'should have colbreak node');
  });
});

describe('DexParser — Endnote references', () => {
  it('parses {endnote id:1} as an endnote node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\nSome text{endnote id:1} continues.\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const enNode = para.runs.find(r => r.type === 'endnote');
    assert.ok(enNode, 'should have endnote node');
    assert.equal(enNode.id, 1);
  });
});

describe('DexParser — Symbols', () => {
  it('parses {sym F0B7} as a sym node with char code', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\nBullet: {sym F0B7}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const symNode = para.runs.find(r => r.type === 'sym');
    assert.ok(symNode, 'should have sym node');
    assert.equal(symNode.char, 'F0B7');
  });
});

describe('DexParser — Strikethrough', () => {
  it('parses {strike}text{/strike} as a strike node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{strike}removed text{/strike}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const strikeNode = para.runs.find(r => r.type === 'strike');
    assert.ok(strikeNode, 'should have strike node');
    assert.equal(strikeNode.text, 'removed text');
  });

  it('parses {dstrike}text{/dstrike} as a dstrike node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{dstrike}double struck{/dstrike}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const dstrikeNode = para.runs.find(r => r.type === 'dstrike');
    assert.ok(dstrikeNode, 'should have dstrike node');
    assert.equal(dstrikeNode.text, 'double struck');
  });
});

describe('DexParser — Font size', () => {
  it('parses {size 28}text{/size} as a size node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{size 28}large text{/size}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const sizeNode = para.runs.find(r => r.type === 'size');
    assert.ok(sizeNode, 'should have size node');
    assert.equal(sizeNode.size, '28');
    assert.equal(sizeNode.text, 'large text');
  });
});

describe('DexParser — Small caps / All caps', () => {
  it('parses {smallcaps}text{/smallcaps} as a smallcaps node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{smallcaps}small caps text{/smallcaps}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const scNode = para.runs.find(r => r.type === 'smallcaps');
    assert.ok(scNode, 'should have smallcaps node');
    assert.equal(scNode.text, 'small caps text');
  });

  it('parses {caps}text{/caps} as a caps node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{caps}all caps text{/caps}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const capsNode = para.runs.find(r => r.type === 'caps');
    assert.ok(capsNode, 'should have caps node');
    assert.equal(capsNode.text, 'all caps text');
  });
});

describe('DexParser — Hidden text', () => {
  it('parses {hidden}text{/hidden} as a hidden node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{hidden}secret text{/hidden}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const hiddenNode = para.runs.find(r => r.type === 'hidden');
    assert.ok(hiddenNode, 'should have hidden node');
    assert.equal(hiddenNode.text, 'secret text');
  });
});

describe('DexParser — Underline types', () => {
  it('parses {u double}text{/u} as an underline node with type', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{u double}double underlined{/u}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const uNode = para.runs.find(r => r.type === 'underline');
    assert.ok(uNode, 'should have underline node');
    assert.equal(uNode.underlineType, 'double');
    assert.equal(uNode.text, 'double underlined');
  });
});

describe('DexParser — Paragraph properties', () => {
  it('parses {p align:center style:"Quote" list-id:1 list-level:0 bidi keepnext}', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:CC align:center style:"Quote" list-id:1 list-level:0 bidi keepnext}\nQuoted text.\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    assert.ok(para, 'should have a paragraph');
    assert.equal(para.id, 'CC');
    assert.equal(para.align, 'center');
    assert.equal(para.style, 'Quote');
    assert.equal(para.listId, '1');
    assert.equal(para.listLevel, '0');
    assert.equal(para.bidi, true);
    assert.equal(para.keepnext, true);
  });
});

describe('DexParser — Move tracking', () => {
  it('parses {movefrom id:1 by:"Author" date:"2026-01-01"}text{/movefrom}', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{movefrom id:1 by:"Author" date:"2026-01-01"}moved text{/movefrom}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const mfNode = para.runs.find(r => r.type === 'movefrom');
    assert.ok(mfNode, 'should have movefrom node');
    assert.equal(mfNode.id, 1);
    assert.equal(mfNode.author, 'Author');
    assert.equal(mfNode.date, '2026-01-01');
    assert.equal(mfNode.text, 'moved text');
  });
});

describe('DexParser — Field codes', () => {
  it('parses {field "PAGE"}3{/field} as a field node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\nPage {field "PAGE"}3{/field} of 10\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const fieldNode = para.runs.find(r => r.type === 'field');
    assert.ok(fieldNode, 'should have field node');
    assert.equal(fieldNode.text, '3');
    assert.equal(typeof fieldNode.instr, 'string', 'instr should be a string');
  });

  it('parses {field instr:"TOC"}table of contents{/field} with key:value attr', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{field instr:"TOC"}table of contents{/field}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const fieldNode = para.runs.find(r => r.type === 'field');
    assert.ok(fieldNode, 'should have field node');
    assert.equal(fieldNode.instr, 'TOC');
    assert.equal(fieldNode.text, 'table of contents');
  });
});

describe('DexParser — Content controls (SDT)', () => {
  it('parses {sdt "Bibliography"}text{/sdt} as an sdt node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{sdt "Bibliography"}References here{/sdt}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const sdtNode = para.runs.find(r => r.type === 'sdt');
    assert.ok(sdtNode, 'should have sdt node');
    assert.equal(sdtNode.text, 'References here');
    assert.equal(typeof sdtNode.name, 'string', 'name should be a string');
  });

  it('parses {sdt name:"Bibliography"}text{/sdt} with key:value attr', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{sdt name:"Bibliography"}References here{/sdt}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const sdtNode = para.runs.find(r => r.type === 'sdt');
    assert.ok(sdtNode, 'should have sdt node');
    assert.equal(sdtNode.name, 'Bibliography');
    assert.equal(sdtNode.text, 'References here');
  });
});

describe('DexParser — Math', () => {
  it('parses {math data:AAAA}x+y{/math} as a math node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{math data:AAAA}x+y{/math}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const mathNode = para.runs.find(r => r.type === 'math');
    assert.ok(mathNode, 'should have math node');
    assert.equal(mathNode.data, 'AAAA');
    assert.equal(mathNode.text, 'x+y');
  });
});

describe('DexParser — Text boxes', () => {
  it('parses {textbox}content{/textbox} as a textbox node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{textbox}box content{/textbox}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const tbNode = para.runs.find(r => r.type === 'textbox');
    assert.ok(tbNode, 'should have textbox node');
    assert.equal(tbNode.text, 'box content');
  });
});

describe('DexParser — Ruby text', () => {
  it('parses {ruby base:"text"}annotation{/ruby} as a ruby node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{ruby base:"text"}annotation{/ruby}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const rubyNode = para.runs.find(r => r.type === 'ruby');
    assert.ok(rubyNode, 'should have ruby node');
    assert.equal(rubyNode.base, 'text');
    assert.equal(rubyNode.text, 'annotation');
  });
});

describe('DexParser — Embedded objects', () => {
  it('parses {object type:"Excel.Sheet"} as an object node', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{object type:"Excel.Sheet"}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const para = ast.body.find(n => n.type === 'paragraph');
    const objNode = para.runs.find(r => r.type === 'object');
    assert.ok(objNode, 'should have object node');
    assert.equal(objNode.progId, 'Excel.Sheet');
  });
});

// ============================================================================
// 3. COMPILER TESTS
// ============================================================================

describe('DexCompiler', () => {
  it('compiles AST to valid .docx zip', () => {
    const dex = '---\ndocex: "0.4.0"\ntitle: "Test"\n---\n\n# Introduction {id:3A2F001B}\n\n{p id:3A2F001C}\nThis is a test paragraph.\n{/p}\n';
    const ast = DexParser.parse(dex);
    const outPath = path.join(OUTPUT_DIR, 'dex-compiled.docx');
    const result = DexCompiler.compile(ast, { output: outPath });
    assert.ok(fs.existsSync(result.path), 'output file should exist');
    assert.ok(result.paragraphCount >= 2, 'should have at least 2 paragraphs');
    execFileSync('unzip', ['-t', result.path], { stdio: 'pipe' });
  });

  it('compiles tracked changes into correct XML', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:3A2F001C}\nWe collected {del id:5 by:"Fabio Votta" date:"2026-03-27T10:00:00Z"}268,635{/del}{ins id:6 by:"Fabio Votta" date:"2026-03-27T10:00:00Z"}300,000{/ins} advertisements.\n{/p}\n';
    const ast = DexParser.parse(dex);
    const outPath = path.join(OUTPUT_DIR, 'dex-tracked.docx');
    DexCompiler.compile(ast, { output: outPath });
    const tmp = fs.mkdtempSync('/tmp/dex-test-');
    execFileSync('unzip', ['-o', outPath, '-d', tmp], { stdio: 'pipe' });
    const docXml = fs.readFileSync(path.join(tmp, 'word', 'document.xml'), 'utf-8');
    execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
    assert.ok(docXml.includes('<w:del'), 'should contain w:del');
    assert.ok(docXml.includes('<w:ins'), 'should contain w:ins');
    assert.ok(docXml.includes('Fabio Votta'), 'should contain author');
    assert.ok(docXml.includes('268,635'), 'should contain deleted text');
    assert.ok(docXml.includes('300,000'), 'should contain inserted text');
  });

  it('compiles comments into correct XML', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n# Introduction {id:AA}\n\n{p id:BB}\nSome text here.\n{/p}\n\n{comment id:17 by:"Reviewer 1" date:"2026-03-24T10:00:00Z"}\nAdd citation\n{/comment}\n\n{reply id:18 parent:17 by:"Fabio Votta" date:"2026-03-25T09:00:00Z"}\nDone.\n{/reply}\n';
    const ast = DexParser.parse(dex);
    const outPath = path.join(OUTPUT_DIR, 'dex-comments.docx');
    DexCompiler.compile(ast, { output: outPath });
    const tmp = fs.mkdtempSync('/tmp/dex-test-');
    execFileSync('unzip', ['-o', outPath, '-d', tmp], { stdio: 'pipe' });
    const commentsXml = fs.readFileSync(path.join(tmp, 'word', 'comments.xml'), 'utf-8');
    const commentsExtXml = fs.readFileSync(path.join(tmp, 'word', 'commentsExtended.xml'), 'utf-8');
    execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
    assert.ok(commentsXml.includes('Reviewer 1'), 'comments.xml should contain author');
    assert.ok(commentsXml.includes('Add citation'), 'comments.xml should contain text');
    assert.ok(commentsXml.includes('w:id="17"'), 'comments.xml should preserve id');
    assert.ok(commentsXml.includes('w:id="18"'), 'comments.xml should have reply id');
    assert.ok(commentsExtXml.includes('w15:commentEx'), 'commentsExtended.xml should have entries');
    assert.ok(commentsExtXml.includes('w15:paraIdParent'), 'commentsExtended.xml should have threading');
  });

  it('compiles footnotes correctly', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\nThis is text{footnote id:2}Footnote content.{/footnote}.\n{/p}\n';
    const ast = DexParser.parse(dex);
    const outPath = path.join(OUTPUT_DIR, 'dex-footnotes.docx');
    DexCompiler.compile(ast, { output: outPath });
    const tmp = fs.mkdtempSync('/tmp/dex-test-');
    execFileSync('unzip', ['-o', outPath, '-d', tmp], { stdio: 'pipe' });
    const footnotesXml = fs.readFileSync(path.join(tmp, 'word', 'footnotes.xml'), 'utf-8');
    const docXml = fs.readFileSync(path.join(tmp, 'word', 'document.xml'), 'utf-8');
    execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
    assert.ok(footnotesXml.includes('Footnote content'), 'footnotes.xml should contain text');
    assert.ok(footnotesXml.includes('w:id="2"'), 'footnotes.xml should preserve id');
    assert.ok(docXml.includes('w:footnoteReference'), 'document.xml should have footnote ref');
  });

  it('compiles tables correctly', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{table style:booktabs cols:2}\n| Name | Value |\n|---|---|\n| A | 1 |\n{/table}\n';
    const ast = DexParser.parse(dex);
    const outPath = path.join(OUTPUT_DIR, 'dex-tables.docx');
    DexCompiler.compile(ast, { output: outPath });
    const tmp = fs.mkdtempSync('/tmp/dex-test-');
    execFileSync('unzip', ['-o', outPath, '-d', tmp], { stdio: 'pipe' });
    const docXml = fs.readFileSync(path.join(tmp, 'word', 'document.xml'), 'utf-8');
    execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
    assert.ok(docXml.includes('<w:tbl>'), 'should contain table');
    assert.ok(docXml.includes('<w:tr>'), 'should contain table rows');
    assert.ok(docXml.includes('Name'), 'should contain header text');
  });

  it('compiles inline formatting correctly', () => {
    const dex = '---\ndocex: "0.4.0"\n---\n\n{p id:AA}\n{b}bold{/b} and {i}italic{/i} and {u}underlined{/u}\n{/p}\n';
    const ast = DexParser.parse(dex);
    const outPath = path.join(OUTPUT_DIR, 'dex-formatting.docx');
    DexCompiler.compile(ast, { output: outPath });
    const tmp = fs.mkdtempSync('/tmp/dex-test-');
    execFileSync('unzip', ['-o', outPath, '-d', tmp], { stdio: 'pipe' });
    const docXml = fs.readFileSync(path.join(tmp, 'word', 'document.xml'), 'utf-8');
    execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
    assert.ok(docXml.includes('<w:b/>'), 'should contain bold formatting');
    assert.ok(docXml.includes('<w:i/>'), 'should contain italic formatting');
    assert.ok(docXml.includes('w:u w:val="single"'), 'should contain underline formatting');
  });
});

// ============================================================================
// 4. ROUND-TRIP TESTS
// ============================================================================

describe('Round-trip', () => {
  it('decompile -> parse -> compile -> decompile yields same structure', () => {
    const ws1 = Workspace.open(FIXTURE);
    const dex1 = DexDecompiler.decompile(ws1);
    ws1.cleanup();
    const ast = DexParser.parse(dex1);
    assert.ok(ast.body.length > 0, 'AST should have body nodes');
    const outPath = path.join(OUTPUT_DIR, 'dex-roundtrip.docx');
    DexCompiler.compile(ast, { output: outPath });
    const ws2 = Workspace.open(outPath);
    const dex2 = DexDecompiler.decompile(ws2);
    ws2.cleanup();
    const headings1 = dex1.match(/^#{1,6}\s+.*/gm) || [];
    const headings2 = dex2.match(/^#{1,6}\s+.*/gm) || [];
    assert.equal(headings1.length, headings2.length, 'heading count should match');
    const paras1 = (dex1.match(/\{p\b/g) || []).length;
    const paras2 = (dex2.match(/\{p\b/g) || []).length;
    assert.equal(paras1, paras2, 'paragraph count should match');
  });

  it('round-trip preserves heading text', () => {
    const ws1 = Workspace.open(FIXTURE);
    const dex1 = DexDecompiler.decompile(ws1);
    ws1.cleanup();
    const ast = DexParser.parse(dex1);
    const outPath = path.join(OUTPUT_DIR, 'dex-roundtrip-headings.docx');
    DexCompiler.compile(ast, { output: outPath });
    const ws2 = Workspace.open(outPath);
    const dex2 = DexDecompiler.decompile(ws2);
    ws2.cleanup();
    assert.ok(dex2.includes('Introduction'), 'should preserve Introduction heading');
    assert.ok(dex2.includes('Methods'), 'should preserve Methods heading');
    assert.ok(dex2.includes('Results'), 'should preserve Results heading');
  });

  it('round-trip preserves tracked changes', () => {
    const sourcePath = createDocxWithTrackedChanges('roundtrip-tracked');
    const ws1 = Workspace.open(sourcePath);
    const dex1 = DexDecompiler.decompile(ws1);
    ws1.cleanup();
    assert.ok(dex1.includes('{del'), 'first decompile should have del');
    assert.ok(dex1.includes('{ins'), 'first decompile should have ins');
    const ast = DexParser.parse(dex1);
    const outPath = path.join(OUTPUT_DIR, 'dex-roundtrip-tracked.docx');
    DexCompiler.compile(ast, { output: outPath });
    const ws2 = Workspace.open(outPath);
    const dex2 = DexDecompiler.decompile(ws2);
    ws2.cleanup();
    assert.ok(dex2.includes('{del'), 'second decompile should have del');
    assert.ok(dex2.includes('{ins'), 'second decompile should have ins');
    assert.ok(dex2.includes('268,635'), 'should preserve deleted text');
    assert.ok(dex2.includes('300,000'), 'should preserve inserted text');
    assert.ok(dex2.includes('Fabio Votta'), 'should preserve author');
  });

  it('round-trip on fixture preserves paragraph count', () => {
    const ws1 = Workspace.open(FIXTURE);
    const dex1 = DexDecompiler.decompile(ws1);
    ws1.cleanup();
    const ast = DexParser.parse(dex1);
    const outPath = path.join(OUTPUT_DIR, 'dex-roundtrip-paracount.docx');
    const result = DexCompiler.compile(ast, { output: outPath });
    const bodyParas = ast.body.filter(n => n.type === 'paragraph' || n.type === 'heading');
    assert.ok(bodyParas.length > 0, 'should have body paragraphs');
    assert.ok(result.paragraphCount >= bodyParas.length,
      'compiled should have >= ' + bodyParas.length + ' paragraphs (got ' + result.paragraphCount + ')');
  });
});

// ============================================================================
// 5. DIRECT API INTEGRATION TESTS
// ============================================================================

describe('Direct API integration', () => {
  it('decompileToDex works (via DexDecompiler)', () => {
    const ws = Workspace.open(FIXTURE);
    const dex = DexDecompiler.decompile(ws);
    ws.cleanup();
    assert.ok(typeof dex === 'string', 'should return string');
    assert.ok(dex.startsWith('---'), 'should start with frontmatter');
    assert.ok(dex.includes('Introduction'), 'should contain heading text');
  });

  it('buildFromDex works (via DexParser + DexCompiler)', () => {
    const ws = Workspace.open(FIXTURE);
    const dex = DexDecompiler.decompile(ws);
    ws.cleanup();
    const dexPath = path.join(OUTPUT_DIR, 'api-test.dex');
    fs.writeFileSync(dexPath, dex, 'utf-8');
    const dexString = fs.readFileSync(dexPath, 'utf-8');
    const ast = DexParser.parse(dexString);
    const outPath = path.join(OUTPUT_DIR, 'api-test-built.docx');
    const result = DexCompiler.compile(ast, { output: outPath });
    assert.ok(fs.existsSync(result.path), 'output should exist');
    assert.ok(result.paragraphCount > 0, 'should have paragraphs');
  });

  it('full roundtrip works', () => {
    const ws1 = Workspace.open(FIXTURE);
    const dex1 = DexDecompiler.decompile(ws1);
    ws1.cleanup();
    const ast = DexParser.parse(dex1);
    const tmpDocx = path.join(OUTPUT_DIR, 'api-roundtrip.docx');
    DexCompiler.compile(ast, { output: tmpDocx });
    const ws2 = Workspace.open(tmpDocx);
    const dex2 = DexDecompiler.decompile(ws2);
    ws2.cleanup();
    assert.ok(typeof dex1 === 'string', 'should have dex1');
    assert.ok(typeof dex2 === 'string', 'should have dex2');
    const h1 = dex1.match(/^#{1,6}\s+.*/gm) || [];
    const h2 = dex2.match(/^#{1,6}\s+.*/gm) || [];
    assert.equal(h1.length, h2.length, 'heading count should match');
  });
});
