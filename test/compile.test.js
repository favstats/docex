/**
 * compile.test.js -- Tests for v0.4.0 features
 *
 * Tests:
 *   1. Compile from .tex produces valid .docx (requires pandoc)
 *   2. Compile with style applies formatting
 *   3. Decompile preserves tracked changes as comments
 *   4. Batch operations apply to all documents
 *   5. fromTemplate creates proper document structure
 *   6. responseLetter generates formatted response
 *   7. create() produces minimal valid .docx
 *   8. toHtml produces HTML
 *   9. toMarkdown produces Markdown
 *  10. Factory methods on docex function
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Tests requiring pandoc skip gracefully if pandoc is not installed.
 */

const { describe, it, before, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const FIXTURE = path.join(__dirname, 'fixtures', 'test-manuscript.docx');
const OUTPUT_DIR = path.join(__dirname, 'output');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

/** Helper: create a fresh copy of the fixture for each test */
function freshCopy(testName) {
  const out = path.join(OUTPUT_DIR, `${testName}.docx`);
  fs.copyFileSync(FIXTURE, out);
  return out;
}

/** Check if pandoc is available */
function hasPandoc() {
  try {
    execFileSync('which', ['pandoc'], { stdio: 'pipe', encoding: 'utf-8' });
    return true;
  } catch {
    return false;
  }
}

/** Helper: unzip a docx and read a specific XML file */
function readDocxXml(docxPath, xmlFile) {
  const tmp = fs.mkdtempSync('/tmp/docex-test-');
  execFileSync('unzip', ['-o', docxPath, '-d', tmp], { stdio: 'pipe' });
  const content = fs.readFileSync(path.join(tmp, xmlFile), 'utf8');
  execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
  return content;
}

/** Helper: create a minimal .tex file for testing */
function createTexFile(name, content) {
  const texPath = path.join(OUTPUT_DIR, name);
  fs.writeFileSync(texPath, content, 'utf-8');
  return texPath;
}

// ============================================================================
// 1. COMPILE FROM LATEX
// ============================================================================

describe('Compile.fromLatex', () => {
  const pandocAvailable = hasPandoc();

  it('compiles a simple .tex to valid .docx', { skip: !pandocAvailable && 'pandoc not installed' }, async () => {
    const texPath = createTexFile('compile-basic.tex', `
\\documentclass{article}
\\begin{document}
\\section{Introduction}
This is a test document for docex compile.
\\section{Methods}
We used a variety of approaches.
\\end{document}
`);
    const outputPath = path.join(OUTPUT_DIR, 'compile-basic.docx');

    const { Compile } = require('../src/compile');
    const result = await Compile.fromLatex(texPath, { output: outputPath });

    assert.ok(fs.existsSync(result.path), 'Output .docx should exist');
    assert.ok(result.fileSize > 0, 'Output should have non-zero size');
    assert.ok(result.paragraphCount > 0, 'Should have paragraphs');

    // Verify it's a valid zip/docx
    const docXml = readDocxXml(result.path, 'word/document.xml');
    assert.ok(docXml.includes('<w:document'), 'Should contain w:document element');
    assert.ok(docXml.includes('Introduction'), 'Should contain Introduction text');
    assert.ok(docXml.includes('test document'), 'Should contain body text');
  });

  it('applies style preset during compile', { skip: !pandocAvailable && 'pandoc not installed' }, async () => {
    const texPath = createTexFile('compile-styled.tex', `
\\documentclass{article}
\\begin{document}
\\section{Introduction}
A styled test document.
\\end{document}
`);
    const outputPath = path.join(OUTPUT_DIR, 'compile-styled.docx');

    const { Compile } = require('../src/compile');
    const result = await Compile.fromLatex(texPath, {
      output: outputPath,
      style: 'polcomm',
    });

    assert.equal(result.style, 'polcomm');
    assert.ok(fs.existsSync(result.path));

    // Verify style was applied (check styles.xml for Times New Roman)
    const stylesXml = readDocxXml(result.path, 'word/styles.xml');
    assert.ok(stylesXml.includes('Times New Roman') || stylesXml.includes('TimesNewRoman'),
      'Should have Times New Roman font from polcomm preset');
  });

  it('throws if pandoc is not available when called with wrong path', async () => {
    const { Compile } = require('../src/compile');
    await assert.rejects(
      () => Compile.fromLatex('/nonexistent/file.tex'),
      /not found/i,
    );
  });

  it('throws for nonexistent .tex file', async () => {
    const { Compile } = require('../src/compile');
    await assert.rejects(
      () => Compile.fromLatex('/tmp/definitely-not-exist.tex'),
      /not found/i,
    );
  });
});

// ============================================================================
// 2. DECOMPILE TO LATEX
// ============================================================================

describe('Compile.decompile', () => {
  it('converts .docx to .tex string', async () => {
    const docxPath = freshCopy('decompile-basic');
    const { Compile } = require('../src/compile');

    const result = await Compile.decompile(docxPath);

    assert.ok(result.tex.length > 0, 'Should produce LaTeX output');
    assert.ok(result.tex.includes('\\begin{document}'), 'Should have \\begin{document}');
    assert.ok(result.tex.includes('\\end{document}'), 'Should have \\end{document}');
    assert.equal(typeof result.changes, 'number');
    assert.equal(typeof result.comments, 'number');
  });

  it('writes to file when output is specified', async () => {
    const docxPath = freshCopy('decompile-file');
    const outputPath = path.join(OUTPUT_DIR, 'decompile-output.tex');
    const { Compile } = require('../src/compile');

    const result = await Compile.decompile(docxPath, { output: outputPath });

    assert.ok(fs.existsSync(outputPath), '.tex output file should exist');
    assert.equal(result.path, outputPath);
    const content = fs.readFileSync(outputPath, 'utf-8');
    assert.ok(content.includes('\\begin{document}'));
  });

  it('toLatex without preserve flags skips change/comment injection', async () => {
    const docxPath = freshCopy('tolatex-nopreserve');
    const { Compile } = require('../src/compile');

    const result = await Compile.toLatex(docxPath, {
      preserveChanges: false,
      preserveComments: false,
    });

    assert.ok(result.tex.includes('\\begin{document}'));
    assert.equal(result.changes, 0, 'Should not report changes when not preserving');
    assert.equal(result.comments, 0, 'Should not report comments when not preserving');
  });
});

// ============================================================================
// 3. BATCH OPERATIONS
// ============================================================================

describe('Batch', () => {
  it('creates a batch from file paths', () => {
    const { Batch } = require('../src/batch');
    const paths = [freshCopy('batch-1'), freshCopy('batch-2')];
    const batch = new Batch(paths);

    assert.equal(batch.length, 2);
    assert.equal(batch.paths.length, 2);
  });

  it('throws for empty paths array', () => {
    const { Batch } = require('../src/batch');
    assert.throws(() => new Batch([]), /non-empty/);
  });

  it('chains author and style', () => {
    const { Batch } = require('../src/batch');
    const paths = [freshCopy('batch-chain-1')];
    const batch = new Batch(paths);
    const result = batch.author('Test').style('academic');
    assert.equal(result, batch, 'Should return self for chaining');
  });

  it('applies operations to all documents and saves', async () => {
    const paths = [freshCopy('batch-save-1'), freshCopy('batch-save-2')];
    const { Batch } = require('../src/batch');
    const batch = new Batch(paths);
    batch.author('Batch Test');

    const results = await batch.saveAll();

    assert.equal(results.length, 2);
    for (const r of results) {
      assert.ok(r.path, 'Each result should have a path');
      assert.ok(r.fileSize > 0, 'Each result should have non-zero size');
      assert.equal(r.error, null, 'Should not have errors');
    }
  });

  it('saves with suffix option', async () => {
    const paths = [freshCopy('batch-suffix-1')];
    const { Batch } = require('../src/batch');
    const batch = new Batch(paths);

    const results = await batch.saveAll({ suffix: '_formatted' });

    assert.equal(results.length, 1);
    assert.ok(results[0].path.includes('_formatted'), 'Output should have suffix');
    assert.ok(fs.existsSync(results[0].path), 'Output file should exist');

    // Clean up
    try { fs.unlinkSync(results[0].path); } catch (_) {}
  });

  it('verify returns per-file results', async () => {
    const paths = [freshCopy('batch-verify-1'), freshCopy('batch-verify-2')];
    const { Batch } = require('../src/batch');
    const batch = new Batch(paths);

    const results = await batch.verify('academic');

    assert.equal(results.length, 2);
    for (const r of results) {
      assert.ok(r.path, 'Each result should have a path');
      assert.ok('result' in r, 'Each result should have a result object');
      assert.ok('pass' in r.result, 'Result should have pass property');
    }
  });

  it('replaceAll queues operation for all documents', async () => {
    const paths = [freshCopy('batch-replace-1')];
    const { Batch } = require('../src/batch');
    const batch = new Batch(paths);
    batch.author('Test');
    batch.replaceAll('the', 'THE');

    const results = await batch.saveAll();
    assert.equal(results.length, 1);
    assert.equal(results[0].error, null);
  });
});

// ============================================================================
// 4. BATCH VIA FACTORY
// ============================================================================

describe('docex.batch() factory', () => {
  it('creates a Batch instance', () => {
    const docex = require('../src/docex');
    const paths = [freshCopy('factory-batch-1')];
    const batch = docex.batch(paths);
    assert.equal(batch.length, 1);
  });
});

// ============================================================================
// 5. TEMPLATE
// ============================================================================

describe('Template.create', () => {
  it('creates a .docx from a preset with metadata', async () => {
    const outputPath = path.join(OUTPUT_DIR, 'template-polcomm.docx');
    const { Template } = require('../src/template');

    const result = await Template.create('polcomm', {
      title: 'The Ban That Wasn\'t',
      authors: [
        { name: 'Fabio Votta', affiliation: 'UvA', email: 'f.votta@uva.nl' },
        { name: 'Simon Munzert', affiliation: 'JGU Mainz' },
      ],
      abstract: 'This paper examines political advertising enforcement failures.',
      keywords: ['political advertising', 'platform governance'],
      output: outputPath,
    });

    assert.ok(fs.existsSync(result.path), 'Template .docx should exist');
    assert.ok(result.fileSize > 0);
    assert.ok(result.paragraphCount > 5, 'Should have multiple paragraphs');

    // Verify contents
    const docXml = readDocxXml(result.path, 'word/document.xml');
    assert.ok(docXml.includes('Ban'), 'Should contain title');
    assert.ok(docXml.includes('Fabio Votta'), 'Should contain author name');
    assert.ok(docXml.includes('Abstract'), 'Should contain Abstract heading');
    assert.ok(docXml.includes('Introduction'), 'Should contain Introduction heading');
    assert.ok(docXml.includes('Methods'), 'Should contain Methods heading');
    assert.ok(docXml.includes('References'), 'Should contain References heading');
    assert.ok(docXml.includes('political advertising'), 'Should contain keywords');
  });

  it('creates template with minimal metadata', async () => {
    const outputPath = path.join(OUTPUT_DIR, 'template-minimal.docx');
    const { Template } = require('../src/template');

    const result = await Template.create('academic', {
      title: 'Minimal Paper',
      output: outputPath,
    });

    assert.ok(fs.existsSync(result.path));
    assert.ok(result.paragraphCount > 0);
  });

  it('throws for unknown preset', async () => {
    const { Template } = require('../src/template');
    await assert.rejects(
      () => Template.create('nonexistent-journal', {}),
      /Unknown preset/,
    );
  });
});

// ============================================================================
// 6. FROM TEMPLATE FACTORY
// ============================================================================

describe('docex.fromTemplate() factory', () => {
  it('creates a document from template via factory method', async () => {
    const outputPath = path.join(OUTPUT_DIR, 'factory-template.docx');
    const docex = require('../src/docex');

    const result = await docex.fromTemplate('apa7', {
      title: 'Factory Test Paper',
      authors: ['Test Author'],
      output: outputPath,
    });

    assert.ok(fs.existsSync(result.path));
    assert.ok(result.fileSize > 0);
  });
});

// ============================================================================
// 7. RESPONSE LETTER
// ============================================================================

describe('ResponseLetter.generate', () => {
  it('generates a response letter from comments and responses', async () => {
    const outputPath = path.join(OUTPUT_DIR, 'response-letter.docx');
    const { ResponseLetter } = require('../src/response-letter');

    const comments = [
      { id: 1, author: 'Reviewer 1', text: 'The methodology section needs more detail.', date: '2026-01-15' },
      { id: 2, author: 'Reviewer 1', text: 'Please add more citations.', date: '2026-01-15' },
      { id: 3, author: 'Reviewer 2', text: 'The abstract is too long.', date: '2026-01-16' },
    ];

    const responses = {
      1: { action: 'agree', text: 'We expanded the methodology section significantly.', changes: ['Added 2 paragraphs to Methods'] },
      2: { action: 'partial', text: 'We added several citations but maintained our framing.', changes: ['Added Smith 2024', 'Added Jones 2023'] },
      3: { action: 'disagree', text: 'The abstract is within the word limit at 248 words.', changes: [] },
    };

    const result = await ResponseLetter.generate(comments, responses, {
      title: 'The Ban That Wasn\'t',
      journal: 'Political Communication',
      authors: ['Fabio Votta'],
      output: outputPath,
    });

    assert.ok(fs.existsSync(result.path), 'Response letter should exist');
    assert.ok(result.fileSize > 0);
    assert.equal(result.reviewers, 2, 'Should have 2 reviewers');
    assert.equal(result.commentsAddressed, 3, 'Should address all 3 comments');

    // Verify contents
    const docXml = readDocxXml(result.path, 'word/document.xml');
    assert.ok(docXml.includes('Response to Reviewer Comments'), 'Should have header');
    assert.ok(docXml.includes('Reviewer 1'), 'Should mention Reviewer 1');
    assert.ok(docXml.includes('Reviewer 2'), 'Should mention Reviewer 2');
    assert.ok(docXml.includes('methodology'), 'Should include comment text');
    assert.ok(docXml.includes('expanded'), 'Should include response text');
  });

  it('handles empty responses gracefully', async () => {
    const outputPath = path.join(OUTPUT_DIR, 'response-letter-empty.docx');
    const { ResponseLetter } = require('../src/response-letter');

    const comments = [
      { id: 1, author: 'Reviewer 1', text: 'Fix this.', date: '2026-01-15' },
    ];

    const result = await ResponseLetter.generate(comments, {}, {
      output: outputPath,
    });

    assert.ok(fs.existsSync(result.path));
    assert.equal(result.commentsAddressed, 0);

    const docXml = readDocxXml(result.path, 'word/document.xml');
    assert.ok(docXml.includes('No response provided'), 'Should indicate missing response');
  });

  it('throws for invalid arguments', async () => {
    const { ResponseLetter } = require('../src/response-letter');
    await assert.rejects(
      () => ResponseLetter.generate('not-an-array', {}, {}),
      /array/,
    );
    await assert.rejects(
      () => ResponseLetter.generate([], null, {}),
      /object/,
    );
  });
});

// ============================================================================
// 8. RESPONSE LETTER FACTORY
// ============================================================================

describe('docex.responseLetter() factory', () => {
  it('generates via factory method', async () => {
    const outputPath = path.join(OUTPUT_DIR, 'factory-response.docx');
    const docex = require('../src/docex');

    const comments = [
      { id: 1, author: 'R1', text: 'Test comment.', date: '2026-01-01' },
    ];
    const responses = {
      1: { action: 'agree', text: 'Done.', changes: ['Fixed'] },
    };

    const result = await docex.responseLetter(comments, responses, { output: outputPath });
    assert.ok(fs.existsSync(result.path));
  });
});

// ============================================================================
// 9. CREATE EMPTY DOCUMENT
// ============================================================================

describe('createEmpty', () => {
  it('produces a minimal valid .docx', async () => {
    const outputPath = path.join(OUTPUT_DIR, 'create-empty.docx');
    const { createEmpty } = require('../src/template');

    const result = await createEmpty({ output: outputPath });

    assert.ok(fs.existsSync(result.path), 'Should create .docx file');
    assert.ok(result.fileSize > 0, 'Should have non-zero size');
    assert.equal(result.paragraphCount, 1, 'Should have one paragraph');

    // Verify it's a valid docx by opening with workspace
    const { Workspace } = require('../src/workspace');
    const ws = Workspace.open(result.path);
    assert.ok(ws.docXml.includes('<w:document'), 'Should be valid OOXML');
    ws.cleanup();
  });

  it('creates at temp path when no output specified', async () => {
    const { createEmpty } = require('../src/template');

    const result = await createEmpty();

    assert.ok(fs.existsSync(result.path), 'Should create file at temp path');
    assert.ok(result.path.includes('docex-create'), 'Should use docex temp pattern');

    // Clean up
    try { fs.unlinkSync(result.path); } catch (_) {}
  });
});

// ============================================================================
// 10. docex.create() FACTORY
// ============================================================================

describe('docex.create() factory', () => {
  it('creates a minimal .docx via factory method', async () => {
    const outputPath = path.join(OUTPUT_DIR, 'factory-create.docx');
    const docex = require('../src/docex');

    const result = await docex.create({ output: outputPath });
    assert.ok(fs.existsSync(result.path));
    assert.ok(result.fileSize > 0);
  });
});

// ============================================================================
// 11. TO HTML
// ============================================================================

describe('toHtml', () => {
  const pandocAvailable = hasPandoc();

  it('converts .docx to HTML string', { skip: !pandocAvailable && 'pandoc not installed' }, async () => {
    const docxPath = freshCopy('tohtml-basic');
    const docex = require('../src/docex');
    const doc = docex(docxPath);

    const html = await doc.toHtml();

    assert.ok(html.length > 0, 'Should produce HTML output');
    assert.ok(html.includes('<'), 'Should contain HTML tags');
    doc.discard();
  });

  it('writes HTML to file when output specified', { skip: !pandocAvailable && 'pandoc not installed' }, async () => {
    const docxPath = freshCopy('tohtml-file');
    const outputPath = path.join(OUTPUT_DIR, 'tohtml-output.html');
    const docex = require('../src/docex');
    const doc = docex(docxPath);

    const html = await doc.toHtml({ output: outputPath });

    assert.ok(fs.existsSync(outputPath), 'HTML file should exist');
    const content = fs.readFileSync(outputPath, 'utf-8');
    assert.ok(content.includes('<'), 'File should contain HTML');
    doc.discard();
  });
});

// ============================================================================
// 12. TO MARKDOWN
// ============================================================================

describe('toMarkdown', () => {
  const pandocAvailable = hasPandoc();

  it('converts .docx to Markdown string', { skip: !pandocAvailable && 'pandoc not installed' }, async () => {
    const docxPath = freshCopy('tomd-basic');
    const docex = require('../src/docex');
    const doc = docex(docxPath);

    const md = await doc.toMarkdown();

    assert.ok(md.length > 0, 'Should produce Markdown output');
    doc.discard();
  });

  it('writes Markdown to file when output specified', { skip: !pandocAvailable && 'pandoc not installed' }, async () => {
    const docxPath = freshCopy('tomd-file');
    const outputPath = path.join(OUTPUT_DIR, 'tomd-output.md');
    const docex = require('../src/docex');
    const doc = docex(docxPath);

    const md = await doc.toMarkdown({ output: outputPath });

    assert.ok(fs.existsSync(outputPath), 'Markdown file should exist');
    doc.discard();
  });
});

// ============================================================================
// 13. WATCH MODE (just test creation/teardown, not file watching)
// ============================================================================

describe('Compile.watch', () => {
  const pandocAvailable = hasPandoc();

  it('creates a watcher that can be closed', { skip: !pandocAvailable && 'pandoc not installed' }, async () => {
    const texPath = createTexFile('watch-test.tex', `
\\documentclass{article}
\\begin{document}
Test
\\end{document}
`);
    const outputPath = path.join(OUTPUT_DIR, 'watch-output.docx');
    const { Compile } = require('../src/compile');

    const watcher = Compile.watch(texPath, { output: outputPath });

    assert.ok(watcher, 'Should return a watcher object');
    assert.ok(typeof watcher.close === 'function', 'Watcher should have close()');

    // Wait a bit for initial compile
    await new Promise(resolve => setTimeout(resolve, 3000));

    watcher.close();

    assert.ok(fs.existsSync(outputPath), 'Should have compiled output');
  });

  it('throws for nonexistent .tex file', () => {
    const { Compile } = require('../src/compile');
    assert.throws(
      () => Compile.watch('/tmp/nonexistent.tex'),
      /not found/i,
    );
  });
});

// ============================================================================
// 14. COMPILE/DECOMPILE FACTORIES
// ============================================================================

describe('docex.compile and docex.decompile factories', () => {
  const pandocAvailable = hasPandoc();

  it('docex.compile is a function', () => {
    const docex = require('../src/docex');
    assert.equal(typeof docex.compile, 'function');
  });

  it('docex.decompile is a function', () => {
    const docex = require('../src/docex');
    assert.equal(typeof docex.decompile, 'function');
  });

  it('docex.watch is a function', () => {
    const docex = require('../src/docex');
    assert.equal(typeof docex.watch, 'function');
  });

  it('round-trip: .tex -> .docx -> .tex', { skip: !pandocAvailable && 'pandoc not installed' }, async () => {
    const docex = require('../src/docex');

    // Create a .tex file
    const texPath = createTexFile('roundtrip.tex', `
\\documentclass{article}
\\begin{document}
\\section{Introduction}
This is a round-trip test.
\\section{Methods}
Testing the pipeline.
\\end{document}
`);
    const docxPath = path.join(OUTPUT_DIR, 'roundtrip.docx');
    const texOutPath = path.join(OUTPUT_DIR, 'roundtrip-out.tex');

    // Compile .tex -> .docx
    const compileResult = await docex.compile(texPath, { output: docxPath });
    assert.ok(fs.existsSync(compileResult.path));

    // Decompile .docx -> .tex
    const decompileResult = await docex.decompile(docxPath, { output: texOutPath });
    assert.ok(fs.existsSync(decompileResult.path));
    assert.ok(decompileResult.tex.includes('\\begin{document}'));
  });
});

// ============================================================================
// 15. EXPORTS CHECK
// ============================================================================

describe('module exports', () => {
  it('exports Compile class', () => {
    const docex = require('../src/docex');
    assert.ok(docex.Compile);
  });

  it('exports Batch class', () => {
    const docex = require('../src/docex');
    assert.ok(docex.Batch);
  });

  it('exports Template class', () => {
    const docex = require('../src/docex');
    assert.ok(docex.Template);
  });

  it('exports ResponseLetter class', () => {
    const docex = require('../src/docex');
    assert.ok(docex.ResponseLetter);
  });
});
