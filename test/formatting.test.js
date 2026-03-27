/**
 * formatting.test.js -- Tests for the Formatting module
 *
 * Tests inline formatting operations: bold, italic, underline,
 * highlight, color, code, superscript, subscript, smallCaps,
 * strikethrough. Covers single-run and cross-run text, tracked
 * and untracked changes, and multiple formatting on the same text.
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/formatting.test.js
 */

const { describe, it, before } = require('node:test');
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

/** Helper: unzip a docx and read a specific XML file */
function readDocxXml(docxPath, xmlFile) {
  const tmp = fs.mkdtempSync('/tmp/docex-fmt-test-');
  execFileSync('unzip', ['-o', docxPath, '-d', tmp], { stdio: 'pipe' });
  const content = fs.readFileSync(path.join(tmp, xmlFile), 'utf8');
  execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
  return content;
}

// ============================================================================
// 1. BASIC FORMATTING
// ============================================================================

describe('formatting - bold', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds w:b to rPr for bold text', () => {
    const out = freshCopy('fmt-bold');
    const ws = Workspace.open(out);
    Formatting.bold(ws, '268,635');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    // Find the run containing 268,635 and check for w:b
    const runMatch = docXml.match(/<w:r>[^]*?268,635[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with 268,635');
    assert.ok(runMatch[0].includes('<w:b/>'), 'run should have w:b in rPr');
  });

  it('bold with tracked change wraps in rPrChange', () => {
    const out = freshCopy('fmt-bold-tracked');
    const ws = Workspace.open(out);
    Formatting.bold(ws, '268,635', { tracked: true, author: 'Fabio Votta' });
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    assert.ok(docXml.includes('<w:rPrChange'), 'should have rPrChange');
    assert.ok(docXml.includes('Fabio Votta'), 'should attribute to author');
    assert.ok(docXml.includes('<w:b/>'), 'should have bold element');
  });
});

describe('formatting - italic', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds w:i to rPr for italic text', () => {
    const out = freshCopy('fmt-italic');
    const ws = Workspace.open(out);
    Formatting.italic(ws, 'platform governance');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?platform governance[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('<w:i/>'), 'run should have w:i in rPr');
  });
});

describe('formatting - underline', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds w:u to rPr for underlined text', () => {
    const out = freshCopy('fmt-underline');
    const ws = Workspace.open(out);
    Formatting.underline(ws, 'electoral transparency');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?electoral transparency[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('<w:u w:val="single"/>'), 'run should have w:u');
  });
});

describe('formatting - highlight', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds w:highlight for highlighted text', () => {
    const out = freshCopy('fmt-highlight');
    const ws = Workspace.open(out);
    Formatting.highlight(ws, 'platform governance', 'yellow');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?platform governance[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('<w:highlight w:val="yellow"/>'), 'run should have highlight');
  });

  it('rejects invalid highlight color', () => {
    const out = freshCopy('fmt-highlight-invalid');
    const ws = Workspace.open(out);
    assert.throws(
      () => Formatting.highlight(ws, 'test', 'neon'),
      /Invalid highlight color/
    );
    ws.cleanup();
  });
});

describe('formatting - color', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds w:color with correct hex for named color', () => {
    const out = freshCopy('fmt-color');
    const ws = Workspace.open(out);
    Formatting.color(ws, '268,635', 'red');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?268,635[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('<w:color w:val="FF0000"/>'), 'should have red color');
  });

  it('accepts raw hex color', () => {
    const out = freshCopy('fmt-color-hex');
    const ws = Workspace.open(out);
    Formatting.color(ws, '268,635', '1A2B3C');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    assert.ok(docXml.includes('<w:color w:val="1A2B3C"/>'), 'should have custom hex color');
  });

  it('rejects invalid color value', () => {
    const out = freshCopy('fmt-color-invalid');
    const ws = Workspace.open(out);
    assert.throws(
      () => Formatting.color(ws, 'test', 'not-a-color'),
      /Invalid color/
    );
    ws.cleanup();
  });
});

describe('formatting - code', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds Courier New font for code formatting', () => {
    const out = freshCopy('fmt-code');
    const ws = Workspace.open(out);
    Formatting.code(ws, '268,635');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?268,635[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('Courier New'), 'should have Courier New font');
    assert.ok(runMatch[0].includes('w:ascii="Courier New"'), 'should have ascii attr');
    assert.ok(runMatch[0].includes('w:hAnsi="Courier New"'), 'should have hAnsi attr');
  });
});

describe('formatting - superscript and subscript', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds vertAlign superscript', () => {
    const out = freshCopy('fmt-superscript');
    const ws = Workspace.open(out);
    Formatting.superscript(ws, '268,635');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?268,635[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(
      runMatch[0].includes('<w:vertAlign w:val="superscript"/>'),
      'should have superscript vertAlign'
    );
  });

  it('adds vertAlign subscript', () => {
    const out = freshCopy('fmt-subscript');
    const ws = Workspace.open(out);
    Formatting.subscript(ws, '268,635');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?268,635[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(
      runMatch[0].includes('<w:vertAlign w:val="subscript"/>'),
      'should have subscript vertAlign'
    );
  });
});

describe('formatting - strikethrough', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds w:strike for strikethrough text', () => {
    const out = freshCopy('fmt-strike');
    const ws = Workspace.open(out);
    Formatting.strikethrough(ws, 'platform governance');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?platform governance[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('<w:strike/>'), 'should have strikethrough');
  });
});

describe('formatting - smallCaps', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('adds w:smallCaps for small caps text', () => {
    const out = freshCopy('fmt-smallcaps');
    const ws = Workspace.open(out);
    Formatting.smallCaps(ws, 'platform governance');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?platform governance[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('<w:smallCaps/>'), 'should have smallCaps');
  });
});

// ============================================================================
// 2. MULTIPLE FORMATTING
// ============================================================================

describe('formatting - multiple formats on same text', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('applies bold + italic to the same text', () => {
    const out = freshCopy('fmt-bold-italic');
    const ws = Workspace.open(out);
    Formatting.bold(ws, '268,635');
    Formatting.italic(ws, '268,635');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?268,635[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('<w:b/>'), 'should have bold');
    assert.ok(runMatch[0].includes('<w:i/>'), 'should have italic');
  });

  it('applies bold + color + highlight to the same text', () => {
    const out = freshCopy('fmt-multi-format');
    const ws = Workspace.open(out);
    Formatting.bold(ws, '268,635');
    Formatting.color(ws, '268,635', 'red');
    Formatting.highlight(ws, '268,635', 'yellow');
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    const runMatch = docXml.match(/<w:r>[^]*?268,635[^]*?<\/w:r>/);
    assert.ok(runMatch, 'should find run with text');
    assert.ok(runMatch[0].includes('<w:b/>'), 'should have bold');
    assert.ok(runMatch[0].includes('<w:color w:val="FF0000"/>'), 'should have color');
    assert.ok(runMatch[0].includes('<w:highlight w:val="yellow"/>'), 'should have highlight');
  });
});

// ============================================================================
// 3. CROSS-RUN TEXT
// ============================================================================

describe('formatting - text spanning runs', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('formats text that spans multiple runs', () => {
    // The fixture has "Meta's advertising ban" split across runs
    // due to the apostrophe entity
    const out = freshCopy('fmt-cross-run');
    const ws = Workspace.open(out);
    Formatting.bold(ws, "Meta's advertising ban");
    ws.save(out);

    const docXml = readDocxXml(out, 'word/document.xml');
    // The text should be formatted -- all pieces should have w:b
    // After formatting, the matched portion gets w:b
    assert.ok(docXml.includes('<w:b/>'), 'document should contain bold formatting');
    // Verify text is still present
    assert.ok(docXml.includes('advertising ban'), 'text content preserved');
  });
});

// ============================================================================
// 4. ERROR HANDLING
// ============================================================================

describe('formatting - error handling', () => {
  let Formatting, Workspace;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
    Workspace = require('../src/workspace').Workspace;
  });

  it('throws when text is not found', () => {
    const out = freshCopy('fmt-not-found');
    const ws = Workspace.open(out);
    assert.throws(
      () => Formatting.bold(ws, 'this text absolutely does not exist in the document'),
      /Text not found/
    );
    ws.cleanup();
  });
});

// ============================================================================
// 5. STATIC PROPERTIES
// ============================================================================

describe('formatting - static properties', () => {
  let Formatting;

  before(() => {
    Formatting = require('../src/formatting').Formatting;
  });

  it('COLORS has standard named colors', () => {
    assert.equal(Formatting.COLORS.red, 'FF0000');
    assert.equal(Formatting.COLORS.blue, '0000FF');
    assert.equal(Formatting.COLORS.green, '008000');
    assert.equal(Formatting.COLORS.black, '000000');
  });

  it('HIGHLIGHTS has Word built-in highlight colors', () => {
    assert.ok(Formatting.HIGHLIGHTS.includes('yellow'));
    assert.ok(Formatting.HIGHLIGHTS.includes('green'));
    assert.ok(Formatting.HIGHLIGHTS.includes('cyan'));
    assert.ok(Formatting.HIGHLIGHTS.includes('darkBlue'));
    assert.equal(Formatting.HIGHLIGHTS.length, 14);
  });
});
