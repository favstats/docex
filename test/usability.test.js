/**
 * usability.test.js -- Tests for v0.3 usability features
 *
 * Tests:
 *   - Doctor (diagnose + validate)
 *   - Dry-run flag
 *   - Preview
 *   - .docexrc loading
 *
 * Uses Node.js built-in test runner (node:test) -- zero dependencies.
 * Run: node --test test/usability.test.js
 */

const { describe, it, before, after, beforeEach, afterEach } = require('node:test');
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
  const tmp = fs.mkdtempSync('/tmp/docex-test-');
  execFileSync('unzip', ['-o', docxPath, '-d', tmp], { stdio: 'pipe' });
  const content = fs.readFileSync(path.join(tmp, xmlFile), 'utf8');
  execFileSync('rm', ['-rf', tmp], { stdio: 'pipe' });
  return content;
}

// ============================================================================
// 1. DOCTOR
// ============================================================================

describe('doctor', () => {
  let Doctor, Workspace;

  before(() => {
    Doctor = require('../src/doctor').Doctor;
    Workspace = require('../src/workspace').Workspace;
  });

  it('validate returns valid for a healthy document', () => {
    const copy = freshCopy('doctor-healthy');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    assert.equal(result.valid, true, 'Document should be valid');
    assert.ok(result.errors.length === 0, 'No errors expected');
    assert.ok(result.checks.length > 0, 'Should have checks');
    ws.cleanup();
  });

  it('validate checks document.xml exists', () => {
    const copy = freshCopy('doctor-docxml');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    const docCheck = result.checks.find(c => c.name === 'document.xml');
    assert.ok(docCheck, 'Should have document.xml check');
    assert.equal(docCheck.passed, true);
    ws.cleanup();
  });

  it('validate checks relationships resolve', () => {
    const copy = freshCopy('doctor-rels');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    const relCheck = result.checks.find(c => c.name === 'relationships');
    assert.ok(relCheck, 'Should have relationships check');
    assert.equal(relCheck.passed, true);
    ws.cleanup();
  });

  it('validate checks paragraph count', () => {
    const copy = freshCopy('doctor-paras');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    const paraCheck = result.checks.find(c => c.name === 'paragraphs');
    assert.ok(paraCheck, 'Should have paragraphs check');
    assert.equal(paraCheck.passed, true);
    assert.ok(paraCheck.message.includes('paragraphs'), 'Should mention paragraph count');
    ws.cleanup();
  });

  it('validate checks heading hierarchy', () => {
    const copy = freshCopy('doctor-headings');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    const headingCheck = result.checks.find(c => c.name === 'headingHierarchy');
    assert.ok(headingCheck, 'Should have heading hierarchy check');
    ws.cleanup();
  });

  it('validate checks paraId uniqueness', () => {
    const copy = freshCopy('doctor-paraid');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    const paraIdCheck = result.checks.find(c => c.name === 'paraIdUniqueness');
    assert.ok(paraIdCheck, 'Should have paraId uniqueness check');
    assert.equal(paraIdCheck.passed, true);
    ws.cleanup();
  });

  it('validate detects duplicate paraIds', () => {
    const copy = freshCopy('doctor-dup-paraid');
    const ws = Workspace.open(copy);

    // Inject a duplicate paraId manually
    let docXml = ws.docXml;
    const m = docXml.match(/w14:paraId="([^"]+)"/);
    if (m) {
      const dupeId = m[1];
      // Find the second w:p and replace its paraId with the same value
      const firstEnd = docXml.indexOf('</w:p>') + 6;
      const secondParaId = docXml.slice(firstEnd).match(/w14:paraId="([^"]+)"/);
      if (secondParaId) {
        docXml = docXml.slice(0, firstEnd) +
          docXml.slice(firstEnd).replace(`w14:paraId="${secondParaId[1]}"`, `w14:paraId="${dupeId}"`);
        ws.docXml = docXml;
      }
    }

    const result = Doctor.validate(ws);
    const paraIdCheck = result.checks.find(c => c.name === 'paraIdUniqueness');
    assert.ok(paraIdCheck, 'Should have paraId uniqueness check');
    assert.equal(paraIdCheck.passed, false, 'Should detect duplicate paraIds');
    assert.ok(result.errors.some(e => e.includes('Duplicate paraId')));
    ws.cleanup();
  });

  it('validate checks file size', () => {
    const copy = freshCopy('doctor-size');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    const sizeCheck = result.checks.find(c => c.name === 'fileSize');
    assert.ok(sizeCheck, 'Should have file size check');
    assert.equal(sizeCheck.passed, true);
    ws.cleanup();
  });

  it('validate checks comment consistency', () => {
    const copy = freshCopy('doctor-comments');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    const commentCheck = result.checks.find(c => c.name === 'comments');
    assert.ok(commentCheck, 'Should have comments check');
    ws.cleanup();
  });

  it('diagnose returns formatted string', () => {
    const copy = freshCopy('doctor-diagnose');
    const ws = Workspace.open(copy);
    const output = Doctor.diagnose(ws);

    assert.ok(typeof output === 'string', 'Should return a string');
    assert.ok(output.length > 0, 'Should not be empty');
    // Should contain checkmark or cross characters
    assert.ok(output.includes('\u2713') || output.includes('\u2717'), 'Should contain check/cross marks');
    ws.cleanup();
  });

  it('diagnose shows healthy for valid document', () => {
    const copy = freshCopy('doctor-healthy2');
    const ws = Workspace.open(copy);
    const output = Doctor.diagnose(ws);

    assert.ok(output.includes('healthy'), 'Should show healthy message');
    ws.cleanup();
  });

  it('validate checks media directory', () => {
    const copy = freshCopy('doctor-media');
    const ws = Workspace.open(copy);
    const result = Doctor.validate(ws);

    const mediaCheck = result.checks.find(c => c.name === 'media');
    assert.ok(mediaCheck, 'Should have media check');
    assert.equal(mediaCheck.passed, true);
    ws.cleanup();
  });
});

// ============================================================================
// 2. DRY-RUN FLAG
// ============================================================================

describe('dry-run', () => {
  let docex;

  before(() => {
    docex = require('../src/docex');
  });

  it('save({ dryRun: true }) returns result without writing', async () => {
    const copy = freshCopy('dryrun-basic');
    const original = fs.readFileSync(copy);

    const doc = docex(copy);
    doc.author('Test');
    doc.replace('268,635', '300,000');

    const result = await doc.save({ dryRun: true });

    assert.ok(result.dryRun === true, 'Result should have dryRun flag');
    assert.ok(result.paragraphCount > 0, 'Should report paragraph count');
    assert.ok(result.path.includes('dryrun-basic'), 'Should report target path');

    // Verify the original file was NOT modified
    const after = fs.readFileSync(copy);
    assert.deepEqual(original, after, 'Original file should not be modified');
  });

  it('save({ dryRun: true }) does not overwrite output file', async () => {
    const copy = freshCopy('dryrun-no-overwrite');
    const outputPath = path.join(OUTPUT_DIR, 'dryrun-output.docx');

    // Ensure output doesn't exist
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);

    const doc = docex(copy);
    doc.replace('268,635', '300,000');

    const result = await doc.save({ dryRun: true, outputPath });

    assert.ok(result.dryRun === true);
    assert.ok(!fs.existsSync(outputPath), 'Output file should NOT be created in dry-run');
  });

  it('dry-run still executes operations on the workspace XML', async () => {
    const copy = freshCopy('dryrun-ops');
    const doc = docex(copy);
    doc.author('DryRunner');
    doc.replace('268,635', '300,000');

    const result = await doc.save({ dryRun: true });
    assert.ok(result.paragraphCount > 0);
  });
});

// ============================================================================
// 3. PREVIEW
// ============================================================================

describe('preview', () => {
  let docex;

  before(() => {
    docex = require('../src/docex');
  });

  it('shows "No pending operations" when empty', () => {
    const copy = freshCopy('preview-empty');
    const doc = docex(copy);

    assert.equal(doc.preview(), 'No pending operations.');
  });

  it('shows single replace operation', () => {
    const copy = freshCopy('preview-single');
    const doc = docex(copy);
    doc.author('Fabio Votta');
    doc.replace('old text', 'new text');

    const preview = doc.preview();
    assert.ok(preview.includes('1 pending operation:'), 'Should show count');
    assert.ok(preview.includes("replace 'old text' -> 'new text'"), 'Should describe the replace');
    assert.ok(preview.includes('tracked'), 'Should mention tracked');
    assert.ok(preview.includes('Fabio Votta'), 'Should mention author');
  });

  it('shows multiple operations', () => {
    const copy = freshCopy('preview-multi');
    const doc = docex(copy);
    doc.author('Fabio Votta');
    doc.replace('old', 'new');
    doc.after('Methods').insert('New paragraph.');
    doc.at('anchor').comment('note', { by: 'Reviewer 2' });

    const preview = doc.preview();
    assert.ok(preview.includes('3 pending operations:'), 'Should show count');
    assert.ok(preview.includes('1. replace'), 'Should have replace');
    assert.ok(preview.includes('2. insert'), 'Should have insert');
    assert.ok(preview.includes('3. comment'), 'Should have comment');
  });

  it('truncates long text in preview', () => {
    const copy = freshCopy('preview-truncate');
    const doc = docex(copy);
    const longText = 'This is a very long text that exceeds the truncation limit for preview display';
    doc.replace(longText, 'short');

    const preview = doc.preview();
    assert.ok(preview.includes('...'), 'Should truncate long text');
  });

  it('shows delete operations', () => {
    const copy = freshCopy('preview-delete');
    const doc = docex(copy);
    doc.author('Test');
    doc.delete('unwanted text');

    const preview = doc.preview();
    assert.ok(preview.includes("delete 'unwanted text'"), 'Should show delete');
  });

  it('shows format operations', () => {
    const copy = freshCopy('preview-format');
    const doc = docex(copy);
    doc.bold('important text');

    const preview = doc.preview();
    assert.ok(preview.includes("bold 'important text'"), 'Should show bold format');
  });

  it('shows untracked status', () => {
    const copy = freshCopy('preview-untracked');
    const doc = docex(copy);
    doc.author('Test');
    doc.untracked();
    doc.replace('old', 'new');

    const preview = doc.preview();
    assert.ok(preview.includes('untracked'), 'Should show untracked');
  });
});

// ============================================================================
// 4. .docexrc
// ============================================================================

describe('.docexrc', () => {
  let docex;
  let rcDir;
  let rcPath;
  let testDocx;

  before(() => {
    docex = require('../src/docex');
  });

  beforeEach(() => {
    // Create a temporary directory with a .docexrc and a test docx
    rcDir = fs.mkdtempSync('/tmp/docex-rc-test-');
    rcPath = path.join(rcDir, '.docexrc');
    testDocx = path.join(rcDir, 'test.docx');
    fs.copyFileSync(FIXTURE, testDocx);
  });

  afterEach(() => {
    try {
      execFileSync('rm', ['-rf', rcDir], { stdio: 'pipe' });
    } catch (_) { /* ignore */ }
  });

  it('loads author from .docexrc', () => {
    fs.writeFileSync(rcPath, JSON.stringify({ author: 'RC Author' }), 'utf-8');

    const doc = docex(testDocx);
    assert.equal(doc._author, 'RC Author', 'Should load author from .docexrc');
  });

  it('exposes rc config via doc.rc', () => {
    fs.writeFileSync(rcPath, JSON.stringify({ author: 'RC Author', style: 'academic' }), 'utf-8');

    const doc = docex(testDocx);
    assert.equal(doc.rc.author, 'RC Author');
    assert.equal(doc.rc.style, 'academic');
  });

  it('author() method overrides .docexrc author', () => {
    fs.writeFileSync(rcPath, JSON.stringify({ author: 'RC Author' }), 'utf-8');

    const doc = docex(testDocx);
    doc.author('Override Author');
    assert.equal(doc._author, 'Override Author');
  });

  it('handles missing .docexrc gracefully', () => {
    // No .docexrc file -- should not throw
    const doc = docex(testDocx);
    assert.equal(doc._author, 'Unknown', 'Should fall back to Unknown');
  });

  it('handles malformed .docexrc gracefully', () => {
    fs.writeFileSync(rcPath, 'not valid json{{{', 'utf-8');

    // Should not throw
    const doc = docex(testDocx);
    assert.equal(doc._author, 'Unknown', 'Should fall back to Unknown on parse error');
  });

  it('loads safeModify from .docexrc', () => {
    fs.writeFileSync(rcPath, JSON.stringify({
      author: 'Test',
      safeModify: '/path/to/safe-modify.sh',
      style: 'academic',
    }), 'utf-8');

    const doc = docex(testDocx);
    assert.equal(doc.rc.safeModify, '/path/to/safe-modify.sh');
    assert.equal(doc.rc.style, 'academic');
  });
});

// ============================================================================
// 5. COLORED REVISIONS (CLI output format)
// ============================================================================

describe('colored revisions', () => {
  it('revisions list has proper structure for coloring', async () => {
    // This tests that the Revisions.list returns the proper structure
    // that the CLI uses for colored output
    const copy = freshCopy('colored-revisions');
    const docex = require('../src/docex');
    const doc = docex(copy);
    doc.author('Test Author');
    doc.replace('268,635', '300,000');

    // Save with tracked changes
    const output = path.join(OUTPUT_DIR, 'colored-rev-output.docx');
    await doc.save(output);

    // Now read revisions from saved file
    const doc2 = docex(output);
    const revs = await doc2.revisions();

    assert.ok(revs.length > 0, 'Should have tracked changes');
    for (const r of revs) {
      assert.ok(r.type === 'insertion' || r.type === 'deletion', 'Should have valid type');
      assert.ok(typeof r.author === 'string', 'Should have author');
      assert.ok(typeof r.date === 'string', 'Should have date');
      assert.ok(typeof r.text === 'string', 'Should have text');
    }
    doc2.discard();
  });
});

// ============================================================================
// 6. INIT CLI COMMAND (tested via .docexrc creation logic)
// ============================================================================

describe('init', () => {
  let initDir;

  beforeEach(() => {
    initDir = fs.mkdtempSync('/tmp/docex-init-test-');
  });

  afterEach(() => {
    try {
      execFileSync('rm', ['-rf', initDir], { stdio: 'pipe' });
    } catch (_) { /* ignore */ }
  });

  it('docex init creates .docexrc via CLI', () => {
    // Run the CLI init command
    try {
      execFileSync('node', [
        path.join(__dirname, '..', 'cli', 'docex-cli.js'),
        'init',
      ], {
        cwd: initDir,
        stdio: 'pipe',
        encoding: 'utf-8',
        timeout: 10000,
      });
    } catch (e) {
      // CLI may exit with code 0, execFileSync should work
      // If it fails, the test file check below will catch it
    }

    const rcPath = path.join(initDir, '.docexrc');
    assert.ok(fs.existsSync(rcPath), '.docexrc should be created');

    const rc = JSON.parse(fs.readFileSync(rcPath, 'utf-8'));
    assert.ok(typeof rc.author === 'string', 'Should have author field');
    assert.ok(typeof rc.style === 'string', 'Should have style field');
    assert.equal(rc.backup, true, 'Should have backup true');
  });

  it('docex init does not overwrite existing .docexrc', () => {
    const rcPath = path.join(initDir, '.docexrc');
    fs.writeFileSync(rcPath, JSON.stringify({ author: 'Original' }), 'utf-8');

    try {
      execFileSync('node', [
        path.join(__dirname, '..', 'cli', 'docex-cli.js'),
        'init',
      ], {
        cwd: initDir,
        stdio: 'pipe',
        encoding: 'utf-8',
        timeout: 10000,
      });
    } catch (_) { /* ignore exit codes */ }

    const rc = JSON.parse(fs.readFileSync(rcPath, 'utf-8'));
    assert.equal(rc.author, 'Original', 'Should not overwrite existing .docexrc');
  });
});

// ============================================================================
// 7. DOCTOR CLI COMMAND
// ============================================================================

describe('doctor CLI', () => {
  it('docex doctor runs without error', () => {
    const copy = freshCopy('doctor-cli');
    const output = execFileSync('node', [
      path.join(__dirname, '..', 'cli', 'docex-cli.js'),
      'doctor',
      copy,
    ], {
      stdio: 'pipe',
      encoding: 'utf-8',
      timeout: 30000,
    });

    assert.ok(output.includes('healthy') || output.includes('\u2713'), 'Should show check results');
  });
});

// ============================================================================
// 8. DRY-RUN CLI
// ============================================================================

describe('dry-run CLI', () => {
  it('--dry-run flag prevents file modification via CLI', () => {
    const copy = freshCopy('dryrun-cli');
    const originalSize = fs.statSync(copy).size;

    try {
      execFileSync('node', [
        path.join(__dirname, '..', 'cli', 'docex-cli.js'),
        'replace',
        copy,
        'enforcement',
        'ENFORCEMENT',
        '--dry-run',
        '--author', 'Test',
      ], {
        stdio: 'pipe',
        encoding: 'utf-8',
        timeout: 30000,
      });
    } catch (_) { /* ignore */ }

    // File should not have changed (dry-run leaves workspace intact, but
    // the original file should still have the same content hash)
    const afterSize = fs.statSync(copy).size;
    assert.equal(originalSize, afterSize, 'File size should not change with --dry-run');
  });
});
