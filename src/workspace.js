/**
 * workspace.js -- Manages the unzip/rezip lifecycle for .docx files.
 *
 * Wraps the temp-directory pattern from suggest-edit-safe.js and docx-patch.js
 * into a clean class with lazy-loaded XML accessors.
 *
 * Zero external dependencies. Uses execFileSync for zip operations.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const crypto = require('crypto');
const xml = require('./xml');

// ============================================================================
// EMPTY DOCUMENT TEMPLATES
// ============================================================================

/**
 * Minimal comments.xml content for a .docx that has no comments yet.
 */
const EMPTY_COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  + `<w:comments xmlns:w="${xml.NS.w}" xmlns:r="${xml.NS.r}" `
  + `xmlns:w14="${xml.NS.w14}" xmlns:mc="${xml.NS.mc}" `
  + `mc:Ignorable="w14">`
  + '</w:comments>';

/**
 * Minimal commentsExtended.xml for .docx files that need threaded replies.
 */
const EMPTY_COMMENTS_EXT_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  + `<w15:commentsEx xmlns:w15="${xml.NS.w15}" `
  + `xmlns:mc="${xml.NS.mc}" mc:Ignorable="w15">`
  + '</w15:commentsEx>';

/**
 * Minimal commentsIds.xml for stable comment identification.
 */
const EMPTY_COMMENTS_IDS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  + `<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" `
  + `xmlns:mc="${xml.NS.mc}" mc:Ignorable="w16cid">`
  + '</w16cid:commentsIds>';

// ============================================================================
// WORKSPACE CLASS
// ============================================================================

class Workspace {

  /**
   * Open a .docx file for editing.
   * Unzips to a temp directory and counts original paragraphs.
   * @param {string} docxPath - Absolute or relative path to the .docx file
   * @returns {Workspace}
   */
  static open(docxPath) {
    const absPath = path.resolve(docxPath);
    if (!fs.existsSync(absPath)) {
      throw new Error(`File not found: ${absPath}`);
    }
    const ws = new Workspace(absPath);
    ws._unzip();
    ws._countOriginalParagraphs();
    return ws;
  }

  // --------------------------------------------------------------------------
  // Constructor (private-ish; use Workspace.open())
  // --------------------------------------------------------------------------

  /**
   * @param {string} docxPath - Resolved absolute path
   */
  constructor(docxPath) {
    /** @type {string} Original .docx path */
    this._docxPath = docxPath;

    /** @type {string} Temp directory holding unzipped contents */
    this._tmpDir = '';

    /** @type {number} Paragraph count at open time */
    this._originalParagraphCount = 0;

    // Lazy-loaded XML caches (null = not yet loaded)
    /** @type {string|null} */ this._docXml = null;
    /** @type {string|null} */ this._commentsXml = null;
    /** @type {string|null} */ this._commentsExtXml = null;
    /** @type {string|null} */ this._commentsIdsXml = null;
    /** @type {string|null} */ this._relsXml = null;
    /** @type {string|null} */ this._contentTypesXml = null;
    /** @type {string|null} */ this._stylesXml = null;

    // Track which XML files have been modified
    /** @type {Set<string>} */ this._dirty = new Set();
  }

  // --------------------------------------------------------------------------
  // XML Accessors (lazy loading)
  // --------------------------------------------------------------------------

  /** @returns {string} word/document.xml content */
  get docXml() {
    if (this._docXml === null) {
      this._docXml = this._readFile('word/document.xml');
    }
    return this._docXml;
  }

  /** @param {string} val */
  set docXml(val) {
    this._docXml = val;
    this._dirty.add('docXml');
  }

  /** @returns {string} word/comments.xml content (created if missing) */
  get commentsXml() {
    if (this._commentsXml === null) {
      const filePath = path.join(this._tmpDir, 'word', 'comments.xml');
      if (fs.existsSync(filePath)) {
        this._commentsXml = fs.readFileSync(filePath, 'utf-8');
      } else {
        this._commentsXml = EMPTY_COMMENTS_XML;
        this._dirty.add('commentsXml');
      }
    }
    return this._commentsXml;
  }

  /** @param {string} val */
  set commentsXml(val) {
    this._commentsXml = val;
    this._dirty.add('commentsXml');
  }

  /** @returns {string} word/commentsExtended.xml content (created if missing) */
  get commentsExtXml() {
    if (this._commentsExtXml === null) {
      const filePath = path.join(this._tmpDir, 'word', 'commentsExtended.xml');
      if (fs.existsSync(filePath)) {
        this._commentsExtXml = fs.readFileSync(filePath, 'utf-8');
      } else {
        this._commentsExtXml = EMPTY_COMMENTS_EXT_XML;
        this._dirty.add('commentsExtXml');
      }
    }
    return this._commentsExtXml;
  }

  /** @param {string} val */
  set commentsExtXml(val) {
    this._commentsExtXml = val;
    this._dirty.add('commentsExtXml');
  }

  /** @returns {string} word/commentsIds.xml content (created if missing) */
  get commentsIdsXml() {
    if (this._commentsIdsXml === null) {
      const filePath = path.join(this._tmpDir, 'word', 'commentsIds.xml');
      if (fs.existsSync(filePath)) {
        this._commentsIdsXml = fs.readFileSync(filePath, 'utf-8');
      } else {
        this._commentsIdsXml = EMPTY_COMMENTS_IDS_XML;
        this._dirty.add('commentsIdsXml');
      }
    }
    return this._commentsIdsXml;
  }

  /** @param {string} val */
  set commentsIdsXml(val) {
    this._commentsIdsXml = val;
    this._dirty.add('commentsIdsXml');
  }

  /** @returns {string} word/_rels/document.xml.rels content */
  get relsXml() {
    if (this._relsXml === null) {
      this._relsXml = this._readFile('word/_rels/document.xml.rels');
    }
    return this._relsXml;
  }

  /** @param {string} val */
  set relsXml(val) {
    this._relsXml = val;
    this._dirty.add('relsXml');
  }

  /** @returns {string} [Content_Types].xml content */
  get contentTypesXml() {
    if (this._contentTypesXml === null) {
      this._contentTypesXml = this._readFile('[Content_Types].xml');
    }
    return this._contentTypesXml;
  }

  /** @param {string} val */
  set contentTypesXml(val) {
    this._contentTypesXml = val;
    this._dirty.add('contentTypesXml');
  }

  /** @returns {string} word/styles.xml content (read-only) */
  get stylesXml() {
    if (this._stylesXml === null) {
      const filePath = path.join(this._tmpDir, 'word', 'styles.xml');
      if (fs.existsSync(filePath)) {
        this._stylesXml = fs.readFileSync(filePath, 'utf-8');
      } else {
        this._stylesXml = '';
      }
    }
    return this._stylesXml;
  }

  /** @returns {string} Absolute path to word/media/ directory */
  get mediaDir() {
    const dir = path.join(this._tmpDir, 'word', 'media');
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    return dir;
  }

  /** @returns {string} Absolute path to the temp directory */
  get tmpDir() {
    return this._tmpDir;
  }

  /** @returns {number} Paragraph count recorded at open time */
  get originalParagraphCount() {
    return this._originalParagraphCount;
  }

  // --------------------------------------------------------------------------
  // Save and verify
  // --------------------------------------------------------------------------

  /**
   * Write all modified XML files back to the temp directory, rezip,
   * verify integrity, and clean up.
   * @param {string} [outputPath] - Target path (defaults to original docx path)
   * @returns {{path: string, fileSize: number, paragraphCount: number, verified: boolean}}
   */
  save(outputPath) {
    const target = outputPath ? path.resolve(outputPath) : this._docxPath;

    // Write all dirty XML files back to disk
    this._flush();

    // Rezip
    if (fs.existsSync(target)) {
      fs.unlinkSync(target);
    }
    this._rezip(target);

    // Gather stats
    const stat = fs.statSync(target);
    const newCount = this._countParagraphsInXml(this.docXml);

    // Verify
    const verified = this._verify(target, newCount, stat.size);

    // Cleanup temp dir
    this.cleanup();

    return {
      path: target,
      fileSize: stat.size,
      paragraphCount: newCount,
      verified,
    };
  }

  /**
   * Verify document integrity without saving.
   * Checks: paragraph count >= original.
   * @returns {{valid: boolean, errors: string[]}}
   */
  verify() {
    const newCount = this._countParagraphsInXml(this.docXml);
    const errors = [];

    if (newCount < this._originalParagraphCount) {
      errors.push(
        `Paragraph count dropped: ${newCount} < ${this._originalParagraphCount} (original)`
      );
    }

    return { valid: errors.length === 0, errors };
  }

  /**
   * Remove the temp directory. Safe to call multiple times.
   */
  cleanup() {
    if (this._tmpDir && fs.existsSync(this._tmpDir)) {
      execFileSync('rm', ['-rf', this._tmpDir], { stdio: 'pipe' });
      this._tmpDir = '';
    }
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Create temp directory and unzip the .docx into it.
   * @private
   */
  _unzip() {
    const id = crypto.randomBytes(8).toString('hex');
    this._tmpDir = path.join('/tmp', `docex-ws-${id}`);
    fs.mkdirSync(this._tmpDir, { recursive: true });
    execFileSync('unzip', ['-o', this._docxPath, '-d', this._tmpDir], {
      stdio: 'pipe',
    });
  }

  /**
   * Rezip the temp directory into a .docx file.
   * @param {string} outputPath - Target file path
   * @private
   */
  _rezip(outputPath) {
    execFileSync('zip', ['-r', '-q', outputPath, '.'], {
      cwd: this._tmpDir,
      stdio: 'pipe',
    });
  }

  /**
   * Read a file from inside the temp directory.
   * @param {string} relPath - Path relative to temp dir root
   * @returns {string} File content as UTF-8 string
   * @private
   */
  _readFile(relPath) {
    const absPath = path.join(this._tmpDir, relPath);
    if (!fs.existsSync(absPath)) {
      throw new Error(`Missing required file in docx: ${relPath}`);
    }
    return fs.readFileSync(absPath, 'utf-8');
  }

  /**
   * Write a string to a file in the temp directory, creating parent dirs.
   * @param {string} relPath - Path relative to temp dir root
   * @param {string} content - File content
   * @private
   */
  _writeFile(relPath, content) {
    const absPath = path.join(this._tmpDir, relPath);
    const dir = path.dirname(absPath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(absPath, content, 'utf-8');
  }

  /**
   * Write all dirty XML caches back to the temp directory.
   * @private
   */
  _flush() {
    const fileMap = {
      docXml:          'word/document.xml',
      commentsXml:     'word/comments.xml',
      commentsExtXml:  'word/commentsExtended.xml',
      commentsIdsXml:  'word/commentsIds.xml',
      relsXml:         'word/_rels/document.xml.rels',
      contentTypesXml: '[Content_Types].xml',
    };

    for (const [key, relPath] of Object.entries(fileMap)) {
      if (this._dirty.has(key) && this['_' + key] !== null) {
        this._writeFile(relPath, this['_' + key]);
      }
    }
  }

  /**
   * Count <w:p> elements in the current document XML and store the count.
   * @private
   */
  _countOriginalParagraphs() {
    const docContent = this.docXml;
    this._originalParagraphCount = this._countParagraphsInXml(docContent);
  }

  /**
   * Count <w:p> elements in an XML string.
   * @param {string} xmlStr - XML content to scan
   * @returns {number}
   * @private
   */
  _countParagraphsInXml(xmlStr) {
    let count = 0;
    const re = /<w:p[\s>]/g;
    while (re.exec(xmlStr) !== null) {
      count++;
    }
    return count;
  }

  /**
   * Verify the output file after save.
   * @param {string} outputPath - Path to the saved .docx
   * @param {number} newCount - New paragraph count
   * @param {number} fileSize - Output file size in bytes
   * @returns {boolean} true if all checks pass
   * @private
   */
  _verify(outputPath, newCount, fileSize) {
    const errors = [];

    // Check valid zip
    try {
      execFileSync('unzip', ['-t', outputPath], { stdio: 'pipe' });
    } catch (e) {
      errors.push('Output is not a valid zip file');
    }

    // Check paragraph count
    if (newCount < this._originalParagraphCount) {
      errors.push(
        `Paragraph count dropped: ${newCount} < ${this._originalParagraphCount}`
      );
    }

    // Check file size ratio (reject suspiciously small output)
    try {
      const origSize = fs.statSync(this._docxPath).size;
      if (origSize > 0) {
        const ratio = fileSize / origSize;
        if (ratio < 0.2) {
          errors.push(
            `File size suspiciously small: ${fileSize} bytes vs original ${origSize} bytes`
          );
        }
      }
    } catch (e) {
      // Original file may have been overwritten; skip size check
    }

    if (errors.length > 0) {
      for (const err of errors) {
        console.error(`[workspace] VERIFY WARN: ${err}`);
      }
      return false;
    }

    return true;
  }
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = { Workspace };
