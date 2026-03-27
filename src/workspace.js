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
const os = require('os');
const { execFileSync } = require('child_process');
const crypto = require('crypto');
const xml = require('./xml');

// ============================================================================
// LOCK FILE HELPERS
// ============================================================================

/**
 * Check if a process with the given PID is still alive.
 * @param {number} pid
 * @returns {boolean}
 */
function _isPidAlive(pid) {
  try {
    process.kill(pid, 0); // signal 0 = existence check, no actual signal sent
    return true;
  } catch (_) {
    return false;
  }
}

/**
 * Build lock file path for a given docx path.
 * Lock file is `.filename.docx.docex-lock` in the same directory.
 * @param {string} docxPath - Absolute path to the docx
 * @returns {string}
 */
function _lockFilePath(docxPath) {
  const dir = path.dirname(docxPath);
  const base = path.basename(docxPath);
  return path.join(dir, '.' + base + '.docex-lock');
}

/**
 * Reference counter for lock files within the same process.
 * Tracks how many open workspaces hold a lock on each file path.
 * @type {Map<string, number>}
 */
const _lockRefCounts = new Map();

// ============================================================================
// BACKUP HELPERS
// ============================================================================

const MAX_BACKUPS = 20;

/**
 * Create a timestamped backup of a file in .docex-backups/ next to it.
 * Prunes backups older than MAX_BACKUPS.
 * @param {string} docxPath - Absolute path to the original docx
 */
function _createBackup(docxPath) {
  if (!fs.existsSync(docxPath)) return;

  const dir = path.dirname(docxPath);
  const base = path.basename(docxPath, '.docx');
  const backupDir = path.join(dir, '.docex-backups');

  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }

  // Timestamp: YYYYMMDD_HHMMSS
  const now = new Date();
  const ts = now.getFullYear().toString()
    + String(now.getMonth() + 1).padStart(2, '0')
    + String(now.getDate()).padStart(2, '0')
    + '_'
    + String(now.getHours()).padStart(2, '0')
    + String(now.getMinutes()).padStart(2, '0')
    + String(now.getSeconds()).padStart(2, '0');

  const backupName = `${base}_${ts}.docx`;
  const backupPath = path.join(backupDir, backupName);
  fs.copyFileSync(docxPath, backupPath);

  // Prune: keep only the newest MAX_BACKUPS
  _pruneBackups(backupDir, base);
}

/**
 * Remove oldest backups so at most MAX_BACKUPS remain.
 * @param {string} backupDir
 * @param {string} baseName - Filename prefix (without extension)
 */
function _pruneBackups(backupDir, baseName) {
  const prefix = baseName + '_';
  let files;
  try {
    files = fs.readdirSync(backupDir)
      .filter(f => f.startsWith(prefix) && f.endsWith('.docx'))
      .sort(); // lexicographic sort = chronological for YYYYMMDD_HHMMSS
  } catch (_) {
    return;
  }

  while (files.length > MAX_BACKUPS) {
    const oldest = files.shift();
    try {
      fs.unlinkSync(path.join(backupDir, oldest));
    } catch (_) { /* ignore */ }
  }
}

// Late-loaded to avoid circular dependency (docmap requires paragraphs which is fine,
// but docmap is only needed at open time for paraId injection).
let _DocMap = null;
function getDocMap() {
  if (!_DocMap) {
    _DocMap = require('./docmap').DocMap;
  }
  return _DocMap;
}

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

/**
 * Minimal footnotes.xml content for a .docx that has no footnotes yet.
 * Includes the two required built-in separator footnotes (id=0 and id=1).
 */
const EMPTY_FOOTNOTES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  + `<w:footnotes xmlns:w="${xml.NS.w}" xmlns:r="${xml.NS.r}" `
  + `xmlns:w14="${xml.NS.w14}" xmlns:mc="${xml.NS.mc}" `
  + `mc:Ignorable="w14">`
  + '<w:footnote w:type="separator" w:id="0">'
  + '<w:p><w:r><w:separator/></w:r></w:p>'
  + '</w:footnote>'
  + '<w:footnote w:type="continuationSeparator" w:id="1">'
  + '<w:p><w:r><w:continuationSeparator/></w:r></w:p>'
  + '</w:footnote>'
  + '</w:footnotes>';

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
  static open(docxPath, opts = {}) {
    const absPath = path.resolve(docxPath);
    if (!fs.existsSync(absPath)) {
      throw new Error(`File not found: ${absPath}`);
    }

    const ws = new Workspace(absPath);
    ws._lockPath = null;
    ws._unzip();
    // Inject w14:paraId on every <w:p> that lacks one (stable addressing)
    const DocMap = getDocMap();
    DocMap.injectParaIds(ws);
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

    /** @type {string|null} Lock file path (set by open()) */
    this._lockPath = null;

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
    /** @type {string|null} */ this._corePropsXml = null;
    /** @type {string|null} */ this._rootRelsXml = null;
    /** @type {string|null} */ this._footnotesXml = null;

    // Track which XML files have been modified
    /** @type {Set<string>} */ this._dirty = new Set();

    // Snapshot stack for in-memory rollback (LIFO)
    /** @type {Array<object>} */ this._snapshots = [];
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

  /** @returns {string} word/styles.xml content */
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

  /** @param {string} val */
  set stylesXml(val) {
    this._stylesXml = val;
    this._dirty.add('stylesXml');
  }

  /** @returns {string|null} docProps/core.xml content, or null if not present */
  get corePropsXml() {
    if (this._corePropsXml === null) {
      const filePath = path.join(this._tmpDir, 'docProps', 'core.xml');
      if (fs.existsSync(filePath)) {
        this._corePropsXml = fs.readFileSync(filePath, 'utf-8');
      } else {
        // Return null to indicate file doesn't exist (unlike comments which auto-create)
        return null;
      }
    }
    return this._corePropsXml;
  }

  /** @param {string} val */
  set corePropsXml(val) {
    this._corePropsXml = val;
    this._dirty.add('corePropsXml');
  }

  /** @returns {string|null} _rels/.rels content (root-level relationships) */
  get rootRelsXml() {
    if (this._rootRelsXml === null) {
      const filePath = path.join(this._tmpDir, '_rels', '.rels');
      if (fs.existsSync(filePath)) {
        this._rootRelsXml = fs.readFileSync(filePath, 'utf-8');
      } else {
        return null;
      }
    }
    return this._rootRelsXml;
  }

  /** @param {string} val */
  set rootRelsXml(val) {
    this._rootRelsXml = val;
    this._dirty.add('rootRelsXml');
  }

  /** @returns {string} word/footnotes.xml content (created with separators if missing) */
  get footnotesXml() {
    if (this._footnotesXml === null) {
      const filePath = path.join(this._tmpDir, 'word', 'footnotes.xml');
      if (fs.existsSync(filePath)) {
        this._footnotesXml = fs.readFileSync(filePath, 'utf-8');
      } else {
        this._footnotesXml = EMPTY_FOOTNOTES_XML;
        this._dirty.add('footnotesXml');
      }
    }
    return this._footnotesXml;
  }

  /** @param {string} val */
  set footnotesXml(val) {
    this._footnotesXml = val;
    this._dirty.add('footnotesXml');
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
  // Snapshot / Rollback (in-memory state save/restore)
  // --------------------------------------------------------------------------

  /**
   * Save a copy of the current document state in memory.
   * Multiple snapshots stack (LIFO). Use rollback() to restore.
   *
   * Captures: docXml, commentsXml, commentsExtXml, commentsIdsXml,
   * stylesXml, footnotesXml, and the dirty set.
   */
  snapshot() {
    this._snapshots.push({
      docXml: this.docXml,
      commentsXml: this.commentsXml,
      commentsExtXml: this.commentsExtXml,
      commentsIdsXml: this.commentsIdsXml,
      stylesXml: this._stylesXml,
      footnotesXml: this._footnotesXml,
      dirty: new Set(this._dirty),
    });
  }

  /**
   * Restore the most recent snapshot, discarding current state.
   * @returns {boolean} true if a snapshot was restored, false if stack was empty
   */
  rollback() {
    const snap = this._snapshots.pop();
    if (!snap) return false;

    this._docXml = snap.docXml;
    this._commentsXml = snap.commentsXml;
    this._commentsExtXml = snap.commentsExtXml;
    this._commentsIdsXml = snap.commentsIdsXml;
    this._stylesXml = snap.stylesXml;
    this._footnotesXml = snap.footnotesXml;
    this._dirty = new Set(snap.dirty);
    return true;
  }

  // --------------------------------------------------------------------------
  // Save and verify
  // --------------------------------------------------------------------------

  /**
   * Write all modified XML files back to the temp directory, rezip,
   * verify integrity, and clean up.
   *
   * @param {string|object} [outputPathOrOpts] - Target path string, or options object:
   *   - outputPath {string}    Target path (defaults to original docx path)
   *   - safeModify {string}    Path to safe-modify.sh script
   *   - description {string}   Description for safe-modify.sh commit message
   * @returns {{path: string, fileSize: number, paragraphCount: number, verified: boolean}}
   */
  save(outputPathOrOpts) {
    let target;
    let safeModify = null;
    let description = 'docex edit';
    let doBackup = true;

    let dryRun = false;

    if (typeof outputPathOrOpts === 'object' && outputPathOrOpts !== null) {
      target = outputPathOrOpts.outputPath
        ? path.resolve(outputPathOrOpts.outputPath)
        : this._docxPath;
      safeModify = outputPathOrOpts.safeModify || null;
      description = outputPathOrOpts.description || 'docex edit';
      if (outputPathOrOpts.backup === false) doBackup = false;
      dryRun = !!outputPathOrOpts.dryRun;
    } else {
      target = outputPathOrOpts ? path.resolve(outputPathOrOpts) : this._docxPath;
    }

    // Read .docexrc in the target directory to check backup setting
    if (doBackup) {
      try {
        const rcPath = path.join(path.dirname(target), '.docexrc');
        if (fs.existsSync(rcPath)) {
          const rc = JSON.parse(fs.readFileSync(rcPath, 'utf-8'));
          if (rc.backup === false) doBackup = false;
        }
      } catch (_) { /* ignore rc parse errors */ }
    }

    // Create backup before overwriting
    if (doBackup && fs.existsSync(target)) {
      _createBackup(target);
    }

    // Acquire lock file on target path to prevent concurrent writes
    if (!dryRun) {
      const lockPath = _lockFilePath(target);
      try {
        const lockContent = fs.readFileSync(lockPath, 'utf-8');
        const lockInfo = JSON.parse(lockContent);
        if (lockInfo.pid && lockInfo.pid !== process.pid && _isPidAlive(lockInfo.pid)) {
          throw new Error(
            `File is being edited by another docex process (PID ${lockInfo.pid}, started ${lockInfo.started || 'unknown'})`
          );
        }
      } catch (err) {
        if (err.message && err.message.includes('being edited')) throw err;
        // No lock file or corrupt -- proceed
      }
      const lockFileData = {
        pid: process.pid,
        started: new Date().toISOString(),
        user: os.userInfo().username,
      };
      try {
        fs.writeFileSync(lockPath, JSON.stringify(lockFileData, null, 2), 'utf-8');
        this._lockPath = lockPath;
      } catch (_) { /* advisory lock, non-critical */ }
    }

    // Write all dirty XML files back to disk
    this._flush();

    // Dry-run mode: return result without writing to disk
    if (dryRun) {
      const newCount = this._countParagraphsInXml(this.docXml);
      return {
        path: target,
        fileSize: 0,
        paragraphCount: newCount,
        verified: true,
        dryRun: true,
      };
    }

    if (safeModify) {
      // Safe-modify path: save to temp file, then use safe-modify.sh to copy over
      const tmpFile = path.join('/tmp', `docex-safe-${crypto.randomBytes(8).toString('hex')}.docx`);
      try {
        this._rezip(tmpFile);

        // Call safe-modify.sh: bash <script> "<description>" cp <temp> <target>
        execFileSync('bash', [safeModify, description, 'cp', tmpFile, target], {
          stdio: 'inherit',
          timeout: 60000,
        });

        // Gather stats from the target (safe-modify.sh copied tmpFile there)
        const stat = fs.statSync(target);
        const newCount = this._countParagraphsInXml(this.docXml);
        const verified = this._verify(target, newCount, stat.size);

        this.cleanup();

        return {
          path: target,
          fileSize: stat.size,
          paragraphCount: newCount,
          verified,
        };
      } finally {
        // Clean up temp file
        try { fs.unlinkSync(tmpFile); } catch (_) { /* already removed */ }
      }
    }

    // Direct save path (original behavior)
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
   * Remove the temp directory and lock file. Safe to call multiple times.
   */
  cleanup() {
    if (this._tmpDir && fs.existsSync(this._tmpDir)) {
      execFileSync('rm', ['-rf', this._tmpDir], { stdio: 'pipe' });
      this._tmpDir = '';
    }
    // Remove lock file only when reference count reaches 0
    if (this._lockPath) {
      const count = (_lockRefCounts.get(this._lockPath) || 1) - 1;
      if (count <= 0) {
        _lockRefCounts.delete(this._lockPath);
        try { fs.unlinkSync(this._lockPath); } catch (_) { /* already removed */ }
      } else {
        _lockRefCounts.set(this._lockPath, count);
      }
      this._lockPath = null;
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
      stylesXml:       'word/styles.xml',
      footnotesXml:    'word/footnotes.xml',
      corePropsXml:    'docProps/core.xml',
      rootRelsXml:     '_rels/.rels',
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
