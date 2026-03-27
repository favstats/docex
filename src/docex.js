/**
 * docex -- LaTeX for .docx
 *
 * A clean, fluent API for programmatic document editing.
 * Tracked changes are the default. Author is set once.
 * Position selectors read like English.
 *
 * Usage:
 *   const doc = docex("manuscript.docx");
 *   doc.author("Fabio Votta");
 *
 *   doc.replace("old text", "new text");
 *   doc.after("Methods").insert("New paragraph.");
 *   doc.after("Enforcement").figure("fig03.png", "Figure 3. Status");
 *   doc.at("platform regulation").comment("Needs work", { by: "Reviewer 2" });
 *
 *   await doc.save();
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const crypto = require('crypto');

const { Workspace } = require('./workspace');
const { Paragraphs } = require('./paragraphs');
const { Comments } = require('./comments');
const { Figures } = require('./figures');
const { Tables } = require('./tables');
const { Citations } = require('./citations');
const { Latex } = require('./latex');
const { TextMap } = require('./textmap');
const xml = require('./xml');

// ============================================================================
// POSITION SELECTOR
// ============================================================================

/**
 * A position selector returned by doc.at() and doc.after().
 * Chains operations that apply at a specific location in the document.
 *
 * Like CSS selectors for documents:
 *   doc.after("Methods").insert("text")
 *   doc.at("some phrase").comment("note")
 */
class PositionSelector {
  constructor(engine, anchor, mode) {
    this._engine = engine;
    this._anchor = anchor;
    this._mode = mode; // 'at' or 'after'
  }

  /** Insert a new paragraph at this position */
  insert(text, opts = {}) {
    return this._engine._insertAt(this._anchor, this._mode, text, opts);
  }

  /** Add a comment anchored to text at this position */
  comment(text, opts = {}) {
    return this._engine._commentAt(this._anchor, text, opts);
  }

  /** Insert a figure at this position */
  figure(imagePath, caption, opts = {}) {
    if (typeof caption === 'object') { opts = caption; caption = opts.caption; }
    return this._engine._figureAt(this._anchor, this._mode, imagePath, caption, opts);
  }

  /** Insert a table at this position */
  table(data, opts = {}) {
    return this._engine._tableAt(this._anchor, this._mode, data, opts);
  }

  /** Reply to a comment at this position (finds comment by anchor text) */
  reply(text, opts = {}) {
    return this._engine._replyAt(this._anchor, text, opts);
  }
}

// ============================================================================
// DOCEX ENGINE
// ============================================================================

class DocexEngine {
  constructor(docxPath) {
    this._docxPath = path.resolve(docxPath);
    this._workspace = null;
    this._author = 'Unknown';
    this._date = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
    this._tracked = true; // tracked changes ON by default (this is an editing tool)
    this._operations = []; // queued operations for single-pass execution
    this._paragraphs = null;
    this._comments = null;
    this._figures = null;
    this._tables = null;
  }

  // ── Configuration ────────────────────────────────────────────────────────

  /** Set the author name for all subsequent operations */
  author(name) {
    this._author = name;
    return this;
  }

  /** Set the date for all subsequent operations (default: now) */
  date(isoDate) {
    this._date = isoDate;
    return this;
  }

  /** Disable tracked changes (edits become direct modifications) */
  untracked() {
    this._tracked = false;
    return this;
  }

  /** Re-enable tracked changes */
  tracked() {
    this._tracked = true;
    return this;
  }

  // ── Position Selectors ───────────────────────────────────────────────────

  /**
   * Select a position AT the given text (for anchoring comments, etc.)
   * @param {string} text - Text to anchor to
   * @returns {PositionSelector}
   *
   * Example: doc.at("platform regulation").comment("Needs citation")
   */
  at(text) {
    return new PositionSelector(this, text, 'at');
  }

  /**
   * Select a position AFTER the given text or heading
   * @param {string} text - Text or heading to position after
   * @returns {PositionSelector}
   *
   * Example: doc.after("Methods").insert("New methodology paragraph.")
   */
  after(text) {
    return new PositionSelector(this, text, 'after');
  }

  /**
   * Select a position BEFORE the given text or heading
   * @param {string} text - Text or heading to position before
   * @returns {PositionSelector}
   */
  before(text) {
    return new PositionSelector(this, text, 'before');
  }

  // ── Direct Operations (apply to whole document) ──────────────────────────

  /**
   * Replace text anywhere in the document.
   * Tracked by default (shows as strikethrough + insertion).
   *
   * Like LaTeX's \replaced{old}{new}
   *
   * @param {string} oldText - Text to find
   * @param {string} newText - Replacement text
   * @param {object} [opts] - Options
   * @param {string} [opts.author] - Override author for this operation
   * @param {boolean} [opts.tracked] - Override tracked setting
   */
  replace(oldText, newText, opts = {}) {
    this._operations.push({
      type: 'replace',
      oldText,
      newText,
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : this._tracked,
      date: this._date,
    });
    return this;
  }

  /**
   * Delete text from the document.
   * Tracked by default (shows as strikethrough).
   *
   * @param {string} text - Text to delete
   * @param {object} [opts] - Options
   */
  delete(text, opts = {}) {
    this._operations.push({
      type: 'delete',
      text,
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : this._tracked,
      date: this._date,
    });
    return this;
  }

  /**
   * Add a comment to text in the document.
   *
   * Like LaTeX's \marginpar{note}
   *
   * @param {string} anchor - Text to anchor the comment to
   * @param {string} text - Comment text
   * @param {object} [opts] - Options
   * @param {string} [opts.by] - Comment author (overrides doc author)
   * @param {string} [opts.initials] - Author initials
   */
  comment(anchor, text, opts = {}) {
    this._operations.push({
      type: 'comment',
      anchor,
      text,
      author: opts.by || opts.author || this._author,
      initials: opts.initials || (opts.by || opts.author || this._author).split(' ').map(w => w[0]).join(''),
      date: this._date,
    });
    return this;
  }

  // ── Internal operation executors ─────────────────────────────────────────

  _insertAt(anchor, mode, text, opts) {
    this._operations.push({
      type: 'insert',
      anchor,
      mode,
      text,
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : this._tracked,
      date: this._date,
    });
    return this;
  }

  _commentAt(anchor, text, opts) {
    this._operations.push({
      type: 'comment',
      anchor,
      text,
      author: opts.by || opts.author || this._author,
      initials: opts.initials || (opts.by || opts.author || this._author).split(' ').map(w => w[0]).join(''),
      date: this._date,
    });
    return this;
  }

  _figureAt(anchor, mode, imagePath, caption, opts) {
    this._operations.push({
      type: 'figure',
      anchor,
      mode,
      imagePath: path.resolve(imagePath),
      caption,
      width: opts.width || 6, // inches
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : this._tracked,
      date: this._date,
    });
    return this;
  }

  _tableAt(anchor, mode, data, opts) {
    this._operations.push({
      type: 'table',
      anchor,
      mode,
      data,
      caption: opts.caption,
      headers: opts.headers !== false,
      style: opts.style || 'booktabs',
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : this._tracked,
      date: this._date,
    });
    return this;
  }

  _replyAt(anchor, text, opts) {
    this._operations.push({
      type: 'reply',
      anchor,
      text,
      author: opts.by || opts.author || this._author,
      date: this._date,
    });
    return this;
  }

  // ── Inspection (read-only, no save needed) ───────────────────────────────

  /**
   * List all paragraphs with their text
   * @returns {Array<{index: number, text: string, style: string}>}
   */
  async paragraphs() {
    const ws = await this._ensureWorkspace();
    return Paragraphs.list(ws);
  }

  /**
   * List all comments
   * @returns {Array<{id: number, author: string, text: string, date: string}>}
   */
  async comments() {
    const ws = await this._ensureWorkspace();
    return Comments.list(ws);
  }

  /**
   * List all images/figures
   * @returns {Array<{rId: string, filename: string, width: number, height: number}>}
   */
  async figures() {
    const ws = await this._ensureWorkspace();
    return Figures.list(ws);
  }

  /**
   * List all headings
   * @returns {Array<{level: number, text: string, index: number}>}
   */
  async headings() {
    const ws = await this._ensureWorkspace();
    return Paragraphs.headings(ws);
  }

  /**
   * Get the full text of the document
   * @returns {string}
   */
  async text() {
    const ws = await this._ensureWorkspace();
    return Paragraphs.fullText(ws);
  }

  /**
   * List all citation patterns found in the document text.
   * No network access required -- just pattern matching.
   *
   * @returns {Array<{text: string, paragraph: number, pattern: string, authors: string, year: string}>}
   */
  async citations() {
    const ws = await this._ensureWorkspace();
    return Citations.list(ws);
  }

  /**
   * Inject Zotero citation field codes into the document.
   * Finds plain-text citation patterns, queries the Zotero Web API,
   * and replaces them with OOXML ZOTERO_CITATION field codes.
   *
   * @param {object} options
   * @param {string} options.zoteroApiKey - Zotero API key
   * @param {string} options.zoteroUserId - Zotero user ID
   * @param {string} [options.collectionId] - Limit matching to this Zotero collection
   * @param {boolean} [options.bibliography] - Insert bibliography field (default: true)
   * @returns {Promise<{found: number, matched: number, injected: number, unmatched: Array<string>}>}
   */
  async injectCitations(options) {
    const ws = await this._ensureWorkspace();
    return Citations.inject(ws, options);
  }

  // ── Export ──────────────────────────────────────────────────────────────

  /**
   * Convert the document to LaTeX.
   * Read-only export -- does not modify the document or require save().
   *
   * @param {object} [options] - Conversion options
   * @param {string} [options.documentClass='article'] - LaTeX document class
   * @param {string[]} [options.packages] - Additional LaTeX packages
   * @param {string} [options.bibFile='references'] - Bibliography file name (without .bib)
   * @returns {Promise<string>} Complete LaTeX document string
   *
   * @example
   *   const doc = docex("manuscript.docx");
   *   const tex = await doc.toLatex();
   *   // tex is a full LaTeX document string
   */
  async toLatex(options) {
    const ws = await this._ensureWorkspace();
    return Latex.convert(ws, options);
  }

  // ── Lifecycle ────────────────────────────────────────────────────────────

  /**
   * Execute all queued operations and save the document.
   * Single pass: unzip once, apply all operations, rezip once, verify.
   *
   * @param {string|object} [outputPathOrOpts] - Save path string, or options object:
   *   - outputPath {string}    Save to a different path (default: overwrite original)
   *   - safeModify {string}    Path to safe-modify.sh script for manuscript protection
   *   - description {string}   Description for safe-modify.sh commit message
   * @returns {object} - { path, operations, paragraphCount, fileSize, verified }
   *
   * @example
   *   // Simple save
   *   await doc.save();
   *   await doc.save("output.docx");
   *
   *   // Safe-modify save (wraps through safe-modify.sh)
   *   await doc.save({ safeModify: "/path/to/safe-modify.sh", description: "Fix typo" });
   */
  async save(outputPathOrOpts) {
    const ws = await this._ensureWorkspace();

    // Normalize arguments: string -> { outputPath }, object -> pass through
    let saveOpts;
    if (typeof outputPathOrOpts === 'object' && outputPathOrOpts !== null) {
      saveOpts = outputPathOrOpts;
    } else {
      // String or undefined -- original behavior
      saveOpts = outputPathOrOpts || undefined;
    }

    const target = (typeof saveOpts === 'object' && saveOpts !== null)
      ? (saveOpts.outputPath ? path.resolve(saveOpts.outputPath) : this._docxPath)
      : (saveOpts ? path.resolve(saveOpts) : this._docxPath);

    console.log(`[docex] Applying ${this._operations.length} operations...`);

    // Apply operations in order (forward for comments/figures, backward index not needed
    // because we operate on XML strings with text search, not character offsets)
    let opCount = { replace: 0, insert: 0, delete: 0, comment: 0, figure: 0, table: 0, reply: 0 };

    for (const op of this._operations) {
      try {
        switch (op.type) {
          case 'replace':
            Paragraphs.replace(ws, op.oldText, op.newText, op);
            opCount.replace++;
            break;
          case 'insert':
            Paragraphs.insert(ws, op.anchor, op.mode, op.text, op);
            opCount.insert++;
            break;
          case 'delete':
            Paragraphs.remove(ws, op.text, op);
            opCount.delete++;
            break;
          case 'comment':
            Comments.add(ws, op.anchor, op.text, op);
            opCount.comment++;
            break;
          case 'reply':
            Comments.reply(ws, op.anchor, op.text, op);
            opCount.reply++;
            break;
          case 'figure':
            Figures.insert(ws, op.anchor, op.mode, op.imagePath, op.caption, op);
            opCount.figure++;
            break;
          case 'table':
            Tables.insert(ws, op.anchor, op.mode, op.data, op);
            opCount.table++;
            break;
        }
      } catch (err) {
        console.error(`[docex] WARN: ${op.type} operation failed: ${err.message}`);
        console.error(`[docex]   anchor/text: "${(op.oldText || op.anchor || op.text || '').slice(0, 50)}"`);
      }
    }

    // Save and verify -- pass through options to workspace
    const result = ws.save(saveOpts);

    const summary = Object.entries(opCount).filter(([, v]) => v > 0).map(([k, v]) => `${v} ${k}`).join(', ');
    console.log(`[docex] Done: ${summary}`);
    console.log(`[docex] Output: ${result.path} (${result.fileSize} bytes, ${result.paragraphCount} paragraphs)`);
    if (result.verified) {
      console.log(`[docex] Verified: valid zip, paragraph count OK, size OK`);
    }

    // Clear operations after save
    this._operations = [];
    return result;
  }

  /**
   * Discard all pending operations without saving
   */
  discard() {
    this._operations = [];
    if (this._workspace) {
      this._workspace.cleanup();
      this._workspace = null;
    }
    return this;
  }

  // ── Internal ─────────────────────────────────────────────────────────────

  async _ensureWorkspace() {
    if (!this._workspace) {
      this._workspace = Workspace.open(this._docxPath);
    }
    return this._workspace;
  }
}

// ============================================================================
// FACTORY FUNCTION
// ============================================================================

/**
 * Open a document for editing.
 *
 * @param {string} docxPath - Path to the .docx file
 * @returns {DocexEngine}
 *
 * @example
 *   const doc = docex("manuscript.docx");
 *   doc.author("Fabio Votta");
 *   doc.replace("old", "new");
 *   await doc.save();
 */
function docex(docxPath) {
  return new DocexEngine(docxPath);
}

// Also expose as docex.open() for people who prefer that style
docex.open = function(docxPath) {
  return new DocexEngine(docxPath);
};

module.exports = docex;
module.exports.DocexEngine = DocexEngine;
module.exports.PositionSelector = PositionSelector;
