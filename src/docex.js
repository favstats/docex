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
const { Revisions } = require('./revisions');
const { Formatting } = require('./formatting');
const { Footnotes } = require('./footnotes');
const { Metadata } = require('./metadata');
const { Diff } = require('./diff');
const { DocMap } = require('./docmap');
const { ParagraphHandle } = require('./handle');
const { Doctor } = require('./doctor');
const { CrossRef } = require('./crossref');
const { Lists } = require('./lists');
const { Macros } = require('./macros');
const { Presets } = require('./presets');
const { Verify } = require('./verify');
const { Submission } = require('./submission');
const { Compile } = require('./compile');
const { Batch } = require('./batch');
const { Template, createEmpty } = require('./template');
const { ResponseLetter } = require('./response-letter');
const { Layout } = require('./layout');
const { Provenance } = require('./provenance');
const { Workflow } = require('./workflow');
const { Sections } = require('./sections');
const { Redact } = require('./redact');
const { Quality } = require('./quality');
const { Production } = require('./production');
const { Transaction } = require('./transaction');
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

  // ── Formatting at position ──────────────────────────────────────────────

  /** Make the anchor text bold */
  bold(opts) { return this._engine._format('bold', this._anchor, opts); }

  /** Make the anchor text italic */
  italic(opts) { return this._engine._format('italic', this._anchor, opts); }

  /** Underline the anchor text */
  underline(opts) { return this._engine._format('underline', this._anchor, opts); }

  /** Strikethrough the anchor text */
  strikethrough(opts) { return this._engine._format('strikethrough', this._anchor, opts); }

  /** Make the anchor text superscript */
  superscript(opts) { return this._engine._format('superscript', this._anchor, opts); }

  /** Make the anchor text subscript */
  subscript(opts) { return this._engine._format('subscript', this._anchor, opts); }

  /** Make the anchor text small caps */
  smallCaps(opts) { return this._engine._format('smallCaps', this._anchor, opts); }

  /** Make the anchor text monospace (code) */
  code(opts) { return this._engine._format('code', this._anchor, opts); }

  /** Set font color on the anchor text */
  color(colorName, opts) { return this._engine._format('color', this._anchor, { ...opts, colorName }); }

  /** Highlight the anchor text */
  highlight(colorName, opts) { return this._engine._format('highlight', this._anchor, { ...opts, colorName }); }

  // ── Footnote at position ────────────────────────────────────────────────

  /** Add a footnote anchored to text at this position */
  footnote(text, opts) { return this._engine._footnoteAt(this._anchor, text, opts); }

  // ── Lists at position ─────────────────────────────────────────────────

  /** Insert a bullet list at this position */
  bulletList(items, opts) { return this._engine._bulletListAt(this._anchor, this._mode, items, opts); }

  /** Insert a numbered list at this position */
  numberedList(items, opts) { return this._engine._numberedListAt(this._anchor, this._mode, items, opts); }
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
    this._rc = null; // .docexrc config

    // Load .docexrc from the directory containing the docx file
    this._loadRc();
  }

  /**
   * Load .docexrc configuration from the docx file's directory,
   * falling back to ~/.docexrc for global defaults.
   * @private
   */
  _loadRc() {
    const localRcPath = path.join(path.dirname(this._docxPath), '.docexrc');
    const homeRcPath = path.join(require('os').homedir(), '.docexrc');

    let globalRc = {};
    let localRc = {};

    try {
      if (fs.existsSync(homeRcPath)) {
        globalRc = JSON.parse(fs.readFileSync(homeRcPath, 'utf-8'));
      }
    } catch (_) { /* ignore */ }

    try {
      if (fs.existsSync(localRcPath)) {
        localRc = JSON.parse(fs.readFileSync(localRcPath, 'utf-8'));
      }
    } catch (_) { /* ignore */ }

    this._rc = { ...globalRc, ...localRc };

    // Apply rc defaults
    if (this._rc.author) {
      this._author = this._rc.author;
    }
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

  // ── Direct Formatting (apply to first match in document) ────────────────

  /**
   * Make text bold anywhere in the document.
   * @param {string} text - Text to make bold
   * @param {object} [opts] - Options (tracked, author)
   */
  bold(text, opts = {}) {
    this._operations.push({
      type: 'format',
      formatType: 'bold',
      text,
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : false,
      date: this._date,
    });
    return this;
  }

  /**
   * Make text italic anywhere in the document.
   * @param {string} text - Text to make italic
   * @param {object} [opts] - Options (tracked, author)
   */
  italic(text, opts = {}) {
    this._operations.push({
      type: 'format',
      formatType: 'italic',
      text,
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : false,
      date: this._date,
    });
    return this;
  }

  /**
   * Highlight text anywhere in the document.
   * @param {string} text - Text to highlight
   * @param {string} color - Highlight color name (e.g. 'yellow')
   * @param {object} [opts] - Options (tracked, author)
   */
  highlight(text, color, opts = {}) {
    this._operations.push({
      type: 'format',
      formatType: 'highlight',
      text,
      colorName: color || 'yellow',
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : false,
      date: this._date,
    });
    return this;
  }

  /**
   * Set font color on text anywhere in the document.
   * @param {string} text - Text to color
   * @param {string} colorName - Named color or hex
   * @param {object} [opts] - Options (tracked, author)
   */
  color(text, colorName, opts = {}) {
    this._operations.push({
      type: 'format',
      formatType: 'color',
      text,
      colorName: colorName || 'red',
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : false,
      date: this._date,
    });
    return this;
  }

  /**
   * Replace all occurrences of text in the document.
   * Unlike replace() which only replaces the first match, this replaces all.
   *
   * @param {string} oldText - Text to find
   * @param {string} newText - Replacement text
   * @param {object} [opts] - Options
   */
  replaceAll(oldText, newText, opts = {}) {
    this._operations.push({
      type: 'replaceAll',
      oldText,
      newText,
      author: opts.author || this._author,
      tracked: opts.tracked !== undefined ? opts.tracked : this._tracked,
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

  _format(formatType, text, opts = {}) {
    this._operations.push({
      type: 'format',
      formatType,
      text,
      colorName: opts && opts.colorName,
      author: (opts && opts.author) || this._author,
      tracked: opts && opts.tracked !== undefined ? opts.tracked : false,
      date: this._date,
    });
    return this;
  }

  _footnoteAt(anchor, text, opts = {}) {
    this._operations.push({
      type: 'footnote',
      anchor,
      text,
      author: (opts && opts.author) || this._author,
      date: this._date,
    });
    return this;
  }

  _bulletListAt(anchor, mode, items, opts = {}) {
    this._operations.push({
      type: 'bulletList',
      anchor,
      mode,
      items,
      author: (opts && opts.author) || this._author,
      tracked: opts && opts.tracked !== undefined ? opts.tracked : false,
      date: this._date,
    });
    return this;
  }

  _numberedListAt(anchor, mode, items, opts = {}) {
    this._operations.push({
      type: 'numberedList',
      anchor,
      mode,
      items,
      author: (opts && opts.author) || this._author,
      tracked: opts && opts.tracked !== undefined ? opts.tracked : false,
      date: this._date,
    });
    return this;
  }

  // ── Stable Addressing (v0.3) ────────────────────────────────────────────

  /**
   * Generate a structured map of the document.
   * Returns sections, paragraphs, figures, tables, and comments.
   *
   * @returns {Promise<{sections: Array, allParagraphs: Array, allFigures: Array, allTables: Array, allComments: Array}>}
   */
  async map() {
    const ws = await this._ensureWorkspace();
    return DocMap.generate(ws);
  }

  /**
   * Get a ParagraphHandle for a specific paragraph by its w14:paraId.
   * The handle provides stable, chainable operations.
   *
   * @param {string} paraId - The w14:paraId of the target paragraph
   * @returns {ParagraphHandle}
   */
  id(paraId) {
    return new ParagraphHandle(this, paraId);
  }

  table(n) {
    var TH = require('./table-handle').TableHandle;
    return new TH(this, n);
  }

  figure(n) {
    var FH = require('./figure-handle').FigureHandle;
    return new FH(this, n);
  }

  /**
   * Find a heading by text and return a PositionSelector.
   * Only matches heading paragraphs.
   *
   * @param {string} text - Heading text to match
   * @returns {PositionSelector}
   */
  afterHeading(text) {
    return new PositionSelector(this, text, 'after');
  }

  /**
   * Find a body paragraph by text and return a PositionSelector.
   * Only matches body paragraphs (not headings).
   *
   * @param {string} text - Body text to match
   * @returns {PositionSelector}
   */
  afterText(text) {
    return new PositionSelector(this, text, 'after');
  }

  /**
   * Find paragraphs containing the given text.
   * Returns matches with section context and surrounding text.
   *
   * @param {string} text - Text to search for
   * @returns {Promise<Array<{id: string, index: number, section: string, context: string}>>}
   */
  async find(text) {
    const ws = await this._ensureWorkspace();
    return DocMap.find(ws, text);
  }

  /**
   * Return a tree-view string of the document structure.
   *
   * @returns {Promise<string>}
   */
  async structure() {
    const ws = await this._ensureWorkspace();
    return DocMap.structure(ws);
  }

  /**
   * Find text and show the XML structure around it.
   * Useful for debugging.
   *
   * @param {string} text - Text to find and explain
   * @returns {Promise<string>}
   */
  async explain(text) {
    const ws = await this._ensureWorkspace();
    return DocMap.explain(ws, text);
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

  // ── Revisions ────────────────────────────────────────────────────────────

  /**
   * List all tracked changes in the document.
   * @returns {Array<{id: number, type: string, author: string, date: string, text: string}>}
   */
  async revisions() {
    return Revisions.list(await this._ensureWorkspace());
  }

  /**
   * Accept tracked changes. If id is provided, accept only that change.
   * If id is undefined, accept ALL changes.
   * @param {number} [id] - Specific change ID to accept
   */
  async accept(id) {
    Revisions.accept(await this._ensureWorkspace(), id);
    return this;
  }

  /**
   * Reject tracked changes. If id is provided, reject only that change.
   * If id is undefined, reject ALL changes.
   * @param {number} [id] - Specific change ID to reject
   */
  async reject(id) {
    Revisions.reject(await this._ensureWorkspace(), id);
    return this;
  }

  /**
   * Produce a clean copy: accept all changes, remove all comments.
   */
  async cleanCopy() {
    Revisions.cleanCopy(await this._ensureWorkspace());
    return this;
  }

  // ── Footnotes ───────────────────────────────────────────────────────────

  /**
   * List all footnotes in the document.
   * @returns {Array<{id: number, text: string}>}
   */
  async footnotes() {
    return Footnotes.list(await this._ensureWorkspace());
  }

  // ── Metadata ────────────────────────────────────────────────────────────

  /**
   * Get or set document metadata.
   * If props is provided, sets metadata and returns this for chaining.
   * If props is omitted, returns the metadata object.
   *
   * @param {object} [props] - Properties to set (title, creator, keywords, etc.)
   * @returns {Promise<object|DocexEngine>}
   */
  async metadata(props) {
    const ws = await this._ensureWorkspace();
    if (props) {
      Metadata.set(ws, props);
      return this;
    }
    return Metadata.get(ws);
  }

  // ── Validation ─────────────────────────────────────────────────────────

  /**
   * Validate the document's integrity.
   * Alias for Doctor.validate(). Does not modify the document.
   *
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[]}>}
   */
  async validate() {
    const ws = await this._ensureWorkspace();
    return Doctor.validate(ws);
  }

  // ── Word Count ──────────────────────────────────────────────────────────

  /**
   * Count words in the document, categorized by type.
   * @returns {{total: number, body: number, headings: number, abstract: number, captions: number, footnotes: number}}
   */
  async wordCount() {
    return Paragraphs.wordCount(await this._ensureWorkspace());
  }

  // ── Stats ──────────────────────────────────────────────────────────────

  /**
   * Comprehensive document statistics.
   * Combines word count, paragraph count, figure count, table count,
   * citation count, comment count, and revision count.
   *
   * @returns {Promise<{words: object, paragraphs: number, headings: number, figures: number, tables: number, citations: number, comments: number, revisions: number, pages: null}>}
   */
  async stats() {
    const ws = await this._ensureWorkspace();

    const words = Paragraphs.wordCount(ws);
    const paraList = Paragraphs.list(ws);
    const headingList = Paragraphs.headings(ws);
    const figureList = Figures.list(ws);

    // Count tables in document
    const docXml = ws.docXml;
    const tableCount = (docXml.match(/<w:tbl[\s>]/g) || []).length;

    const citationList = Citations.list(ws);
    const commentList = Comments.list(ws);
    const revisionList = Revisions.list(ws);

    return {
      words,
      paragraphs: paraList.length,
      headings: headingList.length,
      figures: figureList.length,
      tables: tableCount,
      citations: citationList.length,
      comments: commentList.length,
      revisions: revisionList.length,
      pages: null, // requires PDF rendering, not available
    };
  }

  // ── Contributors ───────────────────────────────────────────────────────

  /**
   * Scan all tracked changes and comments to find unique contributors.
   * Returns authors with change/comment counts and last-active dates.
   *
   * @returns {Promise<Array<{name: string, changes: number, comments: number, lastActive: string}>>}
   */
  async contributors() {
    const ws = await this._ensureWorkspace();
    return Revisions.contributors(ws);
  }

  // ── Timeline ───────────────────────────────────────────────────────────

  /**
   * Combine comment dates and revision dates into a single
   * chronological timeline.
   *
   * @returns {Promise<Array<{date: string, type: string, author: string, text: string}>>}
   */
  async timeline() {
    const ws = await this._ensureWorkspace();
    return Revisions.timeline(ws);
  }

  // ── Export Comments ────────────────────────────────────────────────────

  /**
   * Export all comments as CSV or JSON string.
   *
   * @param {string} [format='json'] - 'csv' or 'json'
   * @returns {Promise<string>}
   */
  async exportComments(format) {
    const ws = await this._ensureWorkspace();
    return Comments.exportComments(ws, format);
  }

  // ── Diff ────────────────────────────────────────────────────────────────

  /**
   * Compare this document with another document.
   * Produces tracked changes (w:del + w:ins) showing what changed.
   *
   * @param {string} otherDocxPath - Path to the other .docx file
   * @param {object} [opts] - Options
   * @param {string} [opts.author] - Author name for tracked changes
   * @returns {{added: number, removed: number, modified: number, unchanged: number}}
   */
  async diff(otherDocxPath, opts = {}) {
    const ws = await this._ensureWorkspace();
    const ws2 = Workspace.open(otherDocxPath);
    const result = Diff.compare(ws, ws2, {
      author: opts.author || this._author,
      date: this._date,
    });
    ws2.cleanup();
    return result;
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

  // ── Pandoc-based Export ────────────────────────────────────────────────

  /**
   * Convert the document to HTML via pandoc.
   * Requires pandoc to be installed.
   *
   * @param {object} [opts] - Options
   * @param {string} [opts.output] - Write HTML to this file path
   * @returns {Promise<string>} HTML string
   */
  async toHtml(opts = {}) {
    await this._ensureWorkspace();
    const tmpDocx = path.join('/tmp', `docex-html-${crypto.randomBytes(8).toString('hex')}.docx`);
    try {
      // Open a separate workspace copy so the active one is not consumed by save()
      const exportWs = Workspace.open(this._docxPath);
      // Copy any in-memory modifications from the active workspace
      exportWs.docXml = this._workspace.docXml;
      try { exportWs.commentsXml = this._workspace.commentsXml; } catch (_) { /* may not exist */ }
      try { exportWs.commentsExtXml = this._workspace.commentsExtXml; } catch (_) { /* may not exist */ }
      try { exportWs.footnotesXml = this._workspace.footnotesXml; } catch (_) { /* may not exist */ }
      exportWs.save({ outputPath: tmpDocx, backup: false });

      const result = execFileSync('pandoc', [tmpDocx, '--from', 'docx', '--to', 'html5', '--standalone'], {
        encoding: 'utf-8',
        stdio: ['pipe', 'pipe', 'pipe'],
        timeout: 60000,
      });

      if (opts.output) {
        fs.writeFileSync(path.resolve(opts.output), result, 'utf-8');
      }

      return result;
    } finally {
      try { if (fs.existsSync(tmpDocx)) fs.unlinkSync(tmpDocx); } catch (_) { /* ignore */ }
    }
  }

  /**
   * Convert the document to Markdown via pandoc.
   * Requires pandoc to be installed.
   *
   * @param {object} [opts] - Options
   * @param {string} [opts.output] - Write Markdown to this file path
   * @returns {Promise<string>} Markdown string
   */
  async toMarkdown(opts = {}) {
    await this._ensureWorkspace();
    const tmpDocx = path.join('/tmp', `docex-md-${crypto.randomBytes(8).toString('hex')}.docx`);
    try {
      // Open a separate workspace copy so the active one is not consumed by save()
      const exportWs = Workspace.open(this._docxPath);
      exportWs.docXml = this._workspace.docXml;
      try { exportWs.commentsXml = this._workspace.commentsXml; } catch (_) { /* may not exist */ }
      try { exportWs.commentsExtXml = this._workspace.commentsExtXml; } catch (_) { /* may not exist */ }
      try { exportWs.footnotesXml = this._workspace.footnotesXml; } catch (_) { /* may not exist */ }
      exportWs.save({ outputPath: tmpDocx, backup: false });

      const result = execFileSync('pandoc', [tmpDocx, '--from', 'docx', '--to', 'markdown'], {
        encoding: 'utf-8',
        stdio: ['pipe', 'pipe', 'pipe'],
        timeout: 60000,
      });

      if (opts.output) {
        fs.writeFileSync(path.resolve(opts.output), result, 'utf-8');
      }

      return result;
    } finally {
      try { if (fs.existsSync(tmpDocx)) fs.unlinkSync(tmpDocx); } catch (_) { /* ignore */ }
    }
  }

  // ── Preview ─────────────────────────────────────────────────────────────

  /**
   * Format pending operations as a human-readable summary string.
   * Does not execute the operations or modify the document.
   *
   * @returns {string} Human-readable summary of pending operations
   *
   * @example
   *   doc.replace("old", "new");
   *   doc.after("Methods").insert("text");
   *   console.log(doc.preview());
   *   // "2 pending operations:
   *   //   1. replace 'old' -> 'new' (tracked, by Fabio Votta)
   *   //   2. insert after 'Methods': 'text' (tracked, by Fabio Votta)"
   */
  preview() {
    if (this._operations.length === 0) {
      return 'No pending operations.';
    }

    const lines = [`${this._operations.length} pending operation${this._operations.length !== 1 ? 's' : ''}:`];

    for (let i = 0; i < this._operations.length; i++) {
      const op = this._operations[i];
      const num = i + 1;
      const trackedStr = op.tracked !== undefined
        ? (op.tracked ? 'tracked' : 'untracked')
        : '';
      const authorStr = op.author ? `by ${op.author}` : '';
      const meta = [trackedStr, authorStr].filter(Boolean).join(', ');
      const metaSuffix = meta ? ` (${meta})` : '';

      let desc;
      switch (op.type) {
        case 'replace':
          desc = `replace '${DocexEngine._truncate(op.oldText, 30)}' -> '${DocexEngine._truncate(op.newText, 30)}'${metaSuffix}`;
          break;
        case 'replaceAll':
          desc = `replaceAll '${DocexEngine._truncate(op.oldText, 30)}' -> '${DocexEngine._truncate(op.newText, 30)}'${metaSuffix}`;
          break;
        case 'insert':
          desc = `insert ${op.mode} '${DocexEngine._truncate(op.anchor, 30)}': '${DocexEngine._truncate(op.text, 30)}'${metaSuffix}`;
          break;
        case 'delete':
          desc = `delete '${DocexEngine._truncate(op.text, 30)}'${metaSuffix}`;
          break;
        case 'comment':
          desc = `comment at '${DocexEngine._truncate(op.anchor, 30)}': '${DocexEngine._truncate(op.text, 30)}'${metaSuffix}`;
          break;
        case 'reply':
          desc = `reply at '${DocexEngine._truncate(op.anchor, 30)}': '${DocexEngine._truncate(op.text, 30)}'${metaSuffix}`;
          break;
        case 'figure':
          desc = `figure ${op.mode} '${DocexEngine._truncate(op.anchor, 30)}'${metaSuffix}`;
          break;
        case 'table':
          desc = `table ${op.mode} '${DocexEngine._truncate(op.anchor, 30)}'${metaSuffix}`;
          break;
        case 'format':
          desc = `${op.formatType} '${DocexEngine._truncate(op.text, 30)}'${metaSuffix}`;
          break;
        case 'footnote':
          desc = `footnote at '${DocexEngine._truncate(op.anchor, 30)}': '${DocexEngine._truncate(op.text, 30)}'${metaSuffix}`;
          break;
        default:
          desc = `${op.type}${metaSuffix}`;
      }

      lines.push(`  ${num}. ${desc}`);
    }

    return lines.join('\n');
  }

  // ── Transactions ────────────────────────────────────────────────────────

  /**
   * Create a new transaction for atomic multi-operation edits.
   * All operations are queued and applied atomically on commit.
   * On failure, the document is rolled back to its pre-transaction state.
   *
   * @returns {Transaction}
   */
  transaction() {
    return new Transaction(this);
  }

  /** @private */
  static _truncate(str, max) {
    if (!str) return '';
    if (str.length <= max) return str;
    return str.slice(0, max) + '...';
  }

  // ── RC Config ──────────────────────────────────────────────────────────

  /**
   * Get the loaded .docexrc configuration.
   * @returns {object} The merged .docexrc config
   */
  get rc() {
    return this._rc || {};
  }

  // ── Cross-References & Auto-Numbering (v0.3) ──────────────────────────

  /**
   * Assign a label to a paragraph for cross-referencing.
   * @param {string} paraId - The w14:paraId of the paragraph
   * @param {string} labelName - Label name, e.g. "fig:funnel"
   */
  async label(paraId, labelName) {
    const ws = await this._ensureWorkspace();
    CrossRef.label(ws, paraId, labelName);
    return this;
  }

  /**
   * Insert a cross-reference to a labeled element.
   * @param {string} labelName - Label to reference
   * @param {object} [opts] - { insertAt: paraId, after: "text" }
   */
  async ref(labelName, opts) {
    const ws = await this._ensureWorkspace();
    CrossRef.ref(ws, labelName, opts);
    return this;
  }

  /**
   * Auto-number all figure and table captions using SEQ field codes.
   * @returns {{figures: number, tables: number}}
   */
  async autoNumber() {
    const ws = await this._ensureWorkspace();
    return CrossRef.autoNumber(ws);
  }

  /**
   * List all labels in the document.
   * @returns {Array<{name: string, type: string, number: number|null, paraId: string|null}>}
   */
  async listLabels() {
    const ws = await this._ensureWorkspace();
    return CrossRef.listLabels(ws);
  }

  // ── Variables & Macros (v0.3) ──────────────────────────────────────────

  /**
   * Define a variable for later expansion.
   * @param {string} name - Variable name (e.g. "NUM_ADS")
   * @param {string} value - Variable value (e.g. "268,635")
   */
  async define(name, value) {
    const ws = await this._ensureWorkspace();
    Macros.define(ws, name, value);
    return this;
  }

  /**
   * Expand all {{VAR_NAME}} patterns in the document.
   * @param {object} [variables] - Map of variable names to values
   * @returns {number} Count of expansions
   */
  async expand(variables) {
    const ws = await this._ensureWorkspace();
    return Macros.expand(ws, variables);
  }

  /**
   * List all {{VAR}} patterns found in the document.
   * @returns {Array<{name: string, paragraph: number, context: string}>}
   */
  async listVariables() {
    const ws = await this._ensureWorkspace();
    return Macros.listVariables(ws);
  }

  // ── Journal Style Presets (v0.3) ──────────────────────────────────────

  /**
   * Apply a journal style preset to the document.
   * @param {string} presetName - e.g. "polcomm", "apa7", "academic"
   * @returns {{applied: string, changes: string[]}}
   */
  async style(presetName) {
    const ws = await this._ensureWorkspace();
    return Presets.apply(ws, presetName);
  }

  async font(name) { Presets.setFont(await this._ensureWorkspace(), name); return this; }
  async fontSize(pt) { Presets.setFontSize(await this._ensureWorkspace(), pt); return this; }
  async headingFont(name) { Presets.setHeadingFont(await this._ensureWorkspace(), name); return this; }
  async headingColor(hex) { Presets.setHeadingColor(await this._ensureWorkspace(), hex); return this; }
  async linkColor(hex) { Presets.setLinkColor(await this._ensureWorkspace(), hex); return this; }
  async paragraphSpacing(opts) { Presets.setParagraphSpacing(await this._ensureWorkspace(), opts); return this; }

  // ── Submission Validation (v0.3) ──────────────────────────────────────

  /**
   * Validate document against journal requirements.
   * @param {string} presetName - e.g. "polcomm", "apa7"
   * @returns {{ pass: boolean, errors: string[], warnings: string[] }}
   */
  async verify(presetName) {
    const ws = await this._ensureWorkspace();
    return Verify.check(ws, presetName);
  }

  // ── Submission Helpers (v0.3) ──────────────────────────────────────────

  /**
   * Remove author names for blind review.
   * @returns {{ authorsRemoved: string[], locations: string[] }}
   */
  async anonymize() {
    const ws = await this._ensureWorkspace();
    return Submission.anonymize(ws);
  }

  /**
   * Restore author info after anonymize().
   * @returns {{ restored: boolean, authors: string[] }}
   */
  async deanonymize() {
    const ws = await this._ensureWorkspace();
    return Submission.deanonymize(ws);
  }

  /**
   * Highlight all tracked changes (insertions yellow, deletions red).
   * @returns {{ insertions: number, deletions: number }}
   */
  async highlightedChanges() {
    const ws = await this._ensureWorkspace();
    return Submission.highlightedChanges(ws);
  }

  // ── Snapshot / Rollback (v0.4.1) ────────────────────────────────────────

  /**
   * Save the current document state in memory.
   * Multiple snapshots stack (LIFO). Use rollback() to restore.
   * Like git stash but in-memory -- no disk I/O.
   *
   * @returns {DocexEngine} this, for chaining
   */
  async snapshot() {
    const ws = await this._ensureWorkspace();
    ws.snapshot();
    return this;
  }

  /**
   * Restore the most recent snapshot, discarding current state.
   * If an operation goes wrong, roll back without touching disk.
   *
   * @returns {boolean} true if a snapshot was restored, false if stack was empty
   */
  async rollback() {
    const ws = await this._ensureWorkspace();
    return ws.rollback();
  }

  // ── Assert (v0.4.1) ──────────────────────────────────────────────────────

  /**
   * Verify a paragraph contains expected text before operating.
   * Fails fast instead of silently editing the wrong paragraph.
   *
   * @param {string} paraId - The w14:paraId of the paragraph to check
   * @param {string} expectedText - Text the paragraph should contain
   * @throws {Error} Descriptive error with actual text if assertion fails
   * @returns {DocexEngine} this, for chaining
   */
  async assert(paraId, expectedText) {
    const ws = await this._ensureWorkspace();
    const result = DocMap.locateById(ws.docXml, paraId);
    if (!result) {
      throw new Error(
        `assert failed: paragraph "${paraId}" not found in document`
      );
    }
    const actualText = xml.extractTextDecoded(result.xml);
    if (!actualText.includes(expectedText)) {
      throw new Error(
        `assert failed: paragraph "${paraId}" does not contain "${expectedText}"\n`
        + `  actual text: "${actualText}"`
      );
    }
    return this;
  }

  // ── Diff Summary (v0.4.1) ────────────────────────────────────────────────

  /**
   * Compare this document with another and return summary counts.
   * "3 paragraphs changed, 1 added, 2 comments added" -- without
   * producing the full tracked-changes diff or modifying either document.
   *
   * @param {string} otherDocxPath - Path to the other .docx file
   * @returns {Promise<{changed: number, added: number, removed: number, comments: number}>}
   */
  async diffSummary(otherDocxPath) {
    const ws1 = await this._ensureWorkspace();
    const ws2 = Workspace.open(otherDocxPath);

    try {
      // Extract paragraphs and texts from both
      const paras1 = xml.findParagraphs(ws1.docXml);
      const paras2 = xml.findParagraphs(ws2.docXml);
      const texts1 = paras1.map(p => xml.decodeXml(p.text));
      const texts2 = paras2.map(p => xml.decodeXml(p.text));

      // Use the Diff module's paragraph alignment
      const operations = Diff._diffParagraphs(texts1, texts2);

      let changed = 0, added = 0, removed = 0;
      for (const op of operations) {
        switch (op.type) {
          case 'modify': changed++; break;
          case 'add': added++; break;
          case 'remove': removed++; break;
          // 'keep' -- no change
        }
      }

      // Count comments difference
      const comments1 = Comments.list(ws1);
      const comments2 = Comments.list(ws2);
      const commentsDiff = Math.abs(comments2.length - comments1.length);

      return { changed, added, removed, comments: commentsDiff };
    } finally {
      ws2.cleanup();
    }
  }

  // ── Document Manipulation (v0.4.8) ──────────────────────────────────────

  /**
   * Replace the content of the nth table in the document.
   * @param {number} tableNumber - 1-based table number
   * @param {Array<Array<string>>} newData - 2D array of cell values
   * @param {object} [opts] - Options { headers, style }
   */
  async replaceTable(tableNumber, newData, opts = {}) {
    const ws = await this._ensureWorkspace();
    Layout.replaceTable(ws, tableNumber, newData, opts);
    return this;
  }

  /**
   * Add a page break before a heading.
   * @param {string} headingText - Text of the heading
   */
  async pageBreakBefore(headingText) {
    const ws = await this._ensureWorkspace();
    Layout.pageBreakBefore(ws, headingText);
    return this;
  }

  /**
   * Fix heading hierarchy skips (e.g. H1->H3 becomes H1->H2).
   * @returns {number} Count of fixes applied
   */
  async ensureHeadingHierarchy() {
    const ws = await this._ensureWorkspace();
    return Layout.ensureHeadingHierarchy(ws);
  }

  /**
   * Merge two consecutive paragraphs into one.
   * @param {string} id1 - w14:paraId of the first paragraph
   * @param {string} id2 - w14:paraId of the second paragraph
   */
  async mergeParagraphs(id1, id2) {
    const ws = await this._ensureWorkspace();
    Layout.mergeParagraphs(ws, id1, id2);
    return this;
  }

  /**
   * Split a paragraph at the given text into two paragraphs.
   * @param {string} paraId - w14:paraId of the paragraph
   * @param {string} atText - Text at which to split
   * @returns {string} New paraId for the second paragraph
   */
  async splitParagraph(paraId, atText) {
    const ws = await this._ensureWorkspace();
    return Layout.splitParagraph(ws, paraId, atText);
  }

  // ── Provenance (v0.4.9) ──────────────────────────────────────────────────

  /**
   * Get the embedded changelog.
   * @returns {Array<{timestamp: string, operation: string, author: string, description: string}>}
   */
  async changelog() {
    const ws = await this._ensureWorkspace();
    return Provenance.getChangelog(ws);
  }

  /**
   * Get changelog entries since a date.
   * @param {string} dateString - ISO date string
   */
  async changelogSince(dateString) {
    const ws = await this._ensureWorkspace();
    return Provenance.changelogSince(ws, dateString);
  }

  /**
   * Get origin information string.
   * @returns {string}
   */
  async origin() {
    const ws = await this._ensureWorkspace();
    return Provenance.origin(ws);
  }

  /**
   * Set origin information.
   * @param {object} info - { version, date, template, tool }
   */
  async setOrigin(info) {
    const ws = await this._ensureWorkspace();
    Provenance.setOrigin(ws, info);
    return this;
  }

  /**
   * Certify the current document state with a label.
   * @param {string} label - e.g. "submitted to Political Communication"
   */
  async certify(label) {
    const ws = await this._ensureWorkspace();
    Provenance.certify(ws, label);
    return this;
  }

  /**
   * Verify if document matches its last certification.
   * @returns {{ certified: boolean, label: string, date: string, hash: string }}
   */
  async verifyCertification() {
    const ws = await this._ensureWorkspace();
    return Provenance.verifyCertification(ws);
  }

  /**
   * List all certification points.
   * @returns {Array<{label: string, date: string, hash: string}>}
   */
  async certifications() {
    const ws = await this._ensureWorkspace();
    return Provenance.certifications(ws);
  }

  // ── Workflow (v0.4.10) ────────────────────────────────────────────────────

  /**
   * Extract all TODO/FIXME items from comments and body text.
   * @returns {Array<{text: string, source: string, paraId: string, author: string}>}
   */
  async todo() {
    const ws = await this._ensureWorkspace();
    return Workflow.todo(ws);
  }

  /**
   * Per-section progress analysis.
   * @returns {Array<{section: string, status: string, wordCount: number, todoCount: number}>}
   */
  async progress() {
    const ws = await this._ensureWorkspace();
    return Workflow.progress(ws);
  }

  /**
   * Render table of contents as a preview string.
   * @returns {string}
   */
  async tocPreview() {
    const ws = await this._ensureWorkspace();
    return Workflow.tocPreview(ws);
  }

  /**
   * List all figures with captions and estimated page numbers.
   * @returns {string}
   */
  async figureList() {
    const ws = await this._ensureWorkspace();
    return Workflow.figureList(ws);
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
   *   - dryRun {boolean}       If true, return result without writing
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
    let opCount = { replace: 0, insert: 0, delete: 0, comment: 0, figure: 0, table: 0, reply: 0, format: 0, footnote: 0, replaceAll: 0 };

    // Operation receipts (v0.4.1): every mutation returns a receipt
    const receipts = [];

    // Helper: find the paraId of the paragraph containing given text
    const _findParaId = (text) => {
      if (!text) return null;
      try {
        const paras = xml.findParagraphs(ws.docXml);
        for (const p of paras) {
          const decoded = xml.extractTextDecoded(p.xml);
          if (decoded.includes(text)) {
            const m = p.xml.match(/w14:paraId="([^"]+)"/);
            return m ? m[1] : null;
          }
        }
      } catch (_) { /* best effort */ }
      return null;
    };

    for (const op of this._operations) {
      const receipt = {
        success: false,
        type: op.type,
        paraId: op.paraId || null,
        matched: null,
        context: (op.oldText || op.anchor || op.text || '').slice(0, 80),
      };

      try {
        switch (op.type) {
          case 'replace':
            Paragraphs.replace(ws, op.oldText, op.newText, op);
            receipt.success = true;
            receipt.matched = op.oldText;
            receipt.paraId = _findParaId(op.newText) || _findParaId(op.oldText);
            opCount.replace++;
            break;
          case 'replaceAll': {
            // Count actual occurrences first to avoid infinite loop when
            // replacement text contains the search text
            const allParas = xml.findParagraphs(ws.docXml);
            let totalOccurrences = 0;
            for (const p of allParas) {
              const decoded = xml.decodeXml(p.text);
              let searchPos = 0;
              while (true) {
                const idx = decoded.indexOf(op.oldText, searchPos);
                if (idx === -1) break;
                totalOccurrences++;
                searchPos = idx + op.oldText.length;
              }
            }
            let count = 0;
            for (let ri = 0; ri < totalOccurrences; ri++) {
              try {
                Paragraphs.replace(ws, op.oldText, op.newText, op);
                count++;
              } catch (_) {
                break; // no more matches
              }
            }
            receipt.success = count > 0;
            receipt.matched = op.oldText;
            receipt.count = count;
            opCount.replaceAll += count;
            break;
          }
          case 'insert':
            Paragraphs.insert(ws, op.anchor, op.mode, op.text, op);
            receipt.success = true;
            receipt.matched = op.anchor;
            receipt.paraId = _findParaId(op.text) || _findParaId(op.anchor);
            opCount.insert++;
            break;
          case 'delete':
            Paragraphs.remove(ws, op.text, op);
            receipt.success = true;
            receipt.matched = op.text;
            opCount.delete++;
            break;
          case 'comment':
            Comments.add(ws, op.anchor, op.text, op);
            receipt.success = true;
            receipt.matched = op.anchor;
            receipt.paraId = _findParaId(op.anchor);
            opCount.comment++;
            break;
          case 'reply':
            Comments.reply(ws, op.anchor, op.text, op);
            receipt.success = true;
            receipt.matched = op.anchor;
            opCount.reply++;
            break;
          case 'figure':
            Figures.insert(ws, op.anchor, op.mode, op.imagePath, op.caption, op);
            receipt.success = true;
            receipt.matched = op.anchor;
            receipt.paraId = _findParaId(op.anchor);
            opCount.figure++;
            break;
          case 'table':
            Tables.insert(ws, op.anchor, op.mode, op.data, op);
            receipt.success = true;
            receipt.matched = op.anchor;
            receipt.paraId = _findParaId(op.anchor);
            opCount.table++;
            break;
          case 'format': {
            const fmtOpts = { tracked: op.tracked, author: op.author, date: op.date };
            switch (op.formatType) {
              case 'bold':
                Formatting.bold(ws, op.text, fmtOpts);
                break;
              case 'italic':
                Formatting.italic(ws, op.text, fmtOpts);
                break;
              case 'underline':
                Formatting.underline(ws, op.text, fmtOpts);
                break;
              case 'strikethrough':
                Formatting.strikethrough(ws, op.text, fmtOpts);
                break;
              case 'superscript':
                Formatting.superscript(ws, op.text, fmtOpts);
                break;
              case 'subscript':
                Formatting.subscript(ws, op.text, fmtOpts);
                break;
              case 'smallCaps':
                Formatting.smallCaps(ws, op.text, fmtOpts);
                break;
              case 'code':
                Formatting.code(ws, op.text, fmtOpts);
                break;
              case 'color':
                Formatting.color(ws, op.text, op.colorName, fmtOpts);
                break;
              case 'highlight':
                Formatting.highlight(ws, op.text, op.colorName || 'yellow', fmtOpts);
                break;
            }
            receipt.success = true;
            receipt.matched = op.text;
            receipt.paraId = _findParaId(op.text);
            opCount.format++;
            break;
          }
          case 'footnote':
            Footnotes.add(ws, op.anchor, op.text, op);
            receipt.success = true;
            receipt.matched = op.anchor;
            receipt.paraId = _findParaId(op.anchor);
            opCount.footnote++;
            break;
          case 'bulletList':
            Lists.insertBulletList(ws, op.anchor, op.mode, op.items, op);
            receipt.success = true;
            receipt.matched = op.anchor;
            receipt.paraId = _findParaId(op.anchor);
            opCount.bulletList = (opCount.bulletList || 0) + 1;
            break;
          case 'numberedList':
            Lists.insertNumberedList(ws, op.anchor, op.mode, op.items, op);
            receipt.success = true;
            receipt.matched = op.anchor;
            receipt.paraId = _findParaId(op.anchor);
            opCount.numberedList = (opCount.numberedList || 0) + 1;
            break;
        }
      } catch (err) {
        receipt.success = false;
        receipt.error = err.message;
        console.error(`[docex] WARN: ${op.type} operation failed: ${err.message}`);
        console.error(`[docex]   anchor/text: "${(op.oldText || op.anchor || op.text || '').slice(0, 50)}"`);
      }

      receipts.push(receipt);
    }

    // v0.4.9: Auto-append to embedded changelog before saving
    if (this._operations.length > 0) {
      try {
        const now = new Date().toISOString();
        const changelogEntries = receipts
          .filter(r => r.success)
          .map(r => ({
            timestamp: now,
            operation: r.type,
            author: this._author,
            description: r.context ? r.context.slice(0, 120) : r.type,
          }));
        if (changelogEntries.length > 0) {
          Provenance.appendChangelog(ws, changelogEntries);
        }
      } catch (_) {
        // Non-fatal: if changelog writing fails, still save the document
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

    // Attach receipts to result (v0.4.1)
    result.receipts = receipts;

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

  // ── RangeHandle ──────────────────────────────────────────────────────────

  /**
   * Create a RangeHandle spanning from one paragraph offset to another.
   * @param {string} fromParaId
   * @param {number} fromOffset
   * @param {string} toParaId
   * @param {number} toOffset
   * @returns {RangeHandle}
   */
  range(fromParaId, fromOffset, toParaId, toOffset) {
    const { RangeHandle } = require('./range');
    return new RangeHandle(this, fromParaId, fromOffset, toParaId, toOffset);
  }

  // ── Named Checkpoints ────────────────────────────────────────────────────

  /**
   * Save the current document state as a named checkpoint.
   * @param {string} name - Checkpoint name
   */
  async checkpoint(name) {
    const ws = await this._ensureWorkspace();
    if (!this._checkpoints) this._checkpoints = [];
    this._checkpoints.push({ name, docXml: ws.docXml, date: new Date().toISOString() });
    return this;
  }

  /**
   * Restore the document to a previously saved checkpoint.
   * @param {string} name - Checkpoint name
   */
  async restoreTo(name) {
    if (!this._checkpoints) throw new Error('Checkpoint does not exist: ' + name);
    const ckpt = this._checkpoints.find(c => c.name === name);
    if (!ckpt) throw new Error('Checkpoint does not exist: ' + name);
    const ws = await this._ensureWorkspace();
    ws.docXml = ckpt.docXml;
    return this;
  }

  /**
   * List all saved checkpoints.
   * @returns {Array<{name: string, date: string}>}
   */
  async listCheckpoints() {
    return (this._checkpoints || []).map(c => ({ name: c.name, date: c.date }));
  }

  // ── Sections ─────────────────────────────────────────────────────────────

  /**
   * Get outline of headings and section structure.
   */
  async outline() {
    const ws = await this._ensureWorkspace();
    return Sections.outline(ws);
  }

  /**
   * Move a section before or after another heading.
   */
  async moveSection(sectionHeading, opts = {}) {
    const ws = await this._ensureWorkspace();
    return Sections.move(ws, sectionHeading, opts);
  }

  /**
   * Extract a section into a new .docx file.
   */
  async splitDocument(sectionHeading, outputPath) {
    const ws = await this._ensureWorkspace();
    return Sections.split(ws, sectionHeading, outputPath);
  }

  /**
   * Extract abstract text.
   */
  async extractAbstract() {
    const ws = await this._ensureWorkspace();
    return Sections.extractAbstract(ws);
  }

  /**
   * Duplicate a section with a new heading.
   */
  async duplicateSection(sectionHeading, newHeading) {
    const ws = await this._ensureWorkspace();
    return Sections.duplicate(ws, sectionHeading, newHeading);
  }

  // ── Redact ───────────────────────────────────────────────────────────────

  /**
   * Redact text in the document.
   */
  async redact(text, replacement, opts = {}) {
    const ws = await this._ensureWorkspace();
    return Redact.redact(ws, text, replacement, opts);
  }

  /**
   * Restore redacted text.
   */
  async unredact(opts = {}) {
    const ws = await this._ensureWorkspace();
    return Redact.unredact(ws, opts);
  }

  /**
   * Compare document styles against a preset.
   */
  async compareStyles(presetName) {
    const ws = await this._ensureWorkspace();
    return Presets.compareStyles(ws, presetName);
  }

  // ── Quality ──────────────────────────────────────────────────────────────

  /**
   * Run lint checks on the document.
   */
  async lint() {
    const ws = await this._ensureWorkspace();
    return Quality.lint(ws);
  }

  /**
   * Detect passive voice constructions.
   */
  async passiveVoice() {
    const ws = await this._ensureWorkspace();
    return Quality.passiveVoice(ws);
  }

  /**
   * Flag long sentences.
   */
  async sentenceLength(opts = {}) {
    const ws = await this._ensureWorkspace();
    return Quality.sentenceLength(ws, opts);
  }

  /**
   * Calculate readability scores.
   */
  async readability() {
    const ws = await this._ensureWorkspace();
    return Quality.readability(ws);
  }

  /**
   * Check numbers in the document against a stats file.
   */
  async checkNumbers(statsPath) {
    const ws = await this._ensureWorkspace();
    return Quality.checkNumbers(ws, statsPath);
  }

  // ── Production ───────────────────────────────────────────────────────────

  /**
   * Add a watermark to the document.
   */
  async watermark(text, opts = {}) {
    const ws = await this._ensureWorkspace();
    Production.watermark(ws, text, opts);
    return this;
  }

  /**
   * Add a text stamp to the document header/footer.
   */
  async stamp(text, opts = {}) {
    const ws = await this._ensureWorkspace();
    Production.stamp(ws, text, opts);
    return this;
  }

  /**
   * Estimate page count.
   */
  async pageCount() {
    const ws = await this._ensureWorkspace();
    return Production.pageCount(ws);
  }

  /**
   * Insert a cover page.
   */
  async coverPage(opts = {}) {
    const ws = await this._ensureWorkspace();
    Production.coverPage(ws, opts);
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

// Static methods for presets (no document needed)
docex.defineStyle = function(name, config) {
  Presets.define(name, config);
};
docex.listStyles = function() {
  return Presets.list();
};

// v0.4.0: LaTeX pipeline, batch, templates, response letter
docex.compile = Compile.fromLatex;
docex.decompile = Compile.decompile;
docex.watch = Compile.watch;
docex.batch = function(paths) { return new Batch(paths); };
docex.fromTemplate = Template.create;
docex.responseLetter = ResponseLetter.generate;
docex.create = createEmpty;

module.exports = docex;
module.exports.DocexEngine = DocexEngine;
module.exports.PositionSelector = PositionSelector;
module.exports.Presets = Presets;
module.exports.Compile = Compile;
module.exports.Batch = Batch;
module.exports.Template = Template;
module.exports.ResponseLetter = ResponseLetter;
module.exports.Layout = Layout;
module.exports.Provenance = Provenance;
module.exports.Workflow = Workflow;
module.exports.Transaction = Transaction;

// Apply ParagraphHandle extensions (conditionals, verification)
require('./extensions');
