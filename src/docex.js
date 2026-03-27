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

    for (const op of this._operations) {
      try {
        switch (op.type) {
          case 'replace':
            Paragraphs.replace(ws, op.oldText, op.newText, op);
            opCount.replace++;
            break;
          case 'replaceAll': {
            // Replace all occurrences by looping until no more matches
            let count = 0;
            const maxIter = 1000; // safety limit
            while (count < maxIter) {
              try {
                Paragraphs.replace(ws, op.oldText, op.newText, op);
                count++;
              } catch (_) {
                break; // no more matches
              }
            }
            opCount.replaceAll += count;
            break;
          }
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
            opCount.format++;
            break;
          }
          case 'footnote':
            Footnotes.add(ws, op.anchor, op.text, op);
            opCount.footnote++;
            break;
          case 'bulletList':
            Lists.insertBulletList(ws, op.anchor, op.mode, op.items, op);
            opCount.bulletList = (opCount.bulletList || 0) + 1;
            break;
          case 'numberedList':
            Lists.insertNumberedList(ws, op.anchor, op.mode, op.items, op);
            opCount.numberedList = (opCount.numberedList || 0) + 1;
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

// Static methods for presets (no document needed)
docex.defineStyle = function(name, config) {
  Presets.define(name, config);
};
docex.listStyles = function() {
  return Presets.list();
};

module.exports = docex;
module.exports.DocexEngine = DocexEngine;
module.exports.PositionSelector = PositionSelector;
module.exports.Presets = Presets;
