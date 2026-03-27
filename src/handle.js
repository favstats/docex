/**
 * handle.js -- ParagraphHandle and RunHandle for docex stable addressing
 *
 * A ParagraphHandle provides a stable reference to a specific paragraph
 * via its w14:paraId. All operations on the handle locate the paragraph
 * fresh from the current docXml, so they survive mutations.
 *
 * All methods operate through the DocexEngine and Workspace.
 * XML manipulation is done entirely with string operations and regex.
 * Zero external dependencies.
 */

'use strict';

const xml = require('./xml');
const { DocMap } = require('./docmap');
const { Paragraphs } = require('./paragraphs');
const { Comments } = require('./comments');
const { Formatting } = require('./formatting');
const { Footnotes } = require('./footnotes');

// ============================================================================
// PARAGRAPH HANDLE
// ============================================================================

class ParagraphHandle {

  /**
   * @param {object} engine - DocexEngine instance
   * @param {string} paraId - The w14:paraId of the target paragraph
   */
  constructor(engine, paraId) {
    this._engine = engine;
    this._paraId = paraId;
  }

  // --------------------------------------------------------------------------
  // Properties (lazy, read from current doc state)
  // --------------------------------------------------------------------------

  /** @returns {string} The paraId */
  get id() {
    return this._paraId;
  }

  /** @returns {string} Current text of this paragraph */
  get text() {
    const loc = this._locate();
    return xml.extractTextDecoded(loc.xml);
  }

  /** @returns {string} Paragraph type: heading | body | caption | abstract */
  get type() {
    const loc = this._locate();
    const level = Paragraphs._headingLevel(loc.xml);
    if (level > 0) return 'heading';
    const text = xml.extractTextDecoded(loc.xml).trim();
    if (/^(Figure|Table)\s+\d/i.test(text)) return 'caption';
    if (/^abstract$/i.test(text)) return 'abstract';
    return 'body';
  }

  /** @returns {string} Section heading text this paragraph belongs to */
  get section() {
    const ws = this._getWorkspace();
    const docXml = ws.docXml;
    const loc = this._locate();
    const paragraphs = xml.findParagraphs(docXml);

    // Find the index of our paragraph
    let myIndex = -1;
    for (let i = 0; i < paragraphs.length; i++) {
      if (paragraphs[i].start === loc.start) {
        myIndex = i;
        break;
      }
    }

    if (myIndex === -1) return '(unknown)';

    // Walk backwards to find the nearest heading
    for (let j = myIndex - 1; j >= 0; j--) {
      const level = Paragraphs._headingLevel(paragraphs[j].xml);
      if (level > 0) {
        return xml.extractTextDecoded(paragraphs[j].xml);
      }
    }

    return '(before first heading)';
  }

  // --------------------------------------------------------------------------
  // Tier 2C: Text match within this paragraph
  // --------------------------------------------------------------------------

  /**
   * Replace text within this specific paragraph.
   *
   * @param {string} oldText - Text to find in this paragraph
   * @param {string} newText - Replacement text
   * @param {object} [opts] - Options: { nth, tracked, author }
   * @returns {ParagraphHandle} this, for chaining
   */
  replace(oldText, newText, opts = {}) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;
    const author = opts.author || this._engine._author;
    const date = this._engine._date;

    if (tracked) {
      const nextId = xml.nextChangeId(ws.docXml);
      const result = Paragraphs._injectReplacement(
        loc.xml, oldText, newText, nextId, author, date
      );
      if (result.modified) {
        ws.docXml = ws.docXml.slice(0, loc.start) + result.xml + ws.docXml.slice(loc.end);
      } else {
        throw new Error(`Text not found in paragraph ${this._paraId}: "${oldText.slice(0, 60)}"`);
      }
    } else {
      // Direct replacement within this paragraph
      const proxy = {
        _xml: loc.xml,
        get docXml() { return this._xml; },
        set docXml(v) { this._xml = v; },
      };
      // Create a mini-doc with just this paragraph for the replace method
      const decoded = xml.extractTextDecoded(loc.xml);
      if (!decoded.includes(oldText)) {
        throw new Error(`Text not found in paragraph ${this._paraId}: "${oldText.slice(0, 60)}"`);
      }
      // Use the paragraph-level direct replacement
      Paragraphs._replaceDirect(proxy, oldText, newText);
      ws.docXml = ws.docXml.slice(0, loc.start) + proxy.docXml + ws.docXml.slice(loc.end);
    }

    return this;
  }

  /**
   * Delete text within this paragraph.
   *
   * @param {string} text - Text to delete
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {ParagraphHandle} this, for chaining
   */
  delete(text, opts = {}) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;
    const author = opts.author || this._engine._author;
    const date = this._engine._date;

    if (tracked) {
      const nextId = xml.nextChangeId(ws.docXml);
      const result = Paragraphs._injectDeletion(
        loc.xml, text, nextId, author, date
      );
      if (result.modified) {
        ws.docXml = ws.docXml.slice(0, loc.start) + result.xml + ws.docXml.slice(loc.end);
      } else {
        throw new Error(`Text not found for deletion in paragraph ${this._paraId}: "${text.slice(0, 60)}"`);
      }
    } else {
      this.replace(text, '', { tracked: false });
    }

    return this;
  }

  /**
   * Bold text within this paragraph.
   *
   * @param {string} text - Text to make bold
   * @param {object} [opts] - Options
   * @returns {ParagraphHandle} this, for chaining
   */
  bold(text, opts = {}) {
    this._formatInParagraph('bold', text, opts);
    return this;
  }

  /**
   * Italic text within this paragraph.
   *
   * @param {string} text - Text to make italic
   * @param {object} [opts] - Options
   * @returns {ParagraphHandle} this, for chaining
   */
  italic(text, opts = {}) {
    this._formatInParagraph('italic', text, opts);
    return this;
  }

  /**
   * Highlight text within this paragraph.
   *
   * @param {string} text - Text to highlight
   * @param {string} [color='yellow'] - Highlight color
   * @param {object} [opts] - Options
   * @returns {ParagraphHandle} this, for chaining
   */
  highlight(text, color, opts = {}) {
    if (typeof color === 'object') { opts = color; color = 'yellow'; }
    this._formatInParagraph('highlight', text, { ...opts, colorName: color || 'yellow' });
    return this;
  }

  /**
   * Add a comment anchored to text in this paragraph.
   *
   * @param {string} text - Comment text
   * @param {object} [opts] - Options: { by, initials }
   * @returns {ParagraphHandle} this, for chaining
   */
  comment(text, opts = {}) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    const paraText = xml.extractTextDecoded(loc.xml);

    // Use the paragraph's text as anchor (or a portion of it)
    const anchor = paraText.slice(0, 80);
    if (!anchor) {
      throw new Error(`Paragraph ${this._paraId} has no text to anchor a comment to`);
    }

    // Add comment using existing Comments module
    Comments.add(ws, anchor, text, {
      author: opts.by || opts.author || this._engine._author,
      initials: opts.initials || (opts.by || opts.author || this._engine._author).split(' ').map(w => w[0]).join(''),
      date: this._engine._date,
    });

    return this;
  }

  /**
   * Add a footnote anchored to text in this paragraph.
   *
   * @param {string} text - Footnote text
   * @param {object} [opts] - Options
   * @returns {ParagraphHandle} this, for chaining
   */
  footnote(text, opts = {}) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    const paraText = xml.extractTextDecoded(loc.xml);
    const anchor = paraText.slice(0, 80);

    Footnotes.add(ws, anchor, text, {
      author: opts.author || this._engine._author,
      date: this._engine._date,
    });

    return this;
  }

  // --------------------------------------------------------------------------
  // Tier 2A: Character offset
  // --------------------------------------------------------------------------

  /**
   * Replace text at character offsets within this paragraph.
   *
   * @param {number} start - Start character offset (0-based)
   * @param {number} end - End character offset (exclusive)
   * @param {string} newText - Replacement text
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {ParagraphHandle} this, for chaining
   */
  replaceAt(start, end, newText, opts = {}) {
    const loc = this._locate();
    const fullText = xml.extractTextDecoded(loc.xml);

    if (start < 0 || end > fullText.length || start >= end) {
      throw new Error(`Invalid offsets [${start}, ${end}) for paragraph with ${fullText.length} chars`);
    }

    const oldText = fullText.slice(start, end);
    return this.replace(oldText, newText, opts);
  }

  // --------------------------------------------------------------------------
  // Tier 2B: Run-level
  // --------------------------------------------------------------------------

  /**
   * Get a RunHandle for a specific run identified by w14:textId.
   *
   * @param {string} textId - The w14:textId of the run
   * @returns {RunHandle}
   */
  run(textId) {
    return new RunHandle(this._engine, this._paraId, textId);
  }

  // --------------------------------------------------------------------------
  // Structural operations
  // --------------------------------------------------------------------------

  /**
   * Insert a new paragraph after this one.
   *
   * @param {string} text - Text content for the new paragraph
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {string} The paraId of the newly created paragraph
   */
  insertAfter(text, opts = {}) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;
    const author = opts.author || this._engine._author;
    const date = this._engine._date;

    // Extract pPr from this paragraph for formatting consistency
    const pPr = Paragraphs._extractPpr(loc.xml);

    // Generate new paraId
    const newParaId = xml.randomHexId().toUpperCase();
    const newTextId = xml.randomHexId().toUpperCase();
    const escapedText = xml.escapeXml(text);

    let newParaXml;
    if (tracked) {
      const id = xml.nextChangeId(ws.docXml);
      newParaXml = `<w:p w14:paraId="${newParaId}" w14:textId="${newTextId}">`
        + pPr
        + `<w:ins w:id="${id}" w:author="${xml.escapeXml(author)}" w:date="${date}">`
        + '<w:r>'
        + '<w:rPr>'
        + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman" />'
        + '<w:sz w:val="24" />'
        + '</w:rPr>'
        + `<w:t xml:space="preserve">${escapedText}</w:t>`
        + '</w:r>'
        + '</w:ins>'
        + '</w:p>';
    } else {
      newParaXml = `<w:p w14:paraId="${newParaId}" w14:textId="${newTextId}">`
        + pPr
        + '<w:r>'
        + '<w:rPr>'
        + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman" />'
        + '<w:sz w:val="24" />'
        + '</w:rPr>'
        + `<w:t xml:space="preserve">${escapedText}</w:t>`
        + '</w:r>'
        + '</w:p>';
    }

    ws.docXml = ws.docXml.slice(0, loc.end) + newParaXml + ws.docXml.slice(loc.end);
    return newParaId;
  }

  /**
   * Insert a new paragraph before this one.
   *
   * @param {string} text - Text content for the new paragraph
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {string} The paraId of the newly created paragraph
   */
  insertBefore(text, opts = {}) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;
    const author = opts.author || this._engine._author;
    const date = this._engine._date;

    const pPr = Paragraphs._extractPpr(loc.xml);
    const newParaId = xml.randomHexId().toUpperCase();
    const newTextId = xml.randomHexId().toUpperCase();
    const escapedText = xml.escapeXml(text);

    let newParaXml;
    if (tracked) {
      const id = xml.nextChangeId(ws.docXml);
      newParaXml = `<w:p w14:paraId="${newParaId}" w14:textId="${newTextId}">`
        + pPr
        + `<w:ins w:id="${id}" w:author="${xml.escapeXml(author)}" w:date="${date}">`
        + '<w:r>'
        + '<w:rPr>'
        + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman" />'
        + '<w:sz w:val="24" />'
        + '</w:rPr>'
        + `<w:t xml:space="preserve">${escapedText}</w:t>`
        + '</w:r>'
        + '</w:ins>'
        + '</w:p>';
    } else {
      newParaXml = `<w:p w14:paraId="${newParaId}" w14:textId="${newTextId}">`
        + pPr
        + '<w:r>'
        + '<w:rPr>'
        + '<w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman" />'
        + '<w:sz w:val="24" />'
        + '</w:rPr>'
        + `<w:t xml:space="preserve">${escapedText}</w:t>`
        + '</w:r>'
        + '</w:p>';
    }

    ws.docXml = ws.docXml.slice(0, loc.start) + newParaXml + ws.docXml.slice(loc.start);
    return newParaId;
  }

  /**
   * Remove this paragraph from the document.
   *
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {void}
   */
  remove(opts = {}) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;

    if (tracked) {
      // For tracked deletion of a whole paragraph, we wrap its content in w:del
      const author = opts.author || this._engine._author;
      const date = this._engine._date;
      const text = xml.extractTextDecoded(loc.xml);
      if (text.trim()) {
        const nextId = xml.nextChangeId(ws.docXml);
        const result = Paragraphs._injectDeletion(loc.xml, text, nextId, author, date);
        if (result.modified) {
          ws.docXml = ws.docXml.slice(0, loc.start) + result.xml + ws.docXml.slice(loc.end);
        }
      }
    } else {
      // Untracked: just remove the paragraph element
      ws.docXml = ws.docXml.slice(0, loc.start) + ws.docXml.slice(loc.end);
    }
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  /**
   * Find this paragraph in the current docXml by paraId.
   *
   * @returns {{xml: string, start: number, end: number}}
   * @throws {Error} if paraId not found
   * @private
   */
  _locate() {
    const ws = this._getWorkspace();
    const result = DocMap.locateById(ws.docXml, this._paraId);
    if (!result) {
      throw new Error(`Paragraph with paraId "${this._paraId}" not found in document`);
    }
    return result;
  }

  /**
   * Get the workspace from the engine, synchronously.
   *
   * @returns {object} Workspace
   * @private
   */
  _getWorkspace() {
    if (!this._engine._workspace) {
      throw new Error('Document not opened yet. Call await doc.map() or another async method first.');
    }
    return this._engine._workspace;
  }

  /**
   * Apply formatting to text within this paragraph.
   *
   * Creates a mini-workspace scoped to just this paragraph, applies
   * the formatting, then splices it back into the full document.
   *
   * @param {string} formatType - Format type (bold, italic, highlight, etc.)
   * @param {string} text - Text to format
   * @param {object} opts - Options
   * @private
   */
  _formatInParagraph(formatType, text, opts) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    const decoded = xml.extractTextDecoded(loc.xml);

    if (!decoded.includes(text)) {
      throw new Error(`Text not found in paragraph ${this._paraId}: "${text.slice(0, 60)}"`);
    }

    // Create a proxy workspace scoped to this paragraph
    const proxy = {
      _xml: loc.xml,
      get docXml() { return this._xml; },
      set docXml(v) { this._xml = v; },
    };

    const fmtOpts = {
      tracked: opts.tracked !== undefined ? opts.tracked : false,
      author: opts.author || this._engine._author,
      date: this._engine._date,
    };

    switch (formatType) {
      case 'bold':
        Formatting.bold(proxy, text, fmtOpts);
        break;
      case 'italic':
        Formatting.italic(proxy, text, fmtOpts);
        break;
      case 'highlight':
        Formatting.highlight(proxy, text, opts.colorName || 'yellow', fmtOpts);
        break;
    }

    ws.docXml = ws.docXml.slice(0, loc.start) + proxy.docXml + ws.docXml.slice(loc.end);
  }
}

// ============================================================================
// RUN HANDLE
// ============================================================================

class RunHandle {

  /**
   * @param {object} engine - DocexEngine instance
   * @param {string} paraId - The w14:paraId of the parent paragraph
   * @param {string} textId - The w14:textId of the target run
   */
  constructor(engine, paraId, textId) {
    this._engine = engine;
    this._paraId = paraId;
    this._textId = textId;
  }

  /**
   * Make this run bold.
   *
   * @param {object} [opts] - Options
   * @returns {RunHandle} this, for chaining
   */
  bold(opts = {}) {
    this._modifyRunProps('<w:b/>', opts);
    return this;
  }

  /**
   * Make this run italic.
   *
   * @param {object} [opts] - Options
   * @returns {RunHandle} this, for chaining
   */
  italic(opts = {}) {
    this._modifyRunProps('<w:i/>', opts);
    return this;
  }

  /**
   * Replace the text of this run.
   *
   * @param {string} newText - New text for the run
   * @param {object} [opts] - Options
   * @returns {RunHandle} this, for chaining
   */
  replace(newText, opts = {}) {
    const ws = this._getWorkspace();
    const paraLoc = DocMap.locateById(ws.docXml, this._paraId);
    if (!paraLoc) {
      throw new Error(`Paragraph with paraId "${this._paraId}" not found`);
    }

    const searchStr = `w14:textId="${this._textId}"`;
    const runPos = paraLoc.xml.indexOf(searchStr);
    if (runPos === -1) {
      throw new Error(`Run with textId "${this._textId}" not found in paragraph ${this._paraId}`);
    }

    // Find the <w:r start and </w:r> end
    const rStart = paraLoc.xml.lastIndexOf('<w:r', runPos);
    const rEnd = paraLoc.xml.indexOf('</w:r>', rStart);
    if (rStart === -1 || rEnd === -1) {
      throw new Error(`Could not parse run boundaries for textId "${this._textId}"`);
    }

    const runXml = paraLoc.xml.slice(rStart, rEnd + 6);

    // Replace the text content in the run
    const newRunXml = runXml.replace(
      /<w:t[^>]*>[^<]*<\/w:t>/,
      `<w:t xml:space="preserve">${xml.escapeXml(newText)}</w:t>`
    );

    const newParaXml = paraLoc.xml.slice(0, rStart) + newRunXml + paraLoc.xml.slice(rEnd + 6);
    ws.docXml = ws.docXml.slice(0, paraLoc.start) + newParaXml + ws.docXml.slice(paraLoc.end);

    return this;
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------

  _getWorkspace() {
    if (!this._engine._workspace) {
      throw new Error('Document not opened yet.');
    }
    return this._engine._workspace;
  }

  _modifyRunProps(propXml, opts) {
    const ws = this._getWorkspace();
    const paraLoc = DocMap.locateById(ws.docXml, this._paraId);
    if (!paraLoc) {
      throw new Error(`Paragraph with paraId "${this._paraId}" not found`);
    }

    const searchStr = `w14:textId="${this._textId}"`;
    const runPos = paraLoc.xml.indexOf(searchStr);
    if (runPos === -1) {
      throw new Error(`Run with textId "${this._textId}" not found in paragraph ${this._paraId}`);
    }

    const rStart = paraLoc.xml.lastIndexOf('<w:r', runPos);
    const rEnd = paraLoc.xml.indexOf('</w:r>', rStart);
    if (rStart === -1 || rEnd === -1) {
      throw new Error(`Could not parse run boundaries for textId "${this._textId}"`);
    }

    let runXml = paraLoc.xml.slice(rStart, rEnd + 6);

    // If run has rPr, insert the property element inside it
    if (runXml.includes('<w:rPr>')) {
      runXml = runXml.replace('<w:rPr>', `<w:rPr>${propXml}`);
    } else {
      // No rPr yet, add one after <w:r...>
      runXml = runXml.replace(/>/, `><w:rPr>${propXml}</w:rPr>`);
    }

    const newParaXml = paraLoc.xml.slice(0, rStart) + runXml + paraLoc.xml.slice(rEnd + 6);
    ws.docXml = ws.docXml.slice(0, paraLoc.start) + newParaXml + ws.docXml.slice(paraLoc.end);
  }
}

module.exports = { ParagraphHandle, RunHandle };
