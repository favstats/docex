'use strict';

const xml = require('./xml');
const { DocMap } = require('./docmap');
const { Paragraphs } = require('./paragraphs');
const { Comments } = require('./comments');
const { Formatting } = require('./formatting');
const { Footnotes } = require('./footnotes');

// Default run properties for new paragraphs (Times New Roman 12pt)
const DEFAULT_RPR =
  '<w:rPr>'
  + '<w:rFonts w:hint="default" w:ascii="Times New Roman" '
  + 'w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" '
  + 'w:cs="Times New Roman" />'
  + '<w:sz w:val="24" />'
  + '</w:rPr>';

// ============================================================================
// ParagraphHandle -- fluent API for a single paragraph by paraId
// ============================================================================

class ParagraphHandle {

  /**
   * @param {object} engine - The DocexEngine instance
   * @param {string} paraId - The w14:paraId of the target paragraph
   */
  constructor(engine, paraId) {
    this._engine = engine;
    this._paraId = paraId;
  }

  /** The w14:paraId of this paragraph. */
  get id() {
    return this._paraId;
  }

  /** The decoded plain text of this paragraph. */
  get text() {
    return xml.extractTextDecoded(this._locate().xml);
  }

  /** The paragraph type: 'heading', 'caption', 'abstract', or 'body'. */
  get type() {
    const loc = this._locate();
    const level = Paragraphs._headingLevel(loc.xml);
    if (level > 0) return 'heading';
    const trimmedText = xml.extractTextDecoded(loc.xml).trim();
    if (/^(Figure|Table)\s+\d/i.test(trimmedText)) return 'caption';
    if (/^abstract$/i.test(trimmedText)) return 'abstract';
    return 'body';
  }

  /** The heading text of the section containing this paragraph. */
  get section() {
    const ws = this._getWorkspace();
    const paragraphs = xml.findParagraphs(ws.docXml);
    const loc = this._locate();

    let matchIndex = -1;
    for (let i = 0; i < paragraphs.length; i++) {
      if (paragraphs[i].start === loc.start) {
        matchIndex = i;
        break;
      }
    }
    if (matchIndex === -1) return '(unknown)';

    for (let j = matchIndex - 1; j >= 0; j--) {
      if (Paragraphs._headingLevel(paragraphs[j].xml) > 0) {
        return xml.extractTextDecoded(paragraphs[j].xml);
      }
    }
    return '(before first heading)';
  }

  /**
   * Replace text within this paragraph.
   *
   * @param {string} oldText - Text to find
   * @param {string} newText - Replacement text
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {ParagraphHandle} this (for chaining)
   */
  replace(oldText, newText, opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;
    const author = opts.author || this._engine._author;

    if (tracked) {
      const nextId = xml.nextChangeId(ws.docXml);
      const result = Paragraphs._injectReplacement(
        loc.xml, oldText, newText, nextId, author, this._engine._date
      );
      if (result.modified) {
        ws.docXml = ws.docXml.slice(0, loc.start) + result.xml + ws.docXml.slice(loc.end);
      } else {
        throw new Error('Text not found in paragraph ' + this._paraId);
      }
    } else {
      const proxy = {
        _xml: loc.xml,
        get docXml() { return this._xml; },
        set docXml(v) { this._xml = v; },
      };
      if (!xml.extractTextDecoded(loc.xml).includes(oldText)) {
        throw new Error('Text not found in paragraph ' + this._paraId);
      }
      Paragraphs._replaceDirect(proxy, oldText, newText);
      ws.docXml = ws.docXml.slice(0, loc.start) + proxy.docXml + ws.docXml.slice(loc.end);
    }
    return this;
  }

  /**
   * Delete text from this paragraph.
   *
   * @param {string} text - Text to delete
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {ParagraphHandle} this (for chaining)
   */
  delete(text, opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;

    if (tracked) {
      const nextId = xml.nextChangeId(ws.docXml);
      const result = Paragraphs._injectDeletion(
        loc.xml, text, nextId,
        opts.author || this._engine._author,
        this._engine._date
      );
      if (result.modified) {
        ws.docXml = ws.docXml.slice(0, loc.start) + result.xml + ws.docXml.slice(loc.end);
      } else {
        throw new Error('Text not found');
      }
    } else {
      this.replace(text, '', { tracked: false });
    }
    return this;
  }

  /** Apply bold formatting to matched text. */
  bold(text, opts) {
    if (!opts) opts = {};
    this._applyFormatting('bold', text, opts);
    return this;
  }

  /** Apply italic formatting to matched text. */
  italic(text, opts) {
    if (!opts) opts = {};
    this._applyFormatting('italic', text, opts);
    return this;
  }

  /** Apply highlight formatting to matched text. */
  highlight(text, color, opts) {
    if (!opts) opts = {};
    if (typeof color === 'object') {
      opts = color;
      color = 'yellow';
    }
    this._applyFormatting('highlight', text, Object.assign({}, opts, { colorName: color || 'yellow' }));
    return this;
  }

  /**
   * Add a comment anchored to this paragraph.
   *
   * @param {string} text - Comment text
   * @param {object} [opts] - Options: { by, author, initials }
   * @returns {ParagraphHandle} this (for chaining)
   */
  comment(text, opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const loc = this._locate();
    const paraText = xml.extractTextDecoded(loc.xml);
    const anchorText = paraText.slice(0, 80);
    if (!anchorText) throw new Error('No text');

    const commentAuthor = opts.by || opts.author || this._engine._author;
    Comments.add(ws, anchorText, text, {
      author: commentAuthor,
      initials: opts.initials || commentAuthor.split(' ').map(function (w) { return w[0]; }).join(''),
      date: this._engine._date,
    });
    return this;
  }

  /**
   * Add a footnote anchored to this paragraph.
   *
   * @param {string} text - Footnote text
   * @param {object} [opts] - Options: { author }
   * @returns {ParagraphHandle} this (for chaining)
   */
  footnote(text, opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const loc = this._locate();
    const anchorText = xml.extractTextDecoded(loc.xml).slice(0, 80);
    Footnotes.add(ws, anchorText, text, {
      author: opts.author || this._engine._author,
      date: this._engine._date,
    });
    return this;
  }

  /**
   * Get paragraph-level formatting info.
   *
   * @returns {{ style: string, font: string|null, size: number|null, keepWithNext: boolean, pageBreakBefore: boolean }}
   */
  formatting() {
    const loc = this._locate();
    const pXml = loc.xml;

    const style = (pXml.match(/<w:pStyle\s+w:val="([^"]+)"/) || [])[1] || 'Normal';
    const fontMatch = pXml.match(/<w:rFonts[^>]*w:ascii="([^"]+)"/);
    const font = fontMatch ? fontMatch[1] : null;
    const sizeMatch = pXml.match(/<w:sz\s+w:val="(\d+)"/);
    const size = sizeMatch ? parseInt(sizeMatch[1], 10) / 2 : null;
    const keepWithNext = /<w:keepNext\s*\/?>/.test(pXml);
    const pageBreakBefore = /<w:pageBreakBefore\s*\/?>/.test(pXml);

    return { style, font, size, keepWithNext, pageBreakBefore };
  }

  /**
   * Get character-level formatting at a specific character offset.
   *
   * @param {number} [charOffset=0] - Character offset within the paragraph
   * @returns {object} Formatting info
   */
  formattingAt(charOffset) {
    const loc = this._locate();
    const { _formattingAtOffset } = require('./range');
    return _formattingAtOffset(loc.xml, charOffset || 0);
  }

  /** Get the raw XML of this paragraph. */
  getXml() {
    return this._locate().xml;
  }

  /**
   * Replace the entire paragraph XML.
   *
   * @param {string} xmlString - New XML for the paragraph
   * @returns {ParagraphHandle} this (for chaining)
   */
  setXml(xmlString) {
    const ws = this._getWorkspace();
    const loc = this._locate();
    ws.docXml = ws.docXml.slice(0, loc.start) + xmlString + ws.docXml.slice(loc.end);
    return this;
  }

  /**
   * Inject raw XML before or after specific text within the paragraph.
   *
   * @param {string} xmlString - XML to inject
   * @param {object} opts - { before: string } or { after: string }
   * @returns {ParagraphHandle} this (for chaining)
   */
  injectXml(xmlString, opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const loc = this._locate();
    let paraXml = loc.xml;

    const anchorText = opts.before || opts.after;
    if (!anchorText) throw new Error('Need before or after');

    const decodedText = xml.extractTextDecoded(paraXml);
    const anchorIndex = decodedText.indexOf(anchorText);
    if (anchorIndex === -1) throw new Error('Text not found');

    const targetOffset = opts.before ? anchorIndex : anchorIndex + anchorText.length;
    const runs = xml.parseRuns(paraXml);
    let charCount = 0;

    for (let ri = 0; ri < runs.length; ri++) {
      const run = runs[ri];
      if (run.texts.length === 0) continue;
      const runText = run.combinedText;
      const runEnd = charCount + runText.length;

      if (targetOffset >= charCount && targetOffset <= runEnd) {
        const offsetInRun = targetOffset - charCount;

        if (offsetInRun === 0) {
          // Insert before this run
          paraXml = paraXml.slice(0, run.index) + xmlString + paraXml.slice(run.index);
        } else if (offsetInRun === runText.length) {
          // Insert after this run
          const afterPos = run.index + run.fullMatch.length;
          paraXml = paraXml.slice(0, afterPos) + xmlString + paraXml.slice(afterPos);
        } else {
          // Split the run
          const beforeText = runText.slice(0, offsetInRun);
          const afterText = runText.slice(offsetInRun);
          const splitXml =
            '<w:r>' + run.rPr + '<w:t xml:space="preserve">' + beforeText + '</w:t></w:r>'
            + xmlString
            + '<w:r>' + run.rPr + '<w:t xml:space="preserve">' + afterText + '</w:t></w:r>';
          paraXml = paraXml.slice(0, run.index) + splitXml + paraXml.slice(run.index + run.fullMatch.length);
        }
        break;
      }
      charCount += runText.length;
    }

    ws.docXml = ws.docXml.slice(0, loc.start) + paraXml + ws.docXml.slice(loc.end);
    return this;
  }

  /**
   * List all text runs in this paragraph.
   *
   * @returns {Array<{ textId: string|null, text: string, formatting: object }>}
   */
  runs() {
    const loc = this._locate();
    const parsedRuns = xml.parseRuns(loc.xml);
    const results = [];

    for (let i = 0; i < parsedRuns.length; i++) {
      const run = parsedRuns[i];
      if (run.texts.length === 0) continue;
      const textIdMatch = run.fullMatch.match(/w14:textId="([^"]+)"/);
      results.push({
        textId: textIdMatch ? textIdMatch[1] : null,
        text: xml.decodeXml(run.combinedText),
        formatting: RunHandle._parseFormatting(run.rPr),
      });
    }
    return results;
  }

  /**
   * Merge consecutive runs with identical formatting.
   *
   * @returns {number} Number of merges performed
   */
  mergeRuns() {
    const ws = this._getWorkspace();
    const loc = this._locate();
    let paraXml = loc.xml;
    const parsedRuns = xml.parseRuns(paraXml);
    const textRuns = parsedRuns.filter(function (r) { return r.texts.length > 0; });

    if (textRuns.length < 2) return 0;

    let mergeCount = 0;
    for (let i = textRuns.length - 1; i > 0; i--) {
      const current = textRuns[i];
      const previous = textRuns[i - 1];
      const currentRPr = current.rPr.replace(/\s+/g, ' ').trim();
      const previousRPr = previous.rPr.replace(/\s+/g, ' ').trim();

      if (currentRPr === previousRPr) {
        const combinedText = previous.combinedText + current.combinedText;
        const newRun = '<w:r>' + previous.rPr
          + '<w:t xml:space="preserve">' + combinedText + '</w:t></w:r>';
        // Remove current run, then replace previous run
        paraXml = paraXml.slice(0, current.index) + paraXml.slice(current.index + current.fullMatch.length);
        paraXml = paraXml.slice(0, previous.index) + newRun + paraXml.slice(previous.index + previous.fullMatch.length);
        mergeCount++;
      }
    }

    if (mergeCount > 0) {
      ws.docXml = ws.docXml.slice(0, loc.start) + paraXml + ws.docXml.slice(loc.end);
    }
    return mergeCount;
  }

  /**
   * Replace text at specific character offsets.
   *
   * @param {number} start - Start character offset (inclusive)
   * @param {number} end - End character offset (exclusive)
   * @param {string} newText - Replacement text
   * @param {object} [opts] - Options
   * @returns {ParagraphHandle} this (for chaining)
   */
  replaceAt(start, end, newText, opts) {
    if (!opts) opts = {};
    const fullText = xml.extractTextDecoded(this._locate().xml);
    if (start < 0 || end > fullText.length || start >= end) {
      throw new Error('Invalid offsets');
    }
    return this.replace(fullText.slice(start, end), newText, opts);
  }

  /**
   * Get a RunHandle for a specific text run by textId.
   *
   * @param {string} textId - The w14:textId of the run
   * @returns {RunHandle}
   */
  run(textId) {
    return new RunHandle(this._engine, this._paraId, textId);
  }

  /**
   * Insert a new paragraph after this one.
   *
   * @param {string} text - Text for the new paragraph
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {string} The paraId of the new paragraph
   */
  insertAfter(text, opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;
    const pPr = Paragraphs._extractPpr(loc.xml);
    const newParaId = xml.randomHexId();
    const newTextId = xml.randomHexId();
    const escapedText = xml.escapeXml(text);

    let newParagraph;
    if (tracked) {
      const changeId = xml.nextChangeId(ws.docXml);
      const authorName = xml.escapeXml(opts.author || this._engine._author);
      newParagraph = '<w:p w14:paraId="' + newParaId + '" w14:textId="' + newTextId + '">'
        + pPr
        + '<w:ins w:id="' + changeId + '" w:author="' + authorName
        + '" w:date="' + this._engine._date + '">'
        + '<w:r>' + DEFAULT_RPR
        + '<w:t xml:space="preserve">' + escapedText + '</w:t></w:r>'
        + '</w:ins></w:p>';
    } else {
      newParagraph = '<w:p w14:paraId="' + newParaId + '" w14:textId="' + newTextId + '">'
        + pPr
        + '<w:r>' + DEFAULT_RPR
        + '<w:t xml:space="preserve">' + escapedText + '</w:t></w:r>'
        + '</w:p>';
    }

    ws.docXml = ws.docXml.slice(0, loc.end) + newParagraph + ws.docXml.slice(loc.end);
    return newParaId;
  }

  /**
   * Insert a new paragraph before this one.
   *
   * @param {string} text - Text for the new paragraph
   * @param {object} [opts] - Options: { tracked, author }
   * @returns {string} The paraId of the new paragraph
   */
  insertBefore(text, opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;
    const pPr = Paragraphs._extractPpr(loc.xml);
    const newParaId = xml.randomHexId();
    const newTextId = xml.randomHexId();
    const escapedText = xml.escapeXml(text);

    let newParagraph;
    if (tracked) {
      const changeId = xml.nextChangeId(ws.docXml);
      const authorName = xml.escapeXml(opts.author || this._engine._author);
      newParagraph = '<w:p w14:paraId="' + newParaId + '" w14:textId="' + newTextId + '">'
        + pPr
        + '<w:ins w:id="' + changeId + '" w:author="' + authorName
        + '" w:date="' + this._engine._date + '">'
        + '<w:r>' + DEFAULT_RPR
        + '<w:t xml:space="preserve">' + escapedText + '</w:t></w:r>'
        + '</w:ins></w:p>';
    } else {
      newParagraph = '<w:p w14:paraId="' + newParaId + '" w14:textId="' + newTextId + '">'
        + pPr
        + '<w:r>' + DEFAULT_RPR
        + '<w:t xml:space="preserve">' + escapedText + '</w:t></w:r>'
        + '</w:p>';
    }

    ws.docXml = ws.docXml.slice(0, loc.start) + newParagraph + ws.docXml.slice(loc.start);
    return newParaId;
  }

  /**
   * Remove this paragraph from the document.
   *
   * @param {object} [opts] - Options: { tracked, author }
   */
  remove(opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const loc = this._locate();
    const tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;

    if (tracked) {
      const paraText = xml.extractTextDecoded(loc.xml);
      if (paraText.trim()) {
        const result = Paragraphs._injectDeletion(
          loc.xml, paraText, xml.nextChangeId(ws.docXml),
          opts.author || this._engine._author, this._engine._date
        );
        if (result.modified) {
          ws.docXml = ws.docXml.slice(0, loc.start) + result.xml + ws.docXml.slice(loc.end);
        }
      }
    } else {
      ws.docXml = ws.docXml.slice(0, loc.start) + ws.docXml.slice(loc.end);
    }
  }

  // --------------------------------------------------------------------------
  // Private helpers
  // --------------------------------------------------------------------------

  /**
   * Locate this paragraph in the document XML.
   * @returns {{ xml: string, start: number, end: number }}
   * @private
   */
  _locate() {
    const ws = this._getWorkspace();
    const result = DocMap.locateById(ws.docXml, this._paraId);
    if (!result) throw new Error('Paragraph "' + this._paraId + '" not found');
    return result;
  }

  /**
   * Get the workspace, throwing if not opened.
   * @returns {object}
   * @private
   */
  _getWorkspace() {
    if (!this._engine._workspace) throw new Error('Document not opened yet.');
    return this._engine._workspace;
  }

  /**
   * Apply formatting (bold, italic, highlight) to text in the paragraph.
   * @param {string} formatType - 'bold', 'italic', or 'highlight'
   * @param {string} text - Text to format
   * @param {object} opts - Options
   * @private
   */
  _applyFormatting(formatType, text, opts) {
    const ws = this._getWorkspace();
    const loc = this._locate();

    if (!xml.extractTextDecoded(loc.xml).includes(text)) {
      throw new Error('Text not found');
    }

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
// RunHandle -- fluent API for a single text run by textId
// ============================================================================

class RunHandle {

  /**
   * @param {object} engine - The DocexEngine instance
   * @param {string} paraId - The w14:paraId of the containing paragraph
   * @param {string} textId - The w14:textId of the target run
   */
  constructor(engine, paraId, textId) {
    this._engine = engine;
    this._paraId = paraId;
    this._textId = textId;
  }

  /** Apply bold to this run. */
  bold(opts) {
    if (!opts) opts = {};
    this._modifyRunProperties('<w:b/>', opts);
    return this;
  }

  /** Apply italic to this run. */
  italic(opts) {
    if (!opts) opts = {};
    this._modifyRunProperties('<w:i/>', opts);
    return this;
  }

  /** Get the decoded text of this run. */
  text() {
    const runXml = this._locateRun();
    const texts = [];
    const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let m;
    while ((m = re.exec(runXml)) !== null) {
      texts.push(m[1]);
    }
    return xml.decodeXml(texts.join(''));
  }

  /**
   * Set the text of this run.
   * @param {string} newText - New text content
   * @returns {RunHandle} this (for chaining)
   */
  setText(newText) {
    return this.replace(newText);
  }

  /** Get the formatting properties of this run. */
  formatting() {
    const runXml = this._locateRun();
    const rPrMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    return RunHandle._parseFormatting(rPrMatch ? rPrMatch[0] : '');
  }

  /**
   * Set formatting properties on this run.
   *
   * @param {object} props - Formatting properties
   * @returns {RunHandle} this (for chaining)
   */
  setFormatting(props) {
    const ws = this._getWorkspace();
    const paraLoc = DocMap.locateById(ws.docXml, this._paraId);
    if (!paraLoc) throw new Error('Para NF');

    const searchStr = 'w14:textId="' + this._textId + '"';
    const searchPos = paraLoc.xml.indexOf(searchStr);
    if (searchPos === -1) throw new Error('Run NF');

    const runStart = paraLoc.xml.lastIndexOf('<w:r', searchPos);
    const runEndTag = paraLoc.xml.indexOf('</w:r>', runStart);
    if (runStart === -1 || runEndTag === -1) throw new Error('Parse err');

    let runXml = paraLoc.xml.slice(runStart, runEndTag + 6);

    // Build new rPr
    let rPrInner = '';
    if (props.font) {
      rPrInner += '<w:rFonts w:ascii="' + xml.escapeXml(props.font)
        + '" w:hAnsi="' + xml.escapeXml(props.font) + '"/>';
    }
    if (props.size) rPrInner += '<w:sz w:val="' + props.size + '"/>';
    if (props.bold) rPrInner += '<w:b/>';
    if (props.italic) rPrInner += '<w:i/>';
    if (props.underline) rPrInner += '<w:u w:val="single"/>';
    if (props.strike) rPrInner += '<w:strike/>';
    if (props.smallCaps) rPrInner += '<w:smallCaps/>';
    if (props.color) rPrInner += '<w:color w:val="' + props.color.replace(/^#/, '') + '"/>';
    if (props.highlight) rPrInner += '<w:highlight w:val="' + props.highlight + '"/>';
    if (props.vertAlign) rPrInner += '<w:vertAlign w:val="' + props.vertAlign + '"/>';

    const newRPr = rPrInner ? '<w:rPr>' + rPrInner + '</w:rPr>' : '';

    if (runXml.includes('<w:rPr>')) {
      runXml = runXml.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, newRPr);
    } else if (newRPr) {
      const closeAngle = runXml.indexOf('>');
      runXml = runXml.slice(0, closeAngle + 1) + newRPr + runXml.slice(closeAngle + 1);
    }

    const newPara = paraLoc.xml.slice(0, runStart) + runXml + paraLoc.xml.slice(runEndTag + 6);
    ws.docXml = ws.docXml.slice(0, paraLoc.start) + newPara + ws.docXml.slice(paraLoc.end);
    return this;
  }

  /**
   * Split this run at a character offset, producing two runs.
   *
   * @param {number} charOffset - Character position to split at
   * @returns {[string, string]} The textIds of the two resulting runs
   */
  splitAt(charOffset) {
    const ws = this._getWorkspace();
    const paraLoc = DocMap.locateById(ws.docXml, this._paraId);
    if (!paraLoc) throw new Error('Para NF');

    const searchStr = 'w14:textId="' + this._textId + '"';
    const searchPos = paraLoc.xml.indexOf(searchStr);
    if (searchPos === -1) throw new Error('Run NF');

    const runStart = paraLoc.xml.lastIndexOf('<w:r', searchPos);
    const runEndTag = paraLoc.xml.indexOf('</w:r>', runStart);
    if (runStart === -1 || runEndTag === -1) throw new Error('Parse');

    const runXml = paraLoc.xml.slice(runStart, runEndTag + 6);

    // Extract full text
    const texts = [];
    const textRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let textMatch;
    while ((textMatch = textRe.exec(runXml)) !== null) {
      texts.push(textMatch[1]);
    }
    const fullText = texts.join('');

    if (charOffset <= 0 || charOffset >= fullText.length) {
      throw new Error('Offset OOB');
    }

    const rPrMatch = runXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
    const rPr = rPrMatch ? rPrMatch[0] : '';

    const textId1 = xml.randomHexId();
    const textId2 = xml.randomHexId();

    const run1 = '<w:r w14:textId="' + textId1 + '">' + rPr
      + '<w:t xml:space="preserve">' + fullText.slice(0, charOffset) + '</w:t></w:r>';
    const run2 = '<w:r w14:textId="' + textId2 + '">' + rPr
      + '<w:t xml:space="preserve">' + fullText.slice(charOffset) + '</w:t></w:r>';

    const newPara = paraLoc.xml.slice(0, runStart) + run1 + run2 + paraLoc.xml.slice(runEndTag + 6);
    ws.docXml = ws.docXml.slice(0, paraLoc.start) + newPara + ws.docXml.slice(paraLoc.end);
    return [textId1, textId2];
  }

  /**
   * Move this run to a different paragraph.
   *
   * @param {string} targetParaId - The paraId of the destination paragraph
   * @returns {RunHandle} this (for chaining)
   */
  moveTo(targetParaId) {
    const ws = this._getWorkspace();

    // Find and extract run from source paragraph
    const srcPara = DocMap.locateById(ws.docXml, this._paraId);
    if (!srcPara) throw new Error('Src NF');

    const searchStr = 'w14:textId="' + this._textId + '"';
    const searchPos = srcPara.xml.indexOf(searchStr);
    if (searchPos === -1) throw new Error('Run NF');

    const runStart = srcPara.xml.lastIndexOf('<w:r', searchPos);
    const runEndTag = srcPara.xml.indexOf('</w:r>', runStart);
    if (runStart === -1 || runEndTag === -1) throw new Error('Parse');

    const runXml = srcPara.xml.slice(runStart, runEndTag + 6);
    const newSrcPara = srcPara.xml.slice(0, runStart) + srcPara.xml.slice(runEndTag + 6);
    ws.docXml = ws.docXml.slice(0, srcPara.start) + newSrcPara + ws.docXml.slice(srcPara.end);

    // Insert run into target paragraph
    const tgtPara = DocMap.locateById(ws.docXml, targetParaId);
    if (!tgtPara) throw new Error('Tgt NF');

    const closeIndex = tgtPara.xml.lastIndexOf('</w:p>');
    const newTgtPara = tgtPara.xml.slice(0, closeIndex) + runXml + tgtPara.xml.slice(closeIndex);
    ws.docXml = ws.docXml.slice(0, tgtPara.start) + newTgtPara + ws.docXml.slice(tgtPara.end);

    this._paraId = targetParaId;
    return this;
  }

  /**
   * Replace the text content of this run.
   *
   * @param {string} newText - New text content
   * @param {object} [opts] - Options (reserved for future use)
   * @returns {RunHandle} this (for chaining)
   */
  replace(newText, opts) {
    if (!opts) opts = {};
    const ws = this._getWorkspace();
    const paraLoc = DocMap.locateById(ws.docXml, this._paraId);
    if (!paraLoc) throw new Error('Para NF');

    const searchStr = 'w14:textId="' + this._textId + '"';
    const searchPos = paraLoc.xml.indexOf(searchStr);
    if (searchPos === -1) throw new Error('Run NF');

    const runStart = paraLoc.xml.lastIndexOf('<w:r', searchPos);
    const runEndTag = paraLoc.xml.indexOf('</w:r>', runStart);
    if (runStart === -1 || runEndTag === -1) throw new Error('Parse');

    const runXml = paraLoc.xml.slice(runStart, runEndTag + 6);
    const newRunXml = runXml.replace(
      /<w:t[^>]*>[^<]*<\/w:t>/,
      '<w:t xml:space="preserve">' + xml.escapeXml(newText) + '</w:t>'
    );

    const newPara = paraLoc.xml.slice(0, runStart) + newRunXml + paraLoc.xml.slice(runEndTag + 6);
    ws.docXml = ws.docXml.slice(0, paraLoc.start) + newPara + ws.docXml.slice(paraLoc.end);
    return this;
  }

  // --------------------------------------------------------------------------
  // Private helpers
  // --------------------------------------------------------------------------

  /**
   * Get the workspace, throwing if not opened.
   * @returns {object}
   * @private
   */
  _getWorkspace() {
    if (!this._engine._workspace) throw new Error('Not opened');
    return this._engine._workspace;
  }

  /**
   * Locate and return the raw XML of this run.
   * @returns {string}
   * @private
   */
  _locateRun() {
    const ws = this._getWorkspace();
    const paraLoc = DocMap.locateById(ws.docXml, this._paraId);
    if (!paraLoc) throw new Error('Para NF');

    const searchStr = 'w14:textId="' + this._textId + '"';
    const searchPos = paraLoc.xml.indexOf(searchStr);
    if (searchPos === -1) throw new Error('Run NF');

    const runStart = paraLoc.xml.lastIndexOf('<w:r', searchPos);
    const runEndTag = paraLoc.xml.indexOf('</w:r>', runStart);
    if (runStart === -1 || runEndTag === -1) throw new Error('Parse');

    return paraLoc.xml.slice(runStart, runEndTag + 6);
  }

  /**
   * Add a property element to the run's rPr.
   * @param {string} propXml - XML element to add (e.g., '<w:b/>')
   * @param {object} opts - Options (reserved for future use)
   * @private
   */
  _modifyRunProperties(propXml, opts) {
    const ws = this._getWorkspace();
    const paraLoc = DocMap.locateById(ws.docXml, this._paraId);
    if (!paraLoc) throw new Error('Para NF');

    const searchStr = 'w14:textId="' + this._textId + '"';
    const searchPos = paraLoc.xml.indexOf(searchStr);
    if (searchPos === -1) throw new Error('Run NF');

    const runStart = paraLoc.xml.lastIndexOf('<w:r', searchPos);
    const runEndTag = paraLoc.xml.indexOf('</w:r>', runStart);
    if (runStart === -1 || runEndTag === -1) throw new Error('Parse');

    let runXml = paraLoc.xml.slice(runStart, runEndTag + 6);

    if (runXml.includes('<w:rPr>')) {
      runXml = runXml.replace('<w:rPr>', '<w:rPr>' + propXml);
    } else {
      runXml = runXml.replace(/>/, '><w:rPr>' + propXml + '</w:rPr>');
    }

    const newPara = paraLoc.xml.slice(0, runStart) + runXml + paraLoc.xml.slice(runEndTag + 6);
    ws.docXml = ws.docXml.slice(0, paraLoc.start) + newPara + ws.docXml.slice(paraLoc.end);
  }

  /**
   * Parse run properties XML into a structured formatting object.
   *
   * @param {string} rPr - The w:rPr XML string
   * @returns {object} Formatting properties
   */
  static _parseFormatting(rPr) {
    const result = {
      bold: false,
      italic: false,
      underline: false,
      font: null,
      size: null,
      color: null,
      highlight: null,
      strike: false,
      vertAlign: null,
      smallCaps: false,
    };

    if (!rPr) return result;

    result.bold = /<w:b\s*\/>/.test(rPr) || /<w:b\s+/.test(rPr);
    result.italic = /<w:i\s*\/>/.test(rPr) || /<w:i\s+/.test(rPr);
    result.underline = /<w:u\s/.test(rPr);
    result.strike = /<w:strike/.test(rPr);
    result.smallCaps = /<w:smallCaps/.test(rPr);

    const fontMatch = rPr.match(/w:rFonts[^>]*w:ascii="([^"]*)"/);
    if (fontMatch) result.font = fontMatch[1];

    const sizeMatch = rPr.match(/<w:sz\s+w:val="(\d+)"/);
    if (sizeMatch) result.size = parseInt(sizeMatch[1], 10);

    const colorMatch = rPr.match(/<w:color\s+w:val="([^"]*)"/);
    if (colorMatch) result.color = colorMatch[1];

    const highlightMatch = rPr.match(/<w:highlight\s+w:val="([^"]*)"/);
    if (highlightMatch) result.highlight = highlightMatch[1];

    const vertAlignMatch = rPr.match(/<w:vertAlign\s+w:val="([^"]*)"/);
    if (vertAlignMatch) result.vertAlign = vertAlignMatch[1];

    return result;
  }
}

module.exports = { ParagraphHandle, RunHandle };
