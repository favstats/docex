'use strict';
var xml = require('./xml');
var DocMap = require('./docmap').DocMap;
var Formatting = require('./formatting').Formatting;
var Comments = require('./comments').Comments;

function _parseRprFormatting(rPr) {
  return {
    bold: /<w:b\s*\/?>/.test(rPr) || /<w:b\s+/.test(rPr),
    italic: /<w:i\s*\/?>/.test(rPr) || /<w:i\s+/.test(rPr),
    underline: /<w:u\s/.test(rPr),
    strikethrough: /<w:strike/.test(rPr),
    font: (rPr.match(/w:ascii="([^"]+)"/) || [])[1] || null,
    size: (function() { var m = rPr.match(/<w:sz\s+w:val="(\d+)"/); return m ? parseInt(m[1], 10) / 2 : null; })(),
    color: (rPr.match(/<w:color\s+w:val="([^"]+)"/) || [])[1] || null,
    highlight: (rPr.match(/<w:highlight\s+w:val="([^"]+)"/) || [])[1] || null,
    superscript: /w:val="superscript"/.test(rPr),
    subscript: /w:val="subscript"/.test(rPr),
    smallCaps: /<w:smallCaps/.test(rPr),
  };
}

function _formattingAtOffset(pXml, charOffset) {
  var runs = xml.parseRuns(pXml);
  var offset = 0;
  for (var i = 0; i < runs.length; i++) {
    var runText = xml.decodeXml(runs[i].combinedText);
    if (offset + runText.length > charOffset) return _parseRprFormatting(runs[i].rPr);
    offset += runText.length;
  }
  return _parseRprFormatting('');
}

function RangeHandle(engine, fromParaId, fromOffset, toParaId, toOffset) {
  this._engine = engine;
  this._fromParaId = fromParaId;
  this._fromOffset = fromOffset;
  this._toParaId = toParaId;
  this._toOffset = toOffset;
}

RangeHandle.prototype._getWorkspace = function() {
  if (!this._engine._workspace) throw new Error('Document not opened yet.');
  return this._engine._workspace;
};

RangeHandle.prototype._paragraphsInRange = function(ws) {
  var allParas = xml.findParagraphs(ws.docXml);
  var result = [];
  var inRange = false;
  for (var i = 0; i < allParas.length; i++) {
    var paraId = DocMap._extractParaId(allParas[i].xml);
    if (paraId === this._fromParaId) inRange = true;
    if (inRange) result.push({ xml: allParas[i].xml, start: allParas[i].start, end: allParas[i].end, paraId: paraId });
    if (paraId === this._toParaId) break;
  }
  return result;
};

RangeHandle.prototype.text = function() {
  var ws = this._getWorkspace();
  var paragraphs = this._paragraphsInRange(ws);
  var parts = [];
  for (var i = 0; i < paragraphs.length; i++) {
    var fullText = xml.extractTextDecoded(paragraphs[i].xml);
    if (paragraphs.length === 1) parts.push(fullText.slice(this._fromOffset, this._toOffset));
    else if (i === 0) parts.push(fullText.slice(this._fromOffset));
    else if (i === paragraphs.length - 1) parts.push(fullText.slice(0, this._toOffset));
    else parts.push(fullText);
  }
  return parts.join('\n');
};

RangeHandle.prototype.bold = function(opts) {
  if (!opts) opts = {};
  this._applyFormatToRange('<w:b/>', opts);
  return this;
};

RangeHandle.prototype.italic = function(opts) {
  if (!opts) opts = {};
  this._applyFormatToRange('<w:i/>', opts);
  return this;
};

RangeHandle.prototype.highlight = function(color, opts) {
  if (typeof color === 'object') { opts = color; color = 'yellow'; }
  if (!opts) opts = {};
  color = color || 'yellow';
  this._applyFormatToRange('<w:highlight w:val="' + color + '"/>', opts);
  return this;
};

RangeHandle.prototype.delete = function(opts) {
  if (!opts) opts = {};
  var ws = this._getWorkspace();
  var tracked = opts.tracked !== undefined ? opts.tracked : this._engine._tracked;
  var author = opts.author || this._engine._author;
  var date = this._engine._date;
  var paragraphs = this._paragraphsInRange(ws);
  var Paragraphs = require('./paragraphs').Paragraphs;
  for (var i = paragraphs.length - 1; i >= 0; i--) {
    var p = paragraphs[i];
    var fullText = xml.extractTextDecoded(p.xml);
    var textToDelete;
    if (paragraphs.length === 1) textToDelete = fullText.slice(this._fromOffset, this._toOffset);
    else if (i === 0) textToDelete = fullText.slice(this._fromOffset);
    else if (i === paragraphs.length - 1) textToDelete = fullText.slice(0, this._toOffset);
    else textToDelete = fullText;
    if (!textToDelete) continue;
    if (tracked) {
      var nextId = xml.nextChangeId(ws.docXml);
      var result = Paragraphs._injectDeletion(p.xml, textToDelete, nextId, author, date);
      if (result.modified) ws.docXml = ws.docXml.slice(0, p.start) + result.xml + ws.docXml.slice(p.end);
    } else {
      var proxy = { _xml: p.xml, get docXml() { return this._xml; }, set docXml(v) { this._xml = v; } };
      Paragraphs._replaceDirect(proxy, textToDelete, '');
      ws.docXml = ws.docXml.slice(0, p.start) + proxy.docXml + ws.docXml.slice(p.end);
    }
  }
  return this;
};

RangeHandle.prototype.comment = function(text, opts) {
  if (!opts) opts = {};
  var ws = this._getWorkspace();
  var firstLoc = DocMap.locateById(ws.docXml, this._fromParaId);
  if (!firstLoc) throw new Error('Paragraph ' + this._fromParaId + ' not found');
  var fullText = xml.extractTextDecoded(firstLoc.xml);
  var anchorText = fullText.slice(this._fromOffset, this._fromOffset + 40) || fullText.slice(0, 40);
  Comments.add(ws, anchorText, text, { author: opts.by || opts.author || this._engine._author, initials: (opts.by || opts.author || this._engine._author).split(' ').map(function(w) { return w[0]; }).join(''), date: this._engine._date });
  return this;
};

RangeHandle.prototype.cut = function() {
  var cutText = this.text();
  this._engine._clipboard = cutText;
  this.delete({ tracked: false });
  return cutText;
};

RangeHandle.prototype.formatting = function() {
  var ws = this._getWorkspace();
  var firstLoc = DocMap.locateById(ws.docXml, this._fromParaId);
  if (!firstLoc) throw new Error('Paragraph ' + this._fromParaId + ' not found');
  return _formattingAtOffset(firstLoc.xml, this._fromOffset);
};

RangeHandle.prototype._applyFormatToRange = function(formatElement, opts) {
  var ws = this._getWorkspace();
  var paragraphs = this._paragraphsInRange(ws);
  for (var i = paragraphs.length - 1; i >= 0; i--) {
    var p = paragraphs[i];
    var fullText = xml.extractTextDecoded(p.xml);
    var textToFormat;
    if (paragraphs.length === 1) textToFormat = fullText.slice(this._fromOffset, this._toOffset);
    else if (i === 0) textToFormat = fullText.slice(this._fromOffset);
    else if (i === paragraphs.length - 1) textToFormat = fullText.slice(0, this._toOffset);
    else textToFormat = fullText;
    if (!textToFormat) continue;
    var result = Formatting._formatInParagraph(p.xml, textToFormat, formatElement, false, '', '', ws.docXml);
    if (result.modified) ws.docXml = ws.docXml.slice(0, p.start) + result.xml + ws.docXml.slice(p.end);
  }
};

module.exports = { RangeHandle: RangeHandle, _formattingAtOffset: _formattingAtOffset, _parseRprFormatting: _parseRprFormatting };
