/**
 * transaction.js -- Atomic transactions for docex
 *
 * A Transaction collects multiple operations and applies them atomically:
 * either all succeed or none do (auto-rollback on failure).
 *
 * Designed for autonomous AI agents that need reliable, predictable edits.
 *
 * Usage:
 *   const tx = doc.transaction();
 *   tx.id("3A2F0021").replace("old", "new");
 *   tx.at("enforcement").comment("note", { by: "Reviewer 2" });
 *   tx.preview();
 *   await tx.commit();  // atomic: all or nothing
 *
 * Zero external dependencies.
 */

'use strict';

const { ParagraphHandle } = require('./handle');

// ============================================================================
// TRANSACTION HANDLE
// ============================================================================

class TransactionHandle {
  constructor(transaction, paraId) {
    this._tx = transaction;
    this._paraId = paraId;
  }

  get id() { return this._paraId; }

  replace(oldText, newText, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleReplace', paraId: this._paraId,
      args: [oldText, newText, opts],
      desc: "replace '" + _truncate(oldText, 30) + "' -> '" + _truncate(newText, 30) + "' in para " + this._paraId,
    });
    return this;
  }

  delete(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleDelete', paraId: this._paraId,
      args: [text, opts],
      desc: "delete '" + _truncate(text, 30) + "' in para " + this._paraId,
    });
    return this;
  }

  bold(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleBold', paraId: this._paraId,
      args: [text, opts],
      desc: "bold '" + _truncate(text, 30) + "' in para " + this._paraId,
    });
    return this;
  }

  italic(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleItalic', paraId: this._paraId,
      args: [text, opts],
      desc: "italic '" + _truncate(text, 30) + "' in para " + this._paraId,
    });
    return this;
  }

  highlight(text, color, opts) {
    if (typeof color === 'object') { opts = color; color = 'yellow'; }
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleHighlight', paraId: this._paraId,
      args: [text, color || 'yellow', opts],
      desc: "highlight '" + _truncate(text, 30) + "' in para " + this._paraId,
    });
    return this;
  }

  comment(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleComment', paraId: this._paraId,
      args: [text, opts],
      desc: "comment '" + _truncate(text, 30) + "' on para " + this._paraId,
    });
    return this;
  }

  insertAfter(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleInsertAfter', paraId: this._paraId,
      args: [text, opts],
      desc: "insertAfter '" + _truncate(text, 30) + "' after para " + this._paraId,
    });
    return this;
  }

  insertBefore(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleInsertBefore', paraId: this._paraId,
      args: [text, opts],
      desc: "insertBefore '" + _truncate(text, 30) + "' before para " + this._paraId,
    });
    return this;
  }

  remove(opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'handleRemove', paraId: this._paraId,
      args: [opts],
      desc: "remove para " + this._paraId,
    });
    return this;
  }
}

// ============================================================================
// TRANSACTION POSITION SELECTOR
// ============================================================================

class TransactionPositionSelector {
  constructor(transaction, anchor, mode) {
    this._tx = transaction;
    this._anchor = anchor;
    this._mode = mode;
  }

  insert(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'engineInsertAt',
      args: [this._anchor, this._mode, text, opts],
      desc: "insert " + this._mode + " '" + _truncate(this._anchor, 30) + "': '" + _truncate(text, 30) + "'",
    });
    return this._tx;
  }

  comment(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'engineComment',
      args: [this._anchor, text, opts],
      desc: "comment at '" + _truncate(this._anchor, 30) + "': '" + _truncate(text, 30) + "'",
    });
    return this._tx;
  }

  bold(opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'engineFormat',
      args: ['bold', this._anchor, opts],
      desc: "bold '" + _truncate(this._anchor, 30) + "'",
    });
    return this._tx;
  }

  italic(opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'engineFormat',
      args: ['italic', this._anchor, opts],
      desc: "italic '" + _truncate(this._anchor, 30) + "'",
    });
    return this._tx;
  }

  highlight(color, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'engineFormat',
      args: ['highlight', this._anchor, { colorName: color || 'yellow' }],
      desc: "highlight '" + _truncate(this._anchor, 30) + "'",
    });
    return this._tx;
  }

  footnote(text, opts) {
    if (!opts) opts = {};
    this._tx._queue.push({
      method: 'engineFootnote',
      args: [this._anchor, text, opts],
      desc: "footnote at '" + _truncate(this._anchor, 30) + "': '" + _truncate(text, 30) + "'",
    });
    return this._tx;
  }
}

// ============================================================================
// TRANSACTION CLASS
// ============================================================================

class Transaction {
  constructor(engine) {
    this._engine = engine;
    this._queue = [];
    this._committed = false;
    this._aborted = false;
    this._snapshotTaken = false;
  }

  id(paraId) { return new TransactionHandle(this, paraId); }

  replace(oldText, newText, opts) {
    if (!opts) opts = {};
    this._queue.push({
      method: 'engineReplace',
      args: [oldText, newText, opts],
      desc: "replace '" + _truncate(oldText, 30) + "' -> '" + _truncate(newText, 30) + "'",
    });
    return this;
  }

  delete(text, opts) {
    if (!opts) opts = {};
    this._queue.push({
      method: 'engineDelete',
      args: [text, opts],
      desc: "delete '" + _truncate(text, 30) + "'",
    });
    return this;
  }

  bold(text, opts) {
    if (!opts) opts = {};
    this._queue.push({ method: 'engineFormat', args: ['bold', text, opts], desc: "bold '" + _truncate(text, 30) + "'" });
    return this;
  }

  italic(text, opts) {
    if (!opts) opts = {};
    this._queue.push({ method: 'engineFormat', args: ['italic', text, opts], desc: "italic '" + _truncate(text, 30) + "'" });
    return this;
  }

  highlight(text, color, opts) {
    if (typeof color === 'object') { opts = color; color = 'yellow'; }
    if (!opts) opts = {};
    this._queue.push({ method: 'engineFormat', args: ['highlight', text, { colorName: color || 'yellow' }], desc: "highlight '" + _truncate(text, 30) + "'" });
    return this;
  }

  comment(anchor, text, opts) {
    if (!opts) opts = {};
    this._queue.push({
      method: 'engineComment',
      args: [anchor, text, opts],
      desc: "comment at '" + _truncate(anchor, 30) + "': '" + _truncate(text, 30) + "'",
    });
    return this;
  }

  at(text) { return new TransactionPositionSelector(this, text, 'at'); }
  after(text) { return new TransactionPositionSelector(this, text, 'after'); }

  preview() {
    if (this._queue.length === 0) return 'Transaction: no pending operations.';
    var lines = ['Transaction: ' + this._queue.length + ' pending operation' + (this._queue.length !== 1 ? 's' : '') + ':'];
    for (var i = 0; i < this._queue.length; i++) {
      lines.push('  ' + (i + 1) + '. ' + this._queue[i].desc);
    }
    return lines.join('\n');
  }

  async commit(saveOpts) {
    if (this._committed) throw new Error('Transaction already committed');
    if (this._aborted) throw new Error('Transaction already aborted');

    var ws = await this._engine._ensureWorkspace();
    ws.snapshot();
    this._snapshotTaken = true;

    try {
      for (var i = 0; i < this._queue.length; i++) {
        this._execute(this._queue[i], ws);
      }
      var result = ws.save(saveOpts || { backup: false });
      result.operations = this._queue.length;
      this._committed = true;
      this._engine._workspace = null;
      return result;
    } catch (err) {
      ws.rollback();
      this._snapshotTaken = false;
      throw err;
    }
  }

  async abort() {
    if (this._committed) throw new Error('Transaction already committed');
    this._aborted = true;
    this._queue = [];
    if (this._snapshotTaken) {
      var ws = await this._engine._ensureWorkspace();
      ws.rollback();
      this._snapshotTaken = false;
    }
  }

  _execute(op, ws) {
    var Paragraphs = require('./paragraphs').Paragraphs;
    var Comments = require('./comments').Comments;
    var Formatting = require('./formatting').Formatting;
    var Footnotes = require('./footnotes').Footnotes;

    var author = this._engine._author;
    var date = this._engine._date;
    var tracked = this._engine._tracked;

    switch (op.method) {
      case 'handleReplace': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.replace(op.args[0], op.args[1], op.args[2]);
        break;
      }
      case 'handleDelete': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.delete(op.args[0], op.args[1]);
        break;
      }
      case 'handleBold': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.bold(op.args[0], op.args[1]);
        break;
      }
      case 'handleItalic': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.italic(op.args[0], op.args[1]);
        break;
      }
      case 'handleHighlight': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.highlight(op.args[0], op.args[1], op.args[2]);
        break;
      }
      case 'handleComment': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.comment(op.args[0], op.args[1]);
        break;
      }
      case 'handleInsertAfter': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.insertAfter(op.args[0], op.args[1]);
        break;
      }
      case 'handleInsertBefore': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.insertBefore(op.args[0], op.args[1]);
        break;
      }
      case 'handleRemove': {
        var handle = new ParagraphHandle(this._engine, op.paraId);
        handle.remove(op.args[0]);
        break;
      }
      case 'engineReplace': {
        Paragraphs.replace(ws, op.args[0], op.args[1], {
          tracked: op.args[2].tracked !== undefined ? op.args[2].tracked : tracked,
          author: op.args[2].author || author, date: date,
        });
        break;
      }
      case 'engineDelete': {
        Paragraphs.remove(ws, op.args[0], {
          tracked: op.args[1].tracked !== undefined ? op.args[1].tracked : tracked,
          author: op.args[1].author || author, date: date,
        });
        break;
      }
      case 'engineComment': {
        var by = op.args[2].by || op.args[2].author || author;
        Comments.add(ws, op.args[0], op.args[1], {
          author: by,
          initials: op.args[2].initials || by.split(' ').map(function(w) { return w[0]; }).join(''),
          date: date,
        });
        break;
      }
      case 'engineFormat': {
        var fmtOpts = {
          tracked: op.args[2] && op.args[2].tracked !== undefined ? op.args[2].tracked : false,
          author: (op.args[2] && op.args[2].author) || author, date: date,
        };
        switch (op.args[0]) {
          case 'bold': Formatting.bold(ws, op.args[1], fmtOpts); break;
          case 'italic': Formatting.italic(ws, op.args[1], fmtOpts); break;
          case 'highlight': Formatting.highlight(ws, op.args[1], (op.args[2] && op.args[2].colorName) || 'yellow', fmtOpts); break;
          default: throw new Error('Unknown format type: ' + op.args[0]);
        }
        break;
      }
      case 'engineInsertAt': {
        Paragraphs.insert(ws, op.args[0], op.args[1], op.args[2], {
          tracked: op.args[3].tracked !== undefined ? op.args[3].tracked : tracked,
          author: op.args[3].author || author, date: date,
        });
        break;
      }
      case 'engineFootnote': {
        Footnotes.add(ws, op.args[0], op.args[1], {
          author: (op.args[2] && op.args[2].author) || author, date: date,
        });
        break;
      }
      default:
        throw new Error('Unknown transaction operation: ' + op.method);
    }
  }
}

// ============================================================================
// HELPERS
// ============================================================================

function _truncate(str, max) {
  if (!str) return '';
  if (str.length <= max) return str;
  return str.slice(0, max) + '...';
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = { Transaction, TransactionHandle, TransactionPositionSelector };
