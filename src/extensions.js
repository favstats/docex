/**
 * extensions.js -- Conditionals, verification chaining, and transaction support
 *
 * Augments ParagraphHandle and DocexEngine with:
 *   - Conditional operations: if, unless, ifContains, ifEmpty
 *   - Verification chaining: verify
 *   - Transaction factory: transaction()
 *
 * This module monkey-patches existing classes when required.
 * Zero external dependencies.
 */

'use strict';

// ============================================================================
// VERIFICATION ERROR
// ============================================================================

class VerificationError extends Error {
  constructor(message) {
    super(message);
    this.name = 'VerificationError';
    this.paraId = null;
    this.currentText = null;
  }
}

// ============================================================================
// APPLY PATCHES
// ============================================================================

let _patched = false;

function applyPatches() {
  if (_patched) return;
  _patched = true;

  var ParagraphHandle;
  try {
    ParagraphHandle = require('./handle').ParagraphHandle;
  } catch (_) { return; }

  if (!ParagraphHandle) return;

  // -- Conditional operations --

  if (!ParagraphHandle.prototype.ifContains) {
    ParagraphHandle.prototype.if = function(conditionFn, thenFn) {
      if (conditionFn(this)) { thenFn(this); }
      return this;
    };

    ParagraphHandle.prototype.unless = function(conditionFn, thenFn) {
      if (!conditionFn(this)) { thenFn(this); }
      return this;
    };

    ParagraphHandle.prototype.ifContains = function(text, thenFn) {
      return this.if(function(h) { return h.text.includes(text); }, thenFn);
    };

    ParagraphHandle.prototype.ifEmpty = function(thenFn) {
      return this.if(function(h) { return h.text.trim() === ''; }, thenFn);
    };
  }

  // -- Verification chaining --

  if (!ParagraphHandle.prototype.verify) {
    ParagraphHandle.prototype.verify = function(checkFn) {
      if (!checkFn(this)) {
        var text = this.text;
        var err = new VerificationError(
          'Verification failed for paragraph "' + this._paraId + '": ' +
          'check returned false. Current text: "' +
          text.slice(0, 100) + (text.length > 100 ? '...' : '') + '"'
        );
        err.paraId = this._paraId;
        err.currentText = text;
        throw err;
      }
      return this;
    };
  }

  // -- Transaction factory on DocexEngine --

  try {
    var DocexEngine = require('./docex').DocexEngine;
    var Transaction = require('./transaction').Transaction;
    if (DocexEngine && DocexEngine.prototype && !DocexEngine.prototype.transaction) {
      DocexEngine.prototype.transaction = function() {
        return new Transaction(this);
      };
    }
  } catch (_) {
    // Circular dependency -- caller must patch manually
  }
}

// Apply patches immediately on require
applyPatches();

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = {
  VerificationError: VerificationError,
  applyPatches: applyPatches,
};
