/**
 * batch.js -- Batch operations across multiple .docx files
 *
 * Applies the same operations to multiple documents in one pass.
 * Each file gets its own docex engine instance.
 *
 * Usage:
 *   const docs = new Batch(["paper1.docx", "paper2.docx"]);
 *   docs.author("Fabio Votta");
 *   docs.style("polcomm");
 *   docs.replaceAll("old term", "new term");
 *   await docs.saveAll();
 *
 * Zero external dependencies beyond docex internals.
 */

'use strict';

const path = require('path');

// ============================================================================
// BATCH CLASS
// ============================================================================

class Batch {

  /**
   * Create a batch operation context for multiple .docx files.
   *
   * @param {string[]} paths - Array of .docx file paths
   */
  constructor(paths) {
    if (!Array.isArray(paths) || paths.length === 0) {
      throw new Error('Batch requires a non-empty array of .docx file paths');
    }

    /** @type {string[]} Resolved file paths */
    this._paths = paths.map(p => path.resolve(p));

    /** @type {Array<object>} Queued operations to apply */
    this._operations = [];

    /** @type {string} Author name for operations */
    this._author = 'Unknown';

    /** @type {string|null} Style preset name */
    this._style = null;
  }

  /**
   * Set the author name for all documents.
   * @param {string} name - Author name
   * @returns {Batch} this (for chaining)
   */
  author(name) {
    this._author = name;
    return this;
  }

  /**
   * Queue a style preset to apply to all documents.
   * @param {string} presetName - Preset name (e.g. "polcomm", "apa7")
   * @returns {Batch} this (for chaining)
   */
  style(presetName) {
    this._style = presetName;
    return this;
  }

  /**
   * Queue a replaceAll operation for all documents.
   * @param {string} oldText - Text to find
   * @param {string} newText - Replacement text
   * @returns {Batch} this (for chaining)
   */
  replaceAll(oldText, newText) {
    this._operations.push({ type: 'replaceAll', oldText, newText });
    return this;
  }

  /**
   * Queue a replace operation (first occurrence) for all documents.
   * @param {string} oldText - Text to find
   * @param {string} newText - Replacement text
   * @returns {Batch} this (for chaining)
   */
  replace(oldText, newText) {
    this._operations.push({ type: 'replace', oldText, newText });
    return this;
  }

  /**
   * Verify all documents against a preset, returning per-file results.
   * @param {string} presetName - Preset name
   * @returns {Promise<Array<{path: string, result: object}>>}
   */
  async verify(presetName) {
    const docex = require('./docex');
    const results = [];

    for (const filePath of this._paths) {
      const doc = docex(filePath);
      try {
        const result = await doc.verify(presetName);
        results.push({ path: filePath, result });
      } catch (err) {
        results.push({ path: filePath, result: { pass: false, errors: [err.message], warnings: [] } });
      } finally {
        doc.discard();
      }
    }

    return results;
  }

  /**
   * Execute all queued operations on all documents and save.
   *
   * @param {object} [opts] - Save options
   * @param {string} [opts.outputDir] - Save all files to this directory (keep original names)
   * @param {string} [opts.suffix] - Add suffix to filenames (e.g. "_formatted")
   * @returns {Promise<Array<{path: string, fileSize: number, paragraphCount: number, error: string|null}>>}
   */
  async saveAll(opts = {}) {
    const docex = require('./docex');
    const results = [];

    for (const filePath of this._paths) {
      try {
        const doc = docex(filePath);
        doc.author(this._author);

        // Apply style if queued
        if (this._style) {
          await doc.style(this._style);
        }

        // Apply queued operations
        for (const op of this._operations) {
          switch (op.type) {
            case 'replace':
              doc.replace(op.oldText, op.newText);
              break;
            case 'replaceAll':
              doc.replaceAll(op.oldText, op.newText);
              break;
          }
        }

        // Determine output path
        let outputPath;
        if (opts.outputDir) {
          const base = path.basename(filePath);
          outputPath = path.join(path.resolve(opts.outputDir), base);
        } else if (opts.suffix) {
          const dir = path.dirname(filePath);
          const ext = path.extname(filePath);
          const base = path.basename(filePath, ext);
          outputPath = path.join(dir, base + opts.suffix + ext);
        } else {
          outputPath = undefined; // overwrite in place
        }

        const result = await doc.save(outputPath);
        results.push({
          path: result.path,
          fileSize: result.fileSize,
          paragraphCount: result.paragraphCount,
          error: null,
        });
      } catch (err) {
        results.push({
          path: filePath,
          fileSize: 0,
          paragraphCount: 0,
          error: err.message,
        });
      }
    }

    return results;
  }

  /**
   * Get the number of files in the batch.
   * @returns {number}
   */
  get length() {
    return this._paths.length;
  }

  /**
   * Get the file paths in the batch.
   * @returns {string[]}
   */
  get paths() {
    return [...this._paths];
  }
}

module.exports = { Batch };
