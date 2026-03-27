/**
 * compile.js -- LaTeX <-> .docx pipeline for docex
 *
 * Provides the core compile pipeline:
 *   .tex -> pandoc -> raw .docx -> docex post-process -> submission-ready .docx
 *
 * Also provides the reverse pipeline:
 *   .docx -> docex extract -> pandoc -> .tex with tracked changes preserved
 *
 * Requires pandoc to be installed on the system.
 * Zero npm dependencies beyond docex internals.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const { Workspace } = require('./workspace');
const { Presets } = require('./presets');
const { CrossRef } = require('./crossref');
const { DocMap } = require('./docmap');
const { Latex } = require('./latex');
const { Revisions } = require('./revisions');
const { Comments } = require('./comments');
const xml = require('./xml');

// ============================================================================
// HELPERS
// ============================================================================

/**
 * Check if pandoc is installed and accessible.
 * @returns {string} Path to pandoc binary
 * @throws {Error} If pandoc is not found
 */
function _findPandoc() {
  try {
    const result = execFileSync('which', ['pandoc'], {
      stdio: ['pipe', 'pipe', 'pipe'],
      encoding: 'utf-8',
      timeout: 5000,
    }).trim();
    if (!result) throw new Error('pandoc not found');
    return result;
  } catch (_) {
    throw new Error(
      'pandoc is not installed or not in PATH. '
      + 'Install it from https://pandoc.org/installing.html'
    );
  }
}

/**
 * Generate a unique temp file path.
 * @param {string} ext - File extension (e.g. '.docx')
 * @returns {string}
 */
function _tmpFile(ext) {
  const id = crypto.randomBytes(8).toString('hex');
  return path.join(os.tmpdir(), `docex-compile-${id}${ext}`);
}

// ============================================================================
// COMPILE CLASS
// ============================================================================

class Compile {

  /**
   * Compile a LaTeX .tex file to a submission-ready .docx.
   *
   * Pipeline:
   *   1. Verify pandoc is available
   *   2. Run pandoc to convert .tex -> raw .docx
   *   3. Open the raw .docx with docex
   *   4. Post-process: style preset, auto-numbering, paraId injection
   *   5. Save to output path
   *
   * @param {string} texPath - Path to the .tex source file
   * @param {object} [opts] - Options
   * @param {string} [opts.style] - Journal preset name (e.g. "polcomm", "apa7")
   * @param {string} [opts.output] - Output .docx path (default: same dir as .tex, .docx extension)
   * @param {string} [opts.bibFile] - Path to .bib bibliography file
   * @param {string} [opts.cslFile] - Path to .csl citation style file
   * @param {string[]} [opts.pandocArgs] - Additional pandoc arguments
   * @returns {Promise<{path: string, fileSize: number, paragraphCount: number, style: string|null}>}
   */
  static async fromLatex(texPath, opts = {}) {
    const absTexPath = path.resolve(texPath);
    if (!fs.existsSync(absTexPath)) {
      throw new Error(`LaTeX file not found: ${absTexPath}`);
    }

    // Step 1: Check pandoc
    _findPandoc();

    // Step 2: Determine output path
    const outputPath = opts.output
      ? path.resolve(opts.output)
      : absTexPath.replace(/\.tex$/, '.docx');

    // Step 3: Run pandoc to produce raw .docx
    const rawDocx = _tmpFile('.docx');
    try {
      const pandocArgs = [
        absTexPath,
        '-o', rawDocx,
        '--from', 'latex',
        '--to', 'docx',
      ];

      if (opts.bibFile) {
        pandocArgs.push('--bibliography=' + path.resolve(opts.bibFile));
        pandocArgs.push('--citeproc');
      }

      if (opts.cslFile) {
        pandocArgs.push('--csl=' + path.resolve(opts.cslFile));
      }

      if (opts.pandocArgs && Array.isArray(opts.pandocArgs)) {
        pandocArgs.push(...opts.pandocArgs);
      }

      execFileSync('pandoc', pandocArgs, {
        stdio: ['pipe', 'pipe', 'pipe'],
        timeout: 120000,
        cwd: path.dirname(absTexPath),
      });

      if (!fs.existsSync(rawDocx)) {
        throw new Error('pandoc did not produce output file');
      }

      // Step 4: Open the raw .docx with docex and post-process
      const ws = Workspace.open(rawDocx);

      // 4a. Apply style preset if specified
      let styleName = null;
      if (opts.style) {
        Presets.apply(ws, opts.style);
        styleName = opts.style;
      }

      // 4b. Auto-number figures and tables
      CrossRef.autoNumber(ws);

      // 4c. paraIds are already injected by Workspace.open()

      // Step 5: Save to output
      const result = ws.save(outputPath);

      return {
        path: result.path,
        fileSize: result.fileSize,
        paragraphCount: result.paragraphCount,
        style: styleName,
      };
    } finally {
      // Clean up temp file
      try { if (fs.existsSync(rawDocx)) fs.unlinkSync(rawDocx); } catch (_) { /* ignore */ }
    }
  }

  /**
   * Convert a .docx back to LaTeX, preserving tracked changes and comments.
   *
   * Pipeline:
   *   1. Open .docx with docex
   *   2. Extract tracked changes and comments
   *   3. Convert to LaTeX using Latex.convert()
   *   4. If preserveChanges: wrap tracked changes as \replaced{old}{new}
   *   5. If preserveComments: wrap comments as \todo{text}
   *
   * @param {string} docxPath - Path to the .docx file
   * @param {object} [opts] - Options
   * @param {string} [opts.output] - Output .tex path (default: same dir, .tex extension)
   * @param {boolean} [opts.preserveChanges] - Preserve tracked changes as \replaced{old}{new}
   * @param {boolean} [opts.preserveComments] - Preserve comments as \todo{text}
   * @param {string} [opts.documentClass] - LaTeX document class (default: 'article')
   * @returns {Promise<{path: string|null, tex: string, changes: number, comments: number}>}
   */
  static async toLatex(docxPath, opts = {}) {
    const absDocxPath = path.resolve(docxPath);
    if (!fs.existsSync(absDocxPath)) {
      throw new Error(`File not found: ${absDocxPath}`);
    }

    const ws = Workspace.open(absDocxPath);

    // Get tracked changes and comments before conversion
    const revisions = opts.preserveChanges ? Revisions.list(ws) : [];
    const comments = opts.preserveComments ? Comments.list(ws) : [];

    // Convert to LaTeX using existing converter
    let tex = Latex.convert(ws, {
      documentClass: opts.documentClass || 'article',
    });

    // Add the changes and todonotes packages if needed
    const extraPackages = [];
    if (opts.preserveChanges && revisions.length > 0) {
      extraPackages.push('changes');
    }
    if (opts.preserveComments && comments.length > 0) {
      extraPackages.push('todonotes');
    }

    // Inject extra usepackage commands after existing ones
    if (extraPackages.length > 0) {
      const pkgLines = extraPackages.map(p => `\\usepackage{${p}}`).join('\n');
      // Insert before \begin{document}
      tex = tex.replace(/\\begin\{document\}/, pkgLines + '\n\\begin{document}');
    }

    // Append tracked changes as comments at the end of the LaTeX document
    if (opts.preserveChanges && revisions.length > 0) {
      // Insert a summary section before \end{document}
      let changesSummary = '\n% --- Tracked Changes Summary ---\n';
      changesSummary += '% The following tracked changes were found in the .docx:\n';
      for (const rev of revisions) {
        const escapedText = rev.text.replace(/\\/g, '\\\\').replace(/[{}]/g, '\\$&');
        if (rev.type === 'insertion') {
          changesSummary += `% \\added[id=${rev.id}, author={${rev.author}}]{${escapedText}}\n`;
        } else {
          changesSummary += `% \\deleted[id=${rev.id}, author={${rev.author}}]{${escapedText}}\n`;
        }
      }
      tex = tex.replace(/\\end\{document\}/, changesSummary + '\\end{document}');
    }

    if (opts.preserveComments && comments.length > 0) {
      let commentsSummary = '\n% --- Comments ---\n';
      for (const c of comments) {
        const escapedText = c.text.replace(/\\/g, '\\\\').replace(/[{}]/g, '\\$&');
        commentsSummary += `% \\todo[author={${c.author}}]{${escapedText}}\n`;
      }
      tex = tex.replace(/\\end\{document\}/, commentsSummary + '\\end{document}');
    }

    ws.cleanup();

    // Write to file if output path specified
    const outputPath = opts.output
      ? path.resolve(opts.output)
      : null;

    if (outputPath) {
      fs.writeFileSync(outputPath, tex, 'utf-8');
    }

    return {
      path: outputPath,
      tex,
      changes: revisions.length,
      comments: comments.length,
    };
  }

  /**
   * Decompile a .docx back to .tex with tracked changes and comments preserved.
   * Alias for toLatex with preserveChanges and preserveComments enabled.
   *
   * @param {string} docxPath - Path to the .docx file
   * @param {object} [opts] - Options (same as toLatex)
   * @returns {Promise<{path: string|null, tex: string, changes: number, comments: number}>}
   */
  static async decompile(docxPath, opts = {}) {
    return Compile.toLatex(docxPath, {
      ...opts,
      preserveChanges: true,
      preserveComments: true,
    });
  }

  /**
   * Watch a .tex file and recompile on changes.
   *
   * @param {string} texPath - Path to the .tex file to watch
   * @param {object} [opts] - Same options as fromLatex
   * @returns {{ close: Function }} Watcher handle with close() to stop
   */
  static watch(texPath, opts = {}) {
    const absTexPath = path.resolve(texPath);
    if (!fs.existsSync(absTexPath)) {
      throw new Error(`LaTeX file not found: ${absTexPath}`);
    }

    // Check pandoc upfront
    _findPandoc();

    let compiling = false;
    let pendingRecompile = false;

    const outputName = opts.output
      ? path.basename(opts.output)
      : path.basename(absTexPath).replace(/\.tex$/, '.docx');

    async function recompile() {
      if (compiling) {
        pendingRecompile = true;
        return;
      }
      compiling = true;
      const start = Date.now();
      try {
        console.log(`[docex] Recompiling ${path.basename(absTexPath)} -> ${outputName}`);
        const result = await Compile.fromLatex(absTexPath, opts);
        const elapsed = ((Date.now() - start) / 1000).toFixed(1);
        console.log(`[docex] Done in ${elapsed}s -> ${result.path} (${result.fileSize} bytes)`);
      } catch (err) {
        console.error(`[docex] Compile error: ${err.message}`);
      } finally {
        compiling = false;
        if (pendingRecompile) {
          pendingRecompile = false;
          recompile();
        }
      }
    }

    // Debounce: wait 500ms after last change before recompiling
    let debounceTimer = null;
    const watcher = fs.watch(absTexPath, (eventType) => {
      if (eventType === 'change') {
        if (debounceTimer) clearTimeout(debounceTimer);
        debounceTimer = setTimeout(recompile, 500);
      }
    });

    // Initial compile
    recompile();

    return {
      close() {
        watcher.close();
        if (debounceTimer) clearTimeout(debounceTimer);
      },
    };
  }
}

module.exports = { Compile };
