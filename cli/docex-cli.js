#!/usr/bin/env node
/**
 * docex-cli.js -- CLI interface for docex
 *
 * A command-line tool that wraps the docex library API.
 * Feels as natural as LaTeX command-line tools.
 *
 * Usage: docex <command> <file> [arguments] [options]
 *
 * Commands:
 *   replace    <file> <old> <new>              Replace text (tracked by default)
 *   insert     <file> <position> <text>        Insert paragraph (e.g. "after:Methods")
 *   delete     <file> <text>                   Delete text (tracked by default)
 *   comment    <file> <anchor> <text>          Add comment anchored to text
 *   reply      <file> <comment-id> <text>      Reply to existing comment
 *   figure     <file> <position> <image> [caption]  Insert figure
 *   table      <file> <position> <json-file>   Insert table from JSON
 *   cite       <file>                           Inject Zotero citations or list patterns
 *   list       <file> [type]                   List paragraphs|headings|comments|figures
 *
 * Options:
 *   --author <name>     Author name (default: from git config)
 *   --by <name>         Comment author (alias for --author)
 *   --untracked         Disable tracked changes
 *   --output <path>     Save to different file (default: overwrite)
 *   --width <inches>    Figure width in inches (default: 6)
 *   --style <style>     Table style: booktabs|plain (default: booktabs)
 *   --caption <text>    Figure/table caption
 *   --safe <path>       Wrap save through safe-modify.sh for manuscript protection
 *   --help              Show help
 *   --version           Show version
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

// ============================================================================
// ANSI color helpers
// ============================================================================

const C = {
  green:  (s) => `\x1b[32m${s}\x1b[0m`,
  red:    (s) => `\x1b[31m${s}\x1b[0m`,
  yellow: (s) => `\x1b[33m${s}\x1b[0m`,
  bold:   (s) => `\x1b[1m${s}\x1b[0m`,
  dim:    (s) => `\x1b[2m${s}\x1b[0m`,
  cyan:   (s) => `\x1b[36m${s}\x1b[0m`,
};

// ============================================================================
// Argument parser (no external deps)
// ============================================================================

/**
 * Parse process.argv into { command, positionals, options }.
 *
 * Flags like --author "Name" consume the next argument as the value.
 * Boolean flags like --untracked and --help have no value.
 *
 * @param {string[]} argv - process.argv.slice(2)
 * @returns {{ command: string, positionals: string[], options: object }}
 */
function parseArgs(argv) {
  const booleanFlags = new Set(['--untracked', '--help', '--version', '--list']);
  const valueFlags = new Set(['--author', '--by', '--output', '--width', '--style', '--caption', '--safe', '--zotero-key', '--zotero-user', '--collection', '--doc-class', '--bib-file', '--packages']);

  const positionals = [];
  const options = {};
  let i = 0;

  while (i < argv.length) {
    const arg = argv[i];

    if (booleanFlags.has(arg)) {
      const key = arg.replace(/^--/, '');
      options[key] = true;
      i++;
    } else if (valueFlags.has(arg)) {
      const key = arg.replace(/^--/, '');
      i++;
      if (i >= argv.length) {
        die(`Option ${arg} requires a value`);
      }
      options[key] = argv[i];
      i++;
    } else if (arg.startsWith('--')) {
      // Unknown flag -- treat as boolean
      const key = arg.replace(/^--/, '');
      options[key] = true;
      i++;
    } else {
      positionals.push(arg);
      i++;
    }
  }

  const command = positionals.shift() || '';
  return { command, positionals, options };
}

// ============================================================================
// Helpers
// ============================================================================

/**
 * Print an error message and exit with code 1.
 * @param {string} msg
 */
function die(msg) {
  console.error(C.red('Error: ') + msg);
  process.exit(1);
}

/**
 * Print a warning message to stderr.
 * @param {string} msg
 */
function warn(msg) {
  console.error(C.yellow('Warning: ') + msg);
}

/**
 * Get the default author name from git config, falling back to "Unknown".
 * @returns {string}
 */
function getDefaultAuthor() {
  try {
    const name = execFileSync('git', ['config', 'user.name'], {
      stdio: ['pipe', 'pipe', 'pipe'],
      encoding: 'utf-8',
      timeout: 3000,
    }).trim();
    return name || 'Unknown';
  } catch {
    return 'Unknown';
  }
}

/**
 * Resolve the author name from options, with fallback to git config.
 * @param {object} options - Parsed CLI options
 * @returns {string}
 */
function resolveAuthor(options) {
  return options.by || options.author || getDefaultAuthor();
}

/**
 * Build save options from CLI options.
 * When --safe is provided, returns an options object for safe-modify.sh integration.
 * Otherwise returns the output path string (or undefined) for backward compatibility.
 *
 * @param {object} options - Parsed CLI options
 * @param {string} description - Description of the operation for safe-modify.sh
 * @returns {string|object|undefined}
 */
function buildSaveOpts(options, description) {
  if (options.safe) {
    return {
      outputPath: options.output || undefined,
      safeModify: options.safe,
      description: description || 'docex CLI edit',
    };
  }
  return options.output;
}

/**
 * Parse a position string like "after:Methods" or "before:Results".
 * @param {string} posStr
 * @returns {{ mode: string, anchor: string }}
 */
function parsePosition(posStr) {
  const colonIdx = posStr.indexOf(':');
  if (colonIdx === -1) {
    // No prefix -- default to "after"
    return { mode: 'after', anchor: posStr };
  }

  const prefix = posStr.slice(0, colonIdx).toLowerCase();
  const anchor = posStr.slice(colonIdx + 1);

  if (prefix === 'after' || prefix === 'before') {
    return { mode: prefix, anchor };
  }

  // Unknown prefix -- treat the whole thing as the anchor
  return { mode: 'after', anchor: posStr };
}

/**
 * Format file size for human display.
 * @param {number} bytes
 * @returns {string}
 */
function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

/**
 * Print the save result summary.
 * @param {object} result - { path, fileSize, paragraphCount, verified }
 */
function printResult(result) {
  const sizeStr = formatSize(result.fileSize);
  const paraStr = result.paragraphCount + ' paragraphs';
  const verifiedStr = result.verified
    ? C.green('verified')
    : C.yellow('unverified');

  console.log('');
  console.log(C.green('  Saved: ') + result.path);
  console.log(C.green('  Size:  ') + sizeStr);
  console.log(C.green('  Paras: ') + paraStr);
  console.log(C.green('  Check: ') + verifiedStr);
}

/**
 * Get the package version from package.json.
 * @returns {string}
 */
function getVersion() {
  try {
    const pkgPath = path.resolve(__dirname, '..', 'package.json');
    const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf-8'));
    return pkg.version || '0.0.0';
  } catch {
    return '0.0.0';
  }
}

// ============================================================================
// Help text
// ============================================================================

function printHelp() {
  console.log(`
${C.bold('docex')} -- LaTeX for .docx

${C.bold('Usage:')} docex <command> <file> [arguments] [options]

${C.bold('Commands:')}
  ${C.cyan('replace')}    <file> <old> <new>              Replace text (tracked by default)
  ${C.cyan('insert')}     <file> <position> <text>        Insert paragraph (e.g. "after:Methods")
  ${C.cyan('delete')}     <file> <text>                   Delete text (tracked by default)
  ${C.cyan('comment')}    <file> <anchor> <text>          Add comment anchored to text
  ${C.cyan('reply')}      <file> <comment-id> <text>      Reply to existing comment
  ${C.cyan('figure')}     <file> <position> <image> [caption]  Insert figure
  ${C.cyan('table')}      <file> <position> <json-file>   Insert table from JSON
  ${C.cyan('cite')}       <file>                           Inject Zotero citations or list patterns
  ${C.cyan('latex')}      <file>                           Export document as LaTeX
  ${C.cyan('list')}       <file> [type]                   List paragraphs|headings|comments|figures

${C.bold('Options:')}
  --author <name>     Author name (default: from git config)
  --by <name>         Comment author (alias for --author)
  --untracked         Disable tracked changes
  --output <path>     Save to different file (default: overwrite)
  --width <inches>    Figure width in inches (default: 6)
  --style <style>     Table style: booktabs|plain (default: booktabs)
  --caption <text>    Figure/table caption
  --safe <path>       Wrap save through safe-modify.sh for manuscript protection
  --zotero-key <key>  Zotero API key (for cite command)
  --zotero-user <id>  Zotero user ID (for cite command)
  --collection <id>   Zotero collection key (for cite command)
  --list              List citation patterns only (for cite command)
  --doc-class <cls>   LaTeX document class (default: article, for latex command)
  --bib-file <name>   Bibliography file name without .bib (default: references)
  --packages <list>   Comma-separated extra LaTeX packages (for latex command)
  --help              Show help
  --version           Show version

${C.bold('Examples:')}
  ${C.dim('# Replace text with tracked changes')}
  docex replace manuscript.docx "268,635" "300,000" --author "Fabio Votta"

  ${C.dim('# Insert paragraph after a heading')}
  docex insert manuscript.docx "after:Methods" "New methodology paragraph."

  ${C.dim('# Add a reviewer comment')}
  docex comment manuscript.docx "platform governance" "Needs citation" --by "Reviewer 2"

  ${C.dim('# Insert a figure')}
  docex figure manuscript.docx "after:Results" figures/fig03.png --caption "Figure 3. Status"

  ${C.dim('# Inject Zotero citations')}
  docex cite manuscript.docx --zotero-key KEY --zotero-user 6875557 --collection TUWJI72V

  ${C.dim('# List citation patterns without injecting')}
  docex cite manuscript.docx --list

  ${C.dim('# Export to LaTeX (stdout)')}
  docex latex manuscript.docx

  ${C.dim('# Export to LaTeX (file)')}
  docex latex manuscript.docx --output paper.tex

  ${C.dim('# List headings')}
  docex list manuscript.docx headings

  ${C.dim('# List comments')}
  docex list manuscript.docx comments

  ${C.dim('# Save to different file')}
  docex replace manuscript.docx "old" "new" --output manuscript_v2.docx
`);
}

// ============================================================================
// Command handlers
// ============================================================================

/**
 * docex replace <file> <old> <new> [--author] [--untracked] [--output]
 */
async function cmdReplace(positionals, options) {
  if (positionals.length < 3) {
    die('replace requires: <file> <old-text> <new-text>');
  }

  const [file, oldText, newText] = positionals;
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);

  if (options.untracked) {
    doc.untracked();
  }

  doc.replace(oldText, newText);

  console.log(C.dim(`Replacing "${oldText.slice(0, 40)}${oldText.length > 40 ? '...' : ''}" with "${newText.slice(0, 40)}${newText.length > 40 ? '...' : ''}"`));
  console.log(C.dim(`Author: ${author}` + (options.untracked ? ' (untracked)' : ' (tracked)')));

  const saveOpts = buildSaveOpts(options, `Replace: "${oldText.slice(0, 40)}" -> "${newText.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex insert <file> <position> <text> [--author] [--untracked] [--output]
 */
async function cmdInsert(positionals, options) {
  if (positionals.length < 3) {
    die('insert requires: <file> <position> <text>\n  Position format: "after:Heading" or "before:Heading"');
  }

  const [file, posStr, text] = positionals;
  const { mode, anchor } = parsePosition(posStr);
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);

  if (options.untracked) {
    doc.untracked();
  }

  if (mode === 'after') {
    doc.after(anchor).insert(text);
  } else {
    doc.before(anchor).insert(text);
  }

  console.log(C.dim(`Inserting ${mode} "${anchor.slice(0, 40)}${anchor.length > 40 ? '...' : ''}"`));
  console.log(C.dim(`Author: ${author}` + (options.untracked ? ' (untracked)' : ' (tracked)')));

  const saveOpts = buildSaveOpts(options, `Insert ${mode} "${anchor.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex delete <file> <text> [--author] [--untracked] [--output]
 */
async function cmdDelete(positionals, options) {
  if (positionals.length < 2) {
    die('delete requires: <file> <text>');
  }

  const [file, text] = positionals;
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);

  if (options.untracked) {
    doc.untracked();
  }

  doc.delete(text);

  console.log(C.dim(`Deleting "${text.slice(0, 60)}${text.length > 60 ? '...' : ''}"`));
  console.log(C.dim(`Author: ${author}` + (options.untracked ? ' (untracked)' : ' (tracked)')));

  const saveOpts = buildSaveOpts(options, `Delete: "${text.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex comment <file> <anchor> <text> [--by|--author] [--output]
 */
async function cmdComment(positionals, options) {
  if (positionals.length < 3) {
    die('comment requires: <file> <anchor-text> <comment-text>');
  }

  const [file, anchor, text] = positionals;
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);

  doc.at(anchor).comment(text, { by: author });

  console.log(C.dim(`Commenting on "${anchor.slice(0, 40)}${anchor.length > 40 ? '...' : ''}"`));
  console.log(C.dim(`By: ${author}`));

  const saveOpts = buildSaveOpts(options, `Comment on: "${anchor.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex reply <file> <comment-id> <text> [--by|--author] [--output]
 *
 * <comment-id> can be a numeric ID or anchor text to find the comment.
 */
async function cmdReply(positionals, options) {
  if (positionals.length < 3) {
    die('reply requires: <file> <comment-id> <reply-text>');
  }

  const [file, commentIdStr, text] = positionals;
  const author = resolveAuthor(options);

  // Try to parse as numeric comment ID, otherwise treat as anchor text
  const commentId = /^\d+$/.test(commentIdStr)
    ? parseInt(commentIdStr, 10)
    : commentIdStr;

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);

  doc.at(commentId).reply(text, { by: author });

  const idDisplay = typeof commentId === 'number' ? `comment #${commentId}` : `"${commentIdStr.slice(0, 40)}"`;
  console.log(C.dim(`Replying to ${idDisplay}`));
  console.log(C.dim(`By: ${author}`));

  const saveOpts = buildSaveOpts(options, `Reply to ${idDisplay}`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex figure <file> <position> <image> [caption] [--width] [--output]
 */
async function cmdFigure(positionals, options) {
  if (positionals.length < 3) {
    die('figure requires: <file> <position> <image-path> [caption]');
  }

  const [file, posStr, imagePath] = positionals;
  const caption = positionals[3] || options.caption || '';
  const { mode, anchor } = parsePosition(posStr);
  const width = options.width ? parseFloat(options.width) : 6;
  const author = resolveAuthor(options);

  // Validate image path
  const resolvedImage = path.resolve(imagePath);
  if (!fs.existsSync(resolvedImage)) {
    die(`Image file not found: ${resolvedImage}`);
  }

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);

  if (mode === 'after') {
    doc.after(anchor).figure(resolvedImage, caption, { width });
  } else {
    doc.before(anchor).figure(resolvedImage, caption, { width });
  }

  console.log(C.dim(`Inserting figure ${mode} "${anchor.slice(0, 40)}${anchor.length > 40 ? '...' : ''}"`));
  console.log(C.dim(`Image: ${path.basename(imagePath)}` + (caption ? ` | Caption: ${caption.slice(0, 40)}` : '')));
  console.log(C.dim(`Width: ${width} inches`));

  const saveOpts = buildSaveOpts(options, `Insert figure ${mode} "${anchor.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex table <file> <position> <json-file> [--style] [--caption] [--output]
 *
 * JSON file format: [["H1","H2"],["v1","v2"]]
 */
async function cmdTable(positionals, options) {
  if (positionals.length < 3) {
    die('table requires: <file> <position> <json-file>');
  }

  const [file, posStr, jsonPath] = positionals;
  const { mode, anchor } = parsePosition(posStr);
  const style = options.style || 'booktabs';
  const caption = options.caption || '';
  const author = resolveAuthor(options);

  // Read and parse JSON data
  const resolvedJson = path.resolve(jsonPath);
  if (!fs.existsSync(resolvedJson)) {
    die(`JSON file not found: ${resolvedJson}`);
  }

  let data;
  try {
    const raw = fs.readFileSync(resolvedJson, 'utf-8');
    data = JSON.parse(raw);
  } catch (err) {
    die(`Failed to parse JSON: ${err.message}`);
  }

  if (!Array.isArray(data) || data.length === 0 || !Array.isArray(data[0])) {
    die('Table JSON must be a 2D array: [["H1","H2"],["v1","v2"]]');
  }

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);

  const tableOpts = { style };
  if (caption) {
    tableOpts.caption = caption;
  }

  if (mode === 'after') {
    doc.after(anchor).table(data, tableOpts);
  } else {
    doc.before(anchor).table(data, tableOpts);
  }

  console.log(C.dim(`Inserting ${data.length}x${data[0].length} table ${mode} "${anchor.slice(0, 40)}${anchor.length > 40 ? '...' : ''}"`));
  console.log(C.dim(`Style: ${style}` + (caption ? ` | Caption: ${caption.slice(0, 40)}` : '')));

  const saveOpts = buildSaveOpts(options, `Insert table ${mode} "${anchor.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex cite <file> [--zotero-key KEY] [--zotero-user ID] [--collection ID] [--list] [--output]
 *
 * Inject Zotero citations or list found citation patterns.
 */
async function cmdCite(positionals, options) {
  if (positionals.length < 1) {
    die('cite requires: <file>\n  Use --list to show patterns, or --zotero-key + --zotero-user to inject');
  }

  const [file] = positionals;
  const docex = require('../src/docex');
  const doc = docex(file);

  if (options.list) {
    // List-only mode: find citation patterns and print them
    const cites = await doc.citations();
    console.log(C.bold(`Citation patterns (${cites.length}):\n`));

    for (const c of cites) {
      const typeTag = c.pattern === 'narrative' ? C.cyan('narrative') : C.cyan('parenthetical');
      console.log(`  ${C.dim(String(c.paragraph).padStart(3))}  ${c.text}  ${typeTag}`);
      console.log(`        ${C.dim('authors: ' + c.authors + '  year: ' + c.year)}`);
    }

    if (cites.length === 0) {
      console.log(C.dim('  No citation patterns found.'));
    }

    doc.discard();
    return;
  }

  // Injection mode: requires Zotero credentials
  if (!options['zotero-key'] || !options['zotero-user']) {
    die('cite requires --zotero-key and --zotero-user for injection.\n  Use --list to just list found patterns.');
  }

  const injectOpts = {
    zoteroApiKey: options['zotero-key'],
    zoteroUserId: options['zotero-user'],
  };
  if (options.collection) {
    injectOpts.collectionId = options.collection;
  }

  console.log(C.dim('Finding citation patterns...'));
  const result = await doc.injectCitations(injectOpts);

  console.log('');
  console.log(C.green('  Found:    ') + result.found + ' citation patterns');
  console.log(C.green('  Matched:  ') + result.matched + ' to Zotero items');
  console.log(C.green('  Injected: ') + result.injected + ' field codes');

  if (result.unmatched.length > 0) {
    console.log(C.yellow('  Unmatched:'));
    for (const u of result.unmatched) {
      console.log(C.yellow('    - ') + u);
    }
  }

  const saveOpts = buildSaveOpts(options, 'Inject Zotero citations');
  const saveResult = await doc.save(saveOpts);
  printResult(saveResult);
}

/**
 * docex list <file> [type]
 *
 * type: paragraphs (default) | headings | comments | figures
 */
async function cmdList(positionals, options) {
  if (positionals.length < 1) {
    die('list requires: <file> [type]');
  }

  const [file] = positionals;
  const type = (positionals[1] || 'paragraphs').toLowerCase();

  const docex = require('../src/docex');
  const doc = docex(file);

  switch (type) {
    case 'paragraphs':
    case 'paras':
    case 'p': {
      const items = await doc.paragraphs();
      console.log(C.bold(`Paragraphs (${items.length}):\n`));
      for (const p of items) {
        const styleTag = p.style ? C.dim(` [${p.style}]`) : '';
        const preview = p.text.slice(0, 100) + (p.text.length > 100 ? '...' : '');
        console.log(`  ${C.cyan(String(p.index).padStart(3))}  ${preview}${styleTag}`);
      }
      break;
    }

    case 'headings':
    case 'h': {
      const items = await doc.headings();
      console.log(C.bold(`Headings (${items.length}):\n`));
      for (const h of items) {
        const indent = '  '.repeat(h.level);
        console.log(`  ${C.cyan('H' + h.level)} ${indent}${h.text}  ${C.dim('[para ' + h.index + ']')}`);
      }
      break;
    }

    case 'comments':
    case 'c': {
      const items = await doc.comments();
      console.log(C.bold(`Comments (${items.length}):\n`));
      for (const c of items) {
        console.log(`  ${C.cyan('#' + c.id)}  ${C.bold(c.author)} ${C.dim(c.date)}`);
        console.log(`        ${c.text}`);
        console.log('');
      }
      break;
    }

    case 'figures':
    case 'images':
    case 'f': {
      const items = await doc.figures();
      console.log(C.bold(`Figures (${items.length}):\n`));
      for (const f of items) {
        const dims = (f.width && f.height)
          ? C.dim(` (${f.width}x${f.height} EMU)`)
          : '';
        console.log(`  ${C.cyan(f.rId)}  ${f.filename || '(embedded)'}${dims}`);
        if (f.caption) {
          console.log(`        ${C.dim(f.caption)}`);
        }
      }
      break;
    }

    default:
      die(`Unknown list type: "${type}". Choose: paragraphs, headings, comments, figures`);
  }

  // Clean up workspace (list is read-only, no save needed)
  doc.discard();
}

/**
 * docex latex <file> [--output <path>] [--doc-class <cls>] [--bib-file <name>] [--packages <list>]
 *
 * Read-only export: converts the document to LaTeX and writes to stdout or file.
 */
async function cmdLatex(positionals, options) {
  if (positionals.length < 1) {
    die('latex requires: <file>');
  }

  const [file] = positionals;

  const docex = require('../src/docex');
  const doc = docex(file);

  const latexOpts = {};
  if (options['doc-class']) latexOpts.documentClass = options['doc-class'];
  if (options['bib-file']) latexOpts.bibFile = options['bib-file'];
  if (options.packages) latexOpts.packages = options.packages.split(',').map(s => s.trim());

  const tex = await doc.toLatex(latexOpts);

  if (options.output) {
    const outputPath = path.resolve(options.output);
    fs.writeFileSync(outputPath, tex, 'utf-8');
    console.error(C.green('  Wrote: ') + outputPath + ' (' + (tex.length / 1024).toFixed(1) + ' KB)');
  } else {
    process.stdout.write(tex);
  }

  // Clean up workspace (read-only, no save needed)
  doc.discard();
}

// ============================================================================
// Main
// ============================================================================

async function main() {
  const { command, positionals, options } = parseArgs(process.argv.slice(2));

  // Handle --version and --help before anything else
  if (options.version) {
    console.log('docex ' + getVersion());
    process.exit(0);
  }

  if (options.help || command === 'help' || command === '') {
    printHelp();
    process.exit(command === '' && !options.help ? 1 : 0);
  }

  // Dispatch to command handler
  try {
    switch (command) {
      case 'replace':
        await cmdReplace(positionals, options);
        break;

      case 'insert':
        await cmdInsert(positionals, options);
        break;

      case 'delete':
      case 'del':
      case 'rm':
        await cmdDelete(positionals, options);
        break;

      case 'comment':
        await cmdComment(positionals, options);
        break;

      case 'reply':
        await cmdReply(positionals, options);
        break;

      case 'figure':
      case 'fig':
        await cmdFigure(positionals, options);
        break;

      case 'table':
      case 'tbl':
        await cmdTable(positionals, options);
        break;

      case 'cite':
      case 'citations':
        await cmdCite(positionals, options);
        break;

      case 'latex':
      case 'tex':
        await cmdLatex(positionals, options);
        break;

      case 'list':
      case 'ls':
        await cmdList(positionals, options);
        break;

      default:
        die(`Unknown command: "${command}". Run 'docex --help' for usage.`);
    }
  } catch (err) {
    die(err.message);
  }
}

main();
