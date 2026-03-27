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
  green:         (s) => `\x1b[32m${s}\x1b[0m`,
  red:           (s) => `\x1b[31m${s}\x1b[0m`,
  yellow:        (s) => `\x1b[33m${s}\x1b[0m`,
  bold:          (s) => `\x1b[1m${s}\x1b[0m`,
  dim:           (s) => `\x1b[2m${s}\x1b[0m`,
  cyan:          (s) => `\x1b[36m${s}\x1b[0m`,
  redStrike:     (s) => `\x1b[31m\x1b[9m${s}\x1b[0m`,
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
  const booleanFlags = new Set(['--untracked', '--help', '--version', '--list', '--json', '--dry-run']);
  const valueFlags = new Set(['--author', '--by', '--output', '--width', '--style', '--caption', '--safe', '--zotero-key', '--zotero-user', '--collection', '--doc-class', '--bib-file', '--packages', '--title', '--keywords', '--color', '--preset', '--vars']);

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
  if (options['dry-run'] || options.safe) {
    const opts = {};
    if (options.output) opts.outputPath = options.output;
    if (options.safe) {
      opts.safeModify = options.safe;
      opts.description = description || 'docex CLI edit';
    }
    if (options['dry-run']) opts.dryRun = true;
    return opts;
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
  if (result.dryRun) {
    console.log('');
    console.log(C.yellow('  [DRY RUN] No file written.'));
    console.log(C.yellow('  Target: ') + result.path);
    console.log(C.yellow('  Paras:  ') + result.paragraphCount + ' paragraphs');
    return;
  }

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
  ${C.cyan('accept')}     <file> [id]                     Accept tracked changes (all or by ID)
  ${C.cyan('reject')}     <file> [id]                     Reject tracked changes (all or by ID)
  ${C.cyan('clean')}      <file>                           Accept all changes, remove all comments
  ${C.cyan('revisions')}  <file>                           List tracked changes
  ${C.cyan('bold')}       <file> <text>                   Make text bold
  ${C.cyan('italic')}     <file> <text>                   Make text italic
  ${C.cyan('highlight')}  <file> <text> [--color C]       Highlight text (default: yellow)
  ${C.cyan('footnote')}   <file> <anchor> <text>          Add footnote at anchor text
  ${C.cyan('count')}      <file>                           Word count
  ${C.cyan('meta')}       <file> [--title X] [--author X] [--keywords X]  Get/set metadata
  ${C.cyan('diff')}       <file1> <file2>                  Compare two documents
  ${C.cyan('doctor')}     <file>                           Run diagnostic checks
  ${C.cyan('init')}                                        Create .docexrc in current directory

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
  --title <text>      Document title (for meta command)
  --keywords <text>   Document keywords (for meta command)
  --color <color>     Color name for highlight command (default: yellow)
  --dry-run           Preview changes without writing (all mutating commands)
  --json              Output as JSON (for list/meta/count commands)
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

  ${C.dim('# Accept all tracked changes')}
  docex accept manuscript.docx

  ${C.dim('# Highlight text in yellow')}
  docex highlight manuscript.docx "important finding" --color yellow

  ${C.dim('# Add a footnote')}
  docex footnote manuscript.docx "platform governance" "See Gorwa 2019 for details."

  ${C.dim('# Word count')}
  docex count manuscript.docx

  ${C.dim('# Get/set metadata')}
  docex meta manuscript.docx
  docex meta manuscript.docx --title "New Title" --author "Author Name"

  ${C.dim('# Compare two documents')}
  docex diff original.docx revised.docx --output diff.docx

  ${C.dim('# List headings')}
  docex list manuscript.docx headings

  ${C.dim('# List tracked changes')}
  docex revisions manuscript.docx

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
  const jsonOutput = !!options.json;

  const docex = require('../src/docex');
  const doc = docex(file);

  switch (type) {
    case 'paragraphs':
    case 'paras':
    case 'p': {
      const items = await doc.paragraphs();
      if (jsonOutput) {
        console.log(JSON.stringify(items, null, 2));
        break;
      }
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
      if (jsonOutput) {
        console.log(JSON.stringify(items, null, 2));
        break;
      }
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
      if (jsonOutput) {
        console.log(JSON.stringify(items, null, 2));
        break;
      }
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
      if (jsonOutput) {
        console.log(JSON.stringify(items, null, 2));
        break;
      }
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

    case 'revisions':
    case 'changes':
    case 'r': {
      const items = await doc.revisions();
      if (jsonOutput) {
        console.log(JSON.stringify(items, null, 2));
        break;
      }
      console.log(C.bold(`Tracked changes (${items.length}):\n`));
      for (const r of items) {
        const typeTag = r.type === 'insertion' ? C.green('ins') : C.red('del');
        console.log(`  ${C.cyan('#' + String(r.id).padStart(3))}  ${typeTag}  ${C.bold(r.author)} ${C.dim(r.date)}`);
        const preview = r.text.slice(0, 80) + (r.text.length > 80 ? '...' : '');
        console.log(`        ${preview}`);
        console.log('');
      }
      break;
    }

    case 'footnotes':
    case 'fn': {
      const items = await doc.footnotes();
      if (jsonOutput) {
        console.log(JSON.stringify(items, null, 2));
        break;
      }
      console.log(C.bold(`Footnotes (${items.length}):\n`));
      for (const fn of items) {
        console.log(`  ${C.cyan('#' + String(fn.id).padStart(3))}  ${fn.text}`);
      }
      if (items.length === 0) {
        console.log(C.dim('  No footnotes found.'));
      }
      break;
    }

    default:
      die(`Unknown list type: "${type}". Choose: paragraphs, headings, comments, figures, revisions, footnotes`);
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
// New v0.2 command handlers
// ============================================================================

/**
 * docex accept <file> [id]
 * Accept tracked changes (all or by ID).
 */
async function cmdAccept(positionals, options) {
  if (positionals.length < 1) {
    die('accept requires: <file> [id]');
  }

  const [file] = positionals;
  const id = positionals[1] ? parseInt(positionals[1], 10) : undefined;

  const docex = require('../src/docex');
  const doc = docex(file);

  if (id !== undefined) {
    console.log(C.dim(`Accepting change #${id}...`));
    await doc.accept(id);
  } else {
    console.log(C.dim('Accepting all tracked changes...'));
    await doc.accept();
  }

  const saveOpts = buildSaveOpts(options, id ? `Accept change #${id}` : 'Accept all changes');
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex reject <file> [id]
 * Reject tracked changes (all or by ID).
 */
async function cmdReject(positionals, options) {
  if (positionals.length < 1) {
    die('reject requires: <file> [id]');
  }

  const [file] = positionals;
  const id = positionals[1] ? parseInt(positionals[1], 10) : undefined;

  const docex = require('../src/docex');
  const doc = docex(file);

  if (id !== undefined) {
    console.log(C.dim(`Rejecting change #${id}...`));
    await doc.reject(id);
  } else {
    console.log(C.dim('Rejecting all tracked changes...'));
    await doc.reject();
  }

  const saveOpts = buildSaveOpts(options, id ? `Reject change #${id}` : 'Reject all changes');
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex clean <file>
 * Accept all changes, remove all comments.
 */
async function cmdClean(positionals, options) {
  if (positionals.length < 1) {
    die('clean requires: <file>');
  }

  const [file] = positionals;
  const docex = require('../src/docex');
  const doc = docex(file);

  console.log(C.dim('Producing clean copy (accept all changes, remove comments)...'));
  await doc.cleanCopy();

  const saveOpts = buildSaveOpts(options, 'Clean copy: accept all changes, remove comments');
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex revisions <file>
 * List tracked changes.
 */
async function cmdRevisions(positionals, options) {
  if (positionals.length < 1) {
    die('revisions requires: <file>');
  }

  const [file] = positionals;
  const docex = require('../src/docex');
  const doc = docex(file);

  const revs = await doc.revisions();

  if (options.json) {
    console.log(JSON.stringify(revs, null, 2));
  } else {
    console.log(C.bold(`Tracked changes (${revs.length}):\n`));
    for (const r of revs) {
      const typeTag = r.type === 'insertion' ? C.green('ins') : C.red('del');
      console.log(`  ${C.cyan('#' + String(r.id).padStart(3))}  ${typeTag}  ${C.dim(r.author + '  ' + r.date)}`);
      const preview = r.text.slice(0, 80) + (r.text.length > 80 ? '...' : '');
      // Colored text: red strikethrough for deletions, green for insertions
      if (r.type === 'deletion') {
        // Red + strikethrough (ANSI: \x1b[9m = strikethrough)
        console.log(`        \x1b[31m\x1b[9m${preview}\x1b[0m`);
      } else {
        // Green for insertions
        console.log(`        ${C.green(preview)}`);
      }
      console.log('');
    }
    if (revs.length === 0) {
      console.log(C.dim('  No tracked changes found.'));
    }
  }

  doc.discard();
}

/**
 * docex bold <file> <text> [--author] [--output]
 */
async function cmdBold(positionals, options) {
  if (positionals.length < 2) {
    die('bold requires: <file> <text>');
  }

  const [file, text] = positionals;
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);
  doc.bold(text);

  console.log(C.dim(`Making bold: "${text.slice(0, 60)}${text.length > 60 ? '...' : ''}"`));

  const saveOpts = buildSaveOpts(options, `Bold: "${text.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex italic <file> <text> [--author] [--output]
 */
async function cmdItalic(positionals, options) {
  if (positionals.length < 2) {
    die('italic requires: <file> <text>');
  }

  const [file, text] = positionals;
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);
  doc.italic(text);

  console.log(C.dim(`Making italic: "${text.slice(0, 60)}${text.length > 60 ? '...' : ''}"`));

  const saveOpts = buildSaveOpts(options, `Italic: "${text.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex highlight <file> <text> [--color <color>] [--author] [--output]
 */
async function cmdHighlight(positionals, options) {
  if (positionals.length < 2) {
    die('highlight requires: <file> <text>');
  }

  const [file, text] = positionals;
  const color = options.color || 'yellow';
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);
  doc.highlight(text, color);

  console.log(C.dim(`Highlighting "${text.slice(0, 60)}${text.length > 60 ? '...' : ''}" in ${color}`));

  const saveOpts = buildSaveOpts(options, `Highlight: "${text.slice(0, 40)}" in ${color}`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex footnote <file> <anchor> <text> [--author] [--output]
 */
async function cmdFootnote(positionals, options) {
  if (positionals.length < 3) {
    die('footnote requires: <file> <anchor-text> <footnote-text>');
  }

  const [file, anchor, text] = positionals;
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file);
  doc.author(author);
  doc.at(anchor).footnote(text);

  console.log(C.dim(`Adding footnote at "${anchor.slice(0, 40)}${anchor.length > 40 ? '...' : ''}"`));
  console.log(C.dim(`Note: ${text.slice(0, 60)}${text.length > 60 ? '...' : ''}`));

  const saveOpts = buildSaveOpts(options, `Footnote at: "${anchor.slice(0, 40)}"`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

/**
 * docex count <file> [--json]
 * Word count.
 */
async function cmdCount(positionals, options) {
  if (positionals.length < 1) {
    die('count requires: <file>');
  }

  const [file] = positionals;
  const docex = require('../src/docex');
  const doc = docex(file);

  const wc = await doc.wordCount();

  if (options.json) {
    console.log(JSON.stringify(wc, null, 2));
  } else {
    console.log(C.bold('Word count:\n'));
    console.log(C.green('  Total:     ') + wc.total);
    console.log(C.green('  Body:      ') + wc.body);
    console.log(C.green('  Headings:  ') + wc.headings);
    console.log(C.green('  Abstract:  ') + wc.abstract);
    console.log(C.green('  Captions:  ') + wc.captions);
    console.log(C.green('  Footnotes: ') + wc.footnotes);
  }

  doc.discard();
}

/**
 * docex meta <file> [--title X] [--author X] [--keywords X] [--json]
 * Get or set metadata.
 */
async function cmdMeta(positionals, options) {
  if (positionals.length < 1) {
    die('meta requires: <file>');
  }

  const [file] = positionals;
  const docex = require('../src/docex');
  const doc = docex(file);

  // Determine if we're setting or getting
  const setProps = {};
  let isSet = false;
  if (options.title !== undefined) { setProps.title = options.title; isSet = true; }
  if (options.author !== undefined) { setProps.creator = options.author; isSet = true; }
  if (options.by !== undefined) { setProps.creator = options.by; isSet = true; }
  if (options.keywords !== undefined) { setProps.keywords = options.keywords; isSet = true; }

  if (isSet) {
    // Set metadata
    await doc.metadata(setProps);

    const setDesc = Object.entries(setProps).map(([k, v]) => `${k}="${v}"`).join(', ');
    console.log(C.dim(`Setting metadata: ${setDesc}`));

    const saveOpts = buildSaveOpts(options, `Set metadata: ${setDesc}`);
    const result = await doc.save(saveOpts);
    printResult(result);
  } else {
    // Get metadata
    const meta = await doc.metadata();

    if (options.json) {
      console.log(JSON.stringify(meta, null, 2));
    } else {
      console.log(C.bold('Document metadata:\n'));
      for (const [key, val] of Object.entries(meta)) {
        if (val) {
          console.log(C.green('  ' + key.padEnd(16)) + val);
        }
      }
      const empty = Object.values(meta).every(v => !v);
      if (empty) {
        console.log(C.dim('  No metadata properties found.'));
      }
    }
    doc.discard();
  }
}

/**
 * docex diff <file1> <file2> [--author] [--output]
 * Compare two documents.
 */
async function cmdDiff(positionals, options) {
  if (positionals.length < 2) {
    die('diff requires: <file1> <file2>');
  }

  const [file1, file2] = positionals;
  const author = resolveAuthor(options);

  const docex = require('../src/docex');
  const doc = docex(file1);
  doc.author(author);

  console.log(C.dim(`Comparing: ${path.basename(file1)} vs ${path.basename(file2)}`));
  const stats = await doc.diff(file2, { author });

  console.log('');
  console.log(C.green('  Added:     ') + stats.added + ' paragraphs');
  console.log(C.red('  Removed:   ') + stats.removed + ' paragraphs');
  console.log(C.yellow('  Modified:  ') + stats.modified + ' paragraphs');
  console.log(C.dim('  Unchanged: ') + stats.unchanged + ' paragraphs');

  const saveOpts = buildSaveOpts(options, `Diff: ${path.basename(file1)} vs ${path.basename(file2)}`);
  const result = await doc.save(saveOpts);
  printResult(result);
}

// ============================================================================
// v0.3 command handlers: doctor, init
// ============================================================================

/**
 * docex doctor <file>
 * Run diagnostic checks on a .docx file.
 */
async function cmdDoctor(positionals, options) {
  if (positionals.length < 1) {
    die('doctor requires: <file>');
  }

  const [file] = positionals;
  const { Workspace } = require('../src/workspace');
  const { Doctor } = require('../src/doctor');

  const ws = Workspace.open(file);
  const output = Doctor.diagnose(ws);
  console.log(output);
  ws.cleanup();
}

/**
 * docex init
 * Create .docexrc in the current directory.
 */
async function cmdInit(positionals, options) {
  const targetDir = process.cwd();
  const rcPath = path.join(targetDir, '.docexrc');

  if (fs.existsSync(rcPath)) {
    console.log(C.yellow('  .docexrc already exists at: ') + rcPath);
    console.log(C.dim('  Delete it first if you want to reinitialize.'));
    return;
  }

  // Get author from git config
  let author = 'Unknown';
  try {
    author = require('child_process').execFileSync('git', ['config', 'user.name'], {
      stdio: ['pipe', 'pipe', 'pipe'],
      encoding: 'utf-8',
      timeout: 3000,
    }).trim() || 'Unknown';
  } catch (_) { /* ignore */ }

  const rc = {
    author,
    safeModify: '',
    style: 'academic',
    backup: true,
  };

  fs.writeFileSync(rcPath, JSON.stringify(rc, null, 2) + '\n', 'utf-8');
  console.log(C.green('  Created: ') + rcPath);
  console.log('');
  console.log(C.dim('  Contents:'));
  console.log(C.dim('  ' + JSON.stringify(rc, null, 2).replace(/\n/g, '\n  ')));
}

// ============================================================================
// v0.3 Academic command handlers
// ============================================================================

/**
 * docex style <file> <preset> [--output]
 * Apply a journal style preset.
 */
async function cmdStyle(positionals, options) {
  if (positionals.length < 2) {
    die('style requires: <file> <preset-name>\n  Available: academic, polcomm, apa7, jcmc, joc');
  }

  const [file, presetName] = positionals;

  const docex = require('../src/docex');
  const doc = docex(file);

  console.log(C.dim(`Applying style preset: ${presetName}`));
  const result = await doc.style(presetName);

  for (const change of result.changes) {
    console.log(C.green('  + ') + change);
  }

  const saveOpts = buildSaveOpts(options, `Apply style: ${presetName}`);
  const saveResult = await doc.save(saveOpts);
  printResult(saveResult);
}

/**
 * docex verify <file> <preset> [--json]
 * Validate document against journal requirements.
 */
async function cmdVerify(positionals, options) {
  if (positionals.length < 2) {
    die('verify requires: <file> <preset-name>\n  Available: academic, polcomm, apa7, jcmc, joc');
  }

  const [file, presetName] = positionals;

  const docex = require('../src/docex');
  const doc = docex(file);

  const result = await doc.verify(presetName);

  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
  } else {
    const status = result.pass ? C.green('PASS') : C.red('FAIL');
    console.log(C.bold(`Verification (${presetName}): `) + status);

    if (result.errors.length > 0) {
      console.log(C.red('\n  Errors:'));
      for (const e of result.errors) {
        console.log(C.red('    - ') + e);
      }
    }

    if (result.warnings.length > 0) {
      console.log(C.yellow('\n  Warnings:'));
      for (const w of result.warnings) {
        console.log(C.yellow('    - ') + w);
      }
    }

    if (result.pass && result.warnings.length === 0) {
      console.log(C.green('\n  All checks passed.'));
    }
  }

  doc.discard();
}

/**
 * docex anonymize <file> [--output]
 * Remove author names for blind review.
 */
async function cmdAnonymize(positionals, options) {
  if (positionals.length < 1) {
    die('anonymize requires: <file>');
  }

  const [file] = positionals;

  const docex = require('../src/docex');
  const doc = docex(file);

  console.log(C.dim('Anonymizing document for blind review...'));
  const result = await doc.anonymize();

  if (result.authorsRemoved.length > 0) {
    console.log(C.green('  Authors removed:'));
    for (const a of result.authorsRemoved) {
      console.log(C.green('    - ') + a);
    }
    console.log(C.green('  Locations: ') + result.locations.join(', '));
  } else {
    console.log(C.dim('  No author names found to remove.'));
  }

  const saveOpts = buildSaveOpts(options, 'Anonymize for blind review');
  const saveResult = await doc.save(saveOpts);
  printResult(saveResult);
}

/**
 * docex expand <file> [--vars "KEY=VAL,KEY2=VAL2"] [--output]
 * Expand {{VAR}} patterns in the document.
 */
async function cmdExpand(positionals, options) {
  if (positionals.length < 1) {
    die('expand requires: <file> [--vars "KEY=VAL,KEY2=VAL2"]');
  }

  const [file] = positionals;

  const docex = require('../src/docex');
  const doc = docex(file);

  // Parse variables from --vars option
  const variables = {};
  if (options.vars) {
    const pairs = options.vars.split(',');
    for (const pair of pairs) {
      const eqIdx = pair.indexOf('=');
      if (eqIdx > 0) {
        const key = pair.slice(0, eqIdx).trim();
        const val = pair.slice(eqIdx + 1).trim();
        variables[key] = val;
      }
    }
  }

  if (Object.keys(variables).length === 0) {
    // List mode: show all variables in the document
    const vars = await doc.listVariables();
    console.log(C.bold(`Variables found (${vars.length}):\n`));
    for (const v of vars) {
      console.log(`  ${C.cyan('{{' + v.name + '}}')}  para ${v.paragraph}  ${C.dim(v.context)}`);
    }
    if (vars.length === 0) {
      console.log(C.dim('  No {{VAR}} patterns found.'));
    }
    doc.discard();
    return;
  }

  console.log(C.dim(`Expanding ${Object.keys(variables).length} variable(s)...`));
  const count = await doc.expand(variables);
  console.log(C.green(`  Expanded: `) + count + ' occurrence(s)');

  const saveOpts = buildSaveOpts(options, `Expand ${Object.keys(variables).length} variables`);
  const saveResult = await doc.save(saveOpts);
  printResult(saveResult);
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

      case 'accept':
        await cmdAccept(positionals, options);
        break;

      case 'reject':
        await cmdReject(positionals, options);
        break;

      case 'clean':
        await cmdClean(positionals, options);
        break;

      case 'revisions':
      case 'changes':
        await cmdRevisions(positionals, options);
        break;

      case 'bold':
        await cmdBold(positionals, options);
        break;

      case 'italic':
        await cmdItalic(positionals, options);
        break;

      case 'highlight':
        await cmdHighlight(positionals, options);
        break;

      case 'footnote':
      case 'fn':
        await cmdFootnote(positionals, options);
        break;

      case 'count':
      case 'wc':
        await cmdCount(positionals, options);
        break;

      case 'meta':
      case 'metadata':
        await cmdMeta(positionals, options);
        break;

      case 'diff':
      case 'compare':
        await cmdDiff(positionals, options);
        break;

      case 'doctor':
      case 'check':
        await cmdDoctor(positionals, options);
        break;

      case 'init':
        await cmdInit(positionals, options);
        break;

      case 'style':
        await cmdStyle(positionals, options);
        break;

      case 'verify':
      case 'validate':
        await cmdVerify(positionals, options);
        break;

      case 'anonymize':
      case 'anon':
        await cmdAnonymize(positionals, options);
        break;

      case 'expand':
        await cmdExpand(positionals, options);
        break;

      default:
        die(`Unknown command: "${command}". Run 'docex --help' for usage.`);
    }
  } catch (err) {
    die(err.message);
  }
}

main();
