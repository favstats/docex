# docex

**LaTeX for .docx.** A zero-dependency Node.js library that makes .docx files programmable. You describe what you want to do (replace text, add comments, insert figures, apply journal styles) and docex handles the OOXML plumbing.

It was built for academic manuscript workflows: writing papers, responding to peer review, formatting for journal submission. But it works on any .docx file.

v0.4.0 / 28 modules / 403 tests / 14,101 lines of source / zero external dependencies

## What It Does

- Opens a .docx, lets you edit it programmatically, saves it back
- Tracked changes by default (shows as strikethrough + insertion in Word/OnlyOffice)
- Comments with threading and replies (manages 5 OOXML files)
- Figures: insert, replace, list (PNG/JPEG, auto-dimensions)
- Tables: insert with booktabs or plain style
- Inline formatting: bold, italic, underline, highlight, color, strikethrough, superscript, subscript, small caps, code
- Footnotes
- Bullet and numbered lists
- Accept/reject tracked changes, clean copy generation
- Document diff (compare two .docx files, produce tracked changes)
- Citations: detect patterns, inject Zotero field codes
- LaTeX export (built-in converter) and import (via Pandoc)
- HTML and Markdown export (via Pandoc)
- Journal style presets (academic, polcomm, apa7, jcmc, joc)
- Submission validation (word count, abstract length, figure resolution, margins, fonts)
- Anonymize/deanonymize for blind review
- Response letter generation for R&R
- Document templates (create from scratch with journal formatting)
- Batch operations across multiple documents
- Watch mode (live recompile on .tex changes)
- Stable addressing via OOXML paraId (survives other edits)
- Cross-references and auto-numbering for figures and tables
- Variables and macro expansion ({{VAR_NAME}} patterns)
- Word count, metadata, document stats, contributor tracking, timeline
- Document health checks (valid zip, relationships, orphaned media, heading hierarchy)
- Auto-backup, lock files, fuzzy text matching
- Single unzip/rezip cycle per save (no corruption from repeated zip operations)
- .docexrc configuration files for project defaults

## Install

```bash
git clone https://github.com/favstats/docex.git
cd docex
# No npm install needed -- zero dependencies
```

Requires Node.js 18+ (uses `node:test` and `node:zlib`).

Or via npm (once published):

```bash
npm install docex
```

## Quick Start (API)

```js
const docex = require('docex');
const doc = docex("manuscript.docx");
doc.author("Fabio Votta");

doc.replace("old text", "new text");                     // tracked by default
doc.after("Methods").insert("New paragraph.");            // position selector
doc.at("enforcement gap").comment("Cite Suzor 2019", { by: "Reviewer 2" });
doc.after("Results").figure("fig03.png", "Figure 3. Status");
doc.after("Results").table([["Party","Ads"],["PAX","117"]], { style: "booktabs" });

await doc.save();  // single rezip + auto-verify
```

## Quick Start (CLI)

```bash
docex replace manuscript.docx "old" "new" --author "Fabio"
docex insert manuscript.docx "after:Methods" "New paragraph."
docex comment manuscript.docx "anchor text" "Note" --by "Reviewer 2"
docex figure manuscript.docx "after:Results" fig03.png --caption "Figure 3."
docex list manuscript.docx headings
docex list manuscript.docx comments
```

For local use without a global install, prefix with `node`:

```bash
node cli/docex-cli.js replace manuscript.docx "old" "new" --author "Fabio"
```

## Module List

28 source modules in `src/`:

| Module | Lines | Description |
|--------|------:|-------------|
| `docex.js` | 1,421 | Main API: DocexEngine, PositionSelector, factory function |
| `paragraphs.js` | 1,292 | Replace, insert, delete text (tracked and untracked), word count |
| `latex.js` | 993 | OOXML to LaTeX converter |
| `comments.js` | 850 | Add, list, reply, resolve, remove, export comments (5-file management) |
| `workspace.js` | 793 | Zip/unzip lifecycle, temp directory, save, verify, backup, lock files |
| `citations.js` | 703 | Detect citation patterns, inject Zotero field codes |
| `handle.js` | 643 | ParagraphHandle and RunHandle for stable paraId-based addressing |
| `figures.js` | 640 | Insert, replace, list images with auto-dimensions and relationships |
| `xml.js` | 474 | Lightweight XML parser, serializer, escaping, paragraph extraction |
| `docmap.js` | 453 | Document map, paraId injection, structure tree, find, explain |
| `revisions.js` | 451 | List, accept, reject tracked changes; clean copy; contributors; timeline |
| `diff.js` | 439 | Compare two .docx files, produce tracked changes |
| `presets.js` | 437 | Journal style presets (fonts, margins, spacing, headers) |
| `formatting.js` | 434 | Inline formatting: bold, italic, underline, highlight, color, code, etc. |
| `lists.js` | 389 | Bullet and numbered list insertion with numbering definitions |
| `template.js` | 363 | Create .docx from scratch with title page, abstract, sections |
| `doctor.js` | 358 | Document health: valid zip, relationships, orphaned media, paraId uniqueness |
| `compile.js` | 347 | LaTeX pipeline: .tex to .docx (via Pandoc), .docx to .tex, watch mode |
| `tables.js` | 329 | Insert tables with booktabs or plain (grid) style |
| `response-letter.js` | 305 | Generate R&R response letters grouped by reviewer |
| `crossref.js` | 298 | Cross-references, labels, SEQ/REF field codes, auto-numbering |
| `textmap.js` | 286 | Map plain-text offsets to XML runs (solves the run-splitting problem) |
| `verify.js` | 259 | Validate against journal requirements (word count, margins, fonts, etc.) |
| `footnotes.js` | 258 | List and add footnotes |
| `submission.js` | 245 | Anonymize/deanonymize for blind review, highlighted changes |
| `metadata.js` | 223 | Dublin Core metadata: read/write title, creator, keywords, etc. |
| `macros.js` | 217 | Variable definition and {{VAR_NAME}} expansion |
| `batch.js` | 201 | Apply operations to multiple .docx files in one pass |

## API Reference

### Opening a Document

```js
const docex = require('docex');

// Factory function (preferred)
const doc = docex("manuscript.docx");

// Alternative form
const doc = docex.open("manuscript.docx");
```

Both return a `DocexEngine` instance.

### Configuration

All configuration methods are chainable and return `doc`.

```js
doc.author("Fabio Votta");              // Set author for all operations
doc.date("2026-03-15T12:00:00Z");       // Set date (default: now)
doc.tracked();                           // Enable tracked changes (default)
doc.untracked();                         // Disable tracked changes
```

The `.docexrc` file in the document's directory (or `~/.docexrc`) can set defaults:

```json
{ "author": "Fabio Votta" }
```

### Position Selectors

Position selectors target a location in the document by its text content. They return a `PositionSelector` object.

```js
doc.at("some phrase")         // AT the text (for comments, formatting, footnotes)
doc.after("Methods")          // AFTER a heading or text
doc.before("Conclusion")      // BEFORE a heading or text
doc.afterHeading("Methods")   // AFTER a heading specifically
doc.afterText("some text")    // AFTER body text specifically
doc.id("3A7F2B1C")           // By paraId (returns ParagraphHandle)
```

#### PositionSelector methods

```js
// Content insertion
doc.after("Methods").insert("New paragraph text.", opts?)
doc.after("Results").figure("fig03.png", "Figure 3. Caption.", opts?)
doc.after("Results").table([["Col1","Col2"],["a","b"]], opts?)
doc.after("Results").bulletList(["Item 1", "Item 2"], opts?)
doc.after("Results").numberedList(["First", "Second"], opts?)

// Comments
doc.at("anchor text").comment("Comment text.", { by: "Reviewer 2" })
doc.at("anchor text").reply("Reply text.", { by: "Fabio Votta" })

// Inline formatting
doc.at("text").bold()
doc.at("text").italic()
doc.at("text").underline()
doc.at("text").strikethrough()
doc.at("text").superscript()
doc.at("text").subscript()
doc.at("text").smallCaps()
doc.at("text").code()
doc.at("text").color("red")
doc.at("text").highlight("yellow")

// Footnote
doc.at("anchor text").footnote("Footnote text.")
```

### Text Operations

```js
doc.replace("old text", "new text");                           // tracked
doc.replace("old text", "new text", { tracked: false });       // direct
doc.replace("old text", "new text", { author: "Someone" });   // override author
doc.replaceAll("old", "new");                                  // all occurrences
doc.delete("text to remove");                                  // tracked deletion
```

Direct formatting (without position selector):

```js
doc.bold("text");
doc.italic("text");
doc.highlight("text", "yellow");
doc.color("text", "red");
```

### Comments

```js
// Add
doc.comment("anchor text", "Comment body.", { by: "Reviewer 2" });
doc.at("anchor text").comment("Comment body.", { by: "Reviewer 2" });

// Reply
doc.at("original anchor").reply("Reply text.", { by: "Fabio Votta" });

// Read
const comments = await doc.comments();
// [{id, author, text, date, anchor}]

// Export
const csv = await doc.exportComments("csv");
const json = await doc.exportComments("json");
```

### Figures and Tables

```js
// Figures
doc.after("Results").figure("fig03.png", "Figure 3. Caption.", { width: 5 });
const figures = await doc.figures();
// [{rId, filename, width, height}]

// Tables
doc.after("Results").table(
  [["Party", "Ads", "Spend"], ["VVD", "342", "12,450"]],
  { style: "booktabs", caption: "Table 1. Ad counts." }
);
// style: "booktabs" (default) or "plain"
```

### Footnotes

```js
doc.at("anchor text").footnote("Footnote content.");
const footnotes = await doc.footnotes();
// [{id, text}]
```

### Lists

```js
doc.after("Methods").bulletList(["Item one", "Item two", "Item three"]);
doc.after("Methods").numberedList(["First step", "Second step"]);
```

### Revisions (Tracked Changes)

```js
const revisions = await doc.revisions();     // [{id, type, author, date, text}]
await doc.accept();                           // accept all
await doc.accept(5);                          // accept specific change
await doc.reject();                           // reject all
await doc.reject(3);                          // reject specific change
await doc.cleanCopy();                        // accept all changes, remove all comments
```

### Document Diff

```js
const result = await doc.diff("other-version.docx", { author: "Fabio Votta" });
// {added, removed, modified, unchanged}
// The current document now contains tracked changes showing the differences.
await doc.save("diff-output.docx");
```

### Citations

```js
// List citation patterns (no network)
const cites = await doc.citations();
// [{text, paragraph, pattern, authors, year}]

// Inject Zotero field codes
const result = await doc.injectCitations({
  zoteroApiKey: "...",
  zoteroUserId: "6875557",
  collectionId: "TUWJI72V",
  bibliography: true,
});
// {found, matched, injected, unmatched}
```

### Metadata and Word Count

```js
// Get metadata
const meta = await doc.metadata();
// {title, creator, subject, keywords, description, ...}

// Set metadata
await doc.metadata({ title: "My Paper", creator: "Fabio Votta", keywords: "elections, ads" });

// Word count
const wc = await doc.wordCount();
// {total, body, headings, abstract, captions, footnotes}

// Comprehensive stats
const stats = await doc.stats();
// {words, paragraphs, headings, figures, tables, citations, comments, revisions}

// Contributors
const contribs = await doc.contributors();
// [{name, changes, comments, lastActive}]

// Timeline
const events = await doc.timeline();
// [{date, type, author, text}]
```

### Journal Styles and Validation

```js
// Apply a style preset
await doc.style("polcomm");
// Built-in presets: academic, polcomm, apa7, jcmc, joc

// Define a custom preset
docex.defineStyle("myjournal", {
  font: "Garamond", size: 11, spacing: "double",
  margins: { top: 1, bottom: 1, left: 1.25, right: 1.25 },
});

// List available presets
const presets = docex.listStyles();

// Validate against journal requirements
const result = await doc.verify("polcomm");
// {pass, errors, warnings}
```

### Submission Helpers

```js
// Anonymize for blind review
const anon = await doc.anonymize();
// {authorsRemoved, locations}

// Deanonymize (restore author info)
const deanon = await doc.deanonymize();
// {restored, authors}

// Generate highlighted changes document
const hc = await doc.highlightedChanges();
// {insertions, deletions}
```

### Response Letters

```js
const comments = await doc.comments();

const responses = {
  1: { action: "agree", text: "We have added the citation.", changes: ["Added Suzor 2019 to p. 12"] },
  2: { action: "partial", text: "We clarified this but kept the framing.", changes: ["Revised paragraph 3 in Discussion"] },
  3: { action: "disagree", text: "We respectfully maintain our position because...", changes: [] },
};

const result = await docex.responseLetter(comments, responses, {
  title: "Platform Enforcement in Dutch Local Elections",
  journal: "Political Communication",
  authors: ["Fabio Votta", "Simon Munzert"],
  output: "response-letter.docx",
});
// {path, fileSize, paragraphCount, reviewers, commentsAddressed}
```

### Templates and Document Creation

```js
// Create from template
const result = await docex.fromTemplate({
  title: "Platform Enforcement in Dutch Local Elections",
  authors: [
    { name: "Fabio Votta", affiliation: "University of Amsterdam", email: "f.votta@uva.nl" },
  ],
  abstract: "This paper examines...",
  keywords: ["elections", "enforcement", "platforms"],
  sections: ["Introduction", "Literature Review", "Methods", "Results", "Discussion", "Conclusion"],
  journal: "polcomm",
  output: "new-manuscript.docx",
});

// Create a minimal empty .docx
const emptyPath = await docex.create("blank.docx");
```

### Batch Operations

```js
const batch = docex.batch(["paper1.docx", "paper2.docx", "paper3.docx"]);
batch.author("Fabio Votta");
batch.style("polcomm");
batch.replaceAll("old term", "new term");
const results = await batch.saveAll({ suffix: "_formatted" });
// [{path, fileSize, paragraphCount, error}]

// Batch verify
const verifyResults = await batch.verify("polcomm");
// [{path, result: {pass, errors, warnings}}]
```

### LaTeX Pipeline

Requires Pandoc to be installed.

```js
// Compile .tex to submission-ready .docx
const result = await docex.compile("paper.tex", {
  style: "polcomm",
  bibFile: "references.bib",
  cslFile: "apa.csl",
  output: "submission.docx",
});
// {path, fileSize, paragraphCount, style}

// Decompile .docx back to .tex (preserves tracked changes and comments)
const result = await docex.decompile("manuscript.docx", {
  output: "paper.tex",
});
// {path, tex, changes, comments}

// Watch mode: recompile on .tex changes
const watcher = docex.watch("paper.tex", { style: "polcomm" });
// ... edit paper.tex, docex recompiles automatically ...
watcher.close();  // stop watching
```

### Export Formats

```js
// To LaTeX (built-in converter, no Pandoc needed)
const tex = await doc.toLatex({ documentClass: "article", bibFile: "references" });

// To HTML (requires Pandoc)
const html = await doc.toHtml({ output: "paper.html" });

// To Markdown (requires Pandoc)
const md = await doc.toMarkdown({ output: "paper.md" });
```

### Document Health

```js
// Validate document integrity
const result = await doc.validate();
// {valid, errors, warnings}

// Comprehensive stats
const stats = await doc.stats();
```

### Stable Addressing

Every paragraph gets a unique `w14:paraId`. Operations on a paraId survive other edits to the document.

```js
// Get document map
const map = await doc.map();
// {sections, allParagraphs, allFigures, allTables, allComments}

// Get ParagraphHandle by paraId
const handle = doc.id("3A7F2B1C");
handle.replace("old", "new");
handle.comment("Note");
handle.bold();

// Find paragraphs by text
const matches = await doc.find("enforcement gap");
// [{id, index, section, context}]

// Document structure tree
const tree = await doc.structure();
// Human-readable tree string

// Explain XML around text
const explanation = await doc.explain("enforcement gap");
// Shows XML structure for debugging
```

### Cross-References and Auto-Numbering

```js
// Label a paragraph
await doc.label("3A7F2B1C", "fig:funnel");

// Reference a label
await doc.ref("fig:funnel", { insertAt: "4B8E3C2D" });

// Auto-number all figure and table captions
const counts = await doc.autoNumber();
// {figures, tables}

// List all labels
const labels = await doc.listLabels();
// [{name, type, number, paraId}]
```

### Variables and Macros

```js
// Define variables
await doc.define("NUM_ADS", "268,635");
await doc.define("PERIOD", "January-March 2026");

// Expand all {{VAR}} patterns
const count = await doc.expand({ NUM_ADS: "268,635", PERIOD: "January-March 2026" });

// List undefined variables
const vars = await doc.listVariables();
// [{name, paragraph, context}]
```

### Preview and Discard

```js
// Preview pending operations without executing
console.log(doc.preview());
// "3 pending operations:
//   1. replace 'old text' -> 'new text' (tracked, by Fabio Votta)
//   2. insert after 'Methods': 'New paragraph.' (tracked, by Fabio Votta)
//   3. comment at 'anchor': 'Note' (by Reviewer 2)"

// Discard all pending operations
doc.discard();
```

### Saving

```js
// Overwrite original
const result = await doc.save();

// Save to new file
const result = await doc.save("revised.docx");

// Save with safe-modify.sh protection (for manuscripts)
const result = await doc.save({
  safeModify: "/path/to/safe-modify.sh",
  description: "Fix typo in Methods"
});

// Dry run
const result = await doc.save({ dryRun: true });
```

Returns: `{ path, operations, paragraphCount, fileSize, verified }`

## CLI Reference

```
docex <command> <file> [arguments] [options]
```

### Commands

| Command | Aliases | Arguments | Description |
|---------|---------|-----------|-------------|
| `replace` | | `<file> <old> <new>` | Replace text (tracked by default) |
| `insert` | | `<file> <position> <text>` | Insert paragraph at position |
| `delete` | `del`, `rm` | `<file> <text>` | Delete text (tracked by default) |
| `comment` | | `<file> <anchor> <text>` | Add comment anchored to text |
| `reply` | | `<file> <comment-id> <text>` | Reply to existing comment |
| `figure` | `fig` | `<file> <position> <image>` | Insert figure at position |
| `table` | `tbl` | `<file> <position> <json-file>` | Insert table from JSON |
| `cite` | `citations` | `<file>` | List or inject Zotero citations |
| `list` | `ls` | `<file> [type]` | List paragraphs, headings, comments, figures, revisions, footnotes |
| `bold` | | `<file> <text>` | Make text bold |
| `italic` | | `<file> <text>` | Make text italic |
| `highlight` | | `<file> <text>` | Highlight text |
| `footnote` | `fn` | `<file> <anchor> <text>` | Add footnote |
| `accept` | | `<file> [id]` | Accept tracked changes (all or by ID) |
| `reject` | | `<file> [id]` | Reject tracked changes (all or by ID) |
| `clean` | | `<file>` | Accept all changes, remove all comments |
| `revisions` | `changes` | `<file>` | List tracked changes |
| `count` | `wc` | `<file>` | Word count |
| `meta` | `metadata` | `<file>` | Show or set metadata |
| `diff` | `compare` | `<file> <other-file>` | Compare two documents |
| `doctor` | `check` | `<file>` | Document health check |
| `init` | | `<file>` | Create .docexrc in document's directory |
| `style` | | `<file> --preset <name>` | Apply journal style preset |
| `verify` | `validate` | `<file> --preset <name>` | Validate against journal requirements |
| `anonymize` | `anon` | `<file>` | Remove author names for blind review |
| `expand` | | `<file> --vars <json>` | Expand {{VAR}} patterns |
| `compile` | | `<tex-file>` | Compile .tex to .docx |
| `decompile` | | `<file>` | Decompile .docx to .tex |
| `template` | `tpl` | `--title <t> --journal <j>` | Create from template |
| `response-letter` | `response` | `<file> --responses <json>` | Generate response letter |
| `watch` | | `<tex-file>` | Watch .tex and recompile on changes |
| `html` | | `<file>` | Export to HTML (via Pandoc) |
| `markdown` | `md` | `<file>` | Export to Markdown (via Pandoc) |
| `create` | | `<file>` | Create a minimal empty .docx |
| `latex` | `tex` | `<file>` | Export to LaTeX |

### Position Syntax

```
after:Methods       Insert after the "Methods" heading or paragraph
before:Conclusion   Insert before "Conclusion"
```

### Global Options

| Option | Description |
|--------|-------------|
| `--author <name>` | Author name (default: from git config or .docexrc) |
| `--by <name>` | Comment author (alias for --author) |
| `--untracked` | Disable tracked changes |
| `--output <path>` | Save to a different file |
| `--safe <path>` | Path to safe-modify.sh for manuscript protection |
| `--dry-run` | Preview changes without saving |
| `--width <inches>` | Figure width in inches (default: 6) |
| `--style <style>` | Table style: booktabs or plain |
| `--caption <text>` | Figure or table caption |
| `--preset <name>` | Journal preset: academic, polcomm, apa7, jcmc, joc |
| `--color <name>` | Color for highlight/color commands |
| `--json` | Output in JSON format |
| `--help` | Show help |
| `--version` | Show version |

### CLI Examples

```bash
# Replace text with tracked changes
docex replace manuscript.docx "enforcement gap" "regulatory gap" --author "Fabio"

# Insert paragraph after Methods
docex insert manuscript.docx "after:Methods" "We used a mixed-methods approach."

# Add a reviewer comment
docex comment manuscript.docx "platform regulation" "Needs citation" --by "Reviewer 2"

# Insert figure after Results
docex figure manuscript.docx "after:Results" fig03.png --caption "Figure 3. Status"

# Apply journal formatting
docex style manuscript.docx --preset polcomm

# Validate for submission
docex verify manuscript.docx --preset polcomm

# Document health check
docex doctor manuscript.docx

# Compare two versions
docex diff manuscript.docx manuscript-v1.docx --output diff.docx

# Compile LaTeX to docx
docex compile paper.tex --style polcomm --output submission.docx

# Watch and recompile
docex watch paper.tex --style polcomm

# Create from template
docex template --title "My Paper" --journal polcomm --output new.docx

# Word count
docex count manuscript.docx

# List headings
docex list manuscript.docx headings

# Export to LaTeX
docex latex manuscript.docx --output paper.tex
```

## Testing

```bash
# Run all tests (403 tests across 108 suites)
npm test

# Same thing without npm
node --test test/*.test.js
```

Test files:

| File | Description |
|------|-------------|
| `docex.test.js` | Core API: replace, insert, delete, comments, figures, tables |
| `integration.test.js` | End-to-end workflows with real .docx files |
| `formatting.test.js` | Inline formatting: bold, italic, highlight, color, etc. |
| `footnotes.test.js` | Footnote insertion and listing |
| `revisions.test.js` | Accept/reject tracked changes, clean copy |
| `citations.test.js` | Citation pattern detection, Zotero field code injection |
| `diff.test.js` | Document comparison |
| `metadata.test.js` | Dublin Core metadata read/write |
| `addressing.test.js` | Stable addressing: paraId, document map, find, structure |
| `latex.test.js` | LaTeX export and import |
| `compile.test.js` | LaTeX compile pipeline, templates, response letters |
| `academic.test.js` | Journal presets, validation, submission helpers |
| `robustness.test.js` | Edge cases, malformed input, large documents |
| `fuzz.test.js` | Fuzzy text matching, Unicode, special characters |
| `usability.test.js` | CLI argument parsing, error messages, preview |

All tests use Node.js built-in `node:test` and `node:assert`. No test dependencies.

## OOXML Reference

The `reference/` directory contains:

- `ooxml-reference.md` (2,967 lines) covering document body, paragraph structure, run properties, tracked changes, comments, images, tables, content types, and relationships.
- `ooxml-cheatsheet.md` for quick lookup.

These are useful if you need to extend docex or debug OOXML output.

## Project Structure

```
docex/
  src/
    docex.js             Main API: DocexEngine, PositionSelector, factory
    workspace.js         Zip/unzip lifecycle, save, verify, backup, locks
    paragraphs.js        Text replace, insert, delete, word count
    comments.js          Add, list, reply, resolve, remove, export (5-file)
    figures.js           Insert, replace, list images with relationships
    tables.js            Insert tables (booktabs and plain styles)
    formatting.js        Bold, italic, underline, highlight, color, code, etc.
    footnotes.js         List and add footnotes
    lists.js             Bullet and numbered lists with numbering.xml
    revisions.js         Accept, reject, list tracked changes, clean copy
    diff.js              Compare two documents, produce tracked changes
    citations.js         Detect patterns, inject Zotero field codes
    latex.js             OOXML to LaTeX converter
    compile.js           .tex to .docx pipeline, decompile, watch mode
    template.js          Create .docx from scratch with journal formatting
    response-letter.js   Generate R&R response letters
    presets.js           Journal style presets (polcomm, apa7, jcmc, joc)
    verify.js            Validate against journal submission requirements
    submission.js        Anonymize, deanonymize, highlighted changes
    batch.js             Batch operations across multiple files
    metadata.js          Dublin Core metadata read/write
    macros.js            Variable definition and {{VAR}} expansion
    crossref.js          Cross-references, labels, auto-numbering
    docmap.js            Document map, paraId injection, structure tree
    handle.js            ParagraphHandle for stable paraId-based operations
    doctor.js            Document health checks
    textmap.js           Map plain-text offsets to XML runs
    xml.js               Low-level XML utilities (regex-based)
  cli/
    docex-cli.js         CLI entry point (all commands)
  test/
    *.test.js            403 tests across 15 test files
    fixtures/            Test .docx files
  reference/
    ooxml-reference.md   OOXML format reference (2,967 lines)
    ooxml-cheatsheet.md  Quick lookup
  presets/               Style preset files (planned)
  skill/
    SKILL.md             AI agent skill file
```

## Design Principles

- **Zero external dependencies.** Only Node.js built-ins (`fs`, `path`, `child_process`, `crypto`, `zlib`, `os`). No npm install needed.
- **Tracked changes are the default.** Every edit is visible in Word's review pane unless you opt out with `.untracked()`.
- **Author set once, applies everywhere.** Call `doc.author("Name")` once and it propagates to all operations.
- **Position selectors read like English.** `doc.after("Methods").insert("text")` does what it says.
- **Single unzip/rezip cycle per save.** Queue operations, apply all in one pass. No corruption from repeated zip operations.
- **Automatic verification after every save.** Checks valid zip, paragraph count, and file size.
- **Stable addressing.** Every paragraph gets a `w14:paraId` that survives other edits.

## Contributing

Issues and pull requests welcome at [github.com/favstats/docex](https://github.com/favstats/docex).

## License

MIT. See [LICENSE](LICENSE).
