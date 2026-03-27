---
name: docex
description: Use when the user asks to modify, edit, or create .docx documents. Covers text replacement, tracked changes, comments, figures, tables, formatting, footnotes, lists, citations, journal styles, submission preparation, LaTeX export/import, document diff, templates, response letters, batch operations, and any .docx inspection or modification task.
---

# docex v0.4.0: LaTeX for .docx

Zero-dependency Node.js library for programmatic .docx editing. 28 modules, 403 tests, 14,101 lines. Tracked changes by default. Single-pass architecture (unzip once, apply all in memory, rezip once). Built for academic manuscript workflows but works on any .docx.

Location: `/mnt/storage/docex/`
Main entry: `/mnt/storage/docex/src/docex.js`
CLI entry: `/mnt/storage/docex/cli/docex-cli.js`

## When to Use

- User asks to modify a .docx manuscript (replace text, rewrite paragraphs, fix sentences)
- User asks to add, reply to, resolve, or export comments
- User asks to insert or replace figures or tables
- User asks to add tracked changes for a revision
- User asks to inject Zotero citations
- User asks to export a .docx to LaTeX, HTML, or Markdown
- User asks to import a .tex file as .docx
- User asks to format for a journal (apply styles, validate submission)
- User asks to anonymize/deanonymize for blind review
- User asks to generate a response letter for R&R
- User asks to create a new .docx from scratch or from a template
- User asks to compare two .docx versions (diff)
- User asks to accept/reject tracked changes or produce a clean copy
- User asks to inspect a document (list headings, paragraphs, comments, figures, word count, stats)
- User asks to apply inline formatting (bold, italic, highlight, footnotes, lists)
- User asks to batch-process multiple .docx files
- Any .docx editing, creation, or inspection task

## Hard Rules

1. **NEVER regenerate a .docx** once humans have edited it. Use docex for surgical edits only.
2. **Every manuscript.docx edit MUST go through safe-modify.sh** (use the `--safe` CLI flag or the `safeModify` save option). This auto-commits to git, creates a timestamped backup, verifies integrity, and auto-restores on failure.
3. **Figures ALWAYS inline** at the point of first reference in the text. NEVER at the end (exception: appendix figures).
4. **NEVER loop CLI commands** for bulk operations. Use the programmatic API with a single `doc.save()` call. Repeated unzip/rezip cycles corrupt documents (documented incident: 25 pages reduced to 7 after 20 cycles).
5. **Author attribution:** For R&R submissions, use the manuscript author's real name (e.g., "Fabio Votta") so journal editors see the author's revisions. Use "Mira" only for internal review. Always ask the user if unclear.
6. **Verify after every modification:** The `save()` method automatically checks that paragraph count >= original, file size is within 20% of original, and the zip is valid.

## Programmatic API Reference

### Opening a Document

```js
const docex = require('/mnt/storage/docex/src/docex');
const doc = docex("/path/to/manuscript.docx");
```

### Configuration

```js
doc.author("Fabio Votta");              // Set author for all operations (required)
doc.date("2026-03-27T00:00:00Z");       // Set date (default: now)
doc.untracked();                         // Disable tracked changes (direct edits)
doc.tracked();                           // Re-enable tracked changes (default: on)
```

All configuration methods return `doc` for chaining.

### Position Selectors

Position selectors anchor operations to a location in the document. They match by text content (headings or body text).

```js
doc.at("some phrase")       // Anchors AT the text (for comments, formatting, footnotes, replies)
doc.after("Methods")        // Position AFTER a heading or text
doc.before("Conclusion")    // Position BEFORE a heading or text
doc.afterHeading("Methods") // Position AFTER a heading specifically
doc.afterText("some text")  // Position AFTER body text specifically
doc.id("3A7F2B1C")         // ParagraphHandle by w14:paraId (stable addressing)
```

### Text Operations

#### Replace text

```js
doc.replace("old text", "new text");
doc.replace("old text", "new text", { author: "Override Author" });
doc.replace("old text", "new text", { tracked: false }); // direct edit
```

Preserves all formatting (bold, italic, font, size). Handles text spanning multiple XML runs. Tracked by default (shows as strikethrough + insertion).

#### Replace all occurrences

```js
doc.replaceAll("old term", "new term");
```

#### Delete text

```js
doc.delete("text to remove");
```

Tracked by default (shows as strikethrough).

#### Insert paragraph

```js
doc.after("Methods").insert("New paragraph text.");
doc.before("Conclusion").insert("Final paragraph.");
```

New paragraphs inherit the document's default formatting.

### Comments

#### Add comment

```js
// Via position selector
doc.at("enforcement gap").comment("Needs a citation here.", { by: "Reviewer 2" });

// Direct (anchors to the text in the first argument)
doc.comment("platform regulation", "This claim needs evidence.");
```

Comments appear in the sidebar in OnlyOffice/Word. Manages 5 OOXML files: document.xml (ranges), comments.xml, commentsExtended.xml, commentsIds.xml, [Content_Types].xml.

#### Reply to comment

```js
doc.at("original comment anchor text").reply("Thank you. We have corrected this.", { by: "Fabio Votta" });
```

Replies appear as threaded conversations under the parent comment. The anchor text must match text near the original comment. Reply comments are linked via `commentsExtended.xml`; they do NOT need ranges in `document.xml`.

#### Read comments

```js
const comments = await doc.comments();
// [{id, author, text, date, anchor}]
```

#### Export comments

```js
const csv = await doc.exportComments("csv");
const json = await doc.exportComments("json");
```

### Figures

#### Insert figure

```js
doc.after("Results").figure("figures/fig03.png", "Figure 3. Enforcement status.");
doc.after("Results").figure("figures/fig03.png", "Figure 3. Status.", { width: 5 }); // width in inches
```

The image is embedded in `word/media/` with a new relationship. Always insert inline at point of reference. PNG and JPEG supported; dimensions auto-detected.

#### List figures

```js
const figures = await doc.figures();
// [{rId, filename, width, height}]
```

### Tables

```js
doc.after("Results").table(
  [["Party", "Ads", "Spend"], ["VVD", "342", "12,450"], ["D66", "215", "8,300"]],
  { style: "booktabs", caption: "Table 1. Ad spending by party." }
);
```

Options:
- `style`: `"booktabs"` (default, academic) or `"plain"` (grid)
- `caption`: Table caption text
- `headers`: Treat first row as headers (default: true)

### Inline Formatting

Via position selector:

```js
doc.at("text").bold();
doc.at("text").italic();
doc.at("text").underline();
doc.at("text").strikethrough();
doc.at("text").superscript();
doc.at("text").subscript();
doc.at("text").smallCaps();
doc.at("text").code();
doc.at("text").color("red");
doc.at("text").highlight("yellow");
```

Direct (without position selector):

```js
doc.bold("text to format");
doc.italic("text to format");
doc.highlight("text to format", "yellow");
doc.color("text to format", "red");
```

### Footnotes

```js
doc.at("anchor text").footnote("Footnote content here.");
const footnotes = await doc.footnotes();
// [{id, text}]
```

### Lists

```js
doc.after("Methods").bulletList(["Item one", "Item two", "Item three"]);
doc.after("Methods").numberedList(["First step", "Second step", "Third step"]);
```

### Revisions (Tracked Changes)

```js
const revisions = await doc.revisions();     // [{id, type, author, date, text}]
await doc.accept();                           // accept all
await doc.accept(5);                          // accept specific change by ID
await doc.reject();                           // reject all
await doc.reject(3);                          // reject specific change by ID
await doc.cleanCopy();                        // accept all changes, remove all comments
```

### Contributors and Timeline

```js
const contribs = await doc.contributors();
// [{name, changes, comments, lastActive}]

const events = await doc.timeline();
// [{date, type, author, text}]  -- combined chronological view
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
// List citation patterns (no network, just pattern matching)
const cites = await doc.citations();
// [{text, paragraph, pattern, authors, year}]
// pattern is "parenthetical" or "narrative"

// Inject Zotero citation field codes
const result = await doc.injectCitations({
  zoteroApiKey: "...",
  zoteroUserId: "6875557",
  collectionId: "TUWJI72V",  // optional: limit to collection
  bibliography: true,          // insert bibliography field (default: true)
});
// {found, matched, injected, unmatched}
```

### Metadata

```js
// Get metadata
const meta = await doc.metadata();
// {title, creator, subject, keywords, description, ...}

// Set metadata
await doc.metadata({ title: "My Paper", creator: "Fabio Votta", keywords: "elections, ads" });
```

### Word Count and Stats

```js
const wc = await doc.wordCount();
// {total, body, headings, abstract, captions, footnotes}

const stats = await doc.stats();
// {words, paragraphs, headings, figures, tables, citations, comments, revisions, pages}
```

### Journal Styles

Built-in presets: `academic`, `polcomm`, `apa7`, `jcmc`, `joc`

```js
// Apply a style preset
await doc.style("polcomm");

// Define a custom preset
const docex = require('/mnt/storage/docex/src/docex');
docex.defineStyle("myjournal", {
  font: "Garamond", size: 11, spacing: "double",
  margins: { top: 1, bottom: 1, left: 1.25, right: 1.25 },
  indent: 0.5, alignment: "justified",
  titlePage: true, runningHeader: true,
  abstractWordLimit: 200, wordLimit: 8000,
});

// List available presets
const presets = docex.listStyles();
```

### Submission Validation

```js
const result = await doc.verify("polcomm");
// {pass, errors, warnings}
// Checks: word count, abstract length, heading hierarchy, margins, font, spacing,
// running header, title page, line numbering, figure resolution, etc.
```

### Submission Helpers

```js
// Anonymize for blind review
const anon = await doc.anonymize();
// {authorsRemoved, locations}

// Deanonymize (restore author info after review)
const deanon = await doc.deanonymize();
// {restored, authors}

// Generate highlighted changes document (insertions yellow, deletions red)
const hc = await doc.highlightedChanges();
// {insertions, deletions}
```

### Response Letters

Generate a formatted .docx response letter from reviewer comments.

```js
const docex = require('/mnt/storage/docex/src/docex');
const doc = docex("manuscript.docx");
const comments = await doc.comments();

const responses = {
  1: { action: "agree", text: "We have added the citation.", changes: ["Added Suzor 2019 to p. 12"] },
  2: { action: "partial", text: "We clarified but kept the framing.", changes: ["Revised Discussion para 3"] },
  3: { action: "disagree", text: "We maintain our position because...", changes: [] },
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
const docex = require('/mnt/storage/docex/src/docex');

// Create from journal template
const result = await docex.fromTemplate({
  title: "My Paper Title",
  authors: [
    { name: "Fabio Votta", affiliation: "University of Amsterdam", email: "f.votta@uva.nl" },
    { name: "Simon Munzert", affiliation: "JGU Mainz" },
  ],
  abstract: "This paper examines...",
  keywords: ["elections", "enforcement", "platforms"],
  sections: ["Introduction", "Literature Review", "Methods", "Results", "Discussion", "Conclusion"],
  journal: "polcomm",   // optional: apply journal preset
  output: "new-manuscript.docx",
});
// {path, fileSize, paragraphCount, journal}

// Create a minimal empty .docx
const path = await docex.create("blank.docx");
```

### Batch Operations

```js
const docex = require('/mnt/storage/docex/src/docex');

const batch = docex.batch(["paper1.docx", "paper2.docx", "paper3.docx"]);
batch.author("Fabio Votta");
batch.style("polcomm");
batch.replaceAll("old term", "new term");

const results = await batch.saveAll({ suffix: "_formatted" });
// [{path, fileSize, paragraphCount, error}]

// Or save to a directory
const results = await batch.saveAll({ outputDir: "/output/" });

// Batch verify
const verifyResults = await batch.verify("polcomm");
// [{path, result: {pass, errors, warnings}}]
```

### LaTeX Pipeline

Requires Pandoc to be installed on the system.

```js
const docex = require('/mnt/storage/docex/src/docex');

// Compile .tex -> submission-ready .docx
const result = await docex.compile("paper.tex", {
  style: "polcomm",          // apply journal preset
  bibFile: "references.bib", // bibliography
  cslFile: "apa.csl",        // citation style
  output: "submission.docx",
  pandocArgs: ["--number-sections"],  // extra pandoc args
});
// {path, fileSize, paragraphCount, style}

// Decompile .docx -> .tex (preserves tracked changes and comments)
const result = await docex.decompile("manuscript.docx", {
  output: "paper.tex",
  documentClass: "article",
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
const tex = await doc.toLatex({ documentClass: "article", bibFile: "references", packages: ["booktabs"] });

// To HTML (requires Pandoc)
const html = await doc.toHtml({ output: "paper.html" });

// To Markdown (requires Pandoc)
const md = await doc.toMarkdown({ output: "paper.md" });
```

### Document Health

```js
const result = await doc.validate();
// {valid, errors, warnings}
// Checks: valid zip, document.xml exists, relationships resolve, no orphaned media,
// comment consistency, paraId uniqueness, heading hierarchy
```

### Stable Addressing (paraId)

Every paragraph gets a unique `w14:paraId`. Operations on a paraId survive other edits.

```js
// Generate document map
const map = await doc.map();
// {sections, allParagraphs, allFigures, allTables, allComments}

// ParagraphHandle by paraId
const handle = doc.id("3A7F2B1C");
handle.replace("old", "new");
handle.comment("Note text");
handle.bold();

// Find paragraphs by text
const matches = await doc.find("enforcement gap");
// [{id, index, section, context}]

// Document structure tree (human-readable)
const tree = await doc.structure();

// Explain XML around text (for debugging)
const explanation = await doc.explain("enforcement gap");
```

### Cross-References and Auto-Numbering

```js
// Label a paragraph for referencing
await doc.label("3A7F2B1C", "fig:funnel");

// Insert a cross-reference
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

// List all undefined {{VAR}} patterns
const vars = await doc.listVariables();
// [{name, paragraph, context}]
```

### Preview and Discard

```js
// Preview pending operations without executing
console.log(doc.preview());
// "3 pending operations:
//   1. replace 'old text' -> 'new text' (tracked, by Fabio Votta)
//   2. insert after 'Methods': 'New para.' (tracked, by Fabio Votta)
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

// Save with safe-modify.sh protection (REQUIRED for manuscripts)
const result = await doc.save({
  safeModify: "/mnt/storage/nl_local_2026/paper/build/safe-modify.sh",
  description: "Fix typo in Methods section"
});

// Save to different file with safe-modify.sh
const result = await doc.save({
  outputPath: "manuscript_v2.docx",
  safeModify: "/path/to/safe-modify.sh",
  description: "R&R revisions"
});

// Dry run (preview result without writing)
const result = await doc.save({ dryRun: true });
```

Returns: `{ path, operations, paragraphCount, fileSize, verified }`

## CLI Reference

Invoke via Node.js or, if npm-linked, directly:

```bash
node /mnt/storage/docex/cli/docex-cli.js <command> [args...] [options]
```

### All Commands

```bash
# Text operations
docex replace manuscript.docx "old text" "new text" --author "Fabio Votta"
docex insert manuscript.docx "after:Methods" "New methodology paragraph."
docex insert manuscript.docx "before:Conclusion" "Final analysis paragraph."
docex delete manuscript.docx "text to remove"

# Comments
docex comment manuscript.docx "enforcement gap" "Needs citation" --by "Reviewer 2"
docex reply manuscript.docx 5 "Thank you. Corrected." --by "Fabio Votta"

# Figures and tables
docex figure manuscript.docx "after:Results" figures/fig03.png --caption "Figure 3. Status."
docex table manuscript.docx "after:Results" data.json --style booktabs --caption "Table 1."

# Formatting
docex bold manuscript.docx "important phrase"
docex italic manuscript.docx "emphasis here"
docex highlight manuscript.docx "key finding" --color yellow
docex footnote manuscript.docx "anchor text" "Footnote content."

# Revisions
docex accept manuscript.docx          # accept all
docex accept manuscript.docx 5        # accept change ID 5
docex reject manuscript.docx          # reject all
docex reject manuscript.docx 3        # reject change ID 3
docex clean manuscript.docx           # accept all + remove comments
docex revisions manuscript.docx       # list tracked changes

# Citations
docex cite manuscript.docx --list                          # list patterns
docex cite manuscript.docx --zotero-key KEY --zotero-user 6875557 --collection TUWJI72V

# Inspection
docex list manuscript.docx headings
docex list manuscript.docx comments
docex list manuscript.docx figures
docex list manuscript.docx paragraphs
docex list manuscript.docx revisions
docex list manuscript.docx footnotes
docex count manuscript.docx               # word count
docex meta manuscript.docx                # metadata

# Journal formatting and validation
docex style manuscript.docx --preset polcomm
docex verify manuscript.docx --preset polcomm
docex anonymize manuscript.docx

# Document comparison and health
docex diff manuscript.docx manuscript-v1.docx --output diff.docx
docex doctor manuscript.docx

# LaTeX pipeline
docex compile paper.tex --style polcomm --output submission.docx
docex decompile manuscript.docx --output paper.tex
docex watch paper.tex --style polcomm
docex latex manuscript.docx --output paper.tex

# Export
docex html manuscript.docx --output paper.html
docex markdown manuscript.docx --output paper.md

# Templates and creation
docex template --title "My Paper" --journal polcomm --output new.docx
docex create blank.docx

# Response letter
docex response-letter manuscript.docx --responses responses.json --output response.docx

# Variable expansion
docex expand manuscript.docx --vars '{"NUM_ADS":"268,635"}'

# Init .docexrc
docex init manuscript.docx
```

### Global Options

| Flag | Description |
|------|-------------|
| `--author <name>` | Author name (default: from git config or .docexrc) |
| `--by <name>` | Comment/reply author (alias for --author) |
| `--untracked` | Disable tracked changes (direct edits) |
| `--output <path>` | Save to different file |
| `--safe <path>` | Path to safe-modify.sh (mandatory for manuscripts) |
| `--dry-run` | Preview changes without saving |
| `--width <inches>` | Figure width (default: 6) |
| `--style <style>` | Table style: booktabs or plain |
| `--caption <text>` | Figure or table caption |
| `--preset <name>` | Journal preset: academic, polcomm, apa7, jcmc, joc |
| `--color <name>` | Color for highlight/color commands |
| `--json` | Output in JSON format |
| `--help` | Show help |
| `--version` | Show version |

## Bulk Operations Pattern

For more than 5 edits, ALWAYS use the programmatic API in a single script. Never loop CLI commands.

```js
const docex = require('/mnt/storage/docex/src/docex');
const doc = docex("/mnt/storage/nl_local_2026/paper/manuscript.docx");
doc.author("Fabio Votta");

// All operations queued in memory
doc.replace("old sentence one", "new sentence one");
doc.replace("old sentence two", "new sentence two");
doc.replaceAll("Twitter", "X (formerly Twitter)");
doc.after("Methods").insert("New methodology paragraph.");
doc.at("reviewer concern").reply("We have addressed this.", { by: "Fabio Votta" });
doc.delete("redundant phrase");
doc.after("Results").figure("figures/fig05.png", "Figure 5. Updated results.");
doc.at("key finding").bold();
doc.at("important phrase").highlight("yellow");

// Single save: one unzip, all changes applied, one rezip, auto-verify
await doc.save({
  safeModify: "/mnt/storage/nl_local_2026/paper/build/safe-modify.sh",
  description: "R&R batch revisions"
});
```

## Common Workflows

### 1. Replace text with tracked changes

```bash
node /mnt/storage/docex/cli/docex-cli.js replace manuscript.docx \
  "Meta removed only 189 of these ads" \
  "Meta removed only 192 of these ads" \
  --author "Fabio Votta" \
  --safe /mnt/storage/nl_local_2026/paper/build/safe-modify.sh
```

### 2. Add reviewer comments (multiple personas)

```js
const docex = require('/mnt/storage/docex/src/docex');
const doc = docex("manuscript.docx");

doc.at("enforcement gap").comment("This needs a citation to the 2024 DSA report.", { by: "Reviewer 1" });
doc.at("platform regulation").comment("Can you clarify what you mean by 'regulation'?", { by: "Reviewer 2" });
doc.at("14.2% takedown rate").comment("Is this figure from the full sample or the subset?", { by: "Reviewer 2" });

await doc.save();
```

### 3. Full R&R workflow (replace + insert + comments + replies + figures in one save)

```js
const docex = require('/mnt/storage/docex/src/docex');
const doc = docex("manuscript.docx");
doc.author("Fabio Votta");

// Text changes requested by reviewers
doc.replace("conceptualizing", "conceptualising");
doc.replaceAll("Twitter", "X (formerly Twitter)");
doc.after("Literature Review").insert(
  "Recent work by Smith et al. (2025) demonstrates that platform enforcement varies significantly across election cycles."
);

// Reply to reviewer comments in the document
doc.at("British English").reply("We have revised the manuscript to use British English consistently.", { by: "Fabio Votta" });
doc.at("missing theory").reply("We have added a discussion of platform governance theory in Section 2.", { by: "Fabio Votta" });

// Insert updated figure
doc.after("Results").figure("figures/fig03_updated.png", "Figure 3. Updated enforcement status.");

// Apply formatting
doc.at("key contribution").bold();
doc.at("p < .001").italic();

// Single save with safe-modify protection
await doc.save({
  safeModify: "/mnt/storage/nl_local_2026/paper/build/safe-modify.sh",
  description: "R&R revisions: reviewer comments addressed"
});
```

### 4. Compile LaTeX to journal-formatted .docx

```bash
node /mnt/storage/docex/cli/docex-cli.js compile paper.tex \
  --style polcomm \
  --bib references.bib \
  --output submission.docx
```

### 5. Generate response letter from manuscript comments

```js
const docex = require('/mnt/storage/docex/src/docex');
const doc = docex("manuscript.docx");
const comments = await doc.comments();
doc.discard();

const responses = {};
for (const c of comments) {
  responses[c.id] = {
    action: "agree",
    text: "We have addressed this comment.",
    changes: ["See revised manuscript."]
  };
}

await docex.responseLetter(comments, responses, {
  title: "Platform Enforcement",
  journal: "Political Communication",
  authors: ["Fabio Votta"],
  output: "response-letter.docx",
});
```

### 6. Prepare submission package

```js
const docex = require('/mnt/storage/docex/src/docex');

// Validate
const doc = docex("manuscript.docx");
const result = await doc.verify("polcomm");
console.log(result.pass ? "PASS" : "FAIL");
console.log(result.errors);
console.log(result.warnings);

// Create anonymized copy
const doc2 = docex("manuscript.docx");
await doc2.anonymize();
await doc2.save("manuscript-blind.docx");

// Create clean copy (no tracked changes, no comments)
const doc3 = docex("manuscript.docx");
await doc3.cleanCopy();
await doc3.save("manuscript-clean.docx");
```

### 7. Inspect document before editing

Always inspect first to understand the document structure:

```bash
# Show headings (document structure)
node /mnt/storage/docex/cli/docex-cli.js list manuscript.docx headings

# Show all comments (for R&R planning)
node /mnt/storage/docex/cli/docex-cli.js list manuscript.docx comments

# Show all figures (before replacing one)
node /mnt/storage/docex/cli/docex-cli.js list manuscript.docx figures

# Show citation patterns
node /mnt/storage/docex/cli/docex-cli.js cite manuscript.docx --list

# Word count
node /mnt/storage/docex/cli/docex-cli.js count manuscript.docx

# Document health
node /mnt/storage/docex/cli/docex-cli.js doctor manuscript.docx
```

## Architecture

```
/mnt/storage/docex/
  src/
    docex.js             Factory, DocexEngine, PositionSelector (1,421 lines)
    workspace.js         Zip/unzip, temp dir, save, verify, backup, locks (793)
    paragraphs.js        Replace, insert, delete, word count (1,292)
    comments.js          Add, list, reply, resolve, remove, export (850)
    figures.js           Insert, replace, list images (640)
    tables.js            Booktabs and plain tables (329)
    formatting.js        Bold, italic, underline, highlight, color, code (434)
    footnotes.js         List and add footnotes (258)
    lists.js             Bullet and numbered lists (389)
    revisions.js         Accept, reject, list changes, clean copy (451)
    diff.js              Compare two documents (439)
    citations.js         Detect patterns, inject Zotero codes (703)
    latex.js             OOXML to LaTeX converter (993)
    compile.js           .tex to .docx, decompile, watch (347)
    template.js          Create .docx from scratch (363)
    response-letter.js   R&R response letters (305)
    presets.js           Journal style presets (437)
    verify.js            Submission validation (259)
    submission.js        Anonymize, deanonymize, highlighted changes (245)
    batch.js             Multi-file operations (201)
    metadata.js          Dublin Core metadata (223)
    macros.js            {{VAR}} expansion (217)
    crossref.js          Cross-refs, labels, auto-numbering (298)
    docmap.js            Document map, paraId, structure tree (453)
    handle.js            ParagraphHandle for stable addressing (643)
    doctor.js            Document health checks (358)
    textmap.js           Text-to-XML offset mapping (286)
    xml.js               Low-level XML utilities (474)
  cli/
    docex-cli.js         CLI (all commands)
  reference/
    ooxml-reference.md   OOXML spec reference (2,967 lines)
    ooxml-cheatsheet.md  Quick lookup
```

## Safety and Verification

### safe-modify.sh

Each paper project has its own `safe-modify.sh` in `build/`. For example:
- `/mnt/storage/nl_local_2026/paper/build/safe-modify.sh`

It provides: git auto-commit before and after, timestamped backup, integrity verification, auto-restore on failure.

### Auto-verify on save

Every `doc.save()` call automatically verifies:
- Valid zip structure
- Paragraph count >= original
- File size within 20% of original

The result object includes `{ verified: true/false }`.

### Undo

```bash
# Restore from backup
cp manuscript.docx.bak manuscript.docx

# Or use git (safe-modify.sh auto-commits before every change)
git log --oneline manuscript.docx
git checkout <hash> -- manuscript.docx
```

## OOXML Reference

For low-level OOXML format details (XML structure, namespaces, tracked change markup, comment threading), see:

`/mnt/storage/docex/reference/ooxml-reference.md` (2,967 lines)
`/mnt/storage/docex/reference/ooxml-cheatsheet.md` (quick lookup)

## Common Mistakes

- Running build.sh after the .docx has been edited (overwrites human changes)
- Not running `list figures` before replacing a figure (wrong figure matched)
- Using too-short text in `replace` (matches the wrong paragraph; use 30+ characters of context)
- Running CLI commands in a loop instead of using the programmatic API for bulk edits
- Forgetting `--safe` for manuscript edits (no git safety net)
- Using "Mira" as author for R&R submissions (editors will see the name; use the real author)
- Calling `doc.save()` multiple times in the same script (corruption risk; queue all operations, save once)
- Not calling `doc.discard()` after read-only operations (leaves temp files)
