---
name: docex
description: Use when the user asks to modify, edit, or create .docx documents. Covers text replacement, tracked changes, comments, figures, tables, citations, and LaTeX export. Use when user says "update figure", "add a comment", "replace text", "rewrite paragraph", "add table", "inject citations", "export to LaTeX", or any .docx modification.
---

# docex: LaTeX for .docx

Zero-dependency Node.js library for programmatic .docx editing. Tracked changes by default. Single-pass architecture (unzip once, apply all operations in memory, rezip once). Built for academic manuscript workflows.

Location: `/mnt/storage/docex/`

## When to Use

- User asks to modify a .docx manuscript (replace text, rewrite a paragraph, fix a sentence)
- User asks to add, reply to, or resolve comments
- User asks to replace or insert figures
- User asks to add tracked changes for a revision
- User asks to inject Zotero citations
- User asks to export a .docx to LaTeX
- User asks to insert a table
- User asks to inspect a document (list headings, paragraphs, comments, figures)
- Any .docx editing task

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

All methods return `doc` for chaining.

### Position Selectors

Position selectors anchor operations to a location in the document. They match by text content (headings or body text).

```js
doc.at("some phrase")       // Anchors AT the text (for comments, replies)
doc.after("Methods")        // Position AFTER a heading or text
doc.before("Conclusion")    // Position BEFORE a heading or text
```

### Operations

#### Replace text

```js
doc.replace("old text", "new text");
doc.replace("old text", "new text", { author: "Override Author" });
doc.replace("old text", "new text", { tracked: false }); // direct edit
```

Preserves all formatting (bold, italic, font, size). Handles text spanning multiple XML runs. Tracked by default (shows as strikethrough + insertion).

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

#### Add comment

```js
// Via position selector
doc.at("enforcement gap").comment("Needs a citation here.", { by: "Reviewer 2" });

// Direct (anchors to the text in the first argument)
doc.comment("platform regulation", "This claim needs evidence.");
```

Comments appear in the sidebar in OnlyOffice/Word.

#### Reply to comment

```js
doc.at("original comment anchor text").reply("Thank you. We have corrected this.", { by: "Fabio Votta" });
```

Replies appear as threaded conversations under the parent comment. The anchor text must match text near the original comment. Reply comments are linked via `commentsExtended.xml`; they do NOT need ranges in `document.xml`.

#### Insert figure

```js
doc.after("Results").figure("figures/fig03.png", "Figure 3. Enforcement status.");
doc.after("Results").figure("figures/fig03.png", "Figure 3. Status.", { width: 5 }); // width in inches
```

The image is embedded in `word/media/` with a new relationship. Always insert inline at point of reference.

#### Insert table

```js
doc.after("Results").table(
  [["Party", "Ads", "Spend"], ["VVD", "342", "12,450"], ["D66", "215", "8,300"]],
  { style: "booktabs", caption: "Table 1. Ad spending by party." }
);
```

Supports `booktabs` (default) and `plain` styles. First row is treated as headers.

### Read-Only Inspection

These methods do not require `save()`. They return data about the current document.

```js
const paragraphs = await doc.paragraphs();  // [{index, text, style}]
const headings   = await doc.headings();    // [{level, text, index}]
const comments   = await doc.comments();    // [{id, author, text, date}]
const figures    = await doc.figures();      // [{rId, filename, width, height}]
const fullText   = await doc.text();        // string
const cites      = await doc.citations();   // [{text, paragraph, pattern, authors, year}]
```

### Export to LaTeX

```js
const tex = await doc.toLatex();
const tex = await doc.toLatex({ documentClass: "article", bibFile: "references", packages: ["booktabs"] });
```

Read-only export. Returns a complete LaTeX document string.

### Save

```js
// Overwrite original
await doc.save();

// Save to different file
await doc.save("output.docx");

// Save with safe-modify.sh protection (REQUIRED for manuscripts)
await doc.save({
  safeModify: "/mnt/storage/nl_local_2026/paper/build/safe-modify.sh",
  description: "Fix typo in Methods section"
});

// Save to different file with safe-modify.sh
await doc.save({
  outputPath: "manuscript_v2.docx",
  safeModify: "/path/to/safe-modify.sh",
  description: "R&R revisions"
});
```

Returns: `{ path, operations, paragraphCount, fileSize, verified }`

### Discard

```js
doc.discard();  // Discard all pending operations, clean up workspace
```

## CLI Reference

Invoke via Node.js or, if npm-linked, directly:

```bash
node /mnt/storage/docex/cli/docex-cli.js <command> [args...] [options]
```

### Commands

```bash
# Replace text (tracked by default)
docex replace manuscript.docx "old text" "new text" --author "Fabio Votta"

# Insert paragraph after a heading
docex insert manuscript.docx "after:Methods" "New methodology paragraph."

# Insert paragraph before a heading
docex insert manuscript.docx "before:Conclusion" "Final analysis paragraph."

# Delete text (tracked by default)
docex delete manuscript.docx "text to remove"

# Add comment
docex comment manuscript.docx "enforcement gap" "Needs citation" --by "Reviewer 2"

# Reply to comment (by ID or anchor text)
docex reply manuscript.docx 5 "Thank you. Corrected." --by "Fabio Votta"

# Insert figure
docex figure manuscript.docx "after:Results" figures/fig03.png --caption "Figure 3. Status."

# Insert table from JSON
docex table manuscript.docx "after:Results" data.json --style booktabs --caption "Table 1."

# List headings, paragraphs, comments, or figures
docex list manuscript.docx headings
docex list manuscript.docx comments
docex list manuscript.docx figures
docex list manuscript.docx paragraphs

# List citation patterns (no network)
docex cite manuscript.docx --list

# Inject Zotero citations
docex cite manuscript.docx --zotero-key KEY --zotero-user 6875557 --collection TUWJI72V

# Export to LaTeX (stdout)
docex latex manuscript.docx

# Export to LaTeX (file)
docex latex manuscript.docx --output paper.tex
```

### Global Options

| Flag | Description |
|------|-------------|
| `--author <name>` | Author name (default: from git config) |
| `--by <name>` | Comment/reply author (alias for --author) |
| `--untracked` | Disable tracked changes (direct edits) |
| `--output <path>` | Save to different file |
| `--safe <path>` | Path to safe-modify.sh (mandatory for manuscripts) |
| `--width <inches>` | Figure width (default: 6) |
| `--style <style>` | Table style: booktabs or plain |
| `--caption <text>` | Figure or table caption |

## Bulk Operations Pattern

For more than 5 edits, ALWAYS use the programmatic API in a single script. Never loop CLI commands.

```js
const docex = require('/mnt/storage/docex/src/docex');
const doc = docex("/mnt/storage/nl_local_2026/paper/manuscript.docx");
doc.author("Fabio Votta");

// All operations queued in memory
doc.replace("old sentence one", "new sentence one");
doc.replace("old sentence two", "new sentence two");
doc.after("Methods").insert("New methodology paragraph.");
doc.at("reviewer concern").reply("We have addressed this.", { by: "Fabio Votta" });
doc.delete("redundant phrase");
doc.after("Results").figure("figures/fig05.png", "Figure 5. Updated results.");

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

### 3. Insert a figure after a heading

```bash
# First, check current figures
node /mnt/storage/docex/cli/docex-cli.js list manuscript.docx figures

# Then insert
node /mnt/storage/docex/cli/docex-cli.js figure manuscript.docx \
  "after:Enforcement Patterns" \
  figures/fig03_enforcement_status.png \
  --caption "Figure 3. Enforcement status by platform." \
  --safe /mnt/storage/nl_local_2026/paper/build/safe-modify.sh
```

### 4. Full R&R workflow (replace + insert + comments + replies in one save)

```js
const docex = require('/mnt/storage/docex/src/docex');
const doc = docex("manuscript.docx");
doc.author("Fabio Votta");

// Text changes requested by reviewers
doc.replace("conceptualizing", "conceptualising");
doc.replace("Twitter", "X (formerly Twitter)");
doc.after("Literature Review").insert(
  "Recent work by Smith et al. (2025) demonstrates that platform enforcement varies significantly across election cycles."
);

// Reply to reviewer comments in the document
doc.at("British English").reply("We have revised the manuscript to use British English consistently.", { by: "Fabio Votta" });
doc.at("missing theory").reply("We have added a discussion of platform governance theory in Section 2.", { by: "Fabio Votta" });

// Insert updated figure
doc.after("Results").figure("figures/fig03_updated.png", "Figure 3. Updated enforcement status.");

// Single save with safe-modify protection
await doc.save({
  safeModify: "/mnt/storage/nl_local_2026/paper/build/safe-modify.sh",
  description: "R&R revisions: reviewer comments addressed"
});
```

### 5. Export to LaTeX

```bash
# Export to stdout
node /mnt/storage/docex/cli/docex-cli.js latex manuscript.docx

# Export to file with custom options
node /mnt/storage/docex/cli/docex-cli.js latex manuscript.docx \
  --output paper.tex \
  --doc-class article \
  --bib-file references \
  --packages "booktabs,graphicx,hyperref"
```

### 6. Inspect document before editing

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
```

## OOXML Reference

For low-level OOXML format details (XML structure, namespaces, tracked change markup, comment threading), see:

`/mnt/storage/docex/reference/ooxml-reference.md`

## Architecture

```
/mnt/storage/docex/
  src/
    docex.js        # Factory function, DocexEngine, PositionSelector
    workspace.js    # Zip/unzip, temp directory, save + verify
    paragraphs.js   # Replace, insert, delete (tracked and untracked)
    comments.js     # Add comments, threaded replies
    figures.js      # Insert/replace figures, embed in word/media/
    tables.js       # Insert tables (booktabs style)
    citations.js    # Find citation patterns, inject Zotero field codes
    latex.js        # Convert .docx to LaTeX
    textmap.js      # Map plain-text offsets to XML runs
    xml.js          # Low-level XML manipulation utilities
  cli/
    docex-cli.js    # CLI entry point (all commands)
  reference/
    ooxml-reference.md  # 2,967-line OOXML format reference
  test/
    docex.test.js       # Unit tests
    integration.test.js # Integration tests
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

## Common Mistakes

- Running build.sh after the .docx has been edited (overwrites human changes)
- Not running `list figures` before replacing a figure (wrong figure matched)
- Using too-short text in `replace` (matches the wrong paragraph; use 30+ characters of context)
- Running CLI commands in a loop instead of using the programmatic API for bulk edits
- Forgetting `--safe` for manuscript edits (no git safety net)
- Editing `manuscript_draft.md` thinking it updates the .docx (it will not; the .docx is the source of truth)
- Using "Mira" as author for R&R submissions (editors will see the name; use the real author)
