# docex

**LaTeX for .docx** -- a zero-dependency Node.js library for programmatic document editing with tracked changes, comments, figures, and tables. Built for academic manuscript workflows.

docex treats `.docx` files the way LaTeX treats `.tex` files: you describe edits in code, and the tool produces a correctly formatted document. Tracked changes are on by default, so every edit is visible in Word's review pane. The entire edit cycle runs in a single unzip/rezip pass, which avoids the corruption that comes from repeated zip operations.

## Install

```bash
git clone https://github.com/favstats/docex.git
cd docex
# No npm install needed -- zero dependencies
```

Requires Node.js 18 or later (uses `node:test` and `node:zlib`).

## Quick Start (API)

```js
const docex = require('./src/docex');
const doc = docex("manuscript.docx");
doc.author("Fabio Votta");

doc.replace("old text", "new text");                     // tracked by default
doc.after("Methods").insert("New paragraph.");            // position selector
doc.at("enforcement gap").comment("Cite Suzor 2019", { by: "Prof. Strict" });
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

## API Reference

### Opening a Document

```js
const docex = require('./src/docex');

// Factory function (preferred)
const doc = docex("manuscript.docx");

// Alternative form
const doc = docex.open("manuscript.docx");
```

Both return a `DocexEngine` instance.

### Configuration

All configuration methods are chainable.

#### `doc.author(name)`

Set the author name for all subsequent operations. This name appears in tracked changes and comments in Word's review pane.

```js
doc.author("Fabio Votta");
```

#### `doc.date(isoDate)`

Set the date for all subsequent operations. Defaults to the current time.

```js
doc.date("2026-03-15T12:00:00Z");
```

#### `doc.tracked()`

Re-enable tracked changes (the default). Every replace, insert, and delete will appear as a revision in Word.

#### `doc.untracked()`

Disable tracked changes. Edits modify the document directly without leaving revision marks.

```js
doc.untracked();
doc.replace("typo", "correction");  // silent fix, no revision mark
doc.tracked();                      // back to tracked mode
```

### Position Selectors

Position selectors let you target a location in the document by its text content. They read like English.

#### `doc.at(text)`

Select the position of `text` in the document. Used for anchoring comments.

```js
doc.at("platform regulation").comment("Needs citation");
```

#### `doc.after(text)`

Select the position immediately after `text` or after a heading matching `text`. Used for inserting paragraphs, figures, and tables.

```js
doc.after("Methods").insert("We used a mixed-methods approach.");
doc.after("Results").figure("fig01.png", "Figure 1. Overview");
doc.after("Results").table(data, { style: "booktabs" });
```

#### `doc.before(text)`

Select the position immediately before `text` or before a heading matching `text`.

```js
doc.before("Conclusion").insert("One additional finding deserves mention.");
```

### Position Selector Methods

Each position selector returns a `PositionSelector` object with the following methods.

#### `.insert(text, opts?)`

Insert a new paragraph at this position.

```js
doc.after("Methods").insert("New methodology paragraph.");
```

Options:
- `author` (string) -- override author for this operation
- `tracked` (boolean) -- override tracked setting

#### `.comment(text, opts?)`

Add a comment anchored to the selected text.

```js
doc.at("enforcement gap").comment("Cite Suzor 2019", { by: "Reviewer 2" });
```

Options:
- `by` (string) -- comment author (overrides doc author)
- `initials` (string) -- author initials

#### `.figure(imagePath, caption, opts?)`

Insert a figure at this position.

```js
doc.after("Results").figure("fig03.png", "Figure 3. Enforcement status by party");
```

Options:
- `width` (number) -- figure width in inches (default: 6)
- `author` (string) -- override author
- `tracked` (boolean) -- override tracked setting

#### `.table(data, opts?)`

Insert a table at this position. `data` is a 2D array where the first row is treated as headers by default.

```js
doc.after("Results").table(
  [
    ["Party", "Ads", "Violations"],
    ["PAX",   "117", "3"],
    ["VVD",   "204", "12"]
  ],
  { style: "booktabs", caption: "Table 1. Ad counts by party" }
);
```

Options:
- `style` (string) -- `"booktabs"` (default) or `"plain"`
- `caption` (string) -- table caption
- `headers` (boolean) -- treat first row as headers (default: `true`)
- `author` (string) -- override author
- `tracked` (boolean) -- override tracked setting

#### `.reply(text, opts?)`

Reply to a comment at this position. Finds the comment by its anchor text.

```js
doc.at("enforcement gap").reply("Added Suzor 2019 citation.", { by: "Fabio Votta" });
```

### Direct Operations

These methods operate on the whole document without a position selector.

#### `doc.replace(oldText, newText, opts?)`

Replace text anywhere in the document. Tracked by default: the old text appears as strikethrough and the new text as an insertion.

```js
doc.replace("old conclusion", "new conclusion");
doc.replace("typo", "correction", { tracked: false });  // silent fix
```

Options:
- `author` (string) -- override author
- `tracked` (boolean) -- override tracked setting

#### `doc.delete(text, opts?)`

Delete text from the document. Tracked by default: the deleted text appears as strikethrough.

```js
doc.delete("This paragraph is no longer needed.");
```

Options:
- `author` (string) -- override author
- `tracked` (boolean) -- override tracked setting

#### `doc.comment(anchor, text, opts?)`

Add a comment anchored to text in the document. Shorthand for `doc.at(anchor).comment(text, opts)`.

```js
doc.comment("platform regulation", "Needs citation", { by: "Reviewer 2" });
```

### Inspection (Read-Only)

These methods read document content without modifying it. They return promises.

#### `await doc.paragraphs()`

List all paragraphs with their text, index, and style.

```js
const paras = await doc.paragraphs();
// [{ index: 0, text: "Introduction", style: "Heading1" }, ...]
```

#### `await doc.headings()`

List all headings with their level, text, and paragraph index.

```js
const headings = await doc.headings();
// [{ level: 1, text: "Introduction", index: 0 }, ...]
```

#### `await doc.comments()`

List all comments with their ID, author, text, and date.

```js
const comments = await doc.comments();
// [{ id: 1, author: "Reviewer 2", text: "Needs citation", date: "..." }, ...]
```

#### `await doc.figures()`

List all images and figures in the document.

```js
const figs = await doc.figures();
// [{ rId: "rId7", filename: "image1.png", width: 600, height: 400 }, ...]
```

#### `await doc.text()`

Get the full text content of the document as a single string.

```js
const fullText = await doc.text();
```

### Lifecycle

#### `await doc.save(outputPath?)`

Execute all queued operations and save the document. This is the only point where the file is modified: docex unzips once, applies all operations in memory, rezips once, then verifies the output.

```js
// Overwrite the original
const result = await doc.save();

// Save to a new file
const result = await doc.save("revised-manuscript.docx");
```

Returns an object:
```js
{
  path: "/absolute/path/to/output.docx",
  operations: 5,
  paragraphCount: 142,
  fileSize: 28450,
  verified: true
}
```

#### `doc.discard()`

Discard all pending operations and clean up the workspace without saving.

```js
doc.discard();
```

## CLI Reference

```
docex <command> <file> [arguments] [options]
```

### Commands

| Command | Arguments | Description |
|---------|-----------|-------------|
| `replace` | `<file> <old> <new>` | Replace text (tracked by default) |
| `insert` | `<file> <position> <text>` | Insert paragraph at position |
| `delete` | `<file> <text>` | Delete text (tracked by default) |
| `comment` | `<file> <anchor> <text>` | Add comment anchored to text |
| `reply` | `<file> <comment-id> <text>` | Reply to existing comment |
| `figure` | `<file> <position> <image>` | Insert figure at position |
| `table` | `<file> <position> <json-file>` | Insert table from JSON |
| `list` | `<file> [type]` | List paragraphs, headings, comments, or figures |

### Position Syntax

Positions for `insert`, `figure`, and `table` use a prefix syntax:

```
after:Methods       Insert after the "Methods" heading or paragraph
before:Conclusion   Insert before "Conclusion"
```

### Global Options

| Option | Description |
|--------|-------------|
| `--author <name>` | Author name (default: from git config) |
| `--by <name>` | Comment author (alias for `--author`) |
| `--untracked` | Disable tracked changes |
| `--output <path>` | Save to a different file instead of overwriting |
| `--width <inches>` | Figure width in inches (default: 6) |
| `--style <style>` | Table style: `booktabs` or `plain` (default: `booktabs`) |
| `--caption <text>` | Figure or table caption |
| `--help` | Show help |
| `--version` | Show version |

### CLI Examples

```bash
# Replace text with tracked changes
docex replace manuscript.docx "enforcement gap" "regulatory gap" --author "Fabio"

# Insert a paragraph after the Methods section
docex insert manuscript.docx "after:Methods" "We used a mixed-methods approach."

# Add a reviewer comment
docex comment manuscript.docx "platform regulation" "Needs citation" --by "Reviewer 2"

# Insert a figure after Results
docex figure manuscript.docx "after:Results" fig03.png --caption "Figure 3. Status"

# List all headings
docex list manuscript.docx headings

# List all comments
docex list manuscript.docx comments

# Save to a new file instead of overwriting
docex replace manuscript.docx "old" "new" --output revised.docx
```

## Testing

```bash
# Run all tests
npm test

# Run unit tests only (49 tests)
node --test test/docex.test.js

# Run integration tests only (11 tests)
node --test test/integration.test.js
```

All tests use Node.js built-in `node:test` and `node:assert`. No test dependencies.

## OOXML Reference

The `reference/` directory contains a comprehensive OOXML reference document (`ooxml-reference.md`, 2,967 lines) covering the XML structures that docex manipulates:

- Document body and paragraph structure
- Run properties and text formatting
- Tracked changes (insertions, deletions, formatting changes)
- Comments and threaded replies
- Images, drawings, and relationships
- Tables, rows, cells, and styling
- Content types and package relationships

This reference is useful if you need to extend docex or debug OOXML output.

## Design Principles

- **Zero external dependencies.** Only Node.js built-ins (`fs`, `path`, `child_process`, `crypto`, `zlib`). No npm install needed.
- **Tracked changes are the default.** Every edit is visible in Word's review pane unless you explicitly opt out with `.untracked()`.
- **Author set once, applies everywhere.** Call `doc.author("Name")` once and it propagates to all operations.
- **Position selectors read like English.** `doc.after("Methods").insert("text")` does exactly what it says.
- **Single unzip/rezip cycle per save.** Queue as many operations as you want; docex applies them all in one pass. No repeated zip operations means no corruption.
- **Automatic verification after every save.** docex checks that the output is a valid zip, the paragraph count is reasonable, and the file size is within bounds.

## Project Structure

```
docex/
  src/
    docex.js        Main API (DocexEngine, PositionSelector)
    workspace.js    Zip/unzip, file I/O, save + verify
    paragraphs.js   Paragraph manipulation, text replace, insert, delete
    comments.js     Comment add, list, reply (incl. commentsExtended.xml)
    figures.js      Image embedding, relationship management
    tables.js       Table generation (booktabs and plain styles)
    textmap.js      Text-to-XML position mapping
    xml.js          Lightweight XML parser and serializer
  cli/
    docex-cli.js    CLI interface
  test/
    docex.test.js         49 unit tests
    integration.test.js   11 integration tests
    fixtures/             Test .docx files
  reference/
    ooxml-reference.md    OOXML spec reference (2,967 lines)
  presets/                Style presets (planned)
  docs/
    plans/                Development plans
```

## License

MIT. See [LICENSE](LICENSE).
