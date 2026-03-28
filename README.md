# docex

## What is docex?

docex lets you edit Word documents (.docx) with code instead of clicking through menus. Think of it as a scripting layer for .docx files. You write commands, docex does the formatting, tracked changes, comments, and all the fiddly XML work that makes Word documents tick.

It was built for academic manuscript workflows (writing papers, responding to peer review, formatting for journal submission), but it works on any .docx file. Zero external dependencies. Just Node.js.

## The .dex format (the killer feature)

docex can convert any .docx file into a plain text format called `.dex`. You can read it, edit it in any text editor, and convert it back to .docx. Nothing is lost -- every font, color, comment, tracked change, footnote, and formatting detail survives the round trip.

Here is what a real .docx looks like as .dex (from a deliberately chaotic test document):

```
---
docex: "0.4.0"
authors:
  - name: "Un-named"
---

# QUARTERLY Strategic SYNERGY Report {id:BABCF972}

{p id:9E7DD00E}
{font "Lucida Handwriting"}{color 999999}Version 47.3b (DRAFT) (FINAL) (FINAL v2) (ACTUALLY FINAL) (USE THIS ONE){/color}{/font}
{/p}

{p id:9861D842}
{font "Arial Black"}{color FF0000}{u}{b}CONFIDENTIAL{/b}{/u}{/color}{/font}
{/p}

{pagebreak}

### Executive Summary (please actually read this one) {id:043C8211}

{comment id:0 by:"The Intern" date:"2024-05-13T22:00:00.000Z"}
WHO WROTE THIS???
{/comment}
```

Compare that to the Word view:

| In Word | In .dex |
|---------|---------|
| Heading with fancy fonts | `# QUARTERLY Strategic SYNERGY Report {id:BABCF972}` |
| Red bold underlined text | `{font "Arial Black"}{color FF0000}{u}{b}CONFIDENTIAL{/b}{/u}{/color}{/font}` |
| A yellow sticky-note comment | `{comment id:0 by:"The Intern"}WHO WROTE THIS???{/comment}` |
| Page break | `{pagebreak}` |

You can convert any .docx to .dex (plain text), edit it in your favorite text editor, and convert back. Nothing is lost.

```bash
docex decompile document.docx          # Word -> plain text (.tex)
# edit the file...
docex compile document.tex             # plain text -> Word
```

## How do I install it?

```bash
git clone https://github.com/favstats/docex.git
cd docex
# That's it. No npm install needed -- zero dependencies.
```

You need Node.js version 18 or newer. Check with `node --version`.

To verify everything works:

```bash
node --test test/*.test.js   # 648 tests should pass
```

## What can it do?

Here are the 10 commands you will use most, with real examples you can copy and paste.

### Replace text (shows as a tracked change in Word)

```bash
docex replace paper.docx "268,635" "300,000" --author "Fabio"
```

Word will show "268,635" crossed out and "300,000" inserted, attributed to "Fabio".

### Add a reviewer comment

```bash
docex comment paper.docx "platform governance" "Needs citation" --by "Reviewer 2"
```

This anchors a comment to the phrase "platform governance" in the document.

### Insert a new paragraph

```bash
docex insert paper.docx "after:Methods" "We used a mixed-methods approach."
```

The text "after:Methods" means "insert after the heading called Methods".

### See what is in the document

```bash
docex list paper.docx headings      # Show all headings
docex list paper.docx comments      # Show all comments
docex list paper.docx revisions     # Show all tracked changes
docex list paper.docx figures       # Show all images
```

### Get a word count

```bash
docex count paper.docx
```

Returns word counts broken down by body text, headings, abstract, captions, and footnotes.

### Check document health

```bash
docex doctor paper.docx
```

Checks for corrupt zip structure, orphaned images, broken relationships, duplicate paragraph IDs, and heading hierarchy problems.

### Compare two versions

```bash
docex diff paper.docx paper-v1.docx --output diff.docx
```

Produces a new .docx with tracked changes showing what changed between the two versions.

### Apply journal formatting

```bash
docex style paper.docx --preset polcomm
```

Built-in presets: `academic`, `polcomm`, `apa7`, `jcmc`, `joc`. Sets fonts, margins, spacing, and headers.

### Accept tracked changes

```bash
docex accept paper.docx          # Accept all changes
docex accept paper.docx 5        # Accept only change #5
```

### Validate for submission

```bash
docex verify paper.docx --preset polcomm
```

Checks word count limits, abstract length, figure resolution, margins, and fonts against journal requirements.

## How do I use it from JavaScript?

docex has a fluent API. Here is a realistic academic workflow:

```js
const docex = require('./src/docex');
const doc = docex("paper.docx");
doc.author("Fabio Votta");

// Respond to reviewer feedback
doc.replace("enforcement gap", "regulatory gap");
doc.at("platform governance").comment("Added Suzor 2019 citation", { by: "Fabio Votta" });
doc.after("Methods").insert("We used a mixed-methods approach combining content analysis with automated classification.");
doc.after("Results").figure("fig03.png", "Figure 3. Ad removal rates by platform.");
doc.after("Results").table(
  [["Party", "Ads", "Spend"], ["VVD", "342", "\u20AC12,450"], ["D66", "289", "\u20AC9,800"]],
  { style: "booktabs", caption: "Table 1. Ad counts by party." }
);

// Format for journal submission
await doc.style("polcomm");

// Save (single zip operation, auto-verifies)
await doc.save();
```

### Position selectors

Position selectors let you target a location in the document:

```js
doc.at("some phrase")          // AT the text (for comments, formatting)
doc.after("Methods")           // AFTER a heading or text
doc.before("Conclusion")       // BEFORE a heading or text
doc.id("3A7F2B1C")            // By paragraph ID (survives other edits)
```

### What you can do at a position

```js
// Insert content
doc.after("Methods").insert("New paragraph.");
doc.after("Results").figure("fig03.png", "Caption.");
doc.after("Results").table([["A","B"],["1","2"]]);
doc.after("Results").bulletList(["Item 1", "Item 2"]);

// Add comments
doc.at("anchor text").comment("Note.", { by: "Reviewer 2" });
doc.at("anchor text").reply("Fixed.", { by: "Fabio Votta" });

// Format text
doc.at("text").bold();
doc.at("text").italic();
doc.at("text").highlight("yellow");
doc.at("text").color("red");
doc.at("text").underline();
doc.at("text").footnote("See appendix.");
```

### Tracked changes

Every edit is tracked by default (shows as strikethrough + insertion in Word). To disable:

```js
doc.untracked();
doc.replace("old", "new");    // direct replacement, no track record
```

### Saving

```js
await doc.save();                           // overwrite original
await doc.save("revised.docx");             // save to new file
await doc.save({ dryRun: true });           // preview without saving
```

## What is the full feature list?

### Editing

| Feature | CLI | API |
|---------|-----|-----|
| Replace text (tracked) | `docex replace file "old" "new"` | `doc.replace("old", "new")` |
| Replace all occurrences | -- | `doc.replaceAll("old", "new")` |
| Insert paragraph | `docex insert file "after:Heading" "text"` | `doc.after("Heading").insert("text")` |
| Delete text (tracked) | `docex delete file "text"` | `doc.delete("text")` |
| Bold | `docex bold file "text"` | `doc.at("text").bold()` |
| Italic | `docex italic file "text"` | `doc.at("text").italic()` |
| Underline | -- | `doc.at("text").underline()` |
| Highlight | `docex highlight file "text"` | `doc.at("text").highlight("yellow")` |
| Color | -- | `doc.at("text").color("red")` |
| Strikethrough | -- | `doc.at("text").strikethrough()` |
| Superscript / Subscript | -- | `doc.at("text").superscript()` |
| Small caps | -- | `doc.at("text").smallCaps()` |
| Code formatting | -- | `doc.at("text").code()` |

### Comments

| Feature | CLI | API |
|---------|-----|-----|
| Add comment | `docex comment file "anchor" "text" --by "Name"` | `doc.at("anchor").comment("text")` |
| Reply to comment | `docex reply file 1 "text" --by "Name"` | `doc.at("anchor").reply("text")` |
| List comments | `docex list file comments` | `await doc.comments()` |
| Export comments | -- | `await doc.exportComments("csv")` |

### Figures and tables

| Feature | CLI | API |
|---------|-----|-----|
| Insert figure | `docex figure file "after:Results" img.png` | `doc.after("Results").figure("img.png", "Caption")` |
| List figures | `docex list file figures` | `await doc.figures()` |
| Insert table | `docex table file "after:Results" data.json` | `doc.after("Results").table(data)` |
| Bullet list | -- | `doc.after("X").bulletList(["a","b"])` |
| Numbered list | -- | `doc.after("X").numberedList(["1","2"])` |

### Tracked changes

| Feature | CLI | API |
|---------|-----|-----|
| List changes | `docex revisions file` | `await doc.revisions()` |
| Accept all | `docex accept file` | `await doc.accept()` |
| Accept one | `docex accept file 5` | `await doc.accept(5)` |
| Reject all | `docex reject file` | `await doc.reject()` |
| Clean copy | `docex clean file` | `await doc.cleanCopy()` |
| Compare two docs | `docex diff a.docx b.docx` | `await doc.diff("b.docx")` |

### Academic tools

| Feature | CLI | API |
|---------|-----|-----|
| Apply journal style | `docex style file --preset polcomm` | `await doc.style("polcomm")` |
| Validate for submission | `docex verify file --preset polcomm` | `await doc.verify("polcomm")` |
| Anonymize for blind review | `docex anonymize file` | `await doc.anonymize()` |
| Word count | `docex count file` | `await doc.wordCount()` |
| Metadata | `docex meta file` | `await doc.metadata()` |
| Footnotes | `docex footnote file "anchor" "text"` | `doc.at("anchor").footnote("text")` |
| Citations | `docex cite file --list` | `await doc.citations()` |
| Zotero injection | `docex cite file --zotero-key X --zotero-user Y` | `await doc.injectCitations(opts)` |
| Response letter | `docex response-letter file --responses r.json` | `await docex.responseLetter(...)` |
| Create from template | `docex template --title "X" --journal polcomm` | `await docex.fromTemplate(...)` |

### Document structure

| Feature | CLI | API |
|---------|-----|-----|
| List headings | `docex list file headings` | `await doc.headings()` |
| List paragraphs | `docex list file paragraphs` | `await doc.paragraphs()` |
| Document map | -- | `await doc.map()` |
| Find text | -- | `await doc.find("text")` |
| Structure tree | -- | `await doc.structure()` |
| Stable paragraph IDs | -- | `doc.id("3A7F2B1C").replace(...)` |
| Cross-references | -- | `await doc.ref("fig:funnel")` |
| Auto-numbering | -- | `await doc.autoNumber()` |
| Variables | `docex expand file --vars '{"X":"1"}'` | `await doc.expand({X:"1"})` |

### Export

| Feature | CLI | API |
|---------|-----|-----|
| LaTeX export | `docex latex file` | `await doc.toLatex()` |
| HTML export | `docex html file` | `await doc.toHtml()` |
| Markdown export | `docex markdown file` | `await doc.toMarkdown()` |
| LaTeX compile | `docex compile paper.tex` | `await docex.compile("paper.tex")` |
| Watch mode | `docex watch paper.tex` | `docex.watch("paper.tex")` |

### Utilities

| Feature | CLI | API |
|---------|-----|-----|
| Doctor (health check) | `docex doctor file` | `await doc.validate()` |
| Create empty .docx | `docex create file` | `await docex.create("file.docx")` |
| Batch operations | -- | `docex.batch(["a.docx","b.docx"])` |
| Preview pending ops | -- | `doc.preview()` |
| Discard pending ops | -- | `doc.discard()` |
| Document stats | -- | `await doc.stats()` |
| Contributors | -- | `await doc.contributors()` |
| Timeline | -- | `await doc.timeline()` |

## How is the source code organized?

47 source modules in `src/`, totaling about 21,000 lines. Grouped by purpose:

| Category | Modules | What they do |
|----------|---------|--------------|
| Core | `docex.js`, `workspace.js`, `xml.js`, `textmap.js` | Main API, zip/unzip lifecycle, XML parsing, text-to-XML mapping |
| Editing | `paragraphs.js`, `formatting.js`, `handle.js`, `range.js` | Replace, insert, delete, bold, italic, highlight, stable addressing |
| Comments | `comments.js` | Add, reply, resolve, export (manages 5 OOXML files) |
| Media | `figures.js`, `figure-handle.js`, `tables.js`, `table-handle.js` | Images, tables with auto-dimensions and relationships |
| Structure | `docmap.js`, `crossref.js`, `sections.js`, `lists.js`, `footnotes.js`, `headers.js`, `fields.js` | Document map, cross-references, lists, footnotes, headers |
| Revisions | `revisions.js`, `diff.js` | Tracked changes, document comparison |
| Academic | `presets.js`, `verify.js`, `submission.js`, `citations.js`, `response-letter.js`, `template.js` | Journal styles, validation, anonymize, citations, R&R letters |
| Export | `latex.js`, `compile.js`, `metadata.js`, `layout.js` | LaTeX/HTML/Markdown export, compile pipeline, metadata |
| Dex format | `dex-decompiler.js`, `dex-compiler.js`, `dex-lossless.js`, `dex-parser.js`, `dex-markdown-parser.js` | .docx to .dex round-trip format |
| Workflow | `batch.js`, `macros.js`, `production.js`, `workflow.js`, `transaction.js`, `provenance.js`, `quality.js`, `redact.js`, `extensions.js` | Batch ops, variables, production workflows |

## How do I run the tests?

```bash
node --test test/*.test.js   # 648 tests, 0 dependencies
```

All tests use Node.js built-in `node:test` and `node:assert`. No test framework needed.

## What are the design principles?

- **Zero external dependencies.** Only Node.js built-ins. No `npm install` needed.
- **Tracked changes by default.** Every edit shows in Word's review pane unless you opt out.
- **Author set once.** Call `doc.author("Name")` and it applies to everything.
- **Position selectors read like English.** `doc.after("Methods").insert("text")` does what it says.
- **Single zip cycle per save.** All operations queue up, then apply in one pass. No corruption from repeated zip/unzip.
- **Auto-verify after save.** Checks valid zip, paragraph count, and file size.
- **Stable addressing.** Every paragraph gets a unique ID that survives other edits.

## License

MIT. See [LICENSE](LICENSE).
