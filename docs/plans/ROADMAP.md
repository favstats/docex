# docex Roadmap

> v1.0.0 ships only after Fabio has personally verified everything works across real workflows.

## The Vision

**.docx files are opaque blobs.** You can't script them, test them, version them, or automate them. Academics are stuck manually formatting, manually responding to reviewers, manually checking word counts, manually switching between journal styles.

**docex makes .docx as programmable as source code.**

### What nobody else can do

| Capability | LaTeX | Word | Pandoc | docex |
|-----------|-------|------|--------|-------|
| Programmatic tracked changes | No | Manual only | No | Yes |
| AI persona peer review | No | No | No | Yes |
| One-command journal formatting | Partially (class files) | Templates (manual) | No | `doc.style("polcomm")` |
| Automated R&R workflow | No | Manual | No | Extract comments, apply changes, generate response letter |
| Compare two document versions | `latexdiff` (PDF only) | Word Compare (GUI) | No | `doc.diff("v2.docx")` produces tracked .docx |
| Batch operations across documents | Makefiles | No | No | One script, 10 manuscripts |
| Verify submission requirements | No | Manual eyeballing | No | `doc.verify("polcomm")` |
| Round-trip: .tex -> .docx -> edit -> .tex | Lossy via pandoc | No | Lossy | Planned (preserve tracked changes) |

### The big picture

```
Academic writes manuscript
       |
       v
  docex.fromTemplate("polcomm")     <- or doc.style() on existing
       |
       v
  Collaborators edit in Word/OnlyOffice (tracked changes, comments)
       |
       v
  docex extracts comments, runs AI persona review, adds more comments
       |
       v
  Submit to journal
       |
       v
  Reviewers respond (more .docx comments)
       |
       v
  docex.responseLetter() + doc.replace/insert (tracked) <- automated R&R
       |
       v
  doc.verify("polcomm")             <- check before resubmission
       |
       v
  doc.cleanCopy()                   <- final version, no markup
       |
       v
  Accepted! doc.style("published")  <- camera-ready formatting
```

Every step is scriptable. Every step is testable. Every step is reproducible.

---

## Version Summary

| Version | Focus | Status |
|---------|-------|--------|
| **v0.1.0** | Core editing | DONE |
| **v0.2.0** | Features + robustness + polish | DONE |
| **v0.3.0** | Academic intelligence | Next |
| **v0.4.0** | LaTeX pipeline + completeness | Planned |
| **v1.0.0** | Fabio-verified release | Only after manual verification |

---

## v0.1.0 (DONE)

Core editing library. 144 tests. Zero dependencies.

Replace, insert, delete (all tracked by default). Comments (add, reply, resolve, remove across 5 XML files). Figures (insert, replace, list). Tables (booktabs/plain). Citations (Zotero field code injection). LaTeX export. CLI (10 commands). safe-modify.sh integration. Auto-verification. OOXML reference manual (2,967 lines). TextMap for cross-run text finding.

---

## v0.2.0 (DONE)

238 tests. 15 modules. 21 CLI commands.

- **Revisions**: accept/reject tracked changes (all or by ID), cleanCopy, list revisions
- **Document diff**: compare two .docx files, produce tracked changes via LCS
- **Inline formatting**: bold, italic, underline, strikethrough, superscript, subscript, smallCaps, code, color, highlight (all with optional tracked formatting via rPrChange)
- **Footnotes**: add, list (manages word/footnotes.xml)
- **Metadata**: get/set title, author, subject, keywords, dates (docProps/core.xml)
- **Word count**: total, body, abstract, headings, captions, footnotes
- **Robustness**: exact comment anchoring (run-level), mixed-state text handling, format-preserving replace, closest-match error messages ("Did you mean: ...")
- **API enhancements**: replaceAll, regex replace, nth occurrence
- **Fuzz testing**: 50 random operations, verify invariants survive
- **CLI**: accept, reject, clean, revisions, bold, italic, highlight, footnote, count, meta, diff, --json flag
- **Publish ready**: .npmignore, OOXML cheatsheet (200 lines)

---

## v0.3.0: Academic Intelligence

### Cross-References and Labels
- `doc.label("fig:funnel")` -- assign label to figure/table/heading
- `doc.ref("fig:funnel")` -- insert cross-reference ("Figure 3")
- Uses OOXML SEQ + REF field codes
- Auto-updates when figures/tables reorder
- `doc.cref("fig:funnel")` -- clever reference ("Figure 3" or "Figures 3 and 4")

### Auto-Numbering
- Figures auto-numbered: "Figure 1", "Figure 2"
- Tables auto-numbered: "Table 1", "Table 2"
- SEQ field codes (same as Word's built-in caption numbering)
- `doc.after().figure()` automatically assigns next number

### Semantic Sections
- `doc.section("Methods")` -- returns all paragraphs in Methods section
- `doc.section("Methods").wordCount()` -- word count for just that section
- `doc.section("Abstract").text()` -- extract abstract text
- `doc.section("Results").figures()` -- figures in Results only
- `doc.goto("Methods").replace(...)` -- scoped operations within a section

### Lists
- `doc.after("text").bulletList(["item 1", "item 2"])`
- `doc.after("text").numberedList(["first", "second"])`
- Nested lists, custom markers

### Document Composition
- `doc.include("chapters/methods.docx")` -- insert another document's content
- `doc.include("chapter.docx", { after: "Introduction" })` -- at position
- Handles relationship/style conflicts

### Variables and Macros
- `doc.define("NUM_ADS", "268,635")` -- define variable
- `doc.expand()` -- replace all `{{NUM_ADS}}` in document
- Conditional: `{{#if APPENDIX}}...{{/if}}`

### Table Enhancements
- Cell merging (colspan/rowspan)
- Column alignment (left/center/right per column)
- Table from CSV/JSON data file
- Table notes (below table, like threeparttable)

### Journal Style Presets
- `doc.style("academic")` -- Times New Roman 12pt, double-spaced, justified, 1in margins
- `doc.style("polcomm")` -- Political Communication format
- `doc.style("jcmc")` -- Journal of Computer-Mediated Communication
- `doc.style("joc")` -- Journal of Communication
- `doc.style("apa7")` -- APA 7th edition
- `docex.defineStyle("myjournal", { ... })` -- custom style definition

### Submission Validation
```js
doc.verify("polcomm")
// { pass: false, errors: [
//   "Word count 8,234 exceeds limit 8,000",
//   "Abstract 312 words, limit is 250",
//   "Figure 3 resolution 72 DPI, minimum 300",
//   "Missing running header",
//   "References not in APA format"
// ]}
```

### Batch Processing
```js
const docs = docex.batch(["paper1.docx", "paper2.docx", "paper3.docx"]);
docs.author("Fabio Votta");
docs.style("apa7");
docs.replaceAll("self-regulation", "private governance");
await docs.saveAll();
```

### Document Templates
```js
const doc = docex.fromTemplate("polcomm", {
  title: "The Ban That Wasn't",
  authors: [{ name: "Fabio Votta", affiliation: "UvA" }],
  abstract: "In October 2025...",
  keywords: ["political advertising", "platform governance"]
});
```

### Response Letter Generator
```js
const comments = await doc.comments();
const responses = {
  1: { action: "agree", text: "We added the citation.", changes: ["para 3"] },
  2: { action: "partial", text: "Expanded but maintained framing.", changes: ["Discussion"] },
};
const letter = docex.responseLetter(comments, responses, {
  title: "The Ban That Wasn't", journal: "Political Communication"
});
await letter.save("response_letter.docx");
```

### Page Layout
- Page breaks, section breaks
- Margins control
- Headers/footers with running title
- Continuous line numbering
- TOC, list of figures, list of tables

### Title Page and Front Matter
- `doc.titlePage({ title, authors, affiliations, date, corresponding })`
- `doc.abstract("text", { wordLimit: 250 })`
- `doc.keywords([...])`
- `doc.appendix()`

### Accessibility
```js
doc.accessibility()
// { issues: ["Figure 3 missing alt text", "Heading levels skip H2 to H4"] }
doc.figures.setAltText("fig03", "Bar chart showing enforcement rates");
```

### Citation Enhancements
- Citation verification (check all refs resolve)
- Orphan/unused reference detection
- DOI auto-linking
- APA/Chicago/Harvard style formatting
- Et al. threshold control

### Submission Helpers
- `doc.anonymize()` / `doc.deanonymize()` -- blind review
- `doc.highlightedChanges()` -- version showing all changes
- `doc.dataAvailability("...")`, `doc.fundingStatement("...")`, `doc.conflictOfInterest("...")`
- Figure resolution check (warn if < 300 DPI)

### Developer Experience
- TypeScript type definitions (.d.ts)
- ESM module support
- Config file (.docexrc) for default author, safe-modify path, style
- `docex init` CLI command
- `--dry-run` flag on every CLI command
- `doc.save({ dryRun: true })` -- preview changes without writing

### Convenience
- `doc.undo()` -- revert last operation before saving
- `doc.preview()` -- print summary of pending operations
- `doc.find("text")` -- check if/where text exists (paragraph number, heading context)
- `doc.structure()` -- tree view of document structure
- `doc.explain("text")` -- show XML structure around text (debugging)

### Usability
- Colored terminal diff for revisions (red deletions, green insertions)
- `docex doctor manuscript.docx` -- diagnose problems (corrupt zip, orphaned images, broken refs)
- `doc.validate()` -- check document health without modifying

### Robustness
- Automatic backup before every save
- Lock file to prevent concurrent edits
- Retry with fuzzy matching when exact match fails

### Academic Team Features
- `doc.contributors()` -- list unique authors from changes and comments
- `doc.timeline()` -- chronological list of edits and comments
- `doc.stats()` -- combined: word count + figures + tables + citations + comments + revisions
- `doc.exportComments("csv")` -- export comments for coding/analysis
- `doc.compareStyles("polcomm")` -- preview what would change

### Operation History
```js
const log = doc.history();
// [{ op: "replace", old: "268,635", new: "300,000", author: "Fabio", date: "..." }, ...]
// Stored as custom XML part in the .docx
```

---

## v0.4.0: LaTeX Pipeline + Completeness

v0.4.0 replaces the original .dex format vision with a more practical approach: use LaTeX as the source format, with Pandoc doing the heavy lifting and docex handling the "last mile" that Pandoc can't.

### LaTeX -> .docx Pipeline
```
.tex source
  |
  pandoc (parses LaTeX, produces raw .docx)
  |
  docex post-process:
    - Journal-specific formatting
    - Proper cross-references via field codes
    - Auto-numbered figures/tables
    - Zotero citation fields
    - Tracked changes for R&R
    - Comment injection
  |
  submission-ready .docx
```

- `docex compile paper.tex --style polcomm --output submission.docx`
- `docex compile paper.tex --bib references.bib --csl apa.csl`
- Pandoc handles LaTeX parsing; docex handles OOXML features Pandoc can't do

### .docx -> .tex Reverse Pipeline
```
.docx (from collaborators with tracked changes)
  |
  docex extract (comments, tracked changes, metadata)
  |
  pandoc (converts to .tex)
  |
  docex post-process (preserves changes as \replaced{}{})
  |
  .tex with changes visible
```

- `docex decompile manuscript.docx --output paper.tex`
- Preserves tracked changes as `\replaced{old}{new}` (changes package)
- Preserves comments as `\todo{text}` (todonotes package)

### Additional v0.4.0 Features
- Math support: inline and display equations (OMML generation)
- Code listings with syntax highlighting
- Subfigures
- Long tables (page-spanning)
- Document creation from scratch: `docex.create()`
- Watch mode: `docex watch paper.tex` (rebuild on change)
- Landscape pages
- Bookmarks and hyperlinks
- .docx to HTML export
- .docx to Markdown export
- Tab completion for CLI
- Progress bar for batch operations

---

## v1.0.0: Verified Release

**No new features.** This version ships only after:

- [ ] Fabio has used docex across all active paper projects
- [ ] All 7 manuscripts edit correctly with tracked changes visible in OnlyOffice
- [ ] R&R workflow tested end-to-end on a real journal submission
- [ ] LaTeX compile pipeline tested on at least 3 papers
- [ ] npm package published and installable by others
- [ ] README verified by someone who isn't Fabio or Mira
- [ ] At least one external user has tested it

---

## Feature Count by Version

| Version | New | Cumulative | Tests |
|---------|-----|------------|-------|
| v0.1.0 | 40 | 40 | 144 |
| v0.2.0 | ~45 | ~85 | 238 |
| v0.3.0 | ~60 | ~145 | ~500 (est.) |
| v0.4.0 | ~30 | ~175 | ~650 (est.) |
| v1.0.0 | 0 | 175 | ~650 (est.) |
