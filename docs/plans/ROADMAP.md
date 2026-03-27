# docex Roadmap

> v1.0.0 ships only after Fabio has personally verified everything works across real workflows.

## Version Summary

| Version | Focus | Features | Status |
|---------|-------|----------|--------|
| **v0.1.0** | Core editing | 40 | DONE |
| **v0.2.0** | Features + robustness + polish | ~45 | Next |
| **v0.3.0** | Academic intelligence | ~40 | Planned |
| **v0.4.0** | The .dex source format | ~35 | Vision |
| **v1.0.0** | Fabio-verified release | 0 new | Only after manual verification |

---

## v0.1.0 (DONE)

Core editing library. 144 tests. Zero dependencies.

**What shipped:**
- Fluent API: `docex("file.docx").author("X").replace("old","new").save()`
- Position selectors: `doc.at()`, `doc.after()`, `doc.before()`
- Tracked changes by default (w:ins/w:del with author attribution)
- Comments: add, reply, resolve, remove (5-file management)
- Figures: insert, replace, list (PNG/JPEG, auto-dimensions)
- Tables: insert with booktabs/plain style
- Citations: pattern detection + Zotero field code injection
- LaTeX export: full .docx to .tex conversion
- CLI: replace, insert, delete, comment, reply, figure, table, list, cite, latex
- safe-modify.sh integration (--safe flag)
- Auto-verification: paragraph count, file size, valid zip
- OOXML reference manual (2,967 lines)
- TextMap for cross-run text finding

---

## v0.2.0: Features + Robustness + Polish

### Tracked Changes Lifecycle
- `doc.accept()` / `doc.accept(id)` -- accept all or specific tracked change
- `doc.reject()` / `doc.reject(id)` -- reject all or specific
- `doc.revisions()` -- list all tracked changes with IDs, authors, dates
- `doc.cleanCopy()` -- accept all changes, resolve all comments (for final submission)
- CLI: `docex accept`, `docex reject`, `docex clean`

### Document Diff
- `doc.diff("other.docx")` -- compare two documents, produce tracked changes
- Like `latexdiff` but for .docx
- CLI: `docex diff original.docx revised.docx --output diff.docx`

### Inline Formatting API
- `doc.bold("text")` -- make text bold (tracked)
- `doc.italic("text")` -- make text italic
- `doc.underline("text")` -- underline
- `doc.strikethrough("text")` -- strikethrough
- `doc.superscript("text")` -- superscript
- `doc.subscript("text")` -- subscript
- `doc.smallCaps("text")` -- small caps
- `doc.code("text")` -- monospace/code inline
- `doc.color("text", "red")` -- text color
- `doc.highlight("text", "yellow")` -- background highlight

### Footnotes
- `doc.at("text").footnote("Footnote content")` -- insert footnote
- `doc.footnotes()` -- list all footnotes
- Manages word/footnotes.xml and relationships

### Word Count
- `doc.wordCount()` -- returns `{ total, body, abstract, headings, captions, footnotes }`
- Excludes headers, figure captions, table captions by default
- CLI: `docex count manuscript.docx`

### Document Metadata
- `doc.metadata({ title, author, subject, keywords, created, modified })`
- Reads/writes docProps/core.xml
- CLI: `docex meta manuscript.docx --title "The Ban That Wasn't"`

### Robustness
- Mixed-state TextMap: handle text spanning w:ins boundaries (port from docx-editor)
- Comment anchoring on exact phrase (run-level, not paragraph-level)
- Format-preserving replace across formatting boundaries (bold+plain)
- Closest-match error messages: "Text not found. Did you mean: ..."
- Property-based fuzz testing (100 random operations, verify invariants)

### API Enhancements
- `doc.replaceAll("old", "new")` -- replace every occurrence
- `doc.replace(/pattern/, "new")` -- regex-based find/replace
- `doc.at("text", { nth: 2 })` -- select nth occurrence
- `--json` output flag for CLI (machine-readable)
- `--verbose` / `--quiet` flags
- CLI batch mode: `docex batch script.js`

### Publish Readiness
- .npmignore (exclude test/, docs/, reference/ from package)
- Compact OOXML cheat sheet (200 lines, loads fast in skills)
- npm publish as `docex`

---

## v0.3.0: Academic Intelligence

### Cross-References and Labels
- `doc.label("fig:funnel")` -- assign label to current figure/table/heading
- `doc.ref("fig:funnel")` -- insert cross-reference ("Figure 3")
- Uses OOXML SEQ + REF field codes
- Auto-updates when figures/tables reorder
- `doc.cref("fig:funnel")` -- clever reference ("Figure 3" or "Figures 3 and 4")

### Auto-Numbering
- Figures auto-numbered: "Figure 1", "Figure 2"
- Tables auto-numbered: "Table 1", "Table 2"
- SEQ field codes (same as Word's built-in caption numbering)
- `doc.after().figure()` automatically assigns next number

### Lists
- `doc.after("text").bulletList(["item 1", "item 2", "item 3"])`
- `doc.after("text").numberedList(["first", "second", "third"])`
- Nested lists supported
- Custom markers

### Document Composition
- `doc.include("chapters/methods.docx")` -- insert another document's content
- `doc.include("chapter.docx", { after: "Introduction" })` -- at position
- Handles relationship/style conflicts

### Variables and Macros
- `doc.define("NUM_ADS", "268,635")` -- define variable
- `doc.expand()` -- replace all `{{NUM_ADS}}` in document
- Conditional: `{{#if APPENDIX}}...{{/if}}`
- Useful for keeping numbers consistent across abstract + body + conclusion

### Table Enhancements
- Cell merging (colspan/rowspan)
- Column alignment (left/center/right per column)
- Column width specification
- Table from CSV/JSON data file
- Table notes (below table, like threeparttable)
- Table auto-numbering

### Journal Style Presets
- `doc.style("academic")` -- Times New Roman 12pt, double-spaced, justified, 1in margins
- `doc.style("polcomm")` -- Political Communication format
- `doc.style("jcmc")` -- Journal of Computer-Mediated Communication
- `doc.style("joc")` -- Journal of Communication
- `doc.style("apa7")` -- APA 7th edition
- Custom style definition API: `docex.defineStyle("myjournal", { ... })`

### Page Layout
- Page breaks: `doc.pageBreak()`
- Section breaks: `doc.sectionBreak("nextPage")`
- Margins: `doc.margins({ top: 1, bottom: 1, left: 1, right: 1 })`
- Headers/footers: `doc.header("Running title")`, `doc.footer("Page {{page}}")`
- Continuous line numbering

### TOC and Lists
- `doc.toc()` -- generate table of contents
- `doc.lof()` -- list of figures
- `doc.lot()` -- list of tables

### Title Page and Front Matter
- `doc.titlePage({ title, authors, affiliations, date, corresponding })`
- `doc.abstract("Abstract text", { wordLimit: 250 })`
- `doc.keywords(["political advertising", "platform governance"])`
- `doc.appendix()` -- start appendix section

### Citation Enhancements
- Citation verification: check all `(Author, Year)` resolve to .bib entries
- Orphan citation detection
- Unused reference detection
- DOI auto-linking
- APA/Chicago/Harvard citation style formatting
- Et al. threshold control

### Submission Helpers
- `doc.anonymize()` -- remove author names for blind review
- `doc.deanonymize()` -- restore author info
- `doc.highlightedChanges()` -- generate version showing all changes for reviewers
- `doc.responseLetter(comments, responses)` -- generate R&R response letter
- `doc.dataAvailability("Data available at ...")` -- insert statement
- `doc.fundingStatement("Funded by ...")` -- insert statement
- `doc.conflictOfInterest("None declared")` -- insert statement
- Figure resolution check (warn if < 300 DPI)

### Developer Experience
- TypeScript type definitions (.d.ts)
- ESM module support (`import docex from 'docex'`)
- Config file (.docexrc) for default author, safe-modify path, style
- Dry-run mode: `doc.save({ dryRun: true })` -- preview changes without writing

---

## v0.4.0: The .dex Source Format

The endgame: a plain-text markup language that compiles to .docx.

### The .dex Format

```
---
title: The Ban That Wasn't
authors:
  - name: Fabio Votta
    affiliation: University of Amsterdam
    email: f.r.votta@uva.nl
    orcid: 0000-0002-1085-3002
  - name: Simon Kruschinski
    affiliation: GESIS
style: polcomm
bibliography: references.bib
variables:
  NUM_ADS: "268,635"
  NUM_POLITICAL: "1,329"
---

\abstract{
Purpose. In October 2025, Meta imposed a blanket ban...
We collected {{NUM_ADS}} advertisements.
}

\keywords{political advertising, platform governance, Meta}

# Introduction

In October 2025, Meta imposed a blanket ban on political
advertising [@meta2025]. We collected {{NUM_ADS}} ads
(see @fig:funnel for the classification pipeline).

# Methods

## Data Collection

We developed an automated monitoring system that collected
{{NUM_ADS}} advertisements from Meta's Ad Library.

![Classification funnel](figures/fig01.png){#fig:funnel}

# Results

@tbl:top shows the most active advertisers.

| Party | Ads | Share |
|-------|-----|-------|
| PAX   | 117 | 8.8%  |
| Wakker Emmen | 82 | 6.2% |
{#tbl:top caption="Top political advertisers"}

Meta removed only 189 ads (14.2%)[footnote: This is a
conservative estimate; ads removed before monitoring
began would not be captured.].

# Discussion

The findings document a specific failure mode of platform
self-regulation [tracked: self-regulation -> private governance].

<!-- comment by Reviewer 2: Expand this section -->

# References
```

### Compilation
- `docex build manuscript.dex` -- compile to manuscript.docx
- `docex build manuscript.dex --style polcomm` -- override style
- `docex build manuscript.dex --output submission.docx`

### Decompilation
- `docex decompile manuscript.docx` -- extract .dex from existing .docx
- Preserves tracked changes as `[tracked: old -> new]` syntax
- Preserves comments as `<!-- comment by Author: text -->`
- Preserves cross-references as `@fig:label` / `@tbl:label`

### Round-Trip Workflow
1. Write `manuscript.dex`
2. `docex build` produces `manuscript.docx`
3. Collaborators edit in Word/OnlyOffice (tracked changes, comments)
4. `docex decompile` extracts updated `.dex` with changes preserved
5. Author reviews changes in plain text
6. `docex build` produces clean revised version

### Additional v0.4.0 Features
- Math support: `$x^2$` inline, `$$E = mc^2$$` display (OMML generation)
- Code listings with syntax highlighting
- Equation numbering and references
- Subfigures
- Long tables (page-spanning)
- Document creation from scratch: `docex.create()`
- Watch mode: `docex watch manuscript.dex` (rebuild on change)
- Landscape pages
- Bookmarks and hyperlinks
- `.docx` to HTML export
- `.docx` to Markdown export

---

## v1.0.0: Verified Release

**No new features.** This version ships only after:

- [ ] Fabio has used docex across all active paper projects
- [ ] All 7 manuscripts edit correctly with tracked changes visible in OnlyOffice
- [ ] R&R workflow tested end-to-end on a real journal submission
- [ ] .dex format tested on at least 3 papers
- [ ] npm package published and installable by others
- [ ] README verified by someone who isn't Fabio or Mira
- [ ] At least one external user has tested it

---

## Feature Count by Version

| Version | New | Cumulative | Tests (est.) |
|---------|-----|------------|-------------|
| v0.1.0 | 40 | 40 | 144 |
| v0.2.0 | ~45 | ~85 | ~300 |
| v0.3.0 | ~40 | ~125 | ~500 |
| v0.4.0 | ~35 | ~160 | ~700 |
| v1.0.0 | 0 | 160 | ~700 |
