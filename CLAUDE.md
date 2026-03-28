# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What is docex?

A zero-dependency Node.js library (>=18) for programmatic .docx editing. Tracked changes are on by default. Built for academic manuscript workflows (peer review, journal formatting, citations) but works on any .docx. Includes a CLI, MCP server, and a `.dex` plain-text round-trip format.

## Commands

```bash
# Run all tests (648 tests, ~11s, uses node:test + node:assert)
node --test test/*.test.js

# Run a single test file
node --test test/docex.test.js

# CLI (no install needed)
node cli/docex-cli.js <command> <file> [args] [options]

# MCP server (stdio JSON-RPC)
node src/mcp-server.js

# Website demo (requires npm install in website/)
cd website && npm install && node server.js
```

No `npm install` needed for the library itself. The only dependency is the system `zip`/`unzip` commands (used by workspace.js for .docx pack/unpack).

## Architecture

**Entry point:** `src/docex.js` -- exports the `docex()` factory function and the `DocexEngine` class. The factory returns a `DocexEngine` instance bound to a .docx file path.

**Core data flow:** `.docx` (zip) -> `Workspace` (unzip to temp dir, lazy XML accessors) -> module operations (regex-based XML manipulation) -> `Workspace.save()` (rezip). All operations queue up and apply in a single zip cycle on save.

**Key abstractions:**
- `Workspace` (`workspace.js`) -- unzip/rezip lifecycle, lock files, backup management, lazy XML loading
- `PositionSelector` (`docex.js`) -- fluent `doc.at()` / `doc.after()` / `doc.before()` / `doc.id()` chains
- `TextMap` (`textmap.js`) -- maps plain text offsets to XML node positions for surgical edits
- `xml.js` -- regex-based XML parser/serializer (no DOM, no external deps). Defines OOXML namespace map (`NS`)

**Module categories (48 modules in src/):**

| Layer | Key modules | Role |
|-------|-------------|------|
| Core | `docex.js`, `workspace.js`, `xml.js`, `textmap.js` | API surface, zip lifecycle, XML ops, text mapping |
| Editing | `paragraphs.js`, `formatting.js`, `handle.js`, `range.js` | Text replace/insert/delete, formatting, stable paragraph addressing |
| Comments | `comments.js` | Add/reply/resolve/export (manages 5 OOXML comment files) |
| Media | `figures.js`, `figure-handle.js`, `tables.js`, `table-handle.js` | Image insertion with relationship management, table generation |
| Structure | `docmap.js`, `crossref.js`, `sections.js`, `lists.js`, `footnotes.js`, `headers.js`, `fields.js` | Document map, cross-refs, lists, footnotes |
| Revisions | `revisions.js`, `diff.js` | Tracked changes, two-document comparison |
| Academic | `presets.js`, `verify.js`, `submission.js`, `citations.js`, `response-letter.js`, `template.js` | Journal styles, validation, anonymize, citations |
| Dex format | `dex-decompiler.js`, `dex-compiler.js`, `dex-lossless.js`, `dex-parser.js`, `dex-markdown-parser.js` | `.docx` <-> `.dex` round-trip |
| Export | `latex.js`, `compile.js`, `metadata.js`, `layout.js` | LaTeX/HTML/Markdown export |
| Workflow | `batch.js`, `macros.js`, `production.js`, `workflow.js`, `transaction.js`, `provenance.js`, `quality.js`, `redact.js` | Batch ops, variables, production pipelines |

**The .dex format** is a YAML-frontmatter + markdown-like plain text representation of .docx content. Two compilation paths exist:
- Human-readable: `DexDecompiler._decompileWorkspace()` / `DexParser.parse()` + `DexCompiler._compileHumanReadable()`
- Lossless binary package: `dex-lossless.js` `serializeDex()`/`parseDex()` for perfect round-trips

**Plugin system:** `.claude-plugin/` contains `plugin.json` (skill + MCP server definition) and `marketplace.json` for the Claude Code plugin marketplace.

**Website:** `website/` is an Express app serving a single-page demo. API endpoints: `POST /api/decompile`, `POST /api/build`, `POST /api/preview`. Separate `package.json` with its own `node_modules/`.

**Static docs site:** `docs/` contains GitHub Pages HTML files (`index.html`, `playground.html`, `api.html`, `cli.html`, `dex-format.html`, etc.) with `docs/decompiler.js` providing client-side .dex decompilation.

## Key design decisions

- **Regex-based XML manipulation** -- no DOM parser, no external XML library. `xml.js` provides `parseXml`/`serializeXml` and namespace-aware helpers. This keeps the zero-dependency constraint.
- **Single zip cycle per save** -- all edits queue as pending operations, applied in one `Workspace.save()` call. Prevents corruption from repeated zip/unzip.
- **Lock files** -- `Workspace` uses `.filename.docx.docex-lock` files with PID tracking and reference counting for concurrent access safety.
- **Stable paragraph IDs** -- every paragraph gets a `w14:paraId` that survives edits, enabling `doc.id("3A7F2B1C")` addressing.

## Test fixtures

`test/fixtures/test-manuscript.docx` is the primary test fixture. Tests create temporary .docx files and clean up after themselves. Test output goes to `test/output/` (gitignored).
