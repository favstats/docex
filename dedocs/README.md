# Dedocs

`dedocs` is a fresh single-file text representation for `.docx` packages.

The design target is strict package fidelity with an editing surface an AI can
follow without guessing:

- one file contains both authoring hints and the exact package core
- generated guides expose document structure for fast orientation
- explicit transforms compile onto exact package parts
- every package part is represented explicitly
- XML parts stay raw UTF-8 text
- binary parts are base64
- each part is isolated in a boundary-delimited block
- metadata is rewritten automatically on save

## Shape

```text
\dedocs[version="1", package="docx", fidelity="package-exact", source="paper.docx"]

\guide[name="document-paragraphs", part="word/document.xml", format="paragraphs", boundary=":::DEDOCS_GUIDE_1_abc:::"]
:::DEDOCS_GUIDE_1_abc:::
\p[index="0000", style="Heading1"] Introduction
\p[index="0001"] This is the first paragraph of the introduction.
:::DEDOCS_GUIDE_1_abc:::
\end{guide}

\replace-text[part="word/document.xml", count="1"]
<<<FIND
platform governance and political advertising.
FIND
<<<WITH
platform governance, political advertising, and enforcement.
WITH
\end{replace-text}

\part[path="word/document.xml", mediaType="application/xml", encoding="utf8", bytes="1234", sha256="...", boundary=":::DEDOCS_PART_1_abc:::"]
:::DEDOCS_PART_1_abc:::
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document>...</w:document>
:::DEDOCS_PART_1_abc:::
\end{part}

\part[path="word/media/image1.png", mediaType="image/png", encoding="base64", bytes="2048", sha256="...", boundary=":::DEDOCS_PART_2_def:::"]
:::DEDOCS_PART_2_def:::
iVBORw0KGgoAAAANSUhEUgAA...
:::DEDOCS_PART_2_def:::
\end{part}

\end{dedocs}
```

## Why this syntax

- The outer grammar is regular and command-like.
- Guides give AI-readable structure without becoming the source of truth.
- Transforms express intent explicitly instead of forcing raw XML edits.
- The inner payload is raw and exact.
- Boundaries avoid brace-escaping games.
- AI edits can focus on the XML payload, not on archive internals.

## Guarantees

- `.docx -> .dedocs -> .docx` preserves package part bytes exactly
- guide blocks are advisory only and never affect fidelity
- `replace-text` transforms change only the targeted UTF-8 part
- untouched XML is never normalized or reserialized
- zip container metadata is not preserved; package contents are

## Editing Loop

After hand-editing a `.dedocs` file, run:

```bash
node dedocs/cli.js normalize input.dedocs
```

That refreshes:

- part hashes and byte counts
- safe boundaries
- generated guide blocks
