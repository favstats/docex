# Dedocs

`dedocs` is a fresh single-file text representation for `.docx` packages.

The design target is strict package fidelity with an editing surface an AI can
follow without guessing:

- every package part is represented explicitly
- XML parts stay raw UTF-8 text
- binary parts are base64
- each part is isolated in a boundary-delimited block
- metadata is rewritten automatically on save

## Shape

```text
\dedocs[version="1", package="docx", fidelity="package-exact", source="paper.docx"]

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
- The inner payload is raw and exact.
- Boundaries avoid brace-escaping games.
- AI edits can focus on the XML payload, not on archive internals.

## Guarantees

- `.docx -> .dedocs -> .docx` preserves package part bytes exactly
- untouched XML is never normalized or reserialized
- zip container metadata is not preserved; package contents are

