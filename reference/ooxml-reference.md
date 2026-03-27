# OOXML Reference for .docx Generation and Editing

> The "LaTeX manual" equivalent for Office Open XML (OOXML) WordprocessingML documents.
>
> This document provides everything an AI agent or human developer needs to
> programmatically generate, edit, and validate `.docx` files without guessing.

---

## Table of Contents

1. [Document Structure](#1-document-structure)
2. [Core XML Elements](#2-core-xml-elements)
3. [Namespaces](#3-namespaces)
4. [Common Patterns](#4-common-patterns)
5. [Pitfalls and Gotchas](#5-pitfalls-and-gotchas)
6. [The LaTeX Equivalence Table](#6-the-latex-equivalence-table)

---

## 1. Document Structure

### 1.1 The .docx Zip Archive

A `.docx` file is a ZIP archive (using deflate compression). When you unzip it, you
get a directory tree like this:

```
mydocument.docx (ZIP)
+-- [Content_Types].xml
+-- _rels/
|   +-- .rels
+-- word/
|   +-- document.xml              # Main document body
|   +-- styles.xml                # Style definitions
|   +-- settings.xml              # Document settings
|   +-- fontTable.xml             # Font declarations
|   +-- numbering.xml             # List numbering definitions
|   +-- footnotes.xml             # Footnotes
|   +-- endnotes.xml              # Endnotes
|   +-- comments.xml              # Comment bodies
|   +-- commentsExtended.xml      # Comment threading (w15)
|   +-- commentsIds.xml           # Durable comment IDs (w16cid)
|   +-- commentsExtensible.xml    # Extensible comment data (w16cex)
|   +-- header1.xml               # Header (can be multiple)
|   +-- footer1.xml               # Footer (can be multiple)
|   +-- theme/
|   |   +-- theme1.xml            # Theme colors, fonts
|   +-- media/
|   |   +-- image1.png            # Embedded images
|   |   +-- image2.jpeg
|   +-- _rels/
|       +-- document.xml.rels     # Relationships for document.xml
|       +-- comments.xml.rels     # Relationships for comments (if images in comments)
|       +-- header1.xml.rels      # Relationships for headers
+-- docProps/
    +-- app.xml                   # Application metadata
    +-- core.xml                  # Core properties (title, author, dates)
    +-- custom.xml                # Custom properties
```

Not all files are required. The minimal set is:

- `[Content_Types].xml`
- `_rels/.rels`
- `word/document.xml`
- `word/_rels/document.xml.rels`

### 1.2 Content Types: `[Content_Types].xml`

This file declares the MIME type of every part in the archive. It lives at the root
of the zip. Word will refuse to open the file if a referenced part has no content type.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- Default types by file extension -->
  <Default Extension="rels"
           ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"
           ContentType="application/xml"/>
  <Default Extension="png"
           ContentType="image/png"/>
  <Default Extension="jpeg"
           ContentType="image/jpeg"/>
  <Default Extension="jpg"
           ContentType="image/jpeg"/>
  <Default Extension="gif"
           ContentType="image/gif"/>
  <Default Extension="svg"
           ContentType="image/svg+xml"/>
  <Default Extension="emf"
           ContentType="image/x-emf"/>
  <Default Extension="wmf"
           ContentType="image/x-wmf"/>

  <!-- Override types for specific parts -->
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/numbering.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/footnotes.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/comments.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
  <Override PartName="/word/commentsIds.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>
  <Override PartName="/word/commentsExtensible.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"/>
  <Override PartName="/word/header1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/theme/theme1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml"
            ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>
```

**Key rules:**

- `<Default>` matches by file extension -- applies to ALL files with that extension.
- `<Override>` matches by exact part name (path within the zip, with leading `/`).
- If you add a `.png` image, you need either a `<Default Extension="png" .../>` or
  a specific `<Override>` for that part.
- If you add a new XML part (e.g., `commentsExtended.xml`), you MUST add a
  corresponding `<Override>` entry or Word will ignore/corrupt it.

### 1.3 Relationships

Relationships connect parts to each other. There are two levels:

#### 1.3.1 Package-Level Relationships: `_rels/.rels`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                Target="docProps/core.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                Target="docProps/app.xml"/>
  <Relationship Id="rId4"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
                Target="docProps/custom.xml"/>
</Relationships>
```

#### 1.3.2 Part-Level Relationships: `word/_rels/document.xml.rels`

This file maps relationship IDs (e.g., `rId7`) referenced inside `document.xml` to
their target parts or external URIs.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
                Target="settings.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
                Target="fontTable.xml"/>
  <Relationship Id="rId4"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                Target="theme/theme1.xml"/>
  <Relationship Id="rId5"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
                Target="numbering.xml"/>
  <Relationship Id="rId6"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
                Target="footnotes.xml"/>
  <Relationship Id="rId7"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
                Target="endnotes.xml"/>
  <Relationship Id="rId8"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
                Target="comments.xml"/>
  <Relationship Id="rId9"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentsExtended"
                Target="commentsExtended.xml"/>
  <Relationship Id="rId10"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentsIds"
                Target="commentsIds.xml"/>
  <Relationship Id="rId11"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
                Target="header1.xml"/>
  <Relationship Id="rId12"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
                Target="footer1.xml"/>

  <!-- Images -->
  <Relationship Id="rId20"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>

  <!-- Hyperlinks (external, note TargetMode) -->
  <Relationship Id="rId30"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="https://example.com"
                TargetMode="External"/>
</Relationships>
```

**Key rules:**

- `Id` must be unique within a `.rels` file. Convention is `rId` + integer.
- `Target` is relative to the part that owns the `.rels` file (so
  `word/_rels/document.xml.rels` targets are relative to `word/`).
- External hyperlinks MUST have `TargetMode="External"`.
- When adding a new image, you must: (1) place the file in `word/media/`, (2) add a
  `<Relationship>` in `word/_rels/document.xml.rels`, (3) ensure a content type
  `<Default>` or `<Override>` exists for the file extension.

### 1.4 Common Relationship Types

| Short Name | Full Type URI |
|---|---|
| officeDocument | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument` |
| styles | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles` |
| settings | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings` |
| fontTable | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable` |
| theme | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme` |
| numbering | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering` |
| footnotes | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes` |
| endnotes | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes` |
| comments | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` |
| commentsExtended | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentsExtended` |
| commentsIds | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentsIds` |
| header | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/header` |
| footer | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer` |
| image | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image` |
| hyperlink | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink` |
| chart | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart` |
| oleObject | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject` |

### 1.5 The Main Document: `word/document.xml`

This is the heart of the `.docx`. It contains the document body as a sequence of
block-level elements (paragraphs, tables, section properties).

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
    xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14">
  <w:body>
    <!-- Block-level elements go here: <w:p>, <w:tbl>, <w:sdt> -->

    <w:p w14:paraId="00000001" w14:textId="77777777" w:rsidR="00A00001" w:rsidRDefault="00A00001">
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r w:rsidRPr="00A00001">
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>

    <!-- Last element in body is always sectPr (section properties) -->
    <w:sectPr w:rsidR="00A00001">
      <w:headerReference w:type="default" r:id="rId11"/>
      <w:footerReference w:type="default" r:id="rId12"/>
      <w:pgSz w:w="12240" w:h="15840"/>  <!-- Letter: 8.5" x 11" in twips -->
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>
```

**Units used in sectPr:**

| Unit | Meaning | Conversion |
|------|---------|------------|
| Twips | 1/20 of a point | 1440 twips = 1 inch |
| Half-points | 1/2 of a point | 24 half-points = 12pt font |
| EMUs | English Metric Units | 914400 EMU = 1 inch |

**Page size values (in twips):**

| Paper Size | w:w | w:h |
|-----------|-----|-----|
| US Letter (8.5 x 11) | 12240 | 15840 |
| A4 (210 x 297mm) | 11906 | 16838 |
| US Legal (8.5 x 14) | 12240 | 20160 |

---

## 2. Core XML Elements

### 2.1 Paragraphs and Runs

#### 2.1.1 `<w:p>` - Paragraph

The fundamental block-level element. Every piece of visible text in a `.docx` lives
inside a paragraph.

```xml
<w:p w14:paraId="1A2B3C4D" w14:textId="5E6F7A8B"
     w:rsidR="00AB1234" w:rsidRDefault="00AB1234" w:rsidRPr="00AB1234">
  <w:pPr>
    <!-- paragraph properties -->
  </w:pPr>
  <w:r>
    <!-- one or more runs -->
  </w:r>
</w:p>
```

**Attributes on `<w:p>`:**

| Attribute | Namespace | Description |
|-----------|-----------|-------------|
| `w14:paraId` | w14 | Unique 8-hex-digit paragraph identifier (e.g., `"1A2B3C4D"`). Used by comments (paraIdParent) and change tracking. MUST be unique across the entire document, comments, footnotes, endnotes, headers, and footers. |
| `w14:textId` | w14 | 8-hex-digit text change identifier. Updated when paragraph content changes. `"77777777"` means "unchanged" by convention. |
| `w:rsidR` | w | Revision save ID -- identifies which editing session created this paragraph. 8-hex-digit, e.g., `"00AB1234"`. |
| `w:rsidRDefault` | w | RSID of the session that last set the default run properties. |
| `w:rsidRPr` | w | RSID of the session that last modified run properties. |
| `w:rsidDel` | w | RSID of the session that deleted this paragraph. |
| `w:rsidP` | w | RSID of the session that last modified paragraph properties. |

**Important:** `w14:paraId` values must be unique across ALL parts of the document
(document.xml, comments.xml, footnotes.xml, endnotes.xml, headers, footers). When
generating new paragraphs, scan all parts for existing paraId values and choose one
that does not collide. Use uppercase hex, 8 characters, no leading zeros stripped.

#### 2.1.2 `<w:pPr>` - Paragraph Properties

Controls paragraph-level formatting. Must be the FIRST child of `<w:p>` if present.

```xml
<w:pPr>
  <!-- Style reference -->
  <w:pStyle w:val="Heading1"/>

  <!-- Keep with next paragraph (no page break between) -->
  <w:keepNext/>

  <!-- Keep all lines together on same page -->
  <w:keepLines/>

  <!-- Page break before this paragraph -->
  <w:pageBreakBefore/>

  <!-- Numbering (for lists) -->
  <w:numPr>
    <w:ilvl w:val="0"/>        <!-- Indent level: 0=first, 1=second, etc. -->
    <w:numId w:val="1"/>       <!-- References numbering.xml definition -->
  </w:numPr>

  <!-- Suppresses line numbering for this paragraph -->
  <w:suppressLineNumbers/>

  <!-- Borders -->
  <w:pBdr>
    <w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/>
    <w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/>
    <w:left w:val="single" w:sz="4" w:space="4" w:color="000000"/>
    <w:right w:val="single" w:sz="4" w:space="4" w:color="000000"/>
    <w:between w:val="single" w:sz="4" w:space="1" w:color="000000"/>
  </w:pBdr>

  <!-- Shading (background color) -->
  <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>  <!-- Yellow bg -->

  <!-- Tabs -->
  <w:tabs>
    <w:tab w:val="left" w:pos="720"/>    <!-- 0.5 inch -->
    <w:tab w:val="center" w:pos="4680"/> <!-- center tab -->
    <w:tab w:val="right" w:pos="9360"/>  <!-- right tab -->
  </w:tabs>

  <!-- Suppress auto-hyphenation -->
  <w:suppressAutoHyphens/>

  <!-- Spacing -->
  <w:spacing
    w:before="240"      <!-- Space before paragraph in twips (240 = 12pt) -->
    w:after="200"       <!-- Space after paragraph in twips -->
    w:line="276"        <!-- Line spacing: 240 = single, 276 ~1.15, 360 = 1.5, 480 = double -->
    w:lineRule="auto"   <!-- "auto" = multiple of line height; "exact" = exact twips; "atLeast" = minimum -->
  />

  <!-- Indentation -->
  <w:ind
    w:left="720"        <!-- Left indent in twips (720 = 0.5 inch) -->
    w:right="0"         <!-- Right indent -->
    w:firstLine="360"   <!-- First-line indent (positive = indent) -->
    w:hanging="360"     <!-- Hanging indent (for bullets/lists) -->
  />
  <!-- Note: firstLine and hanging are mutually exclusive -->

  <!-- Justification / Alignment -->
  <w:jc w:val="both"/>
  <!-- Possible values: "left", "center", "right", "both" (justified),
       "distribute" (distributed), "start", "end" -->

  <!-- Outline level (for TOC) -->
  <w:outlineLvl w:val="0"/>  <!-- 0 = Heading 1 level, 1 = Heading 2, etc. -->

  <!-- Default run properties for the paragraph (applied to paragraph mark) -->
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
    <w:sz w:val="24"/>  <!-- 12pt in half-points -->
  </w:rPr>
</w:pPr>
```

**Spacing `w:line` values for common line spacings:**

| Line Spacing | w:line value | w:lineRule |
|-------------|-------------|------------|
| Single | 240 | auto |
| 1.15 (Word default) | 276 | auto |
| 1.5 | 360 | auto |
| Double | 480 | auto |
| Exact 12pt | 240 | exact |
| At least 12pt | 240 | atLeast |

#### 2.1.3 `<w:r>` - Run

A run is a contiguous span of text with uniform formatting. A paragraph typically
contains one or more runs.

```xml
<w:r w:rsidR="00AB1234" w:rsidRPr="00CD5678">
  <w:rPr>
    <!-- run properties (formatting) -->
  </w:rPr>
  <w:t xml:space="preserve">Hello, World!</w:t>
</w:r>
```

**Attributes on `<w:r>`:**

| Attribute | Description |
|-----------|-------------|
| `w:rsidR` | RSID of the session that created this run |
| `w:rsidRPr` | RSID of the session that last modified run properties |
| `w:rsidDel` | RSID of the session that deleted this run |

A run can contain (in addition to `<w:rPr>`):

| Child Element | Description |
|-------------|-------------|
| `<w:t>` | Text content |
| `<w:tab/>` | Tab character |
| `<w:br/>` | Line break (`w:type="page"` for page break, `w:type="column"` for column break, omit type for line break) |
| `<w:cr/>` | Carriage return (legacy, prefer `<w:br/>`) |
| `<w:sym>` | Symbol character |
| `<w:drawing>` | Inline or floating image/shape |
| `<w:footnoteReference>` | Footnote reference |
| `<w:endnoteReference>` | Endnote reference |
| `<w:commentReference>` | Comment anchor |
| `<w:fldChar>` | Field character (begin/separate/end) |
| `<w:instrText>` | Field instruction text |
| `<w:lastRenderedPageBreak/>` | Marker where Word last broke a page (informational only) |
| `<w:softHyphen/>` | Optional hyphen |
| `<w:noBreakHyphen/>` | Non-breaking hyphen |

#### 2.1.4 `<w:rPr>` - Run Properties

Controls character-level formatting. Must be the FIRST child of `<w:r>` if present.

```xml
<w:rPr>
  <!-- Font -->
  <w:rFonts
    w:ascii="CMU Serif"            <!-- Latin characters -->
    w:hAnsi="CMU Serif"            <!-- High-ANSI characters -->
    w:eastAsia="MS Mincho"         <!-- East Asian characters -->
    w:cs="Times New Roman"         <!-- Complex script (Arabic, Hebrew) -->
    w:asciiTheme="minorHAnsi"      <!-- Theme font override for ASCII -->
    w:hAnsiTheme="minorHAnsi"      <!-- Theme font override for HANSI -->
  />

  <!-- Bold -->
  <w:b/>                           <!-- Bold on -->
  <w:b w:val="0"/>                 <!-- Bold explicitly off (overrides style) -->
  <w:bCs/>                         <!-- Bold for complex script -->

  <!-- Italic -->
  <w:i/>                           <!-- Italic on -->
  <w:i w:val="0"/>                 <!-- Italic off -->
  <w:iCs/>                         <!-- Italic for complex script -->

  <!-- Underline -->
  <w:u w:val="single"/>            <!-- single, double, thick, dotted, dash, wave, etc. -->
  <w:u w:val="single" w:color="FF0000"/>  <!-- Colored underline -->
  <w:u w:val="none"/>              <!-- Remove underline -->

  <!-- Strikethrough -->
  <w:strike/>                      <!-- Single strikethrough -->
  <w:dstrike/>                     <!-- Double strikethrough -->

  <!-- Font size (in HALF-POINTS) -->
  <w:sz w:val="24"/>               <!-- 12pt (24 half-points) -->
  <w:szCs w:val="24"/>             <!-- Complex script size -->

  <!-- Font color -->
  <w:color w:val="FF0000"/>        <!-- Red text (6-digit hex, no # prefix) -->
  <w:color w:val="auto"/>          <!-- Automatic (usually black) -->
  <w:color w:themeColor="accent1"/><!-- Theme color reference -->

  <!-- Highlight (preset colors only) -->
  <w:highlight w:val="yellow"/>
  <!-- Values: black, blue, cyan, darkBlue, darkCyan, darkGray, darkGreen,
       darkMagenta, darkRed, darkYellow, green, lightGray, magenta, none,
       red, white, yellow -->

  <!-- Shading (arbitrary RGB background) -->
  <w:shd w:val="clear" w:color="auto" w:fill="FFCCCC"/>

  <!-- Superscript / Subscript -->
  <w:vertAlign w:val="superscript"/>
  <w:vertAlign w:val="subscript"/>
  <w:vertAlign w:val="baseline"/>  <!-- Reset to normal -->

  <!-- Small caps / All caps -->
  <w:smallCaps/>
  <w:caps/>

  <!-- Character spacing adjustments -->
  <w:spacing w:val="20"/>          <!-- Expanded by 1pt (in twips) -->
  <w:spacing w:val="-20"/>         <!-- Condensed by 1pt -->
  <w:kern w:val="24"/>             <!-- Kern above this size (half-points) -->

  <!-- Language -->
  <w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>

  <!-- Run style reference -->
  <w:rStyle w:val="Hyperlink"/>

  <!-- Vanish (hidden text) -->
  <w:vanish/>

  <!-- Revision (format change tracking) properties -->
  <w:rPrChange w:id="50" w:author="Fabio Votta" w:date="2026-03-27T10:00:00Z">
    <w:rPr>
      <!-- Previous formatting before the change -->
      <w:b/>
    </w:rPr>
  </w:rPrChange>
</w:rPr>
```

**Font size reference:**

| Point Size | w:sz value (half-points) |
|-----------|-------------------------|
| 8pt | 16 |
| 9pt | 18 |
| 10pt | 20 |
| 10.5pt | 21 |
| 11pt | 22 |
| 12pt | 24 |
| 14pt | 28 |
| 16pt | 32 |
| 18pt | 36 |
| 20pt | 40 |
| 24pt | 48 |
| 28pt | 56 |
| 36pt | 72 |

#### 2.1.5 `<w:t>` - Text Content

```xml
<w:t xml:space="preserve">Hello, World!</w:t>
```

**Critical:** Always include `xml:space="preserve"` if the text contains leading
spaces, trailing spaces, or consists entirely of whitespace. Without it, XML parsers
will strip leading/trailing whitespace per the XML specification.

- Text MUST be valid XML: escape `<` as `&lt;`, `>` as `&gt;`, `&` as `&amp;`.
- Newlines within `<w:t>` are NOT rendered as line breaks -- use `<w:br/>` instead.
- An empty paragraph still needs at least one `<w:r><w:t/></w:r>` or can be empty
  (just `<w:p><w:pPr>...</w:pPr></w:p>`).

#### 2.1.6 Run Splitting

Word frequently splits what appears to be a single text span into multiple runs.
This happens when:

1. **Spell check** marks a portion of text (rsid tracking)
2. **Revision tracking** records each editing session separately
3. **Formatting changes** apply to only part of the text
4. **Language detection** switches mid-paragraph
5. **Bookmarks or comments** start/end mid-word

Example: "Hello World" may become:

```xml
<w:p>
  <w:r w:rsidR="00A00001">
    <w:t xml:space="preserve">Hel</w:t>
  </w:r>
  <w:r w:rsidR="00B00002">
    <w:t xml:space="preserve">lo </w:t>
  </w:r>
  <w:r w:rsidR="00C00003">
    <w:t>World</w:t>
  </w:r>
</w:p>
```

**Implications for editing:**

- When searching for text in a paragraph, you MUST concatenate all `<w:t>` values
  across all runs in that paragraph.
- When replacing text, you may need to modify multiple runs or collapse them.
- Never assume a word is contained within a single `<w:r>`.

#### 2.1.7 Styles vs. Direct Formatting

OOXML has a cascading formatting model:

1. **Document defaults** (`<w:docDefaults>` in styles.xml)
2. **Table style** (if in a table)
3. **Paragraph style** (`<w:pStyle>` in `<w:pPr>`)
4. **Run style** (`<w:rStyle>` in `<w:rPr>`)
5. **Direct formatting** (explicit properties in `<w:pPr>` or `<w:rPr>`)

Direct formatting overrides everything. A `<w:b/>` in `<w:rPr>` makes text bold
regardless of what the style says.

To REMOVE bold that a style applies, you must explicitly set `<w:b w:val="0"/>`.
Omitting `<w:b>` entirely means "inherit from style."

---

### 2.2 Tracked Changes (Revisions)

#### 2.2.1 `<w:ins>` - Insertion

Wraps content that was inserted with change tracking enabled.

```xml
<w:ins w:id="1" w:author="Fabio Votta" w:date="2026-03-27T14:30:00Z">
  <w:r w:rsidR="00AB1234">
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="24"/>
    </w:rPr>
    <w:t xml:space="preserve">newly inserted text</w:t>
  </w:r>
</w:ins>
```

**Attributes:**

| Attribute | Required | Description |
|-----------|----------|-------------|
| `w:id` | Yes | Unique integer ID across ALL revision elements in the document. |
| `w:author` | Yes | Display name of the author who made the change. |
| `w:date` | Yes | ISO 8601 timestamp (e.g., `"2026-03-27T14:30:00Z"`). |

**Where `<w:ins>` can appear:**

- As a child of `<w:p>` (wrapping one or more `<w:r>` elements)
- As a child of `<w:body>`, `<w:tc>`, `<w:endnote>`, `<w:footnote>`, `<w:comment>`
  -- wrapping an entire `<w:p>` to track paragraph insertion
- As a child of `<w:tr>` (wrapping `<w:tc>` for tracked cell insertion)
- As a child of `<w:tbl>` (wrapping `<w:tr>` for tracked row insertion)

#### 2.2.2 `<w:del>` - Deletion

Wraps content that was deleted with change tracking enabled.

```xml
<w:del w:id="2" w:author="Fabio Votta" w:date="2026-03-27T14:30:00Z">
  <w:r w:rsidR="00AB1234" w:rsidDel="00CD5678">
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="24"/>
    </w:rPr>
    <w:delText xml:space="preserve">old deleted text</w:delText>
  </w:r>
</w:del>
```

**Critical:** Inside `<w:del>`, text MUST use `<w:delText>`, NOT `<w:t>`. Using
`<w:t>` inside a deletion will cause the text to display as normal (not deleted).

#### 2.2.3 `<w:delText>` vs `<w:t>`

| Element | Used Inside | Meaning |
|---------|-------------|---------|
| `<w:t>` | `<w:r>` (normal or inside `<w:ins>`) | Visible/inserted text |
| `<w:delText>` | `<w:r>` inside `<w:del>` | Deleted text (shown with strikethrough in track changes view) |

Both support `xml:space="preserve"`.

#### 2.2.4 `<w:rPrChange>` - Run Property Change

Records that run formatting was changed. Lives INSIDE `<w:rPr>`. The current
`<w:rPr>` shows the NEW formatting; `<w:rPrChange>` records the OLD formatting.

```xml
<w:r>
  <w:rPr>
    <!-- Current (new) formatting: bold + italic -->
    <w:b/>
    <w:i/>
    <w:rPrChange w:id="3" w:author="Fabio Votta" w:date="2026-03-27T14:30:00Z">
      <w:rPr>
        <!-- Previous (old) formatting: only bold, no italic -->
        <w:b/>
      </w:rPr>
    </w:rPrChange>
  </w:rPr>
  <w:t>formatted text</w:t>
</w:r>
```

This records: "Fabio added italic to this run on 2026-03-27."

#### 2.2.5 `<w:pPrChange>` - Paragraph Property Change

Records that paragraph formatting was changed. Lives INSIDE `<w:pPr>`.

```xml
<w:pPr>
  <!-- Current (new) paragraph properties -->
  <w:jc w:val="center"/>
  <w:pPrChange w:id="4" w:author="Fabio Votta" w:date="2026-03-27T14:30:00Z">
    <w:pPr>
      <!-- Previous (old) paragraph properties -->
      <w:jc w:val="left"/>
    </w:pPr>
  </w:pPrChange>
</w:pPr>
```

This records: "Fabio changed alignment from left to center."

#### 2.2.6 Tracked Paragraph Insertion

To track the insertion of an entire paragraph, wrap the `<w:p>` in `<w:ins>`:

```xml
<w:ins w:id="5" w:author="Fabio Votta" w:date="2026-03-27T14:30:00Z">
  <w:p w14:paraId="2A3B4C5D" w14:textId="6E7F8A9B" w:rsidR="00AB1234" w:rsidRDefault="00AB1234">
    <w:pPr>
      <w:pStyle w:val="Normal"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:sz w:val="24"/>
      </w:rPr>
      <w:t>This entire paragraph was inserted.</w:t>
    </w:r>
  </w:p>
</w:ins>
```

**Note:** When `<w:ins>` wraps an entire `<w:p>` at the body/cell level, the runs
inside do NOT also need to be wrapped in `<w:ins>`. The paragraph-level insertion
implies all content is new.

#### 2.2.7 Tracked Paragraph Deletion (Paragraph Mark)

Deleting a paragraph break (merging two paragraphs) is tracked as:

```xml
<w:p>
  <w:pPr>
    <w:rPr>
      <w:del w:id="6" w:author="Fabio Votta" w:date="2026-03-27T14:30:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:t>Text from first paragraph</w:t>
  </w:r>
</w:p>
```

The `<w:del>` inside `<w:pPr><w:rPr>` means "the paragraph mark was deleted" --
i.e., this paragraph was merged with the next.

#### 2.2.8 RSID Attributes

RSIDs (Revision Save IDs) are 8-hex-digit identifiers assigned to each editing
session. They are NOT the same as tracked change IDs. RSIDs help Word reconstruct
editing history even without formal tracked changes.

| Attribute | Applies To | Meaning |
|-----------|-----------|---------|
| `w:rsidR` | paragraph, run | Session that created this element |
| `w:rsidRDefault` | paragraph | Session that last set default run props for this para |
| `w:rsidRPr` | paragraph, run | Session that last modified run/formatting properties |
| `w:rsidDel` | paragraph, run | Session that deleted this element |
| `w:rsidP` | paragraph | Session that last modified paragraph properties |
| `w:rsidTr` | table row | Session that last modified this row |

**RSID generation:** Use any 8-hex-digit value. For programmatic edits, pick a single
RSID for the entire edit session (e.g., `"00FF0001"`) and also register it in
`settings.xml`:

```xml
<w:rsids>
  <w:rsidRoot w:val="00A00001"/>
  <w:rsid w:val="00A00001"/>
  <w:rsid w:val="00FF0001"/>  <!-- Add your session RSID here -->
</w:rsids>
```

#### 2.2.9 Revision ID Generation

Tracked change IDs (`w:id` on `<w:ins>`, `<w:del>`, `<w:rPrChange>`, etc.) must be
unique across the ENTIRE document (all parts). To generate safe IDs:

1. Parse ALL XML parts: `document.xml`, `comments.xml`, `footnotes.xml`,
   `endnotes.xml`, `header*.xml`, `footer*.xml`.
2. Find all elements with `w:id` attributes that are revision markers.
3. Take the maximum value found.
4. Start new IDs at `max + 1`, incrementing for each new revision element.

**Warning:** Do NOT use random IDs. Word may renumber them, but some consumers of
`.docx` files depend on ID ordering.

#### 2.2.10 All 28 Revision Element Types

The full set of revision tracking elements in OOXML:

| Element | Description |
|---------|-------------|
| `w:ins` | Insertion of content (runs, paragraphs, rows, cells) |
| `w:del` | Deletion of content |
| `w:moveFrom` | Source of a moved block |
| `w:moveTo` | Destination of a moved block |
| `w:rPrChange` | Run property change |
| `w:pPrChange` | Paragraph property change |
| `w:tblPrChange` | Table property change |
| `w:tblGridChange` | Table grid change |
| `w:trPrChange` | Table row property change |
| `w:tcPrChange` | Table cell property change |
| `w:sectPrChange` | Section property change |
| `w:tblPrExChange` | Table-level exception property change |
| `w:customXmlInsRangeStart` | Start of custom XML insertion range |
| `w:customXmlInsRangeEnd` | End of custom XML insertion range |
| `w:customXmlDelRangeStart` | Start of custom XML deletion range |
| `w:customXmlDelRangeEnd` | End of custom XML deletion range |
| `w:customXmlMoveFromRangeStart` | Start of custom XML move-from range |
| `w:customXmlMoveFromRangeEnd` | End of custom XML move-from range |
| `w:customXmlMoveToRangeStart` | Start of custom XML move-to range |
| `w:customXmlMoveToRangeEnd` | End of custom XML move-to range |
| `w:moveFromRangeStart` | Start marker for move-from range |
| `w:moveFromRangeEnd` | End marker for move-from range |
| `w:moveToRangeStart` | Start marker for move-to range |
| `w:moveToRangeEnd` | End marker for move-to range |
| `w:numberingChange` | Numbering change |
| `w:cellIns` | Table cell insertion |
| `w:cellDel` | Table cell deletion |
| `w:cellMerge` | Table cell merge change |

All revision elements share the common attributes: `w:id`, `w:author`, `w:date`.

---

### 2.3 Comments

Comments in OOXML span four separate XML files, each serving a distinct purpose.

#### 2.3.1 `comments.xml` - Comment Bodies

Each comment is a block-level container (like a mini-document):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

  <!-- A top-level comment -->
  <w:comment w:id="100" w:author="Reviewer 1" w:date="2026-03-20T09:00:00Z"
             w:initials="R1">
    <w:p w14:paraId="C0000001" w14:textId="77777777">
      <w:pPr>
        <w:pStyle w:val="CommentText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:annotationRef/>
      </w:r>
      <w:r>
        <w:t xml:space="preserve">This sentence is unclear. Could you rephrase?</w:t>
      </w:r>
    </w:p>
  </w:comment>

  <!-- A reply to the above comment -->
  <w:comment w:id="101" w:author="Fabio Votta" w:date="2026-03-27T14:00:00Z"
             w:initials="FV">
    <w:p w14:paraId="C0000002" w14:textId="77777777">
      <w:pPr>
        <w:pStyle w:val="CommentText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:annotationRef/>
      </w:r>
      <w:r>
        <w:t xml:space="preserve">Done. Rephrased for clarity.</w:t>
      </w:r>
    </w:p>
  </w:comment>

</w:comments>
```

**Attributes on `<w:comment>`:**

| Attribute | Required | Description |
|-----------|----------|-------------|
| `w:id` | Yes | Unique comment ID (integer). Referenced by `commentRangeStart`, `commentRangeEnd`, `commentReference`. |
| `w:author` | Yes | Author display name. |
| `w:date` | Yes | ISO 8601 timestamp. |
| `w:initials` | No | Author initials (shown in margin bubble). |

**Inside the comment body:**

- Each `<w:p>` MUST have a unique `w14:paraId` (used for threading in commentsExtended.xml).
- The first run conventionally contains `<w:annotationRef/>` with CommentReference style.
- Comments can contain multiple paragraphs, formatting, even images (via
  `word/_rels/comments.xml.rels`).

#### 2.3.2 Comment Anchors in `document.xml`

Comments are anchored to text ranges in the document body using three elements:

```xml
<w:p>
  <w:r>
    <w:t xml:space="preserve">This is </w:t>
  </w:r>

  <!-- Comment range start -->
  <w:commentRangeStart w:id="100"/>

  <w:r>
    <w:t>the commented text</w:t>
  </w:r>

  <!-- Comment range end -->
  <w:commentRangeEnd w:id="100"/>

  <!-- Comment reference (creates the superscript marker) -->
  <w:r>
    <w:rPr>
      <w:rStyle w:val="CommentReference"/>
    </w:rPr>
    <w:commentReference w:id="100"/>
  </w:r>

  <w:r>
    <w:t xml:space="preserve"> and this is after.</w:t>
  </w:r>
</w:p>
```

**Rules:**

- `w:commentRangeStart` and `w:commentRangeEnd` use the same `w:id` as the
  `<w:comment>` in `comments.xml`.
- `w:commentReference` also uses the same `w:id`.
- `commentRangeStart` and `commentRangeEnd` can span multiple paragraphs.
- `commentReference` should appear in a run immediately after `commentRangeEnd`,
  with the `CommentReference` run style.
- For a reply, you do NOT add `commentRangeStart`/`commentRangeEnd`/`commentReference`
  in document.xml. Replies are linked only through `commentsExtended.xml`.

#### 2.3.3 `commentsExtended.xml` - Threading (w15)

This file links replies to their parent comments using `paraIdParent`. It also
controls the "done" (resolved) state.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="w15">

  <!-- Top-level comment (no paraIdParent) -->
  <w15:commentEx
    w15:paraId="C0000001"
    w15:done="0"/>

  <!-- Reply to the above (linked via paraIdParent) -->
  <w15:commentEx
    w15:paraId="C0000002"
    w15:paraIdParent="C0000001"
    w15:done="0"/>

</w15:commentsEx>
```

**Attributes on `<w15:commentEx>`:**

| Attribute | Required | Description |
|-----------|----------|-------------|
| `w15:paraId` | Yes | The `w14:paraId` of the LAST paragraph in the comment body (from `comments.xml`). This is the linking key. |
| `w15:paraIdParent` | No | The `w15:paraId` of the PARENT comment. Omit for top-level comments. |
| `w15:done` | No | `"0"` = open, `"1"` = resolved/done. |

**How to create a reply:**

1. Add the reply `<w:comment>` to `comments.xml` with a new unique `w:id`.
2. Give its paragraph a unique `w14:paraId`.
3. Add a `<w15:commentEx>` in `commentsExtended.xml` with:
   - `w15:paraId` = the reply's paragraph `w14:paraId`
   - `w15:paraIdParent` = the parent comment's paragraph `w14:paraId`
4. Do NOT add `commentRangeStart`/`commentRangeEnd`/`commentReference` in
   `document.xml` for the reply.

**How to resolve a comment:**

Set `w15:done="1"` on the top-level `<w15:commentEx>` entry. All replies in the
thread are considered resolved when the root is resolved.

#### 2.3.4 `commentsIds.xml` - Durable IDs (w16cid)

Provides persistent identifiers that survive round-tripping through different editors.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w16cid:commentsIds
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="w16cid">

  <w16cid:commentId w16cid:paraId="C0000001" w16cid:durableId="1A2B3C4D"/>
  <w16cid:commentId w16cid:paraId="C0000002" w16cid:durableId="5E6F7A8B"/>

</w16cid:commentsIds>
```

**Attributes:**

| Attribute | Description |
|-----------|-------------|
| `w16cid:paraId` | Links to the comment paragraph's `w14:paraId` |
| `w16cid:durableId` | A stable 8-hex-digit ID that persists across saves. Used for co-authoring. |

Generate `durableId` as any unique 8-hex-digit value. Scan existing values to avoid
collisions.

#### 2.3.5 `commentsExtensible.xml` - Extended Comment Data (w16cex)

Used by newer versions of Word (2019+) for additional comment metadata. This file is
optional but increasingly common.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w16cex:commentsExtensible
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="w16cex">

  <w16cex:comment w16cex:durableId="1A2B3C4D"
                  w16cex:dateUtc="2026-03-20T09:00:00Z"/>
  <w16cex:comment w16cex:durableId="5E6F7A8B"
                  w16cex:dateUtc="2026-03-27T14:00:00Z"/>

</w16cex:commentsExtensible>
```

This file stores UTC timestamps and other metadata that may differ from the local
time stored in `comments.xml`.

#### 2.3.6 Complete Comment Setup Checklist

When adding a comment to a `.docx`, you must touch these files:

| File | What to add |
|------|-------------|
| `word/comments.xml` | `<w:comment>` element with body paragraphs |
| `word/document.xml` | `<w:commentRangeStart>`, `<w:commentRangeEnd>`, `<w:commentReference>` (top-level only, not replies) |
| `word/commentsExtended.xml` | `<w15:commentEx>` for threading and done state |
| `word/commentsIds.xml` | `<w16cid:commentId>` for durable ID |
| `word/commentsExtensible.xml` | `<w16cex:comment>` (optional, for Word 2019+) |
| `[Content_Types].xml` | `<Override>` entries for any new comment parts |
| `word/_rels/document.xml.rels` | `<Relationship>` entries for any new comment parts |

---

### 2.4 Images and Figures

#### 2.4.1 Image Architecture

Images in OOXML involve multiple layers:

1. **The binary image file** in `word/media/` (e.g., `image1.png`)
2. **A relationship** in `word/_rels/document.xml.rels` linking an `rId` to the media file
3. **A `<w:drawing>` element** in the document body containing the DrawingML markup
4. **A content type** in `[Content_Types].xml` for the image format

#### 2.4.2 Inline Image Structure

```xml
<w:r>
  <w:rPr>
    <w:noProof/>  <!-- Suppresses spell-check underlines on the image -->
  </w:rPr>
  <w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0"
               wp14:anchorId="3A4B5C6D" wp14:editId="7E8F9A0B">

      <!-- Image dimensions in EMUs -->
      <wp:extent cx="5486400" cy="3200400"/>
      <!-- 5486400 EMU = 6 inches wide, 3200400 EMU = 3.5 inches tall -->

      <!-- Effective extent (for cropped images) -->
      <wp:effectExtent l="0" t="0" r="0" b="0"/>

      <!-- Document properties (for numbering, accessibility) -->
      <wp:docPr id="1" name="Figure 1" descr="A bar chart showing election results"/>

      <!-- Non-visual graphic frame properties -->
      <wp:cNvGraphicFramePr>
        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                             noChangeAspect="1"/>
      </wp:cNvGraphicFramePr>

      <!-- The actual graphic -->
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">

            <!-- Non-visual properties -->
            <pic:nvPicPr>
              <pic:cNvPr id="1" name="image1.png"/>
              <pic:cNvPicPr/>
            </pic:nvPicPr>

            <!-- Blip fill (the actual image reference) -->
            <pic:blipFill>
              <a:blip r:embed="rId20"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <!-- Optional: compression state -->
                <a:extLst>
                  <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                    <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                                     val="0"/>
                  </a:ext>
                </a:extLst>
              </a:blip>
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </pic:blipFill>

            <!-- Shape properties (size, position, borders) -->
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="5486400" cy="3200400"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </pic:spPr>

          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>
```

#### 2.4.3 Key Image Attributes

**`<wp:inline>` attributes:**

| Attribute | Description |
|-----------|-------------|
| `distT`, `distB`, `distL`, `distR` | Distance from text (in EMUs). Usually `0` for inline. |
| `wp14:anchorId` | Unique 8-hex-digit anchor ID. |
| `wp14:editId` | Unique 8-hex-digit edit tracking ID. |

**`<wp:extent>` attributes:**

| Attribute | Description |
|-----------|-------------|
| `cx` | Width in EMUs |
| `cy` | Height in EMUs |

**`<wp:docPr>` attributes:**

| Attribute | Description |
|-----------|-------------|
| `id` | Unique integer ID across all drawing objects. |
| `name` | Display name (e.g., "Figure 1"). |
| `descr` | Alt text for accessibility. |

**`<a:blip>` attributes:**

| Attribute | Description |
|-----------|-------------|
| `r:embed` | Relationship ID pointing to the image in `document.xml.rels`. |
| `r:link` | Relationship ID for a linked (not embedded) image. Use one or the other. |
| `cstate` | Compression state: `print`, `screen`, `email`, or omit for none. |

#### 2.4.4 EMU Calculations

EMU (English Metric Unit) is the universal unit for drawings in OOXML.

| Measurement | EMUs |
|-------------|------|
| 1 inch | 914400 |
| 1 cm | 360000 |
| 1 mm | 36000 |
| 1 point | 12700 |
| 1 pixel (96 DPI) | 9525 |

**Common conversions:**

```
EMU = inches * 914400
EMU = cm * 360000
EMU = pixels_at_96dpi * 9525
```

**Example:** A 1920x1080 image at 96 DPI displayed at full size:
- cx = 1920 * 9525 = 18,288,000 EMU (20 inches -- too wide, would need scaling)
- To fit 6 inches wide: cx = 6 * 914400 = 5,486,400 EMU
- Maintain aspect ratio: cy = (1080/1920) * 5,486,400 = 3,086,100 EMU

#### 2.4.5 Floating Images (Anchored)

For images that float beside text, use `<wp:anchor>` instead of `<wp:inline>`:

```xml
<w:drawing>
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
             simplePos="0" relativeHeight="251659264"
             behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"
             wp14:anchorId="4A5B6C7D" wp14:editId="8E9F0A1B">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column">
      <wp:align>center</wp:align>
    </wp:positionH>
    <wp:positionV relativeFrom="paragraph">
      <wp:posOffset>0</wp:posOffset>
    </wp:positionV>
    <wp:extent cx="5486400" cy="3200400"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapTopAndBottom/>  <!-- or wrapSquare, wrapTight, wrapNone, wrapThrough -->
    <wp:docPr id="2" name="Figure 2" descr="Description"/>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                           noChangeAspect="1"/>
    </wp:cNvGraphicFramePr>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <!-- Same graphicData/pic:pic structure as inline -->
      ...
    </a:graphic>
  </wp:anchor>
</w:drawing>
```

**Wrapping modes:**

| Element | Behavior |
|---------|----------|
| `<wp:wrapNone/>` | No text wrapping (image floats over text) |
| `<wp:wrapSquare wrapText="bothSides"/>` | Text wraps in rectangle around image |
| `<wp:wrapTight wrapText="bothSides"/>` | Text follows image contour |
| `<wp:wrapThrough wrapText="bothSides"/>` | Text wraps through transparent areas |
| `<wp:wrapTopAndBottom/>` | No text beside image (like a block figure) |

#### 2.4.6 Image Relationship Setup

In `word/_rels/document.xml.rels`:

```xml
<Relationship Id="rId20"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

In `[Content_Types].xml` (if not already present):

```xml
<Default Extension="png" ContentType="image/png"/>
```

---

### 2.5 Tables

#### 2.5.1 Basic Table Structure

```xml
<w:tbl>
  <!-- Table properties -->
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="0" w:type="auto"/>          <!-- Auto width -->
    <w:tblW w:w="5000" w:type="pct"/>         <!-- 100% width (50ths of percent) -->
    <w:tblW w:w="9360" w:type="dxa"/>         <!-- Exact width in twips (6.5 inches) -->
    <w:jc w:val="center"/>                     <!-- Table alignment -->
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/>
      <w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/>
      <w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>
    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"
               w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>

  <!-- Column grid (defines column widths) -->
  <w:tblGrid>
    <w:gridCol w:w="3120"/>  <!-- Column 1: 3120 twips -->
    <w:gridCol w:w="3120"/>  <!-- Column 2: 3120 twips -->
    <w:gridCol w:w="3120"/>  <!-- Column 3: 3120 twips -->
  </w:tblGrid>

  <!-- Header row -->
  <w:tr w:rsidR="00A00001" w14:paraId="T0000001" w14:textId="77777777">
    <w:trPr>
      <w:tblHeader/>  <!-- Repeat this row as header on each page -->
    </w:trPr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Variable</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Coefficient</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Std. Error</w:t></w:r>
      </w:p>
    </w:tc>
  </w:tr>

  <!-- Data row -->
  <w:tr w:rsidR="00A00001">
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Intercept</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>0.543</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>(0.102)</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

#### 2.5.2 Table Property Details

**`<w:tblW>` width types:**

| w:type | Meaning | Example |
|--------|---------|---------|
| `"auto"` | Automatic width | `w:w="0"` |
| `"dxa"` | Width in twips (1/20 point) | `w:w="9360"` (6.5 inches) |
| `"pct"` | Width in 50ths of a percent | `w:w="5000"` (100%) |
| `"nil"` | Zero width | `w:w="0"` |

**Border `w:val` styles:**

| Value | Description |
|-------|-------------|
| `none` | No border |
| `single` | Single line |
| `thick` | Thick line |
| `double` | Double line |
| `dotted` | Dotted line |
| `dashed` | Dashed line |
| `dashSmallGap` | Dashes with small gaps |
| `thinThickSmallGap` | Thin-thick compound |
| `thickThinSmallGap` | Thick-thin compound |
| `threeDEmboss` | 3D embossed |
| `threeDEngrave` | 3D engraved |

**Border `w:sz` values:** In eighths of a point. `w:sz="4"` = 0.5pt, `w:sz="8"` = 1pt,
`w:sz="12"` = 1.5pt, `w:sz="16"` = 2pt.

#### 2.5.3 Cell Properties

```xml
<w:tcPr>
  <!-- Cell width -->
  <w:tcW w:w="3120" w:type="dxa"/>

  <!-- Cell borders (override table borders) -->
  <w:tcBorders>
    <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    <w:left w:val="nil"/>
    <w:right w:val="nil"/>
  </w:tcBorders>

  <!-- Cell shading/background -->
  <w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>

  <!-- Vertical alignment -->
  <w:vAlign w:val="center"/>  <!-- top, center, bottom -->

  <!-- Column span -->
  <w:gridSpan w:val="2"/>  <!-- Span 2 columns -->

  <!-- Vertical merge -->
  <w:vMerge w:val="restart"/>  <!-- Start of vertical merge -->
  <w:vMerge/>                  <!-- Continue vertical merge (omit val) -->

  <!-- Text direction -->
  <w:textDirection w:val="btLr"/>  <!-- Bottom to top, left to right (rotated) -->

  <!-- No wrap -->
  <w:noWrap/>

  <!-- Cell margins (override table defaults) -->
  <w:tcMar>
    <w:top w:w="72" w:type="dxa"/>
    <w:left w:w="115" w:type="dxa"/>
    <w:bottom w:w="72" w:type="dxa"/>
    <w:right w:w="115" w:type="dxa"/>
  </w:tcMar>
</w:tcPr>
```

**Important:** Every `<w:tc>` MUST contain at least one `<w:p>`. A cell with no
paragraph is invalid and will corrupt the document.

#### 2.5.4 Booktabs-Style Academic Tables

Academic tables use horizontal rules only (no vertical lines), with thicker rules at
the top and bottom:

```xml
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/>
    <w:jc w:val="center"/>
    <!-- NO table-level borders (we set them per-cell) -->
    <w:tblBorders>
      <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="40" w:type="dxa"/>
      <w:left w:w="80" w:type="dxa"/>
      <w:bottom w:w="40" w:type="dxa"/>
      <w:right w:w="80" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>

  <w:tblGrid>
    <w:gridCol w:w="3120"/>
    <w:gridCol w:w="3120"/>
    <w:gridCol w:w="3120"/>
  </w:tblGrid>

  <!-- HEADER ROW: toprule (thick top) + midrule (thin bottom) -->
  <w:tr>
    <w:trPr><w:tblHeader/></w:trPr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>    <!-- toprule: 1.5pt -->
          <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>  <!-- midrule: 0.75pt -->
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Variable</w:t></w:r>
      </w:p>
    </w:tc>
    <!-- Repeat for other header cells with same borders -->
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Estimate</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Std. Error</w:t></w:r>
      </w:p>
    </w:tc>
  </w:tr>

  <!-- DATA ROWS: no borders -->
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="nil"/><w:bottom w:val="nil"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:r><w:t>Intercept</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="nil"/><w:bottom w:val="nil"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>2.341***</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="nil"/><w:bottom w:val="nil"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>(0.456)</w:t></w:r></w:p>
    </w:tc>
  </w:tr>

  <!-- LAST DATA ROW: bottomrule (thick bottom) -->
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="nil"/>
          <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>  <!-- bottomrule -->
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:r><w:t>Treatment</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="nil"/>
          <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>0.789*</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="nil"/>
          <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>(0.321)</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

**Booktabs rules in OOXML:**

| LaTeX | OOXML | Border Size |
|-------|-------|-------------|
| `\toprule` | Top border on header cells | `w:sz="12"` (1.5pt) |
| `\midrule` | Bottom border on header cells | `w:sz="6"` (0.75pt) |
| `\bottomrule` | Bottom border on last row cells | `w:sz="12"` (1.5pt) |
| `\cmidrule` | Bottom border on specific header cells only | `w:sz="4"` (0.5pt) |

---

### 2.6 Styles

#### 2.6.1 Style References

In a paragraph:

```xml
<w:pPr>
  <w:pStyle w:val="Heading1"/>
</w:pPr>
```

In a run:

```xml
<w:rPr>
  <w:rStyle w:val="Strong"/>
</w:rPr>
```

The `w:val` refers to the `w:styleId` attribute in `styles.xml`, NOT the display name.

#### 2.6.2 `styles.xml` Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="w14">

  <!-- Document defaults -->
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="CMU Serif" w:eastAsia="CMU Serif"
                  w:hAnsi="CMU Serif" w:cs="Times New Roman"/>
        <w:sz w:val="24"/>     <!-- Default 12pt -->
        <w:szCs w:val="24"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <!-- Latent style defaults (controls which built-in styles are shown in UI) -->
  <w:latentStyles w:defLockedState="0" w:defUIPriority="99"
                  w:defSemiHidden="0" w:defUnhideWhenUsed="0"
                  w:defQFormat="0" w:count="376">
    <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
    <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 2" w:uiPriority="9" w:semiHidden="1"
                    w:unhideWhenUsed="1" w:qFormat="1"/>
    <!-- ... more exceptions ... -->
  </w:latentStyles>

  <!-- Normal style (base for most paragraphs) -->
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="both"/>  <!-- Justified -->
      <w:ind w:firstLine="360"/>  <!-- Paragraph indent -->
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="CMU Serif" w:hAnsi="CMU Serif"/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>

  <!-- Heading 1 -->
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>      <!-- Inherits from Normal -->
    <w:next w:val="Normal"/>          <!-- Next paragraph reverts to Normal -->
    <w:link w:val="Heading1Char"/>    <!-- Linked character style -->
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="480" w:after="120"/>
      <w:ind w:firstLine="0"/>        <!-- Override: no indent for headings -->
      <w:jc w:val="left"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="36"/>             <!-- 18pt -->
      <w:szCs w:val="36"/>
    </w:rPr>
  </w:style>

  <!-- Heading 2 -->
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading2Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="360" w:after="80"/>
      <w:ind w:firstLine="0"/>
      <w:jc w:val="left"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="28"/>             <!-- 14pt -->
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>

  <!-- Character style example -->
  <w:style w:type="character" w:styleId="Strong">
    <w:name w:val="Strong"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="22"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
  </w:style>

  <!-- Comment styles -->
  <w:style w:type="paragraph" w:styleId="CommentText">
    <w:name w:val="annotation text"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="CommentTextChar"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
      <w:spacing w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val="20"/>  <!-- 10pt for comments -->
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>

  <w:style w:type="character" w:styleId="CommentReference">
    <w:name w:val="annotation reference"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:sz w:val="16"/>  <!-- 8pt -->
      <w:szCs w:val="16"/>
    </w:rPr>
  </w:style>

</w:styles>
```

#### 2.6.3 Style Types

| w:type | Description | Referenced By |
|--------|-------------|---------------|
| `"paragraph"` | Paragraph style (controls both paragraph and default run properties) | `<w:pStyle w:val="..."/>` |
| `"character"` | Character/run style (controls run properties only) | `<w:rStyle w:val="..."/>` |
| `"table"` | Table style | `<w:tblStyle w:val="..."/>` |
| `"numbering"` | Numbering/list style | `<w:numStyleLink>` / `<w:styleLink>` |

#### 2.6.4 Style Inheritance

```
docDefaults
  +-- Normal (w:type="paragraph", w:default="1")
        +-- Heading1 (basedOn="Normal")
        +-- Heading2 (basedOn="Normal")
        +-- BodyText (basedOn="Normal")
              +-- BodyTextIndent (basedOn="BodyText")
```

Resolution order (later overrides earlier):

1. `<w:docDefaults>` -- applies to everything
2. Table style (if inside a table)
3. Paragraph style (from `<w:pStyle>`)
4. Character style (from `<w:rStyle>`)
5. Direct formatting (explicit properties in `<w:pPr>` or `<w:rPr>`)

**Toggle properties** (bold, italic, caps, etc.) interact specially: if a style sets
`<w:b/>` and direct formatting also sets `<w:b/>`, they cancel out (toggle off). To
explicitly force bold on, use `<w:b w:val="1"/>` (though `<w:b/>` is the same as
`<w:b w:val="1"/>` -- it is the style interaction that creates the toggle).

---

## 3. Namespaces

### 3.1 Complete Namespace Reference

| Prefix | URI | Used For |
|--------|-----|----------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | Core WordprocessingML elements (paragraphs, runs, tables, styles, tracked changes) |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship references (`r:id`, `r:embed`) |
| `wp` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | Word drawing wrapper (`wp:inline`, `wp:anchor`, `wp:extent`) |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML core (shapes, fills, effects) |
| `pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` | Picture elements (`pic:pic`, `pic:blipFill`) |
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | Markup compatibility (`mc:Ignorable`, `mc:AlternateContent`) |
| `o` | `urn:schemas-microsoft-com:office:office` | Legacy Office elements |
| `v` | `urn:schemas-microsoft-com:vml` | VML (Vector Markup Language) -- legacy drawing |
| `m` | `http://schemas.openxmlformats.org/officeDocument/2006/math` | Office Math (equations) |
| `w10` | `urn:schemas-microsoft-com:office:word` | Legacy Word-specific elements |
| `wne` | `http://schemas.microsoft.com/office/word/2006/wordml` | Word 2006 non-essential extensions |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` | Word 2010 extensions (`w14:paraId`, `w14:textId`, content controls, etc.) |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` | Word 2013 extensions (comment threading: `w15:commentEx`) |
| `w16` | `http://schemas.microsoft.com/office/word/2018/wordml` | Word 2019 general extensions |
| `w16cid` | `http://schemas.microsoft.com/office/word/2016/wordml/cid` | Comment durable IDs (`w16cid:commentId`) |
| `w16cex` | `http://schemas.microsoft.com/office/word/2018/wordml/cex` | Comment extensible data (`w16cex:commentsExtensible`) |
| `w16se` | `http://schemas.microsoft.com/office/word/2015/wordml/symex` | Symbol extensions |
| `w16sdtdh` | `http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash` | Structured document tag data hash |
| `wp14` | `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing` | Word 2010 drawing extensions (`wp14:anchorId`, `wp14:editId`) |
| `wpc` | `http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas` | Word processing canvas |
| `wpg` | `http://schemas.microsoft.com/office/word/2010/wordprocessingGroup` | Word processing group |
| `wps` | `http://schemas.microsoft.com/office/word/2010/wordprocessingShape` | Word processing shape |
| `wpi` | `http://schemas.microsoft.com/office/word/2010/wordprocessingInk` | Ink annotations |
| `cx` | `http://schemas.microsoft.com/office/drawing/2014/chartex` | Chart extensions |
| `a14` | `http://schemas.microsoft.com/office/drawing/2010/main` | DrawingML 2010 extensions |
| `dgm` | `http://schemas.openxmlformats.org/drawingml/2006/diagram` | SmartArt diagrams |
| `c` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | Charts |
| `c16r3` | `http://schemas.microsoft.com/office/drawing/2017/03/chart` | Chart 2017 extensions |

### 3.2 Markup Compatibility (`mc:`)

The `mc:Ignorable` attribute on the root element lists namespace prefixes that older
consumers can safely ignore:

```xml
<w:document ... mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14">
```

This means: if a consumer does not understand `w14:paraId`, it can skip it without
error.

`mc:AlternateContent` provides fallback content:

```xml
<mc:AlternateContent>
  <mc:Choice Requires="w14">
    <!-- Modern content using w14 features -->
  </mc:Choice>
  <mc:Fallback>
    <!-- Fallback for older consumers -->
  </mc:Fallback>
</mc:AlternateContent>
```

---

## 4. Common Patterns

### 4.1 Simple Paragraph with Bold Text

```xml
<w:p w14:paraId="A1B2C3D4" w14:textId="E5F6A7B8"
     w:rsidR="00AA0001" w:rsidRDefault="00AA0001"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:pPr>
    <w:pStyle w:val="Normal"/>
    <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
    <w:jc w:val="both"/>
  </w:pPr>
  <w:r w:rsidRPr="00AA0001">
    <w:t xml:space="preserve">This is normal text with </w:t>
  </w:r>
  <w:r w:rsidRPr="00AA0001">
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
    <w:t>bold emphasis</w:t>
  </w:r>
  <w:r w:rsidRPr="00AA0001">
    <w:t xml:space="preserve"> in the middle.</w:t>
  </w:r>
</w:p>
```

### 4.2 Tracked Replacement (Delete Old + Insert New)

This is the most common tracked change: replacing one word or phrase with another.

```xml
<w:p w14:paraId="B2C3D4E5" w14:textId="F6A7B8C9"
     w:rsidR="00AA0001" w:rsidRDefault="00BB0002"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:pPr>
    <w:pStyle w:val="Normal"/>
  </w:pPr>

  <!-- Unchanged text before the edit -->
  <w:r w:rsidR="00AA0001">
    <w:t xml:space="preserve">The results were </w:t>
  </w:r>

  <!-- DELETION: old text -->
  <w:del w:id="10" w:author="Fabio Votta" w:date="2026-03-27T15:00:00Z">
    <w:r w:rsidR="00AA0001" w:rsidDel="00BB0002">
      <w:delText xml:space="preserve">quite significant</w:delText>
    </w:r>
  </w:del>

  <!-- INSERTION: new text -->
  <w:ins w:id="11" w:author="Fabio Votta" w:date="2026-03-27T15:00:00Z">
    <w:r w:rsidR="00BB0002">
      <w:t xml:space="preserve">statistically significant (p &lt; 0.001)</w:t>
    </w:r>
  </w:ins>

  <!-- Unchanged text after the edit -->
  <w:r w:rsidR="00AA0001">
    <w:t xml:space="preserve"> across all conditions.</w:t>
  </w:r>
</w:p>
```

### 4.3 Tracked Paragraph Insertion

Insert an entirely new paragraph between existing ones, with change tracking:

```xml
<!-- Existing paragraph -->
<w:p w14:paraId="C3D4E5F6" w14:textId="A7B8C9D0"
     w:rsidR="00AA0001" w:rsidRDefault="00AA0001"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:r>
    <w:t>This paragraph existed before.</w:t>
  </w:r>
</w:p>

<!-- NEW paragraph, tracked as insertion -->
<w:ins w:id="12" w:author="Fabio Votta" w:date="2026-03-27T15:00:00Z"
       xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p w14:paraId="D4E5F6A7" w14:textId="B8C9D0E1"
       w:rsidR="00BB0002" w:rsidRDefault="00BB0002"
       xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
    <w:pPr>
      <w:pStyle w:val="Normal"/>
      <w:rPr>
        <w:rFonts w:ascii="CMU Serif" w:hAnsi="CMU Serif"/>
        <w:sz w:val="24"/>
      </w:rPr>
    </w:pPr>
    <w:r w:rsidR="00BB0002">
      <w:rPr>
        <w:rFonts w:ascii="CMU Serif" w:hAnsi="CMU Serif"/>
        <w:sz w:val="24"/>
      </w:rPr>
      <w:t>This entire paragraph was added during revision.</w:t>
    </w:r>
  </w:p>
</w:ins>

<!-- Next existing paragraph -->
<w:p w14:paraId="E5F6A7B8" w14:textId="C9D0E1F2"
     w:rsidR="00AA0001" w:rsidRDefault="00AA0001"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:r>
    <w:t>This paragraph also existed before.</w:t>
  </w:r>
</w:p>
```

### 4.4 Adding a Comment with a Reply

**Step 1: In `comments.xml`, add both the comment and the reply:**

```xml
<!-- Top-level comment -->
<w:comment w:id="200" w:author="Reviewer 2" w:date="2026-03-20T10:00:00Z" w:initials="R2"
           xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p w14:paraId="CA000001" w14:textId="77777777">
    <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
    <w:r>
      <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
      <w:annotationRef/>
    </w:r>
    <w:r>
      <w:t xml:space="preserve">Please add a citation for this claim.</w:t>
    </w:r>
  </w:p>
</w:comment>

<!-- Reply -->
<w:comment w:id="201" w:author="Fabio Votta" w:date="2026-03-27T14:30:00Z" w:initials="FV"
           xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p w14:paraId="CA000002" w14:textId="77777777">
    <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
    <w:r>
      <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
      <w:annotationRef/>
    </w:r>
    <w:r>
      <w:t xml:space="preserve">Added: Votta et al. (2025).</w:t>
    </w:r>
  </w:p>
</w:comment>
```

**Step 2: In `document.xml`, anchor only the top-level comment (NOT the reply):**

```xml
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:r>
    <w:t xml:space="preserve">Social media platforms have become </w:t>
  </w:r>
  <w:commentRangeStart w:id="200"/>
  <w:r>
    <w:t xml:space="preserve">the primary arena for political advertising</w:t>
  </w:r>
  <w:commentRangeEnd w:id="200"/>
  <w:r>
    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
    <w:commentReference w:id="200"/>
  </w:r>
  <w:r>
    <w:t>.</w:t>
  </w:r>
</w:p>
```

**Step 3: In `commentsExtended.xml`, link the reply to its parent:**

```xml
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                mc:Ignorable="w15">
  <w15:commentEx w15:paraId="CA000001" w15:done="0"/>
  <w15:commentEx w15:paraId="CA000002" w15:paraIdParent="CA000001" w15:done="0"/>
</w15:commentsEx>
```

**Step 4: In `commentsIds.xml`, add durable IDs:**

```xml
<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                    mc:Ignorable="w16cid">
  <w16cid:commentId w16cid:paraId="CA000001" w16cid:durableId="D0000001"/>
  <w16cid:commentId w16cid:paraId="CA000002" w16cid:durableId="D0000002"/>
</w16cid:commentsIds>
```

### 4.5 Inserting an Inline Image

**Step 1: Place the image file at `word/media/image1.png`.**

**Step 2: Add a relationship in `word/_rels/document.xml.rels`:**

```xml
<Relationship Id="rId50"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

**Step 3: Add a content type in `[Content_Types].xml` (if not already present):**

```xml
<Default Extension="png" ContentType="image/png"/>
```

**Step 4: Add the drawing element in `document.xml`:**

```xml
<w:p w14:paraId="F1A2B3C4" w14:textId="D5E6F7A8"
     w:rsidR="00CC0003" w:rsidRDefault="00CC0003"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:pPr>
    <w:jc w:val="center"/>  <!-- Center the image -->
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:noProof/>
    </w:rPr>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0"
                 wp14:anchorId="5A6B7C8D" wp14:editId="9E0F1A2B">
        <wp:extent cx="5486400" cy="3657600"/>  <!-- 6" x 4" -->
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="1" name="Figure 1"
                  descr="Bar chart showing ad spend by party"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="1" name="image1.png"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="rId50"/>
                <a:stretch>
                  <a:fillRect/>
                </a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="5486400" cy="3657600"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                  <a:avLst/>
                </a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
```

### 4.6 A Three-Column Academic Table with Booktabs Borders

See section 2.5.4 for the complete booktabs table example.

### 4.7 Headings (Heading 1, 2, 3)

```xml
<!-- Heading 1 -->
<w:p w14:paraId="H1000001" w14:textId="77777777"
     w:rsidR="00AA0001" w:rsidRDefault="00AA0001"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
  </w:pPr>
  <w:bookmarkStart w:id="0" w:name="_Toc_Section1"/>
  <w:r>
    <w:t>1. Introduction</w:t>
  </w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>

<!-- Heading 2 -->
<w:p w14:paraId="H2000001" w14:textId="77777777"
     w:rsidR="00AA0001" w:rsidRDefault="00AA0001"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:pPr>
    <w:pStyle w:val="Heading2"/>
  </w:pPr>
  <w:r>
    <w:t>1.1 Background</w:t>
  </w:r>
</w:p>

<!-- Heading 3 -->
<w:p w14:paraId="H3000001" w14:textId="77777777"
     w:rsidR="00AA0001" w:rsidRDefault="00AA0001"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:pPr>
    <w:pStyle w:val="Heading3"/>
  </w:pPr>
  <w:r>
    <w:t>1.1.1 Political Advertising Online</w:t>
  </w:r>
</w:p>
```

### 4.8 Footnote

**In `document.xml`, add the reference:**

```xml
<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:t xml:space="preserve">significant effects on voter behavior</w:t>
</w:r>
<w:r>
  <w:rPr>
    <w:rStyle w:val="FootnoteReference"/>
  </w:rPr>
  <w:footnoteReference w:id="1"/>
</w:r>
```

**In `word/footnotes.xml`, add the footnote body:**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- Separator footnotes (required, always present) -->
  <w:footnote w:type="separator" w:id="-1">
    <w:p w14:paraId="FN000001" w14:textId="77777777">
      <w:r><w:separator/></w:r>
    </w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p w14:paraId="FN000002" w14:textId="77777777">
      <w:r><w:continuationSeparator/></w:r>
    </w:p>
  </w:footnote>

  <!-- Actual footnote -->
  <w:footnote w:id="1">
    <w:p w14:paraId="FN000003" w14:textId="77777777">
      <w:pPr>
        <w:pStyle w:val="FootnoteText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="FootnoteReference"/>
        </w:rPr>
        <w:footnoteRef/>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> See Votta et al. (2025) for a detailed analysis.</w:t>
      </w:r>
    </w:p>
  </w:footnote>
</w:footnotes>
```

### 4.9 Hyperlink

```xml
<w:hyperlink r:id="rId30" w:history="1"
             xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:r w:rsidRPr="00DD0004">
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>click here</w:t>
  </w:r>
</w:hyperlink>
```

The `r:id` points to a relationship in `document.xml.rels` with
`TargetMode="External"`:

```xml
<Relationship Id="rId30"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="https://example.com"
              TargetMode="External"/>
```

### 4.10 Page Break

```xml
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:r>
    <w:br w:type="page"/>
  </w:r>
</w:p>
```

Or via paragraph property (break before the paragraph):

```xml
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:pageBreakBefore/>
  </w:pPr>
  <w:r>
    <w:t>This paragraph starts on a new page.</w:t>
  </w:r>
</w:p>
```

### 4.11 Numbered and Bulleted Lists

Lists in OOXML use abstract numbering definitions in `numbering.xml`, referenced by
paragraphs:

**In `word/numbering.xml`:**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Abstract definition (template) -->
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <!-- Level 0: bullets -->
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>  <!-- bullet character -->
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <!-- Level 1: en-dash -->
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2013;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>

  <!-- Concrete numbering instance -->
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>

  <!-- Abstract definition for numbered list -->
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>

  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>

</w:numbering>
```

**In `document.xml`, reference the list:**

```xml
<!-- Bullet list item -->
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:pStyle w:val="ListParagraph"/>
    <w:numPr>
      <w:ilvl w:val="0"/>
      <w:numId w:val="1"/>  <!-- bullet list -->
    </w:numPr>
  </w:pPr>
  <w:r>
    <w:t>First bullet point</w:t>
  </w:r>
</w:p>

<!-- Numbered list item -->
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:pStyle w:val="ListParagraph"/>
    <w:numPr>
      <w:ilvl w:val="0"/>
      <w:numId w:val="2"/>  <!-- numbered list -->
    </w:numPr>
  </w:pPr>
  <w:r>
    <w:t>First numbered item</w:t>
  </w:r>
</w:p>
```

**Number format values (`w:numFmt`):**

| Value | Output |
|-------|--------|
| `decimal` | 1, 2, 3 |
| `lowerLetter` | a, b, c |
| `upperLetter` | A, B, C |
| `lowerRoman` | i, ii, iii |
| `upperRoman` | I, II, III |
| `bullet` | Bullet character (from lvlText) |
| `none` | No number |

---

## 5. Pitfalls and Gotchas

### 5.1 Run Splitting

**Problem:** Word splits text across multiple `<w:r>` elements unpredictably.

**Impact:** String searching in XML will miss matches that span runs.

**Solution:** When searching for text in a paragraph:
1. Concatenate all `<w:t>` and `<w:delText>` content across runs.
2. Track character offsets back to specific runs.
3. When replacing, handle the case where the match spans run boundaries:
   - Truncate the first run's text before the match start.
   - Remove intermediate runs entirely.
   - Truncate the last run's text after the match end.
   - Insert new text in a new run (or modify an existing one).

### 5.2 ID Collisions

**Three separate ID spaces exist in OOXML:**

| ID Type | Scope | Format | Used By |
|---------|-------|--------|---------|
| Revision IDs (`w:id`) | All revision elements across all parts | Integer | `<w:ins>`, `<w:del>`, `<w:rPrChange>`, `<w:pPrChange>`, etc. |
| Comment IDs (`w:id`) | All comments | Integer | `<w:comment>`, `<w:commentRangeStart>`, `<w:commentRangeEnd>`, `<w:commentReference>` |
| Relationship IDs (`r:id`) | Per `.rels` file | String (`rIdN`) | `<Relationship>`, `r:embed`, `r:id` |
| Paragraph IDs (`w14:paraId`) | All parts | 8 hex digits | `<w:p>`, `<w15:commentEx>`, `<w16cid:commentId>` |
| Drawing IDs (`wp:docPr id`) | All drawing objects | Integer | `<wp:docPr>` |
| Bookmark IDs (`w:id`) | Bookmark start/end pairs | Integer | `<w:bookmarkStart>`, `<w:bookmarkEnd>` |

**Note:** Comment IDs and revision IDs share the `w:id` attribute name but are
technically in the same ID space in practice -- Word may renumber them. To be safe,
keep all `w:id` values across ALL elements unique.

**Safe generation algorithm:**

```
1. Parse all XML parts
2. Collect all w:id values from all elements
3. maxId = max(collected values)
4. nextId = maxId + 1
5. For each new element needing an ID: assign nextId, increment nextId
```

### 5.3 `xml:space="preserve"`

**Problem:** Without `xml:space="preserve"` on `<w:t>`, XML parsers strip leading
and trailing whitespace.

**Impact:** `<w:t> Hello </w:t>` becomes `"Hello"` (no spaces).

**Solution:** ALWAYS add `xml:space="preserve"`:

```xml
<w:t xml:space="preserve"> Hello </w:t>
```

**When it matters most:**

- Text starting or ending with spaces
- Text that is entirely whitespace (e.g., `" "` between styled runs)
- Text after a deletion/insertion boundary where spacing is critical

**When you can omit it:** When the text has no leading or trailing whitespace. But
it is always safe to include it, so prefer including it unconditionally.

### 5.4 ZIP Requirements

**Problem:** Invalid ZIP construction corrupts the `.docx`.

**Rules:**

- Use deflate compression (standard ZIP).
- Do NOT use ZIP64 extensions unless the file exceeds 4GB.
- File paths in the ZIP must use forward slashes (`word/document.xml`, not
  `word\document.xml`).
- `[Content_Types].xml` MUST be at the root of the ZIP (not in a subdirectory).
- The `_rels/` directory must be at the root.
- Part names are case-sensitive in the ZIP but case-insensitive in relationship
  targets. Use consistent casing.
- When modifying an existing `.docx`, preserve the compression method and level of
  each entry. Recompressing with different settings can change file size significantly
  and may cause validation warnings.

### 5.5 Paragraph Count Must Not Decrease

**Problem:** After editing a `.docx`, if the paragraph count is lower than before,
something was accidentally deleted.

**Validation rule:** After any modification:

```
count(<w:p>) in new document >= count(<w:p>) in original document
```

Exceptions: intentional paragraph merging (which should be tracked via `<w:del>` on
the paragraph mark). Even then, the `<w:p>` elements remain -- they just have a
deletion mark on the paragraph break.

### 5.6 EMU Calculations

**Problem:** Wrong EMU values cause images to display at incorrect sizes.

**Quick reference:**

```
EMU per inch:       914,400
EMU per cm:         360,000
EMU per mm:          36,000
EMU per point:       12,700
EMU per pixel@96dpi:  9,525
EMU per pixel@72dpi: 12,700
```

**Common page width (for fitting images):**

- US Letter with 1-inch margins: 6.5 inches = 5,943,600 EMU max image width
- A4 with 2.54cm margins: 16.51cm text width = 5,943,600 EMU max image width

### 5.7 Content Types for New Media

**Problem:** Adding an image without a corresponding content type entry makes Word
ignore or fail to render it.

**Checklist when adding media:**

1. Check if `[Content_Types].xml` already has a `<Default>` for the file extension.
2. If not, add one:
   ```xml
   <Default Extension="png" ContentType="image/png"/>
   ```
3. Common media types:

| Extension | ContentType |
|-----------|-------------|
| png | `image/png` |
| jpg, jpeg | `image/jpeg` |
| gif | `image/gif` |
| svg | `image/svg+xml` |
| emf | `image/x-emf` |
| wmf | `image/x-wmf` |
| tiff, tif | `image/tiff` |

### 5.8 Relationship Must-Haves for New Parts

When adding a new part to the document:

1. **Add a `<Relationship>` in the appropriate `.rels` file** -- usually
   `word/_rels/document.xml.rels` for parts referenced by `document.xml`.
2. **Add a content type** in `[Content_Types].xml` -- usually an `<Override>`.
3. **Reference the relationship ID** correctly from the XML that uses the part.

Missing any of these three steps will cause the part to be invisible or the
document to be corrupted.

### 5.9 Element Ordering in `<w:pPr>`

The children of `<w:pPr>` MUST appear in a specific order defined by the XSD schema.
The correct order is:

1. `<w:pStyle>`
2. `<w:keepNext>`
3. `<w:keepLines>`
4. `<w:pageBreakBefore>`
5. `<w:framePr>`
6. `<w:widowControl>`
7. `<w:numPr>`
8. `<w:suppressLineNumbers>`
9. `<w:pBdr>`
10. `<w:shd>`
11. `<w:tabs>`
12. `<w:suppressAutoHyphens>`
13. `<w:kinsoku>`
14. `<w:wordWrap>`
15. `<w:overflowPunct>`
16. `<w:topLinePunct>`
17. `<w:autoSpaceDE>`
18. `<w:autoSpaceDN>`
19. `<w:bidi>`
20. `<w:adjustRightInd>`
21. `<w:snapToGrid>`
22. `<w:spacing>`
23. `<w:ind>`
24. `<w:contextualSpacing>`
25. `<w:mirrorIndents>`
26. `<w:suppressOverlap>`
27. `<w:jc>`
28. `<w:textDirection>`
29. `<w:textAlignment>`
30. `<w:textboxTightWrap>`
31. `<w:outlineLvl>`
32. `<w:divId>`
33. `<w:cnfStyle>`
34. `<w:rPr>` (MUST be last)
35. `<w:sectPr>` (only on last paragraph in body)
36. `<w:pPrChange>` (if tracked change, comes after rPr)

**Tip:** Most validators are lenient, but strict XML Schema validation (and some
editors like OnlyOffice) will reject out-of-order elements. Always follow the order.

### 5.10 Element Ordering in `<w:rPr>`

Similarly, `<w:rPr>` children have a prescribed order:

1. `<w:rStyle>`
2. `<w:rFonts>`
3. `<w:b>`, `<w:bCs>`
4. `<w:i>`, `<w:iCs>`
5. `<w:caps>`, `<w:smallCaps>`
6. `<w:strike>`, `<w:dstrike>`
7. `<w:outline>`
8. `<w:shadow>`
9. `<w:emboss>`
10. `<w:imprint>`
11. `<w:noProof>`
12. `<w:snapToGrid>`
13. `<w:vanish>`
14. `<w:webHidden>`
15. `<w:color>`
16. `<w:spacing>`
17. `<w:w>` (character width scaling)
18. `<w:kern>`
19. `<w:position>`
20. `<w:sz>`, `<w:szCs>`
21. `<w:highlight>`
22. `<w:u>`
23. `<w:effect>`
24. `<w:bdr>`
25. `<w:shd>`
26. `<w:fitText>`
27. `<w:vertAlign>`
28. `<w:rtl>`
29. `<w:cs>`
30. `<w:em>`
31. `<w:lang>`
32. `<w:eastAsianLayout>`
33. `<w:specVanish>`
34. `<w:oMath>`
35. `<w:rPrChange>` (MUST be last if present)

### 5.11 The `mc:Ignorable` Trap

If you add elements from extension namespaces (like `w14:paraId`, `w15:commentEx`),
make sure those namespace prefixes are listed in `mc:Ignorable` on the root element.
Otherwise, older consumers that do not understand those namespaces will throw errors
instead of gracefully skipping them.

### 5.12 Empty Paragraphs and Tables

- An empty `<w:p/>` (self-closing) is valid and creates a blank line.
- An empty `<w:p></w:p>` is also valid.
- A `<w:tc>` MUST contain at least one `<w:p>`. A cell with no paragraphs corrupts
  the document.
- A `<w:tr>` MUST contain at least one `<w:tc>`.
- A `<w:tbl>` MUST contain `<w:tblGrid>` and at least one `<w:tr>`.

### 5.13 Date Format for Tracked Changes and Comments

All dates MUST use ISO 8601 format with timezone:

```
2026-03-27T14:30:00Z
```

- The `T` separator is required.
- Use `Z` for UTC or `+HH:MM`/`-HH:MM` for offset.
- Word displays this in the user's local timezone.
- Do NOT omit the time component -- Word may reject it.

### 5.14 Maximum File Sizes and Performance

| Content | Practical Limit | Notes |
|---------|----------------|-------|
| Total paragraphs | ~100,000 | Word slows significantly above this |
| Images | ~500MB total | Word loads all into memory |
| Single image | ~20MB | Larger images cause slow rendering |
| Comments | ~1,000 | Threading UI degrades above this |
| Tracked changes | ~10,000 | Accept/reject becomes very slow |
| Table rows | ~10,000 | Pagination becomes slow |

---

## 6. The LaTeX Equivalence Table

| LaTeX | OOXML | Notes |
|-------|-------|-------|
| `\documentclass{article}` | `word/document.xml` + `word/styles.xml` | Styles define the "class" |
| `\begin{document}...\end{document}` | `<w:body>...</w:body>` | |
| `\section{Title}` | `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>` | Heading styles set `outlineLvl` |
| `\subsection{Title}` | `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>...` | `outlineLvl="1"` |
| `\subsubsection{Title}` | `<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr>...` | `outlineLvl="2"` |
| `\textbf{bold}` | `<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>` | |
| `\textit{italic}` | `<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>` | |
| `\underline{text}` | `<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>text</w:t></w:r>` | |
| `\textsc{Small Caps}` | `<w:r><w:rPr><w:smallCaps/></w:rPr><w:t>Small Caps</w:t></w:r>` | |
| `\texttt{monospace}` | `<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>monospace</w:t></w:r>` | |
| `\emph{emphasis}` | `<w:r><w:rPr><w:i/></w:rPr><w:t>emphasis</w:t></w:r>` | No toggle behavior in OOXML |
| `{\color{red}text}` | `<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>text</w:t></w:r>` | 6-digit hex, no `#` |
| `\footnote{text}` | `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>` + `footnotes.xml` | |
| `\href{url}{text}` | `<w:hyperlink r:id="rIdN"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>text</w:t></w:r></w:hyperlink>` | + relationship |
| `\includegraphics{img}` | `<w:drawing><wp:inline>...<a:blip r:embed="rIdN"/>...</wp:inline></w:drawing>` | See section 4.5 |
| `\begin{itemize}` | `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` (bullet) | Requires `numbering.xml` |
| `\begin{enumerate}` | `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>` (decimal) | Requires `numbering.xml` |
| `\begin{tabular}` | `<w:tbl>...<w:tblGrid>...<w:tr><w:tc>...` | See section 2.5 |
| `\toprule` | `<w:top w:val="single" w:sz="12" ... />` on header cells | 1.5pt rule |
| `\midrule` | `<w:bottom w:val="single" w:sz="6" ... />` on header cells | 0.75pt rule |
| `\bottomrule` | `<w:bottom w:val="single" w:sz="12" ... />` on last row | 1.5pt rule |
| `\cmidrule{2-3}` | Per-cell bottom border on cols 2--3 of header | Selective borders |
| `\centering` | `<w:jc w:val="center"/>` | In `<w:pPr>` |
| `\raggedright` | `<w:jc w:val="left"/>` | |
| `\raggedleft` | `<w:jc w:val="right"/>` | |
| `\justify` (default) | `<w:jc w:val="both"/>` | |
| `\newpage` | `<w:r><w:br w:type="page"/></w:r>` | Or `<w:pageBreakBefore/>` in pPr |
| `\\` (line break) | `<w:r><w:br/></w:r>` | No `w:type` = line break |
| `\hspace{1cm}` | `<w:r><w:tab/></w:r>` with tab stop | Or use spacing |
| `\vspace{12pt}` | `<w:spacing w:before="240"/>` or `<w:spacing w:after="240"/>` | In twips (240 = 12pt) |
| `\setlength{\parindent}{0.5in}` | `<w:ind w:firstLine="720"/>` | In style or direct |
| `\setlength{\parskip}{6pt}` | `<w:spacing w:after="120"/>` | 120 twips = 6pt |
| `\linespread{1.5}` | `<w:spacing w:line="360" w:lineRule="auto"/>` | See spacing table |
| `\fontsize{12}{14.4}\selectfont` | `<w:sz w:val="24"/>` | Half-points (24 = 12pt) |
| `\textsubscript{n}` | `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>` | |
| `\textsuperscript{n}` | `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>` | |
| `\sout{strikethrough}` | `<w:rPr><w:strike/></w:rPr>` | |
| `\hl{highlight}` | `<w:rPr><w:highlight w:val="yellow"/></w:rPr>` | Preset colors only |
| `\label{sec:intro}` | `<w:bookmarkStart w:id="0" w:name="sec_intro"/>...<w:bookmarkEnd w:id="0"/>` | |
| `\ref{sec:intro}` | `<w:fldChar w:fldCharType="begin"/>...<w:instrText> REF sec_intro \\h </w:instrText>...<w:fldChar w:fldCharType="end"/>` | Field codes |
| `\tableofcontents` | `<w:fldChar w:fldCharType="begin"/>...<w:instrText> TOC \\o "1-3" \\h \\z \\u </w:instrText>...` | Field code |
| `\maketitle` | Manual paragraphs with Title/Subtitle styles | No built-in equivalent |

---

## Appendix A: Minimal Valid .docx File List

To create the smallest possible valid `.docx`:

```
[Content_Types].xml
_rels/.rels
word/document.xml
word/_rels/document.xml.rels
```

**`[Content_Types].xml`:**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
```

**`_rels/.rels`:**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
</Relationships>
```

**`word/document.xml`:**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
```

**`word/_rels/document.xml.rels`:**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
```

---

## Appendix B: Unit Conversion Quick Reference

| From | To | Formula |
|------|----|---------|
| Inches | Twips | `inches * 1440` |
| Points | Twips | `points * 20` |
| Inches | EMUs | `inches * 914400` |
| cm | EMUs | `cm * 360000` |
| Points | Half-points | `points * 2` |
| Pixels (96 DPI) | EMUs | `pixels * 9525` |
| Pixels (72 DPI) | EMUs | `pixels * 12700` |
| Twips | EMUs | `twips * 635` |
| EMUs | Inches | `EMU / 914400` |
| EMUs | cm | `EMU / 360000` |
| Half-points | Points | `halfpoints / 2` |
| Border sz (eighths) | Points | `sz / 8` |
| Percentage (w:w pct) | Percent | `pct / 50` |

---

## Appendix C: Commonly Needed Content Type URIs

| Part | Content Type |
|------|-------------|
| Main document | `application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml` |
| Styles | `application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml` |
| Settings | `application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml` |
| Font table | `application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml` |
| Numbering | `application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml` |
| Footnotes | `application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml` |
| Endnotes | `application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml` |
| Comments | `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml` |
| Comments extended | `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml` |
| Comments IDs | `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml` |
| Comments extensible | `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml` |
| Header | `application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml` |
| Footer | `application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml` |
| Theme | `application/vnd.openxmlformats-officedocument.theme+xml` |
| Core properties | `application/vnd.openxmlformats-package.core-properties+xml` |
| Extended properties | `application/vnd.openxmlformats-officedocument.extended-properties+xml` |
| Custom properties | `application/vnd.openxmlformats-officedocument.custom-properties+xml` |
| Relationships | `application/vnd.openxmlformats-package.relationships+xml` |
| Template (dotx) main | `application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml` |
| Macro-enabled (docm) main | `application/vnd.ms-word.document.macroEnabled.main+xml` |

---

## Appendix D: Field Codes Reference

Field codes in OOXML use a begin/separate/end structure:

```xml
<w:r>
  <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
  <w:instrText xml:space="preserve"> PAGE </w:instrText>
</w:r>
<w:r>
  <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
  <w:t>1</w:t>  <!-- Cached result -->
</w:r>
<w:r>
  <w:fldChar w:fldCharType="end"/>
</w:r>
```

**Common field codes:**

| Field | Instruction | Description |
|-------|-------------|-------------|
| Page number | `PAGE` | Current page number |
| Total pages | `NUMPAGES` | Total page count |
| Date | `DATE \@ "yyyy-MM-dd"` | Current date with format |
| TOC | `TOC \o "1-3" \h \z \u` | Table of contents (levels 1-3, hyperlinks) |
| Cross-reference | `REF bookmark_name \h` | Reference to bookmark |
| Bibliography | `BIBLIOGRAPHY` | Zotero/Mendeley bibliography |
| Citation | `ADDIN ZOTERO_ITEM CSL_CITATION {...}` | Zotero citation |
| Merge field | `MERGEFIELD FieldName` | Mail merge field |
| IF | `IF expression "true" "false"` | Conditional |
| SEQ | `SEQ Figure \* ARABIC` | Sequence numbering (for figure/table numbers) |

---

*End of OOXML Reference*
