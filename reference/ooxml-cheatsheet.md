# OOXML Cheat Sheet for docex

> Copy-paste-ready XML snippets for the 5 most common .docx editing operations.
> Extracted from the full 2,967-line reference.

---

## Key Namespaces (top 5)

| Prefix | URI |
|--------|-----|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` |
| `wp` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` |

---

## 1. Tracked Changes: `w:ins` / `w:del`

### Replace text (delete old + insert new)

```xml
<!-- Preceding text unchanged -->
<w:r><w:t xml:space="preserve">The results were </w:t></w:r>

<!-- DELETION -->
<w:del w:id="10" w:author="Fabio Votta" w:date="2026-03-27T15:00:00Z">
  <w:r w:rsidDel="00BB0002">
    <w:delText xml:space="preserve">quite significant</w:delText>
  </w:r>
</w:del>

<!-- INSERTION -->
<w:ins w:id="11" w:author="Fabio Votta" w:date="2026-03-27T15:00:00Z">
  <w:r>
    <w:t xml:space="preserve">statistically significant (p &lt; 0.001)</w:t>
  </w:r>
</w:ins>
```

### Insert entire paragraph (tracked)

```xml
<w:ins w:id="12" w:author="Fabio Votta" w:date="2026-03-27T15:00:00Z">
  <w:p w14:paraId="D4E5F6A7" w14:textId="B8C9D0E1">
    <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
    <w:r><w:t>This entire paragraph was inserted.</w:t></w:r>
  </w:p>
</w:ins>
```

### Format change (tracked)

```xml
<w:rPr>
  <w:b/>
  <w:i/>
  <w:rPrChange w:id="3" w:author="Fabio Votta" w:date="2026-03-27T14:30:00Z">
    <w:rPr><w:b/></w:rPr>  <!-- old: only bold, no italic -->
  </w:rPrChange>
</w:rPr>
```

**Rules:**
- `w:id` must be unique across ALL revision elements in the document
- Inside `w:del`, text uses `<w:delText>` not `<w:t>`
- Always include `xml:space="preserve"` if text has leading/trailing spaces
- IDs should be sequential (max existing + 1), not random

---

## 2. Comments (all 5 files)

### File 1: `word/comments.xml`

```xml
<w:comment w:id="100" w:author="Reviewer 1" w:date="2026-03-20T09:00:00Z"
           w:initials="R1">
  <w:p w14:paraId="C0000001" w14:textId="77777777">
    <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
    <w:r>
      <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
      <w:annotationRef/>
    </w:r>
    <w:r><w:t xml:space="preserve">This sentence is unclear.</w:t></w:r>
  </w:p>
</w:comment>
```

### File 2: `word/document.xml` (anchors, top-level only -- NOT for replies)

```xml
<w:commentRangeStart w:id="100"/>
<w:r><w:t>the commented text</w:t></w:r>
<w:commentRangeEnd w:id="100"/>
<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="100"/>
</w:r>
```

### File 3: `word/commentsExtended.xml` (threading)

```xml
<!-- Top-level comment -->
<w15:commentEx w15:paraId="C0000001" w15:done="0"/>
<!-- Reply (linked via paraIdParent) -->
<w15:commentEx w15:paraId="C0000002" w15:paraIdParent="C0000001" w15:done="0"/>
```

### File 4: `word/commentsIds.xml` (durable IDs)

```xml
<w16cid:commentId w16cid:paraId="C0000001" w16cid:durableId="1A2B3C4D"/>
```

### File 5: `word/commentsExtensible.xml` (optional, Word 2019+)

```xml
<w16cex:comment w16cex:durableId="1A2B3C4D" w16cex:dateUtc="2026-03-20T09:00:00Z"/>
```

**Reply checklist:**
1. Add `<w:comment>` to comments.xml with new `w:id` and unique `w14:paraId`
2. Add `<w15:commentEx>` with `w15:paraIdParent` pointing to parent's paraId
3. Do NOT add commentRangeStart/End/Reference in document.xml for replies

---

## 3. Images: `w:drawing`

```xml
<w:r>
  <w:rPr><w:noProof/></w:rPr>
  <w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="5486400" cy="3200400"/>  <!-- EMU: 6in x 3.5in -->
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="1" name="Figure 1" descr="Alt text"/>
      <wp:cNvGraphicFramePr>
        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                             noChangeAspect="1"/>
      </wp:cNvGraphicFramePr>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:nvPicPr>
              <pic:cNvPr id="1" name="image1.png"/>
              <pic:cNvPicPr/>
            </pic:nvPicPr>
            <pic:blipFill>
              <a:blip r:embed="rId20"/>
              <a:stretch><a:fillRect/></a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="5486400" cy="3200400"/>
              </a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>
```

**Setup required:**
- `word/media/image1.png` -- the binary file
- `word/_rels/document.xml.rels` -- `<Relationship Id="rId20" Type="...image" Target="media/image1.png"/>`
- `[Content_Types].xml` -- `<Default Extension="png" ContentType="image/png"/>`

**EMU conversions:** 1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pixel @96dpi = 9525 EMU

---

## 4. Tables: `w:tbl` (booktabs style)

```xml
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/>  <!-- 100% width -->
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="3120"/>
    <w:gridCol w:w="3120"/>
  </w:tblGrid>

  <!-- HEADER: toprule (thick top) + midrule (thin bottom) -->
  <w:tr>
    <w:trPr><w:tblHeader/></w:trPr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Variable</w:t></w:r></w:p>
    </w:tc>
  </w:tr>

  <!-- DATA ROW: no borders -->
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
  </w:tr>

  <!-- LAST ROW: bottomrule (thick bottom) -->
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3120" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="nil"/>
          <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
      </w:tcPr>
      <w:p><w:r><w:t>Treatment</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

**Booktabs rule sizes:** toprule = `w:sz="12"` (1.5pt), midrule = `w:sz="6"` (0.75pt), bottomrule = `w:sz="12"` (1.5pt)

**Every `<w:tc>` MUST contain at least one `<w:p>`** or the document is corrupt.

---

## 5. Common Pitfalls

### Run splitting
Word splits "Hello World" into multiple `<w:r>` elements (spell check, revision tracking, formatting, bookmarks). When searching for text, concatenate ALL `<w:t>` values across runs in a paragraph. Never assume a word is in a single run.

### xml:space="preserve"
Without `xml:space="preserve"` on `<w:t>`, XML parsers strip leading/trailing whitespace. Always include it if text has spaces at boundaries.

### ID uniqueness
- `w:id` on tracked changes: unique across ALL XML parts (document, comments, footnotes, headers, footers). Use max+1, not random.
- `w14:paraId`: unique 8-hex-digit across ALL parts. Scan before generating.
- `wp:docPr id`: unique integer across all drawing objects.
- Relationship `Id` (rId): unique within each `.rels` file.

### w:delText vs w:t
Inside `<w:del>`, text MUST be `<w:delText>`, not `<w:t>`. Using `<w:t>` inside a deletion causes the text to display as normal (not deleted).

### Content Types
Every new XML part needs a corresponding `<Override>` in `[Content_Types].xml`. Missing content types cause Word to corrupt/ignore the file.

### Relationships
Every new image, comment file, or footnote file needs a `<Relationship>` entry in the appropriate `.rels` file. External hyperlinks require `TargetMode="External"`.
