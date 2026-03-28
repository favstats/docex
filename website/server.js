#!/usr/bin/env node
/**
 * docex demo website server
 *
 * Endpoints:
 *   POST /api/decompile  - Upload .docx, returns { dex, html }
 *   POST /api/build      - Send .dex string, returns .docx file download
 *   POST /api/preview    - Send .dex string, returns { html }
 *   GET  /               - Single-page app
 */

'use strict';

const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');

// docex modules from parent directory
const docex = require('../src/docex');
const { DexDecompiler } = require('../src/dex-decompiler');
const { DexCompiler } = require('../src/dex-compiler');
const { Workspace } = require('../src/workspace');

const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));

// Parse JSON bodies (for /api/preview and /api/build)
app.use(express.json({ limit: '50mb' }));

// --------------------------------------------------------------------------
// POST /api/decompile -- upload a .docx, get back { dex, html }
// --------------------------------------------------------------------------
app.post('/api/decompile', (req, res) => {
  try {
    // Read raw body as binary
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      try {
        const buf = Buffer.concat(chunks);

        // Write to temp file
        const tmpPath = path.join(os.tmpdir(), `docex-upload-${crypto.randomBytes(8).toString('hex')}.docx`);
        fs.writeFileSync(tmpPath, buf);

        // Decompile to .dex using workspace
        const ws = Workspace.open(tmpPath);
        const dexContent = DexDecompiler.decompile(ws);
        ws.cleanup();

        // Clean up temp file
        try { fs.unlinkSync(tmpPath); } catch (_) {}

        // Generate HTML preview
        const html = dexToHtml(dexContent);

        res.json({ dex: dexContent, html });
      } catch (err) {
        res.status(500).json({ error: err.message });
      }
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// --------------------------------------------------------------------------
// POST /api/build -- send { dex: "..." }, get back .docx download
// --------------------------------------------------------------------------
app.post('/api/build', (req, res) => {
  try {
    const { dex } = req.body;
    if (!dex) return res.status(400).json({ error: 'Missing dex field' });

    const result = DexCompiler.compile(dex);
    const docxPath = result.path;

    if (!docxPath || !fs.existsSync(docxPath)) {
      return res.status(500).json({ error: 'Build failed: no output file' });
    }

    const buf = fs.readFileSync(docxPath);
    try { fs.unlinkSync(docxPath); } catch (_) {}

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="document.docx"');
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// --------------------------------------------------------------------------
// POST /api/preview -- send { dex: "..." }, get back { html }
// --------------------------------------------------------------------------
app.post('/api/preview', (req, res) => {
  try {
    const { dex } = req.body;
    if (!dex) return res.status(400).json({ error: 'Missing dex field' });

    const html = dexToHtml(dex);
    res.json({ html });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// --------------------------------------------------------------------------
// GET /api/example -- returns the example .dex content
// --------------------------------------------------------------------------
app.get('/api/example', (req, res) => {
  try {
    const examplePath = path.join(__dirname, '..', 'examples', 'absolute_chaos.dex');
    const content = fs.readFileSync(examplePath, 'utf-8');
    // Return first ~50 lines
    const lines = content.split('\n').slice(0, 50);
    res.json({ dex: lines.join('\n') });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});


// ==========================================================================
// DEX -> HTML CONVERTER (server-side, used by /api/decompile and /api/preview)
// ==========================================================================

function escHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function dexToHtml(dex) {
  const lines = dex.split('\n');
  const parts = [];
  let inFrontmatter = false;
  let inComment = false;
  let commentMeta = '';
  let commentBody = [];
  let inParagraph = false;
  let paragraphContent = [];

  function flushParagraph() {
    if (paragraphContent.length > 0) {
      const text = paragraphContent.join('\n');
      parts.push('<p class="dex-para">' + renderInline(text) + '</p>');
      paragraphContent = [];
    }
    inParagraph = false;
  }

  function flushComment() {
    if (commentBody.length > 0) {
      parts.push('<div class="dex-comment"><span class="dex-comment-author">' + escHtml(commentMeta) + '</span>' + escHtml(commentBody.join('\n')) + '</div>');
      commentBody = [];
    }
    inComment = false;
  }

  function renderInline(text) {
    // Process inline .dex markup to HTML
    let html = escHtml(text);

    // Bold: {b}...{/b}
    html = html.replace(/\{b\}([\s\S]*?)\{\/b\}/g, '<strong>$1</strong>');
    // Italic: {i}...{/i}
    html = html.replace(/\{i\}([\s\S]*?)\{\/i\}/g, '<em>$1</em>');
    // Underline: {u}...{/u}
    html = html.replace(/\{u\}([\s\S]*?)\{\/u\}/g, '<u>$1</u>');
    // Strikethrough: {strike}...{/strike}
    html = html.replace(/\{strike\}([\s\S]*?)\{\/strike\}/g, '<s>$1</s>');
    // Superscript: {sup}...{/sup}
    html = html.replace(/\{sup\}([\s\S]*?)\{\/sup\}/g, '<sup>$1</sup>');
    // Subscript: {sub}...{/sub}
    html = html.replace(/\{sub\}([\s\S]*?)\{\/sub\}/g, '<sub>$1</sub>');

    // Deletions: {del ...}...{/del}
    html = html.replace(/\{del[^}]*\}([\s\S]*?)\{\/del\}/g, '<span class="dex-del">$1</span>');
    // Insertions: {ins ...}...{/ins}
    html = html.replace(/\{ins[^}]*\}([\s\S]*?)\{\/ins\}/g, '<span class="dex-ins">$1</span>');

    // Font tags (strip, keep content)
    html = html.replace(/\{font[^}]*\}/g, '');
    html = html.replace(/\{\/font\}/g, '');
    // Color tags (strip, keep content)
    html = html.replace(/\{color[^}]*\}/g, '');
    html = html.replace(/\{\/color\}/g, '');
    // Highlight tags (strip, keep content)
    html = html.replace(/\{highlight[^}]*\}/g, '');
    html = html.replace(/\{\/highlight\}/g, '');

    // Footnotes: {footnote id:N}...{/footnote}
    html = html.replace(/\{footnote id:(\d+)\}([\s\S]*?)\{\/footnote\}/g,
      '<span class="dex-footnote" title="$2"><sup>[$1]</sup></span>');
    // Empty footnote refs
    html = html.replace(/\{footnote id:(\d+)\}\{\/footnote\}/g,
      '<span class="dex-footnote"><sup>[$1]</sup></span>');

    return html;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    // Frontmatter
    if (i === 0 && trimmed === '---') {
      inFrontmatter = true;
      continue;
    }
    if (inFrontmatter) {
      if (trimmed === '---') {
        inFrontmatter = false;
      }
      continue;
    }

    // Comment start
    if (trimmed.startsWith('{comment ')) {
      flushParagraph();
      const byMatch = trimmed.match(/by:"([^"]+)"/);
      commentMeta = byMatch ? byMatch[1] : 'Comment';
      inComment = true;
      // Check if single-line comment
      if (trimmed.endsWith('{/comment}')) {
        const inner = trimmed.replace(/\{comment[^}]*\}/, '').replace(/\{\/comment\}/, '');
        parts.push('<div class="dex-comment"><span class="dex-comment-author">' + escHtml(commentMeta) + '</span>' + escHtml(inner) + '</div>');
        inComment = false;
      }
      continue;
    }
    if (trimmed === '{/comment}') {
      flushComment();
      continue;
    }
    if (inComment) {
      commentBody.push(line);
      continue;
    }

    // Paragraph start
    if (trimmed.startsWith('{p') && trimmed.endsWith('}') && !trimmed.startsWith('{pagebreak}')) {
      flushParagraph();
      inParagraph = true;
      continue;
    }
    if (trimmed === '{/p}') {
      flushParagraph();
      continue;
    }
    if (inParagraph) {
      paragraphContent.push(line);
      continue;
    }

    // Pagebreak
    if (trimmed === '{pagebreak}') {
      flushParagraph();
      parts.push('<hr class="dex-pagebreak">');
      continue;
    }

    // Headings
    const headingMatch = trimmed.match(/^(#{1,6})\s+(.+?)(?:\s*\{id:[A-F0-9]+\})?\s*$/);
    if (headingMatch) {
      flushParagraph();
      const level = headingMatch[1].length;
      const text = headingMatch[2];
      parts.push('<h' + level + '>' + escHtml(text) + '</h' + level + '>');
      continue;
    }

    // Table (basic detection)
    if (trimmed.startsWith('{table')) {
      flushParagraph();
      parts.push('<div class="dex-table-placeholder">[Table]</div>');
      continue;
    }

    // Figure
    if (trimmed.startsWith('{figure')) {
      flushParagraph();
      const captionMatch = trimmed.match(/caption:"([^"]+)"/);
      const caption = captionMatch ? captionMatch[1] : 'Figure';
      parts.push('<div class="dex-figure-placeholder">[Figure: ' + escHtml(caption) + ']</div>');
      continue;
    }

    // Empty line
    if (trimmed === '') {
      continue;
    }

    // Plain text line (not inside {p})
    parts.push('<p class="dex-para">' + renderInline(trimmed) + '</p>');
  }

  flushParagraph();
  flushComment();

  return parts.join('\n');
}


// --------------------------------------------------------------------------
// Start server
// --------------------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`docex demo server running at http://localhost:${PORT}`);
});
