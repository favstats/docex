/* ============================================================================
   docex -- shared JavaScript (syntax highlighting, .dex parsing, tabs)
   ============================================================================ */

// ---------- Tab switching ----------
function initTabs() {
  document.querySelectorAll('.tabs').forEach(tabGroup => {
    const buttons = tabGroup.querySelectorAll('.tab-btn');
    const parent = tabGroup.parentElement;
    const contents = parent.querySelectorAll('.tab-content');
    buttons.forEach(btn => {
      btn.addEventListener('click', () => {
        buttons.forEach(b => b.classList.remove('active'));
        contents.forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        const target = parent.querySelector('#' + btn.dataset.tab);
        if (target) target.classList.add('active');
      });
    });
  });
}

// ---------- Copy buttons ----------
function initCopyButtons() {
  document.querySelectorAll('pre').forEach(pre => {
    if (pre.querySelector('.copy-btn')) return;
    const btn = document.createElement('button');
    btn.className = 'copy-btn';
    btn.textContent = 'Copy';
    btn.addEventListener('click', () => {
      const code = pre.querySelector('code');
      const text = code ? code.textContent : pre.textContent;
      navigator.clipboard.writeText(text).then(() => {
        btn.textContent = 'Copied!';
        setTimeout(() => { btn.textContent = 'Copy'; }, 2000);
      });
    });
    pre.style.position = 'relative';
    pre.appendChild(btn);
  });
}

// ---------- Syntax highlighting for .dex code blocks ----------
function highlightDex(text) {
  // Escape HTML first
  let html = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  // Frontmatter
  html = html.replace(/^(---)$/gm, '<span class="dex-frontmatter">$1</span>');

  // Headings (lines starting with #)
  html = html.replace(/^(#{1,6}\s.*)$/gm, '<span class="dex-heading">$1</span>');

  // Comments and replies
  html = html.replace(/(\{comment\b[^}]*\})/g, '<span class="dex-comment">$1</span>');
  html = html.replace(/(\{\/comment\})/g, '<span class="dex-comment">$1</span>');
  html = html.replace(/(\{reply\b[^}]*\})/g, '<span class="dex-comment">$1</span>');
  html = html.replace(/(\{\/reply\})/g, '<span class="dex-comment">$1</span>');

  // Tracked changes
  html = html.replace(/(\{del\b[^}]*\})([\s\S]*?)(\{\/del\})/g,
    '<span class="dex-del">$1$2$3</span>');
  html = html.replace(/(\{ins\b[^}]*\})([\s\S]*?)(\{\/ins\})/g,
    '<span class="dex-ins">$1$2$3</span>');

  // Formatting tags
  html = html.replace(/(\{\/?(b|i|u|sub|sup)\})/g, '<span class="dex-tag">$1</span>');

  // Pagebreak
  html = html.replace(/(\{pagebreak\})/g, '<span class="dex-pagebreak">$1</span>');

  // Paragraph IDs
  html = html.replace(/(\{p\b[^}]*\})/g, '<span class="dex-id">$1</span>');
  html = html.replace(/(\{\/p\})/g, '<span class="dex-id">$1</span>');

  // {id:XXXX}
  html = html.replace(/(\{id:[A-Fa-f0-9]+\})/g, '<span class="dex-id">$1</span>');

  // Font/color/highlight
  html = html.replace(/(\{font\s+"[^"]*"\})/g, '<span class="dex-font">$1</span>');
  html = html.replace(/(\{\/font\})/g, '<span class="dex-font">$1</span>');
  html = html.replace(/(\{color\s+[A-Fa-f0-9]+\})/g, '<span class="dex-font">$1</span>');
  html = html.replace(/(\{\/color\})/g, '<span class="dex-font">$1</span>');
  html = html.replace(/(\{highlight\s+\w+\})/g, '<span class="dex-comment">$1</span>');
  html = html.replace(/(\{\/highlight\})/g, '<span class="dex-comment">$1</span>');

  // Figure/table tags
  html = html.replace(/(\{figure\b[^}]*\})/g, '<span class="dex-tag">$1</span>');
  html = html.replace(/(\{\/figure\})/g, '<span class="dex-tag">$1</span>');
  html = html.replace(/(\{table\b[^}]*\})/g, '<span class="dex-tag">$1</span>');
  html = html.replace(/(\{\/table\})/g, '<span class="dex-tag">$1</span>');

  // Footnotes
  html = html.replace(/(\{footnote\b[^}]*\})/g, '<span class="dex-font">$1</span>');
  html = html.replace(/(\{\/footnote\})/g, '<span class="dex-font">$1</span>');

  // Section
  html = html.replace(/(\{section\b[^}]*\})/g, '<span class="dex-id">$1</span>');

  return html;
}

// ---------- .dex to rendered HTML ----------
function renderDex(dexText) {
  const lines = dexText.split('\n');
  const html = [];
  let inFrontmatter = false;
  let frontmatterDone = false;
  let inComment = false;
  let commentMeta = {};
  let commentText = [];
  let inTable = false;
  let tableRows = [];
  let tableHeaderDone = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Frontmatter
    if (line === '---' && !frontmatterDone) {
      if (inFrontmatter) { inFrontmatter = false; frontmatterDone = true; }
      else { inFrontmatter = true; }
      continue;
    }
    if (inFrontmatter) continue;

    // Comments
    if (/^\{comment\b/.test(line)) {
      const byMatch = line.match(/by:"([^"]*)"/);
      commentMeta = { by: byMatch ? byMatch[1] : 'Unknown' };
      commentText = [];
      inComment = true;
      continue;
    }
    if (line === '{/comment}' && inComment) {
      html.push('<div class="preview-comment"><strong>' +
        escapeHtml(commentMeta.by) + ':</strong> ' +
        escapeHtml(commentText.join(' ')) + '</div>');
      inComment = false;
      continue;
    }
    if (/^\{reply\b/.test(line)) {
      const byMatch = line.match(/by:"([^"]*)"/);
      commentMeta = { by: byMatch ? byMatch[1] : 'Unknown' };
      commentText = [];
      inComment = true;
      continue;
    }
    if (line === '{/reply}' && inComment) {
      html.push('<div class="preview-comment" style="margin-left:48px"><strong>' +
        escapeHtml(commentMeta.by) + ' (reply):</strong> ' +
        escapeHtml(commentText.join(' ')) + '</div>');
      inComment = false;
      continue;
    }
    if (inComment) { commentText.push(line); continue; }

    // Table
    if (/^\{table\b/.test(line)) { inTable = true; tableRows = []; tableHeaderDone = false; continue; }
    if (line === '{/table}') {
      if (tableRows.length > 0) {
        let thtml = '<table>';
        thtml += '<tr>' + tableRows[0].map(c => '<th>' + escapeHtml(c) + '</th>').join('') + '</tr>';
        for (let r = 1; r < tableRows.length; r++) {
          thtml += '<tr>' + tableRows[r].map(c => '<td>' + escapeHtml(c) + '</td>').join('') + '</tr>';
        }
        thtml += '</table>';
        html.push(thtml);
      }
      inTable = false;
      continue;
    }
    if (inTable) {
      if (line.startsWith('|') && line.includes('---')) continue; // separator
      if (line.startsWith('|')) {
        const cells = line.split('|').filter(c => c.trim() !== '').map(c => c.trim());
        tableRows.push(cells);
      }
      continue;
    }

    // Pagebreak
    if (line === '{pagebreak}') { html.push('<hr class="pagebreak">'); continue; }

    // Headings
    const headingMatch = line.match(/^(#{1,6})\s+(.*)$/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      let text = headingMatch[2].replace(/\s*\{id:[A-Fa-f0-9]+\}/, '');
      html.push('<h' + level + '>' + escapeHtml(text) + '</h' + level + '>');
      continue;
    }

    // Paragraph start/end
    if (/^\{p\b/.test(line)) continue;
    if (line === '{/p}') continue;
    if (line.trim() === '') continue;

    // Figure
    if (/^\{figure\b/.test(line)) {
      html.push('<p style="text-align:center;color:#888;font-style:italic">[Figure]</p>');
      continue;
    }
    if (line === '{/figure}') continue;

    // Section
    if (/^\{section\b/.test(line)) continue;

    // Content lines - process inline formatting
    html.push('<p>' + renderInline(line) + '</p>');
  }

  return html.join('\n');
}

function renderInline(text) {
  // Strip {id:XXXX}
  let s = text.replace(/\{id:[A-Fa-f0-9]+\}/g, '');

  // Strip font/color/highlight tags (but keep content)
  s = s.replace(/\{font\s+"[^"]*"\}/g, '');
  s = s.replace(/\{\/font\}/g, '');
  s = s.replace(/\{color\s+[A-Fa-f0-9]+\}/g, '');
  s = s.replace(/\{\/color\}/g, '');
  s = s.replace(/\{highlight\s+\w+\}/g, '');
  s = s.replace(/\{\/highlight\}/g, '');

  // Footnotes
  s = s.replace(/\{footnote\b[^}]*\}([\s\S]*?)\{\/footnote\}/g, '<sup title="$1">[fn]</sup>');

  // Tracked changes
  s = s.replace(/\{del\b[^}]*\}([\s\S]*?)\{\/del\}/g, '<del>$1</del>');
  s = s.replace(/\{ins\b[^}]*\}([\s\S]*?)\{\/ins\}/g, '<ins>$1</ins>');

  // Formatting
  s = s.replace(/\{b\}/g, '<strong>');
  s = s.replace(/\{\/b\}/g, '</strong>');
  s = s.replace(/\{i\}/g, '<em>');
  s = s.replace(/\{\/i\}/g, '</em>');
  s = s.replace(/\{u\}/g, '<u>');
  s = s.replace(/\{\/u\}/g, '</u>');
  s = s.replace(/\{sup\}/g, '<sup>');
  s = s.replace(/\{\/sup\}/g, '</sup>');
  s = s.replace(/\{sub\}/g, '<sub>');
  s = s.replace(/\{\/sub\}/g, '</sub>');

  // Escape remaining braces display (but unescape \{ and \})
  s = s.replace(/\\{/g, '{');
  s = s.replace(/\\}/g, '}');

  return s;
}

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ---------- Active nav link ----------
function initActiveNav() {
  const path = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a').forEach(a => {
    const href = a.getAttribute('href');
    if (href === path || (path === '' && href === 'index.html') ||
        (path === 'index.html' && href === 'index.html')) {
      a.classList.add('active');
    }
  });
}

// ---------- Mobile nav toggle (for pages using shared script) ----------
function initMobileNav() {
  const toggle = document.querySelector('.nav-toggle');
  const navLinks = document.querySelector('.nav-links');
  if (toggle && navLinks) {
    toggle.addEventListener('click', function() {
      const isOpen = navLinks.classList.toggle('open');
      toggle.setAttribute('aria-expanded', String(isOpen));
    });
  }
}

// ---------- Scroll reveal (for pages using shared script) ----------
function initScrollReveal() {
  const reveals = document.querySelectorAll('.reveal');
  if (reveals.length === 0) return;
  const observer = new IntersectionObserver(function(entries) {
    entries.forEach(function(entry) {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
      }
    });
  }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });
  reveals.forEach(function(el) { observer.observe(el); });
}

// ---------- Init ----------
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initCopyButtons();
  initActiveNav();
  initMobileNav();
  initScrollReveal();
});
