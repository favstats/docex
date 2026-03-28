#!/usr/bin/env node
'use strict';

const path = require('path');

const { Workspace } = require('../src/workspace');
const { Comments } = require('../src/comments');
const { Footnotes } = require('../src/footnotes');
const { Figures } = require('../src/figures');
const { Tables } = require('../src/tables');
const { Lists } = require('../src/lists');
const { CrossRef } = require('../src/crossref');
const { Formatting } = require('../src/formatting');
const { Layout } = require('../src/layout');
const { Production } = require('../src/production');
const { Provenance } = require('../src/provenance');
const xml = require('../src/xml');

const INPUT = '/mnt/storage/absolute_chaos_v2.docx';
const OUTPUT = '/mnt/storage/absolute_chaos_v3.docx';
const IMAGE = path.resolve(__dirname, '../test/fixtures/test-image.png');

function main() {
  const ws = Workspace.open(INPUT);

  try {
    Production.coverPage(ws, {
      title: 'ABSOLUTE CHAOS v3',
      subtitle: 'For Round-Trip Verification Only',
      author: 'docex fixture forge',
      organization: 'Office of Compounding Bad Decisions',
      date: '2026-03-27',
    });
    Production.watermark(ws, 'PRE-RELEASE CONFUSION', { color: 'D9D9D9', size: 58, angle: -33 });
    Production.stamp(ws, 'CHAOS INDEX: 11/10', { position: 'footer', alignment: 'right' });

    Layout.pageBreakBefore(ws, 'EMERGENCY SECTION (added at 2am before the deadline)');
    Layout.ensureHeadingHierarchy(ws);

    Formatting.bold(ws, 'CONFIDENTIAL');
    Formatting.color(ws, 'CONFIDENTIAL', 'red');
    Formatting.highlight(ws, 'DRAFT', 'yellow');
    Formatting.code(ws, 'Version 47.3b');
    Formatting.underline(ws, 'REFERENCE (broken)');

    const budgetComment = Comments.add(
      ws,
      'Budget: $4.7 million.',
      'Need a source, a definition, and probably an exorcism.',
      { by: 'Chaos QA' }
    );
    Comments.reply(ws, budgetComment.commentId, 'Confirmed: the number is fictional but emotionally resonant.', {
      by: 'Spreadsheet Necromancer',
    });

    const confidentialComment = Comments.add(
      ws,
      'CONFIDENTIAL',
      'This label has now appeared in the title block, headers, footer stamp, and six email threads.',
      { by: 'Risk' }
    );
    Comments.reply(ws, confidentialComment.commentId, 'Please stop forwarding the confidential copy to interns.', {
      by: 'Legal-ish',
    });

    Footnotes.add(
      ws,
      'CONFIDENTIAL',
      'Confidentiality is aspirational in this particular document family.'
    );
    Footnotes.add(
      ws,
      'please actually read this one',
      'Nobody read this one. The telemetry is unambiguous.'
    );

    Lists.insertBulletList(ws, 'TABLE OF CONTENTS (good luck)', 'after', [
      'Errata we noticed and chose to weaponize',
      'Conflicting approvals collected from incompatible timelines',
      'Hyperlinks that were never valid in this jurisdiction',
    ]);

    Lists.insertNestedList(ws, 'SECTION B: The Part Nobody Reads', 'after', [
      {
        text: 'Escalation ladder',
        children: [
          { text: 'Ignore issue for two quarters' },
          { text: 'Create steering committee' },
          { text: 'Rename issue as opportunity' },
        ],
      },
      {
        text: 'Remediation plan',
        children: [
          { text: 'Add dashboard nobody requested' },
          { text: 'Color-code the same problem four ways' },
        ],
      },
    ]);

    Lists.insertNumberedList(ws, '8. Recommendations We\'ll Ignore', 'after', [
      'Re-read the financial section without laughing',
      'Delete duplicate approvals only after archiving them twice',
      'Replace the placeholder insight before the next board meeting',
    ]);

    Tables.insert(ws, 'IV. Financial Overview (numbers may be wrong)', 'after', [
      ['Scenario', 'Budget', 'Confidence', 'Narrative'],
      ['Optimistic', '$4.7M', '12%', 'Depends on three miracles and a PDF merge'],
      ['Likely', '$6.2M', '41%', 'Assumes nobody opens the appendix'],
      ['Catastrophic', '$9.9M', '93%', 'Triggered if Dave replies-all again'],
    ], {
      caption: 'Table 404. Budget states observed during uncontrolled editing.',
      style: 'plain',
    });

    Tables.insert(ws, 'PART THE SEVENTH: CONCLUSIONS???', 'after', [
      ['Owner', 'Action', 'Status'],
      ['Operations', 'Pretend the numbers reconcile', 'In progress'],
      ['Finance', 'Invent cleaner footnotes', 'Blocked by reality'],
      ['Comms', 'Shorten the disclaimer paragraph', 'Escalated'],
    ], {
      caption: 'Table 405. Action items with negative momentum.',
      style: 'booktabs',
    });

    Figures.insert(
      ws,
      'EMERGENCY SECTION (added at 2am before the deadline)',
      'after',
      IMAGE,
      'Figure 9001. Incident heatmap generated after everyone stopped sleeping.',
      { width: 2.8 }
    );

    const recommendationsId = findParaIdByText(ws, '8. Recommendations We\'ll Ignore');
    const emergencyId = findParaIdByText(ws, 'EMERGENCY SECTION (added at 2am before the deadline)');
    CrossRef.label(ws, recommendationsId, 'sec:ignored_recommendations');
    CrossRef.ref(ws, 'sec:ignored_recommendations', {
      insertAt: emergencyId,
      after: 'EMERGENCY SECTION (added at 2am before the deadline)',
    });

    injectHyperlinkParagraph(
      ws,
      'Addendum to the Addendum of the Appendix',
      'Disaster recovery wiki',
      'https://example.com/definitely-not-the-final-plan'
    );

    injectStructuredDocumentTag(
      ws,
      'Addendum to the Addendum of the Appendix',
      'Recovery checkbox',
      'chaos.recovery.checkbox',
      'Unchecked because that felt more accurate.'
    );

    injectEndnote(
      ws,
      'Budget: $4.7 million.',
      'This endnote exists solely to ensure the package-level `.dex` path handles active endnotes, not just dormant endnote parts.'
    );

    Provenance.setOrigin(ws, {
      tool: 'docex chaos builder',
      version: 'v3',
      date: '2026-03-27',
      template: 'absolute_chaos_v2',
    });
    Provenance.appendChangelog(ws, [
      {
        timestamp: '2026-03-27T00:00:00Z',
        operation: 'chaos-escalation',
        author: 'Codex',
        description: 'Layered comments, notes, tables, figure, lists, crossrefs, and raw OOXML constructs.',
      },
      {
        timestamp: '2026-03-27T00:05:00Z',
        operation: 'fixture-hardening',
        author: 'Codex',
        description: 'Embedded provenance data, activated dormant package parts, and preserved dex round-trip fidelity.',
      },
    ]);
    Provenance.certify(ws, 'absolute chaos v3 assembled');

    const result = ws.save({ outputPath: OUTPUT, backup: false });
    const counts = summarize(OUTPUT);
    console.log(JSON.stringify({ output: result.path, verified: result.verified, counts }, null, 2));
  } finally {
    ws.cleanup();
  }
}

function findParaIdByText(ws, needle) {
  for (const para of xml.findParagraphs(ws.docXml)) {
    const text = xml.extractTextDecoded(para.xml).replace(/\s+/g, ' ').trim();
    if (!text.includes(needle)) continue;
    const match = para.xml.match(/w14:paraId="([^"]+)"/);
    if (match) return match[1];
  }
  throw new Error(`Could not find paragraph id for "${needle}"`);
}

function insertBodyXmlAfterAnchor(ws, anchor, fragment) {
  const paragraphs = xml.findParagraphs(ws.docXml);
  for (const para of paragraphs) {
    const text = xml.extractTextDecoded(para.xml).replace(/\s+/g, ' ').trim();
    if (!text.includes(anchor)) continue;
    ws.docXml = ws.docXml.slice(0, para.end) + fragment + ws.docXml.slice(para.end);
    return;
  }
  throw new Error(`Anchor not found for raw insert: "${anchor}"`);
}

function injectHyperlinkParagraph(ws, anchor, label, url) {
  const relId = xml.nextRId(ws.relsXml);
  const rel = `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xml.escapeXml(url)}" TargetMode="External"/>`;
  ws.relsXml = ws.relsXml.replace('</Relationships>', rel + '</Relationships>');

  const paraId = xml.randomHexId().toUpperCase();
  const textId = xml.randomHexId().toUpperCase();
  const paragraph =
    `<w:p w14:paraId="${paraId}" w14:textId="${textId}">`
    + '<w:pPr><w:spacing w:after="120"/><w:jc w:val="left"/></w:pPr>'
    + '<w:r><w:t xml:space="preserve">Disaster recovery link dump: </w:t></w:r>'
    + `<w:hyperlink r:id="${relId}" w:history="1">`
    + '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>'
    + `<w:t>${xml.escapeXml(label)}</w:t></w:r>`
    + '</w:hyperlink>'
    + '<w:r><w:t xml:space="preserve"> (do not click during a meeting).</w:t></w:r>'
    + '</w:p>';

  insertBodyXmlAfterAnchor(ws, anchor, paragraph);
}

function injectStructuredDocumentTag(ws, anchor, alias, tag, text) {
  const paraId = xml.randomHexId().toUpperCase();
  const textId = xml.randomHexId().toUpperCase();
  const sdtId = parseInt(xml.randomHexId(), 16);
  const fragment =
    '<w:sdt>'
    + '<w:sdtPr>'
    + `<w:id w:val="${sdtId}"/>`
    + `<w:alias w:val="${xml.escapeXml(alias)}"/>`
    + `<w:tag w:val="${xml.escapeXml(tag)}"/>`
    + '<w:lock w:val="sdtContentLocked"/>'
    + '<w:text/>'
    + '<w:showingPlcHdr/>'
    + '</w:sdtPr>'
    + '<w:sdtContent>'
    + `<w:p w14:paraId="${paraId}" w14:textId="${textId}">`
    + '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:pPr>'
    + '<w:r><w:rPr><w:b/></w:rPr><w:t>[Structured Panic]</w:t></w:r>'
    + `<w:r><w:t xml:space="preserve"> ${xml.escapeXml(text)}</w:t></w:r>`
    + '</w:p>'
    + '</w:sdtContent>'
    + '</w:sdt>';

  insertBodyXmlAfterAnchor(ws, anchor, fragment);
}

function injectEndnote(ws, anchor, noteText) {
  const paragraphs = xml.findParagraphs(ws.docXml);
  let updated = false;

  for (const para of paragraphs) {
    const plain = xml.extractTextDecoded(para.xml);
    if (!plain.includes(anchor)) continue;

    const runs = xml.parseRuns(para.xml).filter(run => run.texts.length > 0);
    let charCount = 0;
    let endRun = null;
    const anchorStart = plain.indexOf(anchor);
    const anchorEnd = anchorStart + anchor.length;

    for (const run of runs) {
      charCount += run.combinedText.length;
      if (charCount >= anchorEnd) {
        endRun = run;
        break;
      }
    }
    if (!endRun) endRun = runs[runs.length - 1];
    if (!endRun) break;

    const endnoteId = nextEndnoteId(ws._readFile('word/endnotes.xml'));
    const refRun = '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>'
      + `<w:endnoteReference w:id="${endnoteId}"/>`
      + '</w:r>';

    const insertAt = endRun.index + endRun.fullMatch.length;
    const newParaXml = para.xml.slice(0, insertAt) + refRun + para.xml.slice(insertAt);
    ws.docXml = ws.docXml.slice(0, para.start) + newParaXml + ws.docXml.slice(para.end);

    const endnoteParaId = xml.randomHexId().toUpperCase();
    const endnoteTextId = xml.randomHexId().toUpperCase();
    const endnote =
      `<w:endnote w:id="${endnoteId}">`
      + `<w:p w14:paraId="${endnoteParaId}" w14:textId="${endnoteTextId}">`
      + '<w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>'
      + '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>'
      + `<w:r><w:t xml:space="preserve"> ${xml.escapeXml(noteText)}</w:t></w:r>`
      + '</w:p>'
      + '</w:endnote>';

    const endnotesXml = ws._readFile('word/endnotes.xml');
    ws._writeFile('word/endnotes.xml', endnotesXml.replace('</w:endnotes>', endnote + '</w:endnotes>'));
    updated = true;
    break;
  }

  if (!updated) {
    throw new Error(`Anchor not found for endnote: "${anchor}"`);
  }
}

function nextEndnoteId(endnotesXml) {
  let max = 0;
  const matches = endnotesXml.match(/<w:endnote\b[^>]*\bw:id="(-?\d+)"/g) || [];
  for (const match of matches) {
    const idMatch = match.match(/w:id="(-?\d+)"/);
    if (!idMatch) continue;
    const value = parseInt(idMatch[1], 10);
    if (value > max) max = value;
  }
  return max + 1;
}

function summarize(docxPath) {
  const ws = Workspace.open(docxPath);
  try {
    return {
      comments: Comments.list(ws).length,
      footnotes: Footnotes.list(ws).length,
      figures: Figures.list(ws).length,
      labels: CrossRef.listLabels(ws).length,
      paragraphs: xml.findParagraphs(ws.docXml).length,
      tables: (ws.docXml.match(/<w:tbl[\s>]/g) || []).length,
      hyperlinks: (ws.docXml.match(/<w:hyperlink\b/g) || []).length,
      endnoteRefs: (ws.docXml.match(/<w:endnoteReference\b/g) || []).length,
      structuredTags: (ws.docXml.match(/<w:sdt\b/g) || []).length,
    };
  } finally {
    ws.cleanup();
  }
}

main();
