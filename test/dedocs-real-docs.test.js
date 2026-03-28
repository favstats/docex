'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  compareDocxPackages,
  compileDedocsText,
  dedocsFromDocx,
} = require('../dedocs');

const REAL_DOCS = [
  '/mnt/storage/euroalgos/paper/manuscript.docx',
  '/mnt/storage/absolute_chaos.docx',
  '/mnt/storage/absolute_chaos_v2.docx',
  '/mnt/storage/absolute_chaos_v3.docx',
];

function tmpDocxPath(label) {
  return path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'dedocs-real-')), `${label}.docx`);
}

describe('dedocs real documents', () => {
  for (const docxPath of REAL_DOCS) {
    const exists = fs.existsSync(docxPath);
    const label = path.basename(docxPath, '.docx');

    it(`round-trips ${label} package-exactly`, { skip: !exists && `missing ${docxPath}` }, () => {
      const dedocsText = dedocsFromDocx(docxPath, { source: path.basename(docxPath) });
      const rebuiltPath = tmpDocxPath(label);
      compileDedocsText(dedocsText, rebuiltPath, { strictMetadata: true });

      const comparison = compareDocxPackages(docxPath, rebuiltPath);
      assert.equal(comparison.equal, true, JSON.stringify(comparison.diffs, null, 2));
    });
  }
});
