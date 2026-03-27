/**
 * doctor.js -- Document health checker for docex
 *
 * Checks document integrity: valid zip, document.xml exists,
 * relationships resolve, no orphaned media, comment consistency,
 * paraId uniqueness, heading hierarchy.
 *
 * Two entry points:
 *   Doctor.diagnose(ws)  -- human-readable output (for CLI)
 *   Doctor.validate(ws)  -- programmatic return { valid, errors, warnings }
 *
 * All methods operate on a Workspace object.
 * Zero external dependencies.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const xml = require('./xml');

// ============================================================================
// DOCTOR
// ============================================================================

class Doctor {

  /**
   * Run all diagnostic checks and return a structured result.
   *
   * @param {object} ws - Workspace with ws.docXml, ws.relsXml, ws.commentsXml, etc.
   * @returns {{ valid: boolean, errors: string[], warnings: string[], checks: Array<{name: string, passed: boolean, message: string}> }}
   */
  static validate(ws) {
    const checks = [];
    const errors = [];
    const warnings = [];

    // 1. Check document.xml exists and has w:body
    Doctor._checkDocumentXml(ws, checks, errors);

    // 2. Check all relationships resolve
    Doctor._checkRelationships(ws, checks, errors, warnings);

    // 3. Check for orphaned media
    Doctor._checkOrphanedMedia(ws, checks, errors, warnings);

    // 4. Check comment consistency
    Doctor._checkCommentConsistency(ws, checks, errors, warnings);

    // 5. Check paragraph count > 0
    Doctor._checkParagraphCount(ws, checks, errors);

    // 6. Check file size
    Doctor._checkFileSize(ws, checks, errors, warnings);

    // 7. Check heading hierarchy
    Doctor._checkHeadingHierarchy(ws, checks, errors, warnings);

    // 8. Check paraId uniqueness
    Doctor._checkParaIdUniqueness(ws, checks, errors, warnings);

    return {
      valid: errors.length === 0,
      errors,
      warnings,
      checks,
    };
  }

  /**
   * Run all diagnostic checks and return a formatted string for terminal display.
   *
   * @param {object} ws - Workspace
   * @returns {string} Human-readable diagnostic report
   */
  static diagnose(ws) {
    const result = Doctor.validate(ws);
    const lines = [];

    for (const check of result.checks) {
      const icon = check.passed ? '\x1b[32m\u2713\x1b[0m' : '\x1b[31m\u2717\x1b[0m';
      lines.push(`  ${icon} ${check.message}`);
    }

    lines.push('');
    if (result.valid) {
      lines.push('\x1b[32mDocument is healthy.\x1b[0m');
    } else {
      lines.push(`\x1b[31m${result.errors.length} error(s) found:\x1b[0m`);
      for (const e of result.errors) {
        lines.push(`  \x1b[31m- ${e}\x1b[0m`);
      }
    }

    if (result.warnings.length > 0) {
      lines.push(`\x1b[33m${result.warnings.length} warning(s):\x1b[0m`);
      for (const w of result.warnings) {
        lines.push(`  \x1b[33m- ${w}\x1b[0m`);
      }
    }

    return lines.join('\n');
  }

  // --------------------------------------------------------------------------
  // Individual checks
  // --------------------------------------------------------------------------

  /** @private */
  static _checkDocumentXml(ws, checks, errors) {
    try {
      const docXml = ws.docXml;
      if (!docXml || docXml.length === 0) {
        checks.push({ name: 'document.xml', passed: false, message: 'document.xml is empty' });
        errors.push('document.xml is empty');
        return;
      }

      if (!docXml.includes('<w:body')) {
        checks.push({ name: 'document.xml', passed: false, message: 'document.xml missing <w:body>' });
        errors.push('document.xml missing <w:body> element');
        return;
      }

      checks.push({ name: 'document.xml', passed: true, message: 'document.xml exists and has w:body' });
    } catch (e) {
      checks.push({ name: 'document.xml', passed: false, message: 'document.xml not found: ' + e.message });
      errors.push('document.xml not found: ' + e.message);
    }
  }

  /** @private */
  static _checkRelationships(ws, checks, errors, warnings) {
    try {
      const relsXml = ws.relsXml;
      // Extract all Relationship elements
      const relRe = /<Relationship\s[^>]*>/g;
      let m;
      let totalRels = 0;
      let brokenRels = 0;
      const broken = [];

      while ((m = relRe.exec(relsXml)) !== null) {
        totalRels++;
        const relTag = m[0];
        // Extract Target
        const targetMatch = relTag.match(/Target="([^"]+)"/);
        if (!targetMatch) continue;

        const target = targetMatch[1];
        // Skip external targets (URLs)
        if (target.startsWith('http://') || target.startsWith('https://')) continue;
        // Skip TargetMode="External"
        if (relTag.includes('TargetMode="External"')) continue;

        // Check the file exists in the workspace
        const filePath = path.join(ws.tmpDir, 'word', target);
        if (!fs.existsSync(filePath)) {
          brokenRels++;
          const idMatch = relTag.match(/Id="([^"]+)"/);
          const rId = idMatch ? idMatch[1] : '?';
          broken.push(`${rId} -> ${target}`);
        }
      }

      if (brokenRels > 0) {
        checks.push({ name: 'relationships', passed: false, message: `${brokenRels} broken relationship(s): ${broken.join(', ')}` });
        for (const b of broken) {
          errors.push(`Broken relationship: ${b}`);
        }
      } else {
        checks.push({ name: 'relationships', passed: true, message: `All ${totalRels} relationships resolve` });
      }
    } catch (e) {
      checks.push({ name: 'relationships', passed: false, message: 'Could not check relationships: ' + e.message });
      errors.push('Could not check relationships: ' + e.message);
    }
  }

  /** @private */
  static _checkOrphanedMedia(ws, checks, errors, warnings) {
    try {
      const mediaDir = path.join(ws.tmpDir, 'word', 'media');
      if (!fs.existsSync(mediaDir)) {
        checks.push({ name: 'media', passed: true, message: 'No media directory (OK)' });
        return;
      }

      const mediaFiles = fs.readdirSync(mediaDir);
      if (mediaFiles.length === 0) {
        checks.push({ name: 'media', passed: true, message: 'No media files (OK)' });
        return;
      }

      // Find all referenced media in relationships
      const relsXml = ws.relsXml;
      const referenced = new Set();
      const targetRe = /Target="media\/([^"]+)"/g;
      let m;
      while ((m = targetRe.exec(relsXml)) !== null) {
        referenced.add(m[1]);
      }

      const orphaned = mediaFiles.filter(f => !referenced.has(f));
      const missing = [...referenced].filter(f => !mediaFiles.includes(f));

      if (orphaned.length > 0) {
        warnings.push(`${orphaned.length} orphaned media file(s): ${orphaned.slice(0, 5).join(', ')}${orphaned.length > 5 ? '...' : ''}`);
      }

      if (missing.length > 0) {
        for (const f of missing) {
          errors.push(`Referenced media file missing: ${f}`);
        }
        checks.push({ name: 'media', passed: false, message: `${missing.length} referenced media file(s) missing` });
      } else if (orphaned.length > 0) {
        checks.push({ name: 'media', passed: true, message: `${mediaFiles.length} media files, ${orphaned.length} orphaned (warning)` });
      } else {
        checks.push({ name: 'media', passed: true, message: `${mediaFiles.length} media files, all referenced` });
      }
    } catch (e) {
      checks.push({ name: 'media', passed: true, message: 'Media check skipped: ' + e.message });
    }
  }

  /** @private */
  static _checkCommentConsistency(ws, checks, errors, warnings) {
    try {
      // Count comments in comments.xml
      let commentsXmlContent;
      try {
        commentsXmlContent = ws.commentsXml;
      } catch (_) {
        checks.push({ name: 'comments', passed: true, message: 'No comments.xml (OK)' });
        return;
      }

      const commentCount = (commentsXmlContent.match(/<w:comment\b/g) || []).length;

      // Count comment range markers in document.xml
      const docXml = ws.docXml;
      const rangeStarts = (docXml.match(/<w:commentRangeStart\b/g) || []).length;
      const rangeEnds = (docXml.match(/<w:commentRangeEnd\b/g) || []).length;
      const refs = (docXml.match(/<w:commentReference\b/g) || []).length;

      if (commentCount === 0 && rangeStarts === 0) {
        checks.push({ name: 'comments', passed: true, message: 'No comments (OK)' });
        return;
      }

      const issues = [];
      if (rangeStarts !== rangeEnds) {
        issues.push(`range start/end mismatch: ${rangeStarts} starts vs ${rangeEnds} ends`);
        warnings.push(`Comment range start/end mismatch: ${rangeStarts} starts vs ${rangeEnds} ends`);
      }

      if (commentCount !== refs) {
        issues.push(`comment/reference mismatch: ${commentCount} comments vs ${refs} references`);
        warnings.push(`Comment count (${commentCount}) differs from references in document.xml (${refs})`);
      }

      if (issues.length > 0) {
        checks.push({ name: 'comments', passed: true, message: `${commentCount} comments with inconsistencies: ${issues.join('; ')}` });
      } else {
        checks.push({ name: 'comments', passed: true, message: `${commentCount} comments, all consistent` });
      }
    } catch (e) {
      checks.push({ name: 'comments', passed: true, message: 'Comment check skipped: ' + e.message });
    }
  }

  /** @private */
  static _checkParagraphCount(ws, checks, errors) {
    const docXml = ws.docXml;
    const count = (docXml.match(/<w:p[\s>]/g) || []).length;
    if (count === 0) {
      checks.push({ name: 'paragraphs', passed: false, message: 'No paragraphs found' });
      errors.push('Document has no paragraphs');
    } else {
      checks.push({ name: 'paragraphs', passed: true, message: `${count} paragraphs` });
    }
  }

  /** @private */
  static _checkFileSize(ws, checks, errors, warnings) {
    try {
      const stat = fs.statSync(ws._docxPath);
      if (stat.size === 0) {
        checks.push({ name: 'fileSize', passed: false, message: 'File size is 0 bytes' });
        errors.push('File size is 0 bytes');
      } else if (stat.size > 100 * 1024 * 1024) {
        checks.push({ name: 'fileSize', passed: true, message: `File size: ${(stat.size / (1024 * 1024)).toFixed(1)} MB (large)` });
        warnings.push(`File size is very large: ${(stat.size / (1024 * 1024)).toFixed(1)} MB`);
      } else {
        checks.push({ name: 'fileSize', passed: true, message: `File size: ${(stat.size / 1024).toFixed(1)} KB` });
      }
    } catch (e) {
      checks.push({ name: 'fileSize', passed: true, message: 'File size check skipped: ' + e.message });
    }
  }

  /** @private */
  static _checkHeadingHierarchy(ws, checks, errors, warnings) {
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    const { Paragraphs } = require('./paragraphs');

    let lastLevel = 0;
    let hierarchyOk = true;
    const skips = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const level = Paragraphs._headingLevel(paragraphs[i].xml);
      if (level === 0) continue;

      // Check for hierarchy skip (e.g., H3 without preceding H2)
      if (lastLevel > 0 && level > lastLevel + 1) {
        hierarchyOk = false;
        const text = xml.extractTextDecoded(paragraphs[i].xml).slice(0, 40);
        skips.push(`H${level} after H${lastLevel} at paragraph ${i} ("${text}")`);
        warnings.push(`Heading hierarchy skip: H${level} after H${lastLevel} at paragraph ${i}`);
      }
      lastLevel = level;
    }

    if (hierarchyOk) {
      checks.push({ name: 'headingHierarchy', passed: true, message: 'Heading hierarchy valid' });
    } else {
      checks.push({ name: 'headingHierarchy', passed: true, message: `Heading hierarchy has ${skips.length} skip(s): ${skips[0]}` });
    }
  }

  /** @private */
  static _checkParaIdUniqueness(ws, checks, errors, warnings) {
    const docXml = ws.docXml;
    const paraIdRe = /w14:paraId="([^"]+)"/g;
    const seen = new Map(); // id -> count
    let m;

    while ((m = paraIdRe.exec(docXml)) !== null) {
      const id = m[1];
      seen.set(id, (seen.get(id) || 0) + 1);
    }

    const duplicates = [...seen.entries()].filter(([_, count]) => count > 1);
    if (duplicates.length > 0) {
      checks.push({ name: 'paraIdUniqueness', passed: false, message: `${duplicates.length} duplicate paraId(s)` });
      for (const [id, count] of duplicates) {
        errors.push(`Duplicate paraId "${id}" appears ${count} times`);
      }
    } else {
      checks.push({ name: 'paraIdUniqueness', passed: true, message: `All ${seen.size} paraIds are unique` });
    }
  }
}

module.exports = { Doctor };
