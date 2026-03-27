/**
 * provenance.js -- Provenance and changelog tracking for docex (v0.4.9)
 *
 * Embeds changelog, origin, and certification data inside the .docx as
 * custom XML parts (customXml/). This is a standard OOXML extension
 * mechanism -- the data survives file moves and renames.
 *
 * Custom XML parts:
 *   - customXml/docex-changelog.xml
 *   - customXml/docex-origin.xml
 *   - customXml/docex-certifications.xml
 *
 * Zero external dependencies. Uses Node.js built-in crypto for SHA-256.
 */

'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const xml = require('./xml');

class Provenance {

  // ==========================================================================
  // CHANGELOG
  // ==========================================================================

  /**
   * Read the changelog from the customXml part.
   *
   * @param {object} ws - Workspace
   * @returns {Array<{timestamp: string, operation: string, author: string, description: string}>}
   */
  static getChangelog(ws) {
    const content = Provenance._readCustomXml(ws, 'docex-changelog.xml');
    if (!content) return [];
    return Provenance._parseChangelog(content);
  }

  /**
   * Append operation entries to the changelog custom XML part.
   * Creates the custom XML part if it doesn't exist.
   *
   * @param {object} ws - Workspace
   * @param {Array<{timestamp: string, operation: string, author: string, description: string}>} entries
   */
  static appendChangelog(ws, entries) {
    if (!entries || entries.length === 0) return;

    const existing = Provenance.getChangelog(ws);
    const all = existing.concat(entries);
    const xmlContent = Provenance._buildChangelogXml(all);
    Provenance._writeCustomXml(ws, 'docex-changelog.xml', xmlContent);
  }

  /**
   * Filter changelog entries after the given date.
   *
   * @param {object} ws - Workspace
   * @param {string} dateString - ISO date string or date prefix (e.g. "2026-03-18")
   * @returns {Array<{timestamp: string, operation: string, author: string, description: string}>}
   */
  static changelogSince(ws, dateString) {
    const all = Provenance.getChangelog(ws);
    const threshold = new Date(dateString).getTime();
    return all.filter(entry => {
      const entryTime = new Date(entry.timestamp).getTime();
      return entryTime >= threshold;
    });
  }

  // ==========================================================================
  // ORIGIN
  // ==========================================================================

  /**
   * Read the origin info from the custom XML part.
   * Returns a human-readable string.
   *
   * @param {object} ws - Workspace
   * @returns {string} e.g. "Created by docex v0.3.0 on 2026-03-27 from template polcomm. 47 operations across 3 sessions."
   */
  static origin(ws) {
    const content = Provenance._readCustomXml(ws, 'docex-origin.xml');
    if (!content) return 'No origin information recorded.';

    const info = Provenance._parseOrigin(content);
    const changelog = Provenance.getChangelog(ws);

    // Count operations and unique sessions
    const opCount = changelog.length;
    const sessions = new Set();
    for (const entry of changelog) {
      // Group by date (YYYY-MM-DD) as a rough session proxy
      const dateStr = entry.timestamp ? entry.timestamp.slice(0, 10) : 'unknown';
      sessions.add(dateStr);
    }

    let result = 'Created by ' + (info.tool || 'docex') + ' ' + (info.version || 'unknown');
    result += ' on ' + (info.date || 'unknown');
    if (info.template) {
      result += ' from template ' + info.template;
    }
    result += '.';

    if (opCount > 0) {
      result += ' ' + opCount + ' operation' + (opCount !== 1 ? 's' : '');
      result += ' across ' + sessions.size + ' session' + (sessions.size !== 1 ? 's' : '') + '.';
    }

    return result;
  }

  /**
   * Write origin info to the custom XML part.
   *
   * @param {object} ws - Workspace
   * @param {object} info - { version, date, template, tool }
   */
  static setOrigin(ws, info) {
    const xmlContent = Provenance._buildOriginXml(info);
    Provenance._writeCustomXml(ws, 'docex-origin.xml', xmlContent);
  }

  // ==========================================================================
  // CERTIFICATIONS
  // ==========================================================================

  /**
   * Compute SHA-256 of document.xml content and store a certification.
   *
   * @param {object} ws - Workspace
   * @param {string} label - e.g. "submitted to Political Communication"
   */
  static certify(ws, label) {
    const hash = Provenance._hashDocumentXml(ws);
    const date = new Date().toISOString();

    const existing = Provenance.certifications(ws);
    existing.push({ label, date, hash });

    const xmlContent = Provenance._buildCertificationsXml(existing);
    Provenance._writeCustomXml(ws, 'docex-certifications.xml', xmlContent);
  }

  /**
   * Verify if document content still matches the last certification hash.
   *
   * @param {object} ws - Workspace
   * @returns {{ certified: boolean, label: string, date: string, hash: string }}
   */
  static verifyCertification(ws) {
    const certs = Provenance.certifications(ws);
    if (certs.length === 0) {
      return { certified: false, label: '', date: '', hash: '' };
    }

    const latest = certs[certs.length - 1];
    const currentHash = Provenance._hashDocumentXml(ws);

    return {
      certified: currentHash === latest.hash,
      label: latest.label,
      date: latest.date,
      hash: latest.hash,
    };
  }

  /**
   * List all certification points.
   *
   * @param {object} ws - Workspace
   * @returns {Array<{label: string, date: string, hash: string}>}
   */
  static certifications(ws) {
    const content = Provenance._readCustomXml(ws, 'docex-certifications.xml');
    if (!content) return [];
    return Provenance._parseCertifications(content);
  }

  // ==========================================================================
  // CUSTOM XML PART I/O
  // ==========================================================================

  /**
   * Read a custom XML file from the workspace temp directory.
   *
   * @param {object} ws - Workspace
   * @param {string} filename - e.g. 'docex-changelog.xml'
   * @returns {string|null} File content or null
   * @private
   */
  static _readCustomXml(ws, filename) {
    const filePath = path.join(ws.tmpDir, 'customXml', filename);
    if (fs.existsSync(filePath)) {
      return fs.readFileSync(filePath, 'utf-8');
    }
    return null;
  }

  /**
   * Write a custom XML file to the workspace temp directory.
   * Also ensures Content_Types and relationships are updated.
   *
   * @param {object} ws - Workspace
   * @param {string} filename - e.g. 'docex-changelog.xml'
   * @param {string} content - XML content
   * @private
   */
  static _writeCustomXml(ws, filename, content) {
    const customXmlDir = path.join(ws.tmpDir, 'customXml');
    if (!fs.existsSync(customXmlDir)) {
      fs.mkdirSync(customXmlDir, { recursive: true });
    }

    const filePath = path.join(customXmlDir, filename);
    fs.writeFileSync(filePath, content, 'utf-8');

    // Ensure Content_Types.xml has an entry for this file
    Provenance._ensureContentType(ws, filename);
  }

  /**
   * Ensure [Content_Types].xml has an Override entry for a custom XML part.
   *
   * @param {object} ws - Workspace
   * @param {string} filename - Custom XML filename
   * @private
   */
  static _ensureContentType(ws, filename) {
    let ct = ws.contentTypesXml;
    const partName = '/customXml/' + filename;
    const contentType = 'application/xml';

    // Check if already present
    if (ct.includes(partName)) return;

    // Insert before </Types>
    const insertPoint = ct.lastIndexOf('</Types>');
    if (insertPoint === -1) return;

    const override = '<Override PartName="' + partName + '" ContentType="' + contentType + '"/>';
    ws.contentTypesXml = ct.slice(0, insertPoint) + override + ct.slice(insertPoint);
  }

  // ==========================================================================
  // XML BUILDERS
  // ==========================================================================

  /**
   * Build changelog XML content.
   * @param {Array} entries
   * @returns {string}
   * @private
   */
  static _buildChangelogXml(entries) {
    let xmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xmlStr += '<docex-changelog xmlns="https://docex.dev/changelog">';

    for (const entry of entries) {
      xmlStr += '<entry>';
      xmlStr += '<timestamp>' + xml.escapeXml(entry.timestamp || '') + '</timestamp>';
      xmlStr += '<operation>' + xml.escapeXml(entry.operation || '') + '</operation>';
      xmlStr += '<author>' + xml.escapeXml(entry.author || '') + '</author>';
      xmlStr += '<description>' + xml.escapeXml(entry.description || '') + '</description>';
      xmlStr += '</entry>';
    }

    xmlStr += '</docex-changelog>';
    return xmlStr;
  }

  /**
   * Build origin XML content.
   * @param {object} info - { version, date, template, tool }
   * @returns {string}
   * @private
   */
  static _buildOriginXml(info) {
    let xmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xmlStr += '<docex-origin xmlns="https://docex.dev/origin">';
    xmlStr += '<version>' + xml.escapeXml(info.version || '') + '</version>';
    xmlStr += '<date>' + xml.escapeXml(info.date || '') + '</date>';
    xmlStr += '<template>' + xml.escapeXml(info.template || '') + '</template>';
    xmlStr += '<tool>' + xml.escapeXml(info.tool || 'docex') + '</tool>';
    xmlStr += '</docex-origin>';
    return xmlStr;
  }

  /**
   * Build certifications XML content.
   * @param {Array} certs
   * @returns {string}
   * @private
   */
  static _buildCertificationsXml(certs) {
    let xmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xmlStr += '<docex-certifications xmlns="https://docex.dev/certifications">';

    for (const cert of certs) {
      xmlStr += '<certification>';
      xmlStr += '<label>' + xml.escapeXml(cert.label || '') + '</label>';
      xmlStr += '<date>' + xml.escapeXml(cert.date || '') + '</date>';
      xmlStr += '<hash>' + xml.escapeXml(cert.hash || '') + '</hash>';
      xmlStr += '</certification>';
    }

    xmlStr += '</docex-certifications>';
    return xmlStr;
  }

  // ==========================================================================
  // XML PARSERS
  // ==========================================================================

  /**
   * Parse changelog XML into entry objects.
   * @param {string} xmlContent
   * @returns {Array}
   * @private
   */
  static _parseChangelog(xmlContent) {
    const entries = [];
    const entryRe = /<entry>([\s\S]*?)<\/entry>/g;
    let m;
    while ((m = entryRe.exec(xmlContent)) !== null) {
      const body = m[1];
      entries.push({
        timestamp: Provenance._extractTagText(body, 'timestamp'),
        operation: Provenance._extractTagText(body, 'operation'),
        author: Provenance._extractTagText(body, 'author'),
        description: Provenance._extractTagText(body, 'description'),
      });
    }
    return entries;
  }

  /**
   * Parse origin XML into an info object.
   * @param {string} xmlContent
   * @returns {object}
   * @private
   */
  static _parseOrigin(xmlContent) {
    return {
      version: Provenance._extractTagText(xmlContent, 'version'),
      date: Provenance._extractTagText(xmlContent, 'date'),
      template: Provenance._extractTagText(xmlContent, 'template'),
      tool: Provenance._extractTagText(xmlContent, 'tool'),
    };
  }

  /**
   * Parse certifications XML into an array.
   * @param {string} xmlContent
   * @returns {Array}
   * @private
   */
  static _parseCertifications(xmlContent) {
    const certs = [];
    const certRe = /<certification>([\s\S]*?)<\/certification>/g;
    let m;
    while ((m = certRe.exec(xmlContent)) !== null) {
      const body = m[1];
      certs.push({
        label: Provenance._extractTagText(body, 'label'),
        date: Provenance._extractTagText(body, 'date'),
        hash: Provenance._extractTagText(body, 'hash'),
      });
    }
    return certs;
  }

  /**
   * Extract text content from a simple XML tag.
   * @param {string} xmlStr - XML string
   * @param {string} tag - Tag name
   * @returns {string}
   * @private
   */
  static _extractTagText(xmlStr, tag) {
    const re = new RegExp('<' + tag + '>([\\s\\S]*?)</' + tag + '>');
    const m = xmlStr.match(re);
    return m ? xml.decodeXml(m[1]) : '';
  }

  /**
   * Compute SHA-256 hash of the document.xml content.
   * @param {object} ws - Workspace
   * @returns {string} Hex hash
   * @private
   */
  static _hashDocumentXml(ws) {
    return crypto.createHash('sha256').update(ws.docXml).digest('hex');
  }
}

module.exports = { Provenance };
