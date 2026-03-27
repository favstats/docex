/**
 * metadata.js -- Dublin Core metadata operations for docex
 *
 * Static methods for reading and writing document properties stored in
 * docProps/core.xml (Dublin Core metadata). Manages the root-level
 * relationships (_rels/.rels) and content types as needed.
 *
 * All methods operate on a Workspace object. XML manipulation is done
 * entirely with string operations and regex. Zero external dependencies.
 */

'use strict';

const xml = require('./xml');

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const REL_CORE_PROPS = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CT_CORE_PROPS = 'application/vnd.openxmlformats-package.core-properties+xml';

const EMPTY_CORE_PROPS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  + '<cp:coreProperties'
  + ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
  + ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
  + ' xmlns:dcterms="http://purl.org/dc/terms/"'
  + ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"'
  + ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
  + '</cp:coreProperties>';

// ============================================================================
// METADATA
// ============================================================================

class Metadata {

  /**
   * Parse docProps/core.xml and return metadata properties.
   *
   * @param {object} ws - Workspace with ws.corePropsXml
   * @returns {object} Metadata object with title, creator, subject, description,
   *                   keywords, created, modified, lastModifiedBy, revision.
   *                   Returns empty object if core.xml doesn't exist.
   */
  static get(ws) {
    const coreXml = ws.corePropsXml;
    if (!coreXml) return {};

    return {
      title: Metadata._tagContent(coreXml, 'dc:title'),
      creator: Metadata._tagContent(coreXml, 'dc:creator'),
      subject: Metadata._tagContent(coreXml, 'dc:subject'),
      description: Metadata._tagContent(coreXml, 'dc:description'),
      keywords: Metadata._tagContent(coreXml, 'cp:keywords'),
      created: Metadata._tagContent(coreXml, 'dcterms:created'),
      modified: Metadata._tagContent(coreXml, 'dcterms:modified'),
      lastModifiedBy: Metadata._tagContent(coreXml, 'cp:lastModifiedBy'),
      revision: Metadata._tagContent(coreXml, 'cp:revision'),
    };
  }

  /**
   * Create or update docProps/core.xml with the given properties.
   * Also ensures the root-level relationship and content type exist.
   *
   * @param {object} ws - Workspace with ws.corePropsXml, ws.rootRelsXml, ws.contentTypesXml
   * @param {object} props - Properties to set. Keys: title, creator, subject,
   *                         description, keywords, created, modified, lastModifiedBy, revision
   */
  static set(ws, props) {
    // Ensure core.xml exists
    let coreXml = ws.corePropsXml;
    if (!coreXml) {
      coreXml = EMPTY_CORE_PROPS_XML;
    }

    // Set each property
    if (props.title !== undefined) {
      coreXml = Metadata._setTag(coreXml, 'dc:title', props.title);
    }
    if (props.creator !== undefined) {
      coreXml = Metadata._setTag(coreXml, 'dc:creator', props.creator);
    }
    if (props.subject !== undefined) {
      coreXml = Metadata._setTag(coreXml, 'dc:subject', props.subject);
    }
    if (props.description !== undefined) {
      coreXml = Metadata._setTag(coreXml, 'dc:description', props.description);
    }
    if (props.keywords !== undefined) {
      coreXml = Metadata._setTag(coreXml, 'cp:keywords', props.keywords);
    }
    if (props.created !== undefined) {
      coreXml = Metadata._setDateTag(coreXml, 'dcterms:created', props.created);
    }
    if (props.modified !== undefined) {
      coreXml = Metadata._setDateTag(coreXml, 'dcterms:modified', props.modified);
    }
    if (props.lastModifiedBy !== undefined) {
      coreXml = Metadata._setTag(coreXml, 'cp:lastModifiedBy', props.lastModifiedBy);
    }
    if (props.revision !== undefined) {
      coreXml = Metadata._setTag(coreXml, 'cp:revision', String(props.revision));
    }

    ws.corePropsXml = coreXml;

    // Ensure root-level relationship exists in _rels/.rels
    Metadata._ensureRootRel(ws);

    // Ensure content type exists
    Metadata._ensureContentType(ws);
  }

  // --------------------------------------------------------------------------
  // INTERNAL HELPERS
  // --------------------------------------------------------------------------

  /**
   * Extract text content of an XML element by tag name.
   *
   * @param {string} xmlStr - XML content
   * @param {string} tag - Tag name (e.g. 'dc:title')
   * @returns {string} Text content or empty string
   * @private
   */
  static _tagContent(xmlStr, tag) {
    const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp('<' + escaped + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + escaped + '>');
    const m = xmlStr.match(re);
    if (m) return xml.decodeXml(m[1]);
    return '';
  }

  /**
   * Set or insert a simple XML tag within coreProperties.
   *
   * @param {string} coreXml - docProps/core.xml content
   * @param {string} tag - Tag name (e.g. 'dc:title')
   * @param {string} value - Text value
   * @returns {string} Updated XML
   * @private
   */
  static _setTag(coreXml, tag, value) {
    const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp('<' + escaped + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + escaped + '>');
    const replacement = '<' + tag + '>' + xml.escapeXml(value) + '</' + tag + '>';

    if (re.test(coreXml)) {
      return coreXml.replace(re, replacement);
    }
    // Insert before closing tag
    return coreXml.replace('</cp:coreProperties>', replacement + '</cp:coreProperties>');
  }

  /**
   * Set or insert a dcterms date tag (with xsi:type attribute).
   *
   * @param {string} coreXml - docProps/core.xml content
   * @param {string} tag - Tag name (e.g. 'dcterms:created')
   * @param {string} value - ISO date string
   * @returns {string} Updated XML
   * @private
   */
  static _setDateTag(coreXml, tag, value) {
    const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp('<' + escaped + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + escaped + '>');
    const replacement = '<' + tag + ' xsi:type="dcterms:W3CDTF">' + xml.escapeXml(value) + '</' + tag + '>';

    if (re.test(coreXml)) {
      return coreXml.replace(re, replacement);
    }
    return coreXml.replace('</cp:coreProperties>', replacement + '</cp:coreProperties>');
  }

  /**
   * Ensure the root-level _rels/.rels has a relationship for docProps/core.xml.
   *
   * @param {object} ws - Workspace
   * @private
   */
  static _ensureRootRel(ws) {
    let relsXml = ws.rootRelsXml;
    if (!relsXml) return;

    if (relsXml.includes(REL_CORE_PROPS)) return; // already exists

    // Find next rId in root rels
    let max = 0;
    const re = /Id="rId(\d+)"/g;
    let m;
    while ((m = re.exec(relsXml)) !== null) {
      const n = parseInt(m[1], 10);
      if (n > max) max = n;
    }
    const rId = 'rId' + (max + 1);

    const rel = '<Relationship Id="' + rId + '" Type="' + REL_CORE_PROPS + '" Target="docProps/core.xml"/>';
    relsXml = relsXml.replace('</Relationships>', rel + '</Relationships>');
    ws.rootRelsXml = relsXml;
  }

  /**
   * Ensure [Content_Types].xml has an override for docProps/core.xml.
   *
   * @param {object} ws - Workspace
   * @private
   */
  static _ensureContentType(ws) {
    let ctXml = ws.contentTypesXml;
    if (!ctXml) return;

    if (ctXml.includes('/docProps/core.xml')) return; // already exists

    ctXml = ctXml.replace('</Types>',
      '<Override PartName="/docProps/core.xml" ContentType="' + CT_CORE_PROPS + '"/></Types>');
    ws.contentTypesXml = ctXml;
  }
}

module.exports = { Metadata };
