/**
 * citations.js -- Citation operations for docex
 *
 * Static methods for finding plain-text citation patterns in OOXML documents
 * and replacing them with ZOTERO_CITATION field codes that the Zotero plugin
 * can manage.
 *
 * Two main workflows:
 *   1. Citations.list(ws) -- find all (Author, Year) patterns (no network)
 *   2. Citations.inject(ws, options) -- find patterns, query Zotero API,
 *      replace with OOXML field codes
 *
 * Ported from inject_citations.js (nl_local_2026/paper/build/).
 * All methods operate on a Workspace object. Zero external dependencies
 * (uses Node.js built-in https module for API calls).
 */

'use strict';

const https = require('https');
const xml = require('./xml');

// ============================================================================
// CITATION PATTERNS
// ============================================================================

/**
 * Regex for parenthetical citations: (Author, Year) or (Author & Author, Year)
 * Also matches: (Author et al., Year), (Author, Year, p. 10), multi-cite
 * (Author, Year; Author, Year)
 *
 * This finds candidates in decoded (plain) text. The matching is deliberately
 * broad -- we capture the full parenthetical and then parse its contents.
 */
const PARENTHETICAL_RE = /\(([A-Z\u00C0-\u024F][^\)]{2,120},\s*\d{4}[a-z]?(?:[,;][^\)]*)?)\)/g;

/**
 * Regex for narrative citations: Author (Year) or Author and Author (Year)
 * Captures the author portion and the year parenthetical separately.
 */
const NARRATIVE_RE = /([A-Z\u00C0-\u024F][A-Za-z\u00C0-\u024F'\u2019-]+(?:(?:\s+(?:and|&)\s+|\s*,\s*(?:and\s+|&\s+)?)[A-Z\u00C0-\u024F][A-Za-z\u00C0-\u024F'\u2019-]+)*(?:\s+et\s+al\.)?)\s+\((\d{4}[a-z]?)\)/g;

// ============================================================================
// ZOTERO API
// ============================================================================

/**
 * Fetch JSON from the Zotero Web API.
 * Uses Node.js built-in https module (zero dependencies).
 *
 * @param {string} apiPath - API path (e.g. /users/123/collections/ABC/items)
 * @param {string} apiKey - Zotero API key
 * @returns {Promise<object>} Parsed JSON response
 * @private
 */
function _zoteroFetch(apiPath, apiKey) {
  return new Promise(function (resolve, reject) {
    const opts = {
      hostname: 'api.zotero.org',
      path: apiPath,
      headers: { 'Zotero-API-Key': apiKey },
    };
    https.get(opts, function (res) {
      let data = '';
      res.on('data', function (chunk) { data += chunk; });
      res.on('end', function () {
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          reject(new Error('Zotero API parse error: ' + data.substring(0, 200)));
        }
      });
    }).on('error', reject);
  });
}

/**
 * Fetch all items from a Zotero collection, handling pagination.
 *
 * @param {string} userId - Zotero user ID
 * @param {string} collectionId - Zotero collection key
 * @param {string} apiKey - Zotero API key
 * @returns {Promise<Array>} Array of Zotero item objects
 * @private
 */
async function _fetchCollectionItems(userId, collectionId, apiKey) {
  const items = [];
  let start = 0;
  const limit = 100;

  while (true) {
    const batch = await _zoteroFetch(
      '/users/' + userId + '/collections/' + collectionId +
      '/items?format=json&limit=' + limit + '&start=' + start,
      apiKey
    );
    if (!Array.isArray(batch) || batch.length === 0) break;
    items.push(...batch);
    if (batch.length < limit) break;
    start += limit;
  }

  return items;
}

/**
 * Search user's entire library for items matching author/year.
 *
 * @param {string} userId - Zotero user ID
 * @param {string} authorLastName - Author's last name to search for
 * @param {string} year - Publication year
 * @param {string} apiKey - Zotero API key
 * @returns {Promise<Array>} Matching Zotero items
 * @private
 */
async function _searchItems(userId, authorLastName, year, apiKey) {
  const searchPath =
    '/users/' + userId + '/items?format=json&limit=25' +
    '&q=' + encodeURIComponent(authorLastName) +
    '&qmode=titleCreatorYear';
  const results = await _zoteroFetch(searchPath, apiKey);
  if (!Array.isArray(results)) return [];

  // Filter by year
  return results.filter(function (item) {
    const d = item.data;
    const dateStr = d.date || d.dateEnacted || d.dateDecided || '';
    const yearMatch = dateStr.match(/(\d{4})/);
    return yearMatch && yearMatch[1] === year;
  });
}

// ============================================================================
// CSL DATA BUILDERS
// ============================================================================

/**
 * Map Zotero item type to CSL type.
 * @param {string} itemType - Zotero itemType value
 * @returns {string} CSL type string
 * @private
 */
function _zoteroTypeToCsl(itemType) {
  const map = {
    journalArticle: 'article-journal',
    book: 'book',
    bookSection: 'chapter',
    conferencePaper: 'paper-conference',
    report: 'report',
    thesis: 'thesis',
    webpage: 'webpage',
    newspaperArticle: 'article-newspaper',
    magazineArticle: 'article-magazine',
    statute: 'legislation',
    case: 'legal_case',
    document: 'document',
  };
  return map[itemType] || 'article';
}

/**
 * Build a CSL-JSON item from a Zotero API item object.
 * @param {object} zItem - Zotero item from the API
 * @returns {object} CSL-JSON item data
 * @private
 */
function _buildCslItemData(zItem) {
  const d = zItem.data;
  const csl = {
    id: d.key,
    type: _zoteroTypeToCsl(d.itemType),
    title: d.title || d.nameOfAct || d.caseName || '',
  };

  if (d.creators && d.creators.length > 0) {
    csl.author = d.creators
      .filter(function (c) { return c.creatorType === 'author'; })
      .map(function (c) {
        if (c.name) return { literal: c.name };
        return { family: c.lastName, given: c.firstName };
      });
  }

  const dateStr = d.date || d.dateEnacted || d.dateDecided || '';
  if (dateStr) {
    const yearMatch = dateStr.match(/(\d{4})/);
    if (yearMatch) csl.issued = { 'date-parts': [[parseInt(yearMatch[1])]] };
  }

  if (d.publicationTitle) csl['container-title'] = d.publicationTitle;
  if (d.publisher) csl.publisher = d.publisher;
  if (d.volume) csl.volume = d.volume;
  if (d.issue) csl.issue = d.issue;
  if (d.pages) csl.page = d.pages;
  if (d.DOI) csl.DOI = d.DOI;
  if (d.url) csl.URL = d.url;

  return csl;
}

// ============================================================================
// OOXML FIELD CODE BUILDERS
// ============================================================================

/** Counter for unique citation IDs within a single injection session. */
let _citationCounter = 0;

/**
 * Build the ZOTERO_CITATION JSON payload.
 *
 * @param {Array<object>} cslItems - Array of {zoteroKey, cslData} objects
 * @param {string} userId - Zotero user ID
 * @param {boolean} [suppressAuthor] - Whether to suppress author names
 * @returns {object} Zotero citation JSON object
 * @private
 */
function _buildZoteroCitationJson(cslItems, userId, suppressAuthor) {
  const items = cslItems.map(function (entry) {
    const item = {
      id: entry.zoteroKey,
      uris: ['http://zotero.org/users/' + userId + '/items/' + entry.zoteroKey],
      itemData: entry.cslData,
    };
    if (suppressAuthor) item['suppress-author'] = true;
    return item;
  });

  return {
    citationID: 'cite_' + (++_citationCounter),
    citationItems: items,
    properties: { noteIndex: 0 },
    schema: 'https://raw.githubusercontent.com/citation-style-language/schema/master/csl-citation.json',
  };
}

/**
 * Build OOXML field code XML for a ZOTERO_CITATION.
 *
 * @param {object} citationJson - The citation JSON payload
 * @param {string} displayText - Text shown to the user (e.g. "(Author, 2024)")
 * @param {string} rPrXml - Run properties XML to apply to all runs
 * @returns {string} OOXML XML string with field codes
 * @private
 */
function _buildCitationFieldXml(citationJson, displayText, rPrXml) {
  const jsonStr = xml.escapeXml(JSON.stringify(citationJson));
  const instrText = ' ADDIN ZOTERO_CITATION CSL_CITATION ' + jsonStr + ' ';
  const safeDisplay = xml.escapeXml(displayText);

  return (
    '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r>' + rPrXml + '<w:instrText xml:space="preserve">' + instrText + '</w:instrText></w:r>' +
    '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r>' + rPrXml + '<w:t xml:space="preserve">' + safeDisplay + '</w:t></w:r>' +
    '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="end"/></w:r>'
  );
}

/**
 * Build OOXML field code XML for a ZOTERO_BIBL (bibliography).
 *
 * @param {string} rPrXml - Run properties XML
 * @returns {string} OOXML XML string with bibliography field code
 * @private
 */
function _buildBibliographyFieldXml(rPrXml) {
  const instrText = ' ADDIN ZOTERO_BIBL {&quot;uncited&quot;:[],&quot;omitted&quot;:[],&quot;custom&quot;:[]} CSL_BIBLIOGRAPHY ';
  return (
    '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r>' + rPrXml + '<w:instrText xml:space="preserve">' + instrText + '</w:instrText></w:r>' +
    '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r>' + rPrXml + '<w:t xml:space="preserve">[Bibliography will be generated by Zotero]</w:t></w:r>' +
    '<w:r>' + rPrXml + '<w:fldChar w:fldCharType="end"/></w:r>'
  );
}

// ============================================================================
// CITATION FINDING (text-level)
// ============================================================================

/**
 * Find all citation-like patterns in a plain text string.
 * Returns both parenthetical (Author, Year) and narrative Author (Year) forms.
 *
 * @param {string} text - Plain text to scan
 * @returns {Array<{text: string, start: number, end: number, type: string, authors: string, year: string}>}
 * @private
 */
function _findCitationPatternsInText(text) {
  const results = [];

  // Find parenthetical citations: (Author, Year)
  let m;
  const parentRe = new RegExp(PARENTHETICAL_RE.source, 'g');
  while ((m = parentRe.exec(text)) !== null) {
    // Parse the inner content to extract author names and year
    const inner = m[1];
    // Could be multi-cite: (Author, 2020; Author, 2021)
    const parts = inner.split(/\s*;\s*/);
    for (const part of parts) {
      const yearMatch = part.match(/(\d{4}[a-z]?)(?:\s*,\s*p(?:p)?\.?\s*\d+)?$/);
      if (yearMatch) {
        const authorPart = part.substring(0, part.lastIndexOf(yearMatch[1])).replace(/,\s*$/, '').trim();
        if (authorPart) {
          results.push({
            text: m[0],
            start: m.index,
            end: m.index + m[0].length,
            type: 'parenthetical',
            authors: authorPart,
            year: yearMatch[1],
          });
        }
      }
    }
  }

  // Find narrative citations: Author (Year)
  const narrRe = new RegExp(NARRATIVE_RE.source, 'g');
  while ((m = narrRe.exec(text)) !== null) {
    // Skip if this overlaps with a parenthetical already found
    const overlap = results.some(function (r) {
      return m.index < r.end && (m.index + m[0].length) > r.start;
    });
    if (!overlap) {
      results.push({
        text: m[0],
        start: m.index,
        end: m.index + m[0].length,
        type: 'narrative',
        authors: m[1],
        year: m[2],
      });
    }
  }

  // Sort by position
  results.sort(function (a, b) { return a.start - b.start; });

  return results;
}

// ============================================================================
// CITATIONS CLASS
// ============================================================================

class Citations {

  /**
   * Find all citation patterns in the document text.
   * Returns structured data about each citation found, without requiring
   * any network access.
   *
   * @param {object} ws - Workspace with ws.docXml
   * @returns {Array<{text: string, paragraph: number, pattern: string, authors: string, year: string}>}
   */
  static list(ws) {
    const docXml = ws.docXml;
    const paragraphs = xml.findParagraphs(docXml);
    const results = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      const plainText = xml.extractTextDecoded(para.xml);
      if (!plainText) continue;

      const citations = _findCitationPatternsInText(plainText);
      for (const cite of citations) {
        results.push({
          text: cite.text,
          paragraph: i,
          pattern: cite.type,
          authors: cite.authors,
          year: cite.year,
        });
      }
    }

    return results;
  }

  /**
   * Inject Zotero citation field codes into the document.
   *
   * Finds plain-text citation patterns, queries the Zotero Web API to match
   * them to library items, and replaces them with OOXML ZOTERO_CITATION field
   * codes that the Zotero Word plugin can manage.
   *
   * @param {object} ws - Workspace with ws.docXml get/set
   * @param {object} options
   * @param {string} options.zoteroApiKey - Zotero API key
   * @param {string} options.zoteroUserId - Zotero user ID
   * @param {string} [options.collectionId] - Limit matching to this collection
   * @param {boolean} [options.bibliography] - Insert bibliography field after "References" heading
   * @returns {Promise<{found: number, matched: number, injected: number, unmatched: Array<string>}>}
   */
  static async inject(ws, options) {
    if (!options || !options.zoteroApiKey || !options.zoteroUserId) {
      throw new Error('Citations.inject requires options.zoteroApiKey and options.zoteroUserId');
    }

    // Reset citation counter for this session
    _citationCounter = 0;

    const userId = options.zoteroUserId;
    const apiKey = options.zoteroApiKey;

    // 1. Find all citation patterns in the document
    const found = Citations.list(ws);
    if (found.length === 0) {
      return { found: 0, matched: 0, injected: 0, unmatched: [] };
    }

    // 2. Fetch Zotero items (from collection or by searching)
    let zoteroItems = [];
    if (options.collectionId) {
      zoteroItems = await _fetchCollectionItems(userId, options.collectionId, apiKey);
    }

    // Build a lookup: lowercase "lastname year" -> zotero item
    const zoteroMap = new Map();
    for (const item of zoteroItems) {
      const d = item.data;
      if (!d.creators || d.creators.length === 0) continue;
      const dateStr = d.date || d.dateEnacted || d.dateDecided || '';
      const yearMatch = dateStr.match(/(\d{4})/);
      if (!yearMatch) continue;
      const year = yearMatch[1];

      for (const creator of d.creators) {
        if (creator.creatorType !== 'author') continue;
        const name = (creator.lastName || creator.name || '').toLowerCase();
        if (name) {
          const key = name + ' ' + year;
          if (!zoteroMap.has(key)) zoteroMap.set(key, item);
        }
      }
    }

    // 3. Match each citation to a Zotero item
    const matched = [];
    const unmatched = [];

    for (const cite of found) {
      // Extract the first author's last name from the citation text
      const authorStr = cite.authors;
      const firstName = authorStr.split(/\s*(?:,|and|&|et\s+al\.)\s*/)[0].trim();
      // Handle "Dommett and Power" -> "Dommett"
      const lastName = firstName.split(/\s+/).pop();
      const lookupKey = lastName.toLowerCase() + ' ' + cite.year;

      let zItem = zoteroMap.get(lookupKey);

      // If not found in collection, try searching the library
      if (!zItem && !options.collectionId) {
        const searchResults = await _searchItems(userId, lastName, cite.year, apiKey);
        if (searchResults.length > 0) zItem = searchResults[0];
      }

      if (zItem) {
        matched.push({
          cite: cite,
          zoteroItem: zItem,
          cslData: _buildCslItemData(zItem),
        });
      } else {
        unmatched.push(cite.text);
      }
    }

    // 4. Replace citation patterns in the document XML with field codes
    let docXml = ws.docXml;
    let injectedCount = 0;

    // Process replacements within <w:r> elements that contain <w:t>
    docXml = docXml.replace(
      /<w:r>(<w:rPr>[^<]*(?:<[^/][^<]*)*<\/w:rPr>)?<w:t([^>]*)>([^<]*(?:<(?!\/w:t>)[^<]*)*)<\/w:t><\/w:r>/g,
      function (fullMatch, rPrContent, tAttrs, textContent) {
        const rPrXml = rPrContent || '';
        const decodedText = xml.decodeXml(textContent);

        // Check if this run's text contains any of our matched citations
        const segments = _splitTextByCitations(textContent, decodedText, matched, rPrXml, userId);

        if (!segments) return fullMatch;

        // Build replacement XML from segments
        let replacement = '';
        for (const seg of segments) {
          if (seg.type === 'text') {
            if (seg.text.length > 0) {
              replacement += '<w:r>' + rPrXml + '<w:t xml:space="preserve">' + seg.text + '</w:t></w:r>';
            }
          } else if (seg.type === 'citation') {
            replacement += seg.xml;
            injectedCount++;
          }
        }

        return replacement;
      }
    );

    // 5. Optionally insert bibliography field
    if (options.bibliography !== false) {
      docXml = Citations._insertBibliography(docXml);
    }

    // Write back
    ws.docXml = docXml;

    return {
      found: found.length,
      matched: matched.length,
      injected: injectedCount,
      unmatched: unmatched,
    };
  }

  // ── Internal helpers ─────────────────────────────────────────────────────

  /**
   * Insert a ZOTERO_BIBL field code after the "References" heading.
   * @param {string} docXml - Document XML
   * @returns {string} Modified document XML
   * @private
   */
  static _insertBibliography(docXml) {
    const refsPattern = /<w:p[^>]*>(?:[^<]*<[^>]*>)*?[^<]*<w:t[^>]*>References<\/w:t>[^<]*(?:<[^>]*>[^<]*)*?<\/w:p>/;
    const refsMatch = refsPattern.exec(docXml);

    if (!refsMatch) return docXml;

    const refsHeadingEnd = refsMatch.index + refsMatch[0].length;
    const defaultRPr = '<w:rPr><w:rFonts w:hint="default" w:ascii="Times New Roman" ' +
      'w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman" />' +
      '<w:sz w:val="24" /></w:rPr>';

    const afterRefs = docXml.substring(refsHeadingEnd);
    const bodyEndIdx = afterRefs.indexOf('</w:body>');

    const bibParagraph =
      '<w:p><w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>' +
      _buildBibliographyFieldXml(defaultRPr) +
      '</w:p>';

    if (bodyEndIdx !== -1) {
      return docXml.substring(0, refsHeadingEnd) + bibParagraph + afterRefs.substring(bodyEndIdx);
    }
    return docXml.substring(0, refsHeadingEnd) + bibParagraph;
  }
}

// ============================================================================
// INTERNAL: SPLIT TEXT BY CITATIONS
// ============================================================================

/**
 * Split a run's text content into segments of plain text and citation field codes.
 * Returns null if no citations are found in this text.
 *
 * @param {string} xmlText - Raw text as it appears in the XML (XML-escaped)
 * @param {string} decodedText - Same text with XML entities decoded
 * @param {Array} matchedCitations - Array of {cite, zoteroItem, cslData} objects
 * @param {string} rPrXml - Run properties XML
 * @param {string} userId - Zotero user ID
 * @returns {Array|null} Segments array or null if no matches
 * @private
 */
function _splitTextByCitations(xmlText, decodedText, matchedCitations, rPrXml, userId) {
  const matches = [];

  for (const entry of matchedCitations) {
    const cite = entry.cite;
    const searchText = cite.text;

    // Search in decoded text for the citation
    let searchFrom = 0;
    while (true) {
      const idx = decodedText.indexOf(searchText, searchFrom);
      if (idx === -1) break;

      // Find the corresponding position in the XML-encoded text
      const xmlStart = _decodedToXmlOffset(xmlText, decodedText, idx);
      const xmlEnd = _decodedToXmlOffset(xmlText, decodedText, idx + searchText.length);

      matches.push({
        xmlStart: xmlStart,
        xmlEnd: xmlEnd,
        entry: entry,
      });

      searchFrom = idx + searchText.length;
    }
  }

  if (matches.length === 0) return null;

  // Sort by position, prefer longer matches at same position
  matches.sort(function (a, b) {
    if (a.xmlStart !== b.xmlStart) return a.xmlStart - b.xmlStart;
    return b.xmlEnd - a.xmlEnd;
  });

  // Remove overlapping matches
  const filtered = [matches[0]];
  for (let i = 1; i < matches.length; i++) {
    if (matches[i].xmlStart >= filtered[filtered.length - 1].xmlEnd) {
      filtered.push(matches[i]);
    }
  }

  // Build segments
  const segments = [];
  let pos = 0;

  for (const m of filtered) {
    const entry = m.entry;
    const cite = entry.cite;

    // Text before this citation
    if (m.xmlStart > pos) {
      segments.push({ type: 'text', text: xmlText.substring(pos, m.xmlStart) });
    }

    // Build the citation field code
    const cslItem = { zoteroKey: entry.zoteroItem.data.key, cslData: entry.cslData };
    const isNarrative = cite.pattern === 'narrative';

    const citJson = _buildZoteroCitationJson([cslItem], userId, isNarrative);

    if (isNarrative) {
      // Narrative: keep author names as plain text, field-code the year
      const yearPart = '(' + cite.year + ')';
      const authorPart = xmlText.substring(m.xmlStart, m.xmlEnd - xml.escapeXml(yearPart).length).replace(/\s+$/, ' ');
      segments.push({ type: 'text', text: authorPart });
      segments.push({
        type: 'citation',
        xml: _buildCitationFieldXml(citJson, yearPart, rPrXml),
      });
    } else {
      segments.push({
        type: 'citation',
        xml: _buildCitationFieldXml(citJson, cite.text, rPrXml),
      });
    }

    pos = m.xmlEnd;
  }

  // Remaining text
  if (pos < xmlText.length) {
    segments.push({ type: 'text', text: xmlText.substring(pos) });
  }

  return segments;
}

/**
 * Map a character offset in decoded text to the corresponding offset
 * in XML-escaped text. Handles &amp; &lt; &gt; &quot; &apos; entities.
 *
 * @param {string} xmlText - XML-escaped text
 * @param {string} decodedText - Decoded plain text
 * @param {number} decodedOffset - Character offset in decoded text
 * @returns {number} Corresponding offset in XML text
 * @private
 */
function _decodedToXmlOffset(xmlText, decodedText, decodedOffset) {
  let xmlPos = 0;
  let decodedPos = 0;

  while (decodedPos < decodedOffset && xmlPos < xmlText.length) {
    if (xmlText[xmlPos] === '&') {
      // Find the end of the entity
      const semiIdx = xmlText.indexOf(';', xmlPos);
      if (semiIdx !== -1) {
        xmlPos = semiIdx + 1;
        decodedPos++;
        continue;
      }
    }
    xmlPos++;
    decodedPos++;
  }

  return xmlPos;
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = { Citations };

// Also export internals for testing
module.exports._findCitationPatternsInText = _findCitationPatternsInText;
module.exports._buildCslItemData = _buildCslItemData;
module.exports._buildZoteroCitationJson = _buildZoteroCitationJson;
module.exports._buildCitationFieldXml = _buildCitationFieldXml;
module.exports._buildBibliographyFieldXml = _buildBibliographyFieldXml;
module.exports._zoteroTypeToCsl = _zoteroTypeToCsl;
module.exports._decodedToXmlOffset = _decodedToXmlOffset;
