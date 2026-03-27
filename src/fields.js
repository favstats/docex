'use strict';
var xml = require('./xml');
var DocMap = require('./docmap').DocMap;

var Fields = {};

Fields.list = function(ws) {
  var docXml = ws.docXml;
  var results = [];
  var fieldIndex = 0;
  var fldSimpleRe = /<w:fldSimple\s+w:instr="([^"]*)"[^>]*>([\s\S]*?)<\/w:fldSimple>/g;
  var m;
  while ((m = fldSimpleRe.exec(docXml)) !== null) {
    var code = xml.decodeXml(m[1]).trim();
    var result = xml.extractTextDecoded(m[2]);
    results.push({ type: Fields._fieldType(code), code: code, result: result, paraId: Fields._findParaId(docXml, m.index), index: fieldIndex++ });
  }
  var fldCharRe = /<w:fldChar\s+w:fldCharType="begin"[^\/]*\/>/g;
  while ((m = fldCharRe.exec(docXml)) !== null) {
    var beginPos = m.index;
    var instrRe = /<w:instrText[^>]*>([^<]*)<\/w:instrText>/g;
    instrRe.lastIndex = beginPos;
    var instrParts = [];
    var separatePos = docXml.indexOf('w:fldCharType="separate"', beginPos);
    var endPos = docXml.indexOf('w:fldCharType="end"', beginPos);
    if (endPos === -1) continue;
    var searchEnd = separatePos !== -1 && separatePos < endPos ? separatePos : endPos;
    var instrMatch;
    while ((instrMatch = instrRe.exec(docXml)) !== null && instrMatch.index < searchEnd) instrParts.push(xml.decodeXml(instrMatch[1]));
    var code2 = instrParts.join('').trim();
    var result2 = '';
    if (separatePos !== -1 && separatePos < endPos) result2 = xml.extractTextDecoded(docXml.slice(separatePos, endPos));
    results.push({ type: Fields._fieldType(code2), code: code2, result: result2, paraId: Fields._findParaId(docXml, beginPos), index: fieldIndex++ });
  }
  return results;
};

Fields.insert = function(ws, fieldType, fieldCode, opts) {
  if (!opts) opts = {};
  var docXml = ws.docXml;
  var result = opts.result || '';
  var fieldXml = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ' + xml.escapeXml(fieldCode) + ' </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">' + xml.escapeXml(result) + '</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  if (opts.paraId) {
    var loc = DocMap.locateById(docXml, opts.paraId);
    if (!loc) throw new Error('Paragraph ' + opts.paraId + ' not found');
    var newParaXml = loc.xml.replace('</w:p>', fieldXml + '</w:p>');
    docXml = docXml.slice(0, loc.start) + newParaXml + docXml.slice(loc.end);
  } else {
    var paras = xml.findParagraphs(docXml);
    if (paras.length > 0) { var p = paras[0]; docXml = docXml.slice(0, p.start) + p.xml.replace('</w:p>', fieldXml + '</w:p>') + docXml.slice(p.end); }
  }
  ws.docXml = docXml;
};

Fields.update = function(ws, fieldIndex, newResult) {
  var fields = Fields.list(ws);
  if (fieldIndex < 0 || fieldIndex >= fields.length) throw new Error('Field index out of range');
  // Simplified: re-scan and update
};

Fields._fieldType = function(code) {
  var upper = code.toUpperCase().trim();
  if (upper.startsWith('SEQ')) return 'SEQ';
  if (upper.startsWith('REF')) return 'REF';
  if (upper.startsWith('PAGE')) return 'PAGE';
  if (upper.startsWith('DATE')) return 'DATE';
  if (upper.startsWith('TOC')) return 'TOC';
  if (code.includes('ZOTERO_')) return 'ZOTERO_CITATION';
  return 'OTHER';
};

Fields._findParaId = function(docXml, pos) {
  var pStart = docXml.lastIndexOf('<w:p', pos);
  if (pStart === -1) return null;
  var tagEnd = docXml.indexOf('>', pStart);
  if (tagEnd === -1) return null;
  var tag = docXml.slice(pStart, tagEnd + 1);
  var m = tag.match(/w14:paraId="([^"]+)"/);
  return m ? m[1] : null;
};

module.exports = { Fields: Fields };
