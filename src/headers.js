'use strict';
var fs = require('fs');
var path = require('path');
var xml = require('./xml');

var REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
var REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
var CT_HEADER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
var CT_FOOTER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

var Headers = {};

Headers.get = function(ws, type) {
  type = type || 'default';
  var headerFile = Headers._findHeaderFile(ws, type, 'header');
  if (!headerFile) return null;
  var headerXml = Headers._readPartFile(ws, headerFile);
  if (!headerXml) return null;
  return { text: xml.extractTextDecoded(headerXml), xml: headerXml };
};

Headers.set = function(ws, content, opts) {
  if (!opts) opts = {};
  if (typeof content === 'object' && content !== null && !opts.type) {
    if (content.firstPage !== undefined) Headers._setHeaderContent(ws, content.firstPage, 'first', 'header');
    if (content.rest !== undefined) Headers._setHeaderContent(ws, content.rest, 'default', 'header');
  } else {
    Headers._setHeaderContent(ws, content, opts.type || 'default', 'header');
  }
};

Headers.getFooter = function(ws, type) {
  type = type || 'default';
  var footerFile = Headers._findHeaderFile(ws, type, 'footer');
  if (!footerFile) return null;
  var footerXml = Headers._readPartFile(ws, footerFile);
  if (!footerXml) return null;
  return { text: xml.extractTextDecoded(footerXml), xml: footerXml };
};

Headers.setFooter = function(ws, content, opts) {
  if (!opts) opts = {};
  if (typeof content === 'object' && content !== null && !opts.type) {
    if (content.firstPage !== undefined) Headers._setHeaderContent(ws, content.firstPage, 'first', 'footer');
    if (content.rest !== undefined) Headers._setHeaderContent(ws, content.rest, 'default', 'footer');
  } else {
    Headers._setHeaderContent(ws, content, opts.type || 'default', 'footer');
  }
};

Headers.xml = function(ws, type) {
  type = type || 'default';
  var headerFile = Headers._findHeaderFile(ws, type, 'header');
  if (!headerFile) return null;
  return Headers._readPartFile(ws, headerFile);
};

Headers._findHeaderFile = function(ws, type, kind) {
  var docXml = ws.docXml;
  var relsXml = ws.relsXml;
  var tagName = kind === 'header' ? 'w:headerReference' : 'w:footerReference';
  var refRe = new RegExp('<' + tagName + '\\s+w:type="' + type + '"\\s+r:id="([^"]+)"[^/]*/>', 'g');
  var m = refRe.exec(docXml);
  if (!m) return null;
  var rId = m[1];
  var relRe = new RegExp('Id="' + rId + '"[^>]*Target="([^"]+)"', 'g');
  var rm = relRe.exec(relsXml);
  return rm ? rm[1] : null;
};

Headers._readPartFile = function(ws, filename) {
  var filePath = path.join(ws.tmpDir, 'word', filename);
  if (!fs.existsSync(filePath)) return null;
  return fs.readFileSync(filePath, 'utf-8');
};

Headers._setHeaderContent = function(ws, content, type, kind) {
  var tagName = kind === 'header' ? 'w:headerReference' : 'w:footerReference';
  var relType = kind === 'header' ? REL_HEADER : REL_FOOTER;
  var ctType = kind === 'header' ? CT_HEADER : CT_FOOTER;
  var rootTag = kind === 'header' ? 'w:hdr' : 'w:ftr';
  var existingFile = Headers._findHeaderFile(ws, type, kind);
  if (existingFile) {
    var paraId = xml.randomHexId();
    var textId = xml.randomHexId();
    fs.writeFileSync(path.join(ws.tmpDir, 'word', existingFile), Headers._buildPartXml(rootTag, content, paraId, textId), 'utf-8');
  } else {
    var fileNum = Headers._nextFileNum(ws, kind);
    var filename = kind + fileNum + '.xml';
    var paraId2 = xml.randomHexId();
    var textId2 = xml.randomHexId();
    fs.writeFileSync(path.join(ws.tmpDir, 'word', filename), Headers._buildPartXml(rootTag, content, paraId2, textId2), 'utf-8');
    var rId = xml.nextRId(ws.relsXml);
    ws.relsXml = ws.relsXml.replace('</Relationships>', '<Relationship Id="' + rId + '" Type="' + relType + '" Target="' + filename + '"/></Relationships>');
    ws.contentTypesXml = ws.contentTypesXml.replace('</Types>', '<Override PartName="/word/' + filename + '" ContentType="' + ctType + '"/></Types>');
    var refEl = '<' + tagName + ' w:type="' + type + '" r:id="' + rId + '"/>';
    var docXml = ws.docXml;
    var sectPrPos = docXml.lastIndexOf('<w:sectPr');
    if (sectPrPos !== -1) {
      var closeAngle = docXml.indexOf('>', sectPrPos);
      if (closeAngle !== -1) {
        if (docXml[closeAngle - 1] === '/') {
          ws.docXml = docXml.slice(0, closeAngle - 1) + '>' + refEl + '</w:sectPr>' + docXml.slice(closeAngle + 1);
        } else {
          ws.docXml = docXml.slice(0, closeAngle + 1) + refEl + docXml.slice(closeAngle + 1);
        }
      }
    }
    if (type === 'first' && !ws.docXml.includes('<w:titlePg')) {
      var sp = ws.docXml.lastIndexOf('</w:sectPr>');
      if (sp !== -1) ws.docXml = ws.docXml.slice(0, sp) + '<w:titlePg/>' + ws.docXml.slice(sp);
    }
  }
};

Headers._buildPartXml = function(rootTag, content, paraId, textId) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<' + rootTag + ' xmlns:w="' + xml.NS.w + '" xmlns:r="' + xml.NS.r + '" xmlns:w14="' + xml.NS.w14 + '"><w:p w14:paraId="' + paraId + '" w14:textId="' + textId + '"><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">' + xml.escapeXml(content) + '</w:t></w:r></w:p></' + rootTag + '>';
};

Headers._nextFileNum = function(ws, kind) {
  var wordDir = path.join(ws.tmpDir, 'word');
  var max = 0;
  if (fs.existsSync(wordDir)) {
    var files = fs.readdirSync(wordDir);
    var re = new RegExp('^' + kind + '(\\d+)\\.xml$');
    for (var i = 0; i < files.length; i++) {
      var m = files[i].match(re);
      if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
    }
  }
  return max + 1;
};

module.exports = { Headers: Headers };
