'use strict';
const xml = require('./xml');
var EMU = 914400;

class FigureHandle {
  constructor(engine, fi) { this._engine = engine; this._figureIndex = fi; }
  dimensions() { var f = this._locate(); var m = f.paraXml.match(/<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/); if (!m) return { width: 0, height: 0 }; return { width: parseInt(m[1], 10) / EMU, height: parseInt(m[2], 10) / EMU }; }
  resize(w, h) { var ws = this._gw(); var f = this._locate(); var p = f.paraXml; var cx = Math.round(w * EMU); var cy = Math.round(h * EMU);
    p = p.replace(/<wp:extent\s+cx="\d+"\s+cy="\d+"/g, '<wp:extent cx="' + cx + '" cy="' + cy + '"');
    p = p.replace(/(<a:xfrm>[\s\S]*?<a:ext\s+)cx="\d+"\s+cy="\d+"/, '$1cx="' + cx + '" cy="' + cy + '"');
    ws.docXml = ws.docXml.slice(0, f.start) + p + ws.docXml.slice(f.end); return this; }
  altText() { var f = this._locate(); var m = f.paraXml.match(/<wp:docPr\s+[^>]*descr="([^"]*)"/); return m ? xml.decodeXml(m[1]) : ''; }
  setAltText(t) { var ws = this._gw(); var f = this._locate(); var p = f.paraXml; var e = xml.escapeXml(t);
    if (p.includes('descr="')) p = p.replace(/(<wp:docPr\s+[^>]*?)descr="[^"]*"/, '$1descr="' + e + '"');
    else p = p.replace(/(<wp:docPr\s+)/, '$1descr="' + e + '" ');
    ws.docXml = ws.docXml.slice(0, f.start) + p + ws.docXml.slice(f.end); return this; }
  caption() { var ws = this._gw(); var f = this._locate(); var a = ws.docXml.slice(f.end); var m = a.match(/<w:p[\s>][\s\S]*?<\/w:p>/);
    if (!m) return ''; var t = xml.decodeXml(xml.extractText(m[0])); return /^(Figure|Fig\.)\s+\d/i.test(t.trim()) ? t : ''; }
  setCaption(t) { var ws = this._gw(); var f = this._locate(); var d = ws.docXml; var a = d.slice(f.end); var m = a.match(/<w:p[\s>][\s\S]*?<\/w:p>/);
    if (!m) throw new Error('No para after figure'); var s = f.end + m.index; var e = s + m[0].length;
    var nt = xml.extractText(m[0]); if (!/^(Figure|Fig\.)\s+\d/i.test(xml.decodeXml(nt).trim())) throw new Error('Not a caption');
    var nc = '<w:p><w:pPr><w:jc w:val="center"/><w:rPr/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">' + xml.escapeXml(t) + '</w:t></w:r></w:p>';
    ws.docXml = d.slice(0, s) + nc + d.slice(e); return this; }
  position() { var f = this._locate(); return f.paraXml.includes('<wp:inline') ? 'inline' : f.paraXml.includes('<wp:anchor') ? 'floating' : 'inline'; }
  mediaPath() { var ws = this._gw(); var f = this._locate(); var r = this._eR(f.paraXml); if (!r) return '';
    var re = new RegExp('Id="' + r + '"[^>]*Target="([^"]+)"'); var m = re.exec(ws.relsXml); return m ? 'word/' + m[1] : ''; }
  rId() { var f = this._locate(); return this._eR(f.paraXml) || ''; }
  _gw() { if (!this._engine._workspace) throw new Error('Not opened'); return this._engine._workspace; }
  _locate() { var ws = this._gw(); var d = ws.docXml; var pp = xml.findParagraphs(d); var fc = 0;
    for (var i = 0; i < pp.length; i++) { var p = pp[i]; if ((p.xml.includes('<w:drawing>') || p.xml.includes('<w:drawing ') || p.xml.includes('w:drawing>')) && p.xml.includes('a:blip')) {
      if (fc === this._figureIndex) return { paraXml: p.xml, start: p.start, end: p.end }; fc++; } }
    throw new Error('Figure index ' + this._figureIndex + ' out of range (' + fc + ' figures)'); }
  _eR(p) { var m = p.match(/a:blip[^>]+r:embed="(rId\d+)"/); return m ? m[1] : null; }
}
module.exports = { FigureHandle };
