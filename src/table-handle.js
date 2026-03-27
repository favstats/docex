'use strict';
const xml = require('./xml');

class TableHandle {
  constructor(engine, tableIndex) { this._engine = engine; this._tableIndex = tableIndex; }
  dimensions() {
    var tbl = this._locate(); var rows = TableHandle._findRows(tbl.xml);
    var cols = rows.length > 0 ? TableHandle._findCells(rows[0]).length : 0;
    return { rows: rows.length, cols: cols };
  }
  cell(row, col) { return new CellHandle(this._engine, this._tableIndex, row, col); }
  addRow(data) {
    var ws = this._getWorkspace(); var tbl = this._locate();
    var rows = TableHandle._findRows(tbl.xml);
    var numCols = rows.length > 0 ? TableHandle._findCells(rows[0]).length : data.length;
    var colWidth = Math.floor(9360 / numCols); var cellsXml = '';
    for (var c = 0; c < numCols; c++) {
      var t = c < data.length ? String(data[c]) : '';
      cellsXml += '<w:tc><w:tcPr><w:tcW w:w="' + colWidth + '" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">' + xml.escapeXml(t) + '</w:t></w:r></w:p></w:tc>';
    }
    var tblEnd = tbl.xml.lastIndexOf('</w:tbl>');
    var n = tbl.xml.slice(0, tblEnd) + '<w:tr>' + cellsXml + '</w:tr>' + tbl.xml.slice(tblEnd);
    ws.docXml = ws.docXml.slice(0, tbl.start) + n + ws.docXml.slice(tbl.end); return this;
  }
  removeRow(idx) {
    var ws = this._getWorkspace(); var tbl = this._locate(); var rows = TableHandle._findRows(tbl.xml);
    if (idx < 0 || idx >= rows.length) throw new Error('Row index out of range');
    var row = rows[idx]; var n = tbl.xml.slice(0, row.start) + tbl.xml.slice(row.end);
    ws.docXml = ws.docXml.slice(0, tbl.start) + n + ws.docXml.slice(tbl.end); return this;
  }
  addColumn(header) {
    var ws = this._getWorkspace(); var tbl = this._locate(); var x = tbl.xml;
    var rows = TableHandle._findRows(x); var nc = rows.length > 0 ? TableHandle._findCells(rows[0]).length : 0;
    var cw = Math.floor(9360 / (nc + 1));
    var gi = x.indexOf('</w:tblGrid>'); if (gi !== -1) x = x.slice(0, gi) + '<w:gridCol w:w="' + cw + '"/>' + x.slice(gi);
    var rr = TableHandle._findRows(x);
    for (var r = rr.length - 1; r >= 0; r--) {
      var row = rr[r]; var ct = r === 0 ? header : '';
      var nc2 = '<w:tc><w:tcPr><w:tcW w:w="' + cw + '" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">' + xml.escapeXml(ct) + '</w:t></w:r></w:p></w:tc>';
      var ei = row.xml.lastIndexOf('</w:tr>'); var nr = row.xml.slice(0, ei) + nc2 + row.xml.slice(ei);
      x = x.slice(0, row.start) + nr + x.slice(row.end);
    }
    ws.docXml = ws.docXml.slice(0, tbl.start) + x + ws.docXml.slice(tbl.end); return this;
  }
  merge(r1, c1, r2, c2) { return this; }
  data() {
    var tbl = this._locate(); var rows = TableHandle._findRows(tbl.xml); var result = [];
    for (var ri = 0; ri < rows.length; ri++) {
      var cells = TableHandle._findCells(rows[ri].xml); var rd = [];
      for (var ci = 0; ci < cells.length; ci++) rd.push(xml.decodeXml(xml.extractText(cells[ci].xml)));
      result.push(rd);
    }
    return result;
  }
  _getWorkspace() { if (!this._engine._workspace) throw new Error('Not opened'); return this._engine._workspace; }
  _locate() { return this._locateFromXml(this._getWorkspace().docXml); }
  _locateFromXml(d) { var t = TableHandle.findAllTables(d); if (this._tableIndex < 0 || this._tableIndex >= t.length) throw new Error('Table index out of range'); return t[this._tableIndex]; }
  static findAllTables(d) {
    var t = []; var re = /<w:tbl[\s>]/g; var m;
    while ((m = re.exec(d)) !== null) { var s = m.index; var dp = 1; var p = s + m[0].length;
      while (dp > 0 && p < d.length) { var no = d.indexOf('<w:tbl', p); var nc = d.indexOf('</w:tbl>', p); if (nc === -1) break;
        if (no !== -1 && no < nc) { dp++; p = no + 6; } else { dp--; if (dp === 0) t.push({ xml: d.slice(s, nc + 8), start: s, end: nc + 8 }); p = nc + 8; } } }
    return t;
  }
  static _findRows(x) { var r = []; var re = /<w:tr[\s>]/g; var m; while ((m = re.exec(x)) !== null) { var s = m.index; var ci = x.indexOf('</w:tr>', s); if (ci === -1) continue; r.push({ xml: x.slice(s, ci + 7), start: s, end: ci + 7 }); } return r; }
  static _findCells(x) {
    var c = []; var re = /<w:tc[\s>]/g; var m;
    while ((m = re.exec(x)) !== null) { var s = m.index; var dp = 1; var p = s + m[0].length;
      while (dp > 0 && p < x.length) { var no = x.indexOf('<w:tc', p); var nc = x.indexOf('</w:tc>', p); if (nc === -1) break;
        if (no !== -1 && no < nc) { dp++; p = no + 5; } else { dp--; if (dp === 0) c.push({ xml: x.slice(s, nc + 7), start: s, end: nc + 7 }); p = nc + 7; } } }
    return c;
  }
}

class CellHandle {
  constructor(engine, ti, row, col) { this._engine = engine; this._tableIndex = ti; this._row = row; this._col = col; }
  text() { var c = this._locateCell(); return xml.decodeXml(xml.extractText(c.cellXml)); }
  setText(text) {
    var ws = this._gw(); var c = this._locateCell();
    var pm = c.cellXml.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/); var tp = pm ? pm[0] : '<w:tcPr></w:tcPr>';
    var nc = '<w:tc>' + tp + '<w:p><w:r><w:t xml:space="preserve">' + xml.escapeXml(text) + '</w:t></w:r></w:p></w:tc>';
    var nt = c.tblXml.slice(0, c.cas) + nc + c.tblXml.slice(c.cae);
    ws.docXml = ws.docXml.slice(0, c.ts) + nt + ws.docXml.slice(c.te); return this;
  }
  bold() {
    var ws = this._gw(); var c = this._locateCell(); var cx = c.cellXml;
    if (cx.includes('<w:rPr>')) { if (!cx.includes('<w:b/>')) cx = cx.replace(/<w:rPr>/g, '<w:rPr><w:b/>'); }
    else cx = cx.replace(/<w:r>/g, '<w:r><w:rPr><w:b/></w:rPr>');
    var nt = c.tblXml.slice(0, c.cas) + cx + c.tblXml.slice(c.cae);
    ws.docXml = ws.docXml.slice(0, c.ts) + nt + ws.docXml.slice(c.te); return this;
  }
  shading(h) { var ws = this._gw(); var c = this._locateCell(); var cx = c.cellXml; var co = h.replace(/^#/, '');
    var s = '<w:shd w:val="clear" w:color="auto" w:fill="' + co + '"/>';
    if (cx.includes('<w:tcPr>')) { if (cx.includes('<w:shd ')) cx = cx.replace(/<w:shd[^>]*\/?>/, s); else cx = cx.replace('</w:tcPr>', s + '</w:tcPr>'); }
    else cx = cx.replace('<w:tc>', '<w:tc><w:tcPr>' + s + '</w:tcPr>');
    ws.docXml = ws.docXml.slice(0, c.ts) + c.tblXml.slice(0, c.cas) + cx + c.tblXml.slice(c.cae) + ws.docXml.slice(c.te); return this; }
  alignment(a) { var ws = this._gw(); var c = this._locateCell(); var cx = c.cellXml; var j = '<w:jc w:val="' + a + '"/>';
    if (cx.includes('<w:pPr>')) { if (cx.includes('<w:jc ')) cx = cx.replace(/<w:jc[^>]*\/>/, j); else cx = cx.replace('</w:pPr>', j + '</w:pPr>'); }
    else cx = cx.replace(/<w:p>/g, '<w:p><w:pPr>' + j + '</w:pPr>');
    ws.docXml = ws.docXml.slice(0, c.ts) + c.tblXml.slice(0, c.cas) + cx + c.tblXml.slice(c.cae) + ws.docXml.slice(c.te); return this; }
  width(tw) { var ws = this._gw(); var c = this._locateCell(); var cx = c.cellXml; var w = '<w:tcW w:w="' + tw + '" w:type="dxa"/>';
    if (cx.includes('<w:tcPr>')) { if (cx.includes('<w:tcW ')) cx = cx.replace(/<w:tcW[^>]*\/>/, w); else cx = cx.replace('<w:tcPr>', '<w:tcPr>' + w); }
    else cx = cx.replace('<w:tc>', '<w:tc><w:tcPr>' + w + '</w:tcPr>');
    ws.docXml = ws.docXml.slice(0, c.ts) + c.tblXml.slice(0, c.cas) + cx + c.tblXml.slice(c.cae) + ws.docXml.slice(c.te); return this; }
  _gw() { if (!this._engine._workspace) throw new Error('Not opened'); return this._engine._workspace; }
  _locateCell() {
    var ws = this._gw(); var tables = TableHandle.findAllTables(ws.docXml);
    if (this._tableIndex < 0 || this._tableIndex >= tables.length) throw new Error('Table idx');
    var tbl = tables[this._tableIndex]; var rows = TableHandle._findRows(tbl.xml);
    if (this._row < 0 || this._row >= rows.length) throw new Error('Row idx');
    var row = rows[this._row]; var cells = TableHandle._findCells(row.xml);
    if (this._col < 0 || this._col >= cells.length) throw new Error('Col idx');
    var cell = cells[this._col];
    return { tblXml: tbl.xml, ts: tbl.start, te: tbl.end, cellXml: cell.xml, cas: row.start + cell.start, cae: row.start + cell.end };
  }
}
module.exports = { TableHandle, CellHandle };
