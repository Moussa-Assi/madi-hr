// Minimal .xlsx writer for the fake export spreadsheets (values, bold, fills, colours, number formats, widths, frozen panes, merges).
import { zipSync, strToU8 } from 'fflate';

const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const argb = c => { c = String(c || '').replace('#', ''); if (c.length === 3) c = c.split('').map(x => x + x).join(''); return 'FF' + c.toUpperCase(); };

export function toXlsx(ss) {
  const sheets = ss.sheets.filter(s => s.getLastRow() > 0 || ss.sheets.length === 1);
  const fonts = ['<font><sz val="10"/><name val="Arial"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const numFmts = [], xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], xfKey = { '': 0 };
  const fontIdx = {}, fillIdx = {}, fmtIdx = {};
  const styleId = st => {
    if (!st) return 0;
    const key = JSON.stringify(st); if (key in xfKey) return xfKey[key];
    const fk = [st.bold ? 1 : 0, st.color || '', st.size || 10].join('|');
    if (!(fk in fontIdx)) { fontIdx[fk] = fonts.length; fonts.push('<font>' + (st.bold ? '<b/>' : '') + '<sz val="' + (st.size || 10) + '"/>' + (st.color ? '<color rgb="' + argb(st.color) + '"/>' : '') + '<name val="Arial"/></font>'); }
    let fill = 0;
    if (st.bg && !/^#?f{3}(f{3})?$/i.test(st.bg)) { if (!(st.bg in fillIdx)) { fillIdx[st.bg] = fills.length; fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + argb(st.bg) + '"/><bgColor indexed="64"/></patternFill></fill>'); } fill = fillIdx[st.bg]; }
    let nf = 0;
    if (st.numFmt && st.numFmt !== '@') { if (!(st.numFmt in fmtIdx)) { fmtIdx[st.numFmt] = 164 + numFmts.length; numFmts.push('<numFmt numFmtId="' + fmtIdx[st.numFmt] + '" formatCode="' + esc(st.numFmt) + '"/>'); } nf = fmtIdx[st.numFmt]; }
    else if (st.numFmt === '@') nf = 49;
    const al = st.align || st.valign || st.wrap ? '<alignment' + (st.align ? ' horizontal="' + st.align + '"' : '') + (st.valign ? ' vertical="' + (st.valign === 'middle' ? 'center' : st.valign) + '"' : '') + (st.wrap ? ' wrapText="1"' : '') + '/>' : '';
    const id = xfs.length;
    xfs.push('<xf numFmtId="' + nf + '" fontId="' + fontIdx[fk] + '" fillId="' + fill + '" borderId="' + (st.border ? 1 : 0) + '" xfId="0"' + (nf ? ' applyNumberFormat="1"' : '') + ' applyFont="1"' + (fill ? ' applyFill="1"' : '') + (st.border ? ' applyBorder="1"' : '') + (al ? ' applyAlignment="1">' + al + '</xf>' : '/>'));
    xfKey[key] = id; return id;
  };
  const files = {};
  const sheetXml = sheets.map((s, si) => {
    const nr = s.getLastRow(), nc = s.getLastColumn();
    let rows = '';
    for (let r = 1; r <= nr; r++) {
      const v = s.rows[r - 1] ? s.rows[r - 1].v : [];
      let cells = '';
      for (let c = 1; c <= Math.max(nc, 1); c++) {
        const x = v[c - 1], st = s.styles[r + ',' + c], sid = styleId(st);
        const ref = colName(c) + r;
        if (x === '' || x === null || x === undefined) { if (sid) cells += '<c r="' + ref + '" s="' + sid + '"/>'; continue; }
        if (typeof x === 'number' && isFinite(x)) cells += '<c r="' + ref + '" s="' + sid + '"><v>' + x + '</v></c>';
        else if (typeof x === 'boolean') cells += '<c r="' + ref + '" s="' + sid + '" t="b"><v>' + (x ? 1 : 0) + '</v></c>';
        else cells += '<c r="' + ref + '" s="' + sid + '" t="inlineStr"><is><t xml:space="preserve">' + esc(x && x.getTime ? x.toISOString().slice(0, 10) : x) + '</t></is></c>';
      }
      rows += '<row r="' + r + '"' + (s.heights[r] ? ' ht="' + (s.heights[r] * 0.75) + '" customHeight="1"' : '') + '>' + cells + '</row>';
    }
    const fr = s.frozen[0], fc = s.frozen[1];
    const pane = fr || fc ? '<sheetViews><sheetView workbookViewId="0"><pane' + (fc ? ' xSplit="' + fc + '"' : '') + (fr ? ' ySplit="' + fr + '"' : '') + ' topLeftCell="' + colName(fc + 1) + (fr + 1) + '" activePane="' + (fr && fc ? 'bottomRight' : fr ? 'bottomLeft' : 'topRight') + '" state="frozen"/></sheetView></sheetViews>' : '';
    const ws = Object.keys(s.widths);
    const cols = ws.length ? '<cols>' + ws.map(Number).sort((a, b) => a - b).map(c => '<col min="' + c + '" max="' + c + '" width="' + Math.max(3, Math.round(s.widths[c] / 7 * 10) / 10) + '" customWidth="1"/>').join('') + '</cols>' : '';
    const merges = s.merges.length ? '<mergeCells count="' + s.merges.length + '">' + s.merges.map(m => '<mergeCell ref="' + colName(m[1]) + m[0] + ':' + colName(m[1] + m[3] - 1) + (m[0] + m[2] - 1) + '"/>').join('') + '</mergeCells>' : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + pane + cols + '<sheetData>' + rows + '</sheetData>' + merges + '</worksheet>';
  });
  sheetXml.forEach((x, i) => { files['xl/worksheets/sheet' + (i + 1) + '.xml'] = strToU8(x); });
  const names = {}; const sname = n => { let b = String(n || 'Sheet').replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 31) || 'Sheet', k = b, i = 2; while (names[k.toLowerCase()]) k = b.slice(0, 28) + ' ' + i++; names[k.toLowerCase()] = 1; return k; };
  files['[Content_Types].xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map((s, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>');
  files['_rels/.rels'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  files['xl/workbook.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets.map((s, i) => '<sheet name="' + esc(sname(s.name)) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets></workbook>');
  files['xl/_rels/workbook.xml.rels'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') + '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  files['xl/styles.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + (numFmts.length ? '<numFmts count="' + numFmts.length + '">' + numFmts.join('') + '</numFmts>' : '') + '<fonts count="' + fonts.length + '">' + fonts.join('') + '</fonts><fills count="' + fills.length + '">' + fills.join('') + '</fills><borders count="2"><border/><border><left style="thin"><color rgb="FFBBBBBB"/></left><right style="thin"><color rgb="FFBBBBBB"/></right><top style="thin"><color rgb="FFBBBBBB"/></top><bottom style="thin"><color rgb="FFBBBBBB"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs></styleSheet>');
  return zipSync(files);
}
