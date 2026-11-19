// Runs the original Apps Script logic (Code.gs) on Supabase tables.
// Sheets are kept in memory as rows; after a run, only changed rows are written back.
import { createHash, randomUUID } from 'node:crypto';
import { Buffer } from 'node:buffer';

const T = 'yyyy'; // marker
const OFF = 2 * 3600e3; // Africa/Maputo = UTC+2, no DST

/* ---------------- sheet ↔ table mapping ---------------- */
// type codes: i int, n numeric, d date, ts timestamptz, tm time, dt timestamp, da date[], j jsonb, t text, x = only kept in "extra"
const M = (t, cols, head, opt) => Object.assign({ t, cols: cols.split(',').map(s => { const [c, ty] = s.split(':'); return c === 'x' ? null : { c, ty: ty || 't' }; }), head }, opt || {});
const PAY_HEAD = ['Período', 'ID', 'Nome', 'Entidade', 'Departamento', 'Salário base', 'Dias efetivos', 'Dias trabalhados', 'Horas extra', 'Horas extra domingo', 'Valor horas extra', 'Outras rem.', 'Dias férias pagas', 'Valor férias pagas', 'Bónus', 'Subsídios', 'Salário bruto', 'Faltas', 'Desconto faltas', 'Min atraso', 'Desconto atrasos', 'INSS', 'Outros descontos', 'Saldo anterior', 'Dívida contraída', 'Pagamento adiantamento', 'Saldo atual', 'Total descontos', 'Salário líquido', 'Setor', 'Datas de faltas', 'Nota', 'Fechado em'];
const PIC_HEAD = ['ID registo', 'Data', 'Hora', 'Data/hora', 'ID', 'Nome (Hik)', 'Departamento (Hik)', 'Dispositivo', 'Leitor', 'Resultado', 'Sincronizado em'];
const PIC_COLS = 'guid,data:d,hora:tm,data_hora:dt,emp_id:i,nome_hik,dept_hik,dispositivo,leitor,resultado,sincronizado_em:ts';
export const MAP = {
  CONFIG: M('config', 'key,value,note', ['Chave', 'Valor', 'Notas'], { order: '_rid' }),
  FUNCIONARIOS: M('funcionarios', 'id:i,nome,ent,dept,loc,func,hor,isento,folga,adm:d,contrato,sal:n,outras:n,inss,nuit,banco,nib,tel,estado,saida:d,notas,inss_base:n,ef:n,regra,fim_contrato:d',
    ['ID', 'Nome', 'Entidade', 'Departamento', 'Localização', 'Função', 'Horário', 'Isento de ponto', 'Folga fixa', 'Data admissão', 'Tipo contrato', 'Salário base', 'Outras rem.', 'Nº INSS', 'NUIT', 'Banco', 'NIB', 'Telefone', 'Estado', 'Data saída', 'Notas', 'Base INSS', 'Ef.', 'Regra salarial', 'Fim de contrato'], { order: '_rid' }),
  PICAGENS: M('picagens', PIC_COLS, PIC_HEAD, { order: 'data_hora, _rid', where: "data >= (now() at time zone 'Africa/Maputo')::date - 120" }),
  PICAGENS_ARQUIVO: M('picagens', PIC_COLS, PIC_HEAD, { order: 'data_hora, _rid', where: "data < (now() at time zone 'Africa/Maputo')::date - 120", readOnly: true }),
  PRESENCAS: M('presencas', 'x,data:d,emp_id:i,nome,dept,entrada,saida,estado,min_atraso:i,min_saida:i,horas:n,horas_extra:n,fonte,periodo',
    ['Chave', 'Data', 'ID', 'Nome', 'Departamento', 'Entrada', 'Saída', 'Estado', 'Min atraso', 'Min saída antecipada', 'Horas trabalhadas', 'Horas extra', 'Fonte', 'Período'],
    { order: 'data, emp_id, _rid', where: "data >= (now() at time zone 'Africa/Maputo')::date - 400", dflt: { 0: r => (r[2] !== '' && r[1] !== '' ? r[2] + '|' + r[1] : '') } }),
  AJUSTES: M('ajustes', 'n:i,data:d,emp_id:i,nome,tipo,novo_estado,minutos:n,motivo,documento,estado_aprov,aprovado_por,data_aprov',
    ['Nº', 'Data', 'ID', 'Nome', 'Tipo', 'Novo estado', 'Minutos', 'Motivo', 'Documento', 'Estado aprovação', 'Aprovado por', 'Data aprovação']),
  LICENCAS: M('licencas', 'n:i,emp_id:i,nome,tipo,de:d,ate:d,dias:n,documento,estado,aprovado_por,notas,datas:da,pedido_por,pedido_em,ref',
    ['Nº', 'ID', 'Nome', 'Tipo', 'De', 'Até', 'Dias', 'Documento', 'Estado aprovação', 'Aprovado por', 'Notas', 'Datas', 'Pedido por', 'Pedido em', 'Ref.']),
  DIVIDAS: M('dividas', 'n:i,emp_id:i,nome,tipo,data:d,total:n,prest:n,nprest:i,x,x,estado,ref,notas,ent,inicio,pago_antes:n,registado_por,origem',
    ['Nº', 'ID', 'Nome', 'Tipo', 'Data', 'Valor total', 'Prestação mensal', 'Nº prestações', 'Valor pago', 'Saldo', 'Estado', 'Ref.', 'Notas', 'Entidade', 'Início desconto', 'Pago antes', 'Registado por', 'Origem']),
  DESCONTOS: M('descontos', 'periodo,n:i,emp_id:i,nome,valor:n,tipo,data:d,estado,notas,registado_por',
    ['Período', 'Nº dívida', 'ID', 'Nome', 'Valor', 'Tipo', 'Data', 'Estado', 'Notas', 'Registado por']),
  PROCESSAMENTO: M('processamento', 'periodo,emp_id:i', PAY_HEAD, { order: 'periodo, _rid', wide: true }),
  PROC_AJUSTES: M('proc_ajustes', 'periodo,emp_id:i,nome,bonus:n,subs:n,ef:n,outros:n,nota,atualizado_por,atualizado_em:ts,inss_manual:n',
    ['Período', 'ID', 'Nome', 'Bónus', 'Subsídios', 'Dias efetivos', 'Outros descontos', 'Nota', 'Atualizado por', 'Atualizado em', 'INSS manual']),
  HISTORICO: M('historico', 'em:ts,utilizador,tabela,chave,campo,antes,depois,notas', ['Data/hora', 'Utilizador', 'Tabela', 'Chave', 'Campo', 'Valor anterior', 'Valor novo', 'Notas'], { appendOnly: true }),
  FERIADOS: M('feriados', 'data:d,nome,tipo,ambito,ativo,notas', ['Data', 'Nome', 'Tipo', 'Âmbito', 'Ativo', 'Notas']),
  FERIAS_SALDOS: M('ferias_saldos', 'emp_id:i,nome,ano:i,direito:n,transitado:n,usados_antes:n,data_base:d,estado,notas', ['ID', 'Nome', 'Ano', 'Direito', 'Transitado', 'Usados antes', 'Data base', 'Estado', 'Notas']),
  PEDIDOS: M('pedidos', 'n:i,tipo,emp_id:i,nome,dados:j,documento,estado,submetido_por,submetido_em:ts,decidido_por,decidido_em:ts,nota,resultado',
    ['Nº', 'Tipo', 'ID', 'Nome', 'Dados', 'Documento', 'Estado', 'Submetido por', 'Submetido em', 'Decidido por', 'Decidido em', 'Nota da decisão', 'Resultado']),
  HORAS_EXTRA: M('horas_extra', 'n:i,data:d,emp_id:i,nome,horas:n,inicio,fim,responsavel,motivo,documento,aprovado_por,aprovado_em,pedido_n:i',
    ['Nº', 'Data', 'ID', 'Nome', 'Horas', 'Início', 'Fim', 'Responsável', 'Motivo', 'Documento', 'Aprovado por', 'Aprovado em', 'Nº pedido']),
  DOCUMENTOS: M('documentos', 'n:i,emp_id:i,nome,tipo,numero,validade:d,ficheiro,notas,registado_por,registado_em:ts',
    ['Nº', 'ID', 'Nome', 'Tipo', 'Número', 'Validade', 'Ficheiro', 'Notas', 'Registado por', 'Registado em'])
};

/* ---------------- dates (fixed UTC+2) ---------------- */
const pad = (n, w = 2) => String(n).padStart(w, '0');
export function fmtDate(d, tz, pat) {
  if (!(d instanceof Date) && !(d && typeof d.getTime === 'function')) d = new Date(d);
  const off = /utc|gmt/i.test(String(tz)) ? 0 : OFF;
  const x = new Date(d.getTime() + off);
  const Y = x.getUTCFullYear(), Mo = x.getUTCMonth() + 1, D = x.getUTCDate(), H = x.getUTCHours(), Mi = x.getUTCMinutes(), S = x.getUTCSeconds();
  const WK = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  let out = '', i = 0;
  while (i < pat.length) {
    const ch = pat[i];
    if (ch === "'") { const j = pat.indexOf("'", i + 1); out += pat.slice(i + 1, j < 0 ? pat.length : j); i = j < 0 ? pat.length : j + 1; continue; }
    let k = i; while (k < pat.length && pat[k] === ch) k++;
    const n = k - i;
    switch (ch) {
      case 'y': out += n === 2 ? pad(Y % 100) : Y; break;
      case 'M': out += n >= 2 ? pad(Mo) : Mo; break;
      case 'd': out += n >= 2 ? pad(D) : D; break;
      case 'H': out += n >= 2 ? pad(H) : H; break;
      case 'm': out += n >= 2 ? pad(Mi) : Mi; break;
      case 's': out += n >= 2 ? pad(S) : S; break;
      case 'E': out += WK[x.getUTCDay()]; break;
      case 'X': case 'Z': out += off ? '+02:00' : (ch === 'X' ? 'Z' : '+00:00'); break;
      default: out += pat.slice(i, k);
    }
    i = k;
  }
  return out;
}
const isDate = v => v && typeof v === 'object' && typeof v.getTime === 'function';
const localDT = v => fmtDate(v, 'Africa/Maputo', 'yyyy-MM-dd HH:mm:ss');

/* ---------------- cell ↔ db conversion ---------------- */
const NUMRE = /^-?\d+(\.\d+)?$/;
function toDb(v, ty) { // returns [ok, value]
  if (v === '' || v === null || v === undefined) return [true, null];
  switch (ty) {
    case 'i': case 'n': {
      if (typeof v === 'number') return isFinite(v) && (ty === 'n' || Number.isInteger(v)) ? [true, v] : [false];
      if (typeof v === 'string' && NUMRE.test(v.trim())) { const x = Number(v.trim()); return ty === 'n' || Number.isInteger(x) ? [true, x] : [false]; }
      return [false];
    }
    case 'd': {
      if (isDate(v)) return [true, fmtDate(v, 'Africa/Maputo', 'yyyy-MM-dd')];
      if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return [true, v];
      return [false];
    }
    case 'ts': {
      if (isDate(v)) return [true, v.toISOString()];
      const m = typeof v === 'string' && v.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2})(:\d{2})?$/);
      return m ? [true, m[1] + 'T' + m[2] + (m[3] || ':00') + '+02:00'] : [false];
    }
    case 'tm': {
      if (isDate(v)) return [true, fmtDate(v, 'Africa/Maputo', 'HH:mm:ss')];
      return typeof v === 'string' && /^\d{2}:\d{2}(:\d{2})?$/.test(v) ? [true, v.length === 5 ? v + ':00' : v] : [false];
    }
    case 'dt': {
      if (isDate(v)) return [true, localDT(v)];
      return typeof v === 'string' && /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(v) ? [true, v] : [false];
    }
    case 'da': {
      if (typeof v !== 'string') return [false];
      const a = v.split(',').map(s => s.trim());
      return a.every(s => /^\d{4}-\d{2}-\d{2}$/.test(s)) && a.join(',') === v ? [true, a] : [false];
    }
    case 'j': {
      if (typeof v !== 'string') return [false];
      try { const o = JSON.parse(v); return JSON.stringify(o) === v ? [true, o] : [false]; } catch (e) { return [false]; }
    }
    default: {
      if (typeof v === 'string') return [true, v];
      return [false]; // numbers, booleans, dates stay exact in extra
    }
  }
}
function fromDb(v, ty) {
  if (v === null || v === undefined) return '';
  switch (ty) {
    case 'i': case 'n': return Number(v);
    case 'd': return String(v).slice(0, 10);
    case 'ts': return localDT(new Date(v)).slice(0, 16);
    case 'tm': return String(v).slice(0, 8);
    case 'dt': return String(v).replace('T', ' ').slice(0, 19);
    case 'da': return Array.isArray(v) ? v.join(',') : String(v);
    case 'j': return JSON.stringify(v);
    default: return String(v);
  }
}
const encX = v => isDate(v) ? { __d: v.toISOString() } : v;
const decX = v => v && typeof v === 'object' && v.__d ? new Date(v.__d) : v;

export function recToRow(map, rec) {
  const w = Math.max(map.head.length, map.cols.length);
  const row = new Array(w).fill('');
  if (map.wide) {
    const l = Array.isArray(rec.linha) ? rec.linha : [];
    for (let i = 0; i < l.length; i++) row[i] = decX(l[i]) ?? '';
  } else {
    map.cols.forEach((c, i) => { if (c) row[i] = fromDb(rec[c.c], c.ty); });
  }
  const ex = rec.extra || {};
  Object.keys(ex).forEach(k => { row[Number(k)] = decX(ex[k]); });
  if (map.dflt) Object.keys(map.dflt).forEach(k => { if (row[k] === '' && !(k in ex)) row[k] = map.dflt[k](row); });
  return row;
}
export function rowToRec(map, row) {
  const rec = {}, extra = {};
  if (map.wide) {
    rec.periodo = typeof row[0] === 'string' ? row[0] : isDate(row[0]) ? fmtDate(row[0], 'Africa/Maputo', 'yyyy-MM') : String(row[0] ?? '');
    const id = toDb(row[1], 'i'); rec.emp_id = id[0] ? id[1] : null;
    rec.linha = row.map(v => encX(v === undefined ? '' : v));
    let last = row.length - 1; while (last >= 0 && (rec.linha[last] === '' || rec.linha[last] == null)) last--;
    rec.linha = rec.linha.slice(0, last + 1);
    return rec;
  }
  map.cols.forEach((c, i) => { if (c) rec[c.c] = null; });
  for (let i = 0; i < row.length; i++) {
    const v = row[i], c = map.cols[i];
    if (c) {
      const [ok, x] = toDb(v, c.ty);
      if (ok) rec[c.c] = x; else extra[i] = encX(v);
    } else if (v !== '' && v !== null && v !== undefined) {
      if (map.dflt && map.dflt[i] && map.dflt[i](row) === v) continue;
      extra[i] = encX(v);
    }
  }
  rec.extra = Object.keys(extra).length ? extra : null;
  return rec;
}
const emptyRow = r => !r || r.every(v => v === '' || v === null || v === undefined);

/* ---------------- fake Sheets ---------------- */
const colNum = s => s.toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
class Range {
  constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) { const row = this.sh._row(this.r + i - 1, false); const a = []; for (let j = 0; j < this.nc; j++) { const v = row ? row[this.c + j - 1] : ''; a.push(v === undefined || v === null ? '' : v); } out.push(a); }
    return out;
  }
  getDisplayValues() { return this.getValues().map(r => r.map(v => isDate(v) ? fmtDate(v, 'Africa/Maputo', 'dd/MM/yyyy') : String(v))); }
  getValue() { return this.getValues()[0][0]; }
  setValues(v) {
    if (v.length !== this.nr || (v[0] && v[0].length !== this.nc)) throw new Error('The number of rows/columns in the data does not match the range');
    this.sh._write(this.r, this.c, v); return this;
  }
  setValue(x) { const v = []; for (let i = 0; i < this.nr; i++) { const a = []; for (let j = 0; j < this.nc; j++) a.push(x); v.push(a); } this.sh._write(this.r, this.c, v); return this; }
  clearContent() { return this.setValue(''); }
  clear() { return this.clearContent(); }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getLastRow() { return this.r + this.nr - 1; }
  _style(k, v) { if (this.sh.ss.isExport) for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) { const key = (this.r + i) + ',' + (this.c + j); const s = this.sh.styles[key] = this.sh.styles[key] || {}; s[k] = typeof v === 'function' ? v(i, j) : v; } return this; }
  setNumberFormat(f) { return this._style('numFmt', f); }
  setNumberFormats(f) { return this._style('numFmt', (i, j) => f[i][j]); }
  setFontWeight(f) { return this._style('bold', f === 'bold'); }
  setFontWeights(f) { return this._style('bold', (i, j) => f[i][j] === 'bold'); }
  setBackground(c) { return this._style('bg', c); }
  setBackgrounds(c) { return this._style('bg', (i, j) => c[i][j]); }
  setFontColor(c) { return this._style('color', c); }
  setFontColors(c) { return this._style('color', (i, j) => c[i][j]); }
  setFontSize(n) { return this._style('size', n); }
  setHorizontalAlignment(a) { return this._style('align', a); }
  setVerticalAlignment(a) { return this._style('valign', a); }
  setWrap(w) { return this._style('wrap', w); }
  setBorder() { return this._style('border', true); }
  merge() { if (this.sh.ss.isExport) this.sh.merges.push([this.r, this.c, this.nr, this.nc]); return this; }
  copyTo() { return this; }
  setDataValidation() { return this; }
  activate() { return this; }
}
class Sheet {
  constructor(ss, name, head, rows) {
    this.ss = ss; this.name = name; this.rows = rows || []; // rows[0] = header
    if (head) this.rows.unshift({ rid: 'H', v: head.slice() });
    this.extraCols = 0; this.styles = {}; this.merges = []; this.widths = {}; this.heights = {}; this.frozen = [0, 0]; this.dirty = false; this.hidden = false;
  }
  _row(i, create) { if (!this.rows[i]) { if (!create) return null; while (this.rows.length <= i) this.rows.push({ rid: null, v: [] }); } return this.rows[i].v; }
  _write(r, c, v) {
    if (this.readOnly && r > 1) throw new Error('read-only sheet ' + this.name);
    for (let i = 0; i < v.length; i++) { const row = this._row(r + i - 1, true); for (let j = 0; j < v[i].length; j++) { while (row.length < c + j - 1) row.push(''); row[c + j - 1] = v[i][j]; } }
    if (r + v.length - 1 > 1) this.dirty = true;
    this.ss.touched = true;
  }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getSheetId() { return this.name.length; }
  getLastRow() { for (let i = this.rows.length - 1; i >= 0; i--) if (!emptyRow(this.rows[i].v)) return i + 1; return 0; }
  getLastColumn() { let w = 0; this.rows.forEach(r => { for (let j = r.v.length - 1; j >= 0; j--) if (r.v[j] !== '' && r.v[j] !== null && r.v[j] !== undefined) { w = Math.max(w, j + 1); break; } }); return w; }
  getMaxColumns() { return Math.max(this.getLastColumn(), (this.rows[0] ? this.rows[0].v.length : 0) + this.extraCols, 26); }
  getMaxRows() { return Math.max(this.rows.length, 1000); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = a.match(/^([A-Z]+)(\d*)(?::([A-Z]+)(\d*))?$/i);
      if (!m) throw new Error('Range ' + a);
      const c1 = colNum(m[1]), r1 = m[2] ? Number(m[2]) : 1, c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? Number(m[4]) : (m[2] && !m[3] ? r1 : this.getMaxRows());
      return new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new Range(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }
  getDataRange() { return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  appendRow(v) { this._write(this.getLastRow() + 1, 1, [v]); return this; }
  deleteRow(r) { return this.deleteRows(r, 1); }
  deleteRows(r, n) { if (this.readOnly) throw new Error('read-only'); this.rows.splice(r - 1, n); this.dirty = true; this.ss.touched = true; return this; }
  insertRowsAfter() { return this; }
  insertColumnsAfter(c, n) { this.extraCols += n; return this; }
  clearContents() { for (let i = 1; i < this.rows.length; i++) this.rows[i].v = []; this.rows[0] && (this.rows[0].v = []); this.dirty = true; this.ss.touched = true; return this; }
  clear() { return this.clearContents(); }
  setFrozenRows(n) { this.frozen[0] = n; return this; }
  setFrozenColumns(n) { this.frozen[1] = n; return this; }
  setColumnWidth(c, w) { this.widths[c] = w; return this; }
  setColumnWidths(c, n, w) { for (let i = 0; i < n; i++) this.widths[c + i] = w; return this; }
  setRowHeight(r, h) { this.heights[r] = h; return this; }
  autoResizeColumns() { return this; }
  hideSheet() { this.hidden = true; return this; }
  showSheet() { return this; }
  activate() { return this; }
  setTabColor() { return this; }
  getParent() { return this.ss; }
  copyTo(ss) { const s = ss.insertSheet(this.name + ' copy'); s.rows = this.rows.map(r => ({ rid: null, v: r.v.slice() })); return s; }
}
class Spreadsheet {
  constructor(rt, id, name, isExport) { this.rt = rt; this.id = id; this.name = name; this.isExport = isExport; this.sheets = []; }
  getId() { return this.id; }
  getName() { return this.name; }
  getUrl() { return this.isExport ? 'sbfile:exports/' + this.id + '.xlsx' : ''; }
  setSpreadsheetTimeZone() { return this; }
  getSpreadsheetTimeZone() { return 'Africa/Maputo'; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(n) {
    const s = this.sheets.find(x => x.name === n);
    if (s || this.isExport) return s || null;
    return this.rt.loadSheet(n);
  }
  insertSheet(n) {
    if (!this.isExport && MAP[n]) { const s = this.rt.loadSheet(n); if (s) return s; }
    const s = new Sheet(this, n || ('Sheet' + (this.sheets.length + 1)), null, []);
    this.sheets.push(s); return s;
  }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
  copy(name) { return new Spreadsheet(this.rt, 'copy', name, true); }
  toast() {}
  rename(n) { this.name = n; }
}

export class NeedMore extends Error { constructor(m) { super(m); this.needMore = true; } }

/* ---------------- the runtime ---------------- */
export class Runtime {
  constructor(store, ctx) {
    this.store = store; // { tables: {SHEET: [records]}, props: {}, files: {}, fetches: {} }
    this.ctx = ctx;     // { role, user, email, appUrl, now }
  }
  reset() {
    this.needs = { tables: new Set(), files: new Set(), fetch: null };
    this.db = new Spreadsheet(this, 'DB', 'Sistema RH Madi — DB', false);
    this.exports = []; this.uploads = []; this.trash = []; this.logs = [];
    this.props = Object.assign({}, this.store.props); this.propsDirty = new Set();
  }
  loadSheet(name) {
    const map = MAP[name];
    if (!map) return null; // unknown helper sheets (CACHE, COMPARACAO…) don't exist until inserted
    const recs = this.store.tables[name];
    if (!recs) { this.needs.tables.add(name); throw new NeedMore('table ' + name); }
    const rows = recs.map(r => ({ rid: r._rid, v: recToRow(map, r) }));
    const s = new Sheet(this.db, name, map.head, rows);
    s.readOnly = !!map.readOnly; s.map = map;
    s.orig = new Map(recs.map(r => [r._rid, r]));
    this.db.sheets.push(s);
    return s;
  }
  globals() {
    const rt = this, ctx = this.ctx, NOW = ctx.now;
    class FDate extends Date {
      constructor(...a) { if (a.length) super(...a); else super(NOW); }
      static now() { return NOW; }
    }
    const props = {
      getProperty: k => (k in rt.props ? rt.props[k] : null),
      setProperty: (k, v) => { rt.props[k] = String(v); rt.propsDirty.add(k); return props; },
      setProperties: o => { Object.keys(o).forEach(k => props.setProperty(k, o[k])); return props; },
      getProperties: () => Object.assign({}, rt.props),
      getKeys: () => Object.keys(rt.props),
      deleteProperty: k => { delete rt.props[k]; rt.propsDirty.add(k); return props; }
    };
    const cacheKey = k => '__cache__' + k, mem = {}, keep = k => /^pinfail_/.test(k); // only PIN lock counters persist; other caches live for one request
    const cache = {
      get: k => { if (!keep(k)) return k in mem ? mem[k] : null; const v = props.getProperty(cacheKey(k)); if (!v) return null; try { const o = JSON.parse(v); return o.e > NOW ? o.v : null; } catch (e) { return null; } },
      put: (k, v, s) => { if (!keep(k)) { mem[k] = String(v); return; } props.setProperty(cacheKey(k), JSON.stringify({ v: String(v), e: NOW + (s || 600) * 1000 })); },
      remove: k => { if (!keep(k)) { delete mem[k]; return; } if (props.getProperty(cacheKey(k)) !== null) props.deleteProperty(cacheKey(k)); },
      getAll: ks => { const o = {}; ks.forEach(k => { const v = cache.get(k); if (v !== null) o[k] = v; }); return o; },
      putAll: (o, s) => Object.keys(o).forEach(k => cache.put(k, o[k], s)),
      removeAll: ks => ks.forEach(cache.remove)
    };
    const blob = (bytes, mime, name) => ({ bytes, mime, name, getBytes: () => bytes, getContentType: () => mime, getName: () => name, setName(n) { this.name = n; return this; }, getDataAsString: () => new TextDecoder().decode(bytes) });
    const toBytes = s => typeof s === 'string' ? new TextEncoder().encode(s) : s;
    const Utilities = {
      formatDate: fmtDate,
      getUuid: () => randomUUID(),
      sleep: () => {},
      base64Encode: b => Buffer.from(toBytes(b)).toString('base64'),
      base64Decode: s => new Uint8Array(Buffer.from(String(s), 'base64')),
      newBlob: (bytes, mime, name) => blob(toBytes(bytes), mime, name),
      computeDigest: (alg, s) => Array.from(createHash('sha256').update(toBytes(s)).digest()).map(b => b > 127 ? b - 256 : b),
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' }
    };
    const fileObj = (path, name, mime) => ({
      getId: () => 'sb:' + path, getUrl: () => 'sbfile:' + path, getName: () => name || path.split('/').pop(),
      getBlob: () => { const f = rt.store.files[path]; if (!f) { rt.needs.files.add(path); throw new NeedMore('file ' + path); } return blob(f.bytes, f.mime, name || path.split('/').pop()); },
      setTrashed: () => { rt.trash.push(path); }, setName() { return this; }, moveTo() { return this; }, makeCopy() { return this; }, setSharing() { return this; },
      getDateCreated: () => new FDate(), getMimeType: () => mime || ''
    });
    const folder = base => ({
      getId: () => 'folder:' + base, getUrl: () => '', getName: () => base,
      createFile: b => { const ext = (b.name || '').split('.').pop() || 'bin'; const path = base + '/' + randomUUID() + '.' + ext; rt.uploads.push({ path, bytes: b.bytes, mime: b.mime, name: b.name }); return fileObj(path, b.name, b.mime); },
      getFoldersByName: n => { let used = false; return { hasNext: () => !used, next: () => { used = true; return folder(base + '/' + n.replace(/[^\w-]+/g, '_')); } }; },
      createFolder: n => folder(base + '/' + n.replace(/[^\w-]+/g, '_')),
      getFiles: () => ({ hasNext: () => false, next: () => null })
    });
    const DriveApp = {
      getFileById: id => {
        id = String(id);
        const ex = rt.exports.find(s => s.id === id);
        if (ex) return { getUrl: () => ex.getUrl(), getName: () => ex.name, setName(n) { ex.name = n; return this; }, moveTo() { return this; }, setSharing() { return this; }, setTrashed: () => { rt.exports = rt.exports.filter(s => s !== ex); }, makeCopy() { return this; } };
        if (id.indexOf('sb:') === 0) return fileObj(id.slice(3));
        throw new Error('Ficheiro antigo do Google Drive não disponível · Old Drive file not available');
      },
      getFolderById: () => folder('docs'), createFolder: () => folder('docs'), getRootFolder: () => folder('docs'),
      Access: { ANYONE_WITH_LINK: 1 }, Permission: { VIEW: 1 }
    };
    const SpreadsheetApp = {
      openById: id => rt.exports.find(s => s.id === id) || rt.db,
      create: name => { const s = new Spreadsheet(rt, randomUUID(), name, true); s.sheets.push(new Sheet(s, 'Sheet1', null, [])); rt.exports.push(s); return s; },
      getActive: () => rt.db, flush: () => {},
      BorderStyle: { SOLID: 'solid' }
    };
    const UrlFetchApp = {
      fetch: (url, o) => {
        o = o || {};
        const key = JSON.stringify([url, o.method || 'get', o.payload || '', o.headers || {}]);
        const r = rt.store.fetches[key];
        if (!r) { if (!rt.needs.fetch) rt.needs.fetch = { key, url, o }; throw new NeedMore('fetch'); }
        return { getContentText: () => r.text, getResponseCode: () => r.code, getBlob: () => blob(toBytes(r.text), 'text/plain', '') };
      }
    };
    return {
      Date: FDate,
      SpreadsheetApp, DriveApp, UrlFetchApp, Utilities,
      PropertiesService: { getScriptProperties: () => props, getUserProperties: () => props },
      CacheService: { getScriptCache: () => cache, getUserCache: () => cache },
      LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {}, hasLock: () => true }) },
      Session: { getActiveUser: () => ({ getEmail: () => ctx.email || '' }), getEffectiveUser: () => ({ getEmail: () => ctx.ownerEmail || '' }), getScriptTimeZone: () => 'Africa/Maputo' },
      ScriptApp: { getService: () => ({ getUrl: () => ctx.appUrl || '' }), getProjectTriggers: () => [], newTrigger: () => { const b = { timeBased: () => b, everyDays: () => b, atHour: () => b, nearMinute: () => b, inTimezone: () => b, everyMinutes: () => b, everyHours: () => b, create: () => b }; return b; }, deleteTrigger: () => {} },
      HtmlService: { createHtmlOutput: () => ({}) },
      MailApp: { sendEmail: () => {} }, GmailApp: { sendEmail: () => {} },
      Logger: { log: (...a) => { rt.logs.push(a.map(String).join(' ')); } },
      console: { log: (...a) => rt.logs.push(a.map(String).join(' ')), error: (...a) => rt.logs.push('ERR ' + a.map(String).join(' ')), warn: () => {} },
      __ctx: ctx, __rt: rt
    };
  }

  // Diff in-memory sheets against what was loaded → { table: {del:[rid], ins:[rec]} }
  changes() {
    const out = {};
    this.db.sheets.forEach(s => {
      if (!s.map || !s.dirty || s.readOnly) return;
      const map = s.map, t = map.t;
      const now = s.rows.slice(1).filter(r => !emptyRow(r.v));
      const keyOf = rec => JSON.stringify(Object.keys(rec).sort().map(k => [k, rec[k]]));
      const origKeys = new Map(); // key -> [rid]
      s.orig.forEach((rec, rid) => {
        const clean = rowToRec(map, recToRow(map, rec)); // normalise through the same path
        const k = keyOf(clean); (origKeys.get(k) || origKeys.set(k, []).get(k)).push(rid);
      });
      const ins = [];
      now.forEach(r => {
        const rec = rowToRec(map, r.v), k = keyOf(rec);
        const l = origKeys.get(k);
        if (l && l.length) l.shift(); else ins.push(rec);
      });
      const del = []; origKeys.forEach(l => l.forEach(rid => del.push(rid)));
      if (map.appendOnly) del.length = 0;
      if (del.length || ins.length) { const o = out[t] = out[t] || { del: [], ins: [] }; o.del.push(...del); o.ins.push(...ins); }
    });
    return out;
  }
  propChanges() { const o = {}; this.propsDirty.forEach(k => { o[k] = k in this.props ? this.props[k] : null; }); return o; }
}

/* ---------------- compile Code.gs once ---------------- */
const GLOBAL_NAMES = ['Date', 'SpreadsheetApp', 'DriveApp', 'UrlFetchApp', 'Utilities', 'PropertiesService', 'CacheService', 'LockService', 'Session', 'ScriptApp', 'HtmlService', 'MailApp', 'GmailApp', 'Logger', 'console', '__ctx', '__rt'];
export function compile(code, overrides, exportsList) {
  const body = code + '\n;\n' + overrides + '\n;\nreturn {' + exportsList.map(n => n + ': typeof ' + n + " === 'undefined' ? undefined : " + n).join(', ') + '};';
  return new Function(...GLOBAL_NAMES, body);
}
export function instantiate(fn, rt) {
  rt.reset();
  const g = rt.globals();
  return fn(...GLOBAL_NAMES.map(n => g[n]));
}

/* ---------------- run with retries (load tables / files / fetches on demand) ---------------- */
// io = { loadTables(names) -> {SHEET: recs}, loadFiles(paths) -> {path:{bytes,mime}}, doFetch(url,o) -> {code,text} }
export async function run(fn, rt, io, call, maxRounds = 40) {
  for (let round = 0; round < maxRounds; round++) {
    const api = instantiate(fn, rt);
    let result, error = null;
    try { result = call(api); } catch (e) { error = e; }
    const n = rt.needs;
    if (n.tables.size || n.files.size || n.fetch) {
      const jobs = [];
      if (n.tables.size) jobs.push(io.loadTables([...n.tables]).then(t => Object.assign(rt.store.tables, t)));
      if (n.files.size) jobs.push(io.loadFiles([...n.files]).then(f => Object.assign(rt.store.files, f)));
      if (n.fetch) { const f = n.fetch; jobs.push(io.doFetch(f.url, f.o).then(r => { rt.store.fetches[f.key] = r; })); }
      await Promise.all(jobs);
      continue;
    }
    if (error) throw error;
    return result;
  }
  throw new Error('Too many rounds');
}
