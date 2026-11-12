/**
 * Mock minimal Google Apps Script SpreadsheetApp.
 * Menjalankan Code.gs ASLI di Node.js, merekam setiap operasi ke tools/_work/model.json,
 * lalu tools/build_xlsx.py memutar ulang operasi itu menjadi file .xlsx.
 * Sekaligus menulis sample_transactions.csv dari generator data simulasi di Code.gs.
 *
 * Pakai:  TZ=Asia/Jakarta node tools/mock_sheets.js
 */
const fs = require('fs'), path = require('path'), vm = require('vm');
const root = path.join(__dirname, '..');
const work = path.join(__dirname, '_work');
fs.mkdirSync(work, { recursive: true });

const pad = (n) => String(n).padStart(2, '0');
const iso = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const clean = (v) => {
  if (v instanceof Date) return { date: iso(v) };
  if (Array.isArray(v)) return v.map(clean);
  if (v && v.__a1) return v.__a1();
  if (v && typeof v === 'object') { const o = {}; for (const k of Object.keys(v)) o[k] = clean(v[k]); return o; }
  return v;
};
const colNum = (s) => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

class Range {
  constructor(sheet, r1, c1, r2, c2) { Object.assign(this, { sheet, r1, c1, r2, c2 }); }
  __a1() { return `${colName(this.c1)}${this.r1}:${colName(this.c2)}${this.r2}`; }
}
const RECORDED = ['setValue', 'setValues', 'setFormula', 'setFormulas', 'setNumberFormat', 'setBackground', 'setFontColor',
  'setFontWeight', 'setFontStyle', 'setFontSize', 'setFontFamily', 'setHorizontalAlignment', 'setVerticalAlignment',
  'setWrap', 'merge', 'setBorder', 'setDataValidation', 'clearContent'];
for (const name of RECORDED) {
  Range.prototype[name] = function (...args) {
    this.sheet.ops.push({ op: name, r1: this.r1, c1: this.c1, r2: this.r2, c2: this.c2, args: clean(args) });
    return this;
  };
}
class Sheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.ops = []; this.meta = { colWidths: {}, rowHeights: {}, charts: [], cf: [] }; }
  getParent() { return this.ss; }
  setName(n) { this.name = n; return this; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
      if (!m) throw new Error('A1 tidak valid: ' + a);
      const c1 = colNum(m[1]), r1 = +m[2], c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
      return new Range(this, r1, c1, r2, c2);
    }
    return new Range(this, a, b, a + (c || 1) - 1, b + (d || 1) - 1);
  }
  setTabColor(c) { this.meta.tab = c; return this; }
  setHiddenGridlines(v) { this.meta.hideGrid = v; return this; }
  setColumnWidth(c, w) { this.meta.colWidths[c] = w; return this; }
  setRowHeight(r, h) { this.meta.rowHeights[r] = h; return this; }
  setFrozenRows(n) { this.meta.frozen = n; return this; }
  setConditionalFormatRules(rules) { this.meta.cf = clean(rules); return this; }
  newChart() { return chartBuilder(); }
  insertChart(c) { this.meta.charts.push(c); return this; }
}
function chartBuilder() {
  const spec = { ranges: [], options: {} };
  const b = {
    setChartType(t) { spec.type = t; return b; },
    addRange(r) { spec.ranges.push(r.__a1()); return b; },
    setNumHeaders(n) { spec.numHeaders = n; return b; },
    setPosition(r, c, x, y) { spec.pos = { row: r, col: c, x, y }; return b; },
    setOption(k, v) { spec.options[k] = clean(v); return b; },
    build() { return spec; }
  };
  return b;
}
class Spreadsheet {
  constructor(name) { this.name = name; this.sheets = [new Sheet(this, 'Sheet1')]; this.meta = {}; }
  setSpreadsheetLocale(l) { this.meta.locale = l; }
  setSpreadsheetTimeZone(t) { this.meta.tz = t; }
  getSheets() { return this.sheets; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  insertSheet(n) { const s = new Sheet(this, n); this.sheets.push(s); return s; }
  setActiveSheet() {}
  getUrl() { return 'mock://spreadsheet'; }
}
let created;
const SpreadsheetApp = {
  create(name) { created = new Spreadsheet(name); return created; },
  getActiveSpreadsheet() { return null; },
  BorderStyle: { SOLID: 'SOLID', SOLID_THICK: 'SOLID_THICK' },
  newConditionalFormatRule() {
    const s = { ranges: [] };
    const b = {
      whenFormulaSatisfied(f) { s.formula = f; return b; },
      setRanges(r) { s.ranges = r.map((x) => x.__a1()); return b; },
      setBackground(c) { s.bg = c; return b; }, setFontColor(c) { s.fc = c; return b; },
      setBold(v) { s.bold = v; return b; }, build() { return s; }
    };
    return b;
  },
  newDataValidation() {
    const s = {};
    const b = {
      requireValueInList(v) { s.type = 'list'; s.values = v; return b; },
      requireValueInRange(r) { s.type = 'range'; s.sheet = r.sheet.name; s.range = r.__a1(); return b; },
      requireNumberGreaterThanOrEqualTo(n) { s.type = 'gte'; s.n = n; return b; },
      requireDate() { s.type = 'date'; return b; },
      setAllowInvalid(v) { s.allowInvalid = v; return b; },
      setHelpText(t) { s.help = t; return b; },
      build() { return s; }
    };
    return b;
  }
};
const Charts = { ChartType: { COLUMN: 'COLUMN', PIE: 'PIE', BAR: 'BAR' } };
const logs = [];
const ctx = { SpreadsheetApp, Charts, Logger: { log: (m) => logs.push(m) }, Math, Date, Object, JSON, Error, console };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'Code.gs'), 'utf8'), ctx, { filename: 'Code.gs' });
vm.runInContext('setup()', ctx);

const model = { name: created.name, meta: created.meta, sheets: created.sheets.map((s) => ({ name: s.name, ops: s.ops, meta: s.meta })) };
fs.writeFileSync(path.join(work, 'model.json'), JSON.stringify(model));

// CSV dari generator yang sama
const data = vm.runInContext('buatDataSimulasi_()', ctx);
const q = (s) => (/[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
const csv = ['Tanggal,Jenis Transaksi,Kategori,Keterangan,Metode Pembayaran,Nominal']
  .concat(data.map((r) => [iso(r[0]), r[1], r[2], q(r[3]), r[4], r[5]].join(','))).join('\n') + '\n';
fs.writeFileSync(path.join(root, 'sample_transactions.csv'), csv);
console.log('OK: sheets =', model.sheets.map((s) => s.name).join(', '), '| transaksi demo =', data.length, '|', logs.join(' '));
