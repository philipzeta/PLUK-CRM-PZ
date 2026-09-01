import { toISODate } from './utils.js';

// Thin wrapper around the globally-loaded SheetJS (window.XLSX).
// columns: [{ key, header, type }] — type 'date'|'number'|'text'|'checkbox'

function normHeader(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

export function buildSheetFromRows(columns, rows) {
  const data = rows.map((r) => {
    const obj = {};
    columns.forEach((c) => {
      let v = c.get ? c.get(r) : r[c.key];
      if (c.type === 'date') v = toISODate(v) || '';
      if (v === null || v === undefined) v = '';
      obj[c.header] = v;
    });
    return obj;
  });
  return XLSX.utils.json_to_sheet(data, { header: columns.map((c) => c.header) });
}

export function downloadWorkbook(sheets, filename) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, columns, rows }) => {
    const ws = buildSheetFromRows(columns, rows);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}

export function downloadSingleSheet(sheetName, columns, rows, filename) {
  downloadWorkbook([{ name: sheetName, columns, rows }], filename || `${sheetName}.xlsx`);
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export async function readWorkbook(file) {
  const buf = await readFileAsArrayBuffer(file);
  return XLSX.read(buf, { type: 'array', cellDates: true });
}

// Parse a single sheet (by name, or the first sheet if omitted) into row
// objects keyed by the supplied column defs, matching headers flexibly
// (case-insensitive, common aliases).
export function parseSheetRows(workbook, columns, sheetName) {
  const name = sheetName && workbook.Sheets[sheetName] ? sheetName : workbook.SheetNames[0];
  const ws = workbook.Sheets[name];
  if (!ws) return [];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
  if (!aoa.length) return [];

  // find the header row: first row containing at least 2 recognizable headers
  const allAliases = new Map();
  columns.forEach((c) => {
    (c.aliases || [c.header]).forEach((a) => allAliases.set(normHeader(a), c.key));
  });

  let headerRowIdx = 0;
  let bestScore = -1;
  for (let i = 0; i < Math.min(aoa.length, 5); i++) {
    const score = aoa[i].filter((cell) => allAliases.has(normHeader(cell))).length;
    if (score > bestScore) { bestScore = score; headerRowIdx = i; }
  }

  const headerRow = aoa[headerRowIdx].map((h) => allAliases.get(normHeader(h)) || null);
  const rows = [];
  for (let r = headerRowIdx + 1; r < aoa.length; r++) {
    const line = aoa[r];
    if (!line || line.every((v) => v === '' || v === null || v === undefined)) continue;
    const obj = {};
    let any = false;
    headerRow.forEach((key, ci) => {
      if (!key) return;
      let v = line[ci];
      if (v === '') v = null;
      if (v !== null) any = true;
      obj[key] = v;
    });
    if (any) rows.push(obj);
  }
  return rows;
}
