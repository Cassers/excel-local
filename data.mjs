const X = () => globalThis.XLSX;
export const LIMITS = { rows: 100000, cols: 256, cells: 500000, bytes: 20 * 1024 * 1024 };
export const address = (r, c) => X().utils.encode_cell({ r, c });
export const column = c => X().utils.encode_col(c);
export const decode = value => X().utils.decode_cell(value);
export function blankSheet(name = 'Hoja 1') { return { name, rows: 30, cols: 12, cells: {} }; }
export function blankBook() { return { name: 'Sin título', sheets: [blankSheet()] }; }
export function hasFormulas(book) { return book.sheets.some(s => Object.values(s.cells).some(c => c.f)); }
export function checkDimensions(rows, cols) {
  if (rows > LIMITS.rows || cols > LIMITS.cols) throw Error('El límite es de 100.000 filas y 256 columnas por hoja. Divide el archivo para abrirlo.');
}
export function parseCell(value) {
  const raw = String(value);
  if (!raw) return undefined;
  if (raw.startsWith("'")) return { t: 's', v: raw.slice(1) };
  if (raw.startsWith('=') && raw.length > 1) return { t: 'n', f: raw.slice(1) };
  if (/^(true|false)$/i.test(raw)) return { t: 'b', v: raw.toLowerCase() === 'true' };
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(raw) && Number.isFinite(Number(raw)) && raw.replace(/[^0-9]/g, '').length <= 15) return { t: 'n', v: Number(raw) };
  return { t: 's', v: raw };
}
export function rawCell(cell) {
  if (!cell) return '';
  if (cell.f) return '=' + cell.f;
  if (cell.t === 'd') return cell.v instanceof Date ? cell.v.toISOString().slice(0, 10) : String(cell.v);
  const value = String(cell.v ?? '');
  if (cell.t === 's' && (value.startsWith("'") || parseCell(value)?.t !== 's' || value.startsWith('='))) return "'" + value;
  return value;
}
export function displayCell(cell) {
  if (!cell) return '';
  if (cell.f) return '=' + cell.f;
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  return X().utils.format_cell(cell);
}
export function importBook(buffer, filename) {
  if (buffer.byteLength > LIMITS.bytes) throw Error('El archivo supera los 20 MB. Prueba con un archivo más pequeño.');
  const workbook = X().read(buffer, { type: 'array', cellFormula: true, cellNF: true, raw: true });
  if (!workbook.SheetNames.length) throw Error('El archivo no contiene hojas.');
  let count = 0;
  const sheets = workbook.SheetNames.map(name => {
    const source = workbook.Sheets[name];
    const range = X().utils.decode_range(source['!ref'] || 'A1');
    checkDimensions(range.e.r + 1, range.e.c + 1);
    const sheet = { name, rows: Math.max(30, range.e.r + 1), cols: Math.max(12, range.e.c + 1), cells: {} };
    for (const [key, cell] of Object.entries(source)) {
      if (key.startsWith('!') || !cell || cell.t === 'z') continue;
      const clean = {};
      for (const field of ['t', 'v', 'f', 'F', 'D', 'z']) if (cell[field] !== undefined) clean[field] = cell[field];
      sheet.cells[key] = clean;
      if (++count > LIMITS.cells) throw Error('El libro supera las 500.000 celdas ocupadas. Divídelo en archivos más pequeños.');
    }
    return sheet;
  });
  const metadata = {};
  for (const key of ['WBProps', 'Names']) if (workbook.Workbook?.[key]) metadata[key] = structuredClone(workbook.Workbook[key]);
  return { name: filename.replace(/\.(xlsx|xls|csv)$/i, ''), sheets, metadata };
}
export function usedRange(sheet) {
  let r = 0, c = 0;
  for (const key of Object.keys(sheet.cells)) { const pos = decode(key); r = Math.max(r, pos.r); c = Math.max(c, pos.c); }
  return { r, c };
}
export function exportBook(book) {
  const workbook = X().utils.book_new();
  if (book.metadata) workbook.Workbook = structuredClone(book.metadata);
  for (const sheet of book.sheets) {
    const ws = {};
    for (const [key, source] of Object.entries(sheet.cells)) {
      const cell = { ...source };
      // Never export stale cached results after an input cell has changed.
      if (cell.f) { delete cell.v; delete cell.w; cell.t = 'n'; }
      ws[key] = cell;
    }
    const end = usedRange(sheet);
    ws['!ref'] = `A1:${address(end.r, end.c)}`;
    X().utils.book_append_sheet(workbook, ws, sheet.name);
  }
  return workbook;
}
export function exportCSV(sheet) {
  const ws = {};
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (cell.f) ws[key] = { t: 's', v: "'=" + cell.f };
    else if (cell.t === 's' && /^[=+\-@\t\r\n]/.test(cell.v)) ws[key] = { t: 's', v: "'" + cell.v };
    else ws[key] = { ...cell };
  }
  const end = usedRange(sheet);
  ws['!ref'] = `A1:${address(end.r, end.c)}`;
  return '\uFEFF' + X().utils.sheet_to_csv(ws, { blankrows: true });
}
export function validateSheetName(book, name, current = -1) {
  name = name.trim();
  if (!name || name.length > 31 || /[\\/?*\[\]:]/.test(name) || name.startsWith("'") || name.endsWith("'")) throw Error('Usa entre 1 y 31 caracteres, sin \\ / ? * [ ] : ni apóstrofos al principio o al final.');
  if (book.sheets.some((s, i) => i !== current && s.name.toLowerCase() === name.toLowerCase())) throw Error('Ya existe una hoja con ese nombre.');
  return name;
}
export function sortSheet(sheet, col, direction, header) {
  const end = usedRange(sheet).r;
  const first = header ? 1 : 0;
  const rows = Array.from({ length: Math.max(0, end - first + 1) }, (_, i) => i + first);
  const collator = new Intl.Collator('es', { numeric: true, sensitivity: 'base' });
  rows.sort((a, b) => {
    const av = sheet.cells[address(a, col)]?.v, bv = sheet.cells[address(b, col)]?.v;
    if (av === undefined || av === '') return bv === undefined || bv === '' ? 0 : 1;
    if (bv === undefined || bv === '') return -1;
    const comparison = typeof av === 'number' && typeof bv === 'number' ? av - bv : collator.compare(String(av), String(bv));
    return comparison * direction;
  });
  const map = new Map(rows.map((old, i) => [old, i + first]));
  const cells = {};
  for (const [key, value] of Object.entries(sheet.cells)) {
    const p = decode(key);
    cells[address(map.get(p.r) ?? p.r, p.c)] = value;
  }
  return cells;
}
export function parseClipboard(text) {
  // TSV with quoted multiline cells, as copied from Excel / Google Sheets.
  const rows = []; let row = [], value = '', quoted = false;
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"' && (quoted || value === '')) {
      if (quoted && text[i + 1] === '"') { value += '"'; i++; }
      else quoted = !quoted;
    } else if (!quoted && (ch === '\t' || ch === '\n')) {
      row.push(value); value = '';
      if (ch === '\n') { rows.push(row); row = []; }
    } else value += ch;
  }
  if (value || row.length || !rows.length) { row.push(value); rows.push(row); }
  return rows;
}
