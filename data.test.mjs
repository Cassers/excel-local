import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
globalThis.XLSX = require('./vendor/xlsx.full.min.js');
const { parseCell, rawCell, importBook, exportBook, exportCSV, sortSheet, validateSheetName, blankBook, parseClipboard, hasFormulas, checkDimensions } = await import('./data.mjs');
const bytes = wb => XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

test('Excel round trip preserves multiple sheets, types, leading zeros, dates and formulas', () => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([['Nombre', 'Cantidad', 'Código', 'Fecha', 'Activo', 'Total'], ['Café', 12.5, '00123', 46290, true, null]]);
  ws.D2.z = 'yyyy-mm-dd'; ws.F2 = { t: 'n', f: 'B2*2', v: 25 };
  XLSX.utils.book_append_sheet(wb, ws, 'Ventas');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Texto'], ['=no ejecutar']]), 'Datos');
  const imported = importBook(bytes(wb), 'archivo.xlsx');
  imported.sheets[0].cells.B2 = parseCell('20');
  const output = XLSX.read(bytes(exportBook(imported)), { cellNF: true, sheetStubs: true });
  assert.deepEqual(output.SheetNames, ['Ventas', 'Datos']);
  assert.equal(output.Sheets.Ventas.B2.v, 20);
  assert.equal(output.Sheets.Ventas.C2.v, '00123');
  assert.equal(output.Sheets.Ventas.D2.z, 'yyyy-mm-dd');
  assert.equal(output.Sheets.Ventas.E2.v, true);
  assert.equal(output.Sheets.Ventas.F2.f, 'B2*2');
  assert.notEqual(output.Sheets.Ventas.F2.v, 25);
  assert.equal(output.Sheets.Datos.A2.t, 's');
  assert.equal(hasFormulas(imported), true);
});
test('cell entry keeps identifiers and long integers as text, accepts formulas and explicit text', () => {
  assert.deepEqual(parseCell('00123'), { t: 's', v: '00123' });
  assert.deepEqual(parseCell('12345678901234567'), { t: 's', v: '12345678901234567' });
  assert.deepEqual(parseCell('-12.5'), { t: 'n', v: -12.5 });
  assert.deepEqual(parseCell('=SUM(A1:A3)'), { t: 'n', f: 'SUM(A1:A3)' });
  assert.deepEqual(parseCell("'=literal"), { t: 's', v: '=literal' });
  assert.deepEqual(parseCell(rawCell({ t: 's', v: '123' })), { t: 's', v: '123' });
  assert.equal(parseCell(''), undefined);
});
test('sort preserves row associations and header; blank values stay at end', () => {
  const b = blankBook(), s = b.sheets[0];
  s.cells = { A1: parseCell('Nombre'), B1: parseCell('Total'), A2: parseCell('Zeta'), B2: parseCell('12'), A3: parseCell('Alfa'), B3: parseCell('3'), A4: parseCell('Vacío') };
  const sorted = sortSheet(s, 1, 1, true);
  assert.equal(sorted.A1.v, 'Nombre'); assert.equal(sorted.A2.v, 'Alfa'); assert.equal(sorted.B2.v, 3); assert.equal(sorted.A4.v, 'Vacío');
});
test('clipboard supports TSV, quoted newlines, quotes, and trailing blank cells', () => {
  assert.deepEqual(parseClipboard('A\tB\n1\t2\n'), [['A','B'], ['1','2']]);
  assert.deepEqual(parseClipboard('"dos\nlíneas"\t"una ""cita"""\t'), [['dos\nlíneas', 'una "cita"', '']]);
});
test('CSV export protects formula-like strings and exports formulas as explicit text', () => {
  const s = blankBook().sheets[0]; s.cells = { A1: { t: 's', v: '=1+1' }, A2: { t: 'n', f: 'SUM(B1:B2)' }, B1: { t: 'n', v: -5 } };
  const csv = exportCSV(s);
  assert.ok(csv.includes("'=1+1")); assert.ok(csv.includes("'=SUM(B1:B2)")); assert.ok(csv.includes('-5'));
});
test('invalid sheet names, duplicate names and oversized dimensions are rejected', () => {
  assert.throws(() => validateSheetName(blankBook(), 'hoja 1'));
  assert.throws(() => validateSheetName(blankBook(), 'a/b'));
  assert.throws(() => checkDimensions(100001, 10));
  assert.equal(validateSheetName(blankBook(), 'Nueva'), 'Nueva');
});
test('1904 date system and named ranges survive export', () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[100, 40000]]), 'Datos');
  wb.Workbook = { WBProps: { date1904: true }, Names: [{ Name: 'Importe', Ref: 'Datos!$A$1' }] };
  const output = XLSX.read(bytes(exportBook(importBook(bytes(wb), 'fechas.xlsx'))));
  assert.equal(output.Workbook.WBProps.date1904, true);
  assert.equal(output.Workbook.Names[0].Ref, 'Datos!$A$1');
});
