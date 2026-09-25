import { LIMITS, address, column, decode, blankBook, blankSheet, checkDimensions, parseCell, rawCell, displayCell, importBook, exportBook, exportCSV, usedRange, hasFormulas, sortSheet, validateSheetName, parseClipboard } from './data.mjs';

const $ = id => document.getElementById(id);
let book = null, active = 0, selected = { r: 0, c: 0 }, page = 0, filtered = [], dirty = false;
let undoStack = [], redoStack = [], toastTimer, searchTimer, editing = null, importing = false;
const PAGE_SIZE = 100;
const sheet = () => book.sheets[active];

function toast(message, error = false) {
  clearTimeout(toastTimer); $('toast').textContent = message; $('toast').classList.toggle('error', error); $('toast').hidden = false;
  toastTimer = setTimeout(() => $('toast').hidden = true, error ? 6500 : 3500);
}
function safe(action) { try { action(); } catch (error) { toast(error.message, true); } }
function markDirty() { dirty = true; $('save-status').textContent = 'Cambios en memoria · descarga para guardarlos'; }
function updateHistory() { $('undo').disabled = !undoStack.length; $('redo').disabled = !redoStack.length; }
function record(command) {
  command.forward(); undoStack.push(command); if (undoStack.length > 40) undoStack.shift(); redoStack = [];
  markDirty(); render();
}
function undo() { finishEdit(); const command = undoStack.pop(); if (!command) return; command.backward(); redoStack.push(command); active = Math.min(active, book.sheets.length - 1); markDirty(); render(); }
function redo() { finishEdit(); const command = redoStack.pop(); if (!command) return; command.forward(); undoStack.push(command); markDirty(); render(); }
function replaceBook(next) {
  book = next; active = 0; selected = { r: 0, c: 0 }; page = 0; undoStack = []; redoStack = []; dirty = false; editing = null;
  $('search').value = ''; $('filename').value = book.name; $('save-status').textContent = 'Listo para editar · archivo original intacto';
  $('welcome').hidden = true; $('editor').hidden = false; render();
}
function mayReplace() { return !dirty || confirm('Hay cambios sin descargar. ¿Quieres descartarlos y continuar?'); }
function newBook() { finishEdit(); if (mayReplace()) { replaceBook(blankBook()); selectCell(0, 0, true); } }
async function loadFile(file) {
  if (!file || importing) return;
  finishEdit();
  if (!/\.(xlsx|xls|csv)$/i.test(file.name)) return toast('Elige un archivo .xlsx, .xls o .csv.', true);
  if (file.size > LIMITS.bytes) { $('file-input').value = ''; return window.LargeEditor.open(file); }
  if (!(await window.LargeEditor.release())) return;
  if (!mayReplace()) return;
  importing = true; $('busy').hidden = false;
  try {
    await new Promise(resolve => setTimeout(resolve, 40));
    const next = importBook(await file.arrayBuffer(), file.name);
    replaceBook(next); toast(`Archivo abierto · ${next.sheets.length} hoja${next.sheets.length === 1 ? '' : 's'}`);
  } catch (error) { console.error(error); toast('No se pudo abrir: ' + error.message, true); }
  finally { importing = false; $('busy').hidden = true; $('file-input').value = ''; }
}
function matchingRows() {
  const current = sheet(), query = $('search').value.trim().toLocaleLowerCase('es');
  if (!query) return Array.from({ length: current.rows }, (_, i) => i);
  const matches = new Set();
  for (const [key, cell] of Object.entries(current.cells)) if (displayCell(cell).toLocaleLowerCase('es').includes(query)) matches.add(decode(key).r);
  return [...matches].sort((a, b) => a - b);
}
function render() {
  if (!book) return;
  selected.r = Math.max(0, Math.min(selected.r, sheet().rows - 1)); selected.c = Math.max(0, Math.min(selected.c, sheet().cols - 1));
  filtered = matchingRows(); page = Math.max(0, Math.min(page, Math.ceil(filtered.length / PAGE_SIZE) - 1));
  renderGrid(); renderTabs(); updateSelection(); updateHistory();
  $('dimensions').textContent = `${sheet().rows.toLocaleString('es')} filas × ${sheet().cols} columnas · ${book.sheets.length} hoja${book.sheets.length === 1 ? '' : 's'}`;
  const start = page * PAGE_SIZE;
  $('page-label').textContent = filtered.length ? `${start + 1}–${Math.min(start + PAGE_SIZE, filtered.length)} de ${filtered.length.toLocaleString('es')}` : '0 filas';
  $('previous-page').disabled = page === 0; $('next-page').disabled = start + PAGE_SIZE >= filtered.length;
}
function renderGrid() {
  const current = sheet(), fragment = document.createDocumentFragment();
  const head = document.createElement('thead'), headRow = document.createElement('tr');
  headRow.append(document.createElement('th'));
  for (let c = 0; c < current.cols; c++) { const th = document.createElement('th'); th.textContent = column(c); th.scope = 'col'; th.dataset.column = c; headRow.append(th); }
  head.append(headRow); fragment.append(head);
  const body = document.createElement('tbody');
  for (const r of filtered.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE)) {
    const row = document.createElement('tr'), label = document.createElement('th'); label.textContent = r + 1; label.scope = 'row'; row.append(label);
    for (let c = 0; c < current.cols; c++) {
      const td = document.createElement('td'), key = address(r, c), cell = current.cells[key];
      td.dataset.r = r; td.dataset.c = c; td.dataset.address = key; td.tabIndex = -1; td.textContent = displayCell(cell);
      td.setAttribute('aria-label', `${key}: ${displayCell(cell) || 'vacía'}`);
      if (cell?.t === 'n' && !cell.f) td.classList.add('numeric');
      if (cell?.f) td.classList.add('formula');
      if (r === 0 && $('has-header').checked) td.classList.add('header-cell');
      row.append(td);
    }
    body.append(row);
  }
  fragment.append(body); $('grid').replaceChildren(fragment); $('no-results').hidden = filtered.length !== 0;
  if (filtered.length && !filtered.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE).includes(selected.r)) selected.r = filtered[page * PAGE_SIZE];
}
function renderTabs() {
  $('sheet-tabs').replaceChildren(...book.sheets.map((s, i) => {
    const button = document.createElement('button'); button.textContent = s.name; button.role = 'tab'; button.setAttribute('aria-selected', i === active); button.title = 'Doble clic para renombrar';
    button.onclick = () => { if (active === i) return; finishEdit(); active = i; selected = { r: 0, c: 0 }; page = 0; $('search').value = ''; render(); };
    button.ondblclick = () => renameSheet(i); return button;
  }));
}
function updateSelection(focus = false) {
  $('grid').querySelectorAll('.selected').forEach(td => { td.classList.remove('selected'); td.tabIndex = -1; });
  const key = address(selected.r, selected.c), cell = sheet().cells[key], td = $('grid').querySelector(`[data-address="${key}"]`);
  if (td) { td.classList.add('selected'); td.tabIndex = 0; if (focus) { td.focus({ preventScroll: true }); td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); } }
  $('cell-address').value = key; $('cell-value').value = rawCell(cell);
  $('cell-type').textContent = cell?.f ? 'Fórmula' : cell?.t === 'n' ? 'Número' : cell?.t === 'b' ? 'Booleano' : cell ? 'Texto' : 'Vacía';
  $('selection-info').textContent = `Celda ${key}`;
}
function selectCell(r, c, focus = false) {
  selected = { r: Math.max(0, Math.min(r, sheet().rows - 1)), c: Math.max(0, Math.min(c, sheet().cols - 1)) };
  const index = filtered.indexOf(selected.r);
  if (index >= 0 && Math.floor(index / PAGE_SIZE) !== page) { page = Math.floor(index / PAGE_SIZE); render(); }
  updateSelection(focus);
}
function setCells(matrix, r = selected.r, c = selected.c) {
  const target = sheet(), height = matrix.length, width = Math.max(...matrix.map(row => row.length));
  checkDimensions(r + height, c + width);
  const entries = matrix.flatMap((row, ri) => row.map((value, ci) => { const key = address(r + ri, c + ci); return { key, old: target.cells[key], next: parseCell(value) }; }));
  const changes = entries.filter(e => JSON.stringify(e.old) !== JSON.stringify(e.next));
  if (!changes.length) return;
  const count = book.sheets.reduce((n, s) => n + Object.keys(s.cells).length, 0) + changes.reduce((n, e) => n + (!e.old && e.next ? 1 : e.old && !e.next ? -1 : 0), 0);
  if (count > LIMITS.cells) throw Error('El libro supera las 500.000 celdas ocupadas.');
  const oldSize = [target.rows, target.cols], newSize = [Math.max(target.rows, r + height), Math.max(target.cols, c + width)];
  const apply = (side, size) => { for (const e of changes) { if (e[side]) target.cells[e.key] = e[side]; else delete target.cells[e.key]; } [target.rows, target.cols] = size; active = book.sheets.indexOf(target); };
  record({ forward: () => apply('next', newSize), backward: () => apply('old', oldSize) });
}
function beginEdit(td, initial) {
  if (!td) return;
  finishEdit(); selected = { r: +td.dataset.r, c: +td.dataset.c }; updateSelection();
  const original = rawCell(sheet().cells[td.dataset.address]);
  editing = { td, r: selected.r, c: selected.c, original };
  td.contentEditable = 'true'; td.textContent = initial ?? original; td.focus();
  const range = document.createRange(); range.selectNodeContents(td); range.collapse(false); const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(range);
}
function finishEdit(cancel = false) {
  if (!editing) return;
  const { td, r, c, original } = editing, value = td.textContent; editing = null; td.contentEditable = 'false';
  if (!cancel && value !== original) safe(() => setCells([[value]], r, c));
  td.textContent = displayCell(sheet().cells[address(r, c)]); updateSelection();
}
function move(dr, dc) {
  const at = filtered.indexOf(selected.r); const next = dr ? filtered[Math.max(0, Math.min(filtered.length - 1, at + dr))] ?? selected.r : selected.r;
  selectCell(next, selected.c + dc, true);
}
function appendDimension(kind) {
  finishEdit(); const target = sheet(), old = target[kind], next = old + 1;
  checkDimensions(kind === 'rows' ? next : target.rows, kind === 'cols' ? next : target.cols);
  record({ forward: () => { active = book.sheets.indexOf(target); target[kind] = next; }, backward: () => { active = book.sheets.indexOf(target); target[kind] = old; } });
  $('search').value = ''; render(); selectCell(kind === 'rows' ? old : selected.r, kind === 'cols' ? old : selected.c, true);
}
function addSheet() {
  finishEdit(); let number = book.sheets.length + 1; while (book.sheets.some(s => s.name.toLowerCase() === `hoja ${number}`)) number++;
  const next = blankSheet(`Hoja ${number}`), previous = active;
  record({ forward: () => { book.sheets.push(next); active = book.sheets.length - 1; selected = { r: 0, c: 0 }; page = 0; $('search').value = ''; }, backward: () => { book.sheets.splice(book.sheets.indexOf(next), 1); active = previous; } });
}
function renameSheet(index) {
  finishEdit(); if (hasFormulas(book)) return toast('Para proteger las referencias, renombra las hojas con fórmulas en Excel.', true);
  const target = book.sheets[index], old = target.name, answer = prompt('Nombre de la hoja:', old); if (answer === null) return;
  safe(() => { const name = validateSheetName(book, answer, index); if (name !== old) record({ forward: () => target.name = name, backward: () => target.name = old }); });
}
function sort(direction) {
  finishEdit(); if (hasFormulas(book)) return toast('Este libro contiene fórmulas. Ordénalo en Excel para actualizar sus referencias.', true);
  if ($('search').value.trim()) return toast('Quita el filtro antes de ordenar la hoja completa.', true);
  const target = sheet(), old = target.cells, next = sortSheet(target, selected.c, direction, $('has-header').checked);
  record({ forward: () => { target.cells = next; active = book.sheets.indexOf(target); }, backward: () => { target.cells = old; active = book.sheets.indexOf(target); } });
  toast(`Hoja ordenada por la columna ${column(selected.c)}`);
}
function download(kind) {
  finishEdit(); if (!book) return;
  safe(() => {
    const name = ($('filename').value.trim() || 'Sin título').replace(/[\\/:*?"<>|]/g, '_'); book.name = name;
    if (kind === 'xlsx') {
      const data = XLSX.write(exportBook(book), { bookType: 'xlsx', type: 'array', compression: true });
      saveBlob(new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${name}-editado.xlsx`);
      dirty = false; $('save-status').textContent = 'Copia Excel generada · revisa tus descargas'; toast('Copia Excel enviada a descargas');
    } else {
      saveBlob(new Blob([exportCSV(sheet())], { type: 'text/csv;charset=utf-8' }), `${name}-${sheet().name}.csv`);
      toast('CSV de la hoja activa descargado. Las fórmulas se exportan como texto.');
    }
  });
}
function saveBlob(blob, name) { const url = URL.createObjectURL(blob), link = document.createElement('a'); link.href = url; link.download = name; document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 30000); }

$('new-welcome').onclick = newBook; $('new-button').onclick = newBook;
for (const id of ['open-button', 'dropzone']) $(id).onclick = () => $('file-input').click();
$('dropzone').onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file-input').click(); } };
$('file-input').onchange = e => loadFile(e.target.files[0]);
document.addEventListener('dragover', e => { if (e.dataTransfer.types.includes('Files')) { e.preventDefault(); $('dropzone').classList.add('dragging'); } });
document.addEventListener('dragleave', e => { if (!e.relatedTarget) $('dropzone').classList.remove('dragging'); });
document.addEventListener('drop', e => { if (e.dataTransfer.files.length) { e.preventDefault(); $('dropzone').classList.remove('dragging'); loadFile(e.dataTransfer.files[0]); } });
$('grid').addEventListener('click', e => {
  if (editing) return;
  const td = e.target.closest('td'); if (td) selectCell(+td.dataset.r, +td.dataset.c, true);
  const th = e.target.closest('th[data-column]'); if (th) selectCell(selected.r, +th.dataset.column, true);
});
$('grid').addEventListener('dblclick', e => beginEdit(e.target.closest('td')));
$('grid').addEventListener('mousedown', e => {
  const td = e.target.closest('td');
  if (!editing || !td || td === editing.td) return;
  const r = +td.dataset.r, c = +td.dataset.c;
  e.preventDefault(); finishEdit(); selectCell(r, c, true);
});
$('grid').addEventListener('focusout', () => finishEdit());
$('grid').addEventListener('keydown', e => {
  if (e.ctrlKey || e.metaKey || e.altKey) return;
  if (editing) {
    if (e.key === 'Escape') { e.preventDefault(); finishEdit(true); updateSelection(true); }
    else if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); const tab = e.key === 'Tab'; finishEdit(); move(tab ? 0 : e.shiftKey ? -1 : 1, tab ? e.shiftKey ? -1 : 1 : 0); }
    return;
  }
  if (['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Tab','Enter'].includes(e.key)) {
    e.preventDefault(); const dirs = { ArrowUp: [-1,0], ArrowDown: [1,0], ArrowLeft: [0,-1], ArrowRight: [0,1], Tab: [0,e.shiftKey ? -1 : 1], Enter: [e.shiftKey ? -1 : 1,0] }; move(...dirs[e.key]);
  } else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); safe(() => setCells([['']])); updateSelection(true); }
  else if (e.key === 'F2') { e.preventDefault(); beginEdit(e.target.closest('td')); }
  else if (e.key.length === 1) { e.preventDefault(); beginEdit(e.target.closest('td'), e.key); }
});
$('grid').addEventListener('paste', e => {
  const text = e.clipboardData.getData('text/plain'); e.preventDefault();
  if (text.length > 5 * 1024 * 1024) return toast('Pega bloques de menos de 5 MB.', true);
  const matrix = parseClipboard(text);
  if ($('search').value.trim() && matrix.length > 1) return toast('Quita el filtro antes de pegar varias filas.', true);
  finishEdit(); safe(() => setCells(matrix)); updateSelection(true);
});
$('grid').addEventListener('copy', e => { if (editing) return; e.preventDefault(); const cell = sheet().cells[address(selected.r, selected.c)]; e.clipboardData.setData('text/plain', cell?.f ? '=' + cell.f : String(cell?.v ?? '')); toast('Celda copiada'); });
$('cell-value').addEventListener('change', () => safe(() => setCells([[$('cell-value').value]])));
$('cell-value').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); $('cell-value').blur(); updateSelection(true); } if (e.key === 'Escape') { $('cell-value').value = rawCell(sheet().cells[address(selected.r, selected.c)]); $('cell-value').blur(); updateSelection(true); } });
$('cell-address').addEventListener('keydown', e => {
  if (e.key !== 'Enter') return; e.preventDefault();
  safe(() => {
    const value = $('cell-address').value.trim().toUpperCase(); if (!/^[A-Z]{1,3}[1-9]\d{0,5}$/.test(value)) throw Error('Escribe una referencia válida, como B12.');
    const pos = decode(value); if (pos.r >= sheet().rows || pos.c >= sheet().cols) throw Error('La celda está fuera de la hoja. Añade filas o columnas primero.');
    $('search').value = ''; page = Math.floor(pos.r / PAGE_SIZE); selected = pos; render(); updateSelection(true);
  });
});
$('search').oninput = () => { clearTimeout(searchTimer); searchTimer = setTimeout(() => { page = 0; render(); }, 180); };
$('has-header').onchange = () => render();
$('filename').onchange = () => { book.name = $('filename').value.trim() || 'Sin título'; markDirty(); };
$('undo').onclick = undo; $('redo').onclick = redo;
$('add-row').onclick = () => safe(() => appendDimension('rows')); $('add-column').onclick = () => safe(() => appendDimension('cols'));
$('clear-cell').onclick = () => safe(() => setCells([['']]));
$('add-sheet').onclick = addSheet; $('sort-asc').onclick = () => sort(1); $('sort-desc').onclick = () => sort(-1);
$('previous-page').onclick = () => { finishEdit(); page--; render(); }; $('next-page').onclick = () => { finishEdit(); page++; render(); };
$('export-xlsx').onclick = () => download('xlsx'); $('export-csv').onclick = () => download('csv');
$('help-button').onclick = () => $('help-dialog').showModal(); $('close-help').onclick = () => $('help-dialog').close();
document.addEventListener('keydown', e => {
  if ($('editor').hidden) return;
  if (!book || !(e.ctrlKey || e.metaKey)) return;
  if (e.key.toLowerCase() === 's') { e.preventDefault(); document.activeElement?.blur(); download('xlsx'); }
  if (e.target.matches('input, [contenteditable="true"]')) return;
  if (e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
  if (e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); }
});
window.addEventListener('beforeunload', e => { if (dirty || editing) { e.preventDefault(); e.returnValue = ''; } });
if (!globalThis.XLSX) toast('No se cargó el lector Excel. Recarga la página.', true);
window.HojaData = { parseCell, rawCell, displayCell, parseClipboard, address, column, decode };
window.HojaSmall = { release: () => { finishEdit(); if (!mayReplace()) return false; dirty = false; return true; } };
