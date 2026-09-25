/* Local API client: only the visible page and bounded undo history live in JS. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id), D = window.HojaData;
  let token = '', session = null, sid = 0, offset = 0, col = 1, page = null;
  let selection = { r: 1, c: 1 }, undo = [], redo = [], dirty = false, xhr = null, working = false, sequence = 0;
  const message = text => { $('large-message').textContent = text; };
  const busy = (text, cancel = true) => { working = true; $('large-progress').textContent = text; $('large-busy').hidden = false; $('large-cancel').hidden = !cancel; };
  const idle = () => { working = false; $('large-busy').hidden = true; };
  async function attempt(action) {
    if (working) return;
    busy('Procesando…', false);
    try { await action(); } catch (error) { message(error.message); alert(error.message); } finally { idle(); }
  }
  async function api(path, body) {
    const response = await fetch('/api/' + path, { method: body ? 'POST' : 'GET', headers: { 'X-Hoja-Token': token, ...(body ? { 'Content-Type': 'application/json' } : {}) }, ...(body ? { body: JSON.stringify(body) } : {}) });
    const data = await response.json(); if (!response.ok) throw Error(data.error || 'No se pudo completar la operación.'); return data;
  }
  async function available() {
    if (!/^https?:$/.test(location.protocol)) return false;
    try { const state = await api('health'); token = state.token; return state.large; } catch { return false; }
  }
  function upload(file) {
    return new Promise((resolve, reject) => {
      xhr = new XMLHttpRequest(); xhr.open('POST', '/api/upload?name=' + encodeURIComponent(file.name));
      xhr.setRequestHeader('X-Hoja-Token', token); xhr.setRequestHeader('Content-Type', 'application/octet-stream');
      xhr.upload.onprogress = e => { $('large-progress').textContent = `Copiando a tu disco: ${e.lengthComputable ? Math.round(e.loaded / e.total * 100) + '%' : '…'}`; };
      xhr.onload = () => { const status = xhr.status; try { const data = JSON.parse(xhr.responseText); status < 300 ? resolve(data) : reject(Error(data.error)); } catch { reject(Error('El servicio local no respondió correctamente.')); } xhr = null; };
      xhr.onerror = () => { xhr = null; reject(Error('No se pudo conectar con el servicio local.')); };
      xhr.onabort = () => { xhr = null; reject(Error('Transferencia cancelada.')); };
      xhr.send(file);
    });
  }
  async function waitJob() {
    while (session.state === 'working') {
      $('large-progress').textContent = session.stage;
      await new Promise(resolve => setTimeout(resolve, 600));
      session = await api('status?id=' + session.id);
    }
    if (session.error || session.state === 'error') throw Error(session.error || 'No se pudo leer el archivo.');
    if (session.state === 'cancelled') throw Error('Importación cancelada.');
  }
  async function open(file) {
    if (!file || working) return;
    if (!(await available())) { $('large-help').showModal(); return; }
    if (!/\.(xlsx|csv)$/i.test(file.name)) { alert('El modo grande acepta XLSX y CSV. Convierte los XLS a XLSX primero.'); return; }
    if (dirty && !confirm('Hay cambios sin descargar. ¿Cerrar este libro y abrir el nuevo?')) return;
    if (!window.HojaSmall.release()) return;
    await attempt(async () => {
      if (session) await api('close', { id: session.id });
      session = null; dirty = false; $('large-editor').hidden = true; $('welcome').hidden = false; busy('Copiando el archivo a tu disco…');
      session = await upload(file);
      try { await waitJob(); } catch (error) { await api('close', { id: session.id }); session = null; throw error; }
      sid = 0; offset = 0; col = 1; undo = []; redo = []; dirty = false; selection = { r: 1, c: 1 };
      $('welcome').hidden = true; $('editor').hidden = true; $('large-editor').hidden = false;
      $('large-name').textContent = file.name; $('large-status').textContent = 'Modo grande · datos temporales en disco'; $('large-search').value = '';
      $('large-sheet').replaceChildren(...session.meta.sheets.map((sheet, index) => { const option = document.createElement('option'); option.value = index; option.textContent = sheet.name; return option; }));
      await refresh();
    });
    $('large-file').value = ''; $('file-input').value = '';
  }
  function display(cell) {
    if (cell?.t === 'd') return cell.v;
    if (cell?.tableFormula) return '[Fórmula de tabla]';
    if (cell?.t === 'e') return String(cell.v);
    return D.displayCell(cell);
  }
  async function refresh() {
    const request = ++sequence;
    const params = new URLSearchParams({ id: session.id, sheet: sid, offset, col, q: $('large-search').value });
    message('Cargando página…');
    const result = await api('page?' + params); if (request !== sequence) return;
    page = result; col = result.col;
    const fragment = document.createDocumentFragment(), head = document.createElement('thead'), hr = document.createElement('tr'); hr.append(document.createElement('th'));
    for (let c = col; c < col + page.cols; c++) { const th = document.createElement('th'); th.textContent = D.column(c - 1); hr.append(th); }
    head.append(hr); fragment.append(head); const body = document.createElement('tbody');
    for (const r of page.rows) {
      const row = document.createElement('tr'), label = document.createElement('th'); label.textContent = r; row.append(label);
      for (let c = col; c < col + page.cols; c++) {
        const cell = page.cells[`${r}:${c}`], td = document.createElement('td'); td.dataset.r = r; td.dataset.c = c; td.dataset.largeCell = D.address(r - 1, c - 1); td.tabIndex = -1;
        td.textContent = display(cell); if (cell?.t === 'n' && !cell.f) td.classList.add('numeric'); if (cell?.f) td.classList.add('formula'); row.append(td);
      }
      body.append(row);
    }
    fragment.append(body); $('large-grid').replaceChildren(fragment);
    const info = session.meta.sheets[sid];
    if (!page.rows.includes(selection.r) || selection.c < col || selection.c >= col + page.cols) selection = { r: page.rows[0] || 1, c: col };
    select();
    $('large-dimensions').textContent = `${info.rows.toLocaleString('es')} filas × ${info.cols.toLocaleString('es')} columnas · solo esta página en memoria`;
    $('large-col-label').textContent = `${D.column(col - 1)}–${D.column(col + page.cols - 2)}`;
    $('large-page-label').textContent = page.rows.length ? `Filas ${page.rows[0]}–${page.rows.at(-1)}` : 'Sin coincidencias';
    $('large-prev').disabled = offset === 0; $('large-next').disabled = !page.more;
    $('large-col-prev').disabled = col === 1; $('large-col-next').disabled = col + page.cols > info.cols;
    $('large-undo').disabled = !undo.length; $('large-redo').disabled = !redo.length; message('');
  }
  function select() {
    $('large-grid').querySelectorAll('.selected').forEach(td => { td.classList.remove('selected'); td.tabIndex = -1; });
    const key = D.address(selection.r - 1, selection.c - 1), td = $('large-grid').querySelector(`[data-large-cell="${key}"]`);
    if (td) { td.classList.add('selected'); td.tabIndex = 0; td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
    $('large-address').value = key; $('large-value').value = D.rawCell(page?.cells[`${selection.r}:${selection.c}`]);
  }
  async function edit(changes, record = true) {
    const response = await api('edit', { id: session.id, sheet: sid, revision: session.revision, changes }); session.revision = response.revision;
    if (record) { undo.push({ sid, changes, before: response.before }); if (undo.length > 20) undo.shift(); redo = []; }
    dirty = true; $('large-status').textContent = 'Cambios en disco temporal · descarga para conservarlos'; await refresh();
  }
  async function saveValue() {
    const previous = page?.cells[`${selection.r}:${selection.c}`], value = $('large-value').value;
    if (value === D.rawCell(previous)) return;
    busy('Guardando cambio…', false); await edit([{ ...selection, cell: D.parseCell(value) || null }]);
  }
  async function history(backward) {
    const from = backward ? undo : redo, to = backward ? redo : undo, item = from.at(-1); if (!item) return;
    busy('Aplicando cambio…', false); sid = item.sid; $('large-sheet').value = sid;
    await edit(backward ? item.before : item.changes, false); from.pop(); to.push(item); await refresh();
  }
  async function download(kind) {
    busy('Preparando la copia en disco…'); session = await api('export', { id: session.id, sheet: sid, kind }); await waitJob();
    if (!session.download) { message(session.stage); return; }
    const link = document.createElement('a'); link.href = '/api/download?' + new URLSearchParams({ id: session.id, key: session.download.key }); link.download = ''; document.body.append(link); link.click(); link.remove();
    if (kind === 'xlsx') { dirty = false; $('large-status').textContent = 'Copia generada · revisa tus descargas'; } message('Descarga iniciada');
  }
  $('large-button').onclick = () => attempt(async () => { if (await available()) $('large-file').click(); else $('large-help').showModal(); });
  $('large-help-close').onclick = () => $('large-help').close();
  $('large-file').onchange = e => open(e.target.files[0]);
  $('large-grid').onclick = e => { const td = e.target.closest('td'); if (td) { selection = { r: +td.dataset.r, c: +td.dataset.c }; select(); td.focus(); } };
  $('large-grid').ondblclick = e => { if (e.target.closest('td')) { $('large-value').focus(); $('large-value').select(); } };
  $('large-grid').onkeydown = e => {
    if (e.ctrlKey || e.metaKey) return;
    if (e.key === 'Delete') { e.preventDefault(); $('large-clear').click(); }
    else if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); $('large-value').focus(); $('large-value').select(); }
    else if (e.key.length === 1) { e.preventDefault(); $('large-value').focus(); $('large-value').value = e.key; }
  };
  $('large-grid').onpaste = e => {
    e.preventDefault(); const text = e.clipboardData.getData('text/plain');
    if ($('large-search').value) { alert('Quita el filtro antes de pegar tablas.'); return; }
    if (text.length > 5 * 1024 * 1024) { alert('Pega bloques de menos de 5 MB.'); return; }
    attempt(async () => { const rows = D.parseClipboard(text), changes = [];
      rows.forEach((row, ri) => row.forEach((value, ci) => changes.push({ r: selection.r + ri, c: selection.c + ci, cell: D.parseCell(value) || null })));
      busy('Pegando datos…', false); await edit(changes);
    });
  };
  $('large-save').onclick = () => attempt(saveValue);
  $('large-value').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); attempt(saveValue); } if (e.key === 'Escape') select(); };
  $('large-clear').onclick = () => attempt(async () => { busy('Borrando celda…', false); await edit([{ ...selection, cell: null }]); });
  $('large-undo').onclick = () => attempt(() => history(true)); $('large-redo').onclick = () => attempt(() => history(false));
  $('large-xlsx').onclick = () => attempt(() => download('xlsx')); $('large-csv').onclick = () => attempt(() => download('csv'));
  $('large-sheet').onchange = () => attempt(async () => { sid = +$('large-sheet').value; offset = 0; col = 1; $('large-search').value = ''; await refresh(); });
  let timer;
  $('large-search').oninput = () => { clearTimeout(timer); timer = setTimeout(() => attempt(async () => { offset = 0; await refresh(); }), 350); };
  $('large-prev').onclick = () => attempt(async () => { offset = Math.max(0, offset - 100); await refresh(); });
  $('large-next').onclick = () => attempt(async () => { offset += 100; await refresh(); });
  $('large-col-prev').onclick = () => attempt(async () => { col = Math.max(1, col - 24); await refresh(); });
  $('large-col-next').onclick = () => attempt(async () => { col += 24; await refresh(); });
  $('large-address').onkeydown = e => {
    if (e.key !== 'Enter') return; e.preventDefault();
    attempt(async () => { const value = $('large-address').value.trim().toUpperCase(); if (!/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(value)) throw Error('Referencia inválida.');
      const p = D.decode(value), info = session.meta.sheets[sid]; if (p.r >= info.rows || p.c >= info.cols) throw Error('La celda está fuera de la hoja.');
      selection = { r: p.r + 1, c: p.c + 1 }; offset = Math.floor(p.r / 100) * 100; col = Math.floor(p.c / 24) * 24 + 1; $('large-search').value = ''; await refresh();
    });
  };
  $('large-cancel').onclick = async () => { if (xhr) xhr.abort(); else if (session) { try { await api('cancel', { id: session.id }); $('large-progress').textContent = 'Cancelando…'; } catch (error) { message(error.message); } } };
  $('large-close').onclick = () => attempt(async () => {
    await release();
  });
  async function release() {
    if (!session) return true;
    if (dirty && !confirm('Hay cambios sin descargar. ¿Cerrar y borrar los datos temporales?')) return false;
    await api('close', { id: session.id }); session = null; dirty = false; $('large-editor').hidden = true; $('welcome').hidden = false;
    return true;
  }
  document.addEventListener('keydown', e => {
    if ($('large-editor').hidden || !(e.ctrlKey || e.metaKey)) return;
    if (e.key.toLowerCase() === 's') { e.preventDefault(); attempt(() => download('xlsx')); }
    if (e.target.matches('input')) return;
    if (e.key.toLowerCase() === 'z') { e.preventDefault(); attempt(() => history(!e.shiftKey)); }
  });
  window.addEventListener('beforeunload', e => { if (dirty || working) { e.preventDefault(); e.returnValue = ''; } });
  window.LargeEditor = { open, release };
})();
