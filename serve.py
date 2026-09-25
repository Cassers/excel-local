"""Loopback-only web editor with optional disk-backed large-file processing."""
import argparse
import json
import secrets
import shutil
import threading
from concurrent.futures import ThreadPoolExecutor
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlsplit
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent
try:
    import large_store as store
except ImportError:
    store = None

class Jobs:
    def __init__(self, directory):
        self.directory = directory
        directory.mkdir(mode=0o700, exist_ok=True)
        self.sessions = {}
        self.pool = ThreadPoolExecutor(max_workers=1)
        self.token = secrets.token_urlsafe(32)

    def create(self, name):
        key = secrets.token_hex(16)
        folder = self.directory / key
        folder.mkdir(mode=0o700)
        session = dict(id=key, folder=folder, name=Path(name).name, state='uploading', stage='Recibiendo archivo…',
                       cancel=threading.Event(), lock=threading.Lock(), revision=0, meta=None, closing=False)
        self.sessions[key] = session
        return session

    def get(self, key):
        if key not in self.sessions:
            raise ValueError('El libro ya no está abierto. Vuelve a abrirlo.')
        return self.sessions[key]

    def public(self, s):
        return {k: s[k] for k in ('id', 'name', 'state', 'stage', 'revision', 'meta', 'error', 'download') if k in s}

    def discard(self, s):
        s['closing'] = True
        s['cancel'].set()
        future = s.get('future')
        if future and not future.done():
            return
        shutil.rmtree(s['folder'], ignore_errors=True)
        self.sessions.pop(s['id'], None)

    def run(self, s, operation):
        s['cancel'].clear()
        s.pop('error', None); s.pop('download', None)
        s['state'] = 'working'; s['stage'] = 'En cola…'
        def work():
            def check():
                if s['cancel'].is_set():
                    raise store.Cancelled()
            def report(message, _count):
                s['stage'] = message
            try:
                check()
                with s['lock']:
                    if operation == 'import':
                        s['stage'] = 'Leyendo estructura del archivo…'
                        s['meta'] = store.import_file(s['source'], s['folder'] / 'data.sqlite', report, check)
                        s['source'].unlink(missing_ok=True)
                    else:
                        kind, sid = operation
                        output = s['folder'] / ('export.' + kind)
                        store.export_file(s['folder'] / 'data.sqlite', s['meta'], output, kind, sid, report, check)
                        check()
                        s['download'] = {'key': secrets.token_urlsafe(32), 'kind': kind}
                s['state'], s['stage'] = 'ready', 'Listo · datos temporales en este equipo'
            except store.Cancelled:
                s['state'] = 'ready' if s['meta'] else 'cancelled'
                s['stage'] = 'Operación cancelada'
            except Exception as error:
                s['state'] = 'ready' if s['meta'] else 'error'
                s['error'] = str(error)
                s['stage'] = 'No se pudo completar la operación'
            finally:
                if s['closing']:
                    shutil.rmtree(s['folder'], ignore_errors=True)
                    self.sessions.pop(s['id'], None)
        s['future'] = self.pool.submit(work)

    def close(self):
        for s in list(self.sessions.values()):
            s['cancel'].set()
        self.pool.shutdown(wait=True, cancel_futures=True)
        for s in list(self.sessions.values()):
            shutil.rmtree(s['folder'], ignore_errors=True)

def make_handler(jobs):
    class Handler(SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=str(ROOT), **kwargs)

        def log_message(self, fmt, *args):
            pass

        def context(self):
            host = self.headers.get('Host', '')
            allowed = (f'127.0.0.1:{self.server.server_port}', f'localhost:{self.server.server_port}')
            if host not in allowed:
                raise PermissionError('Host no permitido.')
            origin = self.headers.get('Origin')
            if origin and origin != 'http://' + host:
                raise PermissionError('Origen no permitido.')
            return urlsplit(self.path)

        def json(self, data, status=200):
            body = json.dumps(data, ensure_ascii=False).encode()
            self.send_response(status)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers(); self.wfile.write(body)

        def authorized(self):
            if not secrets.compare_digest(self.headers.get('X-Hoja-Token', ''), jobs.token):
                raise PermissionError('Sesión local inválida. Recarga la página.')
            if store is None:
                raise ValueError('Inicia el modo grande con ./abrir.sh.')

        def do_GET(self):
            try:
                url = self.context(); query = parse_qs(url.query)
                arg = lambda key, default='': query.get(key, [default])[0]
                if url.path == '/api/health':
                    return self.json({'large': store is not None, 'token': jobs.token})
                if url.path == '/api/download':
                    s = jobs.get(arg('id')); download = s.get('download', {})
                    if not secrets.compare_digest(arg('key'), download.get('key', '!')):
                        raise PermissionError('Descarga inválida o vencida.')
                    file = s['folder'] / ('export.' + download['kind'])
                    name = Path(s['name']).stem + '-editado.' + download['kind']
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/octet-stream')
                    self.send_header('Content-Disposition', "attachment; filename*=UTF-8''" + quote(name))
                    self.send_header('Content-Length', str(file.stat().st_size))
                    self.send_header('Cache-Control', 'no-store')
                    self.end_headers()
                    with file.open('rb') as source:
                        shutil.copyfileobj(source, self.wfile, length=1024 * 1024)
                    return
                if url.path.startswith('/api/'):
                    self.authorized(); s = jobs.get(arg('id'))
                    if url.path == '/api/status':
                        return self.json(jobs.public(s))
                    if url.path == '/api/page':
                        if s['state'] != 'ready':
                            raise ValueError('Espera a que termine la operación.')
                        sid = int(arg('sheet', '0'))
                        if not 0 <= sid < len(s['meta']['sheets']):
                            raise ValueError('Hoja inválida.')
                        return self.json(store.get_page(s['folder'] / 'data.sqlite', s['meta'], sid, int(arg('offset', '0')), int(arg('col', '1')), arg('q')))
                    return self.json({'error': 'Ruta no encontrada'}, 404)
                allowed = {'/', '/index.html', '/icon.svg', '/styles.css', '/app.bundle.js', '/large.js', '/vendor/xlsx.full.min.js'}
                if url.path not in allowed:
                    return self.send_error(404)
                return super().do_GET()
            except (BrokenPipeError, ConnectionResetError):
                pass
            except Exception as error:
                self.json({'error': str(error)}, 403 if isinstance(error, PermissionError) else 400)

        def do_POST(self):
            s = None
            try:
                url = self.context(); self.authorized()
                size = int(self.headers.get('Content-Length', '-1'))
                if size < 0:
                    raise ValueError('Falta el tamaño de la solicitud.')
                if url.path == '/api/upload':
                    name = parse_qs(url.query).get('name', [''])[0]
                    suffix = Path(name).suffix.lower()
                    if suffix not in ('.xlsx', '.csv'):
                        raise ValueError('El modo grande acepta XLSX y CSV. Convierte los XLS a XLSX primero.')
                    if shutil.disk_usage(jobs.directory).free < size * 3 + 128 * 1024 * 1024:
                        raise ValueError('No hay suficiente espacio libre para este archivo y sus temporales.')
                    s = jobs.create(name); s['source'] = s['folder'] / ('input' + suffix)
                    self.connection.settimeout(120)
                    remaining = size
                    with s['source'].open('wb') as out:
                        while remaining:
                            chunk = self.rfile.read(min(1024 * 1024, remaining))
                            if not chunk:
                                raise ValueError('La transferencia se interrumpió.')
                            out.write(chunk); remaining -= len(chunk)
                    if suffix == '.xlsx':
                        with ZipFile(s['source']) as archive:
                            expanded = sum(x.file_size for x in archive.infolist())
                        if shutil.disk_usage(jobs.directory).free < expanded * 3 + 128 * 1024 * 1024:
                            raise ValueError('El XLSX descomprimido necesita más espacio libre en disco.')
                    jobs.run(s, 'import')
                    return self.json(jobs.public(s), 202)
                if size > 8 * 1024 * 1024:
                    raise ValueError('Pega un bloque de datos más pequeño.')
                body = json.loads(self.rfile.read(size)); s = jobs.get(body['id'])
                if url.path == '/api/close':
                    jobs.discard(s); return self.json({'ok': True})
                if url.path == '/api/cancel':
                    s['cancel'].set(); return self.json({'ok': True})
                if s['state'] != 'ready':
                    raise ValueError('Espera a que termine la operación.')
                sid = int(body.get('sheet', 0))
                if not 0 <= sid < len(s['meta']['sheets']):
                    raise ValueError('Hoja inválida.')
                if url.path == '/api/edit':
                    with s['lock']:
                        if body.get('revision') != s['revision']:
                            raise ValueError('La hoja cambió. Recarga su página antes de editar.')
                        before = store.edit_cells(s['folder'] / 'data.sqlite', s['meta'], sid, body['changes'])
                        s['revision'] += 1
                    return self.json({'before': before, 'revision': s['revision']})
                if url.path == '/api/export':
                    kind = body.get('kind')
                    if kind not in ('xlsx', 'csv'):
                        raise ValueError('Formato inválido.')
                    jobs.run(s, (kind, sid)); return self.json(jobs.public(s), 202)
                self.json({'error': 'Ruta no encontrada'}, 404)
            except (BrokenPipeError, ConnectionResetError):
                if s and s['state'] == 'uploading':
                    jobs.discard(s)
            except Exception as error:
                if s and s['state'] == 'uploading':
                    jobs.discard(s)
                self.json({'error': str(error)}, 403 if isinstance(error, PermissionError) else 400)

        def do_HEAD(self):
            self.send_error(405)
    return Handler

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8766)
    args = parser.parse_args()
    jobs = Jobs(ROOT / '.local-data')
    server = ThreadingHTTPServer(('127.0.0.1', args.port), make_handler(jobs))
    print(f'Hoja: http://127.0.0.1:{args.port} · modo grande: {store is not None}', flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close(); jobs.close()
