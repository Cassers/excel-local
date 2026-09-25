"""Reproducible synthetic benchmark; outputs remain ignored by Git."""
import argparse
import json
import resource
import time
from pathlib import Path
from zipfile import ZipFile, ZIP_STORED
from openpyxl import Workbook
import large_store as store

ROOT = Path(__file__).parent / '.test-output' / 'large'
ROOT.mkdir(parents=True, exist_ok=True)

def generate(kind, mb):
    target = ROOT / f'synthetic-{mb}MB.{kind}'
    payload = 'Datos de prueba 0123456789 abcdefghijklmnopqrstuvwxyz ' * 150
    row_count = (mb * 1000000) // (len(payload) * 4) + 1
    if kind == 'csv':
        with target.open('w') as f:
            for i in range(row_count): f.write(f'{i},' + ','.join([payload] * 4) + '\n')
    else:
        template = ROOT / 'template.xlsx'; wb = Workbook(); wb.save(template)
        with ZipFile(template) as source, ZipFile(target, 'w', ZIP_STORED) as out:
            for name in source.namelist():
                if name != 'xl/worksheets/sheet1.xml': out.writestr(name, source.read(name))
            with out.open('xl/worksheets/sheet1.xml', 'w', force_zip64=True) as sheet:
                sheet.write(b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
                for r in range(1, row_count + 1):
                    line = f'<row r="{r}">' + ''.join(f'<c r="{c}{r}" t="inlineStr"><is><t>{payload}</t></is></c>' for c in 'ABCD') + '</row>'
                    sheet.write(line.encode())
                sheet.write(b'</sheetData></worksheet>')
        template.unlink()
    print(json.dumps({'file': str(target), 'bytes': target.stat().st_size, 'rows': row_count}), flush=True)
    return target

def run(path):
    db = ROOT / (path.name + '.sqlite'); db.unlink(missing_ok=True)
    started = time.monotonic()
    meta = store.import_file(path, db)
    imported = time.monotonic()
    page = store.get_page(db, meta, offset=max(0, meta['sheets'][0]['rows'] - 100))
    assert page['cells']
    store.edit_cells(db, meta, 0, [{'r': 1, 'c': 1, 'cell': {'t': 's', 'v': 'Edición verificada'}}])
    assert store.get_page(db, meta)['cells']['1:1']['v'] == 'Edición verificada'
    exported = ROOT / (path.stem + '-export.csv')
    store.export_file(db, meta, exported, 'csv')
    with exported.open(encoding='utf-8-sig') as f: assert f.readline().startswith('Edición verificada')
    result = {'file': path.name, 'bytes': path.stat().st_size, 'rows': meta['sheets'][0]['rows'],
              'import_seconds': round(imported - started, 2), 'export_and_edit_seconds': round(time.monotonic() - imported, 2),
              'max_rss_mb': round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024, 1), 'database_mb': round(db.stat().st_size / 1e6, 1)}
    (ROOT / (path.name + '.result.json')).write_text(json.dumps(result, indent=2))
    print(json.dumps(result), flush=True)
    exported.unlink(); db.unlink()
    for suffix in ('-wal', '-shm'): Path(str(db) + suffix).unlink(missing_ok=True)

if __name__ == '__main__':
    p = argparse.ArgumentParser(); p.add_argument('action', choices=['generate', 'run']); p.add_argument('--kind', choices=['csv', 'xlsx'], default='csv'); p.add_argument('--mb', type=int, default=100); args = p.parse_args()
    target = ROOT / f'synthetic-{args.mb}MB.{args.kind}'
    if args.action == 'generate': generate(args.kind, args.mb)
    else: run(target)
