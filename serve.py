"""Serve only this app on localhost, with no upload or API endpoints."""
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import argparse

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8766)
    args = parser.parse_args()
    handler = partial(SimpleHTTPRequestHandler, directory=str(Path(__file__).resolve().parent))
    server = ThreadingHTTPServer(('127.0.0.1', args.port), handler)
    print(f'Hoja: http://127.0.0.1:{args.port}', flush=True)
    server.serve_forever()
