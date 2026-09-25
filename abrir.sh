#!/usr/bin/env bash
set -euo pipefail
cd -- "$(dirname -- "${BASH_SOURCE[0]}")"
if [[ ! -x .venv/bin/python ]]; then
  if command -v uv >/dev/null 2>&1; then
    uv venv .venv
  else
    python3 -m venv .venv
  fi
fi
if ! .venv/bin/python -c 'import openpyxl, lxml, defusedxml' >/dev/null 2>&1; then
  if command -v uv >/dev/null 2>&1; then
    uv pip install --python .venv/bin/python -r requirements.txt
  else
    .venv/bin/python -m pip install -r requirements.txt
  fi
fi
port="${HOJA_PORT:-8768}"
mkdir -p -m 700 .local-data/tmp
export TMPDIR="$PWD/.local-data/tmp"
printf 'Abre http://127.0.0.1:%s en tu navegador. Ctrl+C cierra el servicio.\n' "$port"
exec .venv/bin/python serve.py --port "$port"
