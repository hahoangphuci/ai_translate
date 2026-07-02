#!/bin/bash
set -e

ROOT="/home/site/wwwroot"
cd "$ROOT"

PY="$ROOT/antenv/bin/python"
if [ ! -x "$PY" ]; then
  echo "[startup] Creating antenv..."
  python3 -m venv "$ROOT/antenv"
fi

"$PY" -m pip install --upgrade pip >/dev/null
if ! "$PY" -c "import pdf2docx" >/dev/null 2>&1; then
  echo "[startup] Installing Python dependencies (first run may take a few minutes)..."
  "$PY" -m pip install -r "$ROOT/requirements.txt"
fi

"$PY" -c "import pdf2docx; print('[startup] pdf2docx OK')"

cd "$ROOT/api_base"
export BACKEND_PORT="${WEBSITES_PORT:-8000}"
export BACKEND_AUTO_OPEN_BROWSER=0

echo "[startup] Starting Flask on port ${BACKEND_PORT}..."
exec "$PY" run_api.py
