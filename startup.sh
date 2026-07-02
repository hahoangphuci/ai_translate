#!/bin/bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

if [ -f "$ROOT/antenv/bin/activate" ]; then
  # shellcheck disable=SC1091
  source "$ROOT/antenv/bin/activate"
elif [ -f "/home/site/wwwroot/antenv/bin/activate" ]; then
  # shellcheck disable=SC1091
  source "/home/site/wwwroot/antenv/bin/activate"
fi

cd "$ROOT/api_base"
export BACKEND_PORT="${WEBSITES_PORT:-8000}"
export BACKEND_AUTO_OPEN_BROWSER=0

echo "[startup] python=$(command -v python || true)"
python -c "import pdf2docx; print('[startup] pdf2docx OK')" || {
  echo "[startup] ERROR: pdf2docx missing — reinstall dependencies in CI artifact"
  exit 1
}

exec python run_api.py
