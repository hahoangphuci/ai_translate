#!/usr/bin/env bash
# Minimal Azure startup — do NOT apt-install here (causes 503/timeouts).
# DOCX->PDF on Linux uses LibreOffice only if preinstalled; otherwise pipeline returns translated DOCX.

ROOT="/home/site/wwwroot"
if [[ -d "$ROOT" ]]; then
  cd "$ROOT"
fi

export PYTHONPATH="${ROOT}/python_packages:${PYTHONPATH:-}"

if command -v soffice >/dev/null 2>&1; then
  export LIBREOFFICE_PATH="$(command -v soffice)"
  echo "[startup] LibreOffice: ${LIBREOFFICE_PATH}"
else
  echo "[startup] LibreOffice not found — PDF jobs may deliver translated DOCX fallback"
fi

exec python api_base/run_api.py
