"""Persistent document translation job status (survives Azure restarts / multi-instance)."""

from __future__ import annotations

import json
import os
import re
import threading
from typing import Any, Optional

_lock = threading.Lock()
_JOB_ID_RE = re.compile(r"^[0-9a-fA-F-]{36}$")


def _jobs_dir() -> str:
    base = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    path = os.path.join(base, "utils", "jobs")
    os.makedirs(path, exist_ok=True)
    return path


def _job_path(job_id: str) -> str:
    if not _JOB_ID_RE.match(job_id or ""):
        raise ValueError("invalid job_id")
    return os.path.join(_jobs_dir(), f"{job_id}.json")


def save(job_id: str, data: dict[str, Any]) -> None:
    payload = dict(data or {})
    with _lock:
        tmp = f"{_job_path(job_id)}.tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False)
        os.replace(tmp, _job_path(job_id))


def load(job_id: str) -> Optional[dict[str, Any]]:
    if not _JOB_ID_RE.match(job_id or ""):
        return None
    path = _job_path(job_id)
    if not os.path.isfile(path):
        return None
    try:
        with _lock:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        return data if isinstance(data, dict) else None
    except Exception:
        return None


class PersistingJobDict(dict):
    """In-memory job view that writes through to disk on each mutation."""

    def __init__(self, job_id: str, initial: dict[str, Any]):
        super().__init__(initial or {})
        self._job_id = job_id
        save(job_id, self)

    def __setitem__(self, key, value) -> None:
        super().__setitem__(key, value)
        save(self._job_id, self)

    def update(self, *args, **kwargs) -> None:
        super().update(*args, **kwargs)
        save(self._job_id, self)
