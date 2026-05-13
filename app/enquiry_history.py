"""
询价单解析历史：本地 JSON 持久化（runs + index），供 Streamlit 历史 Tab 使用。
"""
from __future__ import annotations

import json
import os
import secrets
import tempfile
from datetime import datetime, timezone
from typing import Any

import config

INDEX_NAME = "index.json"


def _runs_dir() -> str:
    return os.path.abspath(config.ENQUIRY_HISTORY_DIR)


def _index_path() -> str:
    return os.path.join(_runs_dir(), INDEX_NAME)


def ensure_history_dir() -> None:
    os.makedirs(_runs_dir(), exist_ok=True)


def _atomic_write_json(path: str, data: Any) -> None:
    d = os.path.dirname(path) or "."
    fd, tmp = tempfile.mkstemp(prefix=".idx_", suffix=".tmp", dir=d)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except Exception:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _read_index() -> list[dict]:
    p = _index_path()
    if not os.path.isfile(p):
        return []
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []


def _write_index(entries: list[dict]) -> None:
    ensure_history_dir()
    _atomic_write_json(_index_path(), entries)


def _run_file(run_id: str) -> str:
    return os.path.join(_runs_dir(), f"{run_id}.json")


def new_run_id() -> str:
    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")
    return f"{ts}_{secrets.token_hex(3)}"


def append_run(original_filename: str, file_hash: str, enquiry_items: list[dict]) -> str:
    """写入一条新解析记录，更新索引（新在前），超出上限则删最旧文件。"""
    ensure_history_dir()
    run_id = new_run_id()
    created_at = datetime.now(timezone.utc).isoformat()
    payload = {
        "run_id": run_id,
        "created_at": created_at,
        "original_filename": original_filename or "",
        "file_hash": file_hash,
        "item_count": len(enquiry_items or []),
        "enquiry_items": enquiry_items or [],
    }
    path = _run_file(run_id)
    _atomic_write_json(path, payload)

    index = _read_index()
    meta = {
        "run_id": run_id,
        "created_at": created_at,
        "original_filename": payload["original_filename"],
        "file_hash": file_hash,
        "item_count": payload["item_count"],
    }
    index.insert(0, meta)

    max_runs = max(1, int(config.ENQUIRY_HISTORY_MAX_RUNS))
    while len(index) > max_runs:
        old = index.pop()
        old_id = old.get("run_id")
        if old_id:
            try:
                os.unlink(_run_file(old_id))
            except OSError:
                pass

    _write_index(index)
    return run_id


def list_runs() -> list[dict]:
    """按时间倒序（新在前）的索引条目列表。"""
    return list(_read_index())


def load_run(run_id: str) -> dict | None:
    path = _run_file(run_id)
    if not os.path.isfile(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


def clear_all_runs() -> int:
    """删除目录下所有 run 文件与索引。返回删除的 run 文件数。"""
    d = _runs_dir()
    if not os.path.isdir(d):
        return 0
    n = 0
    for name in os.listdir(d):
        if name == INDEX_NAME:
            continue
        if not name.endswith(".json"):
            continue
        fp = os.path.join(d, name)
        try:
            os.unlink(fp)
            n += 1
        except OSError:
            pass
    ip = _index_path()
    if os.path.isfile(ip):
        try:
            os.unlink(ip)
        except OSError:
            pass
    return n
