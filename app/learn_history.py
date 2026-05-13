"""
学习入库本地审计：追加 JSONL，供页面与排障追溯（不含客户整单原文）。
"""
from __future__ import annotations

import json
import os
import time
from typing import Any

DEFAULT_JSONL_PATH = os.path.join("data", "learn_history.jsonl")


def _ensure_parent_dir(path: str) -> None:
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)


def append_learn_event(record: dict, path: str | None = None) -> None:
    """追加一行 JSON（含 ts）。"""
    p = path or DEFAULT_JSONL_PATH
    _ensure_parent_dir(p)
    line = dict(record)
    line.setdefault("ts", time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()))
    with open(p, "a", encoding="utf-8") as f:
        f.write(json.dumps(line, ensure_ascii=False) + "\n")


def read_recent_learn_events(
    limit: int = 50,
    file_hash: str | None = None,
    path: str | None = None,
) -> list[dict[str, Any]]:
    """
    读取最近若干条；若给定 file_hash 则只保留该会话/文件下的记录，再取末尾 limit 条。
    """
    p = path or DEFAULT_JSONL_PATH
    if not os.path.isfile(p):
        return []
    rows: list[dict[str, Any]] = []
    with open(p, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    if file_hash is not None and file_hash != "":
        rows = [r for r in rows if r.get("file_hash") == file_hash]
    if limit > 0 and len(rows) > limit:
        rows = rows[-limit:]
    return rows
