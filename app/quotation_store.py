"""
报价单持久化：索引 + 单条 JSON + 归档 Excel，供「我的报价」使用。
"""
from __future__ import annotations

import json
import os
import secrets
import shutil
import tempfile
from datetime import datetime, timezone
from typing import Any

import config

INDEX_NAME = "index.json"
EXCEL_FILENAME = "quotation.xlsx"

QUOTATION_STATUSES: tuple[str, ...] = ("draft", "ready", "sent")

STATUS_LABELS_CN: dict[str, str] = {
    "draft": "草稿",
    "ready": "待发送",
    "sent": "已发送",
}

# 历史版本曾写入的 status，打开「我的报价」列表时迁移为三态之一
_LEGACY_STATUS_TO_CANONICAL: dict[str, str] = {
    "confirming": "sent",
    "won": "sent",
    "lost": "sent",
    "closed": "sent",
    "expired": "sent",
}


def _store_dir() -> str:
    return os.path.abspath(config.QUOTATION_STORE_DIR)


def _index_path() -> str:
    return os.path.join(_store_dir(), INDEX_NAME)


def _quotation_json_path(quotation_id: str) -> str:
    return os.path.join(_store_dir(), f"{quotation_id}.json")


def _quotation_subdir(quotation_id: str) -> str:
    return os.path.join(_store_dir(), quotation_id)


def ensure_store_dir() -> None:
    os.makedirs(_store_dir(), exist_ok=True)


def _atomic_write_json(path: str, data: Any) -> None:
    d = os.path.dirname(path) or "."
    fd, tmp = tempfile.mkstemp(prefix=".q_", suffix=".tmp", dir=d)
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
    ensure_store_dir()
    _atomic_write_json(_index_path(), entries)


def new_quotation_id() -> str:
    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")
    return f"{ts}_{secrets.token_hex(3)}"


def load_quotation(quotation_id: str) -> dict | None:
    path = _quotation_json_path(quotation_id)
    if not os.path.isfile(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            rec = json.load(f)
        return rec if isinstance(rec, dict) else None
    except (json.JSONDecodeError, OSError):
        return None


def _meta_from_record(rec: dict) -> dict:
    return {
        "quotation_id": rec["quotation_id"],
        "owner": rec.get("owner") or "",
        "status": rec.get("status") or "draft",
        "created_at": rec.get("created_at") or "",
        "updated_at": rec.get("updated_at") or "",
        "original_filename": rec.get("original_filename") or "",
        "file_hash": rec.get("file_hash") or "",
        "item_count": int(rec.get("item_count") or len(rec.get("enquiry_items") or [])),
        "display_title": rec.get("display_title") or "",
        "valid_until": rec.get("valid_until"),
        "has_excel": bool((rec.get("exported_excel_relpath") or "").strip()),
    }


def canonical_quotation_status(raw: str | None) -> str:
    """将磁盘上的 status 规范为当前允许的三态之一。"""
    s = (raw or "").strip() or "draft"
    if s in QUOTATION_STATUSES:
        return s
    return _LEGACY_STATUS_TO_CANONICAL.get(s, "draft")


def migrate_legacy_quotation_statuses() -> None:
    """打开列表前调用：把旧版多状态写入的记录迁移为三态并写回磁盘。"""
    for meta in _read_index():
        qid = meta.get("quotation_id")
        if not qid:
            continue
        rec = load_quotation(qid)
        if not rec:
            continue
        prev = (rec.get("status") or "draft").strip() or "draft"
        new_s = canonical_quotation_status(prev)
        if new_s != prev:
            rec["status"] = new_s
            rec["updated_at"] = datetime.now(timezone.utc).isoformat()
            save_quotation_record(rec, update_index=True)


def save_quotation_record(record: dict, update_index: bool = True) -> None:
    ensure_store_dir()
    qid = record["quotation_id"]
    _atomic_write_json(_quotation_json_path(qid), record)
    if not update_index:
        return
    meta = _meta_from_record(record)
    idx = _read_index()
    out: list[dict] = []
    found = False
    for e in idx:
        if e.get("quotation_id") == qid:
            out.append(meta)
            found = True
        else:
            out.append(e)
    if not found:
        out.insert(0, meta)
    _write_index(out)


def _trim_store_if_needed() -> None:
    max_n = max(1, int(config.QUOTATION_STORE_MAX_RECORDS))
    idx = _read_index()
    if len(idx) <= max_n:
        return
    # 按 updated_at 升序删最旧
    keyed = [(e.get("updated_at") or e.get("created_at") or "", e) for e in idx]
    keyed.sort(key=lambda x: x[0])
    while len(keyed) > max_n:
        _, old = keyed.pop(0)
        oid = old.get("quotation_id")
        if not oid:
            continue
        try:
            os.unlink(_quotation_json_path(oid))
        except OSError:
            pass
        sub = _quotation_subdir(oid)
        if os.path.isdir(sub):
            try:
                shutil.rmtree(sub, ignore_errors=True)
            except OSError:
                pass
    _write_index([e for _, e in keyed])


def upsert_quotation(
    owner: str,
    enquiry_items: list[dict],
    match_results: list[dict],
    file_hash: str,
    original_filename: str,
    source_run_id: str | None,
    status: str,
    excel_src_path: str | None,
    quotation_id: str | None = None,
    customer_name: str | None = None,
    valid_until: str | None = None,
) -> str | None:
    """
    新建或更新报价单。owner 为空时返回 None。
    excel_src_path 为 None 时不更新已归档 Excel（保留已有路径）；新建且无已有路径则为空。
    """
    owner = (owner or "").strip()
    if not owner:
        return None
    if status not in QUOTATION_STATUSES:
        status = "draft"

    ensure_store_dir()
    now = datetime.now(timezone.utc).isoformat()
    existing: dict | None = None
    qid = quotation_id
    if qid:
        existing = load_quotation(qid)
        if existing and (existing.get("owner") or "").strip() != owner:
            return None
    if existing:
        created = existing.get("created_at") or now
        qid = existing["quotation_id"]
    else:
        created = now
        qid = new_quotation_id()

    display_title = ""
    for it in enquiry_items or []:
        if isinstance(it, dict):
            display_title = (it.get("title") or it.get("item_title") or "")[:200]
            if display_title.strip():
                break
    if not (display_title or "").strip():
        display_title = (original_filename or "未命名询价")[:200]

    rel_excel = (existing or {}).get("exported_excel_relpath") or ""
    if excel_src_path and os.path.isfile(excel_src_path):
        sub = _quotation_subdir(qid)
        os.makedirs(sub, exist_ok=True)
        dest = os.path.join(sub, EXCEL_FILENAME)
        shutil.copy2(excel_src_path, dest)
        rel_excel = f"{qid}/{EXCEL_FILENAME}"

    prev_cust = (existing or {}).get("customer_name") or ""
    prev_vu = (existing or {}).get("valid_until")
    record = {
        "quotation_id": qid,
        "owner": owner,
        "status": status,
        "created_at": created,
        "updated_at": now,
        "source_run_id": source_run_id if source_run_id is not None else (existing or {}).get("source_run_id") or "",
        "original_filename": original_filename or (existing or {}).get("original_filename") or "",
        "file_hash": file_hash or (existing or {}).get("file_hash") or "",
        "customer_name": (customer_name if customer_name is not None else prev_cust) or "",
        "valid_until": valid_until if valid_until is not None else prev_vu,
        "display_title": display_title,
        "item_count": len(enquiry_items or []),
        "enquiry_items": list(enquiry_items or []),
        "match_results": list(match_results or []),
        "exported_excel_relpath": rel_excel or "",
    }
    save_quotation_record(record, update_index=True)
    _trim_store_if_needed()
    return qid


def list_quotation_meta_for_owner(owner: str, status_filter: str | None) -> list[dict]:
    migrate_legacy_quotation_statuses()
    owner = (owner or "").strip()
    rows = [e for e in _read_index() if (e.get("owner") or "").strip() == owner]
    if status_filter and status_filter != "all":
        rows = [e for e in rows if e.get("status") == status_filter]
    rows.sort(key=lambda x: x.get("updated_at") or "", reverse=True)
    return rows


def status_counts_for_owner(owner: str) -> dict[str, int]:
    migrate_legacy_quotation_statuses()
    owner = (owner or "").strip()
    rows = [e for e in _read_index() if (e.get("owner") or "").strip() == owner]
    counts: dict[str, int] = {s: 0 for s in QUOTATION_STATUSES}
    counts["all"] = len(rows)
    for e in rows:
        s = e.get("status") or "draft"
        if s in counts:
            counts[s] += 1
    return counts


def set_quotation_status(quotation_id: str, owner: str, new_status: str) -> bool:
    owner = (owner or "").strip()
    if not owner or new_status not in QUOTATION_STATUSES:
        return False
    rec = load_quotation(quotation_id)
    if not rec or (rec.get("owner") or "").strip() != owner:
        return False
    rec["status"] = new_status
    rec["updated_at"] = datetime.now(timezone.utc).isoformat()
    save_quotation_record(rec, update_index=True)
    return True


def exported_excel_abs_path(record: dict) -> str | None:
    rel = (record.get("exported_excel_relpath") or "").strip()
    if not rel:
        return None
    p = os.path.join(_store_dir(), rel.replace("\\", "/"))
    return p if os.path.isfile(p) else None
