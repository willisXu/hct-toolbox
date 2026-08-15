# -*- coding: utf-8 -*-
"""已轉出紀錄 — 記住哪些單號已經轉出過出貨格式表格。

「NetSuite 直接抓取」分頁每次轉換成功，就把該批資料的單號寫進紀錄；
下次抓到同一張單時，表格會標記「已轉出」，勾選要轉換時再額外提醒，
避免同一張單被重複轉出、重複出貨。

紀錄以 JSON 檔保存（預設 ``data/ns_export_log.json``）：

    {"version": 1, "records": {"SO12345": {"count": 2, "first_at": "...",
      "last_at": "...", "last_format": "HCT 銷貨報表格式", "formats": [...]}}}

注意：Streamlit Community Cloud 的檔案系統是暫時的，容器重啟後紀錄會清空。
介面提供下載／匯入紀錄檔，需要長期保存時請自行下載備份再匯入。
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from .xlio import build_header_map, normalize_text

LOG_VERSION = 1

# 紀錄上限：超過就丟掉最舊的（依最後轉出時間），避免檔案無限長大。
MAX_RECORDS = 5000

# 單號欄候選（依序取第一個存在的欄）。文件編號是使用者認得的單號；
# 找不到才退回內部 ID，至少還能唯一識別一張單。
_KEY_ALIASES = ("文件編號", "單號", "交易編號", "document number", "內部 id", "內部id")

_TIME_FORMAT = "%Y-%m-%d %H:%M"


def key_column_index(header_row: list) -> int | None:
    """回傳單號欄的欄索引；報表沒有任何可用單號欄時回 None。"""
    headers = build_header_map(header_row)
    for alias in _KEY_ALIASES:
        index = headers.get(alias.casefold())
        if index is not None:
            return index
    return None


def extract_keys(rows: list[list[object]]) -> list[str]:
    """回傳每一筆資料列（rows[1:]）對應的單號；該列沒有單號就是空字串。"""
    if not rows:
        return []
    index = key_column_index(rows[0])
    if index is None:
        return ["" for _ in rows[1:]]
    return [
        normalize_text(row[index]) if index < len(row) else ""
        for row in rows[1:]
    ]


# ------------------------------------------------------------------ 讀寫

def empty_log() -> dict:
    return {"version": LOG_VERSION, "records": {}}


def load(path: str | Path) -> dict:
    """讀取紀錄檔；檔案不存在或內容壞掉都回空紀錄（紀錄壞掉不該擋住轉換）。"""
    try:
        raw = Path(path).read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return empty_log()
    return _coerce(raw)


def loads(raw: str | bytes) -> dict:
    """解析上傳的紀錄檔內容；格式不對時丟 ValueError 讓呼叫端提示使用者。"""
    if isinstance(raw, bytes):
        try:
            raw = raw.decode("utf-8-sig")
        except UnicodeDecodeError as exc:
            raise ValueError("紀錄檔不是 UTF-8 文字檔。") from exc
    log = _coerce(raw, strict=True)
    return log


def _coerce(raw: str, strict: bool = False) -> dict:
    try:
        data = json.loads(raw)
    except ValueError as exc:
        if strict:
            raise ValueError(f"紀錄檔不是有效的 JSON：{exc}") from exc
        return empty_log()
    if not isinstance(data, dict) or not isinstance(data.get("records"), dict):
        if strict:
            raise ValueError("紀錄檔格式不對，應該是 {\"version\": 1, \"records\": {...}}。")
        return empty_log()
    records = {
        str(key): value
        for key, value in data["records"].items()
        if isinstance(value, dict) and str(key)
    }
    return {"version": LOG_VERSION, "records": records}


def save(path: str | Path, log: dict) -> None:
    """寫回紀錄檔（先寫暫存檔再取代，避免中途失敗留下半個壞檔）。"""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_suffix(target.suffix + ".tmp")
    tmp.write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(target)


def dumps(log: dict) -> bytes:
    return json.dumps(log, ensure_ascii=False, indent=1).encode("utf-8")


# ------------------------------------------------------------------ 更新

def mark(log: dict, keys, output_format: str, when: datetime | None = None) -> dict:
    """把 keys 標記為已轉出（同一批重複的單號只算一次），回傳更新後的 log。"""
    stamp = (when or datetime.now()).strftime(_TIME_FORMAT)
    records = log.setdefault("records", {})
    for key in dict.fromkeys(k for k in keys if k):  # 去重、保序
        record = records.get(key)
        if not isinstance(record, dict):
            record = {"count": 0, "first_at": stamp, "formats": []}
            records[key] = record
        record["count"] = int(record.get("count") or 0) + 1
        record["last_at"] = stamp
        record["last_format"] = output_format
        formats = record.get("formats")
        if not isinstance(formats, list):
            formats = []
        if output_format not in formats:
            formats.append(output_format)
        record["formats"] = formats
    _prune(records)
    log["version"] = LOG_VERSION
    return log


def _prune(records: dict) -> None:
    if len(records) <= MAX_RECORDS:
        return
    ordered = sorted(records.items(), key=lambda kv: str(kv[1].get("last_at") or ""))
    for key, _ in ordered[: len(records) - MAX_RECORDS]:
        records.pop(key, None)


def record_export(
    path: str | Path, keys, output_format: str, when: datetime | None = None
) -> dict:
    """讀檔 → 標記 → 寫回。寫檔失敗會往上丟，由呼叫端提示（轉換結果不受影響）。"""
    log = mark(load(path), keys, output_format, when)
    save(path, log)
    return log


def merge(base: dict, other: dict) -> dict:
    """合併兩份紀錄（匯入備份用）：同一單號取次數相加、時間取較晚的那筆。"""
    result = {"version": LOG_VERSION, "records": dict(base.get("records") or {})}
    for key, incoming in (other.get("records") or {}).items():
        current = result["records"].get(key)
        if not isinstance(current, dict):
            result["records"][key] = dict(incoming)
            continue
        merged = dict(current)
        merged["count"] = int(current.get("count") or 0) + int(incoming.get("count") or 0)
        first_times = [t for t in (current.get("first_at"), incoming.get("first_at")) if t]
        if first_times:
            merged["first_at"] = min(first_times)
        if str(incoming.get("last_at") or "") > str(current.get("last_at") or ""):
            merged["last_at"] = incoming.get("last_at")
            merged["last_format"] = incoming.get("last_format")
        formats = list(current.get("formats") or [])
        for fmt in incoming.get("formats") or []:
            if fmt not in formats:
                formats.append(fmt)
        merged["formats"] = formats
        result["records"][key] = merged
    _prune(result["records"])
    return result


def clear(path: str | Path) -> dict:
    log = empty_log()
    save(path, log)
    return log


# ------------------------------------------------------------------ 顯示

def short_label(record: dict) -> str:
    """表格裡的精簡標記，例如「⚠️ 已轉出 ×2（08/14 15:03）」。"""
    count = int(record.get("count") or 0)
    stamp = str(record.get("last_at") or "")
    when = stamp[5:] if len(stamp) >= 11 else stamp  # 去掉年份，欄寬有限
    return f"⚠️ 已轉出 ×{count}（{when}）" if when else f"⚠️ 已轉出 ×{count}"


def describe(record: dict) -> str:
    """提醒清單裡的完整說明。"""
    count = int(record.get("count") or 0)
    parts = [f"已轉出 {count} 次"]
    if record.get("last_at"):
        parts.append(f"最後 {record['last_at']}")
    formats = record.get("formats") or []
    if formats:
        parts.append("／".join(str(f) for f in formats))
    return "｜".join(parts)
