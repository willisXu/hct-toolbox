# -*- coding: utf-8 -*-
"""「已轉出紀錄」的存放後端。

兩種後端，介面一樣（load / save / record / clear）：

  * ``LocalStore``  存本機 JSON 檔。沒設定 Google 試算表時用它；設定了的話
    它就退居「寫入失敗時的待補紀錄」（見 app.py 的補寫流程）。
  * ``SheetStore``  存 Google 試算表（一列一張單）。雲端部署重啟也不會掉，
    而且可以直接開試算表查、手動修。

紀錄本體的結構與合併邏輯都在 ``export_log``，這裡只負責「存到哪裡」。
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from . import export_log
from .gsheet import SheetClient, SheetError

# 試算表的欄位標題（第 1 列）。認欄名不認順序，使用者手動調欄位也不會壞。
SHEET_HEADER = ["單號", "轉出次數", "首次轉出", "最後轉出", "最後格式", "轉出格式"]
DEFAULT_WORKSHEET = "轉出紀錄"

_FIELD_BY_HEADER = {
    "單號": "key",
    "轉出次數": "count",
    "首次轉出": "first_at",
    "最後轉出": "last_at",
    "最後格式": "last_format",
    "轉出格式": "formats",
}


def rows_to_log(rows: list[list[object]]) -> dict:
    """試算表內容 → 紀錄 dict。無法辨識的列直接略過（手動編輯壞掉不該擋住轉換）。"""
    log = export_log.empty_log()
    if not rows:
        return log
    header = [str(cell).strip() for cell in rows[0]]
    if any(cell in _FIELD_BY_HEADER for cell in header):
        fields = [_FIELD_BY_HEADER.get(cell) for cell in header]
        data_rows = rows[1:]
    else:  # 沒有標題列：假設就是預設欄序
        fields = [_FIELD_BY_HEADER[name] for name in SHEET_HEADER]
        data_rows = rows
    for row in data_rows:
        values: dict[str, str] = {}
        for index, field in enumerate(fields):
            if field and index < len(row):
                values[field] = str(row[index]).strip()
        key = values.get("key", "")
        if not key:
            continue
        try:
            count = int(float(values.get("count") or 0))
        except ValueError:
            count = 0
        formats = [f.strip() for f in (values.get("formats") or "").split("／") if f.strip()]
        last_format = values.get("last_format") or ""
        if last_format and last_format not in formats:
            formats.append(last_format)
        log["records"][key] = {
            "count": max(count, 1),
            "first_at": values.get("first_at") or "",
            "last_at": values.get("last_at") or "",
            "last_format": last_format,
            "formats": formats,
        }
    return log


def log_to_rows(log: dict) -> list[list[str]]:
    """紀錄 dict → 試算表內容（含標題列），最後轉出時間新的排前面。"""
    records = sorted(
        (log.get("records") or {}).items(),
        key=lambda kv: str(kv[1].get("last_at") or ""),
        reverse=True,
    )
    rows = [list(SHEET_HEADER)]
    for key, record in records:
        rows.append([
            key,
            str(record.get("count") or 0),
            str(record.get("first_at") or ""),
            str(record.get("last_at") or ""),
            str(record.get("last_format") or ""),
            "／".join(str(f) for f in (record.get("formats") or [])),
        ])
    return rows


class LocalStore:
    """本機 JSON 檔。"""

    kind = "local"

    def __init__(self, path: str | Path):
        self.path = Path(path)

    @property
    def name(self) -> str:
        return f"local:{self.path}"

    @property
    def label(self) -> str:
        return f"本機檔案 `{self.path}`"

    def load(self) -> dict:
        return export_log.load(self.path)

    def save(self, log: dict) -> None:
        export_log.save(self.path, log)

    def record(self, keys, output_format: str, when: datetime | None = None) -> dict:
        return export_log.record_export(self.path, keys, output_format, when)

    def clear(self) -> dict:
        return export_log.clear(self.path)

    def is_empty(self) -> bool:
        return not self.load().get("records")


class SheetStore:
    """Google 試算表。"""

    kind = "sheet"

    def __init__(self, client: SheetClient, worksheet: str = DEFAULT_WORKSHEET):
        self.client = client
        self.worksheet = worksheet

    @property
    def name(self) -> str:
        return f"sheet:{self.client.spreadsheet_id}/{self.worksheet}"

    @property
    def label(self) -> str:
        return (
            f"Google 試算表 [{self.client.spreadsheet_id}]"
            f"(https://docs.google.com/spreadsheets/d/{self.client.spreadsheet_id}/edit)"
            f" 的「{self.worksheet}」分頁"
        )

    def load(self) -> dict:
        return rows_to_log(self.client.read(self.worksheet))

    def save(self, log: dict) -> None:
        self.client.write(self.worksheet, log_to_rows(log))

    def record(self, keys, output_format: str, when: datetime | None = None) -> dict:
        # 寫入前重讀一次：別人（或另一個分頁）剛寫進去的紀錄才不會被蓋掉。
        log = export_log.mark(self.load(), keys, output_format, when)
        self.save(log)
        return log

    def clear(self) -> dict:
        log = export_log.empty_log()
        self.save(log)
        return log


def _to_plain(value):
    """st.secrets 的物件轉成純 dict/list，才能餵給 google-auth。"""
    if hasattr(value, "items"):
        return {str(k): _to_plain(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_to_plain(v) for v in value]
    return value


def build_sheet_store(config) -> SheetStore:
    """依 secrets 的 [gsheet_log] 設定建立 SheetStore；設定不全時丟 SheetError。"""
    cfg = _to_plain(config) or {}
    spreadsheet = cfg.get("spreadsheet_id") or cfg.get("spreadsheet_url")
    if not spreadsheet:
        raise SheetError("Google 試算表設定缺少 spreadsheet_id（或 spreadsheet_url）。")
    account = cfg.get("service_account")
    if not isinstance(account, dict) or not account:
        raise SheetError(
            "Google 試算表設定缺少 [gsheet_log.service_account] 區塊"
            "（請把服務帳戶金鑰 JSON 的內容貼進去）。"
        )
    missing = [k for k in ("client_email", "private_key", "token_uri") if not account.get(k)]
    if missing:
        raise SheetError(f"服務帳戶金鑰缺少欄位：{'、'.join(missing)}")
    worksheet = str(cfg.get("worksheet") or DEFAULT_WORKSHEET)
    return SheetStore(SheetClient(account, str(spreadsheet)), worksheet)
