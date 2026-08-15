# -*- coding: utf-8 -*-
"""表格篩選 — 「NetSuite 直接抓取」抓回來的資料列篩選邏輯。

抓回來的資料是 [header 列] + 資料列（跟上傳檔案讀出來的格式相同），
這裡提供不依賴 Streamlit 的純函式，方便測試：

  * ``distinct_values``  某欄的相異值清單（給下拉多選用）
  * ``is_date_column``   判斷某欄是不是日期欄（給日期區間篩選用）
  * ``filter_indices``   套用關鍵字／欄位值／欄位包含文字／日期區間，
                         回傳符合條件的「資料列索引」（0 起算，不含 header）

索引一律用欄位「索引」而非欄名，避免 saved search 出現同名欄時對錯欄。
"""
from __future__ import annotations

from datetime import date, datetime

from .xlio import normalize_text, try_parse_date

# 多選清單裡代表「這一格是空白」的選項文字（空字串沒辦法在 UI 上顯示）。
BLANK_LABEL = "（空白）"

# 相異值超過這個數量就不適合用下拉多選（改用「包含文字」輸入框）。
MAX_CHOICES = 300


def _cell(row: list, index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return normalize_text(row[index])


def distinct_values(rows: list[list[object]], col_index: int, limit: int | None = None) -> list[str]:
    """回傳某欄的相異值（已排序）；該欄有空白格時，清單開頭加上 ``BLANK_LABEL``。

    limit 只用來提早中斷掃描（相異值太多時 UI 不會用這份清單），
    回傳的筆數可能會超過 limit 一點點，呼叫端只需判斷「有沒有超過」。
    """
    seen: dict[str, None] = {}
    has_blank = False
    for row in rows[1:]:
        text = _cell(row, col_index)
        if not text:
            has_blank = True
            continue
        if text not in seen:
            seen[text] = None
            if limit is not None and len(seen) > limit:
                break
    values = sorted(seen)
    return ([BLANK_LABEL] + values) if has_blank else values


def _looks_like_date(value: object) -> bool:
    """比 try_parse_date 嚴格：避免把「出貨數量」這種小整數欄當成 Excel 日期序號。

    只認 datetime/date 型別、以及看得出來是日期寫法的文字（2026-08-14、
    2026/8/14、20260814）。純數字只接受 8 碼 YYYYMMDD。
    """
    if isinstance(value, (datetime, date)):
        return True
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return value == int(value) and 10000000 <= int(value) <= 99999999 and try_parse_date(value) is not None
    text = normalize_text(value)
    if not text:
        return False
    if "T" in text:
        text = text.split("T", 1)[0]
    if text.isdigit():
        return len(text) == 8 and try_parse_date(text) is not None
    return try_parse_date(text) is not None


def is_date_column(
    rows: list[list[object]], col_index: int, sample: int = 200, threshold: float = 0.8
) -> bool:
    """抽樣判斷某欄是不是日期欄（非空白格有 threshold 以上看起來是日期）。"""
    checked = 0
    hits = 0
    for row in rows[1:]:
        text = _cell(row, col_index)
        if not text:
            continue
        checked += 1
        if _looks_like_date(row[col_index] if col_index < len(row) else None):
            hits += 1
        if checked >= sample:
            break
    return checked > 0 and hits / checked >= threshold


def _row_text(row: list) -> str:
    return " ".join(normalize_text(value) for value in row).casefold()


def filter_indices(
    rows: list[list[object]],
    keyword: str = "",
    value_filters: dict[int, set[str]] | None = None,
    text_filters: dict[int, str] | None = None,
    date_ranges: dict[int, tuple[date | None, date | None]] | None = None,
) -> list[int]:
    """回傳符合所有條件的資料列索引（0 起算，對應 rows[1:]）。

    * keyword：跨所有欄位的關鍵字，以空白分隔多個詞，全部都要命中（AND）。
    * value_filters：{欄索引: {允許的值}}，值取 normalize 後的字串；
      集合裡含 ``BLANK_LABEL`` 時，空白格也算命中。
    * text_filters：{欄索引: 包含文字}（不分大小寫）。
    * date_ranges：{欄索引: (起, 迄)}，任一端可為 None 表示不限；
      該欄解析不出日期的列一律排除（篩日期時空白/亂碼不該混進來）。
    """
    value_filters = value_filters or {}
    text_filters = text_filters or {}
    date_ranges = date_ranges or {}
    terms = [t.casefold() for t in keyword.split()] if keyword else []

    matched: list[int] = []
    for offset, row in enumerate(rows[1:]):
        if terms:
            haystack = _row_text(row)
            if not all(term in haystack for term in terms):
                continue
        ok = True
        for col_index, allowed in value_filters.items():
            if not allowed:
                continue
            text = _cell(row, col_index)
            if text:
                if text not in allowed:
                    ok = False
                    break
            elif BLANK_LABEL not in allowed:
                ok = False
                break
        if not ok:
            continue
        for col_index, needle in text_filters.items():
            if not needle:
                continue
            if needle.casefold() not in _cell(row, col_index).casefold():
                ok = False
                break
        if not ok:
            continue
        for col_index, (start, end) in date_ranges.items():
            if start is None and end is None:
                continue
            parsed = try_parse_date(row[col_index]) if col_index < len(row) else None
            if parsed is None or (start and parsed < start) or (end and parsed > end):
                ok = False
                break
        if ok:
            matched.append(offset)
    return matched
