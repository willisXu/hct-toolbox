# -*- coding: utf-8 -*-
"""HCT 出貨資料 → M00 出貨格式（第三方物流匯入範本）。

沿用 shipping.py 的來源解析規則（必要欄位、料號/數量/效期正規化），
但「不做跨單合併」：每張原始銷售訂單/調撥單各自成一列群組，
只有同一張單內相同「料號＋效期」的數量才會加總。

固定值（依需求指定）：
  商品狀態 = 良品
  出貨方式 = 件出
  物流方式 = 黑貓宅急便(常溫)

「有效期限」「希望配送日」皆輸出為文字字串 "YYYY-MM-DD"，
對應 M00 範本裡的文字格式(@)，跟 HCT 範本輸出真正日期值不同。
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from .shipping import (
    MODE_ORDER,
    MODE_TRANSFER,
    _REQUIRED_HEADERS,
    _add_distinct,
    _apply_aliases,
    _batch_status_warning,
    _build_header_map,
    _cell,
    _material_and_name,
    _normalize_arrival,
    _normalize_expiry,
    _phone_to_text,
    _positive_quantity,
    _split_postal,
)
from .xlio import first_sheet, is_blank_row, normalize_text, read_workbook

FIXED_PRODUCT_STATUS = "良品"
FIXED_SHIP_METHOD = "件出"
FIXED_CARRIER = "黑貓宅急便(常溫)"

# DR_預計到貨時段（早/中/晚）→ M00 希望配送時段（僅接受上午/下午/不指定）。
# 沒有值，或值無法對應到下列任一項時，預設填「上午」。
_TIME_SLOT_MAP = {
    "早(9-12)": "上午",
    "中(12-17)": "下午",
    "晚(17-20)": "不指定",
}
_DEFAULT_TIME_SLOT = "上午"

_SHEET_NAME = "訂單"
# 客製訂單編號/SKU/有效期限/聯絡號碼/希望配送日/郵遞區號：強制文字格式，避免被當數字處理掉前導 0。
_TEXT_COLUMNS = (1, 2, 3, 10, 11, 15)


class M00Error(ValueError):
    pass


@dataclass
class M00Group:
    order_number: str
    arrival_value: object          # date / str / None
    contact: str
    phone: str
    address: str
    items: dict = field(default_factory=dict)      # (material, expiry_key) -> item dict（保序）
    times: list = field(default_factory=list)
    reminders: list = field(default_factory=list)


@dataclass
class M00Result:
    output_bytes: bytes
    output_name: str
    order_count: int
    output_items: int
    problem_count: int
    input_rows: int
    skipped_rows: int
    warnings: list
    problem_rows: list


def convert(data: bytes, filename: str, mode: str, template_path: Path) -> M00Result:
    rows = first_sheet(read_workbook(data, filename))
    if not rows:
        raise M00Error("來源工作表沒有可轉換的表格資料。")

    headers = _build_header_map(rows[0])
    _apply_aliases(headers, mode)
    missing = [h for h in _REQUIRED_HEADERS[mode] if h.casefold() not in headers]
    if missing:
        raise M00Error("來源檔缺少必要欄位：\n" + "\n".join(f"- {h}" for h in missing))

    state = _build_groups(rows, headers, mode)
    total_items = sum(len(g.items) for g in state["groups"].values())
    if total_items == 0:
        raise M00Error("沒有可輸出的有效品項。")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    mode_label = "訂單" if mode == MODE_ORDER else "調撥單"
    output_name = f"M00出貨格式_{mode_label}_{stamp}.xlsx"
    try:
        output_bytes = _write_output(state, template_path)
    except Exception as exc:
        raise M00Error(f"建立輸出檔時發生錯誤：{exc}") from exc

    return M00Result(
        output_bytes=output_bytes,
        output_name=output_name,
        order_count=len(state["groups"]),
        output_items=total_items,
        problem_count=len(state["problem_rows"]),
        input_rows=state["input_rows"],
        skipped_rows=state["skipped_rows"],
        warnings=state["warnings"],
        problem_rows=state["problem_rows"],
    )


# ------------------------------------------------------------------ 主轉換（不合併）


def _build_groups(rows: list, headers: dict, mode: str) -> dict:
    warnings: list[str] = []
    problem_rows: list[dict] = []
    input_rows = sum(1 for row in rows[1:] if not is_blank_row(row))
    skipped_rows = 0
    groups: dict[str, M00Group] = {}

    for row_number, row in enumerate(rows[1:], start=2):
        if is_blank_row(row):
            continue
        document = normalize_text(_cell(row, headers, "文件編號"))
        if not document:
            document = "ID-" + normalize_text(_cell(row, headers, "內部 ID"))
            _add_distinct(warnings, f"第 {row_number} 列文件編號空白，改用 {document}")
        sales_order = normalize_text(_cell(row, headers, "銷售訂單單號"))
        if not sales_order:
            sales_order = document
            if "銷售訂單單號".casefold() in headers:
                _add_distinct(warnings, f"第 {row_number} 列銷售訂單單號空白，改用 {sales_order}")

        group = groups.get(sales_order)
        if group is None:
            arrival_value, arrival_warning = _normalize_arrival(
                _cell(row, headers, "DR_預計到貨日期")
            )
            group = M00Group(
                order_number=sales_order,
                arrival_value=arrival_value,
                contact=normalize_text(_cell(row, headers, "門市/倉儲聯繫人")),
                phone=_phone_to_text(_cell(row, headers, "門市/倉儲電話")),
                address=normalize_text(_cell(row, headers, "門市/倉儲地址")),
            )
            if arrival_warning:
                _add_distinct(warnings, arrival_warning)
            groups[sales_order] = group

        _add_distinct(group.times, normalize_text(_cell(row, headers, "DR_預計到貨時段")))
        _add_distinct(group.reminders, normalize_text(_cell(row, headers, "DR_溫馨提醒")))

        batch_value = normalize_text(_cell(row, headers, "序號/批號"))
        item_text = normalize_text(_cell(row, headers, "項目"))
        material, name_fallback = _material_and_name(_cell(row, headers, "DR_料號"), item_text)
        product_name = normalize_text(_cell(row, headers, "項目名稱")) or name_fallback

        status_warning = _batch_status_warning(row, headers, row_number)
        if status_warning:
            _add_distinct(warnings, status_warning)

        def record_problem(reason: str) -> None:
            problem_rows.append({
                "來源列": row_number,
                "文件編號": document,
                "銷售訂單單號": sales_order,
                "項目": item_text,
                "DR_料號": normalize_text(_cell(row, headers, "DR_料號")),
                "項目名稱": product_name,
                "序號/批號": batch_value,
                "出貨數量": normalize_text(_cell(row, headers, "出貨數量")),
                "排除原因": reason,
            })

        if not material:
            text = f"第 {row_number} 列無有效料號，已跳過。"
            _add_distinct(warnings, text)
            record_problem("無有效料號")
            skipped_rows += 1
            continue

        quantity = _positive_quantity(_cell(row, headers, "出貨數量"))
        if quantity is None:
            text = f"第 {row_number} 列數量無效，已跳過。"
            _add_distinct(warnings, text)
            record_problem("數量無效")
            skipped_rows += 1
            continue

        expiry_key, expiry_value, expiry_warning = _normalize_expiry(batch_value)
        if expiry_warning:
            _add_distinct(warnings, f"第 {row_number} 列：{expiry_warning}")

        item_key = (material, expiry_key)
        item = group.items.get(item_key)
        if item is not None:
            item["quantity"] += quantity
        else:
            group.items[item_key] = {
                "material": material,
                "expiry_value": expiry_value,
                "quantity": quantity,
            }

    return {
        "groups": groups,
        "warnings": warnings,
        "problem_rows": problem_rows,
        "input_rows": input_rows,
        "skipped_rows": skipped_rows,
    }


# ------------------------------------------------------------------ 輸出


def _date_text(value: object) -> str:
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if value is None:
        return ""
    return normalize_text(value)


def _time_slot_text(raw_values: list) -> str:
    if not raw_values:
        return _DEFAULT_TIME_SLOT
    mapped = dict.fromkeys(_TIME_SLOT_MAP.get(v, _DEFAULT_TIME_SLOT) for v in raw_values)
    return "、".join(mapped)


def _write_output(state: dict, template_path: Path) -> bytes:
    import openpyxl

    wb = openpyxl.load_workbook(template_path)
    ws = wb[_SHEET_NAME]

    row_index = 1
    for group in state["groups"].values():
        postal, address_text = _split_postal(group.address)
        note = "、".join(group.reminders)
        times_text = _time_slot_text(group.times)
        arrival_text = _date_text(group.arrival_value)
        for item in group.items.values():
            row_index += 1
            expiry_text = _date_text(item["expiry_value"])
            values = [
                group.order_number,     # A 客製訂單編號
                item["material"],       # B SKU(商品貨號)
                expiry_text,             # C 有效期限
                FIXED_PRODUCT_STATUS,    # D 商品狀態
                FIXED_SHIP_METHOD,       # E 出貨方式
                item["quantity"],        # F 數量
                FIXED_CARRIER,           # G 物流方式
                group.contact,           # H 收件人
                address_text,            # I 收件地址
                group.phone,             # J 聯絡號碼
                arrival_text,            # K 希望配送日
                times_text,              # L 希望配送時段
                note,                    # M 物流備註
                None,                    # N 是否為B2B
                postal,                  # O 郵遞區號
            ]
            for col, value in enumerate(values, start=1):
                cell = ws.cell(row_index, col)
                cell.value = value
                if col in _TEXT_COLUMNS:
                    cell.number_format = "@"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
