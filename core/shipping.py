# -*- coding: utf-8 -*-
"""HCT 出貨單轉換引擎（訂單 360 格式 / 調撥單 162 格式共用）。

由 VBA modHCTConverter360 / modHCTConverter162 移植，轉換規則完全相同：
  1. 依「預計到貨日 + 出貨客戶 + 門市/倉儲 + 地址」判斷可否跨單合併；
     同收件條件有 >=2 種訂單類型且含「一般銷售訂單」時合併為一張送貨單。
  2. 同送貨單內相同「料號 + 效期」的明細數量加總。
  3. 輸出 36 欄 HCT 銷貨報表（欄位常數取自 HCT範本）。
"""
from __future__ import annotations

import io
import re
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from .xlio import first_sheet, normalize_text, read_workbook, try_parse_date

GENERAL_ORDER_TYPE = "一般銷售訂單"
PLATFORM_NAME = "DR.WU 達爾膚生醫科技股份有限公司"
PLATFORM_URL = "https://www.drwu.com"
PLATFORM_SERVICE_PHONE = "0800-083-999"
CARRIER_NAME = "OMS"
OUTPUT_COLUMN_COUNT = 36

MODE_ORDER = "order"       # 銷售訂單未出貨明細 (360)
MODE_TRANSFER = "transfer"  # 調撥單未出貨明細 (162)

# 每個標準欄名可對應多個候選別名，依序取第一個存在的欄
#（2026-07 NetSuite 報表改版：項目名稱→顯示名稱、備忘錄→備忘錄 (主要)）。
_ALIASES = {
    MODE_ORDER: {
        "項目名稱": ("顯示名稱",),
        "序號/批號": ("交易序號/批號",),
        "DR_溫馨提醒": ("備忘錄", "備忘錄 (主要)"),
    },
    MODE_TRANSFER: {
        "項目名稱": ("顯示名稱",),
        "序號/批號": ("交易序號/批號",),
        "DR_溫馨提醒": ("備忘錄", "備忘錄 (主要)"),
        "出貨客戶": ("目標地點",),
        "門市/倉儲": ("倉儲",),
        "門市/倉儲聯繫人": ("倉儲聯繫人",),
        "門市/倉儲電話": ("倉儲電話",),
        "門市/倉儲地址": ("倉儲地址",),
    },
}

_REQUIRED_HEADERS = {
    MODE_ORDER: [
        "內部 ID", "文件編號",
        "銷售訂單類型", "日期", "項目", "項目名稱", "DR_料號",
        "序號/批號", "出貨客戶",
        "出貨數量", "門市/倉儲", "門市/倉儲聯繫人",
        "門市/倉儲電話", "門市/倉儲地址",
        "DR_預計到貨日期", "DR_預計到貨時段",
    ],
    MODE_TRANSFER: [
        "內部 ID", "文件編號", "日期", "項目",
        "出貨數量", "出貨客戶", "門市/倉儲", "門市/倉儲聯繫人",
        "門市/倉儲電話", "門市/倉儲地址",
    ],
}


@dataclass
class Shipment:
    merged: bool
    arrival_value: object          # date / str / None
    arrival_key: str
    customer: str
    location: str
    contact: str
    phone: str
    address: str
    input_rows: int = 0
    total_quantity: float = 0.0
    items: dict = field(default_factory=dict)          # key -> item dict（保序）
    documents: list = field(default_factory=list)
    sales_orders: list = field(default_factory=list)
    general_sales_orders: list = field(default_factory=list)
    types: list = field(default_factory=list)
    times: list = field(default_factory=list)
    purchase_orders: list = field(default_factory=list)
    reminders: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


@dataclass
class ConvertResult:
    output_bytes: bytes
    output_name: str
    shipments: int
    output_items: int
    problem_count: int
    input_rows: int
    skipped_rows: int
    merged_groups: int
    warnings: list
    problem_rows: list


class ConvertError(ValueError):
    pass


def convert(data: bytes, filename: str, mode: str, template_path: Path) -> ConvertResult:
    rows = first_sheet(read_workbook(data, filename))
    if not rows:
        raise ConvertError("來源工作表沒有可轉換的表格資料。")

    headers = _build_header_map(rows[0])
    _apply_aliases(headers, mode)
    missing = [h for h in _REQUIRED_HEADERS[mode] if h.casefold() not in headers]
    if missing:
        raise ConvertError("來源檔缺少必要欄位：\n" + "\n".join(f"- {h}" for h in missing))

    state = _build_shipments(rows, headers, mode)
    total_items = sum(len(s.items) for s in state["shipments"].values())
    if total_items == 0:
        raise ConvertError("沒有可輸出的有效品項。")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    mode_label = "訂單" if mode == MODE_ORDER else "調撥單"
    output_name = f"HCT銷貨報表_{mode_label}_{stamp}.xlsx"
    try:
        output_bytes = _write_output(state, mode, template_path)
    except Exception as exc:
        raise ConvertError(f"建立輸出檔時發生錯誤：{exc}") from exc

    return ConvertResult(
        output_bytes=output_bytes,
        output_name=output_name,
        shipments=len(state["shipments"]),
        output_items=total_items,
        problem_count=len(state["problem_rows"]),
        input_rows=state["input_rows"],
        skipped_rows=state["skipped_rows"],
        merged_groups=state["merged_groups"],
        warnings=state["warnings"],
        problem_rows=state["problem_rows"],
    )


# ------------------------------------------------------------------ 讀表輔助


def _build_header_map(header_row: list) -> dict[str, int]:
    headers: dict[str, int] = {}
    for index, value in enumerate(header_row):
        text = normalize_text(value)
        if text and text.casefold() not in headers:
            headers[text.casefold()] = index
    return headers


def _apply_aliases(headers: dict[str, int], mode: str) -> None:
    for canonical, aliases in _ALIASES[mode].items():
        key = canonical.casefold()
        if key in headers:
            continue
        for alias in aliases:
            alias_key = alias.casefold()
            if alias_key in headers:
                headers[key] = headers[alias_key]
                break


def _cell(row: list, headers: dict[str, int], name: str) -> object:
    index = headers.get(name.casefold())
    if index is None or index >= len(row):
        return None
    return row[index]


def _add_distinct(values: list, text: str) -> None:
    if text and text not in values:
        values.append(text)


def _normalize_date_key(value: object) -> str:
    parsed = try_parse_date(value)
    return parsed.strftime("%Y-%m-%d") if parsed else ""


def _normalize_arrival(value: object) -> tuple[object, str]:
    raw = normalize_text(value)
    if not raw:
        return None, "預計到貨日期空白；此 SO 不跨單合併。"
    parsed = try_parse_date(value)
    if parsed:
        return parsed, ""
    return raw, f"預計到貨日期無法辨識：{raw}"


def _normalize_expiry(batch_text: str) -> tuple[str, object, str]:
    """回傳 (expiry_key, expiry_value, warning)。批號本身即效期字串。"""
    if not batch_text:
        return "", None, ""
    parsed = try_parse_date(batch_text)
    if parsed:
        return parsed.strftime("%Y-%m-%d"), parsed, ""
    return "", None, f"效期無法辨識，已留白：{batch_text}"


# 「項目」欄的「料號直接連品名」樣態（2026-07 備品列新增）：開頭一串數字
# 料號後面直接接品名文字，如 902224100008ㄈ型試用品壓克力架(圓管)。
_ITEM_CODE_NAME_RE = re.compile(r"^(\d{6,})\s*(\D.*)$")


def _material_and_name(dr_material: object, item_text: object) -> tuple[str, str]:
    """從 DR_料號／項目 解析 (料號, 品名備援)。

    DR_料號空白時退回「項目」欄，該欄有三種樣態：
    料號_序號（取底線前段）、純料號、料號直接連品名（拆開後品名當備援）。
    """
    material = normalize_text(dr_material)
    if material:
        return material, ""
    fallback = normalize_text(item_text)
    matched = _ITEM_CODE_NAME_RE.match(fallback)
    if "_" in fallback:
        head = fallback.split("_", 1)[0]
        if head:
            fallback = head
    elif matched:
        return matched.group(1), matched.group(2).strip()
    return fallback.strip(), ""


BATCH_STATUS_OK = "批號賦予成功"


def _batch_status_warning(row: list, headers: dict[str, int], row_number: int) -> str:
    """調撥單報表的 DR_程式執行狀態非「批號賦予成功」時給警告（欄位不存在則略過）。"""
    if "DR_程式執行狀態".casefold() not in headers:
        return ""
    status = normalize_text(_cell(row, headers, "DR_程式執行狀態"))
    if status and status != BATCH_STATUS_OK:
        return f"第 {row_number} 列批號狀態為「{status}」，批號/效期可能尚未確認。"
    return ""


def _positive_quantity(value: object) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    try:
        quantity = float(str(value).replace(",", "").strip())
    except ValueError:
        return None
    return quantity if quantity > 0 else None


def _phone_to_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    return normalize_text(value)


def _split_phone(raw_phone: str) -> tuple[str, str]:
    compact = raw_phone.strip()
    for ch in ("-", " ", "(", ")", "\xa0"):
        compact = compact.replace(ch, "")
    digits = compact.lstrip("+")
    if digits.startswith("886"):
        digits = "0" + digits[3:]
    if digits and not digits.startswith("0"):
        digits = "0" + digits
    if not digits:
        return "", ""
    if digits.startswith("09"):
        return "", digits
    return digits, ""


def _split_postal(raw_address: str) -> tuple[str, str]:
    clean = raw_address.strip()
    if len(clean) >= 5 and clean[:5].isdigit():
        return clean[:5], clean[5:].strip()
    if len(clean) >= 3 and clean[:3].isdigit():
        return clean[:3], clean[3:].strip()
    return "", clean


# ------------------------------------------------------------------ 主轉換


def _build_shipments(rows: list, headers: dict[str, int], mode: str) -> dict:
    warnings: list[str] = []
    problem_rows: list[dict] = []
    input_rows = len(rows) - 1
    skipped_rows = 0

    candidate_types: dict[tuple, list[str]] = {}
    candidate_has_general: set = set()

    def candidate_key_for(row: list) -> tuple:
        return (
            _normalize_date_key(_cell(row, headers, "DR_預計到貨日期")),
            normalize_text(_cell(row, headers, "出貨客戶")),
            normalize_text(_cell(row, headers, "門市/倉儲")),
            normalize_text(_cell(row, headers, "門市/倉儲地址")),
        )

    for row in rows[1:]:
        key = candidate_key_for(row)
        if not all(key):
            continue
        order_type = normalize_text(_cell(row, headers, "銷售訂單類型"))
        types = candidate_types.setdefault(key, [])
        _add_distinct(types, order_type)
        if order_type.casefold() == GENERAL_ORDER_TYPE.casefold():
            candidate_has_general.add(key)

    merged_groups = sum(
        1
        for key, types in candidate_types.items()
        if len(types) >= 2 and key in candidate_has_general
    )

    shipments: dict[tuple, Shipment] = {}

    for row_number, row in enumerate(rows[1:], start=2):
        key = candidate_key_for(row)
        mergeable = all(key)

        document = normalize_text(_cell(row, headers, "文件編號"))
        if not document:
            document = "ID-" + normalize_text(_cell(row, headers, "內部 ID"))
            _add_distinct(warnings, f"第 {row_number} 列文件編號空白，改用 {document}")
        sales_order = normalize_text(_cell(row, headers, "銷售訂單單號"))
        if not sales_order:
            sales_order = document
            if "銷售訂單單號".casefold() in headers:
                _add_distinct(warnings, f"第 {row_number} 列銷售訂單單號空白，改用 {sales_order}")

        is_merged = (
            mergeable
            and key in candidate_types
            and len(candidate_types[key]) >= 2
            and key in candidate_has_general
        )
        shipment_key = ("M", *key) if is_merged else ("S", sales_order, *key)

        shipment = shipments.get(shipment_key)
        if shipment is None:
            arrival_value, arrival_warning = _normalize_arrival(
                _cell(row, headers, "DR_預計到貨日期")
            )
            shipment = Shipment(
                merged=is_merged,
                arrival_value=arrival_value,
                arrival_key=key[0],
                customer=key[1],
                location=key[2],
                contact=normalize_text(_cell(row, headers, "門市/倉儲聯繫人")),
                phone=_phone_to_text(_cell(row, headers, "門市/倉儲電話")),
                address=key[3],
            )
            if arrival_warning:
                _add_distinct(shipment.warnings, arrival_warning)
                _add_distinct(warnings, arrival_warning)
            if mode == MODE_ORDER and not mergeable:
                text = "合併鍵不完整；此 SO 不跨單合併。"
                _add_distinct(shipment.warnings, text)
                _add_distinct(warnings, text)
            shipments[shipment_key] = shipment

        skipped_rows += _add_row_to_shipment(
            shipment, row, row_number, headers, document, sales_order,
            warnings, problem_rows, mode,
        )

    return {
        "shipments": shipments,
        "warnings": warnings,
        "problem_rows": problem_rows,
        "input_rows": input_rows,
        "skipped_rows": skipped_rows,
        "merged_groups": merged_groups,
    }


def _add_row_to_shipment(
    shipment: Shipment,
    row: list,
    row_number: int,
    headers: dict[str, int],
    document: str,
    sales_order: str,
    warnings: list,
    problem_rows: list,
    mode: str,
) -> int:
    """把一列加進送貨單；回傳跳過列數 (0 或 1)。"""
    shipment.input_rows += 1
    _add_distinct(shipment.documents, document)
    _add_distinct(shipment.sales_orders, sales_order)

    order_type = normalize_text(_cell(row, headers, "銷售訂單類型"))
    _add_distinct(shipment.types, order_type)
    if order_type.casefold() == GENERAL_ORDER_TYPE.casefold():
        _add_distinct(shipment.general_sales_orders, sales_order)

    _add_distinct(shipment.times, normalize_text(_cell(row, headers, "DR_預計到貨時段")))
    _add_distinct(shipment.purchase_orders, normalize_text(_cell(row, headers, "客戶採購單編號")))
    _add_distinct(shipment.reminders, normalize_text(_cell(row, headers, "DR_溫馨提醒")))

    batch_value = normalize_text(_cell(row, headers, "序號/批號"))
    item_text = normalize_text(_cell(row, headers, "項目"))
    material, name_fallback = _material_and_name(_cell(row, headers, "DR_料號"), item_text)
    product_name = normalize_text(_cell(row, headers, "項目名稱")) or name_fallback

    status_warning = _batch_status_warning(row, headers, row_number)
    if status_warning:
        _add_distinct(warnings, status_warning)
        _add_distinct(shipment.warnings, status_warning)

    def record_problem(reason: str) -> None:
        problem_rows.append({
            "來源列": row_number,
            "文件編號": document,
            "銷售訂單單號": sales_order,
            "銷售訂單類型": order_type,
            "項目": item_text,
            "DR_料號": normalize_text(_cell(row, headers, "DR_料號")),
            "項目名稱": product_name,
            "序號/批號": batch_value,
            "出貨數量": normalize_text(_cell(row, headers, "出貨數量")),
            "出貨客戶": shipment.customer,
            "門市/倉儲": shipment.location,
            "門市/倉儲地址": shipment.address,
            "DR_預計到貨日期": normalize_text(_cell(row, headers, "DR_預計到貨日期")),
            "排除原因": reason,
        })

    if not material:
        text = f"第 {row_number} 列無有效料號，已跳過。"
        _add_distinct(warnings, text)
        _add_distinct(shipment.warnings, text)
        record_problem("無有效料號")
        return 1

    quantity = _positive_quantity(_cell(row, headers, "出貨數量"))
    if quantity is None:
        text = f"第 {row_number} 列數量無效，已跳過。"
        _add_distinct(warnings, text)
        _add_distinct(shipment.warnings, text)
        record_problem("數量無效")
        return 1

    expiry_key, expiry_value, expiry_warning = _normalize_expiry(batch_value)
    if expiry_warning:
        text = f"第 {row_number} 列：{expiry_warning}"
        _add_distinct(warnings, text)
        _add_distinct(shipment.warnings, text)

    item_key = (material, expiry_key)
    item = shipment.items.get(item_key)
    if item is not None:
        item["quantity"] += quantity
    else:
        shipment.items[item_key] = {
            "material": material,
            "product_name": product_name,
            "batch": batch_value,
            "expiry_value": expiry_value,
            "expiry_key": expiry_key,
            "quantity": quantity,
        }
    shipment.total_quantity += quantity
    return 0


# ------------------------------------------------------------------ 輸出


def _build_note(shipment: Shipment) -> str:
    return "|".join((
        "、".join(shipment.times),
        "、".join(shipment.purchase_orders),
        "、".join(shipment.reminders),
    ))


def _delivery_number(shipment: Shipment) -> str:
    return "、".join(shipment.sales_orders)


def _order_number(shipment: Shipment, mode: str) -> str:
    if shipment.merged:
        if mode == MODE_ORDER:
            return shipment.general_sales_orders[0] if shipment.general_sales_orders else ""
        return "、".join(shipment.general_sales_orders)
    return "、".join(shipment.sales_orders)


_TEXT_COLUMNS = (3, 5, 7, 11, 13, 15)
_DATE_COLUMNS = (29, 33)
_EXPIRY_COLUMN = 33  # 效期欄顯示為 yyyy/m/d（月日不補零），其他日期欄維持 yyyy-mm-dd
_TEMPLATE_VALUE_COLUMNS = (4, 20, 21, 22, 23, 25, 26, 27, 30, 31, 32, 34)


def _write_output(state: dict, mode: str, template_path: Path) -> bytes:
    import openpyxl

    template_wb = openpyxl.load_workbook(template_path)
    template_ws = template_wb.active

    template_values = {
        col: template_ws.cell(2, col).value for col in _TEMPLATE_VALUE_COLUMNS
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工作表1"

    for col in range(1, OUTPUT_COLUMN_COUNT + 1):
        header_cell = template_ws.cell(1, col)
        cell = ws.cell(1, col, header_cell.value)
        cell.font = copy(header_cell.font)
        cell.fill = copy(header_cell.fill)
        cell.border = copy(header_cell.border)
        cell.alignment = copy(header_cell.alignment)
        letter = openpyxl.utils.get_column_letter(col)
        width = template_ws.column_dimensions[letter].width
        if width:
            ws.column_dimensions[letter].width = width

    row_index = 1
    for shipment in state["shipments"].values():
        postal, address_text = _split_postal(shipment.address)
        daytime, mobile = _split_phone(shipment.phone)
        note = _build_note(shipment)
        for item in shipment.items.values():
            row_index += 1
            values: list[object] = [None] * (OUTPUT_COLUMN_COUNT + 1)
            values[1] = _delivery_number(shipment)
            values[2] = _order_number(shipment, mode)
            values[3] = item["material"]
            values[4] = template_values[4]
            values[5] = item["material"]
            values[6] = item["product_name"]
            values[7] = item["batch"]
            values[8] = item["quantity"]
            values[9] = ""
            values[10] = shipment.contact
            values[11] = postal
            values[12] = address_text
            values[13] = daytime
            values[14] = ""
            values[15] = mobile
            values[16] = PLATFORM_NAME
            values[17] = PLATFORM_URL
            values[18] = PLATFORM_SERVICE_PHONE
            values[19] = note
            for col in (20, 21, 22, 23, 25, 26, 27, 30, 31, 32, 34):
                values[col] = template_values[col]
            values[24] = CARRIER_NAME
            values[28] = ""
            values[29] = shipment.arrival_value
            values[33] = item["expiry_value"]
            values[35] = shipment.customer
            values[36] = 1
            for col in range(1, OUTPUT_COLUMN_COUNT + 1):
                cell = ws.cell(row_index, col)
                value = values[col]
                date_format = "yyyy/m/d" if col == _EXPIRY_COLUMN else "yyyy-mm-dd"
                if isinstance(value, date):
                    cell.value = value
                    cell.number_format = date_format
                else:
                    if col in _TEXT_COLUMNS:
                        cell.number_format = "@"
                        cell.value = "" if value is None else str(value)
                    else:
                        cell.value = value
                if col in _DATE_COLUMNS and not isinstance(value, date):
                    cell.number_format = date_format

    ws.auto_filter.ref = f"A1:AJ{row_index}"
    ws.freeze_panes = "A2"

    _write_problem_sheet(wb, state["problem_rows"])

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


_PROBLEM_HEADERS = [
    "來源列", "文件編號", "銷售訂單單號", "銷售訂單類型",
    "項目", "DR_料號", "項目名稱", "序號/批號", "出貨數量", "出貨客戶",
    "門市/倉儲", "門市/倉儲地址", "DR_預計到貨日期", "排除原因",
]


def _write_problem_sheet(wb, problem_rows: list) -> None:
    from openpyxl.styles import Font

    ws = wb.create_sheet("有問題訂單")
    for col, header in enumerate(_PROBLEM_HEADERS, start=1):
        cell = ws.cell(1, col, header)
        cell.font = Font(bold=True)
    ws.column_dimensions["H"].width = 16
    for col_cells in ws.iter_cols(min_col=8, max_col=8, min_row=2):
        for cell in col_cells:
            cell.number_format = "@"
    if problem_rows:
        for row_offset, record in enumerate(problem_rows, start=2):
            for col, header in enumerate(_PROBLEM_HEADERS, start=1):
                cell = ws.cell(row_offset, col, record.get(header, ""))
                if header == "序號/批號":
                    cell.number_format = "@"
    else:
        ws.cell(2, 1, "（本次沒有無法轉換的明細）")
