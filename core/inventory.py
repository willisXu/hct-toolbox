# -*- coding: utf-8 -*-
"""HCT × NetSuite 庫存核對（由 VBA modInventory* 模組移植）。

  - NetSuite 報表欄位：地點、到期日、DR_料號、項目計數 總和、數量 總和
  - HCT 報表欄位：儲區類別、有效日期、客戶產品編號、可出數量、庫存數量
  - 只核對 G00/G10/G30/G40/G80/G90 倉；其他倉別列入排除數。
  - 以「倉別+料號+到期日」做日期明細核對，另以「倉別+料號」做料號彙總核對。
  - 差異基準：HCT－NetSuite。
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date, datetime

from .xlio import read_workbook

SYSTEM_HCT = "HCT"
SYSTEM_NETSUITE = "NetSuite"
APPROVED_LOCATIONS = {"G00", "G10", "G30", "G40", "G80", "G90"}

STATUS_MATCH = "完全一致"
STATUS_AVAILABLE_DIFF = "僅可用量差異"
STATUS_TOTAL_DIFF = "僅總庫存差異"
STATUS_BOTH_DIFF = "兩者皆不同"
STATUS_HCT_ONLY = "僅 HCT 存在"
STATUS_NS_ONLY = "僅 NetSuite 存在"
ALL_STATUSES = [
    STATUS_MATCH, STATUS_AVAILABLE_DIFF, STATUS_TOTAL_DIFF,
    STATUS_BOTH_DIFF, STATUS_HCT_ONLY, STATUS_NS_ONLY,
]

_NS_HEADERS = ["地點", "到期日", "DR_料號", "項目計數 總和", "數量 總和"]
# 業務助理版 NetSuite 庫存報表（293 格式）：料號取「項目」底線前段，
# 只有一個數量欄「庫存數 總和」，同時當作可用量與總庫存量。
_NS293_HEADERS = ["地點", "到期日", "項目", "庫存數 總和"]
# 物流核對版 NetSuite 庫存報表（663 格式）：料號用 DR_料號（空白時取「項目」
# 底線前段），到期日在「庫存編號」欄，可用量=可用、總庫存量=在庫量。
_NS663_HEADERS = ["項目", "DR_料號", "地點", "庫存編號", "在庫量", "可用"]
_HCT_HEADERS = ["客戶產品編號", "有效日期", "儲區類別", "可出數量", "庫存數量"]

SYSTEM_NETSUITE_293 = "NetSuite293"
SYSTEM_NETSUITE_663 = "NetSuite663"


class InventoryError(ValueError):
    pass


@dataclass
class SourceStats:
    file_name: str
    sheet_name: str = ""
    rows_read: int = 0
    valid_rows: int = 0
    excluded_rows: int = 0
    anomaly_rows: int = 0


@dataclass
class InventoryResult:
    output_bytes: bytes
    output_name: str
    detail_rows: list
    item_rows: list
    anomalies: list
    hct_stats: SourceStats
    ns_stats: SourceStats
    detail_status_counts: dict = field(default_factory=dict)
    item_status_counts: dict = field(default_factory=dict)


# ------------------------------------------------------------------ 正規化


def _normalize_header(value: object) -> str:
    text = "" if value is None else str(value)
    return text.replace("\xa0", " ").replace("　", " ").strip()


def _normalize_item(value: object) -> str:
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return "" if value is None else str(value).strip()


def _normalize_location(value: object) -> str:
    text = ("" if value is None else str(value)).strip().upper()
    if "_" in text:
        text = text.split("_", 1)[0]
    return text.strip()


def _normalize_date(value: object) -> tuple[str, bool]:
    """回傳 (yyyymmdd 或空字串, 是否有效)。空值視為有效的「無到期日」。"""
    if value is None:
        return "", True
    if isinstance(value, datetime):
        return value.strftime("%Y%m%d"), True
    if isinstance(value, date):
        return value.strftime("%Y%m%d"), True
    text = str(value).strip()
    if not text:
        return "", True
    if "T" in text:
        text = text.split("T", 1)[0]
    try:
        if len(text) == 8 and text.isdigit():
            parsed = date(int(text[:4]), int(text[4:6]), int(text[6:]))
        elif "/" in text:
            parts = text.split("/")
            if len(parts) != 3:
                return "", False
            parsed = date(int(parts[0]), int(parts[1]), int(parts[2]))
        elif "-" in text:
            parts = text.split("-")
            if len(parts) != 3:
                return "", False
            parsed = date(int(parts[0]), int(parts[1]), int(parts[2]))
        elif isinstance(value, (int, float)) and 0 < float(value) < 100000:
            from datetime import timedelta
            parsed = (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        else:
            return "", False
    except (ValueError, OverflowError):
        return "", False
    return parsed.strftime("%Y%m%d"), True


def _normalize_quantity(value: object) -> tuple[float, bool]:
    if value is None:
        return 0.0, True
    if isinstance(value, bool):
        return 0.0, False
    if isinstance(value, (int, float)):
        return float(value), True
    text = str(value).strip()
    if not text:
        return 0.0, True
    text = text.replace(",", "")
    try:
        return float(text), True
    except ValueError:
        return 0.0, False


def _expiry_output(expiry_key: str) -> object:
    if not expiry_key:
        return "無到期日"
    return date(int(expiry_key[:4]), int(expiry_key[4:6]), int(expiry_key[6:]))


# ------------------------------------------------------------------ 匯入


def _find_matching_sheet(book: dict, required: list[str]):
    """在所有工作表前 10 列內找欄位列；回傳 (sheet_name, rows, header_row_idx, header_map)。"""
    for sheet_name, rows in book.items():
        limit = min(len(rows), 10)
        for row_index in range(limit):
            header_map: dict[str, int] = {}
            for col_index, value in enumerate(rows[row_index]):
                text = _normalize_header(value)
                if text and text.casefold() not in header_map:
                    header_map[text.casefold()] = col_index
            if all(h.casefold() in header_map for h in required):
                return sheet_name, rows, row_index, header_map
    return None


def detect_source(data: bytes, filename: str) -> str:
    book = read_workbook(data, filename)
    matches = []
    if _find_matching_sheet(book, _NS_HEADERS):
        matches.append(SYSTEM_NETSUITE)
    if _find_matching_sheet(book, _NS663_HEADERS):
        matches.append(SYSTEM_NETSUITE_663)
    if _find_matching_sheet(book, _NS293_HEADERS):
        matches.append(SYSTEM_NETSUITE_293)
    if _find_matching_sheet(book, _HCT_HEADERS):
        matches.append(SYSTEM_HCT)
    if len(matches) > 1:
        raise InventoryError(
            f"{filename}:同時符合多種報表格式（{'、'.join(matches)}），無法自動判斷，"
            "請確認欄位名稱是否重複或誤植。"
        )
    return matches[0] if matches else ""


def _import_source(data: bytes, filename: str, system: str,
                   detail: dict, item_totals: dict, anomalies: list,
                   stats: SourceStats) -> None:
    book = read_workbook(data, filename)
    if system == SYSTEM_HCT:
        required = _HCT_HEADERS
    elif system == SYSTEM_NETSUITE_293:
        required = _NS293_HEADERS
    elif system == SYSTEM_NETSUITE_663:
        required = _NS663_HEADERS
    else:
        required = _NS_HEADERS
    match = _find_matching_sheet(book, required)
    if match is None:
        raise InventoryError(
            f"{filename}:找不到包含必要欄位的工作表:" + "、".join(required))
    sheet_name, rows, header_row, header_map = match
    stats.sheet_name = sheet_name

    def col(name: str) -> int | None:
        return header_map.get(name.casefold())

    strip_item_suffix = False
    item_fallback_col = None
    if system == SYSTEM_HCT:
        location_col, item_col = col("儲區類別"), col("客戶產品編號")
        expiry_col, available_col, total_col = col("有效日期"), col("可出數量"), col("庫存數量")
        desc_col = col("產品名稱")
        field_names = ("儲區類別", "客戶產品編號", "有效日期", "可出數量", "庫存數量")
    elif system == SYSTEM_NETSUITE_293:
        location_col, item_col = col("地點"), col("項目")
        expiry_col = col("到期日")
        available_col = total_col = col("庫存數 總和")
        desc_col = col("項目")
        field_names = ("地點", "項目", "到期日", "庫存數 總和", "庫存數 總和")
        strip_item_suffix = True
    elif system == SYSTEM_NETSUITE_663:
        location_col, item_col = col("地點"), col("DR_料號")
        expiry_col, available_col, total_col = col("庫存編號"), col("可用"), col("在庫量")
        desc_col = col("項目")
        field_names = ("地點", "DR_料號", "庫存編號", "可用", "在庫量")
        item_fallback_col = col("項目")
    else:
        location_col, item_col = col("地點"), col("DR_料號")
        expiry_col, available_col, total_col = col("到期日"), col("項目計數 總和"), col("數量 總和")
        desc_col = col("項目")
        field_names = ("地點", "DR_料號", "到期日", "項目計數 總和", "數量 總和")

    def get(row: list, index: int | None) -> object:
        if index is None or index >= len(row):
            return None
        return row[index]

    for row_index in range(header_row + 1, len(rows)):
        row = rows[row_index]
        raw = [get(row, c) for c in (location_col, item_col, expiry_col, available_col, total_col)]
        if all((v is None or str(v).strip() == "") for v in raw):
            continue

        stats.rows_read += 1
        actual_row = row_index + 1
        has_issue = False

        def anomaly(field_name: str, raw_value: object, issue: str) -> None:
            display_system = SYSTEM_NETSUITE if system.startswith(SYSTEM_NETSUITE) else system
            anomalies.append({
                "來源系統": display_system, "來源檔名": filename, "工作表": sheet_name,
                "原始列號": actual_row, "問題欄位": field_name,
                "原始值": "" if raw_value is None else str(raw_value),
                "異常說明": issue,
            })

        location = _normalize_location(get(row, location_col))
        if not location:
            anomaly(field_names[0], get(row, location_col), "倉別不可空白")
            has_issue = True
        elif location not in APPROVED_LOCATIONS:
            stats.excluded_rows += 1
            continue

        item_code = _normalize_item(get(row, item_col))
        if strip_item_suffix and "_" in item_code:
            item_code = item_code.split("_", 1)[0].strip()
        if not item_code and item_fallback_col is not None:
            # 663 格式：DR_料號空白時，改取「項目」底線前段當料號。
            item_code = _normalize_item(get(row, item_fallback_col)).split("_", 1)[0].strip()
        if not item_code:
            anomaly(field_names[1], get(row, item_col), "料號不可空白")
            has_issue = True

        expiry_key, expiry_ok = _normalize_date(get(row, expiry_col))
        if not expiry_ok:
            anomaly(field_names[2], get(row, expiry_col), "非空日期無法辨識")
            has_issue = True

        available, available_ok = _normalize_quantity(get(row, available_col))
        if not available_ok:
            anomaly(field_names[3], get(row, available_col), "數量不是有效數字")
            has_issue = True

        total, total_ok = _normalize_quantity(get(row, total_col))
        if not total_ok:
            anomaly(field_names[4], get(row, total_col), "數量不是有效數字")
            has_issue = True

        if has_issue:
            stats.anomaly_rows += 1
            continue

        description = _normalize_item(get(row, desc_col)) if desc_col is not None else ""
        _add_aggregate(detail, (location, item_code, expiry_key), available, total, description)
        _add_aggregate(item_totals, (location, item_code), available, total, description)
        stats.valid_rows += 1


def _add_aggregate(aggregate: dict, key: tuple, available: float, total: float,
                   description: str) -> None:
    bucket = aggregate.get(key)
    if bucket is None:
        aggregate[key] = [available, total, 1, description]
    else:
        bucket[0] += available
        bucket[1] += total
        bucket[2] += 1
        if not bucket[3] and description:
            bucket[3] = description


# ------------------------------------------------------------------ 核對


def _classify(has_hct: bool, has_ns: bool, hct_avail: float, hct_total: float,
              ns_avail: float, ns_total: float) -> tuple[str, str]:
    if not has_ns:
        return STATUS_HCT_ONLY, "HCT 有庫存資料，NetSuite 未找到對應資料"
    if not has_hct:
        return STATUS_NS_ONLY, "NetSuite 有庫存資料，HCT 未找到對應資料"
    if hct_avail == ns_avail and hct_total == ns_total:
        return STATUS_MATCH, "HCT 與 NetSuite 可用量及總庫存量一致"
    if hct_total == ns_total:
        return STATUS_AVAILABLE_DIFF, "帳面總量一致，但可出／可用狀態不同"
    if hct_avail == ns_avail:
        return STATUS_TOTAL_DIFF, "可用量一致，但帳面總量不同"
    return STATUS_BOTH_DIFF, "可用量與總庫存量皆有差異"


def _build_reconciliation(hct_data: dict, ns_data: dict, include_expiry: bool) -> list[dict]:
    union_keys = sorted(set(hct_data) | set(ns_data), key=lambda k: tuple(str(p).casefold() for p in k))
    result = []
    empty = [0.0, 0.0, 0, ""]
    for key in union_keys:
        has_hct = key in hct_data
        has_ns = key in ns_data
        hct_bucket = hct_data.get(key, empty)
        ns_bucket = ns_data.get(key, empty)
        description = hct_bucket[3] or ns_bucket[3]
        status, explanation = _classify(
            has_hct, has_ns, hct_bucket[0], hct_bucket[1], ns_bucket[0], ns_bucket[1])
        record = {
            "倉別": key[0],
            "料號": key[1],
            "品名／項目": description,
            "HCT 可出數量": hct_bucket[0],
            "NetSuite 項目計數": ns_bucket[0],
            "可用量差額": hct_bucket[0] - ns_bucket[0],
            "HCT 庫存數量": hct_bucket[1],
            "NetSuite 數量": ns_bucket[1],
            "總庫存差額": hct_bucket[1] - ns_bucket[1],
            "狀態": status,
            "結果說明": explanation,
            "HCT 來源列數": hct_bucket[2],
            "NetSuite 來源列數": ns_bucket[2],
        }
        if include_expiry:
            record["到期日"] = _expiry_output(key[2])
        result.append(record)
    return result


def reconcile(ns_data: bytes, ns_name: str, hct_data: bytes, hct_name: str) -> InventoryResult:
    """核對兩份庫存報表；傳入順序不限，程式自動辨識並對調。"""
    if ns_data == hct_data:
        raise InventoryError("兩次選取的是同一個檔案，請重新選擇不同的來源檔案。")
    first_system = detect_source(ns_data, ns_name)
    if not first_system:
        raise InventoryError(f"無法辨識檔案:{ns_name}(找不到 NetSuite 或 HCT 的必要欄位)")
    second_system = detect_source(hct_data, hct_name)
    if not second_system:
        raise InventoryError(f"無法辨識檔案:{hct_name}(找不到 NetSuite 或 HCT 的必要欄位)")

    ns_variants = (SYSTEM_NETSUITE, SYSTEM_NETSUITE_293, SYSTEM_NETSUITE_663)
    if first_system == SYSTEM_HCT and second_system in ns_variants:
        ns_data, hct_data = hct_data, ns_data
        ns_name, hct_name = hct_name, ns_name
        ns_system = second_system
    elif first_system in ns_variants and second_system == SYSTEM_HCT:
        ns_system = first_system
    else:
        raise InventoryError("兩份檔案無法配成一份 NetSuite 與一份 HCT 報表，請重新選取。")

    hct_detail: dict = {}
    hct_items: dict = {}
    ns_detail: dict = {}
    ns_items: dict = {}
    anomalies: list = []
    hct_stats = SourceStats(file_name=hct_name)
    ns_stats = SourceStats(file_name=ns_name)

    _import_source(ns_data, ns_name, ns_system, ns_detail, ns_items, anomalies, ns_stats)
    _import_source(hct_data, hct_name, SYSTEM_HCT, hct_detail, hct_items, anomalies, hct_stats)

    detail_rows = _build_reconciliation(hct_detail, ns_detail, include_expiry=True)
    item_rows = _build_reconciliation(hct_items, ns_items, include_expiry=False)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        output_bytes = _write_output(detail_rows, item_rows, anomalies, hct_stats, ns_stats)
    except Exception as exc:
        raise InventoryError(f"建立輸出檔時發生錯誤：{exc}") from exc

    def counts(rows):
        counter: dict[str, int] = {}
        for record in rows:
            counter[record["狀態"]] = counter.get(record["狀態"], 0) + 1
        return counter

    return InventoryResult(
        output_bytes=output_bytes,
        output_name=f"庫存核對結果_{stamp}.xlsx",
        detail_rows=detail_rows,
        item_rows=item_rows,
        anomalies=anomalies,
        hct_stats=hct_stats,
        ns_stats=ns_stats,
        detail_status_counts=counts(detail_rows),
        item_status_counts=counts(item_rows),
    )


# ------------------------------------------------------------------ 輸出

_DETAIL_COLUMNS = [
    "倉別", "料號", "品名／項目", "到期日", "HCT 可出數量", "NetSuite 項目計數",
    "可用量差額", "HCT 庫存數量", "NetSuite 數量", "總庫存差額", "狀態", "結果說明",
    "HCT 來源列數", "NetSuite 來源列數",
]
_ITEM_COLUMNS = [
    "倉別", "料號", "品名／項目", "HCT 可出數量", "NetSuite 項目計數",
    "可用量差額", "HCT 庫存數量", "NetSuite 數量", "總庫存差額", "狀態", "結果說明",
    "HCT 來源列數", "NetSuite 來源列數",
]
_ANOMALY_COLUMNS = ["來源系統", "來源檔名", "工作表", "原始列號", "問題欄位", "原始值", "異常說明"]

_STATUS_FILL = {
    STATUS_MATCH: ("C6EFCE", "006100"),
    STATUS_AVAILABLE_DIFF: ("FFEB9C", "9C6500"),
    STATUS_TOTAL_DIFF: ("FFEB9C", "9C6500"),
    STATUS_BOTH_DIFF: ("FFC7CE", "9C0006"),
    STATUS_HCT_ONLY: ("FFC7CE", "9C0006"),
    STATUS_NS_ONLY: ("FFC7CE", "9C0006"),
}


def _write_output(detail_rows, item_rows, anomalies, hct_stats, ns_stats) -> bytes:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")

    wb = openpyxl.Workbook()

    summary = wb.active
    summary.title = "核對摘要"
    summary.merge_cells("A1:F1")
    summary["A1"] = "庫存核對摘要"
    summary["A1"].font = Font(bold=True, size=20, color="FFFFFF")
    summary["A1"].fill = PatternFill("solid", fgColor="1F4E79")
    summary.row_dimensions[1].height = 38

    summary["A3"] = "狀態"
    summary["B3"] = "日期明細"
    summary["C3"] = "料號彙總"

    def count_rows(rows, status):
        return sum(1 for r in rows if r["狀態"] == status)

    for offset, status in enumerate(ALL_STATUSES):
        summary.cell(4 + offset, 1, status)
        summary.cell(4 + offset, 2, count_rows(detail_rows, status))
        summary.cell(4 + offset, 3, count_rows(item_rows, status))
    summary["A10"] = "總筆數"
    summary["B10"] = len(detail_rows)
    summary["C10"] = len(item_rows)

    summary["A12"] = "核對範圍倉別"
    summary["B12"] = "、".join(sorted(APPROVED_LOCATIONS))
    summary["A13"] = "資料統計"
    summary["A14"] = "項目"
    summary["B14"] = "HCT"
    summary["C14"] = "NetSuite"
    stat_labels = [
        ("讀取筆數", "rows_read"), ("有效筆數", "valid_rows"),
        ("排除非 G 倉筆數", "excluded_rows"), ("異常來源列數", "anomaly_rows"),
    ]
    for offset, (label, attr) in enumerate(stat_labels):
        summary.cell(15 + offset, 1, label)
        summary.cell(15 + offset, 2, getattr(hct_stats, attr))
        summary.cell(15 + offset, 3, getattr(ns_stats, attr))
    summary["A19"] = "異常明細筆數"
    summary["B19"] = len(anomalies)
    summary["A20"] = "HCT 檔案"
    summary["B20"] = hct_stats.file_name
    summary["A21"] = "NetSuite 檔案"
    summary["B21"] = ns_stats.file_name
    summary["A22"] = "核對時間"
    summary["B22"] = datetime.now()
    summary["B22"].number_format = "yyyy-mm-dd hh:mm:ss"

    light_fill = PatternFill("solid", fgColor="D9E1F2")
    for ref in ("A3", "B3", "C3", "A13", "A14", "B14", "C14"):
        summary[ref].font = Font(bold=True)
        summary[ref].fill = light_fill
    summary["A4"].fill = PatternFill("solid", fgColor="C6EFCE")
    for ref in ("A5", "A6"):
        summary[ref].fill = PatternFill("solid", fgColor="FFEB9C")
    for ref in ("A7", "A8", "A9"):
        summary[ref].fill = PatternFill("solid", fgColor="FFC7CE")
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 40
    summary.column_dimensions["C"].width = 18

    def write_table(title: str, columns: list, rows: list, status_col: int | None):
        ws = wb.create_sheet(title)
        for col, header in enumerate(columns, start=1):
            cell = ws.cell(1, col, header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 34
        for row_offset, record in enumerate(rows, start=2):
            for col, header in enumerate(columns, start=1):
                value = record.get(header, "")
                cell = ws.cell(row_offset, col, value)
                if header == "到期日" and isinstance(value, date):
                    cell.number_format = "yyyy-mm-dd"
                elif header == "料號":
                    cell.number_format = "@"
                elif isinstance(value, (int, float)) and header not in ("原始列號",):
                    cell.number_format = "#,##0.####"
                if status_col is not None and col == status_col:
                    colors = _STATUS_FILL.get(str(value))
                    if colors:
                        cell.fill = PatternFill("solid", fgColor=colors[0])
                        cell.font = Font(color=colors[1])
        if not rows:
            ws.cell(2, 1, "無資料" if status_col is not None else "無資料異常")
        import openpyxl.utils
        last_col = openpyxl.utils.get_column_letter(len(columns))
        ws.auto_filter.ref = f"A1:{last_col}{max(len(rows) + 1, 1)}"
        ws.freeze_panes = "A2"
        widths = {"倉別": 10, "料號": 18, "品名／項目": 34, "到期日": 14,
                  "狀態": 18, "結果說明": 42, "異常說明": 32, "原始值": 32, "來源檔名": 28}
        for col, header in enumerate(columns, start=1):
            letter = openpyxl.utils.get_column_letter(col)
            ws.column_dimensions[letter].width = widths.get(header, 16)

    write_table("日期明細", _DETAIL_COLUMNS, detail_rows, status_col=11)
    write_table("料號彙總", _ITEM_COLUMNS, item_rows, status_col=10)
    write_table("資料異常", _ANOMALY_COLUMNS, anomalies, status_col=None)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
