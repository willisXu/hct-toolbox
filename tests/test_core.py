# -*- coding: utf-8 -*-
"""core 模組回歸測試。

執行方式（擇一）：
    python tests/test_core.py     # 不需安裝任何套件
    python -m pytest tests/       # 有裝 pytest 的話

涵蓋 2026-07 code review 修正的行為：批號解析失敗不合併、NetSuite 0 筆
結果、數字型別日期、Excel 錯誤值過濾、數量容差、到貨日關鍵字比對等。
"""
from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from core.compare import _normalize_customer, _normalize_order_id
from core.inventory import _normalize_date, _normalize_item, _normalize_location, _qty_equal
from core.m00 import convert_rows as m00_convert_rows
from core.return_compare import _normalize_expiry as ret_normalize_expiry, _strip_dr_prefix
from core.shipping import (
    MODE_ORDER,
    _match_arrival_by_keyword,
    _normalize_expiry as ship_normalize_expiry,
    _split_phone,
    _split_postal,
    convert_rows as shipping_convert_rows,
)
from core.xlio import build_header_map, normalize_text, read_workbook, try_parse_date

TEMPLATE_PATH = BASE_DIR / "mappings" / "HCT範本.xlsx"
M00_TEMPLATE_PATH = BASE_DIR / "mappings" / "M00出貨格式.xlsx"


# ------------------------------------------------------------------ xlio


def test_try_parse_date_variants():
    from datetime import date

    assert try_parse_date("2027-05-01") == date(2027, 5, 1)
    assert try_parse_date("2027/5/1") == date(2027, 5, 1)
    assert try_parse_date("20270501") == date(2027, 5, 1)
    # 數字型別的 8 碼日期（Excel 把日期欄存成 Number）
    assert try_parse_date(20270501.0) == date(2027, 5, 1)
    # Excel 序號
    assert try_parse_date(45000) == date(2023, 3, 15)
    assert try_parse_date("LOT-A1") is None
    assert try_parse_date("") is None


def test_normalize_text_excel_error_and_float():
    assert normalize_text("#REF!") == ""
    assert normalize_text(12345.0) == "12345"
    assert normalize_text(" a\nb ") == "a b"


def test_build_header_map_dedup_and_casefold():
    headers = build_header_map(["DR_料號", "dr_料號", "", None, "數量"])
    assert headers == {"dr_料號": 0, "數量": 4}


def test_read_workbook_utf8_bom_xml():
    xml = (
        '<?xml version="1.0"?>'
        '<Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
        '<ss:Worksheet ss:Name="S1"><ss:Table><ss:Row>'
        '<ss:Cell><ss:Data ss:Type="String">A</ss:Data></ss:Cell>'
        "</ss:Row></ss:Table></ss:Worksheet></Workbook>"
    ).encode("utf-8")
    book = read_workbook(b"\xef\xbb\xbf" + xml, "bom.xls")
    assert book["S1"][0][0] == "A"


# ------------------------------------------------------------------ 批號/效期 key


def test_shipping_expiry_unparseable_batches_stay_distinct():
    key_a, value_a, warning_a = ship_normalize_expiry("LOT-A1")
    key_b, _, _ = ship_normalize_expiry("LOT-B2")
    assert key_a != key_b
    assert key_a.startswith("raw:")
    assert value_a is None and warning_a
    assert ship_normalize_expiry("2027-05-01")[0] == "2027-05-01"
    assert ship_normalize_expiry("")[0] == ""


def test_return_expiry_same_text_matches_but_distinct_stays_apart():
    # 兩邊相同的無法解析文字要對得上（casefold）
    assert ret_normalize_expiry("待確認批號B")[0] == ret_normalize_expiry("待確認批號b")[0]
    # 不同文字不能落到同一 key
    assert ret_normalize_expiry("待確認批號B")[0] != ret_normalize_expiry("效期未確認")[0]
    # 可解析日期輸出 YYYY-MM-DD
    assert ret_normalize_expiry("20270501") == ("2027-05-01", "2027-05-01")
    assert ret_normalize_expiry(None) == ("", "")


# ------------------------------------------------------------------ inventory


def test_inventory_normalize_date_numeric_and_text():
    assert _normalize_date(20270101.0) == ("20270101", True)
    assert _normalize_date("2027/1/1") == ("20270101", True)
    assert _normalize_date(None) == ("", True)
    assert _normalize_date("   ") == ("", True)
    assert _normalize_date("abc") == ("", False)
    assert _normalize_date(45000.0)[1] is True  # Excel 序號


def test_inventory_excel_error_treated_as_blank():
    assert _normalize_item("#REF!") == ""
    assert _normalize_location("#N/A") == ""
    assert _normalize_location("g10_x") == "G10"


def test_inventory_quantity_tolerance():
    assert _qty_equal(0.1 + 0.2, 0.3)
    assert not _qty_equal(10.0, 10.5)


# ------------------------------------------------------------------ 到貨日欄位比對


def test_arrival_keyword_excludes_actual_and_warns():
    headers = {"實際到貨日": 0, "客戶指定到貨日": 1}
    notes = _match_arrival_by_keyword(headers)
    assert headers["dr_預計到貨日期"] == 1
    assert notes and "客戶指定到貨日" in notes[0]


def test_arrival_keyword_no_match_when_only_excluded():
    headers = {"實際到貨日": 0}
    notes = _match_arrival_by_keyword(headers)
    assert "dr_預計到貨日期" not in headers
    assert notes == []


# ------------------------------------------------------------------ compare 正規化


def test_compare_normalizers():
    assert _normalize_customer("CUS-001 康是美") == "康是美"
    assert _normalize_order_id("SO12345-1") == "SO12345"
    assert _normalize_order_id("-abc") == ""
    assert _strip_dr_prefix("DR902224") == "902224"
    assert _strip_dr_prefix("DRW-X") == "DRW-X"


# ------------------------------------------------------------------ 電話/地址


def test_split_phone_and_postal():
    assert _split_phone("+886-912-345-678") == ("", "0912345678")
    assert _split_phone("02-2345-6789") == ("0223456789", "")
    assert _split_postal("10041台北市中正區路1號") == ("10041", "台北市中正區路1號")
    assert _split_postal("台北市") == ("", "台北市")


# ------------------------------------------------------------------ 端對端轉換


_E2E_HEADER = [
    "內部 ID", "文件編號", "銷售訂單類型", "日期", "項目", "項目名稱", "DR_料號",
    "序號/批號", "出貨客戶", "出貨數量", "門市/倉儲", "門市/倉儲聯繫人",
    "門市/倉儲電話", "門市/倉儲地址", "DR_預計到貨日期", "DR_預計到貨時段",
]


def _e2e_row(batch: str, qty: float, slot: str = "早(9-12)") -> list:
    return [
        "1", "SO-1", "一般銷售訂單", "2026-07-01", "SKU001_X", "品名A", "SKU001",
        batch, "客戶A", qty, "門市A", "聯絡人", "0912345678",
        "10041台北市中正區路1號", "2026-07-30", slot,
    ]


def test_e2e_distinct_bad_batches_not_merged():
    rows = [_E2E_HEADER, _e2e_row("LOT-A1", 10), _e2e_row("LOT-B2", 5)]
    result = shipping_convert_rows(rows, MODE_ORDER, TEMPLATE_PATH)
    assert result.output_items == 2
    m00_result = m00_convert_rows(rows, MODE_ORDER, M00_TEMPLATE_PATH)
    assert m00_result.output_items == 2


def test_e2e_same_batch_still_merges():
    rows = [_E2E_HEADER, _e2e_row("2027-05-01", 10), _e2e_row("2027-05-01", 5)]
    result = shipping_convert_rows(rows, MODE_ORDER, TEMPLATE_PATH)
    assert result.output_items == 1


def test_e2e_unknown_time_slot_warns():
    rows = [_E2E_HEADER, _e2e_row("2027-05-01", 10, slot="晚（17-20）")]
    m00_result = m00_convert_rows(rows, MODE_ORDER, M00_TEMPLATE_PATH)
    assert any("無法對應" in w for w in m00_result.warnings)


# ------------------------------------------------------------------ 執行器


def main() -> int:
    failures = 0
    tests = [
        (name, fn) for name, fn in sorted(globals().items())
        if name.startswith("test_") and callable(fn)
    ]
    for name, fn in tests:
        try:
            fn()
        except AssertionError:
            import traceback
            failures += 1
            print(f"FAIL {name}")
            traceback.print_exc()
        else:
            print(f"ok   {name}")
    print(f"\n{len(tests) - failures}/{len(tests)} passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
