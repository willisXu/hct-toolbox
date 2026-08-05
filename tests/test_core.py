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

from core import bundle_split
from core import inventory
from core.compare import _normalize_customer, _normalize_order_id
from core.inventory import _normalize_date, _normalize_item, _normalize_location, _qty_equal
from core.m00 import convert_rows as m00_convert_rows
from core.return_compare import _normalize_expiry as ret_normalize_expiry, _strip_dr_prefix
from core.shipping import (
    MODE_ORDER,
    _match_arrival_by_keyword,
    _normalize_arrival,
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


def test_inventory_contract_reconcile_against_ns293():
    """代工廠庫存核對報表 × NetSuite:只看 D 倉、排除合計列、單一數量欄當兩用。"""
    contract = _spreadsheetml([
        ["庫存日期", "倉別", "倉別名稱", "品號", "品名", "批號", "庫存數量"],
        ["20260724", "D01", "凱芬妮倉", "300020400025", "品A", "", "100"],
        ["20260724", "D01", "凱芬妮倉", "300020400026", "品B", "", "50"],
        ["", "", "凱芬妮倉 合計", "", "", "", "150"],
        ["20260724", "E01", "非代工廠倉", "300020400027", "品C", "", "10"],
        ["", "", "總計", "", "", "", "160"],
    ])
    ns293 = _spreadsheetml([
        ["地點", "到期日", "項目", "庫存數 總和"],
        ["D01_凱芬妮", "", "300020400025_X", "90"],
        ["D01_凱芬妮", "", "300020400026_X", "50"],
        ["G10", "", "300020400028_X", "5"],
    ])
    assert inventory.detect_source(contract, "c.xls") == inventory.SYSTEM_CONTRACT

    result = inventory.reconcile(ns293, "ns.xls", contract, "c.xls")
    assert result.ext_label == inventory.SYSTEM_CONTRACT
    assert result.output_name.startswith("代工廠庫存核對結果_")
    # 代工廠側排除:合計/總計列 2 筆 + 非 D 倉 1 筆;NetSuite 側排除 G10 1 筆
    assert result.hct_stats.excluded_rows == 3
    assert result.ns_stats.excluded_rows == 1
    assert result.anomalies == []
    by_item = {r["料號"]: r for r in result.item_rows}
    assert by_item["300020400025"]["代工廠 庫存數量"] == 100
    assert by_item["300020400025"]["總庫存差額"] == 10
    assert by_item["300020400025"]["狀態"] == inventory.STATUS_BOTH_DIFF
    assert by_item["300020400026"]["狀態"] == inventory.STATUS_MATCH
    assert "僅 代工廠 存在" in result.statuses


def test_inventory_hct_reconcile_output_unchanged():
    """原本的 HCT × NetSuite 流程不受新格式影響(欄名、檔名、狀態清單)。"""
    hct = _spreadsheetml([
        ["儲區類別", "客戶產品編號", "有效日期", "可出數量", "庫存數量", "產品名稱"],
        ["G10", "SKU1", "20270501", "10", "12", "品A"],
        ["Z99", "SKU2", "", "1", "1", "排除"],
    ])
    ns = _spreadsheetml([
        ["地點", "到期日", "DR_料號", "項目計數 總和", "數量 總和"],
        ["G10_X", "2027/5/1", "SKU1", "10", "12"],
    ])
    result = inventory.reconcile(hct, "hct.xls", ns, "ns.xls")
    assert result.ext_label == inventory.SYSTEM_HCT
    assert result.output_name.startswith("庫存核對結果_")
    assert result.statuses == inventory.ALL_STATUSES
    assert result.hct_stats.excluded_rows == 1
    assert result.item_rows[0]["HCT 可出數量"] == 10
    assert result.item_rows[0]["狀態"] == inventory.STATUS_MATCH


# ------------------------------------------------------------------ 到貨日欄位比對


def test_arrival_keyword_excludes_actual_and_warns():
    headers = {"實際到貨日": 0, "客戶指定到貨日": 1}
    notes = _match_arrival_by_keyword(headers)
    assert headers["dr_預計到貨日期"] == 1
    assert notes and "客戶指定到貨日" in notes[0]


def test_arrival_blank_warning_mentions_merge_only_in_order_mode():
    # 訂單模式（真的有跨單合併）才提「不跨單合併」；M00/調撥單用中性訊息
    _, order_warning = _normalize_arrival(None, merge_hint=True)
    _, neutral_warning = _normalize_arrival(None)
    assert "不跨單合併" in order_warning
    assert "不跨單合併" not in neutral_warning
    assert "空白" in neutral_warning


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


# ------------------------------------------------------------------ ED 訂單拆解

_SPLIT_HEADER = ["單號", "商品貨號", "商品名稱", "商品類型", "數量", "結帳單價", "含稅金額"]


def _split_map():
    """兩品套件（比率 0.65/0.35）＋三品套件（0.333/0.333/0.334，每套 2 個）。"""
    return {
        "29900001": [
            bundle_split.BundleComponent("A001", "單品A", "SPA", 1.0, 0.65),
            bundle_split.BundleComponent("A002", "單品B", "SPB", 1.0, 0.35),
        ],
        "29900002": [
            bundle_split.BundleComponent("B001", "面膜1", "", 2.0, 0.333),
            bundle_split.BundleComponent("B002", "面膜2", "", 2.0, 0.333),
            bundle_split.BundleComponent("B003", "面膜3", "", 2.0, 0.334),
        ],
    }


def _split_result(rows):
    return bundle_split.split_rows([_SPLIT_HEADER] + rows, _split_map(), "來源.xls", "對照.xlsx")


def test_split_expands_qty_and_keeps_amount_total():
    result = _split_result([
        ["S1", "29900001", "兩品組", "商品", 3, 500, 1000],
        ["S1", "101000000001", "一般單品", "商品", 2, 300, 600],
    ])
    assert result.output_rows == 3          # 套件拆 2 列 + 原樣保留 1 列
    assert result.split_rows == 1
    assert result.component_rows == 2
    assert result.amount_matched
    assert result.amount_before == result.amount_after == 1600
    qty_index = _SPLIT_HEADER.index("數量")
    sku_index = _SPLIT_HEADER.index("商品貨號")
    amount_index = _SPLIT_HEADER.index("含稅金額")
    exploded = [r for r in result.preview_rows if r[sku_index] in ("A001", "A002")]
    assert [r[qty_index] for r in exploded] == [3, 3]
    # 650/350，且來源金額是整數時拆出來也維持整數
    assert [r[amount_index] for r in exploded] == [650, 350]
    assert all(isinstance(r[amount_index], int) for r in exploded)


def test_split_rounding_remainder_goes_to_largest_ratio():
    # 864 × (0.333, 0.333, 0.334) 個別四捨五入會多 1 元，尾差要吃回比率最大者
    result = _split_result([["S1", "29900002", "三品組", "商品", 1, 999, 864]])
    amount_index = _SPLIT_HEADER.index("含稅金額")
    qty_index = _SPLIT_HEADER.index("數量")
    amounts = [r[amount_index] for r in result.preview_rows]
    assert sum(amounts) == 864
    assert amounts == [288, 288, 288]
    assert [r[qty_index] for r in result.preview_rows] == [2, 2, 2]  # 每套數量 2


def test_split_marks_and_labels_rows():
    result = _split_result([
        ["S1", "29900001", "兩品組", "商品", 1, 500, 1000],
        ["S1", "101000000001", "一般單品", "商品", 1, 300, 300],
    ])
    columns = result.preview_columns
    status = columns.index("拆解狀態")
    source_sku = columns.index("組合來源料號")
    assert result.preview_rows[0][status] == bundle_split.STATUS_SPLIT
    assert result.preview_rows[0][source_sku] == "29900001"
    assert result.preview_rows[0][columns.index("組合來源品名")] == "兩品組"
    assert result.preview_rows[-1][status] == bundle_split.STATUS_KEPT
    assert result.preview_rows[-1][source_sku] == ""


def test_split_gift_zero_amount_and_unmapped_bundle_warning():
    result = _split_result([
        ["S1", "29900001", "兩品組", "贈品", 1, 0, 0],
        ["S1", "29900099", "未對照的架上組", "贈品", 1, 0, 0],
    ])
    amount_index = _SPLIT_HEADER.index("含稅金額")
    assert [r[amount_index] for r in result.preview_rows[:2]] == [0, 0]
    assert [r["商品貨號"] for r in result.unmapped_rows] == ["29900099"]
    assert result.unmapped_rows[0]["出現列數"] == 1
    assert any("不在對照表" in w for w in result.warnings)


def test_split_298_is_a_real_product_not_an_unmapped_bundle():
    """298 開頭是正常組合品（ERP 有自己的品號與庫存），不拆也不該被當成漏對照。"""
    result = _split_result([["S1", "29826012", "杏仁酸煥膚透亮組", "商品", 1, 556, 556]])
    columns = result.preview_columns
    assert result.split_rows == 0
    assert result.preview_rows[0][columns.index("拆解狀態")] == bundle_split.STATUS_KEPT
    assert result.unmapped_rows == []
    assert result.warnings == []


def test_split_missing_required_column_raises():
    try:
        bundle_split.split_rows([["單號", "商品名稱"], ["S1", "X"]], _split_map())
    except bundle_split.BundleSplitError as exc:
        assert "商品貨號" in str(exc)
    else:
        raise AssertionError("缺少商品貨號欄位應該要報錯")


def test_load_bundle_map_normalizes_and_rejects_unknown_format():
    header = ["套件品號", "單品品號", "單品品名", "規格", "每套數量", "分攤比率"]
    headers = build_header_map(header)
    # 比率未正規化（加總 4）→ 載入後各 0.5
    rows = [["29900003", "C001", "單品C", "", 1, 2], ["29900003", "C002", "單品D", "", 1, 2]]
    comps = bundle_split._load_simple(rows, headers)["29900003"]
    assert [round(c.ratio, 6) for c in comps] == [0.5, 0.5]
    # 比率全空時平均分攤
    blank = [["29900004", "E001", "", "", 1, ""], ["29900004", "E002", "", "", 1, ""]]
    comps = bundle_split._load_simple(blank, headers)["29900004"]
    assert [round(c.ratio, 6) for c in comps] == [0.5, 0.5]
    # 欄位認不出來要報錯，而不是安靜回傳空對照表
    unknown = _spreadsheetml([["欄一", "欄二"], ["a", "b"]])
    try:
        bundle_split.load_bundle_map(unknown, "unknown.xls")
    except bundle_split.BundleSplitError as exc:
        assert "無法辨識" in str(exc)
    else:
        raise AssertionError("無法辨識的對照表格式應該要報錯")


def _spreadsheetml(rows: list[list[str]]) -> bytes:
    """組一份最小的 SpreadsheetML（NetSuite .xls 的格式），供讀檔路徑測試。"""
    cells = "".join(
        "<ss:Row>" + "".join(
            f'<ss:Cell><ss:Data ss:Type="String">{value}</ss:Data></ss:Cell>' for value in row
        ) + "</ss:Row>"
        for row in rows
    )
    return (
        '<?xml version="1.0"?>'
        '<Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
        f'<ss:Worksheet ss:Name="S1"><ss:Table>{cells}</ss:Table></ss:Worksheet></Workbook>'
    ).encode("utf-8")


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
