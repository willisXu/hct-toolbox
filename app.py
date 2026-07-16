# -*- coding: utf-8 -*-
"""HCT 工具箱 — 四合一網頁介面（Streamlit）。

四個功能分頁，對應原本四份 Excel VBA 工具：
  1. 訂單轉換     ← HCT出貨單轉換程式(訂單)
  2. 調撥單轉換   ← HCT出貨單轉換程式(調撥單)
  3. 表格核對     ← HCT表格核對工具
  4. 庫存核對     ← 庫存核對工具

注意：轉換結果一律存進 st.session_state 再顯示。
按下載鈕時 Streamlit 會整頁重跑，若結果只活在「開始轉換」的 if 區塊裡,
重跑後就消失,下載鈕也會跟著失效。
"""
from __future__ import annotations

import traceback
from datetime import datetime
from pathlib import Path

import streamlit as st

from core import compare as compare_mod
from core import inventory as inventory_mod
from core import m00 as m00_mod
from core import shipping

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "mappings" / "HCT範本.xlsx"
M00_TEMPLATE_PATH = BASE_DIR / "mappings" / "M00出貨格式.xlsx"
OUTPUT_DIR = BASE_DIR / "outputs"

FORMAT_HCT = "HCT 銷貨報表格式"
FORMAT_M00 = "M00 出貨格式"

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="HCT 工具箱", page_icon="📦", layout="wide")

st.title("📦 HCT 工具箱")
st.caption("訂單轉換 / 調撥單轉換 / 表格核對 / 庫存核對 — 上傳檔案 → 按按鈕 → 下載結果")


def _save_output(name: str, data: bytes) -> Path:
    OUTPUT_DIR.mkdir(exist_ok=True)
    path = OUTPUT_DIR / name
    path.write_bytes(data)
    return path


def _show_error(exc: Exception) -> None:
    if isinstance(exc, ValueError):
        st.error(str(exc))
    else:
        st.error(f"發生未預期的錯誤：{exc}")
        with st.expander("技術細節（回報問題時請截圖此區塊）"):
            st.code(traceback.format_exc())


# ------------------------------------------------------------ 訂單 / 調撥單


def _convert_tab(mode: str, title: str, source_hint: str, key_prefix: str) -> None:
    st.subheader(title)
    st.markdown(
        f"1. 從 NetSuite 匯出 **{source_hint}**（.xls / .xlsx 皆可）\n"
        "2. 把檔案拖進下方框框\n"
        "3. 選擇輸出格式，按「開始轉換」，完成後點「下載結果」"
    )
    uploaded = st.file_uploader(
        f"選擇 {source_hint}", type=["xls", "xlsx", "xlsm"], key=f"{key_prefix}_file"
    )
    st.markdown("**輸出格式**")
    output_format = st.radio(
        "輸出格式",
        [FORMAT_HCT, FORMAT_M00],
        captions=[
            "HCT 系統銷貨報表，可跨單合併",
            "第三方物流匯入用，不合併訂單、固定商品狀態/出貨方式/物流方式",
        ],
        key=f"{key_prefix}_format",
        horizontal=True,
        label_visibility="collapsed",
    )
    state_key = f"{key_prefix}_result"

    if uploaded is not None and st.button("🚀 開始轉換", key=f"{key_prefix}_run", type="primary"):
        try:
            with st.spinner("轉換中..."):
                if output_format == FORMAT_M00:
                    result = m00_mod.convert(uploaded.getvalue(), uploaded.name, mode, M00_TEMPLATE_PATH)
                else:
                    result = shipping.convert(uploaded.getvalue(), uploaded.name, mode, TEMPLATE_PATH)
        except Exception as exc:  # noqa: BLE001 - 給使用者看得懂的訊息
            st.session_state.pop(state_key, None)
            _show_error(exc)
        else:
            saved_path = _save_output(result.output_name, result.output_bytes)
            st.session_state[state_key] = (result, str(saved_path), uploaded.name)

    stored = st.session_state.get(state_key)
    if not stored:
        return
    result, saved_path, source_name = stored
    is_m00 = hasattr(result, "order_count")
    group_label = "訂單數" if is_m00 else "送貨單數"
    group_count = result.order_count if is_m00 else result.shipments
    st.success(
        f"「{source_name}」轉換完成！{group_label} **{group_count}**、"
        f"輸出品項 **{result.output_items}**、"
        f"無法轉換明細 **{result.problem_count}**、警告 **{len(result.warnings)}**"
    )
    st.download_button(
        f"⬇️ 下載結果（{result.output_name}）",
        data=result.output_bytes,
        file_name=result.output_name,
        mime=EXCEL_MIME,
        key=f"{key_prefix}_download",
    )
    st.caption(f"已同時存一份到：{saved_path}")

    if result.problem_rows:
        detail_hint = "（詳見輸出檔的「有問題訂單」工作表）" if not is_m00 else "（M00 格式輸出檔不含問題明細，僅顯示於下方）"
        st.warning(f"有 {len(result.problem_rows)} 筆明細無法轉換{detail_hint}")
        st.dataframe(result.problem_rows, use_container_width=True)
    if result.warnings:
        with st.expander(f"⚠️ 警告訊息（{len(result.warnings)} 則）"):
            for warning in result.warnings:
                st.text(f"• {warning}")


tab_order, tab_transfer, tab_compare, tab_inventory = st.tabs(
    ["🛒 訂單轉換", "🔄 調撥單轉換", "🔍 表格核對", "📊 庫存核對"]
)

with tab_order:
    _convert_tab(
        shipping.MODE_ORDER,
        "銷售訂單 → HCT 銷貨報表",
        "銷售訂單未出貨明細",
        "order",
    )

with tab_transfer:
    _convert_tab(
        shipping.MODE_TRANSFER,
        "調撥單 → HCT 銷貨報表",
        "調撥單未出貨明細",
        "transfer",
    )

# ------------------------------------------------------------ 表格核對

with tab_compare:
    st.subheader("未出貨明細 × 銷貨單明細 核對")
    st.markdown(
        "上傳 **未出貨明細** 與 **銷貨單明細** 各一份（順序不限，程式會自動辨識），"
        "產出數量核對與訂單編號核對結果。"
    )
    col1, col2 = st.columns(2)
    with col1:
        file_a = st.file_uploader("檔案一", type=["xls", "xlsx", "xlsm"], key="cmp_a")
    with col2:
        file_b = st.file_uploader("檔案二", type=["xls", "xlsx", "xlsm"], key="cmp_b")

    if file_a is not None and file_b is not None:
        if st.button("🚀 開始核對", key="cmp_run", type="primary"):
            try:
                with st.spinner("核對中..."):
                    result = compare_mod.compare(
                        file_a.getvalue(), file_a.name,
                        file_b.getvalue(), file_b.name,
                    )
            except Exception as exc:  # noqa: BLE001
                st.session_state.pop("cmp_result", None)
                _show_error(exc)
            else:
                saved_path = _save_output(result.output_name, result.output_bytes)
                st.session_state["cmp_result"] = (result, str(saved_path))

    stored = st.session_state.get("cmp_result")
    if stored:
        result, saved_path = stored
        st.success(
            f"核對完成！數量核對 **{result.quantity_rows}** 筆、"
            f"訂單編號核對 **{result.order_rows}** 筆"
        )
        metric_cols = st.columns(4)
        for idx, status in enumerate(["一致", "數量不符", "僅未出貨明細", "僅銷貨單"]):
            metric_cols[idx].metric(
                f"數量核對:{status}",
                result.quantity_status_counts.get(status, 0),
            )
        st.download_button(
            f"⬇️ 下載核對結果（{result.output_name}）",
            data=result.output_bytes,
            file_name=result.output_name,
            mime=EXCEL_MIME,
            key="cmp_download",
        )
        st.caption(f"已同時存一份到：{saved_path}")

# ------------------------------------------------------------ 庫存核對

with tab_inventory:
    st.subheader("HCT × NetSuite 庫存核對")
    st.markdown(
        "上傳 **HCT 庫存報表** 與 **NetSuite 庫存報表** 各一份，順序不限，程式會自動辨識。\n\n"
        "NetSuite 報表支援三種格式：物流核對版（DR_料號/庫存編號/在庫量/可用）、"
        "業務助理版（項目/庫存數 總和）、舊版（DR_料號/項目計數 總和/數量 總和）。\n\n"
        "只核對 G00 / G10 / G30 / G40 / G80 / G90 倉，差異基準為 **HCT－NetSuite**。"
    )
    col1, col2 = st.columns(2)
    with col1:
        inv_a = st.file_uploader("檔案一", type=["xls", "xlsx", "xlsm"], key="inv_a")
    with col2:
        inv_b = st.file_uploader("檔案二", type=["xls", "xlsx", "xlsm"], key="inv_b")

    if inv_a is not None and inv_b is not None:
        if st.button("🚀 開始核對", key="inv_run", type="primary"):
            try:
                with st.spinner("核對中...（大檔案可能需要一分鐘）"):
                    result = inventory_mod.reconcile(
                        inv_a.getvalue(), inv_a.name,
                        inv_b.getvalue(), inv_b.name,
                    )
            except Exception as exc:  # noqa: BLE001
                st.session_state.pop("inv_result", None)
                _show_error(exc)
            else:
                saved_path = _save_output(result.output_name, result.output_bytes)
                st.session_state["inv_result"] = (result, str(saved_path))

    stored = st.session_state.get("inv_result")
    if stored:
        result, saved_path = stored
        st.success(
            f"核對完成！日期明細 **{len(result.detail_rows)}** 筆、"
            f"料號彙總 **{len(result.item_rows)}** 筆、"
            f"資料異常 **{len(result.anomalies)}** 筆"
        )
        metric_cols = st.columns(6)
        for idx, status in enumerate(inventory_mod.ALL_STATUSES):
            metric_cols[idx].metric(status, result.item_status_counts.get(status, 0))
        st.download_button(
            f"⬇️ 下載核對結果（{result.output_name}）",
            data=result.output_bytes,
            file_name=result.output_name,
            mime=EXCEL_MIME,
            key="inv_download",
        )
        st.caption(f"已同時存一份到：{saved_path}")

st.divider()
st.caption(
    f"HCT 工具箱 v1.1 ｜ 結果檔會同步存放在 outputs 資料夾 ｜ {datetime.now():%Y-%m-%d}"
)
