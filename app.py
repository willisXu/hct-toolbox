# -*- coding: utf-8 -*-
"""HCT 工具箱 — 網頁介面（Streamlit）。

分頁對應原本的 Excel VBA 工具，另外加了 NetSuite 直接抓取與退貨核對：
  1. 訂單轉換     ← HCT出貨單轉換程式(訂單)
  2. 調撥單轉換   ← HCT出貨單轉換程式(調撥單)
  3. 表格核對     ← HCT表格核對工具
  4. 退貨核對     ← 客戶退貨授權明細 × HCT 退貨入庫格式
  5. 庫存核對     ← 庫存核對工具
  6. ED 訂單拆解  ← 每日銷售串接訂單明細的 299 組合料號 → 單品

注意：轉換結果一律存進 st.session_state 再顯示。
按下載鈕時 Streamlit 會整頁重跑，若結果只活在「開始轉換」的 if 區塊裡,
重跑後就消失,下載鈕也會跟著失效。
"""
from __future__ import annotations

import hashlib
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import yaml

from core import bundle_split as bundle_split_mod
from core import compare as compare_mod
from core import export_log as export_log_mod
from core import export_store as export_store_mod
from core import inventory as inventory_mod
from core import m00 as m00_mod
from core import netsuite as netsuite_mod
from core import return_compare as return_compare_mod
from core import shipping
from core import table_filter

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "mappings" / "HCT範本.xlsx"
M00_TEMPLATE_PATH = BASE_DIR / "mappings" / "M00出貨格式.xlsx"
BUNDLE_MAP_PATH = BASE_DIR / "mappings" / "組合對照表.xlsx"

FORMAT_HCT = "HCT 銷貨報表格式"
FORMAT_M00 = "M00 出貨格式"

# 備忘錄來源選項：顯示名稱 → shipping 模組的 memo_source 常數
MEMO_LINE_LABEL = "明細行備忘錄"
MEMO_MAIN_LABEL = "主要備忘錄"
MEMO_SOURCE_MAP = {
    MEMO_LINE_LABEL: shipping.MEMO_SOURCE_LINE,
    MEMO_MAIN_LABEL: shipping.MEMO_SOURCE_MAIN,
}


# 合併方式選項：顯示名稱 → shipping 模組的 merge_by 常數
MERGE_RECIPIENT_LABEL = "收件條件（原規則）"
MERGE_PO_LABEL = "客戶採購單編號"
MERGE_BY_MAP = {
    MERGE_RECIPIENT_LABEL: shipping.MERGE_BY_RECIPIENT,
    MERGE_PO_LABEL: shipping.MERGE_BY_PURCHASE_ORDER,
}


def _merge_by_radio(key: str) -> str:
    """跨單合併方式選擇（回傳 shipping 的 merge_by 常數）。僅 HCT 銷貨報表格式有效。"""
    st.markdown("**合併方式**（僅 HCT 銷貨報表格式有效，M00 一律不合併）")
    label = st.radio(
        "合併方式",
        [MERGE_RECIPIENT_LABEL, MERGE_PO_LABEL],
        captions=[
            "預計到貨日＋出貨客戶＋門市＋地址相同，且訂單類型 ≥2 種並含一般銷售訂單才合併",
            "客戶採購單編號相同即合併為一張送貨單（編號空白的不合併）",
        ],
        key=key,
        horizontal=True,
        label_visibility="collapsed",
        help="以採購單合併時不看訂單類型；合併群組內收件資訊不一致會警告並以第一筆為準。",
    )
    return MERGE_BY_MAP[label]


def _memo_source_radio(key: str) -> str:
    """備忘錄來源選擇（回傳 shipping 的 memo_source 常數）。"""
    st.markdown("**備忘錄來源**")
    label = st.radio(
        "備忘錄來源",
        [MEMO_LINE_LABEL, MEMO_MAIN_LABEL],
        captions=[
            "吃明細行的「備忘錄」欄",
            "吃主要層級的「備忘錄 (主要)」欄",
        ],
        key=key,
        horizontal=True,
        label_visibility="collapsed",
        help=(
            "來源報表本身有「DR_溫馨提醒」欄時，以該欄為準，此選項不生效。\n\n"
            "新版報表的明細行「備忘錄」常被系統回填成品名，"
            "與品名相同的內容會自動忽略不進備註；訂單備註在"
            "「備忘錄 (主要)」時請改選主要備忘錄。"
        ),
    )
    return MEMO_SOURCE_MAP[label]

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="HCT 工具箱", page_icon="📦", layout="wide")

st.title("📦 HCT 工具箱")
st.caption("訂單轉換 / 調撥單轉換 / 表格核對 / 退貨核對 / 庫存核對 — 上傳檔案 → 按按鈕 → 下載結果")


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
        "3. 選擇輸出格式與備忘錄來源，按「開始轉換」，完成後點「下載結果」"
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
    memo_source = _memo_source_radio(f"{key_prefix}_memo_source")
    merge_by = _merge_by_radio(f"{key_prefix}_merge_by")
    state_key = f"{key_prefix}_result"

    if uploaded is not None and st.button("🚀 開始轉換", key=f"{key_prefix}_run", type="primary"):
        try:
            with st.spinner("轉換中..."):
                if output_format == FORMAT_M00:
                    result = m00_mod.convert(
                        uploaded.getvalue(), uploaded.name, mode, M00_TEMPLATE_PATH,
                        memo_source=memo_source,
                    )
                else:
                    result = shipping.convert(
                        uploaded.getvalue(), uploaded.name, mode, TEMPLATE_PATH,
                        memo_source=memo_source, merge_by=merge_by,
                    )
        except Exception as exc:  # noqa: BLE001 - 給使用者看得懂的訊息
            st.session_state.pop(state_key, None)
            _show_error(exc)
        else:
            st.session_state[state_key] = (result, uploaded.name)

    stored = st.session_state.get(state_key)
    if not stored:
        return
    result, source_name = stored
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

    if result.problem_rows:
        detail_hint = "（詳見輸出檔的「有問題訂單」工作表）" if not is_m00 else "（M00 格式輸出檔不含問題明細，僅顯示於下方）"
        st.warning(f"有 {len(result.problem_rows)} 筆明細無法轉換{detail_hint}")
        st.dataframe(result.problem_rows, use_container_width=True)
    if result.warnings:
        with st.expander(f"⚠️ 警告訊息（{len(result.warnings)} 則）"):
            for warning in result.warnings:
                st.text(f"• {warning}")


(
    tab_order, tab_transfer, tab_netsuite, tab_compare,
    tab_return, tab_inventory, tab_bundle,
) = st.tabs(
    ["🛒 訂單轉換", "🔄 調撥單轉換", "🔗 NetSuite 直接抓取", "🔍 表格核對",
     "🔁 退貨核對", "📊 庫存核對", "🧩 ED 訂單拆解"]
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

# ------------------------------------------------------------ NetSuite 直接抓取

# 「已轉出」紀錄：轉換成功的單號記下來，下次抓到同一張單會標記＋提醒。
# 正本優先存 Google 試算表（secrets 有設 [gsheet_log] 時），雲端部署重啟也不會掉；
# 沒設定就退回本機 JSON 檔（雲端的檔案系統是暫時的，重啟會清空）。
NS_EXPORT_LOG_PATH = BASE_DIR / "data" / "ns_export_log.json"
# 寫進 Google 試算表失敗時，先把這一批記在這裡，等連線恢復再補寫回去。
NS_PENDING_LOG_PATH = BASE_DIR / "data" / "ns_export_pending.json"

EXPORT_FILTER_ALL = "全部"
EXPORT_FILTER_NEW = "只看未轉出過"
EXPORT_FILTER_DONE = "只看已轉出過"


def _load_saved_searches() -> list[dict]:
    path = BASE_DIR / "mappings" / "netsuite_saved_searches.yaml"
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or []
    except FileNotFoundError:
        return []
    return [item for item in data if item.get("search_id")]


def _netsuite_client() -> tuple[netsuite_mod.NetSuiteClient, str]:
    try:
        cfg = st.secrets["netsuite"]
    except Exception as exc:
        raise ValueError(
            "尚未設定 NetSuite 連線資訊。請依 .streamlit/secrets.toml.example，"
            "在 Streamlit 的 Secrets 設定裡填入 netsuite 區塊。"
        ) from exc
    required = ("account_id", "consumer_key", "consumer_secret", "token_id", "token_secret", "restlet_url")
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        raise ValueError(f"NetSuite 設定缺少：{'、'.join(missing)}")
    return netsuite_mod.NetSuiteClient(cfg), cfg["restlet_url"]


@st.cache_resource(show_spinner=False)
def _sheet_store(fingerprint: str):
    """Google 試算表 store。憑證與連線重用，不然每次重跑都要重新取 token。"""
    del fingerprint  # 只用來讓設定變更時換一個 cache 項目
    return export_store_mod.build_sheet_store(st.secrets["gsheet_log"])


def _export_store() -> tuple[object, str | None]:
    """回傳（紀錄存放後端, 設定錯誤訊息）。沒設定 Google 試算表就用本機檔案。"""
    try:
        cfg = st.secrets["gsheet_log"]
    except Exception:
        return export_store_mod.LocalStore(NS_EXPORT_LOG_PATH), None
    try:
        fingerprint = f"{cfg.get('spreadsheet_id') or cfg.get('spreadsheet_url')}|{cfg.get('worksheet')}"
        return _sheet_store(fingerprint), None
    except Exception as exc:  # noqa: BLE001 - 設定錯不該讓整個分頁掛掉
        return export_store_mod.LocalStore(NS_EXPORT_LOG_PATH), str(exc)


def _export_state(store, refresh: bool = False) -> dict:
    """讀一次紀錄後存在 session：Google 試算表每次重跑都連線會讓介面很鈍。

    寫入（轉換、匯入、清除）之後會帶 refresh=True 重讀，所以自己操作的
    結果一定看得到；別人同時寫進試算表的，要按「🔄 重新整理」才會出現。
    """
    cache = st.session_state.get("ns_log_cache")
    if refresh or not cache or cache.get("name") != store.name:
        try:
            log = store.load()
            error = None
        except Exception as exc:  # noqa: BLE001 - 讀不到紀錄不該擋住轉換
            log = export_log_mod.empty_log()
            error = str(exc)
        cache = {
            "name": store.name,
            "log": log,
            "error": error,
            "at": datetime.now().strftime("%H:%M:%S"),
        }
        st.session_state["ns_log_cache"] = cache
    return cache


def _unique_columns(header_row: list) -> list[str]:
    """欄名去重＋補空欄名（saved search 可能有同名欄，DataFrame 不允許重複欄名）。"""
    used: dict[str, int] = {}
    result: list[str] = []
    for index, value in enumerate(header_row):
        name = str(value).strip() if value is not None else ""
        if not name:
            name = f"欄{index + 1}"
        if name in used:
            used[name] += 1
            name = f"{name} ({used[name]})"
        else:
            used[name] = 1
        result.append(name)
    return result


def _date_range(picked) -> tuple[object, object]:
    """st.date_input 的區間選擇：使用者只點了起日時回 (起, None)，沒選回 (None, None)。"""
    if isinstance(picked, (list, tuple)):
        if len(picked) >= 2:
            return picked[0], picked[1]
        if len(picked) == 1:
            return picked[0], None
        return None, None
    return (picked, picked) if picked else (None, None)


def _ns_filter_panel(
    ns_rows: list[list[object]], columns: list[str], fetch_seq: int, exported_flags: list[bool]
) -> tuple[list[int], str]:
    """篩選面板。回傳（通過篩選的資料列索引, 篩選條件簽章）。

    簽章併進 data_editor 的 key：篩選條件一變就換一個 widget，避免
    Streamlit 把上一組篩選結果的勾選差異（依列位置記錄）套到不同的列上。
    """
    prefix = f"ns_filter_{fetch_seq}"
    total = len(ns_rows) - 1
    with st.expander("🔎 篩選條件（先篩出要的資料，再勾選轉換）", expanded=True):
        keyword = st.text_input(
            "關鍵字",
            key=f"{prefix}_keyword",
            placeholder="跨所有欄位搜尋，空白分隔多個詞＝要同時符合",
            help="例如輸入「SO12345 台北」＝該列同時出現這兩個詞才留下。",
        )
        chosen = st.multiselect(
            "依欄位篩選（可複選欄位）",
            columns,
            key=f"{prefix}_columns",
            help="選了哪些欄位，下面就出現對應的篩選器：一般欄位是值多選，日期欄位是日期區間，相異值太多的欄位改用「包含文字」。",
        )
        value_filters: dict[int, set[str]] = {}
        text_filters: dict[int, str] = {}
        date_ranges: dict[int, tuple[object, object]] = {}
        if chosen:
            grid = st.columns(min(len(chosen), 3))
            for position, name in enumerate(chosen):
                col_index = columns.index(name)
                with grid[position % len(grid)]:
                    if table_filter.is_date_column(ns_rows, col_index):
                        picked = st.date_input(
                            f"{name}（日期區間）",
                            value=(),
                            key=f"{prefix}_date_{col_index}",
                            format="YYYY-MM-DD",
                        )
                        start, end = _date_range(picked)
                        if start or end:
                            date_ranges[col_index] = (start, end)
                    else:
                        values = table_filter.distinct_values(
                            ns_rows, col_index, limit=table_filter.MAX_CHOICES
                        )
                        if len(values) <= table_filter.MAX_CHOICES:
                            picked_values = st.multiselect(
                                name, values, key=f"{prefix}_value_{col_index}"
                            )
                            if picked_values:
                                value_filters[col_index] = set(picked_values)
                        else:
                            typed = st.text_input(
                                f"{name}（包含文字）",
                                key=f"{prefix}_text_{col_index}",
                                help=f"這一欄相異值超過 {table_filter.MAX_CHOICES} 種，改用包含文字比對。",
                            )
                            if typed.strip():
                                text_filters[col_index] = typed.strip()

        export_filter = st.radio(
            "轉出紀錄",
            [EXPORT_FILTER_ALL, EXPORT_FILTER_NEW, EXPORT_FILTER_DONE],
            key=f"{prefix}_export",
            horizontal=True,
            help="依「這張單先前有沒有轉出過出貨格式表格」篩選。",
        )

        visible = table_filter.filter_indices(
            ns_rows, keyword, value_filters, text_filters, date_ranges
        )
        if export_filter == EXPORT_FILTER_NEW:
            visible = [i for i in visible if not exported_flags[i]]
        elif export_filter == EXPORT_FILTER_DONE:
            visible = [i for i in visible if exported_flags[i]]

        reset_col, info_col = st.columns([1, 4])
        with reset_col:
            if st.button("🧹 清除篩選", key=f"{prefix}_reset", use_container_width=True):
                for state_key in [k for k in st.session_state if k.startswith(prefix)]:
                    del st.session_state[state_key]
                st.rerun()
        with info_col:
            st.caption(f"篩選後 **{len(visible)}** / 共 **{total}** 筆")

    signature_source = (
        keyword,
        sorted((index, sorted(values)) for index, values in value_filters.items()),
        sorted(text_filters.items()),
        sorted((index, str(start), str(end)) for index, (start, end) in date_ranges.items()),
        export_filter,
    )
    signature = hashlib.md5(repr(signature_source).encode("utf-8")).hexdigest()[:8]
    return visible, signature


with tab_netsuite:
    st.subheader("Saved Search → 直接抓取轉換")
    st.markdown(
        "1. 選擇要抓取的 saved search，按「抓取資料」\n"
        "2. 用「篩選條件」篩出要的資料（關鍵字／欄位值／日期區間／有沒有轉出過）\n"
        "3. 在表格勾選要轉換的列（預設全選；「全選／取消全選」只作用在目前篩選結果）\n"
        "4. 選擇輸出格式與備忘錄來源，按「開始轉換」，完成後點「下載結果」\n\n"
        "轉換成功的單號會記進「轉出紀錄」，下次抓到同一張單會在表格標記 ⚠️，"
        "勾選時另外跳出提醒，避免重複轉出。"
    )

    ns_store, ns_store_config_error = _export_store()
    if ns_store_config_error:
        st.error(
            "Google 試算表設定有問題，這次改用本機檔案暫存轉出紀錄"
            f"（伺服器重啟後會消失）：{ns_store_config_error}"
        )
    ns_log_state = _export_state(ns_store)
    export_records = ns_log_state["log"].get("records", {})
    if ns_log_state["error"]:
        st.error(
            f"讀取轉出紀錄失敗：{ns_log_state['error']}\n\n"
            "這次不會顯示「已轉出」標記與提醒，請確認後按「🗂️ 轉出紀錄管理 → 🔄 重新整理」。"
        )
    ns_pending_store = export_store_mod.LocalStore(NS_PENDING_LOG_PATH)
    ns_pending_records = ns_pending_store.load().get("records", {})
    if ns_pending_records:
        st.warning(
            f"有 **{len(ns_pending_records)}** 張單的轉出紀錄還沒寫進 {ns_store.label}"
            "（先前寫入失敗，先暫存在伺服器本機）。"
            "請到下方「🗂️ 轉出紀錄管理」按「📤 補寫回正本」，"
            "否則伺服器重啟後這幾筆紀錄就沒了。"
        )

    searches = _load_saved_searches()
    if not searches:
        st.info("尚未設定 saved search，請編輯 mappings/netsuite_saved_searches.yaml 填入 search_id。")
    else:
        labels = [item["label"] for item in searches]
        chosen_label = st.selectbox("選擇 saved search", labels, key="ns_search_select")
        chosen = next(item for item in searches if item["label"] == chosen_label)

        if st.button("📥 抓取資料", key="ns_fetch", type="primary"):
            try:
                with st.spinner("正在向 NetSuite 抓取資料..."):
                    client, restlet_url = _netsuite_client()
                    rows = client.run_saved_search(restlet_url, chosen["search_id"])
            except Exception as exc:
                st.session_state.pop("ns_rows", None)
                st.session_state.pop("ns_result", None)
                _show_error(exc)
            else:
                st.session_state.pop("ns_result", None)
                if len(rows) <= 1:
                    # 只有 header 列：saved search 沒有符合條件的結果，是正常
                    # 業務狀態（例如當天訂單都已出貨），不是錯誤。
                    st.session_state.pop("ns_rows", None)
                    st.info(f"「{chosen_label}」目前沒有符合條件的資料。")
                else:
                    # 每次抓取遞增序號並併入 widget key：若只靠標籤+列數，
                    # 「重抓後列數剛好相同」會沿用上一批資料的勾選狀態，
                    # 悄悄把舊的取消勾選套到內容不同的新資料列上。
                    st.session_state["ns_fetch_seq"] = st.session_state.get("ns_fetch_seq", 0) + 1
                    st.session_state["ns_rows"] = (rows, chosen["mode"], chosen_label)

        fetched = st.session_state.get("ns_rows")
        if fetched:
            ns_rows, ns_mode, ns_label = fetched
            fetch_seq = st.session_state.get("ns_fetch_seq", 0)
            total_rows = len(ns_rows) - 1
            st.success(f"已抓取「{ns_label}」共 {total_rows} 筆資料列。")

            # 表格與單號只在「新抓一批」時重算，篩選/勾選重跑不必重建。
            ns_columns = _unique_columns(ns_rows[0])

            # 品名欄漏抓時要講清楚，否則轉出的檔案品名整欄空白卻沒有任何提示。
            # 常見原因：saved search 該欄沒設自訂 Label，RESTlet 只能回欄位內部
            # ID（netsuite.py 的 _HEADER_ALIASES 已收常見寫法）；若欄名不在對照
            # 表裡、或 saved search 根本沒加這欄，就得回 NetSuite 補。
            if not any(name in ns_columns for name in ("項目名稱", "顯示名稱")):
                st.warning(
                    "這批資料裡找不到品名欄（需要「項目名稱」或「顯示名稱」），"
                    "轉出的檔案品名會是空白。請到 NetSuite 確認這個 saved search "
                    "有加品名欄位、並替它設定中文 Label。實際抓回來的欄名："
                    + "、".join(str(c) for c in ns_columns)
                )
            cache = st.session_state.get("ns_cache")
            if not cache or cache.get("seq") != fetch_seq:
                cache = {
                    "seq": fetch_seq,
                    "df": pd.DataFrame(ns_rows[1:], columns=ns_columns),
                    "keys": export_log_mod.extract_keys(ns_rows),
                }
                st.session_state["ns_cache"] = cache
            base_df = cache["df"]
            row_keys = cache["keys"]

            exported_flags = [bool(key) and key in export_records for key in row_keys]
            if any(exported_flags):
                st.info(
                    f"這批資料裡有 **{sum(exported_flags)}** 筆明細所屬的單號先前已經轉出過"
                    "（表格「轉出紀錄」欄標記 ⚠️）；只想看沒轉過的，可在篩選條件選「只看未轉出過」。"
                )

            visible_idx, filter_sig = _ns_filter_panel(
                ns_rows, ns_columns, fetch_seq, exported_flags
            )

            # 勾選狀態存原始列索引，不隨篩選改變而遺失（切換篩選再切回來還在）。
            select_key = f"ns_selected_{fetch_seq}"
            if select_key not in st.session_state:
                st.session_state[select_key] = set(range(total_rows))
            selected: set[int] = st.session_state[select_key]

            # nonce：按下「全選/取消全選」時 +1，換一組新的 data_editor key，
            # 強制 Streamlit 重建 widget、整批套用新的勾選狀態（否則殘留的
            # 個別勾選差異會蓋掉程式剛設好的值）。
            nonce_key = f"ns_select_nonce_{fetch_seq}"
            if nonce_key not in st.session_state:
                st.session_state[nonce_key] = 0

            btn_col1, btn_col2, _btn_col3 = st.columns([1, 1, 5])
            with btn_col1:
                if st.button("☑️ 全選（篩選結果）", key="ns_select_all", use_container_width=True):
                    selected.update(visible_idx)
                    st.session_state[nonce_key] += 1
            with btn_col2:
                if st.button("⬜ 取消全選（篩選結果）", key="ns_deselect_all", use_container_width=True):
                    selected.difference_update(visible_idx)
                    st.session_state[nonce_key] += 1

            if not visible_idx:
                st.info("目前的篩選條件沒有任何符合的資料列，請調整條件或按「清除篩選」。")
            else:
                view = base_df.loc[visible_idx].copy()
                view.insert(0, "選取", [index in selected for index in visible_idx])
                if any(exported_flags):
                    view.insert(1, "轉出紀錄", [
                        export_log_mod.short_label(export_records[row_keys[index]])
                        if exported_flags[index] else ""
                        for index in visible_idx
                    ])
                editor_key = f"ns_editor_{fetch_seq}_{filter_sig}_{st.session_state[nonce_key]}"
                edited = st.data_editor(
                    view,
                    key=editor_key,
                    use_container_width=True,
                    hide_index=True,
                    height=420,
                    disabled=[c for c in view.columns if c != "選取"],
                    column_config={
                        "選取": st.column_config.CheckboxColumn(
                            "✅ 選取", default=True, width="small", help="勾選要轉換的資料列"
                        ),
                        "轉出紀錄": st.column_config.TextColumn(
                            "⚠️ 轉出紀錄", width="medium", help="這張單先前已經轉出過出貨格式表格的次數與時間"
                        ),
                    },
                )
                # 把這次的勾選寫回原始列索引（data_editor 回傳的列序與傳入相同）。
                edited_flags = list(edited["選取"])
                for position, row_index in enumerate(visible_idx):
                    if position >= len(edited_flags):
                        break
                    if bool(edited_flags[position]):
                        selected.add(row_index)
                    else:
                        selected.discard(row_index)

                visible_selected = [index for index in visible_idx if index in selected]
                hidden_selected = len(selected) - len(visible_selected)
                st.caption(
                    f"已勾選 **{len(visible_selected)}** / 篩選後 **{len(visible_idx)}** 筆"
                    f"（這批資料共 {total_rows} 筆）"
                )
                if hidden_selected:
                    st.caption(
                        f"ℹ️ 另有 **{hidden_selected}** 筆已勾選、但被目前篩選條件隱藏，"
                        "**不會**被轉換（只轉換上面表格裡勾選的列）。"
                    )

                # 已轉出提醒：勾到先前轉過的單就明白列出來，讓人再確認一次。
                repeat: dict[str, int] = {}
                for index in visible_selected:
                    if exported_flags[index]:
                        key = row_keys[index]
                        repeat[key] = repeat.get(key, 0) + 1
                if repeat:
                    st.warning(
                        f"⚠️ 已勾選的資料裡有 **{len(repeat)}** 張單（共 {sum(repeat.values())} 筆明細）"
                        "先前已經轉出過出貨格式表格，請確認是不是真的要再轉一次。"
                    )
                    st.dataframe(
                        pd.DataFrame([
                            {
                                "單號": key,
                                "本次勾選明細": count,
                                "轉出紀錄": export_log_mod.describe(export_records[key]),
                            }
                            for key, count in repeat.items()
                        ]),
                        use_container_width=True,
                        hide_index=True,
                    )

                st.markdown("**輸出格式**")
                ns_format = st.radio(
                    "輸出格式",
                    [FORMAT_HCT, FORMAT_M00],
                    key="ns_format",
                    horizontal=True,
                    label_visibility="collapsed",
                )
                ns_memo_source = _memo_source_radio("ns_memo_source")
                ns_merge_by = _merge_by_radio("ns_merge_by")

                if st.button("🚀 開始轉換", key="ns_run", type="primary"):
                    if not visible_selected:
                        st.warning("目前篩選結果裡沒有勾選任何資料列。")
                    else:
                        # 直接取原始列（不經過 DataFrame），避免型別被 pandas 動到。
                        rows_for_convert = [list(ns_rows[0])] + [
                            list(ns_rows[index + 1]) for index in visible_selected
                        ]
                        try:
                            with st.spinner("轉換中..."):
                                if ns_format == FORMAT_M00:
                                    result = m00_mod.convert_rows(
                                        rows_for_convert, ns_mode, M00_TEMPLATE_PATH,
                                        memo_source=ns_memo_source,
                                    )
                                else:
                                    result = shipping.convert_rows(
                                        rows_for_convert, ns_mode, TEMPLATE_PATH,
                                        memo_source=ns_memo_source, merge_by=ns_merge_by,
                                    )
                        except Exception as exc:
                            st.session_state.pop("ns_result", None)
                            _show_error(exc)
                        else:
                            st.session_state["ns_result"] = (result, ns_label)
                            # 轉換成功才記錄。寫紀錄失敗不影響已經產好的結果，
                            # 只警告並把這一批存進待補紀錄，之後可以補寫回正本。
                            converted_keys = [row_keys[index] for index in visible_selected]
                            try:
                                ns_store.record(converted_keys, ns_format)
                            except Exception as exc:  # noqa: BLE001
                                st.session_state["ns_log_error"] = str(exc)
                                try:
                                    ns_pending_store.record(converted_keys, ns_format)
                                except OSError:
                                    pass  # 連本機都寫不了，警告訊息已經夠了
                            else:
                                st.session_state.pop("ns_log_error", None)
                            _export_state(ns_store, refresh=True)
                            st.rerun()

        stored = st.session_state.get("ns_result")
        if stored:
            result, source_label = stored
            is_m00 = hasattr(result, "order_count")
            group_label = "訂單數" if is_m00 else "送貨單數"
            group_count = result.order_count if is_m00 else result.shipments
            st.success(
                f"「{source_label}」轉換完成！{group_label} **{group_count}**、"
                f"輸出品項 **{result.output_items}**、"
                f"無法轉換明細 **{result.problem_count}**、警告 **{len(result.warnings)}**"
            )
            log_error = st.session_state.get("ns_log_error")
            if log_error:
                st.warning(
                    f"⚠️ 轉出紀錄寫入失敗（**不影響這次的轉換結果**）：{log_error}\n\n"
                    "這一批已暫存成「待補紀錄」，請到下方「🗂️ 轉出紀錄管理」補寫回正本。"
                )
            st.download_button(
                f"⬇️ 下載結果（{result.output_name}）",
                data=result.output_bytes,
                file_name=result.output_name,
                mime=EXCEL_MIME,
                key="ns_download",
            )
            if result.problem_rows:
                detail_hint = "（詳見輸出檔的「有問題訂單」工作表）" if not is_m00 else "（M00 格式輸出檔不含問題明細，僅顯示於下方）"
                st.warning(f"有 {len(result.problem_rows)} 筆明細無法轉換{detail_hint}")
                st.dataframe(result.problem_rows, use_container_width=True)
            if result.warnings:
                with st.expander(f"⚠️ 警告訊息（{len(result.warnings)} 則）"):
                    for warning in result.warnings:
                        st.text(f"• {warning}")

    # ---------------------------------------------------------- 轉出紀錄管理
    with st.expander("🗂️ 轉出紀錄管理", expanded=bool(ns_pending_records)):
        current_log = ns_log_state["log"]
        current_records = current_log.get("records", {})
        info_col, refresh_col = st.columns([4, 1])
        with info_col:
            st.caption(
                f"目前記錄 **{len(current_records)}** 張單，存放在 {ns_store.label}"
                f"（讀取於 {ns_log_state['at']}）。"
            )
            if ns_store.kind == "sheet":
                st.caption(
                    "紀錄存在 Google 試算表，伺服器重啟／重新部署都不會消失，"
                    "也可以直接開試算表查或手動修改（改完按 🔄 重新整理）。"
                )
            else:
                st.caption(
                    "尚未設定 Google 試算表，紀錄存在伺服器本機檔案 —— "
                    "**雲端部署重啟後會清空**，請定期下載備份，或依 README 設定 Google 試算表。"
                )
        with refresh_col:
            if st.button("🔄 重新整理", key="ns_log_refresh", use_container_width=True):
                _export_state(ns_store, refresh=True)
                st.rerun()

        if ns_pending_records:
            st.warning(
                f"待補紀錄：**{len(ns_pending_records)}** 張單還沒寫進正本。"
            )
            if st.button("📤 補寫回正本", key="ns_log_sync", type="primary"):
                try:
                    merged = export_log_mod.merge(ns_store.load(), ns_pending_store.load())
                    ns_store.save(merged)
                except Exception as exc:  # noqa: BLE001
                    st.error(f"補寫失敗：{exc}")
                else:
                    ns_pending_store.clear()
                    _export_state(ns_store, refresh=True)
                    st.rerun()

        if current_records:
            recent = sorted(
                current_records.items(),
                key=lambda kv: str(kv[1].get("last_at") or ""),
                reverse=True,
            )[:200]
            st.dataframe(
                pd.DataFrame([
                    {
                        "單號": key,
                        "轉出次數": record.get("count"),
                        "最後轉出": record.get("last_at"),
                        "最後格式": record.get("last_format"),
                        "首次轉出": record.get("first_at"),
                    }
                    for key, record in recent
                ]),
                use_container_width=True,
                hide_index=True,
                height=240,
            )
            if len(current_records) > len(recent):
                st.caption(f"（只顯示最近 {len(recent)} 張，完整內容請下載備份或直接看試算表）")
        else:
            st.caption("目前還沒有任何轉出紀錄。")

        log_col1, log_col2 = st.columns(2)
        with log_col1:
            st.download_button(
                "⬇️ 下載紀錄備份（JSON）",
                data=export_log_mod.dumps(current_log),
                file_name=NS_EXPORT_LOG_PATH.name,
                mime="application/json",
                key="ns_log_download",
                use_container_width=True,
            )
        with log_col2:
            confirm_clear = st.checkbox("我確定要清除全部轉出紀錄", key="ns_log_clear_confirm")
            if st.button(
                "🗑️ 清除全部紀錄",
                key="ns_log_clear",
                disabled=not confirm_clear,
                use_container_width=True,
            ):
                try:
                    ns_store.clear()
                except Exception as exc:  # noqa: BLE001
                    st.error(f"清除失敗：{exc}")
                else:
                    st.session_state["ns_log_clear_confirm"] = False
                    _export_state(ns_store, refresh=True)
                    st.rerun()

        uploaded_log = st.file_uploader(
            "匯入紀錄備份（與現有紀錄合併，次數相加、時間取較晚者）",
            type=["json"],
            key="ns_log_upload",
        )
        if uploaded_log is not None and st.button("📤 合併匯入", key="ns_log_import"):
            try:
                incoming = export_log_mod.loads(uploaded_log.getvalue())
                merged = export_log_mod.merge(ns_store.load(), incoming)
                ns_store.save(merged)
            except Exception as exc:  # noqa: BLE001
                st.error(f"匯入失敗：{exc}")
            else:
                _export_state(ns_store, refresh=True)
                st.success(
                    f"已匯入 {len(incoming.get('records', {}))} 張單的紀錄，"
                    f"合併後共 {len(merged.get('records', {}))} 張。"
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
                st.session_state["cmp_result"] = result

    result = st.session_state.get("cmp_result")
    if result:
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

# ------------------------------------------------------------ 退貨核對

with tab_return:
    st.subheader("客戶退貨授權明細 × HCT 退貨入庫格式 核對")
    st.markdown(
        "上傳 **客戶退貨授權明細報表**（NetSuite RA）與 **HCT 退貨入庫格式** 各一份"
        "（順序不限，程式會自動辨識），以「料號＋效期＋倉別」加總數量核對退貨是否入庫。\n\n"
        "兩邊沒有可以互相對應的單號欄位，因此不核對單號，只核對數量。"
    )
    col1, col2 = st.columns(2)
    with col1:
        ret_a = st.file_uploader("檔案一", type=["xls", "xlsx", "xlsm"], key="ret_a")
    with col2:
        ret_b = st.file_uploader("檔案二", type=["xls", "xlsx", "xlsm"], key="ret_b")

    if ret_a is not None and ret_b is not None:
        if st.button("🚀 開始核對", key="ret_run", type="primary"):
            try:
                with st.spinner("核對中..."):
                    result = return_compare_mod.compare(
                        ret_a.getvalue(), ret_a.name,
                        ret_b.getvalue(), ret_b.name,
                    )
            except Exception as exc:  # noqa: BLE001
                st.session_state.pop("ret_result", None)
                _show_error(exc)
            else:
                st.session_state["ret_result"] = result

    result = st.session_state.get("ret_result")
    if result:
        st.success(f"核對完成！共 **{result.total_rows}** 筆料號＋效期＋倉別組合。")
        metric_cols = st.columns(4)
        for idx, status in enumerate(["一致", "數量不符", "僅退貨授權", "僅入庫記錄"]):
            metric_cols[idx].metric(status, result.status_counts.get(status, 0))
        st.download_button(
            f"⬇️ 下載核對結果（{result.output_name}）",
            data=result.output_bytes,
            file_name=result.output_name,
            mime=EXCEL_MIME,
            key="ret_download",
        )

# ------------------------------------------------------------ 庫存核對

with tab_inventory:
    st.subheader("HCT／代工廠／M00 × NetSuite 庫存核對")
    st.markdown(
        "上傳 **HCT（或代工廠、M00）庫存報表** 與 **NetSuite 庫存報表** 各一份，順序不限，"
        "程式會自動辨識。\n\n"
        "NetSuite 報表支援三種格式：物流核對版（DR_料號/庫存編號/在庫量/可用）、"
        "業務助理版（項目/庫存數 總和）、舊版（DR_料號/項目計數 總和/數量 總和）。\n\n"
        "另一份可以是 **HCT 庫存報表**（儲區類別/客戶產品編號/有效日期/可出數量/庫存數量）、"
        "**代工廠庫存核對報表**（庫存日期/倉別/品號/品名/批號/庫存數量），"
        "或 **M00 電商物流庫存報表**（InventorySummaryReport，取「庫存詳情」分頁，"
        "庫存量＝數量＋組合保留）。\n\n"
        "HCT 對帳只核對 G00 / G10 / G30 / G40 / G80 / G90 倉；"
        "代工廠對帳只核對 **D 開頭代工廠倉**（合計／總計列自動排除，"
        "「庫存數量」同時當作可出與庫存數量）；"
        "M00 對帳只核對 **M00 開頭倉別**。差異基準為 **HCT／代工廠／M00－NetSuite**。"
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
                st.session_state["inv_result"] = result

    result = st.session_state.get("inv_result")
    if result:
        st.success(
            f"核對完成！日期明細 **{len(result.detail_rows)}** 筆、"
            f"料號彙總 **{len(result.item_rows)}** 筆、"
            f"資料異常 **{len(result.anomalies)}** 筆"
        )
        metric_cols = st.columns(6)
        statuses = getattr(result, "statuses", inventory_mod.ALL_STATUSES)
        for idx, status in enumerate(statuses):
            metric_cols[idx].metric(status, result.item_status_counts.get(status, 0))
        st.download_button(
            f"⬇️ 下載核對結果（{result.output_name}）",
            data=result.output_bytes,
            file_name=result.output_name,
            mime=EXCEL_MIME,
            key="inv_download",
        )

# ------------------------------------------------------------ ED 訂單拆解


@st.cache_data(show_spinner=False)
def _builtin_bundle_map_bytes() -> bytes | None:
    """讀內建組合對照表；檔案不在（例如部署時漏帶）就回 None，改請使用者上傳。"""
    try:
        return BUNDLE_MAP_PATH.read_bytes()
    except OSError:
        return None


with tab_bundle:
    st.subheader("ED 訂單明細 299 組合料號 → 單品明細")
    st.markdown(
        "把 NetSuite「每日銷售串接訂單明細」裡的架上組（299 開頭的虛擬組合料號）"
        "展開成單品列：**數量 × 每套數量**、**含稅金額依分攤比率拆分**"
        "（尾差併入比率最大的單品，拆解前後金額總計完全一致），"
        "其餘欄位原樣複製。\n\n"
        "只有 **299 開頭**的架上組需要拆解；其餘料號（含 298 開頭的正常組合品）明細列原樣保留。\n\n"
        "1. 上傳 **每日銷售串接訂單明細**（.xls / .xlsx）\n"
        "2. 組合對照表預設用內建的那份；有更新時可在下方展開區上傳新的覆蓋\n"
        "3. 按「開始拆解」，完成後點「下載結果」"
    )

    bundle_file = st.file_uploader(
        "選擇每日銷售串接訂單明細", type=["xls", "xlsx", "xlsm"], key="bundle_order_file"
    )

    builtin_bytes = _builtin_bundle_map_bytes()
    with st.expander("組合對照表（預設用內建版本，需要更新時才上傳）"):
        if builtin_bytes:
            st.download_button(
                f"⬇️ 下載目前內建的對照表（{BUNDLE_MAP_PATH.name}）",
                data=builtin_bytes,
                file_name=BUNDLE_MAP_PATH.name,
                mime=EXCEL_MIME,
                key="bundle_map_download",
            )
            st.caption(
                "欄位：套件品號／單品品號／單品品名／規格／每套數量／分攤比率。"
                "同一套件的分攤比率加總應為 1（程式載入時會再正規化一次）。"
                "也可直接上傳 ERP 匯出的 MDxxx 原始格式。"
            )
        else:
            st.warning(f"找不到內建對照表（{BUNDLE_MAP_PATH.name}），請在下方上傳一份。")
        map_file = st.file_uploader(
            "上傳新的組合對照表（可留空）", type=["xls", "xlsx", "xlsm"], key="bundle_map_file"
        )

    map_data = map_file.getvalue() if map_file is not None else builtin_bytes
    map_name = map_file.name if map_file is not None else BUNDLE_MAP_PATH.name

    if bundle_file is not None and map_data is None:
        st.error("沒有可用的組合對照表，請先上傳一份再拆解。")
    elif bundle_file is not None and st.button("🚀 開始拆解", key="bundle_run", type="primary"):
        try:
            with st.spinner("拆解中..."):
                result = bundle_split_mod.split(
                    bundle_file.getvalue(), bundle_file.name, map_data, map_name
                )
        except Exception as exc:  # noqa: BLE001
            st.session_state.pop("bundle_result", None)
            _show_error(exc)
        else:
            st.session_state["bundle_result"] = result

    result = st.session_state.get("bundle_result")
    if result:
        st.success(
            f"「{result.source_name}」拆解完成！"
            f"來源明細 **{result.source_rows}** 列 → 拆解後 **{result.output_rows}** 列"
        )
        metric_cols = st.columns(5)
        metric_cols[0].metric("對照表套件數", result.bundle_count)
        metric_cols[1].metric("被拆解的組合列", result.split_rows)
        metric_cols[2].metric("展開的單品列", result.component_rows)
        metric_cols[3].metric("用到的套件品號", result.used_bundles)
        metric_cols[4].metric(
            "含稅金額合計",
            f"{result.amount_after:,.0f}",
            delta=None if result.amount_matched else f"{result.amount_after - result.amount_before:,.2f}",
        )
        if result.amount_matched:
            st.caption(f"✅ 拆解前後含稅金額一致：{result.amount_before:,.2f}")
        else:
            st.error(
                f"拆解前後含稅金額不一致（前 {result.amount_before:,.2f}、"
                f"後 {result.amount_after:,.2f}），請回報此問題。"
            )
        st.download_button(
            f"⬇️ 下載結果（{result.output_name}）",
            data=result.output_bytes,
            file_name=result.output_name,
            mime=EXCEL_MIME,
            key="bundle_download",
        )
        if result.unmapped_rows:
            st.warning(
                f"有 {len(result.unmapped_rows)} 個 299 開頭的架上組料號不在對照表裡，"
                "已原樣保留未拆解；如需拆解請把它補進組合對照表。"
            )
            st.dataframe(result.unmapped_rows, use_container_width=True)
        if result.preview_rows:
            with st.expander(
                f"預覽拆解後明細（前 {len(result.preview_rows)} 列，完整結果請下載）"
            ):
                st.dataframe(
                    pd.DataFrame(result.preview_rows, columns=result.preview_columns),
                    use_container_width=True,
                    hide_index=True,
                )
        if result.warnings:
            with st.expander(f"⚠️ 警告訊息（{len(result.warnings)} 則）"):
                for warning in result.warnings:
                    st.text(f"• {warning}")

st.divider()
st.caption(f"HCT 工具箱 v1.2 ｜ 請用「下載結果」按鈕保存檔案 ｜ {datetime.now():%Y-%m-%d}")
