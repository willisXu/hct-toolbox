# -*- coding: utf-8 -*-
"""NetSuite REST 客戶端 — OAuth 1.0 (TBA) HMAC-SHA256 簽章。

沿用「NS 匯入程式」工具的 core/netsuite.py（已在 sandbox / 正式區實測），
簽章邏輯不動；這裡只加一個 run_saved_search()，呼叫使用者自行部署在
NetSuite 裡的 RESTlet。RESTlet 一次只回傳一頁（見
netsuite_restlet/saved_search_restlet.js），逐頁回傳
{"columns": [...], "rows": [[...], ...], "page": N, "pageCount": M}；
run_saved_search() 逐頁呼叫、組成 [header 列] + 所有資料列，並把 header
裡常見的英文預設欄名（Document Number、Date 等）轉成中文欄名，格式對齊
本地上傳檔案的讀取結果。
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import secrets
import time
import urllib.parse

import requests

from .shipping import ALL_REQUIRED_HEADERS, KNOWN_SOURCE_HEADERS
from .xlio import (
    NETSUITE_HEADER_ALIASES,
    has_cjk,
    header_aliases,
    save_header_alias_cache,
)


def _pct(s: str) -> str:
    return urllib.parse.quote(str(s), safe="~-._")


class NetSuiteError(ValueError):
    pass


# saved search 欄位若沒有在 NetSuite 設定自訂 Label，search.load() 回傳的
# col.label 會是該欄位的英文預設名稱，跟手動匯出檔案的中文欄名對不上
# （既有轉換邏輯認的是中文欄名）。對照表放在 xlio，跟「上傳匯出檔」那條路徑
# 共用同一份：同樣的英文欄名兩邊都要認得，否則同一支 saved search 直接抓取
# 會過、下載成檔案再上傳卻缺欄。
_HEADER_ALIASES = NETSUITE_HEADER_ALIASES

# 內部 ID 補救只作用在「缺了就一定轉不成」的必要欄位上。
_REQUIRED_KEYS = {h.casefold() for h in ALL_REQUIRED_HEADERS}


def _column_parts(column: object) -> tuple[str, str]:
    """把一個欄位定義拆成 (Label, 欄位內部 ID)。

    新版 RESTlet 回 {"name": ..., "label": ...}；舊版（還沒重新部署）只回
    `col.label || col.name` 一個字串，分不出拿到的是哪一種，這時 ID 留空、
    行為跟以前完全一樣。
    """
    if isinstance(column, dict):
        label = str(column.get("label") or "").strip()
        name = str(column.get("name") or "").strip()
        return label, name
    return str(column or "").strip(), ""


def _normalize_headers(rows: list[list[object]]) -> list[list[object]]:
    """欄位定義 → 轉換邏輯認得的中文欄名。

    以 Label 為準（自訂中文標籤 出貨數量／來源倉別… 必須原樣保留，不能被欄位
    內部 ID 的通用譯名蓋掉），Label 空白才退到內部 ID；兩者都認不得時，若內部
    ID 對得上某個「必要欄位」而該欄目前整份報表都沒有，才用內部 ID 補救——
    只在「不補就一定轉不成」時介入，不會亂改看得懂的欄名。
    """
    if not rows:
        return rows
    aliases = header_aliases()
    header: list[object] = []
    ids: list[str] = []
    for column in rows[0]:
        label, name = _column_parts(column)
        text = label or name
        header.append(aliases.get(text.casefold(), text))
        ids.append(name)

    present = {str(h).strip().casefold() for h in header}
    for index, name in enumerate(ids):
        if not name or str(header[index]).strip().casefold() in _REQUIRED_KEYS:
            continue
        canonical = aliases.get(name.casefold())
        if not canonical:
            continue
        key = canonical.casefold()
        if key in _REQUIRED_KEYS and key not in present:
            header[index] = canonical
            present.add(key)
    return [header] + rows[1:]


# ------------------------------------------------------------ 欄名對照自動學習


def _column_join(column: object) -> str:
    return str(column.get("join") or "").strip() if isinstance(column, dict) else ""


def learn_header_aliases(
    columns_by_search: list[tuple[str, list[object]]],
) -> tuple[dict[str, str], list[dict]]:
    """從各支 saved search 的欄位定義推出「英文欄名 → 中文欄名」對照。

    原理：同一個欄位內部 ID（tranid、memomain…）在 A 搜尋有人設了中文 Label
    「文件編號」、在 B 搜尋沒設所以顯示英文預設名「Document Number」，兩邊一
    join 就知道 Document Number 該當成文件編號。全部都是從 NetSuite 現有資料推
    出來的，不用人工維護對照表。

    刻意不碰的：
      * 公式欄（內部 ID 都叫 formulatext/formuladate…，不同欄會撞在一起）
      * 產出的對照鍵一定不含中文——中文欄名永遠以來源檔本身為準，不會被改掉
      * 同一個鍵推出兩個不同中文欄名時，優先採用這套工具本來就認得的那個

    回傳 (對照表, 明細列)；明細列給介面顯示「學到了什麼、從哪支搜尋學到的」。
    """
    # (join, 內部 ID) -> {中文 Label 次數, 英文 Label, 來源搜尋}
    fields: dict[tuple[str, str], dict] = {}
    for search_label, columns in columns_by_search:
        for column in columns or []:
            label, name = _column_parts(column)
            if not name or name.casefold().startswith("formula"):
                continue
            entry = fields.setdefault(
                (_column_join(column).casefold(), name.casefold()),
                {"cjk": {}, "plain": {}, "searches": [], "name": name},
            )
            if search_label not in entry["searches"]:
                entry["searches"].append(search_label)
            if not label:
                continue
            if has_cjk(label):
                entry["cjk"][label] = entry["cjk"].get(label, 0) + 1
            else:
                entry["plain"].setdefault(label.casefold(), label)

    # 鍵 -> {中文欄名: (是否為工具認得的欄名, 次數, 來源搜尋)}
    candidates: dict[str, dict[str, list]] = {}
    for (_join, _name), entry in fields.items():
        if not entry["cjk"]:
            continue
        canonical = max(
            entry["cjk"].items(),
            key=lambda item: (item[0].casefold() in KNOWN_SOURCE_HEADERS, item[1]),
        )[0]
        keys = set(entry["plain"]) | {entry["name"].casefold()}
        for key in keys:
            if has_cjk(key) or key == canonical.casefold():
                continue
            slot = candidates.setdefault(key, {})
            record = slot.setdefault(canonical, [False, 0, []])
            record[0] = canonical.casefold() in KNOWN_SOURCE_HEADERS
            record[1] += 1
            for search_label in entry["searches"]:
                if search_label not in record[2]:
                    record[2].append(search_label)

    aliases: dict[str, str] = {}
    rows: list[dict] = []
    for key, slot in sorted(candidates.items()):
        canonical, record = max(slot.items(), key=lambda item: (item[1][0], item[1][1]))
        aliases[key] = canonical
        rows.append({
            "來源欄名（英文／內部 ID）": key,
            "對應中文欄名": canonical,
            "工具認得": "✅" if record[0] else "",
            "衝突的其他中文欄名": "、".join(n for n in slot if n != canonical),
            "來源 saved search": "、".join(record[2]),
        })
    return aliases, rows


def refresh_header_aliases(
    client: "NetSuiteClient", restlet_url: str, searches: list[dict],
) -> tuple[dict[str, str], list[dict], list[str]]:
    """掃過所有 saved search 的欄位定義，學出對照表並寫進快取。

    單一支搜尋失敗（權限、searchId 打錯）不該讓整批中斷，改成收集訊息回報。
    """
    columns_by_search: list[tuple[str, list[object]]] = []
    notes: list[str] = []
    for item in searches:
        label = str(item.get("label") or item.get("search_id"))
        try:
            columns = client.fetch_search_columns(restlet_url, str(item["search_id"]))
        except NetSuiteError as exc:
            notes.append(f"「{label}」讀取欄位定義失敗：{exc}")
            continue
        if not any(isinstance(column, dict) for column in columns):
            notes.append(
                f"「{label}」回傳的欄位定義沒有欄位內部 ID，"
                "表示 NetSuite 上的 RESTlet 還是舊版，請重新部署 "
                "netsuite_restlet/saved_search_restlet.js 後再試一次。"
            )
            continue
        columns_by_search.append((label, columns))
        notes.append(f"「{label}」讀到 {len(columns)} 個欄位。")

    aliases, rows = learn_header_aliases(columns_by_search)
    if aliases:
        save_header_alias_cache(aliases, [
            {"label": label, "columns": len(columns)} for label, columns in columns_by_search
        ])
    return aliases, rows, notes


class NetSuiteClient:
    # 空白金鑰簽出來的 Authorization header 長 oauth_consumer_key=""，NetSuite 判定
    # 為 malformed、回 400 INVALID_REQUEST——那個訊息跟「script id 根本不存在」一模
    # 一樣，很容易誤判成 RESTlet 部署壞掉。設定不全就在這裡直接講清楚。
    _REQUIRED_CFG = ("account_id", "consumer_key", "consumer_secret", "token_id", "token_secret")

    def __init__(self, cfg: dict):
        missing = [key for key in self._REQUIRED_CFG if not str(cfg.get(key) or "").strip()]
        if missing:
            raise NetSuiteError(
                "NetSuite 連線設定不完整，缺少：" + "、".join(missing) + "。"
                "請在 .streamlit/secrets.toml（本機）或 Streamlit Cloud 的 Secrets 補齊。"
            )
        self.account = cfg["account_id"]
        self.ck = cfg["consumer_key"]
        self.cs = cfg["consumer_secret"]
        self.tk = cfg["token_id"]
        self.ts_ = cfg["token_secret"]

    def _auth_header(self, method: str, url: str) -> str:
        base_url, _, query = url.partition("?")
        oauth = {
            "oauth_consumer_key": self.ck,
            "oauth_token": self.tk,
            "oauth_signature_method": "HMAC-SHA256",
            "oauth_timestamp": str(int(time.time())),
            "oauth_nonce": secrets.token_hex(16),
            "oauth_version": "1.0",
        }
        # 簽章參數 = oauth 參數 + URL query 參數
        params = dict(urllib.parse.parse_qsl(query, keep_blank_values=True))
        params.update(oauth)
        norm = "&".join(
            f"{k}={v}"
            for k, v in sorted((_pct(k), _pct(v)) for k, v in params.items())
        )
        base_string = "&".join([method.upper(), _pct(base_url), _pct(norm)])
        key = f"{_pct(self.cs)}&{_pct(self.ts_)}".encode()
        sig = base64.b64encode(
            hmac.new(key, base_string.encode(), hashlib.sha256).digest()
        ).decode()
        oauth["oauth_signature"] = sig
        header = ", ".join(f'{k}="{_pct(v)}"' for k, v in sorted(oauth.items()))
        return f'OAuth realm="{self.account}", {header}'

    def _request_payload(self, url: str, search_id: str) -> dict:
        """呼叫 RESTlet 並把回應解析成 dict；各種失敗都轉成看得懂的訊息。"""
        try:
            resp = requests.get(
                url,
                headers={"Authorization": self._auth_header("GET", url)},
                timeout=120,
            )
        except requests.RequestException as exc:
            raise NetSuiteError(f"連線 NetSuite 失敗：{exc}") from exc

        if resp.status_code != 200:
            raise NetSuiteError(
                f"NetSuite RESTlet 回應錯誤（HTTP {resp.status_code}）：\n{resp.text[:1000]}"
            )
        try:
            payload = resp.json()
        except ValueError as exc:
            raise NetSuiteError(f"NetSuite RESTlet 回應不是有效的 JSON：{resp.text[:1000]}") from exc

        if not isinstance(payload, dict):
            raise NetSuiteError(f"NetSuite RESTlet 回應格式錯誤：{resp.text[:1000]}")
        if payload.get("error"):
            raise NetSuiteError(
                f"NetSuite saved search「{search_id}」執行失敗"
                f"（{payload.get('name', 'UNKNOWN_ERROR')}）：{payload.get('message', '未知錯誤')}"
            )
        return payload

    def fetch_search_columns(self, restlet_url: str, search_id: str) -> list[object]:
        """只取 saved search 的欄位定義（columnsOnly=1），不執行查詢。

        建立欄名對照表用：一次掃過所有 saved search 也幾乎不耗 governance。
        """
        sep = "&" if "?" in restlet_url else "?"
        url = f"{restlet_url}{sep}searchId={_pct(search_id)}&columnsOnly=1"
        payload = self._request_payload(url, search_id)
        columns = payload.get("columns") or []
        if not columns:
            raise NetSuiteError(
                f"saved search「{search_id}」沒有回傳任何欄位定義，"
                "請確認 RESTlet 部署與 searchId 是否正確。"
            )
        return list(columns)

    def run_saved_search(
        self, restlet_url: str, search_id: str, max_pages: int = 200
    ) -> list[list[object]]:
        """呼叫已部署的 saved-search RESTlet，逐頁抓取並組成 rows（header 列 + 資料列）。

        RESTlet 一次只回傳一頁（見 netsuite_restlet/saved_search_restlet.js），
        讓每次呼叫都拿到全新的 governance 額度，避免資料量大的 saved search
        在單次執行內抓完所有分頁而撞到 NetSuite 系統層級的執行單位上限
        （這種系統層級中斷連 RESTlet 自己的 try/catch 都攔不到，呼叫端只會
        看到籠統的 UNEXPECTED_ERROR）。

        每一頁都是獨立呼叫、各自重新 search.load()/runPaged()，不是同一個
        伺服器端游標；如果 saved search 底層資料在抓頁之間被改動（正式區
        隨時可能有人異動單據），pageCount 會跟著變，跨頁組出來的結果就可能
        不一致。這裡偵測到 pageCount 中途變動就直接報錯，請使用者重新抓取，
        避免悄悄回傳漏列/重複列的資料。
        """
        columns: list[object] | None = None
        data_rows: list[list[object]] = []
        page = 0
        page_count = 1
        expected_page_count: int | None = None

        while page < page_count:
            if page >= max_pages:
                raise NetSuiteError(
                    f"saved search「{search_id}」超過 {max_pages} 頁（每頁 500 列），"
                    "資料量過大，請在 NetSuite 縮小這個 saved search 的篩選範圍。"
                )
            sep = "&" if "?" in restlet_url else "?"
            url = f"{restlet_url}{sep}searchId={_pct(search_id)}&page={page}"
            payload = self._request_payload(url, search_id)

            if columns is None and payload.get("columns"):
                columns = list(payload["columns"])
            data_rows.extend(payload.get("rows") or [])
            # pageCount 為 0 是合法值（saved search 沒有符合條件的結果時
            # pageRanges 是空陣列），不能用 `or 1` 把 0 蓋成 1。
            raw_page_count = payload.get("pageCount")
            page_count = (
                int(raw_page_count)
                if isinstance(raw_page_count, (int, float)) and not isinstance(raw_page_count, bool)
                else 1
            )
            if expected_page_count is None:
                expected_page_count = page_count
            elif page_count != expected_page_count:
                raise NetSuiteError(
                    f"saved search「{search_id}」在分頁抓取途中資料筆數發生變化"
                    "（可能有人同時異動了單據），為避免抓到不一致的結果，請重新抓取一次。"
                )
            page += 1

        if not columns:
            raise NetSuiteError(
                f"saved search「{search_id}」沒有回傳任何欄位定義，"
                "請確認 RESTlet 部署與 searchId 是否正確。"
            )
        # data_rows 可以是空的：saved search 沒有符合條件的結果是正常業務
        # 狀態（例如當天訂單都已出貨），由呼叫端顯示「沒有資料」提示。
        return _normalize_headers([list(columns)] + data_rows)
