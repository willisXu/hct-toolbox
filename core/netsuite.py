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


def _pct(s: str) -> str:
    return urllib.parse.quote(str(s), safe="~-._")


class NetSuiteError(ValueError):
    pass


# saved search 欄位若沒有在 NetSuite 設定自訂 Label，search.load() 回傳的
# col.label 會是該欄位的英文預設名稱，跟手動匯出檔案的中文欄名對不上
# （既有轉換邏輯認的是中文欄名）。這裡把常見的英文預設名稱轉成對應中文欄名。
_HEADER_ALIASES = {
    "document number": "文件編號",
    "date": "日期",
    "internal id": "內部 ID",
    "item": "項目",
}


def _normalize_headers(rows: list[list[object]]) -> list[list[object]]:
    if not rows:
        return rows
    header = [
        _HEADER_ALIASES.get(str(h).strip().casefold(), h) for h in rows[0]
    ]
    return [header] + rows[1:]


class NetSuiteClient:
    def __init__(self, cfg: dict):
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

            if columns is None:
                columns = payload.get("columns") or []
            data_rows.extend(payload.get("rows") or [])
            page_count = payload.get("pageCount") or 1
            if expected_page_count is None:
                expected_page_count = page_count
            elif page_count != expected_page_count:
                raise NetSuiteError(
                    f"saved search「{search_id}」在分頁抓取途中資料筆數發生變化"
                    "（可能有人同時異動了單據），為避免抓到不一致的結果，請重新抓取一次。"
                )
            page += 1

        if not columns or not data_rows:
            raise NetSuiteError(f"saved search「{search_id}」沒有回傳任何資料列。")
        return _normalize_headers([list(columns)] + data_rows)
