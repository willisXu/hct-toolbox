# -*- coding: utf-8 -*-
"""NetSuite REST 客戶端 — OAuth 1.0 (TBA) HMAC-SHA256 簽章。

沿用「NS 匯入程式」工具的 core/netsuite.py（已在 sandbox / 正式區實測），
簽章邏輯不動；這裡只加一個 run_saved_search()，呼叫使用者自行部署在
NetSuite 裡的 RESTlet（RESTlet 需回傳 {"columns": [...], "rows": [[...], ...]}，
rows[0] 為欄位標題、其餘為資料列，格式對齊本地上傳檔案的讀取結果）。
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

    def run_saved_search(self, restlet_url: str, search_id: str) -> list[list[object]]:
        """呼叫已部署的 saved-search RESTlet，回傳 rows（header 列 + 資料列）。"""
        sep = "&" if "?" in restlet_url else "?"
        url = f"{restlet_url}{sep}searchId={_pct(search_id)}"
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

        rows = payload.get("rows") if isinstance(payload, dict) else None
        if not rows:
            raise NetSuiteError(f"saved search「{search_id}」沒有回傳任何資料列。")
        return rows
