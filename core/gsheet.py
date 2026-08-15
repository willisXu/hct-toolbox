# -*- coding: utf-8 -*-
"""Google 試算表讀寫（服務帳戶 / Sheets API v4）。

只用到三個動作：讀整個分頁、清空分頁、寫回整個分頁 —— 「已轉出紀錄」
的模型是「整份讀進來、在記憶體合併、整份寫回去」，不需要逐格更新。

用服務帳戶（service account）而不是 OAuth：Streamlit 是無人值守的網頁
程式，沒有人可以在旁邊點「允許」。設定方式見 README。
"""
from __future__ import annotations

import re
import urllib.parse

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
API_ROOT = "https://sheets.googleapis.com/v4/spreadsheets"
DEFAULT_TIMEOUT = 30


class SheetError(ValueError):
    """Google 試算表操作失敗（訊息直接顯示給使用者看）。"""


def extract_spreadsheet_id(value: str) -> str:
    """接受完整試算表網址或純 ID，一律回傳 ID。"""
    text = str(value or "").strip()
    if not text:
        raise SheetError("沒有填 Google 試算表 ID 或網址。")
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", text)
    if match:
        return match.group(1)
    if "/" in text or " " in text:
        raise SheetError(f"看不出 Google 試算表 ID：{text[:80]}")
    return text


def _quote_range(worksheet: str, a1: str) -> str:
    # 分頁名稱含中文/空白時要用單引號包起來，名稱裡的單引號要重複兩次。
    escaped = str(worksheet).replace("'", "''")
    return urllib.parse.quote(f"'{escaped}'!{a1}", safe="")


class SheetClient:
    """單一試算表的最小 API 客戶端。憑證與連線在第一次呼叫時才建立。"""

    def __init__(self, service_account_info: dict, spreadsheet_id: str, timeout: int = DEFAULT_TIMEOUT):
        self.info = dict(service_account_info or {})
        # 有些貼上方式會讓 private_key 變成字面上的 \n（而不是真的換行），
        # 直接拿去解析會噴 "No key could be detected"，這裡順手救回來。
        key = self.info.get("private_key")
        if isinstance(key, str) and "\\n" in key and "\n" not in key:
            self.info["private_key"] = key.replace("\\n", "\n")
        self.spreadsheet_id = extract_spreadsheet_id(spreadsheet_id)
        self.timeout = timeout
        self._session = None

    # ------------------------------------------------------------ 連線

    def _get_session(self):
        if self._session is not None:
            return self._session
        try:
            from google.auth.transport.requests import AuthorizedSession
            from google.oauth2 import service_account
        except ImportError as exc:  # pragma: no cover - 只在沒裝套件時發生
            raise SheetError(
                "缺少 google-auth 套件，無法連線 Google 試算表。"
                "請執行 pip install -r requirements.txt（雲端部署會自動安裝）。"
            ) from exc
        try:
            creds = service_account.Credentials.from_service_account_info(self.info, scopes=SCOPES)
        except Exception as exc:
            raise SheetError(
                f"Google 服務帳戶憑證不正確：{exc}。"
                "請確認 secrets 裡的 [gsheet_log.service_account] 是完整貼上的金鑰 JSON 內容。"
            ) from exc
        self._session = AuthorizedSession(creds)
        return self._session

    def _request(self, method: str, url: str, **kwargs):
        import requests

        session = self._get_session()
        try:
            resp = session.request(method, url, timeout=self.timeout, **kwargs)
        except requests.RequestException as exc:
            raise SheetError(f"連線 Google 試算表失敗：{exc}") from exc
        if resp.status_code >= 400:
            raise SheetError(self._describe_error(resp))
        try:
            return resp.json() if resp.content else {}
        except ValueError as exc:
            raise SheetError(f"Google 試算表回應不是有效的 JSON：{resp.text[:300]}") from exc

    def _describe_error(self, resp) -> str:
        detail = ""
        try:
            payload = resp.json()
            detail = str(payload.get("error", {}).get("message") or "")
        except ValueError:
            detail = resp.text[:300]
        email = self.info.get("client_email", "（服務帳戶）")
        if resp.status_code == 403:
            return (
                f"Google 試算表沒有權限（HTTP 403）：{detail}\n"
                f"請把試算表分享給服務帳戶 {email}，權限給「編輯者」。"
            )
        if resp.status_code == 404:
            return (
                f"找不到這份 Google 試算表（HTTP 404）：{detail}\n"
                f"請確認 spreadsheet_id 是否正確（目前：{self.spreadsheet_id}）。"
            )
        return f"Google 試算表操作失敗（HTTP {resp.status_code}）：{detail}"

    # ------------------------------------------------------------ 操作

    def worksheet_titles(self) -> list[str]:
        payload = self._request(
            "GET", f"{API_ROOT}/{self.spreadsheet_id}", params={"fields": "sheets.properties.title"}
        )
        return [
            str(sheet.get("properties", {}).get("title", ""))
            for sheet in payload.get("sheets", [])
        ]

    def ensure_worksheet(self, worksheet: str) -> None:
        """分頁不存在就建立（第一次使用時會用到）。"""
        if worksheet in self.worksheet_titles():
            return
        self._request(
            "POST",
            f"{API_ROOT}/{self.spreadsheet_id}:batchUpdate",
            json={"requests": [{"addSheet": {"properties": {"title": worksheet}}}]},
        )

    def read(self, worksheet: str, a1: str = "A:Z") -> list[list[str]]:
        """讀整個分頁；分頁不存在時回空表（第一次寫入前是正常狀態）。"""
        try:
            payload = self._request(
                "GET",
                f"{API_ROOT}/{self.spreadsheet_id}/values/{_quote_range(worksheet, a1)}",
                params={"majorDimension": "ROWS"},
            )
        except SheetError as exc:
            if "Unable to parse range" in str(exc) or "HTTP 400" in str(exc):
                return []
            raise
        return [[("" if cell is None else str(cell)) for cell in row] for row in payload.get("values", [])]

    def write(self, worksheet: str, values: list[list[object]], a1: str = "A:Z") -> None:
        """整份覆寫：先清空範圍再寫入，避免舊資料殘留在下方。"""
        self.ensure_worksheet(worksheet)
        quoted = _quote_range(worksheet, a1)
        self._request("POST", f"{API_ROOT}/{self.spreadsheet_id}/values/{quoted}:clear")
        if not values:
            return
        target = _quote_range(worksheet, "A1")
        self._request(
            "PUT",
            f"{API_ROOT}/{self.spreadsheet_id}/values/{target}",
            params={"valueInputOption": "RAW"},
            json={
                "range": urllib.parse.unquote(target),
                "majorDimension": "ROWS",
                "values": [[("" if cell is None else str(cell)) for cell in row] for row in values],
            },
        )
