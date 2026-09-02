# 部署 saved search RESTlet

供「HCT 工具箱」網頁直接用 saved search ID 抓資料，不必手動匯出 .xls。

> **2026-09 更新**：`columns` 改回傳每欄的 `{name, label, join}`（原本只回一個字串），
> 並支援 `&columnsOnly=1` 只取欄位定義、不執行查詢——「🔤 欄名對照表」靠這個自動推導
> 中英文欄名對照。**請重新上傳一次 `saved_search_restlet.js` 覆蓋舊版**；沒更新的話
> 抓取仍可正常運作（呼叫端相容舊格式），只是學不到對照表。

## 部署步驟（在 NetSuite 後台，需 SuiteScript 開發/部署權限）

1. Setup → File Cabinet，上傳 `saved_search_restlet.js` 到任一資料夾（例如 SuiteScripts）。
2. Customization → Scripts → New，Type 選 Restlet，指向剛上傳的檔案，儲存後記下 **Script ID**（例如 `customscript_hct_saved_search`）。
3. 該 Script 的 Deployments 分頁 → New Deployment：
   - Status：Released
   - Log Level：Error（或依需求）
   - Audience：至少要包含會呼叫這支 API 的角色/使用者（通常是產生 Token 的那個角色）
   - 儲存後記下網址列的 **script=** 與 **deploy=** 數字（例如 `script=1090&deploy=1`）。
4. 完整 RESTlet URL（帳號 ID 依 sandbox / 正式區而不同）：

   ```text
   https://<帳號ID小寫、底線改連字號>.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=<Script ID>&deploy=<Deploy ID>
   ```

   實際的帳號 ID、Script ID、Deploy ID 只填進 Streamlit Secrets（見下方），不要寫進這份文件或其他會進版控的檔案。
5. 確認呼叫此 RESTlet 的角色（Token 綁定的角色）有權限「執行」該 script，且對要查的 saved search 所屬的 record type 有 View 權限。

## 設定回 HCT 工具箱

把上面組好的完整 URL，填進 Streamlit Secrets 的 `netsuite.restlet_url`
（見專案根目錄 `.streamlit/secrets.toml.example`），不要寫進程式碼或 git repo。

## 找 saved search 的 ID

NetSuite 裡打開該 saved search 編輯頁，網址列 `id=` 後面的字串（例如
`customsearch1234`）就是 searchId，填進 `mappings/netsuite_saved_searches.yaml`。
