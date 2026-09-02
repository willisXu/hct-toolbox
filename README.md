# 📦 HCT 工具箱

網頁工具（Streamlit），對應原本的 Excel VBA 工具：

| 分頁 | 功能 |
|------|------|
| 🛒 訂單轉換 | NetSuite 銷售訂單未出貨明細 → HCT 銷貨報表 / M00 出貨格式 |
| 🔄 調撥單轉換 | NetSuite 調撥單未出貨明細 → HCT 銷貨報表 / M00 出貨格式 |
| 🔗 NetSuite 直接抓取 | 免手動匯出：選 saved search → 抓資料 → 篩選 → 勾選要轉換的列 → 轉換（已轉出過的單會標記提醒） |
| 🔍 表格核對 | 未出貨明細 × 銷貨單明細，數量與訂單編號核對 |
| 🔁 退貨核對 | 客戶退貨授權明細（NetSuite RA）× HCT 退貨入庫格式，料號＋效期數量核對 |
| 📊 庫存核對 | HCT × NetSuite 庫存報表核對（不限定倉別；M00 除外） |
| 🧩 ED 訂單拆解 | 每日銷售串接訂單明細的 299 組合料號 → 單品明細（數量展開＋金額分攤） |

## ED 訂單拆解（299 組合料號）

上傳 NetSuite「每日銷售串接訂單明細」，把架上組（299 開頭的虛擬組合料號）展開成單品列：

- **數量** ＝ 套件數量 × 每套數量
- **含稅金額** ＝ 依分攤比率拆分，四捨五入尾差併入比率最大的單品，拆解前後金額總計完全一致
  （來源金額是整數時，拆出來也維持整數）
- **結帳單價** ＝ 拆分後金額 ÷ 拆分後數量
- 其餘欄位（單號、客戶、發票資訊…）原樣複製；非組合料號的明細列原樣保留
- 輸出檔在原欄位右邊加上 `拆解狀態／組合來源料號／組合來源品名／每套數量／分攤比率` 供稽核

組合對照表預設用內建的 `mappings/組合對照表.xlsx`（欄位：套件品號／單品品號／單品品名／規格／
每套數量／分攤比率），對照表更新時可在介面的展開區上傳新檔覆蓋，或直接換掉 `mappings/` 裡那份
重新部署。也可以直接丟 ERP 匯出的 `MDxxx_中文` 原始格式，程式會自動判別，同一套件有多版時
取「生效日期最新」那版。

只有 **299 開頭**（架上組虛擬組合）需要拆解。299 開頭但不在對照表裡的，會原樣保留不拆解，
並列在輸出檔的「未對應組合料號」工作表與介面警告區，提醒把它補進對照表。

對照表缺料號時，除了等 ERP 匯出的 `bundles.csv`，也可以用
`tools/derive_bundle_map.py` 從 **NS 品項主檔的「規格」欄**推導補上——該欄本身
就是組合公式（`HHS030A*4+HHS015A`）。詳見 `handoff.md`。
298 開頭是正常組合品（ERP 有自己的品號與庫存），不拆解也不會被當成漏對照。

## 出貨格式的虛擬倉（HCT 銷貨報表 AD 欄）

虛擬倉逐列從來源報表帶出，依序找第一個「有值」的欄：
`虛擬倉／倉別／來源倉別／出貨倉／出貨倉別／地點／來源地點／出貨地點`，
欄位值只要含 `G+兩碼數字`（如 `G30`、`G30 台北倉`、`G30_台北倉`）就辨識得出來。
調撥單 saved search 回的是 **來源倉別／來源地點**（出貨端）——「目標倉別／目標地點」
是收貨端，不會拿來當虛擬倉。

來源報表整份都沒有這些欄位、或有欄位但某些列是空白時，那些列的虛擬倉用範本
預設值 **G10**（料號 9 開頭的陳列/宣傳品仍輸出 G90），兩種情況都會在警告區
提示一次；出貨倉不是 G10 時，請到 saved search 或匯出報表補上來源倉別欄，
不要靠預設值。

## NetSuite 直接抓取：篩選與轉出紀錄

抓回來的資料列可以先篩選再轉換：

- **關鍵字**：跨所有欄位搜尋，空白分隔多個詞＝該列要同時符合（AND）。
- **依欄位篩選**：挑要篩的欄位後出現對應篩選器 —— 一般欄位是相異值多選
  （該欄有空白格時會多一個「（空白）」選項）、日期欄位是日期區間、
  相異值超過 300 種的欄位改用「包含文字」。
- **轉出紀錄**：只看未轉出過／只看已轉出過。

勾選狀態記在「原始列」上，切換篩選條件不會遺失；「全選／取消全選」只作用在
目前篩選結果。**轉換只會轉出目前表格裡（＝通過篩選）且有勾選的列**，若有已勾選
但被篩選隱藏的列，表格下方會明白提示有幾筆不會被轉換。

轉換成功後，該批資料的 **文件編號**（沒有這欄時退回內部 ID）會寫進轉出紀錄
`data/ns_export_log.json`。下次抓到同一張單時：

- 表格多一欄「⚠️ 轉出紀錄」，顯示轉出過幾次、最後一次的時間。
- 勾選到這些列時，上方跳出提醒並列出單號、勾選明細數與轉出紀錄，避免重複出貨。

紀錄最多保留 5000 張單（超過就丟最舊的）。存放位置有兩種：

| 有沒有設定 Google 試算表 | 存在哪 | 能保存多久 |
|---|---|---|
| **有**（建議） | 你的 Google 試算表，一列一張單 | 一直都在。重新部署、休眠喚醒、換伺服器都不受影響，也能直接開試算表查或手改 |
| 沒有 | 伺服器本機 `data/ns_export_log.json` | 本機執行是一直都在；**Streamlit Community Cloud 重新部署／休眠喚醒／重啟就會清空** |

不管哪一種，「🗂️ 轉出紀錄管理」都可以下載 JSON 備份、匯入合併（同一單號次數
相加、時間取較晚者）、清除全部紀錄。

### 把轉出紀錄存到 Google 試算表

1. 開一份 Google 試算表（分頁不用先建，程式第一次寫入時會自動建「轉出紀錄」分頁）。
2. 到 [Google Cloud Console](https://console.cloud.google.com/) 建一個專案 →
   啟用 **Google Sheets API** → 建立**服務帳戶（service account）** →
   在該服務帳戶的「金鑰」頁建立 **JSON 金鑰**並下載。
3. 把試算表**分享**給金鑰 JSON 裡的 `client_email`（那個
   `xxx@xxx.iam.gserviceaccount.com`），權限選「編輯者」。
   —— 這一步最常漏，漏了會看到 HTTP 403。
4. 依 `.streamlit/secrets.toml.example` 的 `[gsheet_log]` 區塊，把試算表網址
   與整份金鑰 JSON 的內容填進 `.streamlit/secrets.toml`（本機）或
   Streamlit Cloud 的 App settings → Secrets（雲端）。

設定好之後分頁下方的「轉出紀錄管理」會顯示「存放在 Google 試算表…」。
紀錄讀取後會存在該次連線的 session 裡（避免每次重跑都連線拖慢介面），
別人同時寫進去的紀錄要按「🔄 重新整理」才會出現；不過**寫入前一定會重讀
試算表再合併**，所以兩個人同時轉換也不會互相蓋掉。

寫入失敗（網路不通、token 過期、忘了分享試算表）時：**轉換照常完成、結果照常
可以下載**，只會跳警告，並把這一批暫存成本機「待補紀錄」；修好之後到
「🗂️ 轉出紀錄管理」按「📤 補寫回正本」就會補上去（不會重複計次）。

## NetSuite 直接抓取設定

1. 依 `netsuite_restlet/README.md` 把 `netsuite_restlet/saved_search_restlet.js` 部署成 NetSuite RESTlet。
2. 複製 `.streamlit/secrets.toml.example` 為 `.streamlit/secrets.toml`（本機）或貼進
   Streamlit Community Cloud 的 App settings → Secrets（雲端），填入 NetSuite 帳號、
   OAuth 1.0 Token 與組好的 RESTlet 網址。**此檔含機密，不會進版控。**
3. 需要抓別的 saved search 時，編輯 `mappings/netsuite_saved_searches.yaml` 增加項目。
4. （選用）要讓轉出紀錄長期保存，依上一節設定 `[gsheet_log]` 存到 Google 試算表。

## 中英文欄名自動對照

saved search 的欄位若沒在 NetSuite 設中文 Label，匯出／抓取拿到的會是 NetSuite 的
英文預設名稱（`Document Number`、`Date`、`Item`、`Memo (Main)`、
`Transaction Serial/Lot Number`…），同一份報表常常中英混雜。工具兩條路徑都會自動
把英文欄名對回中文欄名：

* **內建對照表**（`core/xlio.py` 的 `NETSUITE_HEADER_ALIASES`）收常見的英文預設名稱
  與欄位內部 ID。
* **自動學習**：「🔗 NetSuite 直接抓取 → 🔤 欄名對照表 → 🔄 從 NetSuite 更新欄名對照表」
  會讀每一支 saved search 的欄位定義（只讀欄位、不跑查詢），用「同一個欄位內部 ID 在
  別支 saved search 有中文 Label」推出對照，存進 `mappings/欄位對照快取.json`，
  **上傳檔案**與**直接抓取**兩條路徑共用。saved search 欄位改名後重按一次即可。
  公式欄（內部 ID 都是 `formulatext`/`formuladate`，不同欄會撞在一起）刻意不學，
  那些欄位請直接在 saved search 設中文 Label。
  此功能需要 2026-09 之後版本的 RESTlet（會回傳欄位內部 ID），舊版請重新部署。
  Streamlit Community Cloud 的檔案系統是暫時的，重啟後快取會消失；要讓它常駐，
  把 `mappings/欄位對照快取.json` 提交進版控即可（內容只有欄名，不含機密）。

轉換前的「欄位對照檢查」面板會列出每個欄位對到來源檔哪一欄。必要欄缺少會直接報錯，
**選用欄對不上則不會報錯**（批號、效期、備註會靜默空白），所以有缺時面板會自動展開。

## 線上使用

部署於 Streamlit Community Cloud，開啟網址即可使用（閒置後首次開啟需等待喚醒約 30 秒～1 分鐘）。

結果檔一律用各分頁的「下載結果」按鈕取得，伺服器不留檔。

## 本機開發

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 專案結構

```
app.py              # Streamlit 介面（七個分頁）
core/
  shipping.py       # 訂單/調撥單 → HCT 銷貨報表（convert 讀檔、convert_rows 共用轉換邏輯）
  m00.py            # 訂單/調撥單 → M00 出貨格式（同上）
  compare.py        # 表格核對
  return_compare.py # 退貨核對（客戶退貨授權明細 × HCT 退貨入庫格式）
  inventory.py      # 庫存核對
  bundle_split.py   # ED 訂單明細 299 組合料號拆解（split 讀檔、split_rows 共用拆解邏輯）
  xlio.py           # Excel 讀寫（NS 的 .xls 是 XML、HCT 的 .xls 是 BIFF）
  netsuite.py       # NetSuite REST OAuth 1.0 (TBA) 客戶端，呼叫 saved search RESTlet
  table_filter.py   # 抓回來的資料列篩選（關鍵字／欄位值／包含文字／日期區間）
  export_log.py     # 已轉出紀錄的資料結構與合併邏輯（單號 → 轉出次數/時間/格式）
  export_store.py   # 紀錄存放後端：本機 JSON 檔 / Google 試算表
  gsheet.py         # Google Sheets API v4（服務帳戶）讀寫
tools/              # 維護用腳本(derive_bundle_map.py:從品項主檔規格公式補組合對照表)
netsuite_restlet/   # 需部署到 NetSuite 的 SuiteScript RESTlet + 部署說明
data/               # 本機紀錄／待補紀錄 JSON（自動產生、不進版控）
mappings/           # 輸出範本（HCT範本.xlsx、M00出貨格式.xlsx）、組合對照表、saved search 對照表
tests/              # 回歸測試（python tests/test_core.py，免安裝 pytest）
```
