# 📦 HCT 工具箱

網頁工具（Streamlit），對應原本的 Excel VBA 工具：

| 分頁 | 功能 |
|------|------|
| 🛒 訂單轉換 | NetSuite 銷售訂單未出貨明細 → HCT 銷貨報表 / M00 出貨格式 |
| 🔄 調撥單轉換 | NetSuite 調撥單未出貨明細 → HCT 銷貨報表 / M00 出貨格式 |
| 🔗 NetSuite 直接抓取 | 免手動匯出：選 saved search → 抓資料 → 勾選要轉換的列 → 轉換 |
| 🔍 表格核對 | 未出貨明細 × 銷貨單明細，數量與訂單編號核對 |
| 🔁 退貨核對 | 客戶退貨授權明細（NetSuite RA）× HCT 退貨入庫格式，料號＋效期數量核對 |
| 📊 庫存核對 | HCT × NetSuite 庫存報表核對（G00/G10/G30/G40/G80/G90 倉） |
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

料號是 298／299 開頭、但不在對照表裡的，會原樣保留不拆解，並列在輸出檔的「未對應組合料號」
工作表與介面警告區，提醒把它補進對照表。

## NetSuite 直接抓取設定

1. 依 `netsuite_restlet/README.md` 把 `netsuite_restlet/saved_search_restlet.js` 部署成 NetSuite RESTlet。
2. 複製 `.streamlit/secrets.toml.example` 為 `.streamlit/secrets.toml`（本機）或貼進
   Streamlit Community Cloud 的 App settings → Secrets（雲端），填入 NetSuite 帳號、
   OAuth 1.0 Token 與組好的 RESTlet 網址。**此檔含機密，不會進版控。**
3. 需要抓別的 saved search 時，編輯 `mappings/netsuite_saved_searches.yaml` 增加項目。

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
netsuite_restlet/   # 需部署到 NetSuite 的 SuiteScript RESTlet + 部署說明
mappings/           # 輸出範本（HCT範本.xlsx、M00出貨格式.xlsx）、組合對照表、saved search 對照表
tests/              # 回歸測試（python tests/test_core.py，免安裝 pytest）
```
