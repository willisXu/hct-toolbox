# 📦 HCT 工具箱

四合一網頁工具（Streamlit），對應原本四份 Excel VBA 工具：

| 分頁 | 功能 |
|------|------|
| 🛒 訂單轉換 | NetSuite 銷售訂單未出貨明細 → HCT 銷貨報表 / M00 出貨格式 |
| 🔄 調撥單轉換 | NetSuite 調撥單未出貨明細 → HCT 銷貨報表 / M00 出貨格式 |
| 🔗 NetSuite 直接抓取 | 免手動匯出：選 saved search → 抓資料 → 勾選要轉換的列 → 轉換 |
| 🔍 表格核對 | 未出貨明細 × 銷貨單明細，數量與訂單編號核對 |
| 📊 庫存核對 | HCT × NetSuite 庫存報表核對（G00/G10/G30/G40/G80/G90 倉） |

## NetSuite 直接抓取設定

1. 依 `netsuite_restlet/README.md` 把 `netsuite_restlet/saved_search_restlet.js` 部署成 NetSuite RESTlet。
2. 複製 `.streamlit/secrets.toml.example` 為 `.streamlit/secrets.toml`（本機）或貼進
   Streamlit Community Cloud 的 App settings → Secrets（雲端），填入 NetSuite 帳號、
   OAuth 1.0 Token 與組好的 RESTlet 網址。**此檔含機密，不會進版控。**
3. 需要抓別的 saved search 時，編輯 `mappings/netsuite_saved_searches.yaml` 增加項目。

## 線上使用

部署於 Streamlit Community Cloud，開啟網址即可使用（閒置後首次開啟需等待喚醒約 30 秒～1 分鐘）。

## 本機執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 打包成 exe（Windows）

執行 `建置EXE.bat`，產物在 `release/` 資料夾。

## 專案結構

```
app.py              # Streamlit 介面（五個分頁）
core/
  shipping.py       # 訂單/調撥單 → HCT 銷貨報表（convert 讀檔、convert_rows 共用轉換邏輯）
  m00.py            # 訂單/調撥單 → M00 出貨格式（同上）
  compare.py        # 表格核對
  inventory.py      # 庫存核對
  xlio.py           # Excel 讀寫（NS 的 .xls 是 XML、HCT 的 .xls 是 BIFF）
  netsuite.py       # NetSuite REST OAuth 1.0 (TBA) 客戶端，呼叫 saved search RESTlet
netsuite_restlet/   # 需部署到 NetSuite 的 SuiteScript RESTlet + 部署說明
mappings/           # 輸出範本（HCT範本.xlsx、M00出貨格式.xlsx）、saved search 對照表
```
