/**
 * @NApiVersion 2.0
 * @NScriptType Restlet
 *
 * 依 searchId 執行既有的 saved search，一次只回傳一頁（最多 500 列），
 * 讓每次呼叫都拿到全新的 governance 額度，避免大量資料的 saved search
 * 在單次執行裡把所有分頁抓完而撞到系統執行單位上限（NetSuite 對這種
 * governance 超限的中斷屬於系統層級，使用者程式的 try/catch 攔不到，
 * 只會看到籠統的 UNEXPECTED_ERROR）。
 *
 * 呼叫方式：GET .../restlet.nl?script=<此腳本ID>&deploy=1&searchId=customsearchXXX&page=0
 *           加上 &columnsOnly=1 則只回欄位定義、不執行查詢（建立欄名對照表用）
 * 回應：{ columns（每欄 {name, label, join}）, rows（此頁資料列，不含標題）, page, pageCount }
 * 呼叫端見 rows 為空或 page+1 >= pageCount 時停止翻頁。
 */
define(['N/search'], function (search) {

    var PAGE_SIZE = 500;

    function get(context) {
        var searchId = context.searchId;
        if (!searchId) {
            return JSON.stringify({ error: true, name: 'MISSING_PARAM', message: '缺少 searchId 參數' });
        }
        var page = parseInt(context.page, 10);
        if (isNaN(page) || page < 0) {
            page = 0;
        }

        try {
            var loadedSearch = search.load({ id: searchId });
            // 回傳完整欄位定義而不是壓成一個字串：label 會隨語言與自訂標籤變動
            // （沒設中文 Label 的欄拿到的是英文預設名稱 Document Number / Date /
            // Memo (Main)…），name 是欄位內部 ID（tranid / trandate / memomain…），
            // 跟語言無關。呼叫端優先用 label、label 空白才退到 name，並在 label
            // 認不得、內部 ID 卻對得上必要欄位時用內部 ID 補救。
            // 舊版只回字串陣列，呼叫端仍相容（見 core/netsuite.py _normalize_headers）。
            var columns = loadedSearch.columns.map(function (col) {
                return {
                    name: col.name || null,
                    label: col.label || null,
                    join: col.join || null
                };
            });

            // columnsOnly=1：只要欄位定義、不跑查詢。用來建立「英文欄名 → 中文欄名」
            // 對照表（同一個欄位內部 ID 在別支 saved search 有中文 Label 時就能自動
            // 對上），幾乎不耗 governance，可以一次掃過所有 saved search。
            var columnsOnly = context.columnsOnly;
            if (columnsOnly === '1' || columnsOnly === 'true' || columnsOnly === true) {
                return JSON.stringify({
                    columns: columns, rows: [], page: 0, pageCount: 0, columnsOnly: true
                });
            }

            var pagedData = loadedSearch.runPaged({ pageSize: PAGE_SIZE });
            var pageCount = pagedData.pageRanges.length;
            var rows = [];

            if (page < pageCount) {
                var resultPage = pagedData.fetch({ index: page });
                resultPage.data.forEach(function (result) {
                    var rowValues = loadedSearch.columns.map(function (col) {
                        var text = result.getText(col);
                        if (text !== null && text !== undefined && text !== '') {
                            return text;
                        }
                        var value = result.getValue(col);
                        return value === undefined ? null : value;
                    });
                    rows.push(rowValues);
                });
            }

            // RESTlet 的 get() 規定要回傳字串（"Invalid data format. You should
            // return text." 是回傳原始物件時 NetSuite 丟出的錯誤），所以自己
            // JSON.stringify 再回傳，呼叫端用 resp.json() 解析一樣能還原成物件。
            return JSON.stringify({ columns: columns, rows: rows, page: page, pageCount: pageCount });
        } catch (e) {
            // 包住所有例外自己回傳細節，避免 NetSuite 把原始錯誤壓成籠統的
            // UNEXPECTED_ERROR，讓呼叫端看不出真正原因（權限不足、欄位不存在等）。
            // 注意：governance/執行時間超限屬於系統層級中斷，這層 catch 攔不到，
            // 呼叫端仍會看到系統的 UNEXPECTED_ERROR——這正是改成分頁抓取要解決的問題。
            return JSON.stringify({
                error: true,
                name: e.name || 'UNKNOWN_ERROR',
                message: e.message || String(e),
                stack: e.stack || null,
            });
        }
    }

    return { get: get };
});
