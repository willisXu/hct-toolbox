/**
 * @NApiVersion 2.1
 * @NScriptType Restlet
 *
 * 依 searchId 執行既有的 saved search，一次只回傳一頁（最多 500 列），
 * 讓每次呼叫都拿到全新的 governance 額度，避免大量資料的 saved search
 * 在單次執行裡把所有分頁抓完而撞到系統執行單位上限（NetSuite 對這種
 * governance 超限的中斷屬於系統層級，使用者程式的 try/catch 攔不到，
 * 只會看到籠統的 UNEXPECTED_ERROR）。
 *
 * 呼叫方式：GET .../restlet.nl?script=<此腳本ID>&deploy=1&searchId=customsearchXXX&page=0
 * 回應：{ columns, rows（此頁資料列，不含標題）, page, pageCount }
 * 呼叫端見 rows 為空或 page+1 >= pageCount 時停止翻頁。
 */
define(['N/search'], function (search) {

    var PAGE_SIZE = 500;

    function get(context) {
        var searchId = context.searchId;
        if (!searchId) {
            return { error: true, name: 'MISSING_PARAM', message: '缺少 searchId 參數' };
        }
        var page = parseInt(context.page, 10);
        if (isNaN(page) || page < 0) {
            page = 0;
        }

        try {
            var loadedSearch = search.load({ id: searchId });
            var columns = loadedSearch.columns.map(function (col) {
                return col.label || col.name;
            });

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

            return { columns: columns, rows: rows, page: page, pageCount: pageCount };
        } catch (e) {
            // 包住所有例外自己回傳細節，避免 NetSuite 把原始錯誤壓成籠統的
            // UNEXPECTED_ERROR，讓呼叫端看不出真正原因（權限不足、欄位不存在等）。
            // 注意：governance/執行時間超限屬於系統層級中斷，這層 catch 攔不到，
            // 呼叫端仍會看到系統的 UNEXPECTED_ERROR——這正是改成分頁抓取要解決的問題。
            return {
                error: true,
                name: e.name || 'UNKNOWN_ERROR',
                message: e.message || String(e),
                stack: e.stack || null,
            };
        }
    }

    return { get: get };
});
