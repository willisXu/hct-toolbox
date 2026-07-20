/**
 * @NApiVersion 2.1
 * @NScriptType Restlet
 *
 * 依 searchId 執行既有的 saved search，回傳
 *   { "columns": [欄位標題...], "rows": [[欄位標題...], [資料...], ...] }
 * rows[0] 為欄位標題、其餘為資料列——格式對齊 HCT 工具箱本機上傳檔案
 * 讀進來的表格資料，前端可直接把 rows 丟進既有轉換邏輯。
 *
 * 呼叫方式：GET .../restlet.nl?script=<此腳本ID>&deploy=1&searchId=customsearchXXX
 * 逐欄優先取 getText()（清單/日期欄位的顯示文字），沒有文字才退回 getValue()。
 */
define(['N/search'], function (search) {

    function get(context) {
        var searchId = context.searchId;
        if (!searchId) {
            return { error: true, name: 'MISSING_PARAM', message: '缺少 searchId 參數' };
        }

        try {
            var loadedSearch = search.load({ id: searchId });
            var columns = loadedSearch.columns.map(function (col) {
                return col.label || col.name;
            });

            var rows = [columns];
            var pagedData = loadedSearch.runPaged({ pageSize: 1000 });
            pagedData.pageRanges.forEach(function (pageRange) {
                var page = pagedData.fetch({ index: pageRange.index });
                page.data.forEach(function (result) {
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
            });

            return { columns: columns, rows: rows };
        } catch (e) {
            // 包住所有例外自己回傳細節，避免 NetSuite 把原始錯誤壓成籠統的
            // UNEXPECTED_ERROR，讓呼叫端看不出真正原因（權限不足、欄位不存在等）。
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
