var hash;

function doGet(request) {
    var pageSelector = new jw.MaterialsTracker.Utilities.PageSelector(request);

    var page = pageSelector.getPage();

    var template = HtmlService.createTemplateFromFile(page.templateName);

    template.data = page.data;

    return template.evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getCoreListData(filter) {
    var centralPurchasingSSID = jw.MaterialsTracker.Config.ConfigurationManager.getSetting('CentralPurchasingSSID');

    var centralPurchasingSS = SpreadsheetApp.openById(centralPurchasingSSID);

    var coreListSheet = centralPurchasingSS.getSheetByName('CoreList');

    var lastRow = coreListSheet.getLastRow();

    var lastColumn = coreListSheet.getLastColumn();

    var coreListRange = coreListSheet.getRange(1, 1, lastRow, lastColumn);

    var filteredRows = null;

    var processedTrades = [];

    //Get the trades from the core list
    var rangeUtils;

    rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(coreListRange);

    var coreListData = rangeUtils.convertToObjectArray();

    coreListData.forEach(function (value, index, arr) {
        if (processedTrades.indexOf(value.trade.toString().trim()) === -1) {
            processedTrades.push(value.trade.toString().trim());
        }
    });

    //If no filter has been passed only retrieve the trades
    if (typeof filter != 'undefined') {
        filteredRows = jw.MaterialsTracker.Utilities.RangeUtilties.findRowsMatchingKey(coreListRange, filter.trade, 0, true);
    } else {
        return {
            trades: processedTrades
        };
    }

    if (filteredRows == null) {
        rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(coreListRange);
    } else {
        rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(filteredRows);
    }

    coreListData = rangeUtils.convertToObjectArray();

    return {
        coreListData: coreListData,
        trades: processedTrades
    };
}
//# sourceMappingURL=Code.js.map
