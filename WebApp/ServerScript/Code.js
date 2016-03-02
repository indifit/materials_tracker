var hash;

function doGet(request) {
    var pageSelector = new jw.MaterialsTracker.Utilities.PageSelector(request);

    var page = pageSelector.getPage();

    var template = HtmlService.createTemplateFromFile(page.templateName);

    template.data = page.data;

    return template.evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getCoreListData() {
    var centralPurchasingSSID = jw.MaterialsTracker.Config.ConfigurationManager.getSetting('CentralPurchasingSSID');

    var centralPurchasingSS = SpreadsheetApp.openById(centralPurchasingSSID);

    var coreListSheet = centralPurchasingSS.getSheetByName('CoreList');

    var lastRow = coreListSheet.getLastRow();

    var coreListRange = coreListSheet.getRange('A1:T' + lastRow);

    var rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(coreListRange);

    var coreListData = rangeUtils.convertToObjectArray();

    return coreListData;
}
//# sourceMappingURL=Code.js.map
