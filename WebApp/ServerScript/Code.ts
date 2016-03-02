var hash: string;

function doGet(request: GoogleAppsScript.Script.IParameters)
{
    var pageSelector: jw.MaterialsTracker.Utilities.PageSelector = new jw.MaterialsTracker.Utilities.PageSelector(request);

    var page: jw.MaterialsTracker.Interfaces.IPage = pageSelector.getPage();    

    var template = HtmlService.createTemplateFromFile(page.templateName);

    template.data = page.data;

    return template.evaluate().setTitle('Materials Tracker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getCoreListData() : any
{
    var centralPurchasingSSID: string = jw.MaterialsTracker.Config.ConfigurationManager.getSetting('CentralPurchasingSSID');

    var centralPurchasingSS = SpreadsheetApp.openById(centralPurchasingSSID);

    var coreListSheet = centralPurchasingSS.getSheetByName('CoreList');

    var lastRow: number = coreListSheet.getLastRow();

    var coreListRange: GoogleAppsScript.Spreadsheet.Range = coreListSheet.getRange('A1:T' + lastRow);

    var rangeUtils = new jw.MaterialsTracker.Utilities.RangeUtilties(coreListRange);

    var coreListData: Object[] = rangeUtils.convertToObjectArray();

    var processedTrades: string[] = [];

    coreListData.forEach((value: any, index: number, arr: Object[]): boolean =>
    {
        if (processedTrades.indexOf(value.trade.toString().trim()) === -1)
        {
            processedTrades.push(value.trade.toString().trim());
        }
    });       

    return {
        coreListData: coreListData,
        trades: processedTrades
    };
}